import pandas as pd
import yaml
from collections import OrderedDict


class ReportingEngine:

    SUPPORTED_AGGREGATIONS = {
        "sum",
        "mean",
        "min",
        "max",
        "count",
        "percentage",
    }

    def __init__(self, config):

        self.config = config["report"]

        self.report_name = self.config["name"]

        self.dimensions = [
            item["name"]
            for item in self.config.get("dimensions", [])
        ]

        self.dimension_ids = {
            dimension: i + 1
            for i, dimension in enumerate(self.dimensions)
        }

        self.value_columns = (
            self.config.get("value_columns", [])
        )

        self.merge_config = (
            self.config.get("merge_filters", [])
        )

        self.hierarchy_mapping_config = (
            self.config.get("hierarchy_mapping", [])
        )

        self.mapping_config = (
            self.config.get("mapping_filters", [])
        )

        self.rollup_config = (
            self.config.get("rollup_rules", [])
        )

        self.hiding_config = (
            self.config.get("display_hiding", [])
        )

        self.order_config = (
            self.config.get("order", {})
        )

        self._validate_config()

    def _validate_config(self):
        if not self.dimensions or len(set(self.dimensions)) != len(self.dimensions):
            raise ValueError("Configure at least one dimension and use distinct dimension names.")
        for metric in self.value_columns:
            aggregation = metric.get("aggregation", "sum").lower()
            if aggregation not in self.SUPPORTED_AGGREGATIONS:
                raise ValueError(f"Unsupported aggregation for {metric['name']}: {aggregation}")

    def _validate_dataframe(self, df):
        required = set(self.dimensions)
        for metric in self.value_columns:
            if metric.get("aggregation", "sum").lower() == "percentage":
                required.update((metric["numerator"], metric["denominator"]))
            else:
                required.add(metric["name"])
        missing = required.difference(df.columns)
        if missing:
            raise ValueError(f"Missing report columns: {', '.join(sorted(missing))}")

    # =========================================================
    # PUBLIC
    # =========================================================

    def generate_report(self, df):

        df = df.copy()

        self._validate_dataframe(df)

        # Preserve original values for hierarchy mapping
        for dimension in self.dimensions:
            df[f"__original_{dimension}"] = df[dimension]

        # 1. Merge
        df = self._apply_merge(df)

        # 2. Hierarchy mapping
        df = self._apply_hierarchy_mapping(df)

        # 3. Mapping
        df = self._apply_mapping(df)

        # 4. Numeric preparation
        df = self._prepare_numeric_columns(df)

        # 5. Build tree
        report, _ = self._build_node(
            df=df,
            level=0,
            value=None
        )

        # 6. Display hiding
        self._apply_hiding_to_tree(report)

        # 7. Ordering
        self._apply_order_to_tree(report)

        return report

    # =========================================================
    # BUILD NODE
    # =========================================================

    def _build_node(
        self,
        df,
        level,
        value
    ):

        # =====================================================
        # ROOT
        # =====================================================

        if level == 0:

            node = OrderedDict()

            node["root_id"] = 0
            node["root_value"] = self.report_name

            child_dimension = self.dimensions[0]

            root_state = self._empty_state()

            children = []

            grouped = df.groupby(
                child_dimension,
                sort=False,
                dropna=False
            )

            for child_value, child_df in grouped:

                if pd.isna(child_value):
                    child_value = "N/A"

                child_node, child_state = (
                    self._build_node(
                        child_df,
                        level=1,
                        value=child_value
                    )
                )

                children.append(child_node)

                # Root is the parent.
                # Apply edge-level exclusions here too.
                child_state = (
                    self._apply_child_rollup_exclusions(
                        parent_dimension=None,
                        parent_value=self.report_name,
                        child_dimension=child_dimension,
                        child_value=child_value,
                        child_state=child_state
                    )
                )

                root_state = self._combine_states(
                    root_state,
                    child_state
                )

            node["values"] = (
                self._state_to_values(root_state)
            )

            if children:
                node["child"] = children

            return node, root_state

        # =====================================================
        # NORMAL NODE
        # =====================================================

        current_dimension = self.dimensions[level - 1]

        node = OrderedDict()

        node["root_id"] = (
            self.dimension_ids[current_dimension]
        )

        node["root_name"] = current_dimension
        node["root_value"] = value

        # =====================================================
        # LEAF
        # =====================================================

        if level == len(self.dimensions):

            own_state = self._dataframe_to_state(df)

            # Display always shows the actual node value.
            node["values"] = (
                self._state_to_values(own_state)
            )

            # A leaf's state is returned unchanged.
            #
            # The PARENT decides whether this child contributes.
            #
            return node, own_state

        # =====================================================
        # PARENT
        # =====================================================

        child_dimension = self.dimensions[level]

        parent_state = self._empty_state()

        children = []

        grouped = df.groupby(
            child_dimension,
            sort=False,
            dropna=False
        )

        for child_value, child_df in grouped:

            if pd.isna(child_value):
                child_value = "N/A"

            child_node, child_state = (
                self._build_node(
                    child_df,
                    level + 1,
                    child_value
                )
            )

            children.append(child_node)

            # =================================================
            # CRITICAL SECTION
            # =================================================
            #
            # We now know:
            #
            # parent_dimension = current_dimension
            # parent_value     = value
            #
            # child_dimension  = child_dimension
            # child_value      = child_value
            #
            # Therefore the configuration can be evaluated
            # against the exact parent-child edge.
            # =================================================

            child_state = (
                self._apply_child_rollup_exclusions(
                    parent_dimension=current_dimension,
                    parent_value=value,
                    child_dimension=child_dimension,
                    child_value=child_value,
                    child_state=child_state
                )
            )

            parent_state = self._combine_states(
                parent_state,
                child_state
            )

        # -----------------------------------------------------
        # Parent display value is based on what actually
        # rolls up from its children.
        # -----------------------------------------------------

        node["values"] = (
            self._state_to_values(parent_state)
        )

        if children:
            node["child"] = children

        return node, parent_state

    # =========================================================
    # CHILD ROLLUP EXCLUSION
    # =========================================================

    def _apply_child_rollup_exclusions(
        self,
        parent_dimension,
        parent_value,
        child_dimension,
        child_value,
        child_state
    ):

        # Root-level exclusions are optional.
        # If parent_dimension is None, there is no configured
        # parent dimension to match.
        if parent_dimension is None:
            return child_state

        excluded_columns = set()

        # -----------------------------------------------------
        # Find rules for this exact parent.
        # -----------------------------------------------------

        for rule in self.rollup_config:

            if rule.get(
                "parent_dimension"
            ) != parent_dimension:

                continue

            parent_values = rule.get(
                "parent_values",
                []
            )

            if parent_value not in parent_values:

                continue

            # -------------------------------------------------
            # Check every exclusion under this parent.
            # -------------------------------------------------

            for exclusion in rule.get(
                "exclusions",
                []
            ):

                if exclusion.get(
                    "dimension"
                ) != child_dimension:

                    continue

                child_values = exclusion.get(
                    "values",
                    []
                )

                if child_value not in child_values:

                    continue

                columns = exclusion.get(
                    "value_columns",
                    []
                )

                if columns == "all":

                    for item in self.value_columns:

                        aggregation = item.get(
                            "aggregation",
                            "sum"
                        ).lower()

                        if aggregation != "percentage":

                            excluded_columns.add(
                                item["name"]
                            )

                else:

                    excluded_columns.update(
                        columns
                    )

        if not excluded_columns:

            return child_state

        # -----------------------------------------------------
        # Copy state so the child's own state remains intact.
        # -----------------------------------------------------

        result = {
            column: metric.copy()
            for column, metric
            in child_state.items()
        }

        # -----------------------------------------------------
        # Remove only the requested metrics from the
        # contribution to this parent.
        # -----------------------------------------------------

        for column in excluded_columns:

            if column not in result:
                continue

            result[column]["sum"] = 0
            result[column]["count"] = 0
            result[column]["min"] = None
            result[column]["max"] = None

        return result

    # =========================================================
    # STATE FROM DATAFRAME
    # =========================================================

    def _dataframe_to_state(self, df):

        state = {}

        for item in self.value_columns:

            name = item["name"]

            aggregation = item.get(
                "aggregation",
                "sum"
            ).lower()

            if aggregation == "percentage":
                continue

            series = pd.to_numeric(
                df[name],
                errors="coerce"
            )

            state[name] = {
                "aggregation": aggregation,
                "sum": series.sum(skipna=True),
                "count": series.count(),
                "min": (
                    series.min()
                    if series.count() > 0
                    else None
                ),
                "max": (
                    series.max()
                    if series.count() > 0
                    else None
                )
            }

        return state

    # =========================================================
    # EMPTY STATE
    # =========================================================

    def _empty_state(self):

        state = {}

        for item in self.value_columns:

            aggregation = item.get(
                "aggregation",
                "sum"
            ).lower()

            if aggregation == "percentage":
                continue

            name = item["name"]

            state[name] = {
                "aggregation": aggregation,
                "sum": 0,
                "count": 0,
                "min": None,
                "max": None
            }

        return state

    # =========================================================
    # COMBINE STATES
    # =========================================================

    def _combine_states(
        self,
        left,
        right
    ):

        result = {}

        for item in self.value_columns:

            name = item["name"]

            aggregation = item.get(
                "aggregation",
                "sum"
            ).lower()

            if aggregation == "percentage":
                continue

            l = left.get(
                name,
                self._empty_metric(aggregation)
            )

            r = right.get(
                name,
                self._empty_metric(aggregation)
            )

            result[name] = {
                "aggregation": aggregation,

                "sum": (
                    l["sum"] +
                    r["sum"]
                ),

                "count": (
                    l["count"] +
                    r["count"]
                ),

                "min": self._combine_min(
                    l["min"],
                    r["min"]
                ),

                "max": self._combine_max(
                    l["max"],
                    r["max"]
                )
            }

        return result

    # =========================================================
    # EMPTY METRIC
    # =========================================================

    def _empty_metric(self, aggregation):

        return {
            "aggregation": aggregation,
            "sum": 0,
            "count": 0,
            "min": None,
            "max": None
        }

    # =========================================================
    # MIN
    # =========================================================

    def _combine_min(self, left, right):

        values = [
            x
            for x in [left, right]
            if x is not None
            and not pd.isna(x)
        ]

        return min(values) if values else None

    # =========================================================
    # MAX
    # =========================================================

    def _combine_max(self, left, right):

        values = [
            x
            for x in [left, right]
            if x is not None
            and not pd.isna(x)
        ]

        return max(values) if values else None

    # =========================================================
    # STATE → VALUES
    # =========================================================

    def _state_to_values(self, state):

        result = OrderedDict()

        for item in self.value_columns:

            name = item["name"]

            aggregation = item.get(
                "aggregation",
                "sum"
            ).lower()

            if aggregation == "percentage":
                continue

            metric = state.get(name)

            if metric is None:
                result[name] = None
                continue

            if aggregation == "sum":

                result[name] = metric["sum"]

            elif aggregation == "mean":

                if metric["count"] == 0:
                    result[name] = None
                else:
                    result[name] = (
                        metric["sum"]
                        / metric["count"]
                    )

            elif aggregation == "min":

                result[name] = metric["min"]

            elif aggregation == "max":

                result[name] = metric["max"]

            elif aggregation == "count":

                result[name] = metric["count"]

        # -----------------------------------------------------
        # Percentage
        # -----------------------------------------------------

        for item in self.value_columns:

            name = item["name"]

            if item.get(
                "aggregation",
                ""
            ).lower() != "percentage":

                continue

            numerator = item["numerator"]
            denominator = item["denominator"]

            numerator_state = state.get(
                numerator
            )

            denominator_state = state.get(
                denominator
            )

            if (
                numerator_state is None
                or denominator_state is None
            ):

                result[name] = None
                continue

            numerator_value = (
                numerator_state["sum"]
            )

            denominator_value = (
                denominator_state["sum"]
            )

            if (
                denominator_value is None
                or denominator_value == 0
            ):

                result[name] = None

            else:

                result[name] = round(
                    numerator_value
                    / denominator_value
                    * 100,
                    item.get(
                        "decimal_places",
                        2
                    )
                )

        return result

    # =========================================================
    # MERGE
    # =========================================================

    def _apply_merge(self, df):

        for rule in self.merge_config:

            dimension = rule["dimension"]

            for group in rule.get(
                "groups",
                []
            ):

                target = group["name"]

                values = group.get(
                    "values",
                    []
                )

                if not values:
                    continue

                mask = df[
                    dimension
                ].isin(values)

                df.loc[
                    mask,
                    dimension
                ] = target

        return df

    # =========================================================
    # HIERARCHY MAPPING
    # =========================================================

    def _apply_hierarchy_mapping(self, df):

        for rule in self.hierarchy_mapping_config:

            parent_dimension = (
                rule["parent_dimension"]
            )

            child_dimension = (
                rule["child_dimension"]
            )

            for mapping_rule in rule.get(
                "rules",
                []
            ):

                parent_value = (
                    mapping_rule["parent_value"]
                )

                source_dimension = (
                    mapping_rule.get(
                        "source_dimension",
                        parent_dimension
                    )
                )

                source_values = (
                    mapping_rule.get(
                        "source_values",
                        []
                    )
                )

                child_config = (
                    mapping_rule.get(
                        "child_value",
                        {}
                    )
                )

                original_source = (
                    f"__original_{source_dimension}"
                )

                mask = df[
                    original_source
                ].isin(source_values)

                df.loc[
                    mask,
                    parent_dimension
                ] = parent_value

                child_type = child_config.get(
                    "type"
                )

                if child_type == "source":

                    source_column = (
                        child_config["column"]
                    )

                    original_child = (
                        f"__original_{source_column}"
                    )

                    df.loc[
                        mask,
                        child_dimension
                    ] = df.loc[
                        mask,
                        original_child
                    ]

                elif child_type == "constant":

                    df.loc[
                        mask,
                        child_dimension
                    ] = child_config.get(
                        "value"
                    )

        return df

    # =========================================================
    # MAPPING
    # =========================================================

    def _apply_mapping(self, df):

        for rule in self.mapping_config:

            dimension = rule["dimension"]

            mapping = rule.get(
                "mapping",
                {}
            )

            if mapping:

                df[dimension] = (
                    df[dimension].replace(
                        mapping
                    )
                )

        return df

    # =========================================================
    # NUMERIC PREPARATION
    # =========================================================

    def _prepare_numeric_columns(self, df):

        columns = set()

        for item in self.value_columns:

            aggregation = item.get(
                "aggregation",
                "sum"
            ).lower()

            if aggregation == "percentage":

                columns.add(
                    item["numerator"]
                )

                columns.add(
                    item["denominator"]
                )

            else:

                columns.add(
                    item["name"]
                )

        for column in columns:

            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            )

        return df

    # =========================================================
    # DISPLAY HIDING
    # =========================================================

    def _apply_hiding_to_tree(self, node):

        if "child" not in node:
            return

        visible = []

        for child in node["child"]:

            dimension = child["root_name"]
            value = child["root_value"]

            if self._is_hidden(
                dimension,
                value
            ):
                continue

            self._apply_hiding_to_tree(child)

            visible.append(child)

        if visible:
            node["child"] = visible
        else:
            node.pop("child", None)

    def _is_hidden(
        self,
        dimension,
        value
    ):

        for rule in self.hiding_config:

            if rule["dimension"] != dimension:
                continue

            if value in rule.get(
                "values",
                []
            ):
                return True

        return False

    # =========================================================
    # ORDERING
    # =========================================================

    def _apply_order_to_tree(self, node):

        if "child" not in node:
            return

        children = node["child"]

        if not children:
            return

        dimension = children[0]["root_name"]

        node["child"] = self._sort_children(
            children,
            dimension
        )

        for child in node["child"]:
            self._apply_order_to_tree(child)

    def _sort_children(
        self,
        children,
        dimension
    ):

        rule = self.order_config.get(
            dimension
        )

        if not rule:
            return children

        if rule.get("type") != "custom":
            return children

        configured_order = rule.get(
            "values",
            []
        )

        order_map = {
            value: index
            for index, value
            in enumerate(configured_order)
        }

        return sorted(
            children,
            key=lambda child:
                order_map.get(
                    child["root_value"],
                    len(order_map)
                )
        )


# =============================================================
# CONFIG LOADER
# =============================================================

def load_config(config_path):

    with open(
        config_path,
        "r",
        encoding="utf-8"
    ) as file:

        return yaml.safe_load(file)


# =============================================================
# GENERATE REPORT
# =============================================================

def generate_report(
    df,
    config_path
):

    config = load_config(
        config_path
    )

    engine = ReportingEngine(
        config
    )

    return engine.generate_report(df)
