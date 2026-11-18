# EmailTemplateGenerator — HTML and Outlook OFT

Simple Global Markets reporting email package. The existing Excel-configured layout, placeholders, dynamic links and report-specific image sizes remain unchanged.

## Setup

```
pip install -r requirements.txt
python example.py
```

For `.oft` creation **on Windows** only, install the desktop version of Microsoft Outlook and `pip install pywin32`. Outlook must be configured for the Windows user running the script. No Excel or LibreOffice installation is required.

## Public methods

1. `generate_html(filepath, cobdate, version, reportid, imagedict, link_url, link_text, output_path) -> tuple[str, str]`
   - Saves the complete HTML file at `output_path` (the existing side effect).
   - Returns `(html_string, subject_string)`, **not** the saved path.
   - Uses `Subject` from the single-sheet config, resolving `{{COBDate}}` / `{{Version}}`.
   - Excel screenshots are rendered entirely in memory; no temporary image files in the HTML stage.
2. `generate_email(html, subject, output_path) -> str`
   - Calls private `_generate_oft()` to create an Outlook `.oft` file and returns the absolute output path.
   - **Windows + Outlook desktop + pywin32** required. On Linux, raises `OSError` for `.oft` creation (HTML generation still works).
   - Does **not** send, open or display an email.
   - Converts embedded Base64 images to Outlook-compatible CID attachments. Outlook COM's attachment API requires file paths, so this OFT packaging step uses temporary PNG files and automatically removes them after saving the `.oft`. These files are not created in the HTML generation stage.

## Example

```python
html, subject = generator.generate_html(
    filepath='sample_report.xlsx', cobdate='2026-11-18', version='Daily',
    reportid='GM_REV', imagedict=images,
    link_url='https://example.com/revenue', link_text='View Revenue Report',
    output_path='outputs/revenue.html',
)
# On Windows with Outlook desktop and pywin32:
oft_path = generator.generate_email(html, subject, 'outputs/revenue.oft')
```

## Configuration

One worksheet, one row per `ReportID`. The `EmailLayout` cell controls the order of `{{OpeningLine}}`, `{{Body}}`, `{{Image1}}`, `{{Image2}}`, `{{Image3}}`, `{{AdditionalComment}}`, `{{Link}}`, `{{Ending}}`. Images not supplied in `imagedict` are removed with their caption and spacing. `{{COBDate}}` and `{{Version}}` resolve in text and subject. `link_url` and `link_text` are supplied dynamically.

Image sizes are controlled per report by `Image1Width`, `Image1Height`, `Image2Width`, `Image2Height`, `Image3Width`, `Image3Height` (pixels). Blank width defaults to 650px; blank height uses automatic aspect ratio. Specifying both width and height can stretch images. The source workbook may contain multiple report worksheets.

## Limitations

- `.oft` creation cannot be tested end to end outside Windows with Outlook. Test the resulting `.oft` in the Outlook installation that will consume it; Outlook versions and policies can affect inline images.
- Pillow does not exactly replicate all Excel cell formatting, conditional formatting, charts or shapes. Formula values require valid cached values in the workbook.
- Generated HTML uses Base64 PNGs. Browsers typically display these; Outlook desktop frequently does not, which is why the OFT conversion uses CID images.
