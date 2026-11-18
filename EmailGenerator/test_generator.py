"""Run using `python -m unittest -v test_generator.py`."""
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from email_template_generator import EmailTemplateGenerator

BASE = Path(__file__).resolve().parent

class GeneratorTests(unittest.TestCase):
    def setUp(self):
        self.gen = EmailTemplateGenerator(BASE / 'email_config.xlsx', BASE / 'email_template.html')
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def test_html_returns_string_and_subject_and_saves(self):
        path = Path(self.tmp.name) / 'report.html'
        html, subject = self.gen.generate_html(
            BASE / 'sample_report.xlsx', '2026-11-18', 'Daily', 'GM_RISK',
            {'Image1': {'Sheetname':'Risk Exposure','Imagerange':'B2:F10','Imageinfo':'Risk Summary'}},
            'https://example.com/report', 'View report', path)
        self.assertEqual(path.read_text(encoding='utf-8'), html)
        self.assertIn('<html', html.lower())
        self.assertIn('18 Nov 2026', subject)
        self.assertEqual(html.count('data:image/png;base64,'), 1)
        self.assertNotIn('{{Image3}}', html)

    def test_generate_email_calls_private_backend(self):
        target = Path(self.tmp.name) / 'report.oft'
        with patch.object(self.gen, '_generate_oft', return_value=str(target)) as mocked:
            self.assertEqual(self.gen.generate_email('<html>Hi</html>', 'Subject', target), str(target))
            mocked.assert_called_once_with('<html>Hi</html>', 'Subject', target.resolve())

    def test_oft_requires_windows(self):
        if sys.platform != 'win32':
            with self.assertRaisesRegex(OSError, 'Windows'):
                self.gen.generate_email('<html>Hi</html>', 'Subject', Path(self.tmp.name) / 'test.oft')

    def test_bad_oft_path(self):
        with self.assertRaises(ValueError):
            self.gen.generate_email('<html>Hi</html>', 'Subject', 'not-oft.html')

if __name__ == '__main__':
    unittest.main()
