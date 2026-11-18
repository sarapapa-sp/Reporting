"""Generate HTML on any platform; optionally create .oft on Windows + Outlook."""
from pathlib import Path
import sys
from email_template_generator import EmailTemplateGenerator

folder = Path(__file__).resolve().parent
generator = EmailTemplateGenerator(folder / 'email_config.xlsx', folder / 'email_template.html')
reports = {
    'GM_REV': {
        'Image1': {'Sheetname': 'Revenue Summary', 'Imagerange': 'B2:F11', 'Imageinfo': 'Revenue by business line (kEUR)'},
        'Image2': {'Sheetname': 'Revenue Drivers', 'Imagerange': 'B2:F9', 'Imageinfo': 'Revenue movement drivers'},
        'Image3': {'Sheetname': 'Revenue Controls', 'Imagerange': 'B2:F9', 'Imageinfo': 'Revenue control results'},
    },
    'GM_RISK': {
        'Image1': {'Sheetname': 'Risk Exposure', 'Imagerange': 'B2:F10', 'Imageinfo': 'Risk exposure by business line'},
        'Image2': {'Sheetname': 'Limit Monitoring', 'Imagerange': 'B2:F9', 'Imageinfo': 'Limit monitoring'},
    },
}

for reportid, images in reports.items():
    html_path = folder / 'outputs' / f'{reportid}_20261118.html'
    html_string, subject = generator.generate_html(
        filepath=folder / 'sample_report.xlsx',
        cobdate='2026-11-18',
        version='Daily',
        reportid=reportid,
        imagedict=images,
        link_url=f'https://example.com/global-markets/{reportid.lower()}',
        link_text='View detailed report',
        output_path=html_path,
    )
    print('HTML saved:', html_path)
    print('Subject:', subject)
    print('HTML length:', len(html_string))

    if sys.platform == 'win32':
        # Outlook desktop and pywin32 must be installed to produce a real .oft.
        oft_path = generator.generate_email(html_string, subject, folder / 'outputs' / f'{reportid}_20261118.oft')
        print('OFT saved:', oft_path)
    else:
        print('OFT skipped: Windows + Outlook desktop required')
