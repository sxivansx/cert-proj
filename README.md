# Certificate Generator

A simple Python-based tool to generate certificates from a template image and an Excel sheet. The script reads user details from `recipients.xlsx`, prints Name, USN, and Team Name on the certificate template, and exports the generated certificates into the `output/` folder.

## Features
- Reads Name, USN, and Team Name from an Excel sheet.
- Renders text on a certificate PNG template.
- Supports custom fonts, colors, and positioning.
- Automatically exports final certificates as PNG files.

## Project Structure
```
cert-proj/
│
├── generate_and_send.py
├── data/
│   └── recipients.xlsx
├── template/
│   └── certificate.png
├── fonts/
│   └── FontName.ttf
└── output/
```

## Requirements
- Python 3
- Virtual environment (recommended)

Install dependencies:
```
pip install pillow pandas openpyxl
```

## Usage
1. Place your certificate template inside the `template/` folder.
2. Place your Excel file (`recipients.xlsx`) inside the `data/` folder.
3. Update coordinates, font paths, and settings inside `generate_and_send.py` as needed.
4. Run the script:
```
python3 generate_and_send.py
```

Certificates will be generated inside the `output/` folder.

## Excel Format
The script expects the following column names:
```
Name
USN
TEAM NAME
```

## Customization
You can edit:
- Text coordinates
- Font file and size
- Text color
- Output file naming

All settings are inside the configuration section of the script.

## License
Free to use 
