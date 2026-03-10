# SDS Chemical Extractor

A web app that extracts **Product Name**, **Chemical Name**, and **CAS Numbers** from SDS (Safety Data Sheet) PDFs and downloads the results as an Excel file.

## How it works

1. Upload one or more SDS PDF files via the browser interface
2. The app parses Section 1 (product name) and Section 3 (composition/ingredients)
3. Downloads `chemicals_output.xlsx` with 3 columns:

| Product Name | Chemical Name | CAS Number |
|---|---|---|
| Caustic Soda | Sodium Hydroxide | 1310-73-2 |
| Hydrogen peroxide | Water | 7732-18-5 |

## Run locally

```bash
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`

## Command-line usage

```bash
# Single PDF
python pdf_to_excel.py input.pdf output.xlsx

# Entire folder
python pdf_to_excel.py SDS/ output.xlsx
```

## Tech stack

- **pdfplumber** — PDF text & table extraction
- **Flask** — web interface
- **pandas + openpyxl** — Excel output
