# PdfExcelXlator

A Python web app that extracts data from PDF documents — including Traditional Chinese text — and converts them into clean, structured Excel files.

Built with Streamlit, so it runs in the browser with no command-line experience needed.

---

## What it does

- Extracts text and/or tables from PDF files
- Supports Traditional Chinese characters
- Outputs a structured `.xlsx` file ready for budgeting, analysis, or reporting
- Optional data sanitization (redacts emails, phone numbers, ID numbers)
- Processes files entirely in memory — no data is stored

## Why I built it

Reviewing monthly credit card statements meant manually copying rows into a spreadsheet. This tool automates that entire process — what used to take 15+ minutes now takes seconds.

## How to run it locally

**Requirements:** Python 3.8+

```bash
git clone https://github.com/chingyichiang/PdfExcelXlator.git
cd PdfExcelXlator
pip install -r requirements.txt
streamlit run app.py
```

Then open your browser at `http://localhost:8501`.

## How to use

1. Accept the security agreement
2. Upload a PDF file (max 50MB)
3. Choose an extraction method: Text, Tables, or Both
4. Click **Convert to Excel**
5. Preview the extracted data and download the `.xlsx` file

## Project structure

```
app.py               # Streamlit app — UI and main workflow
pdf_processor.py     # Extracts text and tables from PDF
data_sanitizer.py    # Cleans and redacts sensitive data
excel_converter.py   # Writes structured output to Excel
```

## Built with

- [Streamlit](https://streamlit.io) — web app framework
- [pdfplumber](https://github.com/jsvine/pdfplumber) — PDF text and table extraction
- [pandas](https://pandas.pydata.org) — data processing
- [openpyxl](https://openpyxl.readthedocs.io) — Excel file generation
- [Replit](https://replit.com) — development environment
