# Excel Spell & Grammar Checker

A small Streamlit app that scans Excel spreadsheets for spelling and grammar issues.
Cells with problems are highlighted and can be downloaded as a new workbook.

## Features
- Checks spelling with [pyspellchecker](https://github.com/barrust/pyspellchecker)
- Uses LanguageTool's public API for grammar checks (no Java required)
- Highlights spelling and grammar issues directly in the Excel file
- Works across multiple sheets and supports basic language selection

## Installation
```bash
pip install -r requirements.txt
```

## Usage
```bash
streamlit run app.py
```
1. Open the app in your browser.
2. Choose a language (currently `en-US` or `en-GB`).
3. Upload one or more Excel files.
4. Review the summary and download the highlighted workbook.

## Notes
The grammar checker relies on the public LanguageTool API which has usage limits. For large-scale use, consider running your own LanguageTool server.
