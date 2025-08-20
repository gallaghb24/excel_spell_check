import streamlit as st
import pandas as pd
import io
import logging
from spellchecker import SpellChecker
from language_tool_python import LanguageToolPublicAPI
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import re
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)

# Initialize spell checker and grammar tool
@st.cache_resource
def load_checkers(language: str):
    spell_lang = language.split('-')[0]
    spell = SpellChecker(language=spell_lang)
    tool = LanguageToolPublicAPI(language)
    return spell, tool

def check_spelling(text: str, spell_checker) -> List[str]:
    """Check spelling in text and return list of misspelled words"""
    if not isinstance(text, str) or not text.strip():
        return []
    
    # Extract words (remove punctuation and numbers)
    words = re.findall(r'\b[a-zA-Z]+\b', text.lower())
    misspelled = spell_checker.unknown(words)
    return list(misspelled)

def check_grammar(text: str, grammar_tool) -> List[str]:
    """Check grammar in text and return list of issues"""
    if not isinstance(text, str) or not text.strip():
        return []

    try:
        matches = grammar_tool.check(text)
        issues = [match.message for match in matches]
        return issues
    except Exception as e:
        logger.error("Grammar check failed: %s", e)
        return []

def process_excel_data(excel_data: Dict[str, pd.DataFrame], spell_checker, grammar_tool) -> Tuple[Dict, Dict]:
    """Process Excel data and return dictionaries of issues by sheet"""

    spelling_issues = {}
    grammar_issues = {}

    for sheet_name, df in excel_data.items():
        spelling_issues[sheet_name] = {}
        grammar_issues[sheet_name] = {}

        for row_idx, row in enumerate(df.itertuples(index=False, name=None)):
            for col_idx, cell_value in enumerate(row):
                if pd.notna(cell_value) and isinstance(cell_value, str):
                    misspelled = check_spelling(cell_value, spell_checker)
                    if misspelled:
                        spelling_issues[sheet_name][(row_idx, col_idx)] = misspelled

                    if len(cell_value.split()) > 2:
                        grammar_errs = check_grammar(cell_value, grammar_tool)
                        if grammar_errs:
                            grammar_issues[sheet_name][(row_idx, col_idx)] = grammar_errs

    return spelling_issues, grammar_issues

def create_highlighted_excel(excel_data: Dict[str, pd.DataFrame], spelling_issues, grammar_issues) -> io.BytesIO:
    """Create Excel file with highlighted issues"""
    
    # Create a new workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Define fill colors
    spelling_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Light red
    grammar_fill = PatternFill(start_color="CCCCFF", end_color="CCCCFF", fill_type="solid")   # Light blue
    both_fill = PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid")     # Light purple
    
    for sheet_name, df in excel_data.items():
        ws = wb.create_sheet(title=sheet_name)
        
        # Add data to worksheet
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Apply highlighting
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                cell = ws.cell(row=row_idx + 2, column=col_idx + 1)  # +2 for header and 0-based indexing
                
                has_spelling = (row_idx, col_idx) in spelling_issues.get(sheet_name, {})
                has_grammar = (row_idx, col_idx) in grammar_issues.get(sheet_name, {})
                
                if has_spelling and has_grammar:
                    cell.fill = both_fill
                elif has_spelling:
                    cell.fill = spelling_fill
                elif has_grammar:
                    cell.fill = grammar_fill
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("📝 Excel Spell & Grammar Checker")
    st.write("Upload Excel files to check for spelling and grammar issues. Issues will be highlighted in the output file.")

    language = st.sidebar.selectbox("Language", ["en-US", "en-GB"])

    with st.spinner("Loading spell and grammar checkers..."):
        spell_checker, grammar_tool = load_checkers(language)

    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Upload one or more Excel files for spell and grammar checking",
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.subheader(f"Processing: {uploaded_file.name}")

            with st.spinner(f"Analyzing {uploaded_file.name}..."):
                try:
                    excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
                    spelling_issues, grammar_issues = process_excel_data(excel_data, spell_checker, grammar_tool)
                    display_results(excel_data, spelling_issues, grammar_issues, uploaded_file.name)
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {str(e)}")

            st.write("---")


def display_results(excel_data, spelling_issues, grammar_issues, filename):
    total_spelling = sum(len(issues) for issues in spelling_issues.values())
    total_grammar = sum(len(issues) for issues in grammar_issues.values())

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Sheets Processed", len(spelling_issues))
    with col2:
        st.metric("Spelling Issues", total_spelling)
    with col3:
        st.metric("Grammar Issues", total_grammar)

    if total_spelling > 0 or total_grammar > 0:
        with st.expander("View Detailed Issues"):
            for sheet_name in spelling_issues.keys():
                if spelling_issues[sheet_name] or grammar_issues[sheet_name]:
                    st.write(f"**Sheet: {sheet_name}**")

                    if spelling_issues[sheet_name]:
                        st.write("🔤 Spelling Issues:")
                        for (row, col), words in spelling_issues[sheet_name].items():
                            st.write(f"  Row {row+1}, Column {col+1}: {', '.join(words)}")

                    if grammar_issues[sheet_name]:
                        st.write("📝 Grammar Issues:")
                        for (row, col), issues in grammar_issues[sheet_name].items():
                            st.write(f"  Row {row+1}, Column {col+1}: {'; '.join(issues)}")

                    st.write("---")

        with st.spinner("Creating highlighted Excel file..."):
            highlighted_file = create_highlighted_excel(excel_data, spelling_issues, grammar_issues)

        download_name = filename.replace('.xlsx', '_checked.xlsx').replace('.xls', '_checked.xlsx')
        st.download_button(
            label=f"📥 Download {download_name}",
            data=highlighted_file,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.write("**Legend:**")
        st.write("🔴 Light Red = Spelling Issues")
        st.write("🔵 Light Blue = Grammar Issues")
        st.write("🟣 Light Purple = Both Spelling & Grammar Issues")

    else:
        st.success("✅ No spelling or grammar issues found!")

if __name__ == "__main__":
    main()
