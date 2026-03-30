import streamlit as st
import pandas as pd
import tempfile
import re
import io

def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the correct sheet.")
        return st.selectbox("Select the correct worksheet:", xl.sheet_names, key=preferred_name) 
    
def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def get_clean_cols(filepath, sheet, col_prs, col_urs=None, skiprows=0):
    xl = pd.ExcelFile(filepath)
    if sheet not in xl.sheet_names:
        st.error(f"Worksheet named '{sheet}' not found in the uploaded file. Available sheets: {xl.sheet_names}")
        st.stop()
    df = pd.read_excel(filepath, sheet_name=sheet, skiprows=skiprows)
    if col_prs not in df.columns:
        st.error(f"Column '{col_prs}' not found in sheet '{sheet}'. Available columns: {list(df.columns)}")
        st.stop()
    prs = df[col_prs].dropna().astype(str).str.strip()
    if col_urs:
        if col_urs not in df.columns:
            st.error(f"Column '{col_urs}' not found in sheet '{sheet}'. Available columns: {list(df.columns)}")
            st.stop()
        urs = df[col_urs].dropna().astype(str).str.strip()
        return prs, urs
    return prs 

IGNORE_SENTENCES = [
    'Note (N/A*): The columns “Actual Result /Description”, "Version tested", "Date tested", "Tester", "Conclusion (Pass/Fail)", "Defect/Enhancement Number", and "Defect/Enhancement Status" from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.',
    'Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.',
    'Note: Defect/Enhancement are linked to test cases and their respective execution versions',
    'NA* Closed Service Orders',
    'Note *: From Version Release 5.01.1835.00 onward all information related to the release version will consist of the 4 positions of the version ID (X.YY.ZZZZ.AAAA) intesd of 2 positions as it was used (X.YY)'
] 

IGNORE_SENTENCES_VAL_TP_TM = IGNORE_SENTENCES + ['* Not applicable']

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def normalize_spaces(series):
    return pd.Series(series).astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))    
    
def run_comparison():    
    st.markdown("### PRS DOC x Requirements TM Comparison")
    st.info("Upload your PRS DOC and Requirements TM files to compare requirements.")
    col1, col2 = st.columns(2)
    with col1:
        prs_file = st.file_uploader("Upload PRS DOC (.xlsx)", type=["xlsx"], key="prsdoc_ver")
    with col2:
        tm_file = st.file_uploader("Upload Requirements TM (.xlsx)", type=["xlsx"], key="tm_ver")
    if prs_file and tm_file:
        if st.button("🔍 Run Comparison", key="run_prs_tm_ver"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_prs:
                tmp_prs.write(prs_file.read())
                prs_path = tmp_prs.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tm:
                tmp_tm.write(tm_file.read())
                tm_path = tmp_tm.name
            xl_tm = pd.ExcelFile(prs_path)
            xl_tp = pd.ExcelFile(tm_path)
            prs_sheet = get_sheet_with_fallback(xl_tm, 'Functional Requirements')
            tm_sheet = get_sheet_with_fallback(xl_tp, 'Traceability Matrix')
            PRScol, PRSurs = get_clean_cols(prs_path, prs_sheet, 'PRS ID', 'URS ID')
            TMcol, TMurs = get_clean_cols(tm_path, tm_sheet, 'PRS Requirement ID', 'URS Requirement ID', skiprows=1)
            PRScol = normalize_spaces(filter_ignored(PRScol, IGNORE_SENTENCES_VAL_TP_TM))
            PRSurs = normalize_spaces(filter_ignored(PRSurs, IGNORE_SENTENCES_VAL_TP_TM))
            TMcol = normalize_spaces(filter_ignored(TMcol, IGNORE_SENTENCES_VAL_TP_TM))
            TMurs = normalize_spaces(filter_ignored(TMurs, IGNORE_SENTENCES_VAL_TP_TM))
            # Check if all TMcol are present in PRScol
            missing_prs = sorted(set(TMcol) - set(PRScol))
            # Check if all TMurs are present in PRSurs
            missing_urs = sorted(set(TMurs) - set(PRSurs))

            st.markdown("#### PRS IDs in TM not found in PRS DOC:")
            if missing_prs:
                st.dataframe(pd.DataFrame({'PRS in TM not in PRS DOC': missing_prs}), use_container_width=True)
            else:
                st.success("All PRS IDs from TM are present in PRS DOC.")

            st.markdown("#### URS IDs in TM not found in PRS DOC:")
            if missing_urs:
                st.dataframe(pd.DataFrame({'URS in TM not in PRS DOC': missing_urs}), use_container_width=True)
            else:
                st.success("All URS IDs from TM are present in PRS DOC.")

            # Optionally, allow download of missing items
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                if missing_prs:
                    pd.DataFrame({'PRS in TM not in PRS DOC': missing_prs}).to_excel(writer, sheet_name='Missing PRS', index=False)
                if missing_urs:
                    pd.DataFrame({'URS in TM not in PRS DOC': missing_urs}).to_excel(writer, sheet_name='Missing URS', index=False)
            buffer.seek(0)
            if missing_prs or missing_urs:
                st.download_button(
                    "⬇️ Download Missing IDs Excel",
                    data=buffer,
                    file_name="Missing_PRS_URS_in_PRS_DOC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No missing PRS or URS IDs to download.")