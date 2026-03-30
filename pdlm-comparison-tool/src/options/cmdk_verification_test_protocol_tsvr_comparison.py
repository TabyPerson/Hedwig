import streamlit as st
import pandas as pd
import tempfile
import re
import io
from bs4 import BeautifulSoup

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

def extract_pairs_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    pairs = set()
    for th in soup.find_all('th', string=re.compile(r'Test Case ID:', re.IGNORECASE)):
        td_id = th.find_next_sibling('td')
        td_req = None
        next_td = td_id.find_next_sibling('td') if td_id else None
        if next_td and next_td.find('b', string=re.compile(r'Requirements:', re.IGNORECASE)):
            td_req = next_td
        if td_id and td_req:
            tc_id_vals = [v.strip() for v in re.split(r'[;,]', td_id.get_text()) if v.strip()]
            req_b = td_req.find('b', string=re.compile(r'Requirements:', re.IGNORECASE))
            req_texts = []
            found_b = False
            for elem in td_req.contents:
                if elem == req_b:
                    found_b = True
                elif found_b:
                    if isinstance(elem, str):
                        req_texts.append(elem)
                    else:
                        req_texts.append(elem.get_text())
            req_vals = [v.strip() for v in re.split(r'[;,]', ''.join(req_texts)) if v.strip()]
            for tc_id in tc_id_vals:
                for req in req_vals:
                    pairs.add((tc_id, req))
    return pairs

def run_comparison():
    st.markdown("### TSVR x Verification Test Protocol Comparison (HTML x HTML)")
    st.info("Upload your TSVR (.html) and Verification Test Protocol (.html) to compare Test Case IDs and Requirements.")
    col1, col2 = st.columns(2)
    with col1:
        tsvr_file = st.file_uploader("Upload TSVR (.html)", type=["html"], key="tsvr_html")
    with col2:
        protocol_file = st.file_uploader("Upload Verification Test Protocol (.html)", type=["html"], key="protocol_html")
    if tsvr_file and protocol_file:
        if st.button("🔍 Run Comparison", key="run_html_html"):
            tsvr_html = tsvr_file.read().decode('utf-8', errors='ignore')
            protocol_html = protocol_file.read().decode('utf-8', errors='ignore')

            tsvr_pairs = extract_pairs_from_html(tsvr_html)
            protocol_pairs = extract_pairs_from_html(protocol_html)

            only_in_tsvr = sorted(tsvr_pairs - protocol_pairs)
            only_in_protocol = sorted(protocol_pairs - tsvr_pairs)
            in_both = sorted(tsvr_pairs & protocol_pairs)

            def pairs_to_df(pairs, col1, col2):
                df = pd.DataFrame(pairs, columns=[col1, col2]) if pairs else pd.DataFrame({col1: [''], col2: ['']})
                return df.drop_duplicates().reset_index(drop=True)

            df_only_in_tsvr = pairs_to_df(only_in_tsvr, 'Test Case ID', 'PRS Requirement ID')
            df_only_in_protocol = pairs_to_df(only_in_protocol, 'Test Case ID', 'PRS Requirement ID')
            df_in_both = pairs_to_df(in_both, 'Test Case ID', 'PRS Requirement ID')

            st.write('---')
            st.write('### Pares combinados Test Case ID x PRS Requirement ID')
            st.write('**Pares apenas no TSVR:**')
            st.dataframe(df_only_in_tsvr, use_container_width=True)
            st.write('**Pares apenas no Protocol:**')
            st.dataframe(df_only_in_protocol, use_container_width=True)
            st.write('**Pares em ambos:**')
            st.dataframe(df_in_both, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                df_only_in_tsvr.to_excel(writer, sheet_name='TSVR Only', index=False)
                df_only_in_protocol.to_excel(writer, sheet_name='Protocol Only', index=False)
                df_in_both.to_excel(writer, sheet_name='Common', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="TSVR_vs_Protocol_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
