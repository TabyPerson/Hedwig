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
    
def get_all_prs_ids(prs_path, sheet_names, col_name):
    xl = pd.ExcelFile(prs_path)
    all_ids = []
    for sheet in sheet_names:
        if sheet in xl.sheet_names:
            df = pd.read_excel(prs_path, sheet_name=sheet)
            if col_name in df.columns:
                # Normaliza espaços e remove caracteres invisíveis
                ids = df[col_name].dropna().astype(str).apply(lambda x: re.sub(r'\s+', '', x).strip())
                all_ids.extend(ids)
    return pd.Series(all_ids)

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

            prs_sheets = [
                'Functional Requirements',
                'Interface Requirements',
                'Distribution and Installation',
                'Labeling and Learning Materials'
            ]
            prs_ids = get_all_prs_ids(prs_path, prs_sheets, 'ID')
            prs_ids = prs_ids.apply(lambda x: re.sub(r'\s+', '', str(x)).strip())
            prs_ids = filter_ignored(prs_ids, IGNORE_SENTENCES_VAL_TP_TM)

            xl_tm = pd.ExcelFile(tm_path)
            tm_sheet = get_sheet_with_fallback(xl_tm, 'Traceability-Verification')
            tm_df = pd.read_excel(tm_path, sheet_name=tm_sheet, skiprows=1)
            if 'PRS Requirement ID' not in tm_df.columns:
                st.error(f"Column 'PRS Requirement ID' not found in sheet '{tm_sheet}'. Available columns: {list(tm_df.columns)}")
                st.stop()
            tm_ids = tm_df['PRS Requirement ID'].dropna().astype(str).apply(lambda x: re.sub(r'\s+', '', x).strip())
            tm_ids = filter_ignored(tm_ids, IGNORE_SENTENCES_VAL_TP_TM)

            prs_results = {
                'PRS Only in PRS DOC': sorted(set(prs_ids) - set(tm_ids)),
                'PRS Only in TM': sorted(set(tm_ids) - set(prs_ids)),
                'PRS Common to Both': sorted(set(prs_ids) & set(tm_ids))
            }

            def force_to_strings(lst):
                result = []
                for item in lst:
                    try:
                        text = str(item)
                        if (not "method" in text and 
                            not "descriptor" in text and
                            not text.startswith("<") and
                            not text.startswith("[") and
                            not text.startswith("(") and
                            text.strip() != ""):
                            result.append(text)
                    except:
                        pass
                return result

            max_len = max(len(force_to_strings(v)) for v in prs_results.values())
            for k in prs_results:
                clean_list = force_to_strings(prs_results[k])
                if not clean_list:
                    clean_list = [""]
                prs_results[k] = clean_list + [""] * (max_len - len(clean_list))

            prs_df = pd.DataFrame(prs_results)
            st.write("PRS Comparison")
            st.dataframe(prs_df, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                prs_df.to_excel(writer, sheet_name='PRS COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VER_PRSxTM_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )