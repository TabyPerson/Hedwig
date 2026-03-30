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

def get_combined_df(filepath, sheets, col_tc, col_prs, col_desc, col_step, col_exp):
    xl = pd.ExcelFile(filepath)    
    df_list = []
    for sheet in sheets:
        if sheet in xl.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet)
            # Filtra apenas as colunas necessárias e remove linhas nulas
            cols = [col_tc, col_prs, col_desc, col_step, col_exp]
            missing = [c for c in cols if c not in df.columns]
            if missing:
                st.warning(f"Sheet '{sheet}' missing columns: {missing}")
                continue
            df = df[cols].dropna(subset=[col_tc, col_prs])
            df.columns = ['Test Case', 'PRS', 'Brief Description', 'Action / Step', 'Expected Result']
            df_list.append(df)
        else:
            st.warning(f"Sheet '{sheet}' not found in the uploaded file.")
    if df_list:
        return pd.concat(df_list, ignore_index=True)
    else:
        return pd.DataFrame(columns=['Test Case', 'PRS', 'Brief Description', 'Action / Step', 'Expected Result'])

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

def pad_lists_to_same_length(d):
    max_len = max(len(v) for v in d.values())
    for k in d:
        d[k] = list(d[k]) + [""] * (max_len - len(d[k]))
    return d

def run_comparison():    
    st.markdown("### Verification Test Protocol x TSVR Comparison")
    st.info("Upload your Verification Test Protocol and TSVR files to compare.")
    col1, col2 = st.columns(2)
    with col1:
        tp_file = st.file_uploader("Upload Verification Test Protocol (.xlsx)", type=["xlsx"], key="tp_ver_tsvr")
    with col2:
        tsvr_file = st.file_uploader("Upload TSVR (.xlsx)", type=["xlsx"], key="tsvr_ver_tsvr")
    if tp_file and tsvr_file:
        if st.button("🔍 Run Comparison", key="run_ver_tp_tsvr"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tp:
                tmp_tp.write(tp_file.read())
                tp_path = tmp_tp.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tsvr:
                tmp_tsvr.write(tsvr_file.read())
                tsvr_path = tmp_tsvr.name

            # Sheets to read
            protocol_sheets = ["Test Case Report - URS MD", "Test Case Report - URS NMD"]
            tsvr_sheets = ["Test Case Report - MD", "Test Case Report - NMD"]

            # Helper to read and combine columns from multiple sheets
            def get_combined_cols(filepath, sheets, col_tc, col_prs):
                xl = pd.ExcelFile(filepath)
                tc_list = []
                prs_list = []
                for sheet in sheets:
                    if sheet in xl.sheet_names:
                        df = pd.read_excel(filepath, sheet_name=sheet)
                        if col_tc in df.columns and col_prs in df.columns:
                            tc_list.append(df[col_tc].dropna().astype(str).str.strip())
                            prs_list.append(df[col_prs].dropna().astype(str).str.strip())
                        else:
                            st.warning(f"Sheet '{sheet}' does not contain required columns: '{col_tc}' , '{col_prs}'. Found: {df.columns.tolist()}")
                    else:
                        st.warning(f"Sheet '{sheet}' not found in the uploaded file.")
                tc_combined = pd.concat(tc_list).drop_duplicates() if tc_list else pd.Series(dtype=str)
                prs_combined = pd.concat(prs_list).drop_duplicates() if prs_list else pd.Series(dtype=str)
                return tc_combined, prs_combined

            # Get and clean columns
            TPcol, TPprs = get_combined_cols(tp_path, protocol_sheets, 'Test Case ', 'Traceability (PRS)')
            TSVRcol, TSVRprs = get_combined_cols(tsvr_path, tsvr_sheets, 'Manual Test Case ID', 'Requirement Coverage')
            TPcol = normalize_spaces(filter_ignored(TPcol))
            TPprs = normalize_spaces(filter_ignored(TPprs))
            TSVRcol = normalize_spaces(filter_ignored(TSVRcol))
            TSVRprs = normalize_spaces(filter_ignored(TSVRprs))

            # Results
            tc_results = {
                'Protocol Only': sorted(set(TPcol) - set(TSVRcol)),
                'TSVR Only': sorted(set(TSVRcol) - set(TPcol)),
                'Common to Both': sorted(set(TPcol) & set(TSVRcol))
            }
            prs_results = {
                'Protocol Only': sorted(set(TPprs) - set(TSVRprs)),
                'TSVR Only': sorted(set(TSVRprs) - set(TPprs)),
                'Common to Both': sorted(set(TPprs) & set(TSVRprs))
            }

            tc_results = pad_lists_to_same_length(tc_results)
            prs_results = pad_lists_to_same_length(prs_results)

            # DataFrames
            tc_df = pd.DataFrame(tc_results)
            prs_df = pd.DataFrame(prs_results)

            st.write("Manual Test Case ID Comparison")
            st.dataframe(tc_df, use_container_width=True)
            st.write("Requirement Coverage Comparison")
            st.dataframe(prs_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Test Case COMP', index=False)
                prs_df.to_excel(writer, sheet_name='Requirement Coverage COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VER_TP_TSVR_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )