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
    st.markdown("### Verification TM APP x Test Protocol Comparison")
    st.info("Upload your TM APP and Verification Test Protocol files to compare test protocols.")
    col1, col2 = st.columns(2)
    with col1:
        tm_file = st.file_uploader("Upload TM APP (.xlsx)", type=["xlsx"], key="tmapp_ver2")
    with col2:
        tp_file = st.file_uploader("Upload Verification Test Protocol (.xlsx)", type=["xlsx"], key="tp_ver2")
    if tm_file and tp_file:
        if st.button("🔍 Run Comparison", key="run_ver_tp_tm"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tm:
                tmp_tm.write(tm_file.read())
                tm_path = tmp_tm.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tp:
                tmp_tp.write(tp_file.read())
                tp_path = tmp_tp.name

            # Read TM APP
            xl_tm = pd.ExcelFile(tm_path)
            tm_sheet = get_sheet_with_fallback(xl_tm, 'Design Verification')
            TMcol, TMprs = get_clean_cols(tm_path, tm_sheet, 'Verification Test ID', 'PRS Requirement ID', skiprows=1)
            TMcol = normalize_spaces(filter_ignored(TMcol, IGNORE_SENTENCES_VAL_TP_TM))
            TMprs = normalize_spaces(filter_ignored(TMprs, IGNORE_SENTENCES_VAL_TP_TM))

            # Read both sheets from Verification Test Protocol
            xl_tp = pd.ExcelFile(tp_path)
            sheets_to_read = ["Test Case Report - URS MD", "Test Case Report - URS NMD"]
            test_case_list = []
            traceability_list = []
            for sheet in sheets_to_read:
                if sheet in xl_tp.sheet_names:
                    df = pd.read_excel(tp_path, sheet_name=sheet)
                    if "Test Case " in df.columns and "Traceability (PRS)" in df.columns:
                        test_case_list.append(df["Test Case "].dropna().astype(str).str.strip())
                        traceability_list.append(df["Traceability (PRS)"].dropna().astype(str).str.strip())
                    else:
                        st.warning(f"Sheet '{sheet}' does not contain required columns 'Test Case ' and 'Traceability (PRS)'. Found: {df.columns.tolist()}")
                else:
                    st.warning(f"Sheet '{sheet}' not found in the uploaded Verification Test Protocol file.")

                # Combine and clean
            TPcol = pd.concat(test_case_list).drop_duplicates()
            TPprs = pd.concat(traceability_list).drop_duplicates()
            TPcol = normalize_spaces(filter_ignored(TPcol, IGNORE_SENTENCES_VAL_TP_TM))
            TPprs = normalize_spaces(filter_ignored(TPprs, IGNORE_SENTENCES_VAL_TP_TM))

            # Test Case comparison
            tc_results = {
                'Verification Test Case Only in Traceability Application': sorted(set(TMcol) - set(TPcol)),
                'Verification Test Case Only in Verification Test Protocol': sorted(set(TPcol) - set(TMcol)),
                'Verification Test Case Common to Both Documents': sorted(set(TMcol) & set(TPcol))
            }
            # PRS comparison
            prs_results = {
                'Product Requirements Only in Traceability Application': sorted(set(TMprs) - set(TPprs)),
                'Product Requirements Only in Verification Test Protocol': sorted(set(TPprs) - set(TMprs)),
                'Product Requirements Common to Both Documents': sorted(set(TMprs) & set(TPprs))
            }
            # Função para garantir apenas strings válidas e nunca None
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
            # Para cada resultado, aplique o filtro e preencha até o tamanho máximo
            for result_dict in [prs_results, tc_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                # Preencher até o tamanho máximo
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))

            # Agora pode criar os DataFrames sem risco de erro
            prs_df = pd.DataFrame(prs_results)
            tc_df = pd.DataFrame(tc_results)
            st.write("PRS Comparison")
            st.dataframe(prs_df, use_container_width=True)
            st.write("Test Case Comparison")
            st.dataframe(tc_df, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Test Case COMP', index=False)
            prs_df.to_excel(writer, sheet_name='PRS COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VER_TPxTMAPP_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )