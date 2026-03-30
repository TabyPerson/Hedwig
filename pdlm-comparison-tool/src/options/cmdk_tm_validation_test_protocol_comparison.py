import streamlit as st
import pandas as pd
import tempfile
import io
import re

def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the sheet that you want to analyze.")
        return st.selectbox("Select the worksheet that you want to analyze:", xl.sheet_names, key=preferred_name) 

def get_clean_cols(filepath, sheet, col_tc, col_urs=None, skiprows=0):
    xl = pd.ExcelFile(filepath)
    if sheet not in xl.sheet_names:
        st.error(f"Worksheet named '{sheet}' not found in the uploaded file. Available sheets: {xl.sheet_names}")
        st.stop()
    df = pd.read_excel(filepath, sheet_name=sheet, skiprows=skiprows)
    if col_tc not in df.columns:
        st.error(f"Column '{col_tc}' not found in sheet '{sheet}'. Available columns: {list(df.columns)}")
        st.stop()
    tc = df[col_tc].dropna().astype(str).str.strip()
    if col_urs:
        if col_urs not in df.columns:
            st.error(f"Column '{col_urs}' not found in sheet '{sheet}'. Available columns: {list(df.columns)}")
            st.stop()
        urs = df[col_urs].dropna().astype(str).str.strip()
        return tc, urs
    return tc 

IGNORE_SENTENCES = [
    'Note (N/A*): The columns “Actual Result /Description”, "Version tested", "Date tested", "Tester", "Conclusion (Pass/Fail)", "Defect/Enhancement Number", and "Defect/Enhancement Status" from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.',
    'Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.',
    'Note: Defect/Enhancement are linked to test cases and their respective execution versions',
    'NA* Closed Service Orders'
] 

IGNORE_SENTENCES_VAL_TP_TM = IGNORE_SENTENCES + ['* Not applicable']

def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def pad_list(lst, target_len):
    lst = lst if lst is not None else []
    if target_len is None or not isinstance(target_len, int) or target_len < 1:
        target_len = 1
    return lst + [""] * max(target_len - len(lst), 0)
   
def run_comparison():
    st.markdown("### CMDK TM x Validation Test Protocol Comparison")
    st.info("Upload your TM and Validation Test Protocol files to compare test protocols.")
    col1, col2 = st.columns(2)
    with col1:
        tm_file = st.file_uploader("Upload TM (.xlsx)", type=["xlsx"], key="tmapp_val")
    with col2:
        tp_file = st.file_uploader("Upload Validation Test Protocol (.xlsx)", type=["xlsx"], key="tp_val")
    if tm_file and tp_file:
        if st.button("🔍 Run Comparison", key="run_val_tp_tm"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tm:
                tmp_tm.write(tm_file.read())
                tm_path = tmp_tm.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tp:
                tmp_tp.write(tp_file.read())
                tp_path = tmp_tp.name
            xl_tm = pd.ExcelFile(tm_path)
            xl_tp = pd.ExcelFile(tp_path)
            tm_sheet = get_sheet_with_fallback(xl_tm, 'Traceability-Validation')
            tp_sheet = get_sheet_with_fallback(xl_tp, 'Test Case Protocol')
            TMcol, TMurs = get_clean_cols(tm_path, tm_sheet, 'Validation Test ID', 'URS Requirement ID', skiprows=1)
            TPcol, TPurs = get_clean_cols(tp_path, tp_sheet, 'Test case ID ', 'URS ID ')
            TMcol = normalize_spaces(filter_ignored(TMcol, IGNORE_SENTENCES_VAL_TP_TM))
            TMurs = normalize_spaces(filter_ignored(TMurs, IGNORE_SENTENCES_VAL_TP_TM))
            TPcol = normalize_spaces(filter_ignored(TPcol, IGNORE_SENTENCES_VAL_TP_TM))
            TPurs = normalize_spaces(filter_ignored(TPurs, IGNORE_SENTENCES_VAL_TP_TM))
            tc_results = {
                'VAL Protocol Only in TM': sorted(set(TMcol) - set(TPcol)),
                'VAL Protocol Only in Protocol': sorted(set(TPcol) - set(TMcol)),
                'VAL Protocol Common to Both': sorted(set(TMcol) & set(TPcol))
            }
            urs_results = {
                'VAL URS Only in TM': sorted(set(TMurs) - set(TPurs)),
                'VAL URS Only in Protocol': sorted(set(TPurs) - set(TMurs)),
                'VAL URS Common to Both': sorted(set(TMurs) & set(TPurs))
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
            for result_dict in [tc_results, urs_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                # Preencher até o tamanho máximo
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))

            # Agora pode criar os DataFrames sem risco de erro
            tc_df = pd.DataFrame(tc_results)
            urs_df = pd.DataFrame(urs_results)
            st.write("Test Case Comparison")
            st.dataframe(tc_df, use_container_width=True)
            st.write("URS Comparison")
            st.dataframe(urs_df, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Protocol COMP', index=False)
                urs_df.to_excel(writer, sheet_name='URS COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="CMDK_VAL_TPxTM_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )