# File: /pdlm-comparison-tool/pdlm-comparison-tool/src/options/val_records_pdsr_comparison.py

import streamlit as st
import pandas as pd
import tempfile
import io
import re


def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))



def get_clean_cols(filepath, sheet, col_defect, skiprows=0):
    xl = pd.ExcelFile(filepath)
    if sheet not in xl.sheet_names:
        st.error(f"Worksheet named '{sheet}' not found in the uploaded file. Available sheets: {xl.sheet_names}")
        st.stop()
    df = pd.read_excel(filepath, sheet_name=sheet, skiprows=skiprows)
    if col_defect not in df.columns:
        st.error(f"Column '{col_defect}' not found in sheet '{sheet}'. Available columns: {list(df.columns)}")
        st.stop()
    return df[col_defect].dropna().astype(str).str.strip()

def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the correct sheet.")
        return st.selectbox("Select the correct worksheet:", xl.sheet_names, key=preferred_name) 
    
IGNORE_SENTENCES = [
    'Note (N/A*): The columns “Actual Result /Description”, "Version tested", "Date tested", "Tester", "Conclusion (Pass/Fail)", "Defect/Enhancement Number", and "Defect/Enhancement Status" from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.',
    'Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.',
    'Note: Defect/Enhancement are linked to test cases and their respective execution versions',
    'NA* Closed Service Orders'
]

IGNORE_SENTENCES_VAL_TP_TM = IGNORE_SENTENCES + ['* Not applicable']
IGNORE_SENTENCES_VAL_TR_PDSR = IGNORE_SENTENCES + ['* No Defect/Enhancement raised']

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES_VAL_TR_PDSR):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def pad_list(lst, target_len):
    lst = lst if lst is not None else []
    if target_len is None or not isinstance(target_len, int) or target_len < 1:
        target_len = 1
    return lst + [""] * max(target_len - len(lst), 0)

def run_comparison():
    st.markdown("### Validation Test Records x PDSR Comparison")
    st.info("Upload your Validation Test Records and Product Defect Status Report files to compare.")
    col1, col2 = st.columns(2)
    with col1:
        tr_file = st.file_uploader("Upload Validation Test Records (.xlsx)", type=["xlsx"], key="tr_val_pdsr")
    with col2:
        pdsr_file = st.file_uploader("Upload Product Defect Status Report (.xlsx)", type=["xlsx"], key="pdsr_val")
    if tr_file and pdsr_file:
        if st.button("🔍 Run Comparison", key="run_val_tr_pdsr"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_pdsr:
                tmp_pdsr.write(pdsr_file.read())
                pdsr_path = tmp_pdsr.name
            xl_tr = pd.ExcelFile(tr_path)
            xl_pdsr = pd.ExcelFile(pdsr_path)
            tr_sheet = get_sheet_with_fallback(xl_tr, 'Test Case Report')
            # Find the sheet with 'Defect ID' column, skipping 2 rows
            pdsr_sheet = None
            for sheet in xl_pdsr.sheet_names:
                df_tmp = pd.read_excel(pdsr_path, sheet_name=sheet, skiprows=2)
                if 'Defect ID' in df_tmp.columns:
                    pdsr_sheet = sheet
                    break
            if pdsr_sheet is None:
                st.error("No sheet with 'Defect ID' column found in the PDSR file.")
                st.stop()
            TRcol = get_clean_cols(tr_path, tr_sheet, 'Defect / Enhancement Number')
            PDSRcol = get_clean_cols(pdsr_path, pdsr_sheet, 'Defect ID', skiprows=2)
            TRcol = normalize_spaces(filter_ignored(TRcol, IGNORE_SENTENCES_VAL_TR_PDSR))
            PDSRcol = normalize_spaces(filter_ignored(PDSRcol, IGNORE_SENTENCES_VAL_TR_PDSR))
            df_results = {            
                'Defects Only in Records': sorted(set(TRcol) - set(PDSRcol)),
                'Defects Only in PDSR': sorted(set(PDSRcol) - set(TRcol)),
                'Defectscommon to Both': sorted(set(TRcol) & set(PDSRcol))
            }

            # Versão mais robusta para garantir apenas strings
            def force_to_strings(lst):
                result = []
                for item in lst:
                    try:
                        # Converter para string e verificar se é uma string válida
                        text = str(item)
                        # Filtrar strings que parecem objetos ou métodos
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
            for result_dict in [df_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                # Preencher até o tamanho máximo
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))

            # Agora pode criar os DataFrames sem risco de erro
            df_df = pd.DataFrame(df_results)
            st.write("Defect Comparison")
            st.dataframe(df_results, use_container_width=True)   
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                df_df.to_excel(writer, sheet_name='Defeito COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Result as Excel",
                data=buffer,
                file_name="VAL_TRxPDSR_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )