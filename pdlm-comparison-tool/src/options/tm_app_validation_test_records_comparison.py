import streamlit as st
import pandas as pd
import tempfile
import io
import re

def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the correct sheet.")
        return st.selectbox("Select the correct worksheet:", xl.sheet_names, key=preferred_name) 

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
    'Note (N/A*): The columns “Actual Result /Description”, "Version tested", "Date tested", "Tester", "Result (Pass/Fail)", "Defect/Enhancement Number", and "Defect/Enhancement Status" from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.',
    'Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.',
    'Note: Defect/Enhancement are linked to test cases and their respective execution versions',
    'NA* Closed Service Orders'
] 

IGNORE_SENTENCES_VAL_TP_TM = IGNORE_SENTENCES + ['* Not applicable']

def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()).lower())

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def pad_list(lst, target_len):
    lst = lst if lst is not None else []
    if target_len is None or not isinstance(target_len, int) or target_len < 1:
        target_len = 1
    return lst + [""] * max(target_len - len(lst), 0)

def normalize_version(ver):
    # Remove .00 do final e espaços
    if isinstance(ver, str):
        ver = ver.strip()
        if ver.endswith('.00'):
            ver = ver[:-3]
    return ver.lower()
   
def run_comparison():
    st.markdown("### Validation TM APP x Validation Test Records Comparison")
    st.info("Upload your TM APP and Validation Test Records files to compare test records.")
    col1, col2 = st.columns(2)
    with col1:
        tm_file = st.file_uploader("Upload TM APP (.xlsx)", type=["xlsx"], key="tmapp_val_rec")
    with col2:
        tr_file = st.file_uploader("Upload Validation Test Records (.xlsx)", type=["xlsx"], key="tr_val_tm")
    if tm_file and tr_file:
        if st.button("🔍 Run Comparison", key="run_val_tr_tm"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tm:
                tmp_tm.write(tm_file.read())
                tm_path = tmp_tm.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name
            xl_tm = pd.ExcelFile(tm_path)
            xl_tr = pd.ExcelFile(tr_path)
            tm_sheet = get_sheet_with_fallback(xl_tm, 'Design Validation')
            tr_sheet = get_sheet_with_fallback(xl_tr, 'Test Case Report')
            df_tm = pd.read_excel(tm_path, sheet_name=tm_sheet, skiprows=1)
            df_tr = pd.read_excel(tr_path, sheet_name=tr_sheet)

            tm_validation_id_col = 'Validation Test ID'
            if tm_validation_id_col not in df_tm.columns:
                st.warning(f"Coluna '{tm_validation_id_col}' não encontrada. Por favor, selecione a coluna correta.")
                tm_validation_id_col = st.selectbox("Selecione a coluna que contém os IDs de validação na TM APP:", list(df_tm.columns))
            tm_urs_id_col = 'URS Requirement ID'
            if tm_urs_id_col not in df_tm.columns:
                st.warning(f"Coluna '{tm_urs_id_col}' não encontrada. Por favor, selecione a coluna correta.")
                tm_urs_id_col = st.selectbox("Selecione a coluna que contém os IDs de requisitos URS na TM APP:", list(df_tm.columns))
            tm_version_col = 'Validation Version'
            if tm_version_col not in df_tm.columns:
                st.warning(f"Coluna '{tm_version_col}' não encontrada. Por favor, selecione a coluna correta.")
                tm_version_col = st.selectbox("Selecione a coluna que contém a versão de validação na TM APP:", list(df_tm.columns))
            tm_result_col = 'Validation Test Result (Pass/Fail)'
            if tm_result_col not in df_tm.columns:
                st.warning(f"Coluna '{tm_result_col}' não encontrada. Por favor, selecione a coluna correta.")
                tm_result_col = st.selectbox("Selecione a coluna que contém os resultados de teste na TM APP:", list(df_tm.columns))

            tr_validation_id_col = 'Validation Protocol ID'
            if tr_validation_id_col not in df_tr.columns:
                st.warning(f"Coluna '{tr_validation_id_col}' não encontrada. Por favor, selecione a coluna correta.")
                tr_validation_id_col = st.selectbox("Selecione a coluna que contém os IDs de protocolo de validação nos Test Records:", list(df_tr.columns))
            tr_urs_id_col = 'URS'
            if tr_urs_id_col not in df_tr.columns:
                st.warning(f"Coluna '{tr_urs_id_col}' não encontrada. Por favor, selecione a coluna correta.")
                tr_urs_id_col = st.selectbox("Selecione a coluna que contém os IDs de requisitos URS nos Test Records:", list(df_tr.columns))
            tr_version_col = 'Version Tested'
            if tr_version_col not in df_tr.columns:
                potential_version_cols = [col for col in df_tr.columns if 'version' in col.lower()]
                if potential_version_cols:
                    st.warning(f"Coluna '{tr_version_col}' não encontrada, mas encontramos possíveis alternativas.")
                    tr_version_col = st.selectbox("Selecione a coluna que contém a versão testada nos Test Records:", potential_version_cols + [col for col in df_tr.columns if col not in potential_version_cols])
                else:
                    st.warning(f"Coluna '{tr_version_col}' não encontrada. Por favor, selecione a coluna correta.")
                    tr_version_col = st.selectbox("Selecione a coluna que contém a versão testada nos Test Records:", list(df_tr.columns))
            tr_result_col = 'Result (Pass/Fail)'
            if tr_result_col not in df_tr.columns:
                potential_result_cols = [col for col in df_tr.columns if 'result' in col.lower() or 'pass' in col.lower() or 'conclusion' in col.lower()]
                if potential_result_cols:
                    st.warning(f"Coluna '{tr_result_col}' não encontrada, mas encontramos possíveis alternativas.")
                    tr_result_col = st.selectbox("Selecione a coluna que contém os resultados de teste nos Test Records:", potential_result_cols + [col for col in df_tr.columns if col not in potential_result_cols])
                else:
                    st.warning(f"Coluna '{tr_result_col}' não encontrada. Por favor, selecione a coluna correta.")
                    tr_result_col = st.selectbox("Selecione a coluna que contém os resultados de teste nos Test Records:", list(df_tr.columns))

            # Comparação básica de IDs e URS
            TMcol = normalize_spaces(filter_ignored(df_tm[tm_validation_id_col].dropna().astype(str).str.strip(), IGNORE_SENTENCES_VAL_TP_TM))
            TMurs = normalize_spaces(filter_ignored(df_tm[tm_urs_id_col].dropna().astype(str).str.strip(), IGNORE_SENTENCES_VAL_TP_TM))
            TRcol = normalize_spaces(filter_ignored(df_tr[tr_validation_id_col].dropna().astype(str).str.strip(), IGNORE_SENTENCES_VAL_TP_TM))
            TRurs = normalize_spaces(filter_ignored(df_tr[tr_urs_id_col].dropna().astype(str).str.strip(), IGNORE_SENTENCES_VAL_TP_TM))
            TMcol = normalize_spaces(filter_ignored(TMcol, IGNORE_SENTENCES_VAL_TP_TM))
            TMurs = normalize_spaces(filter_ignored(TMurs, IGNORE_SENTENCES_VAL_TP_TM))
            TRcol = normalize_spaces(filter_ignored(TRcol, IGNORE_SENTENCES_VAL_TP_TM))
            TRurs = normalize_spaces(filter_ignored(TRurs, IGNORE_SENTENCES_VAL_TP_TM))
            tc_results = {
                'VAL Record Only in TM': sorted(set(TMcol) - set(TRcol)),
                'VAL Record Only in Record': sorted(set(TRcol) - set(TMcol)),
                'VAL Record Common to Both': sorted(set(TMcol) & set(TRcol))
            }
            urs_results = {
                'VAL URS Only in TM': sorted(set(TMurs) - set(TRurs)),
                'VAL URS Only in Record': sorted(set(TRurs) - set(TMurs)),
                'VAL URS Common to Both': sorted(set(TMurs) & set(TRurs))
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
                            result.append(text.upper())  # <-- MAIÚSCULA
                    except:
                        pass
                return result
            for result_dict in [tc_results, urs_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))
            tc_df = pd.DataFrame(tc_results)
            urs_df = pd.DataFrame(urs_results)
            st.write("Test Case Comparison")
            st.dataframe(tc_df, use_container_width=True)
            st.write("URS Comparison")
            st.dataframe(urs_df, use_container_width=True)
            non_functional_urs = [
                "A_0_6.1.598",
                "A_0_6.1.599",
                "A_0_6.10.601",
                "A_0_6.10.602",
                "A_0_6.10.603",
                "A_0_6.10.604"
            ]
            if 'VAL URS Only in TM' in urs_df.columns:
                if urs_df['VAL URS Only in TM'].isin([x.upper() for x in non_functional_urs]).any():
                    st.markdown(
                        """
                        <span style="color:black"><b>Note*:</b> The following 6 URSs listed abaixo do not require to be validated as they are non-functional requirements and are related to software infrastructure. The non-functional requirements are applicable for System Security and Recovery capabilities to the system and are designed in accordance with the Software product technical design for the Tasy EMR:</span>

                        - A_0_6.1.598
                        - A_0_6.1.599
                        - A_0_6.10.601
                        - A_0_6.10.602
                         - A_0_6.10.603
                        - A_0_6.10.604
                        """,
                        unsafe_allow_html=True
                    )
            # Comparação de versões e resultados
            st.write("### Comparação de Versões e Resultados por Validation Protocol/Test ID")
            df_tm_clean = df_tm.copy()
            df_tr_clean = df_tr.copy()
            df_tm_clean[tm_validation_id_col] = df_tm_clean[tm_validation_id_col].astype(str).str.strip()
            df_tr_clean[tr_validation_id_col] = df_tr_clean[tr_validation_id_col].astype(str).str.strip()
            tm_ignore_mask = df_tm_clean[tm_validation_id_col].str.lower().isin([s.strip().lower() for s in IGNORE_SENTENCES_VAL_TP_TM])
            tr_ignore_mask = df_tr_clean[tr_validation_id_col].str.lower().isin([s.strip().lower() for s in IGNORE_SENTENCES_VAL_TP_TM])
            df_tm_clean = df_tm_clean[~tm_ignore_mask]
            df_tr_clean = df_tr_clean[~tr_ignore_mask]

            # Sempre usar a coluna 'Round' para pegar a última execução
            round_col = 'Round' if 'Round' in df_tr_clean.columns else None

            def normalize_version(ver):
                if isinstance(ver, str):
                    ver = ver.strip()
                    if ver.endswith('.00'):
                        ver = ver[:-3]
                return ver.lower()

            version_diff_data = []
            result_diff_data = []
            common_ids = set(df_tm_clean[tm_validation_id_col]) & set(df_tr_clean[tr_validation_id_col])
            if not common_ids:
                st.warning("Não foram encontrados IDs comuns para comparar versões e resultados.")
            else:
                st.success(f"Encontrados {len(common_ids)} IDs comuns entre os dois arquivos.")
                for test_id in sorted(common_ids):
                    tm_rows = df_tm_clean[df_tm_clean[tm_validation_id_col] == test_id]
                    tr_rows = df_tr_clean[df_tr_clean[tr_validation_id_col] == test_id]
                    # Se tiver coluna de Round, pegar apenas o último round para cada teste
                    if round_col:
                        try:
                            valid_rounds = tr_rows[tr_rows[round_col].notna()]
                            if not valid_rounds.empty:
                                try:
                                    valid_rounds[round_col] = pd.to_numeric(valid_rounds[round_col], errors='coerce')
                                    for i, row in valid_rounds.iterrows():
                                        if pd.isna(row[round_col]):
                                            valid_rounds.loc[i, round_col] = tr_rows.loc[i, round_col]
                                except:
                                    pass
                                sorted_rounds = valid_rounds.sort_values(by=round_col, ascending=False)
                                if not sorted_rounds.empty:
                                    tr_rows = sorted_rounds.head(1)
                        except Exception as e:
                            st.error(f"Erro ao processar dados de round para o teste {test_id}: {str(e)}")
                    for _, tm_row in tm_rows.iterrows():
                        tm_version = str(tm_row[tm_version_col]).strip() if pd.notna(tm_row[tm_version_col]) else 'N/A'
                        tm_result = str(tm_row[tm_result_col]).strip() if pd.notna(tm_row[tm_result_col]) else 'N/A'
                        for _, tr_row in tr_rows.iterrows():
                            tr_version = str(tr_row[tr_version_col]).strip() if pd.notna(tr_row[tr_version_col]) else 'N/A'
                            tr_result = str(tr_row[tr_result_col]).strip() if pd.notna(tr_row[tr_result_col]) else 'N/A'
                            tm_version_norm = normalize_version(tm_version)
                            tr_version_norm = normalize_version(tr_version)
                            version_match = tm_version_norm == tr_version_norm
                            result_match = tm_result.lower() == tr_result.lower()
                            if not version_match:
                                version_diff_data.append({
                                    'Validation Protocol ID': str(test_id).upper(),
                                    'Version Tested': tr_version.upper(),
                                    'Validation Version': tm_version.upper()
                                })
                            if not result_match:
                                result_diff_data.append({
                                    'Validation Protocol ID': str(test_id).upper(),
                                    'Result (Pass/Fail)': tr_result.upper(),
                                    'Validation Test Result (Pass/Fail)': tm_result.upper()
                                })
            if version_diff_data:
                st.write("#### Diferenças de Versão")
                st.dataframe(pd.DataFrame(version_diff_data), use_container_width=True)
            else:
                st.success("Todas as versões coincidem!")
            if result_diff_data:
                st.write("#### Diferenças de Resultado")
                st.dataframe(pd.DataFrame(result_diff_data), use_container_width=True)
            else:
                st.success("Todos os Result (Pass/Fail) coincidem!")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Record COMP', index=False)
                urs_df.to_excel(writer, sheet_name='URS COMP', index=False)
                if version_diff_data:
                    pd.DataFrame(version_diff_data).to_excel(writer, sheet_name='Version_DIFF', index=False)
                if result_diff_data:
                    pd.DataFrame(result_diff_data).to_excel(writer, sheet_name='Result_DIFF', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VAL_TRxTM_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_val_comparison"
            )