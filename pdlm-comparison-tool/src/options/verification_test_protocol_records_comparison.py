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
    'Note *: From Version Release 5.01.1835.00 onward all information related to the release version will consist of the 4 positions of the version ID (X.YY.ZZZZ.AAAA) intesd of 2 positions as it was used (X.YY)',
    'Note (N/A*): The columns “Actual Result /Description”, \'Version tested\', \'Date tested\', \'Tester\', \'Conclusion (Pass/Fail)\', \'Defect/Enhancement Number\', and \'Defect/Enhancement Status\' from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.'
] 

IGNORE_SENTENCES_VAL_TP_TM = IGNORE_SENTENCES + ['* Not applicable']

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def normalize_spaces(series):
    return pd.Series(series).astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def run_comparison():    
    st.markdown("### Verification Test Protocol x Records Comparison")
    st.info("Upload your Verification Test Protocol and Verification Test Records files to compare.")
    col1, col2 = st.columns(2)
    with col1:
        tp_file = st.file_uploader("Upload Verification Test Protocol (.xlsx)", type=["xlsx"], key="tp_ver_rec")
    with col2:
        tr_file = st.file_uploader("Upload Verification Test Records (.xlsx)", type=["xlsx"], key="tr_ver_rec")
    if tp_file and tr_file:
        if st.button("🔍 Run Comparison", key="run_ver_tp_rec"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tp:
                tmp_tp.write(tp_file.read())
                tp_path = tmp_tp.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name

            # Sheets to read
            protocol_sheets = ["Test Case Report - URS MD", "Test Case Report - URS NMD"]
            records_sheets = ["Test Case Report - URS MD", "Test Case Report - URS NMD", "Test Case Report - RMM"]

            # Helper to read and combine columns from multiple sheets
            def get_combined_cols(filepath, sheets, col_tc, col_prs, col_desc, col_step, col_exp):
                xl = pd.ExcelFile(filepath)
                tc_list = []
                prs_list = []
                desc_list = []
                step_list = []
                exp_list = []
                for sheet in sheets:
                    if sheet in xl.sheet_names:
                        df = pd.read_excel(filepath, sheet_name=sheet)
                        if col_tc in df.columns and col_prs in df.columns:
                            tc_list.append(df[col_tc].dropna().astype(str).str.strip())
                            prs_list.append(df[col_prs].dropna().astype(str).str.strip())
                            desc_list.append(df[col_tc].dropna().astype(str).str.strip())
                            step_list.append(df[col_prs].dropna().astype(str).str.strip())
                            exp_list.append(df[col_tc].dropna().astype(str).str.strip())
                        else:
                            st.warning(f"Sheet '{sheet}' does not contain required columns: '{col_tc}' , '{col_prs}' , '{col_desc}' , '{col_step}' and '{col_exp}'. Found: {df.columns.tolist()}")
                    else:
                        st.warning(f"Sheet '{sheet}' not found in the uploaded file.")
                tc_combined = pd.concat(tc_list).drop_duplicates() if tc_list else pd.Series(dtype=str)
                prs_combined = pd.concat(prs_list).drop_duplicates() if prs_list else pd.Series(dtype=str)
                desc_combined = pd.concat(desc_list).drop_duplicates() if desc_list else pd.Series(dtype=str)
                step_combined = pd.concat(step_list).drop_duplicates() if step_list else pd.Series(dtype=str)
                exp_combined = pd.concat(exp_list).drop_duplicates() if exp_list else pd.Series(dtype=str)
                return tc_combined, prs_combined, desc_combined, step_combined, exp_combined                
              
            # Get and clean columns
            TPcol, TPprs,TPdesc,TPstep,TPexp = get_combined_cols(tp_path, protocol_sheets, 'Test Case ', 'Traceability (PRS)', 'Brief Description', 'Action / Step (Description)', 'Expected Result (Description)')
            TRcol, TRprs,TRdesc,TRstep,TRexp = get_combined_cols(tr_path, records_sheets, 'Test Case ', 'Traceability (PRS)', 'Brief Description', 'Action / Step (Description)', 'Expected Result (Description)')
            TPcol = normalize_spaces(filter_ignored(TPcol))
            TPprs = normalize_spaces(filter_ignored(TPprs))
            TRcol = normalize_spaces(filter_ignored(TRcol))
            TRprs = normalize_spaces(filter_ignored(TRprs))
            TPdesc = normalize_spaces(filter_ignored(TPdesc))
            TRdesc = normalize_spaces(filter_ignored(TRdesc))
            TPstep = normalize_spaces(filter_ignored(TPstep))
            TRstep = normalize_spaces(filter_ignored(TRstep))
            TPexp = normalize_spaces(filter_ignored(TPexp))
            TRexp = normalize_spaces(filter_ignored(TRexp))
            # Monte os DataFrames completos
            protocol_df = get_combined_df(
                tp_path, protocol_sheets,
                'Test Case ', 'Traceability (PRS)', 'Brief Description', 'Action / Step (Description)', 'Expected Result (Description)'
            )
            records_df = get_combined_df(
                tr_path, records_sheets,
                'Test Case ', 'Traceability (PRS)', 'Brief Description', 'Action / Step (Description)', 'Expected Result (Description)'
            )
            # --- EXTRA CHECK: Test Case with inconsistent Conclusion/Defect Status ---
            def get_status_mismatches(filepath, sheets):
                xl = pd.ExcelFile(filepath)
                mismatch_rows = []
                for sheet in sheets:
                    if sheet in xl.sheet_names:
                        df = pd.read_excel(filepath, sheet_name=sheet)                        
                        df.columns = df.columns.astype(str).str.strip()
                         # Só tenta mostrar se for DataFrame, tiver colunas e pelo menos uma linha
                        required_cols = ['Test Case', 'Conclusion (Pass / Fail)', 'Defect/Enhancement Status']
                        if all(col in df.columns for col in required_cols):
                            # processamento normal
                            df['Test Case'] = df['Test Case'].astype(str).str.strip()
                            df['Conclusion (Pass / Fail)'] = df['Conclusion (Pass / Fail)'].astype(str).str.strip().str.lower()
                            df['Defect/Enhancement Status'] = df['Defect/Enhancement Status'].astype(str).str.strip().str.lower()
                            # Filtro direto, sem agrupar
                            mask_passed_open = (df['Conclusion (Pass / Fail)'] == 'passed') & (df['Defect/Enhancement Status'] == 'open')
                            mask_failed_closed = (df['Conclusion (Pass / Fail)'] == 'failed') & (df['Defect/Enhancement Status'] == 'closed')
                            mask_failed_no_defect = (df['Conclusion (Pass / Fail)'] == 'failed') & (
                                df['Defect/Enhancement Status'].str.contains('no defect', case=False, na=False)
                            )
                            mismatches = df[mask_passed_open | mask_failed_closed | mask_failed_no_defect][
                                ['Test Case', 'Conclusion (Pass / Fail)', 'Defect/Enhancement Status']
                            ]
                            mismatch_rows.append(mismatches)                        
                if mismatch_rows:
                    return pd.concat(mismatch_rows, ignore_index=True)
                else:
                    return pd.DataFrame()
            # Adicione após records_df ser criado
            mismatch_df = get_status_mismatches(tr_path, ["Test Case Report - URS MD", "Test Case Report - URS NMD", "Test Case Report - RMM"])
            if not mismatch_df.empty:
                st.markdown("### Test Cases with inconsistent Conclusion/Defect Status")
                st.dataframe(mismatch_df, use_container_width=True)
            else:
                st.info("No inconsistent Conclusion/Defect Status found in the selected sheets.")

            tc_results = {
                'VER Protocol Only in Protocol': sorted(set(TPcol) - set(TRcol)),
                'VER Protocol Only in Records': sorted(set(TRcol) - set(TPcol)),
                'VER Protocol Common to Both': sorted(set(TPcol) & set(TRcol))
            }
            prs_results = {
                'VER PRS Only in Protocol': sorted(set(TPprs) - set(TRprs)),
                'VER PRS Only in Records': sorted(set(TRprs) - set(TPprs)),
                'VER PRS Common to Both': sorted(set(TPprs) & set(TRprs))
            }
            desc_results = {
                'Brief Description Only in Protocol': sorted(set(TPdesc) - set(TRdesc)),
                'Brief Description Only in Records': sorted(set(TRdesc) - set(TPdesc)),
                'Brief Description Common to Both': sorted(set(TPdesc) & set(TRdesc))
            }            
            step_results = {
                'Action / Step Only in Protocol': sorted(set(TPstep) - set(TRstep)),
                'Action / Step Only in Records': sorted(set(TRstep) - set(TPstep)),
                'Action / Step Common to Both': sorted(set(TPstep) & set(TRstep))
            }
            exp_results = {
                'Expected Result Only in Protocol': sorted(set(TPexp) - set(TRexp)),
                'Expected Result Only in Records': sorted(set(TRexp) - set(TPexp)),
                'Expected Result Common to Both': sorted(set(TPexp) & set(TRexp))
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
            
            revision_mismatches = pd.DataFrame()               
            # --- Revision History Section ---
            st.markdown("## Revision History")
            # --- REVISION HISTORY TABLE 1 ---
            #PRSs only in protocol
            prs_only_in_protocol = force_to_strings(sorted(set(TPprs) - set(TRprs)))
            shown_tc_1 = set()
            shown_prs_1 = set()
            if prs_only_in_protocol:
                prs_to_tc = {}
                for prs in force_to_strings(prs_only_in_protocol):
                    linked_tcs = [tc for tc, p in zip(TPcol, TPprs) if p == prs]
                    linked_tcs = force_to_strings(linked_tcs)
                    prs_to_tc[prs] = linked_tcs
                    shown_tc_1.update(linked_tcs)
                    shown_prs_1.add(prs)

                table_data_1 = []
                for prs, tcs in prs_to_tc.items():
                    prs_str = str(prs) if prs is not None else ''
                    for tc in tcs if tcs is not None else ['']:
                        tc_str = str(tc) if tc is not None else ''
                        table_data_1.append({
                            'PRS Only in Protocol': prs_str,
                            'Test Case Linked': tc_str
                    })      
                        
                table_df_1 = pd.DataFrame(table_data_1) 
                        
                # Garante que o DataFrame está pronto para exibição
                table_df_1 = table_df_1.fillna('')  # Substitui None por string vazia

                total_tc = len(TPcol)
                total_prs = len(prs_only_in_protocol)
                if 'Test Case Linked' in table_df_1.columns:
                    total_tc_linked = len(table_df_1['Test Case Linked'].dropna().unique())
                else:
                    total_tc_linked = 0

                st.markdown(
                    f"""Out of the total of <b>{total_tc}</b> test cases listed in version XX.XX.XXXX.XX in the Product Verification Test Protocol [REF-X] for Tasy EMR, <b>{total_tc_linked}</b> test case(s) (listed below) linked to <b>{total_prs}</b> PRSs were removed from the Product Verification Test Records [REF-X] for Tasy EMR as they were linked to requirements changed in the revision X of the Product Requirement [REF-X].""",
                    unsafe_allow_html=True
                )
                # Exibir o DataFrame apenas depois do texto explicativo
                if table_df_1 is not None and not table_df_1.empty:
                    st.dataframe(table_df_1, use_container_width=True)
                else:
                    st.write("DataFrame is empty or None")
                
            # --- REVISION HISTORY TABLE 2 ---
            #Teste Cases only in protocol (excluding those linked to PRSs only in protocol)
            protocols_only_in_protocol = force_to_strings(sorted(set(TPcol) - set(TRcol)))
            protocols_only_in_protocol = [p for p in protocols_only_in_protocol if p not in shown_tc_1]

            protocol_to_prs = {}
            if protocols_only_in_protocol:
                for prot in force_to_strings(protocols_only_in_protocol):
                    linked_prs = [p for tc, p in zip(TPcol, TPprs) if tc == prot]
                    linked_prs = [p for p in linked_prs if p not in shown_prs_1]
                    linked_prs = force_to_strings(linked_prs)
                    if linked_prs:
                        protocol_to_prs[prot] = linked_prs

                table_data_2 = []
                for prot, prs_list in protocol_to_prs.items():
                    for prs in prs_list:
                        table_data_2.append({'PRS Linked': prs, 'Test Case ID': prot})
                        
                # Garante pelo menos uma linha
                if not table_data_2:
                    table_data_2 = [{'PRS Linked': '', 'Test Case ID': ''}]

                table_df_2 = pd.DataFrame(table_data_2)
                total_protocols = len(TPcol)
                total_protocols_only = len(protocol_to_prs)
                total_prs_linked = len(set([row['PRS Linked'] for row in table_data_2]))

                st.markdown(
                    f"""Out of the total of <b>{total_protocols}</b> Test cases listed in the Product Verification Test Protocol [REF-X] for Tasy EMR, <b>{total_protocols_only}</b> test cases linked to <b>{total_prs_linked}</b> PRSs listed below and in the Product Verification Test Records [REF-x] were not verified in release XX.XX.XXXX.XX. This occurred because the test cases were not part of the scope of this release. The processes is covered in other test cases linked to the PRSs.""",
                    unsafe_allow_html=True
                )
                # Exibir o texto explicativo antes do DataFrame
                st.dataframe(table_df_2, use_container_width=True)
                  
            # --- REVISION HISTORY TABLE 3 ---
            # --- Added in the test scope of the release (PRS only in records) ---
            prs_only_in_records = force_to_strings(sorted(set(TRprs) - set(TPprs)))
            shown_tc_2 = set()
            shown_prs_2 = set()
            if prs_only_in_records:
                prs_to_tc = {}
                for prs in force_to_strings(prs_only_in_records):
                    linked_tcs = [tc for tc, p in zip(TRcol, TRprs) if p == prs]
                    linked_tcs = force_to_strings(linked_tcs)
                    prs_to_tc[prs] = linked_tcs
                    shown_tc_2.update(linked_tcs)
                    shown_prs_2.add(prs)

                table_data_3 = []
                for prs, tcs in prs_to_tc.items():
                    prs_str = str(prs) if prs is not None else ''
                    for tc in tcs if tcs is not None else ['']:
                        tc_str = str(tc) if tc is not None else ''
                        table_data_3.append({
                            'PRS Only in Records': prs_str,
                            'Test Case Linked': tc_str
                    })      
                        
                table_df_3 = pd.DataFrame(table_data_3) 
                        
                # Garante que o DataFrame está pronto para exibição
                table_df_3 = table_df_3.fillna('')  # Substitui None por string vazia

                # Exibir o dataframe com parâmetros explícitos
                if table_df_3 is not None and not table_df_3.empty:
                    st.dataframe(table_df_3, use_container_width=True)
                else:
                    st.write("DataFrame is empty or None")
                    
                total_prs = len(prs_only_in_records)
                if 'Test Case Linked' in table_df_3.columns:
                    total_tc_linked = len(table_df_3['Test Case Linked'].dropna().unique())
                else:
                    total_tc_linked = 0

                st.markdown(
                    f"""Added in the test scope of the release <b>{total_tc_linked}</b> Test case(s) linked to the <b>{total_prs}</b> PRS(s) that were added from the Product Verification Test Records [REF-X] for Tasy EMR as they were linked to requirements changed in the revision M of the Product Requirement [REF-X].""",
                    unsafe_allow_html=True
                )
                # Exibir o texto explicativo antes do DataFrame
                st.dataframe(table_df_3, use_container_width=True)
                
            # --- REVISION HISTORY TABLE 4 ---
            #Teste Cases only in Records (excluding those linked to PRSs only in Records)
            protocols_only_in_records = force_to_strings(sorted(set(TRcol) - set(TPcol)))
            protocols_only_in_records = [p for p in protocols_only_in_records if p not in shown_tc_2]

            protocol_to_prs_records = {}
            if protocols_only_in_records:
                for prot in force_to_strings(protocols_only_in_records):
                    linked_prs = [p for tc, p in zip(TRcol, TRprs) if tc == prot]
                    linked_prs = [p for p in linked_prs if p not in shown_prs_2]
                    linked_prs = force_to_strings(linked_prs)
                    if linked_prs:
                        protocol_to_prs_records[prot] = linked_prs

                table_data_4 = []
                for prot, prs_list in protocol_to_prs_records.items():
                    for prs in prs_list:
                        table_data_4.append({'PRS Linked': prs, 'Test Case ID': prot})
                        
                # Garante pelo menos uma linha
                if not table_data_4:
                    table_data_4 = [{'PRS Linked': '', 'Test Case ID': ''}]

                table_df_4 = pd.DataFrame(table_data_4)
                total_protocols_records = len(protocol_to_prs_records)
                total_prs_linked = len(set([row['PRS Linked'] for row in table_data_4]))

                st.markdown(
                    f"""Added in the test scope of the release <b>{total_protocols_records}</b> Test case(s) linked to the <b>{total_prs_linked}</b> PRS(s) (listed below), the results were documented in the Product Verification Test Records [REF-X]:""",
                    unsafe_allow_html=True
                )
                # Exibir o texto explicativo antes do DataFrame
                st.dataframe(table_df_4, use_container_width=True)
            
            # --- REVISION HISTORY TABLE 5: Editorial changes ---

            editorial_pairs = set()

            common_test_cases = set(protocol_df['Test Case']) & set(records_df['Test Case'])

            for tc in force_to_strings(common_test_cases):
                proto_tc_df = protocol_df[protocol_df['Test Case'] == tc]
                rec_tc_df = records_df[records_df['Test Case'] == tc]
                prs_common = set(proto_tc_df['PRS']) & set(rec_tc_df['PRS'])
                for prs in force_to_strings(prs_common):
                    proto_row = proto_tc_df[proto_tc_df['PRS'] == prs]
                    rec_row = rec_tc_df[rec_tc_df['PRS'] == prs]
                    # Pega os campos relevantes
                    proto_desc = force_to_strings(proto_row['Brief Description'].tolist())
                    rec_desc = force_to_strings(rec_row['Brief Description'].tolist())
                    proto_step = force_to_strings(proto_row['Action / Step'].tolist())
                    rec_step = force_to_strings(rec_row['Action / Step'].tolist())
                    proto_exp = force_to_strings(proto_row['Expected Result'].tolist())
                    rec_exp = force_to_strings(rec_row['Expected Result'].tolist())
                    # Se qualquer campo for diferente, adiciona o par único
                    if (proto_desc != rec_desc) or (proto_step != rec_step) or (proto_exp != rec_exp):
                        editorial_pairs.add((str(tc), str(prs)))

            # Monta a tabela final apenas com Test Case ID e PRS únicos
            editorial_list = [{'PRS': prs, 'Test Case ID': tc} for tc, prs in editorial_pairs]
            editorial_df = pd.DataFrame(editorial_list).drop_duplicates()

            st.markdown(
                f"""The following test cases presented in the Product Verification Protocol [REF-X] have undergone editorial changes to provide more detail on the test case flow. This brings no changes to the test case flows linked to the PRS. The table below lists the affected protocols and their linked PRSs.""",
            )
            if not editorial_df.empty:
                st.dataframe(editorial_df, use_container_width=True)
            else:
                st.info("No editorial changes found between Protocol and Records for the selected fields.")
                
            # --- Columns Comparison Section ---
            st.markdown("## Columns comparison")
            tc_results = {
                'Test Case Only in Protocol': protocols_only_in_protocol,
                'Test Case Only in Records': protocols_only_in_records,
                'Test Case Common to Both': sorted(set(TPcol) & set(TRcol))
            }
            prs_results = {
                'Product Requirements Only in Protocol': prs_only_in_protocol,
                'Product Requirements Only in Records': prs_only_in_records,
                'Product Requirements Common to Both': sorted(set(TPprs) & set(TRprs))
            }
            desc_results = {
                'Brief Description Only in Protocol': sorted(set(TPdesc) - set(TRdesc)),
                'Brief Description Only in Records': sorted(set(TRdesc) - set(TPdesc)),
                'Brief Description Common to Both': sorted(set(TPdesc) & set(TRdesc))
            }            
            step_results = {
                'Action / Step Only in Protocol': sorted(set(TPstep) - set(TRstep)),
                'Action / Step Only in Records': sorted(set(TRstep) - set(TPstep)),
                'Action / Step Common to Both': sorted(set(TPstep) & set(TRstep))
            }
            exp_results = {
                'Expected Result Only in Protocol': sorted(set(TPexp) - set(TRexp)),
                'Expected Result Only in Records': sorted(set(TRexp) - set(TPexp)),
                'Expected Result Common to Both': sorted(set(TPexp) & set(TRexp))
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
            for result_dict in [tc_results, prs_results, step_results, desc_results, step_results, exp_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                    # Preencher até o tamanho máximo
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))

            # Agora pode criar os DataFrames sem risco de erro
            tc_df = pd.DataFrame(tc_results)
            prs_df = pd.DataFrame(prs_results)
            # --- Brief Description Comparison: apenas casos com diferença ---
            brief_diff = []
            common_test_cases = set(protocol_df['Test Case']) & set(records_df['Test Case'])
            for tc in force_to_strings(common_test_cases):
                proto_desc = force_to_strings(protocol_df[protocol_df['Test Case'] == tc]['Brief Description'].tolist())
                rec_desc = force_to_strings(records_df[records_df['Test Case'] == tc]['Brief Description'].tolist())
                if proto_desc != rec_desc:
                    brief_diff.append({'Test Case ID': tc,
                                        'Brief Description Protocol': proto_desc[0] if proto_desc else '',
                                        'Brief Description Records': rec_desc[0] if rec_desc else ''})
            desc_df = pd.DataFrame(brief_diff)
            # --- Action / Step Comparison: apenas casos com diferença ---
            step_diff = []
            for tc in force_to_strings(common_test_cases):
                proto_step = force_to_strings(protocol_df[protocol_df['Test Case'] == tc]['Action / Step'].tolist())
                rec_step = force_to_strings(records_df[records_df['Test Case'] == tc]['Action / Step'].tolist())
                if proto_step != rec_step:
                    step_diff.append({'Test Case ID': tc,
                                      'Action / Step Protocol': proto_step[0] if proto_step else '',
                                      'Action / Step Records': rec_step[0] if rec_step else ''})
            step_df = pd.DataFrame(step_diff)
            # --- Expected Result Comparison: apenas casos com diferença ---
            exp_diff = []
            for tc in force_to_strings(common_test_cases):
                proto_exp = force_to_strings(protocol_df[protocol_df['Test Case'] == tc]['Expected Result'].tolist())
                rec_exp = force_to_strings(records_df[records_df['Test Case'] == tc]['Expected Result'].tolist())
                if proto_exp != rec_exp:
                    exp_diff.append({'Test Case ID': tc,
                                    'Expected Result Protocol': proto_exp[0] if proto_exp else '',
                                    'Expected Result Records': rec_exp[0] if rec_exp else ''})
            exp_df = pd.DataFrame(exp_diff)
            # --- Revision Mismatches ---
            st.write("Test Case Comparison")
            st.dataframe(tc_df, use_container_width=True)
            st.write("PRS Comparison")
            st.dataframe(prs_df, use_container_width=True)
            st.write("Brief Description Comparison")
            st.dataframe(desc_df, use_container_width=True)
            st.write("Action / Step Comparison")
            st.dataframe(step_df, use_container_width=True)
            st.write("Expected Result Comparison")
            st.dataframe(exp_df, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Test Case COMP', index=False)
                prs_df.to_excel(writer, sheet_name='PRS COMP', index=False)
                desc_df.to_excel(writer, sheet_name='Brief Description COMP', index=False)
                step_df.to_excel(writer, sheet_name='Action_Step_COMP', index=False)  # <-- Corrigido
                exp_df.to_excel(writer, sheet_name='Expected_Result_COMP', index=False)  # <-- Corrigido
                if not revision_mismatches.empty:
                    revision_mismatches.to_excel(writer, sheet_name='Revision Mismatches', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VER_TPR_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )