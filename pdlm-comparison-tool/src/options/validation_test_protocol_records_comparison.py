import streamlit as st
import pandas as pd
import tempfile
import io
import re

def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

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

def filter_ignored(series, ignore_sentences=IGNORE_SENTENCES):
    ignore_lower = [s.strip().lower() for s in ignore_sentences]
    return series[~series.str.strip().str.lower().isin(ignore_lower)]

def run_comparison():
    st.markdown("### Validation Test Protocol x Records Comparison")
    st.info("Upload your Validation Test Protocol and Validation Test Records files to compare, including Revision History.")
    col1, col2 = st.columns(2)
    with col1:
        tp_file = st.file_uploader("Upload Validation Test Protocol (.xlsx)", type=["xlsx"], key="tp_val_rec")
    with col2:
        tr_file = st.file_uploader("Upload Validation Test Records (.xlsx)", type=["xlsx"], key="tr_val_rec")
    if tp_file and tr_file:
        if st.button("🔍 Run Comparison", key="run_val_tp_rec"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tp:
                tmp_tp.write(tp_file.read())
                tp_path = tmp_tp.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name
            xl_tp = pd.ExcelFile(tp_path)
            xl_tr = pd.ExcelFile(tr_path)
            revision_mismatches = pd.DataFrame()
            tp_sheet = get_sheet_with_fallback(xl_tp, 'Test Case Report')
            tr_sheet = get_sheet_with_fallback(xl_tr, 'Test Case Report')
            # Read DataFrames and strip columns
            tp_df = pd.read_excel(tp_path, sheet_name=tp_sheet)
            tr_df = pd.read_excel(tr_path, sheet_name=tr_sheet)
            tp_df.columns = tp_df.columns.str.strip()
            tr_df.columns = tr_df.columns.str.strip()
            # --- EXTRA CHECK: Validation Protocol ID with last round status mismatch ---
            if all(col in tr_df.columns for col in ['Validation Protocol ID', 'Round', 'Result (Pass/Fail)', 'Defect / Enhancement Status']):
                tr_df['Round'] = pd.to_numeric(tr_df['Round'], errors='coerce')
                last_rounds = tr_df.sort_values('Round').groupby('Validation Protocol ID').tail(1)
                mask_passed_open = (
                    (last_rounds['Result (Pass/Fail)'].str.strip().str.lower() == 'passed') &
                    (last_rounds['Defect / Enhancement Status'].str.strip().str.lower() == 'open')
                )
                mask_failed_closed_or_no_defect = (
                    (last_rounds['Result (Pass/Fail)'].str.strip().str.lower() == 'failed') &
                        (
                            (last_rounds['Defect / Enhancement Status'].str.strip().str.lower() == 'closed') |
                            (last_rounds['Defect / Enhancement Status'].str.contains('no defect/enhancement raised', case=False, na=False))
                        )
                )
                mismatch_df = last_rounds[mask_passed_open | mask_failed_closed_or_no_defect][
                    ['Validation Protocol ID', 'Round', 'Result (Pass/Fail)', 'Defect / Enhancement Status']
                ]
                if not mismatch_df.empty:
                    st.markdown("### Protocols with inconsistent Result/Defect Status in last Round")
                    st.dataframe(mismatch_df, use_container_width=True)
                else:
                    st.warning("Columns for extra check not found in Validation Test Records file.")
                            
            TPcol, TPurs = get_clean_cols(tp_path, tp_sheet, 'Validation Protocol ID', 'URS')
            TRcol, TRurs = get_clean_cols(tr_path, tr_sheet, 'Validation Protocol ID', 'URS')
            TPcol = normalize_spaces(filter_ignored(TPcol))
            TPurs = normalize_spaces(filter_ignored(TPurs))
            TRcol = normalize_spaces(filter_ignored(TRcol))
            TRurs = normalize_spaces(filter_ignored(TRurs))
            TPstep = get_clean_cols(tp_path, tp_sheet, 'Step')
            TRstep = get_clean_cols(tr_path, tr_sheet, 'Step')
            TPpre = get_clean_cols(tp_path, tp_sheet, 'Precondition')
            TRpre = get_clean_cols(tr_path, tr_sheet, 'Precondition')
            TPact = get_clean_cols(tp_path, tp_sheet, 'Activity')
            TRact = get_clean_cols(tr_path, tr_sheet, 'Activity')
            TPexp = get_clean_cols(tp_path, tp_sheet, 'Expected Result')
            TRexp = get_clean_cols(tr_path, tr_sheet, 'Expected Result')
            TPstep = normalize_spaces(filter_ignored(TPstep))
            TRstep = normalize_spaces(filter_ignored(TRstep))
            TPpre = normalize_spaces(filter_ignored(TPpre))
            TRpre = normalize_spaces(filter_ignored(TRpre))
            TPact = normalize_spaces(filter_ignored(TPact))
            TRact = normalize_spaces(filter_ignored(TRact))
            TPexp = normalize_spaces(filter_ignored(TPexp))
            TRexp = normalize_spaces(filter_ignored(TRexp))
            step_results = {
                'Step Only in Protocol': sorted(set(TPstep) - set(TRstep)),
                'Step Only in Records': sorted(set(TRstep) - set(TPstep)),
                'Step Common to Both': sorted(set(TPstep) & set(TRstep))
            }
            pre_results = {
                'Precondition Only in Protocol': sorted(set(TPpre) - set(TRpre)),
                'Precondition Only in Records': sorted(set(TRpre) - set(TPpre)),
                'Precondition Common to Both': sorted(set(TPpre) & set(TRpre))
            }
            act_results = {
                'Activity Only in Protocol': sorted(set(TPact) - set(TRact)),
                'Activity Only in Records': sorted(set(TRact) - set(TPact)),
                'Activity Common to Both': sorted(set(TPact) & set(TRact))
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
            
            # --- Revision History Section ---
            st.markdown("## Revision History")
            # --- REVISION HISTORY TABLE 1 ---
            urs_only_in_protocol = force_to_strings(sorted(set(TPurs) - set(TRurs)))
            shown_tc_1 = set()
            shown_urs_1 = set()
            if urs_only_in_protocol:
                urs_to_tc = {}
                for urs in force_to_strings(urs_only_in_protocol):
                    linked_tcs = [tc for tc, u in zip(TPcol, TPurs) if u == urs]
                    linked_tcs = force_to_strings(linked_tcs)
                    urs_to_tc[urs] = linked_tcs
                    shown_tc_1.update(linked_tcs)
                    shown_urs_1.add(urs)

                table_data_1 = []
                for urs, tcs in urs_to_tc.items():
                    urs_str = str(urs) if urs is not None else ''
                    for tc in tcs if tcs is not None else ['']:
                        tc_str = str(tc) if tc is not None else ''
                        table_data_1.append({
                            'URS Only in Protocol': urs_str,
                            'Validation Protocol ID Linked': tc_str
                    })      
                        
                table_df_1 = pd.DataFrame(table_data_1) 
                        
                # Garante que o DataFrame está pronto para exibição
                table_df_1 = table_df_1.fillna('')  # Substitui None por string vazia

                # Exibir o dataframe com parâmetros explícitos
                #if table_df_1 is not None and not table_df_1.empty:
                    #st.dataframe(table_df_1, use_container_width=True)
                #else:
                    #st.write("DataFrame is empty or None")
                    
                total_tc = len(TPcol)
                total_urs = len(urs_only_in_protocol)
                if 'Test Case Linked' in table_df_1.columns:
                    total_tc_linked = len(table_df_1['Test Case Linked'].dropna().unique())
                else:
                    total_tc_linked = 0

                st.markdown(
                    f"""Out of the total of <b>{total_tc}</b> Protocols listed in version xx.xx.xxxx.xx in the Product validation Protocol [REF-X] for Tasy EMR, <b>{total_tc_linked}</b> protocol(s) (listed below) linked to <b>{total_urs}</b> URSs were removed from the Product Validation Records [REF-X] for Tasy EMR as they were linked to requirements changed in the revision X of the User Requirement [REF-X].""",
                    unsafe_allow_html=True
                )
                st.dataframe(table_df_1, use_container_width=True)

            # --- REVISION HISTORY TABLE 2 ---
            protocols_only_in_protocol = force_to_strings(sorted(set(TPcol) - set(TRcol)))
            protocols_only_in_protocol = [p for p in protocols_only_in_protocol if p not in shown_tc_1]

            protocol_to_urs = {}
            if protocols_only_in_protocol:
                for prot in force_to_strings(protocols_only_in_protocol):
                    linked_urs = [u for tc, u in zip(TPcol, TPurs) if tc == prot]
                    linked_urs = [u for u in linked_urs if u not in shown_urs_1]
                    linked_urs = force_to_strings(linked_urs)
                    if linked_urs:
                        protocol_to_urs[prot] = linked_urs

                table_data_2 = []
                for prot, urs_list in protocol_to_urs.items():
                    for urs in urs_list:
                        table_data_2.append({'URS Linked': urs, 'Validation Protocol ID': prot})
                        
                # Garante pelo menos uma linha
                if not table_data_2:
                    table_data_2 = [{'URS Linked': '', 'Validation Protocol ID': ''}]

                table_df_2 = pd.DataFrame(table_data_2)
                total_protocols = len(TPcol)
                total_protocols_only = len(protocol_to_urs)
                total_urs_linked = len(set([row['URS Linked'] for row in table_data_2]))

                st.markdown(
                    f"""Out of the total of <b>{total_protocols}</b> Protocol listed in the Product Validation Protocol [REF-X] for Tasy EMR, <b>{total_protocols_only}</b> protocols linked to <b>{total_urs_linked}</b> URSs listed below and in the Product Validation Records [REF-X] were not validated in release X.XX.XXXX.XX. This occurred because the protocols were not part of the scope of this release. The processes is covered in other protocol linked to the URSs.""",
                    unsafe_allow_html=True
                )
                st.dataframe(table_df_2, use_container_width=True)

            # --- REVISION HISTORY TABLE 3: Editorial changes ---
            protocols_editorial = set()
            editorial_rows = []

            def add_editorial_pairs(col_name, only_in_records):
                if col_name not in tr_df.columns:
                    st.warning(f"Column '{col_name}' not found in Validation Test Records file.")
                    return
                if not only_in_records:
                    return
                for value in force_to_strings(only_in_records):
                    mask = tr_df[col_name].astype(str).str.strip() == str(value)
                    for idx, row in tr_df[mask].iterrows():
                        urs_id = str(row.get('URS', '')).strip()
                        prot_id = str(row.get('Validation Protocol ID', '')).strip()
                        if prot_id:
                            pair = (prot_id, urs_id)
                            if pair not in protocols_editorial:
                                protocols_editorial.add(pair)
                                editorial_rows.append({'URS ID': urs_id,'Validation Protocol ID': prot_id})


            add_editorial_pairs('Step', step_results['Step Only in Records'])
            add_editorial_pairs('Precondition', pre_results['Precondition Only in Records'])
            add_editorial_pairs('Activity', act_results['Activity Only in Records'])
            add_editorial_pairs('Expected Result', exp_results['Expected Result Only in Records'])

            # Limpeza final dos dados editoriais
            editorial_rows = [
                {'URS ID': str(row['URS ID']), 'Validation Protocol ID': str(row['Validation Protocol ID'])}
                for row in editorial_rows
                if row['URS ID'] and row['Validation Protocol ID']
            ]

            editorial_df = pd.DataFrame(editorial_rows).drop_duplicates()

            if not editorial_df.empty:
                st.markdown(
                    """The following protocols presented in the Product Validation Protocol [REF-X] have undergone editorial changes to provide more detail on the protocol flow. This brings no changes to the protocol flows linked to the URS. The table below lists the affected protocols and their linked URSs.""",
                    unsafe_allow_html=True
                )
                st.dataframe(editorial_df, use_container_width=True)

            # --- Columns Comparison Section ---
            st.markdown("## Columns comparison")
            tc_results = {
                'VAL Protocol Only in Protocol': sorted(set(TPcol) - set(TRcol)),
                'VAL Protocol Only in Records': sorted(set(TRcol) - set(TPcol)),
                'VAL Protocol Common to Both': sorted(set(TPcol) & set(TRcol))
            }
            urs_results = {
                'VAL URS Only in Protocol': urs_only_in_protocol,
                'VAL URS Only in Records': sorted(set(TRurs) - set(TPurs)),
                'VAL URS Common to Both': sorted(set(TPurs) & set(TRurs))
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
            for result_dict in [tc_results, urs_results, step_results, pre_results, act_results, exp_results]:
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
            def get_field_diff(tp_df, tr_df, field):
                # IDs presentes em ambos
                common_ids = set(tp_df['Validation Protocol ID']).intersection(set(tr_df['Validation Protocol ID']))
                diff_rows = []
                for vp_id in common_ids:
                    tp_vals = tp_df[tp_df['Validation Protocol ID'] == vp_id][field].dropna().astype(str).str.strip().unique()
                    tr_vals = tr_df[tr_df['Validation Protocol ID'] == vp_id][field].dropna().astype(str).str.strip().unique()
                    if set(tp_vals) != set(tr_vals):
                        diff_rows.append({
                            'Validation Protocol ID': vp_id,
                            f'{field} in Protocol': "; ".join(tp_vals),
                            f'{field} in Records': "; ".join(tr_vals)
                        })
                return pd.DataFrame(diff_rows)

            step_diff_df = get_field_diff(tp_df, tr_df, 'Step')
            precondition_diff_df = get_field_diff(tp_df, tr_df, 'Precondition')
            activity_diff_df = get_field_diff(tp_df, tr_df, 'Activity')
            expected_result_diff_df = get_field_diff(tp_df, tr_df, 'Expected Result')
            st.write("### Protocols Comparison")
            st.dataframe(tc_df, use_container_width=True)
            st.write("### URS Comparison")
            st.dataframe(urs_df, use_container_width=True)
            if not step_diff_df.empty:
                st.markdown("### Step Differences")
                st.dataframe(step_diff_df, use_container_width=True)
            if not precondition_diff_df.empty:
                st.markdown("### Precondition Differences")
                st.dataframe(precondition_diff_df, use_container_width=True)
            if not activity_diff_df.empty:
                st.markdown("### Activity Differences")
                st.dataframe(activity_diff_df, use_container_width=True)
            if not expected_result_diff_df.empty:
                st.markdown("### Expected Result Differences")
                st.dataframe(expected_result_diff_df, use_container_width=True)
            step_df = pd.DataFrame(step_results)
            pre_df = pd.DataFrame(pre_results)
            act_df = pd.DataFrame(act_results)
            exp_df = pd.DataFrame(exp_results)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Protocol COMP', index=False)
                urs_df.to_excel(writer, sheet_name='URS COMP', index=False)
                step_df.to_excel(writer, sheet_name='Step COMP', index=False)
                pre_df.to_excel(writer, sheet_name='Precondition COMP', index=False)
                act_df.to_excel(writer, sheet_name='Activity COMP', index=False)
                exp_df.to_excel(writer, sheet_name='Expected Result COMP', index=False)
                if not revision_mismatches.empty:
                    revision_mismatches.to_excel(writer, sheet_name='Revision Mismatches', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VAL_TPR_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )