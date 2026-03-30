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

def run_comparison():       
    st.markdown("### CMDK Verification TM requirements x Test Protocol Comparison")
    st.info("Upload your TM requirements (.xlsx) and Verification Test Protocol (.html) to compare requirements.")
    col1, col2 = st.columns(2)
    with col1:
        tm_file = st.file_uploader("Upload TM requirements (.xlsx)", type=["xlsx"], key="tmapp_ver2")
    with col2:
        tp_file = st.file_uploader("Upload Verification Test Protocol (.html)", type=["html"], key="tp_ver2_html")
    if tm_file and tp_file:
        if st.button("🔍 Run Comparison", key="run_ver_tp_tm_html"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tm:
                tmp_tm.write(tm_file.read())
                tm_path = tmp_tm.name
            tm_xl = pd.ExcelFile(tm_path)
            tm_sheet = get_sheet_with_fallback(tm_xl, 'Traceability-Verification')
            TMprs, TMids = get_clean_cols(tm_path, tm_sheet, 'PRS Requirement ID', 'Verification Test ID', skiprows=1)
            TMprs = normalize_spaces(filter_ignored(TMprs, IGNORE_SENTENCES_VAL_TP_TM))
            TMids = normalize_spaces(filter_ignored(TMids, IGNORE_SENTENCES_VAL_TP_TM))

            # Extrai Requirements e Test Case ID do HTML
            html_content = tp_file.read().decode('utf-8', errors='ignore')
            soup = BeautifulSoup(html_content, 'html.parser')
            requirements = []
            test_case_ids = []
            # Requirements
            for btag in soup.find_all('b', string=re.compile(r'Requirements:', re.IGNORECASE)):
                parent_td = btag.find_parent('td')
                if parent_td:
                    texts = []
                    found_b = False
                    for elem in parent_td.contents:
                        if elem == btag:
                            found_b = True
                        elif found_b:
                            if isinstance(elem, str):
                                texts.append(elem)
                            else:
                                texts.append(elem.get_text())
                    value = ''.join(texts).strip()
                    for req in re.split(r'[;,]', value):
                        val = req.strip()
                        if val:
                            requirements.append(val)
            # Test Case ID
            for th in soup.find_all('th', string=re.compile(r'Test Case ID:', re.IGNORECASE)):
                td = th.find_next_sibling('td')
                if td:
                    for val in re.split(r'[;,]', td.get_text()):
                        val = val.strip()
                        if val:
                            test_case_ids.append(val)
            requirements = pd.Series(requirements).drop_duplicates().astype(str).str.strip()
            requirements = normalize_spaces(requirements)
            test_case_ids = pd.Series(test_case_ids).drop_duplicates().astype(str).str.strip()
            test_case_ids = normalize_spaces(test_case_ids)

            # PRS comparison
            prs_results = {
                'Product Requirements Only in Traceability Application': sorted(set(TMprs) - set(requirements)),
                'Product Requirements Only in Verification Test Protocol (HTML)': sorted(set(requirements) - set(TMprs)),
                'Product Requirements Common to Both Documents': sorted(set(TMprs) & set(requirements))
            }
            # Test Case ID comparison
            tc_results = {
                'Test Case ID Only in Traceability Application': sorted(set(TMids) - set(test_case_ids)),
                'Test Case ID Only in Verification Test Protocol (HTML)': sorted(set(test_case_ids) - set(TMids)),
                'Test Case ID Common to Both Documents': sorted(set(TMids) & set(test_case_ids))
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
            # Preencher até o tamanho máximo para ambos
            max_len_prs = max(len(force_to_strings(v)) for v in prs_results.values())
            for k in prs_results:
                clean_list = force_to_strings(prs_results[k])
                if not clean_list:
                    clean_list = [""]
                prs_results[k] = clean_list + [""] * (max_len_prs - len(clean_list))
            max_len_tc = max(len(force_to_strings(v)) for v in tc_results.values())
            for k in tc_results:
                clean_list = force_to_strings(tc_results[k])
                if not clean_list:
                    clean_list = [""]
                tc_results[k] = clean_list + [""] * (max_len_tc - len(clean_list))

            prs_df = pd.DataFrame(prs_results)
            tc_df = pd.DataFrame(tc_results)
            st.write("**PRS Comparison**")
            st.dataframe(prs_df, use_container_width=True)
            st.write("**Test Case ID Comparison**")
            st.dataframe(tc_df, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                prs_df.to_excel(writer, sheet_name='PRS COMP', index=False)
                tc_df.to_excel(writer, sheet_name='Test Case ID COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VER_TPxTMAPP_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # COMPARAÇÃO DE PARES COMBINADOS PRS x TEST CASE ID
            html_pairs = set()
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
                            html_pairs.add((req, tc_id))

            # Só compara pares da TM cujo Test Case ID existe no HTML
            tm_pairs = set((prs, tcid) for prs, tcid in zip(TMprs, TMids) if tcid in set(test_case_ids))

            only_in_tm = sorted(tm_pairs - html_pairs)
            only_in_html = sorted(html_pairs - tm_pairs)
            in_both = sorted(tm_pairs & html_pairs)

            def pairs_to_df(pairs, col1, col2):
                df = pd.DataFrame(pairs, columns=[col1, col2]) if pairs else pd.DataFrame({col1: [''], col2: ['']})
                return df.drop_duplicates().reset_index(drop=True)

            df_only_in_tm = pairs_to_df(only_in_tm, 'PRS Requirement ID', 'Verification Test ID')
            df_only_in_html = pairs_to_df(only_in_html, 'PRS Requirement ID', 'Verification Test ID')
            df_in_both = pairs_to_df(in_both, 'PRS Requirement ID', 'Verification Test ID')

            st.write('---')
            st.write('### Pares combinados PRS Requirement ID x Verification Test ID')
            st.write('**Pares apenas na TM:**')
            st.dataframe(df_only_in_tm, use_container_width=True)
            st.write('**Pares apenas no HTML:**')
            st.dataframe(df_only_in_html, use_container_width=True)
            st.write('**Pares em ambos:**')
            st.dataframe(df_in_both, use_container_width=True)
            
            # Conferência dos campos com valor diferente de N/A no HTML
            fields_to_check = [
                ("Tester", r'Tester:', False),
                ("Date Tested", r'Date Tested:', True),
                ("Actual Result", r'Actual Result:', False),
                ("Objective Evidence", r'Objective Evidence:', False),
                ("Verification Status", r'Verification Status:', False),
                ("Defect ID", r'Defect ID:', True),
            ]
            results_not_na = []
            for field_name, field_regex, is_bold in fields_to_check:
                if is_bold:
                    # Busca <b>Field:</b> seguido de valor diferente de N/A
                    for btag in soup.find_all('b', string=re.compile(field_regex, re.IGNORECASE)):
                        parent_td = btag.find_parent('td')
                        if parent_td:
                            value = parent_td.get_text().replace(btag.get_text(), '').strip()
                            if value and value != "N/A":
                                results_not_na.append(f"{field_name}: {value}")
                else:
                    # Busca <th>Field:</th> seguido de <td>valor diferente de N/A</td>
                    for th in soup.find_all('th', string=re.compile(field_regex, re.IGNORECASE)):
                        td = th.find_next_sibling('td')
                        if td:
                            val = td.get_text(strip=True)
                            if val and val != "N/A":
                                results_not_na.append(f"{field_name}: {val}")
            # Conferência de <td colspan="2">valor diferente de N/A</td> para Actual Result e Objective Evidence
            for field_name in ["Actual Result", "Objective Evidence"]:
                for th in soup.find_all('th', string=re.compile(field_name, re.IGNORECASE)):
                    td = th.find_next_sibling('td')
                    if td:
                        val = td.get_text(strip=True)
                        if val and val != "N/A":
                            results_not_na.append(f"{field_name}: {val}")
            # Exibe resultado
            st.write('---')
            st.write('### Conferência do campos: Tester, Date Tested, Actual Result, Objective Evidence, Verification Status e Defect ID com valor diferente de N/A')
            if results_not_na:
                for res in results_not_na:
                    st.write(f"- {res}")
            else:
                st.write("Nenhum campo diferente de N/A encontrado.")
            
            # Conferência se algum campo está vazio no HTML
            fields_to_check_empty = [
                ("Test Name", r'Test Name:', False),
                ("Test Case ID", r'Test Case ID:', False),
                ("Requirements", r'Requirements:', True),
                ("Tester", r'Tester:', False),
                ("Date Tested", r'Date Tested:', True),
                ("Test Type", r'Test Type:', False),
                ("Steps", r'Steps:', False),
                ("Expected Result", r'Expected Result:', False),
                ("Actual Result", r'Actual Result:', False),
                ("Objective Evidence", r'Objective Evidence:', False),
                ("Verification Status", r'Verification Status:', False),
                ("Defect ID", r'Defect ID:', True),
            ]
            fields_empty = []
            for field_name, field_regex, is_bold in fields_to_check_empty:
                found = False
                if field_name == "Steps":
                    # Busca Steps no <td> do <tbody> se não encontrar no <th>
                    steps_val = ""
                    # Primeiro tenta pelo <th>
                    for th in soup.find_all('th', string=re.compile(field_regex, re.IGNORECASE)):
                        td = th.find_next_sibling('td')
                        if td:
                            steps_val = td.get_text(strip=True)
                            found = True
                    # Se não achou, busca no <tbody> <td colspan="3">
                    if not steps_val:
                        for tbody in soup.find_all('tbody'):
                            for td in tbody.find_all('td', colspan=True):
                                steps_val = td.get_text(strip=True)
                                if steps_val:
                                    found = True
                                    break
                    if not steps_val:
                        fields_empty.append(field_name)
                elif is_bold:
                    for btag in soup.find_all('b', string=re.compile(field_regex, re.IGNORECASE)):
                        parent_td = btag.find_parent('td')
                        if parent_td:
                            value = parent_td.get_text().replace(btag.get_text(), '').strip()
                            if value == "":
                                fields_empty.append(field_name)
                            found = True
                else:
                    for th in soup.find_all('th', string=re.compile(field_regex, re.IGNORECASE)):
                        td = th.find_next_sibling('td')
                        if td:
                            val = td.get_text(strip=True)
                            if val == "":
                                fields_empty.append(field_name)
                            found = True
                # Se não encontrou o campo, considera como vazio
                if not found:
                    fields_empty.append(field_name)

            st.write('---')
            st.write('### Conferência de campos vazios')
            if fields_empty:
                st.warning("Os seguintes campos estão vazios ou ausentes:")
                for f in fields_empty:
                    st.write(f"- {f}")
            else:
                st.success("Todos os campos estão preenchidos.")
            
            # Mapeia pares do HTML: {Test Case ID: [Requirement IDs]}
            html_tc_to_prs = dict()
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
                        html_tc_to_prs.setdefault(tc_id, set()).update(req_vals)

            # Mapeia pares da TM: {(Verification Test ID, PRS Requirement ID)}
            tm_pairs = set()
            for prs, tcid in zip(TMprs, TMids):
                if tcid in html_tc_to_prs:  # Só compara se o TC existe no HTML
                    tm_pairs.add((tcid, prs))

            # Pares do HTML: {(Test Case ID, PRS Requirement ID)}
            html_pairs = set()
            for tc_id, prs_set in html_tc_to_prs.items():
                for prs in prs_set:
                    html_pairs.add((tc_id, prs))

            # Pares apenas na TM (TC existe no HTML, mas PRS não está no HTML para aquele TC)
            only_in_tm = sorted(tm_pairs - html_pairs)
            # Pares apenas no HTML (PRS para TC não está na TM)
            only_in_html = sorted(html_pairs - tm_pairs)
            # Pares em ambos
            in_both = sorted(tm_pairs & html_pairs)

            def pairs_to_df(pairs, col1, col2):
                df = pd.DataFrame(pairs, columns=[col1, col2]) if pairs else pd.DataFrame({col1: [''], col2: ['']})
                return df.drop_duplicates().reset_index(drop=True)

            df_only_in_tm = pairs_to_df(only_in_tm, 'Verification Test ID', 'PRS Requirement ID')
            df_only_in_html = pairs_to_df(only_in_html, 'Verification Test ID', 'PRS Requirement ID')
            df_in_both = pairs_to_df(in_both, 'Verification Test ID', 'PRS Requirement ID')

            st.write('---')
            st.write('### Pares combinados Verification Test ID x PRS Requirement ID')
            st.write('**Pares apenas na TM (TC existe no HTML, mas PRS não está no HTML para aquele TC):**')
            st.dataframe(df_only_in_tm, use_container_width=True)
            st.write('**Pares apenas no HTML (PRS para TC não está na TM):**')
            st.dataframe(df_only_in_html, use_container_width=True)
            st.write('**Pares em ambos:**')
            st.dataframe(df_in_both, use_container_width=True)