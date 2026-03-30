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

def validate_verification_status(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    issues = []
    for tr in soup.find_all('tr'):
        th_status = tr.find('th', string=re.compile(r'Verification Status:', re.IGNORECASE))
        if th_status:
            tds = tr.find_all('td')
            status = ''
            defect_id = ''
            for td in tds:
                text = td.get_text(strip=True)
                if text.lower() in ['pass', 'fail']:
                    status = text.lower()
                if td.find('b', string=re.compile(r'Defect ID:', re.IGNORECASE)):
                    defect_id = text.replace('Defect ID:', '').strip()
            # Busca o Test Case ID subindo as linhas anteriores
            test_case_id = ''
            prev_tr = tr.find_previous_sibling('tr')
            while prev_tr:
                th_tc = prev_tr.find('th', string=re.compile(r'Test Case ID:', re.IGNORECASE))
                if th_tc:
                    tc_td = th_tc.find_next_sibling('td')
                    if tc_td:
                        test_case_id = tc_td.get_text(strip=True)
                        break
                prev_tr = prev_tr.find_previous_sibling('tr')
            if status == 'pass' and defect_id.lower() != 'n/a':
                issues.append(f"Test Case ID={test_case_id}: Status=Pass, Defect ID='{defect_id}' (deveria ser N/A)")
            if status == 'fail' and (defect_id.lower() == 'n/a' or defect_id == ''):
                issues.append(f"Test Case ID={test_case_id}: Status=Fail, Defect ID='{defect_id}' (deveria ser preenchido e diferente de N/A)")
    return list(set(issues))

def validate_date_tested(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    issues = []
    last_test_case_id = ''
    for tr in soup.find_all('tr'):
        # Atualiza o último Test Case ID encontrado
        th_tc = tr.find('th', string=re.compile(r'Test Case ID:', re.IGNORECASE))
        if th_tc:
            tc_td = th_tc.find_next_sibling('td')
            if tc_td:
                last_test_case_id = tc_td.get_text(strip=True)
        # Busca Date Tested
        for td in tr.find_all('td'):
            b_date = td.find('b', string=re.compile(r'Date Tested:', re.IGNORECASE))
            if b_date:
                # Extrai o texto após "Date Tested:"
                date_text = td.get_text(strip=True).replace('Date Tested:', '').strip()
                # Verifica formato DD-MMM-AAAA
                if not re.match(r'^\d{2}-[A-Za-z]{3}-\d{4}$', date_text):
                    issues.append(f"Test Case ID={last_test_case_id}: Date Tested='{date_text}' (formato inválido, esperado DD-MMM-AAAA)")
    return issues

def extract_tcname_pairs_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    tcname_dict = {}
    last_test_name = ''
    for tr in soup.find_all('tr'):
        th_name = tr.find('th', string=re.compile(r'Test Name:', re.IGNORECASE))
        if th_name:
            td_name = th_name.find_next_sibling('td')
            if td_name:
                last_test_name = td_name.get_text(strip=True)
        th_tc = tr.find('th', string=re.compile(r'Test Case ID:', re.IGNORECASE))
        if th_tc:
            tc_td = th_tc.find_next_sibling('td')
            if tc_td:
                tc_id = tc_td.get_text(strip=True)
                tcname_dict[tc_id] = last_test_name
    return tcname_dict

def run_comparison():
    st.markdown("### Verification Test Records x Verification Test Protocol Comparison (HTML x HTML)")
    st.info("Upload your Verification Test Records (.html) and Verification Test Protocol (.html) to compare Test Case IDs and Requirements.")
    col1, col2 = st.columns(2)
    with col1:
        protocol_file = st.file_uploader("Upload Verification Test Protocol (.html)", type=["html"], key="protocol_html") 
        if not protocol_file:
            st.warning("Please upload the Verification Test Protocol file.")
    with col2:
        records_file = st.file_uploader("Upload Verification Test Records (.html)", type=["html"], key="records_html") 
        if not records_file:
            st.warning("Please upload the Verification Test Records file.")
    if protocol_file and records_file:
        if st.button("🔍 Run Comparison", key="run_html_html"):
            records_html = records_file.read().decode('utf-8', errors='ignore')
            protocol_html = protocol_file.read().decode('utf-8', errors='ignore')

            records_pairs = extract_pairs_from_html(records_html)
            protocol_pairs = extract_pairs_from_html(protocol_html)

            only_in_records = sorted(records_pairs - protocol_pairs)
            only_in_protocol = sorted(protocol_pairs - records_pairs)
            in_both = sorted(records_pairs & protocol_pairs)

            def pairs_to_df(pairs, col1, col2):
                df = pd.DataFrame(pairs, columns=[col1, col2]) if pairs else pd.DataFrame({col1: [''], col2: ['']})
                return df.drop_duplicates().reset_index(drop=True)

            df_only_in_records = pairs_to_df(only_in_records, 'Test Case ID', 'PRS Requirement ID')
            df_only_in_protocol = pairs_to_df(only_in_protocol, 'Test Case ID', 'PRS Requirement ID')
            df_in_both = pairs_to_df(in_both, 'Test Case ID', 'PRS Requirement ID')

            # Comparação separada dos campos
            records_tc_ids = set(tc for tc, _ in records_pairs)
            protocol_tc_ids = set(tc for tc, _ in protocol_pairs)
            records_prs_ids = set(prs for _, prs in records_pairs)
            protocol_prs_ids = set(prs for _, prs in protocol_pairs)

            tc_ids_only_in_records = sorted(records_tc_ids - protocol_tc_ids)
            tc_ids_only_in_protocol = sorted(protocol_tc_ids - records_tc_ids)
            tc_ids_in_both = sorted(records_tc_ids & protocol_tc_ids)

            prs_ids_only_in_records = sorted(records_prs_ids - protocol_prs_ids)
            prs_ids_only_in_protocol = sorted(protocol_prs_ids - records_prs_ids)
            prs_ids_in_both = sorted(records_prs_ids & protocol_prs_ids)

            # Agrupando Test Case ID
            max_tc_len = max(len(tc_ids_only_in_protocol), len(tc_ids_only_in_records), len(tc_ids_in_both))
            tc_df = pd.DataFrame({
                'Test Case only Protocols': tc_ids_only_in_protocol + [''] * (max_tc_len - len(tc_ids_only_in_protocol)),
                'Test Case only Records': tc_ids_only_in_records + [''] * (max_tc_len - len(tc_ids_only_in_records)),
                'Test Case In Both': tc_ids_in_both + [''] * (max_tc_len - len(tc_ids_in_both))
            })

            # Agrupando PRS Requirement ID
            max_prs_len = max(len(prs_ids_only_in_protocol), len(prs_ids_only_in_records), len(prs_ids_in_both))
            prs_df = pd.DataFrame({
                'PRS Only in Protocols': prs_ids_only_in_protocol + [''] * (max_prs_len - len(prs_ids_only_in_protocol)),
                'PRS Only in Records': prs_ids_only_in_records + [''] * (max_prs_len - len(prs_ids_only_in_records)),
                'PRS In Both': prs_ids_in_both + [''] * (max_prs_len - len(prs_ids_in_both))
            })

            st.write('---')
            st.write('### Test Case ID Comparison')
            st.dataframe(tc_df, use_container_width=True)
            st.write('### PRS Requirement ID Comparison')
            st.dataframe(prs_df, use_container_width=True)            
            st.write('---')
            st.write('### Test Case ID x PRS Requirement ID pairs')
            st.write('**Pairs only in Verification Test Records:**')
            st.dataframe(df_only_in_records, use_container_width=True)
            st.write('**Pairs only in Verification Test Protocol:**')
            st.dataframe(df_only_in_protocol, use_container_width=True)
            st.write('**Pairs in both:**')
            st.dataframe(df_in_both, use_container_width=True)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                df_only_in_records.to_excel(writer, sheet_name='Verification Test Records Only', index=False)
                df_only_in_protocol.to_excel(writer, sheet_name='Verification Test Protocol Only', index=False)
                df_in_both.to_excel(writer, sheet_name='Common', index=False)
                tc_df.to_excel(writer, sheet_name='Test Case ID Comparison', index=False)
                prs_df.to_excel(writer, sheet_name='PRS Requirement ID Comparison', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="Verification_Test_Records_vs_Verification_Test_Protocol_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            issues = validate_verification_status(records_html)
            if issues:
                st.warning("Foram encontrados problemas nos campos de Verification Status e Defect ID:")
                for issue in issues:
                    st.write(f"- {issue}")
            else:
                st.success("Todos os campos de Verification Status e Defect ID estão preenchidos corretamente.")
            
            date_issues = validate_date_tested(records_html)
            if date_issues:
                st.warning("Foram encontrados problemas no campo Date Tested:")
                for issue in date_issues:
                    st.write(f"- {issue}")
            else:
                st.success("Todos os campos de Date Tested estão preenchidos corretamente.")
            
            records_tcname = extract_tcname_pairs_from_html(records_html)
            protocol_tcname = extract_tcname_pairs_from_html(protocol_html)

            # IDs presentes em ambos
            common_tc_ids = set(records_tcname.keys()) & set(protocol_tcname.keys())
            tcname_issues = []
            for tc_id in common_tc_ids:
                if records_tcname[tc_id] != protocol_tcname[tc_id]:
                    tcname_issues.append(
                        f"Test Case ID={tc_id}: Test Name diferente! Records='{records_tcname[tc_id]}', Protocols='{protocol_tcname[tc_id]}'"
                    )

            if tcname_issues:
                st.warning("Foram encontrados Test Case IDs com Test Name diferente entre Records e Protocols:")
                for issue in tcname_issues:
                    st.write(f"- {issue}")
            else:
                st.success("Todos os Test Case IDs possuem Test Name igual entre Records e Protocols.")
