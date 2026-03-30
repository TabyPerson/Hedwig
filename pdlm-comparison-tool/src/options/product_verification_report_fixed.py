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
    
def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

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


def run_comparison():
    st.markdown("### Product Verification Report")
    st.info("Upload your Product Verification Protocol, Product Verification Records, and Product Defect Status Report files.")

    col1, col2, col3 = st.columns(3)
    with col1:
        protocol_file = st.file_uploader("Upload Product Verification Protocol (.xlsx)", type=["xlsx"], key="pvr_protocol")
    with col2:
        records_file = st.file_uploader("Upload Product Verification Records (.xlsx)", type=["xlsx"], key="pvr_records")
    with col3:
        defect_file = st.file_uploader("Upload Product Defect Status Report (.xlsx)", type=["xlsx"], key="pvr_defect")

    if protocol_file and records_file and defect_file:
        if st.button("🔍 Run Product Verification Report", key="run_pvr"):
            # Load Protocol (merge sheets)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_protocol:
                tmp_protocol.write(protocol_file.read())
                protocol_path = tmp_protocol.name
            xl_protocol = pd.ExcelFile(protocol_path)
            protocol_sheets = ["Test Case Report - URS MD", "Test Case Report - URS NMD"]
            protocol_dfs = []
            for sheet in protocol_sheets:
                if sheet in xl_protocol.sheet_names:
                    df = pd.read_excel(protocol_path, sheet_name=sheet)
                    protocol_dfs.append(df)
                else:
                    st.warning(f"Worksheet named '{sheet}' not found in Protocol file.")
            if protocol_dfs:
                protocol_df = pd.concat(protocol_dfs, ignore_index=True)
            else:
                st.error("No valid Protocol sheets found.")
                st.stop()
                return

            # Load Records (merge sheets)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_records:
                tmp_records.write(records_file.read())
                records_path = tmp_records.name
            xl_records = pd.ExcelFile(records_path)
            records_sheets = ["Test Case Report - URS MD", "Test Case Report - RMM", "Test Case Report - URS NMD"]
            records_dfs = []
            for sheet in records_sheets:
                if sheet in xl_records.sheet_names:
                    df = pd.read_excel(records_path, sheet_name=sheet)
                    records_dfs.append(df)
                else:
                    st.warning(f"Worksheet named '{sheet}' not found in Records file.")
            if records_dfs:
                records_df = pd.concat(records_dfs, ignore_index=True)
            else:
                st.error("No valid Records sheets found.")
                st.stop()
                return

            # Normalize columns
            for df in [protocol_df, records_df]:
                df.columns = df.columns.str.replace(r'\s+', ' ', regex=True).str.strip()

            # Use Test Case and Traceability (PRS) columns
            for df in [protocol_df, records_df]:
                if 'Test Case' not in df.columns or 'Traceability (PRS)' not in df.columns:
                    st.error("Required columns 'Test Case' and 'Traceability (PRS)' not found in one of the files.")
                    st.stop()
                    return
                df['Test Case'] = df['Test Case'].astype(str).str.strip()
                df['Traceability (PRS)'] = df['Traceability (PRS)'].astype(str).str.strip()

            # Load Defect Status Report
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_defect:
                tmp_defect.write(defect_file.read())
                defect_path = tmp_defect.name
            xl_defect = pd.ExcelFile(defect_path)
            defect_sheet = "Defect List"
            if defect_sheet not in xl_defect.sheet_names:
                st.error(f"Sheet '{defect_sheet}' not found in the uploaded file. Available sheets: {xl_defect.sheet_names}")
                st.stop()
                return
            defect_df = pd.read_excel(defect_path, sheet_name=defect_sheet, skiprows=2)
            defect_df.columns = defect_df.columns.str.replace(r'\s+', ' ', regex=True).str.strip().str.lower()
            status_col = None
            defect_id_col = None
            for col in defect_df.columns:
                if 'status' in col:
                    status_col = col
                if 'defect id' in col:
                    defect_id_col = col
            if not status_col or not defect_id_col:
                st.error(f"Could not find required columns in Defect Status Report. Columns found: {defect_df.columns.tolist()}")
                st.stop()
                return
            defect_df[status_col] = defect_df[status_col].astype(str).str.strip().str.lower()
            open_defects = defect_df[defect_df[status_col] == 'open'][defect_id_col].nunique()

            # Clean records_df for both columns using boolean masks
            def mask_test_case_column(series):
                return series.str.strip().str.startswith('TASY_VTC', na=False)
            def mask_prs_column(series):
                return series.str.strip().str.startswith('A', na=False)

            records_df_clean = records_df[
                mask_test_case_column(records_df['Test Case']) &
                mask_prs_column(records_df['Traceability (PRS)'])
            ]

            # 1. Overall Protocol executed (in records)
            records_concat = []
            missing_test_case_sheets = []
            missing_prs_sheets = []
            for sheet in records_sheets:
                if sheet in xl_records.sheet_names:
                    df = pd.read_excel(records_path, sheet_name=sheet)
                    # Normaliza os nomes das colunas para evitar problemas com espaços invisíveis
                    df.columns = (
                        df.columns
                        .str.replace(r'\s+', ' ', regex=True)
                        .str.replace('\u200b', '', regex=False)
                        .str.strip()
                    )
                    if 'Test Case' in df.columns and 'Traceability (PRS)' in df.columns:
                        records_concat.append(df)
                    else:
                        if 'Test Case' not in df.columns:
                            missing_test_case_sheets.append(sheet)
                        if 'Traceability (PRS)' not in df.columns:
                            missing_prs_sheets.append(sheet)
            if missing_test_case_sheets:
                st.warning(f"As seguintes abas não possuem a coluna 'Test Case': {missing_test_case_sheets}")
            if missing_prs_sheets:
                st.warning(f"As seguintes abas não possuem a coluna 'Traceability (PRS)': {missing_prs_sheets}")
            if records_concat:
                records_all = pd.concat(records_concat, ignore_index=True)
                records_all['Test Case'] = records_all['Test Case'].astype(str).str.strip()
                records_all['Traceability (PRS)'] = records_all['Traceability (PRS)'].astype(str).str.strip()
                notes_to_ignore = [
                    "Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.",
                    "Note: Defect/Enhancement are linked to test cases and their respective execution versions"
                ]
                # Remove linhas com títulos, notas, vazias ou só espaços
                mask_valid = (
                    (~records_all['Test Case'].str.strip().str.lower().isin(['', 'test case'] + [n.lower() for n in notes_to_ignore])) &
                    (~records_all['Traceability (PRS)'].str.strip().str.lower().isin(['', 'traceability (prs)']))
                )
                records_all = records_all[mask_valid]
                # Remove duplicatas
                records_all = records_all.drop_duplicates(subset=['Test Case', 'Traceability (PRS)'])
                protocols_executed = records_all['Test Case'].nunique()
                prs_executed = records_all['Traceability (PRS)'].nunique()

                # 2. Overall Not Executed (in protocol but not in records)
                test_cases_in_protocol = protocol_df.loc[mask_test_case_column(protocol_df['Test Case']), 'Test Case']
                test_cases_in_records = records_df.loc[mask_test_case_column(records_df['Test Case']), 'Test Case']
                protocols_not_executed = set(test_cases_in_protocol) - set(test_cases_in_records)
                prs_not_executed = protocol_df.loc[
                    protocol_df['Test Case'].isin(protocols_not_executed) & mask_prs_column(protocol_df['Traceability (PRS)']),
                    'Traceability (PRS)'
                ].nunique() if 'Traceability (PRS)' in protocol_df.columns else 0

                # 3. Overall Passed Results (Result == Pass in records)
                if 'Conclusion (Pass / Fail)' in records_df.columns:
                    records_df['Conclusion (Pass / Fail)'] = records_df['Conclusion (Pass / Fail)'].astype(str).str.strip()
                    passed_mask = (
                        records_df['Conclusion (Pass / Fail)'].str.lower().isin(['pass', 'passed']) &
                        records_df['Test Case'].str.len().gt(0) &
                        records_df['Traceability (PRS)'].str.len().gt(0)
                    )
                    passed_df = records_df[passed_mask]
                    protocols_passed = passed_df['Test Case'].nunique()
                    prs_passed = passed_df['Traceability (PRS)'].nunique()
                else:
                    protocols_passed = prs_passed = 0

                # 4. Overall Failed Results (not all passed, no round logic)
                if 'Conclusion (Pass / Fail)' in records_df.columns:
                    valid_mask = (
                        records_df['Test Case'].str.len().gt(0) &
                        records_df['Traceability (PRS)'].str.len().gt(0)
                    )
                    df_valid = records_df[valid_mask].copy()
                    grouped = df_valid.groupby(['Traceability (PRS)', 'Test Case'])['Conclusion (Pass / Fail)'].apply(
                        lambda x: all(r.lower() == 'passed' for r in x)
                    )
                    failed_pairs = grouped[~grouped].reset_index()[['Traceability (PRS)', 'Test Case']]
                    protocols_failed = failed_pairs['Test Case'].nunique()
                    prs_failed = failed_pairs['Traceability (PRS)'].nunique()
                else:
                    protocols_failed = prs_failed = 0

                # 5. Overall Number of Open Anomalies classified as Defects
                open_defects = defect_df[defect_df[status_col] == 'open'][defect_id_col].nunique()

                parameter_data = [
                    [1, "Overall Test Case executed", f"{protocols_executed} (related to {prs_executed} PRS)"],
                    [2, "Overall Not Executed", f"{len(protocols_not_executed)} (related to {prs_not_executed} PRS)"],
                    [3, "Overall Passed Results", f"{protocols_passed} (related to {prs_passed} PRS)"],
                    [4, "Overall Failed Results", f"{protocols_failed} (related to {prs_failed} PRS)"],
                    [5, "Overall Number of Open Anomalies classified as Defects", f"{open_defects}"]
                ]
                parameter_data = [force_to_strings(row) for row in parameter_data]
                parameter_df = pd.DataFrame(parameter_data, columns=["#", "Parameter", "Value"])
                st.markdown("#### Parameter Table")
                st.dataframe(parameter_df, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer) as writer:
                    parameter_df.to_excel(writer, sheet_name='Parameter', index=False)
                buffer.seek(0)
                st.success("Product Verification Report generated! Download your results below.")
                st.download_button(
                    "⬇️ Download Parameter Table",
                    data=buffer,
                    file_name="product_verification_report_parameter.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Please click on the 'Run Product Verification Report' button to generate the report.")
    else:
        st.warning("Please upload all required files first.")
