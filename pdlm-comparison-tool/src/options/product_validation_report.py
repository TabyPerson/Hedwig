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
    st.markdown("### Product Validation Report")
    st.info("Upload your Product Validation Protocol, Product Validation Records, and Product Defect Status Report files.")

    col1, col2, col3 = st.columns(3)
    with col1:
        protocol_file = st.file_uploader("Upload Product Validation Protocol (.xlsx)", type=["xlsx"], key="pvr_protocol")
    with col2:
        records_file = st.file_uploader("Upload Product Validation Records (.xlsx)", type=["xlsx"], key="pvr_records")
    with col3:
        defect_file = st.file_uploader("Upload Product Defect Status Report (.xlsx)", type=["xlsx"], key="pvr_defect")

    if protocol_file and records_file and defect_file:
        if st.button("🔍 Run Product Validation Report", key="run_pvr"):
            # Load Protocol
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_protocol:
                tmp_protocol.write(protocol_file.read())
                protocol_path = tmp_protocol.name
            xl_protocol = pd.ExcelFile(protocol_path)
            protocol_sheet = get_sheet_with_fallback(xl_protocol, 'Test Case Report')
            protocol_df = pd.read_excel(protocol_path, sheet_name=protocol_sheet)
            protocol_df['Validation Protocol ID'] = protocol_df['Validation Protocol ID'].astype(str).fillna('')
            protocol_df['URS'] = protocol_df['URS'].astype(str).fillna('')
            if 'URS' in protocol_df.columns:
                protocol_df['URS'] = normalize_spaces(protocol_df['URS'].astype(str))
            else:
                protocol_df['URS'] = ""

            # Load Records
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_records:
                tmp_records.write(records_file.read())
                records_path = tmp_records.name
            xl_records = pd.ExcelFile(records_path)
            records_sheet = get_sheet_with_fallback(xl_records, 'Test Case Report')
            records_df = pd.read_excel(records_path, sheet_name=records_sheet)
            records_df['Validation Protocol ID'] = records_df['Validation Protocol ID'].astype(str).fillna('')
            records_df['URS'] = records_df['URS'].astype(str).fillna('')
            if 'URS' in records_df.columns:
                records_df['URS'] = normalize_spaces(records_df['URS'].astype(str))
            else:
                 records_df['URS'] = ""

            # Load Defect Status Report
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_defect:
                tmp_defect.write(defect_file.read())
                defect_path = tmp_defect.name

            xl_defect = pd.ExcelFile(defect_path)

            # Explicitly select the correct sheet
            defect_sheet = "Defect List"
            if defect_sheet not in xl_defect.sheet_names:
                st.error(f"Sheet '{defect_sheet}' not found in the uploaded file. Available sheets: {xl_defect.sheet_names}")
                st.stop()

            # Read the sheet (adjust skiprows if your header is not the first row)
            defect_df = pd.read_excel(defect_path, sheet_name=defect_sheet, skiprows=2)

            # Normalize column names: remove line breaks and extra spaces, lowercase
            defect_df.columns = defect_df.columns.str.replace(r'\s+', ' ', regex=True).str.strip().str.lower()

            # Find the correct columns
            status_col = None
            defect_id_col = None
            for col in defect_df.columns:
                if 'status' in col:
                    status_col = col
                if 'defect id' in col:
                    defect_id_col = col

            if not status_col or not defect_id_col:
                st.error(f"Could not find required columns in Defect Status Report. Columns found: {defect_df.columns.tolist()}")
            else:
                defect_df[status_col] = defect_df[status_col].astype(str).str.strip().str.lower()
                #st.text(f"Unique status values: {defect_df[status_col].unique()}")
                open_defects = defect_df[defect_df[status_col] == 'open'][defect_id_col].nunique()

            # Boolean mask functions
            def mask_protocol_id_column(protocol_series):
                return protocol_series.str.strip().str.startswith('TASY_VTC', na=False)

            def mask_urs_column(urs_series):
                return urs_series.str.strip().str.startswith('A', na=False)

            # Clean records_df for both columns using boolean masks
            records_df_clean = records_df[
                mask_protocol_id_column(records_df['Validation Protocol ID']) &
                mask_urs_column(records_df['URS'])
            ]

            # 1. Overall Protocol executed (in records)
            protocols_executed = records_df_clean['Validation Protocol ID'].nunique()
            urs_executed = records_df_clean['URS'].nunique()

            # 2. Overall Not Executed (in protocol but not in records)
            protocol_ids_in_protocol = protocol_df.loc[mask_protocol_id_column(protocol_df['Validation Protocol ID']), 'Validation Protocol ID']
            protocol_ids_in_records = records_df.loc[mask_protocol_id_column(records_df['Validation Protocol ID']), 'Validation Protocol ID']
            protocols_not_executed = set(protocol_ids_in_protocol) - set(protocol_ids_in_records)
            urs_not_executed = protocol_df.loc[
                protocol_df['Validation Protocol ID'].isin(protocols_not_executed) & mask_urs_column(protocol_df['URS']),
                'URS'
            ].nunique() if 'URS' in protocol_df.columns else 0

            # 3. Overall Passed Results (Result == Pass in records)
            records_df['Result (Pass/Fail)'] = records_df['Result (Pass/Fail)'].astype(str).str.strip()
            records_df['Validation Protocol ID'] = records_df['Validation Protocol ID'].astype(str).str.strip()
            records_df['URS'] = records_df['URS'].astype(str).str.strip()
            passed_mask = (
                records_df['Result (Pass/Fail)'].str.lower().isin(['pass', 'passed']) &
                records_df['Validation Protocol ID'].str.len().gt(0) &
                records_df['URS'].str.len().gt(0)
            )
            passed_df = records_df[passed_mask]
            protocols_passed = passed_df['Validation Protocol ID'].nunique()
            urs_passed = passed_df['URS'].nunique()

            # 4. Overall Failed Results (latest round not all passed)
            records_df['Round'] = records_df['Round'].astype(str).str.strip() if 'Round' in records_df.columns else ""
            if 'Result (Pass/Fail)' in records_df.columns and 'Round' in records_df.columns:
                valid_mask = (
                    records_df['Validation Protocol ID'].str.len().gt(0) &
                    records_df['URS'].str.len().gt(0) &
                    records_df['Round'].str.len().gt(0)
                )
                df_valid = records_df[valid_mask].copy()
                df_valid['Round_sort'] = pd.to_numeric(df_valid['Round'], errors='coerce')
                latest_round_idx = df_valid.groupby(['URS', 'Validation Protocol ID'])['Round_sort'].transform('max') == df_valid['Round_sort']
                latest_df = df_valid[latest_round_idx]
                grouped = latest_df.groupby(['URS', 'Validation Protocol ID'])['Result (Pass/Fail)'].apply(
                    lambda x: all(r.lower() == 'passed' for r in x)
                )
                failed_pairs = grouped[~grouped].reset_index()[['URS', 'Validation Protocol ID']]
                protocols_failed = failed_pairs['Validation Protocol ID'].nunique()
                urs_failed = failed_pairs['URS'].nunique()
            else:
                protocols_failed = urs_failed = 0

            # 5. Overall Number of Open Anomalies classified as Defects
            open_defects = 0
            if defect_id_col and status_col:
                defect_df[status_col] = defect_df[status_col].astype(str).str.strip().str.lower()
                open_defects = defect_df[defect_df[status_col] == 'open'][defect_id_col].nunique()

            parameter_data = [
                [1, "Overall Protocol executed", f"{protocols_executed} (related to {urs_executed} URS)"],
                [2, "Overall Not Executed", f"{len(protocols_not_executed)} (related to {urs_not_executed} URS)"],
                [3, "Overall Passed Results", f"{protocols_passed} (related to {urs_passed} URS)"],
                [4, "Overall Failed Results", f"{protocols_failed} (related to {urs_failed} URS)"],
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
            st.success("Product Validation Report generated! Download your results below.")
            st.download_button(
                "⬇️ Download Parameter Table",
                data=buffer,
                file_name="product_validation_report_parameter.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )