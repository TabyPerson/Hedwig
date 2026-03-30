import streamlit as st
import pandas as pd
import tempfile
from utils.file_utils import get_clean_cols, normalize_spaces, filter_ignored
from utils.comparison_utils import get_sheet_with_fallback

IGNORE_SENTENCES = [
    'Note (N/A*): The columns “Actual Result /Description”, "Version tested", "Date tested", "Tester", "Conclusion (Pass/Fail)", "Defect/Enhancement Number", and "Defect/Enhancement Status" from the table below are a placeholder to record the Test Results. Once verification activities are completed, this document will be updated.',
    'Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.',
    'Note: Defect/Enhancement are linked to test cases and their respective execution versions',
    'NA* Closed Service Orders'
]

def run_tm_val_tp_comparison():
    st.markdown("### TM APP x Validation Test Protocol Comparison")
    st.info("Upload your TM APP and Validation Test Protocol files to compare test protocols.")
    col1, col2 = st.columns(2)
    with col1:
        tm_file = st.file_uploader("Upload TM APP (.xlsx)", type=["xlsx"], key="tmapp_val")
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
            tm_sheet = get_sheet_with_fallback(xl_tm, 'Design Validation')
            tp_sheet = get_sheet_with_fallback(xl_tp, 'Test Case Report')
            TMcol, TMurs = get_clean_cols(tm_path, tm_sheet, 'Validation Test ID', 'URS Requirement ID', skiprows=1)
            TPcol, TPurs = get_clean_cols(tp_path, tp_sheet, 'Validation Protocol ID', 'URS')
            TMcol = normalize_spaces(filter_ignored(TMcol, IGNORE_SENTENCES))
            TMurs = normalize_spaces(filter_ignored(TMurs, IGNORE_SENTENCES))
            TPcol = normalize_spaces(filter_ignored(TPcol, IGNORE_SENTENCES))
            TPurs = normalize_spaces(filter_ignored(TPurs, IGNORE_SENTENCES))
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
            tc_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in tc_results.items()]))
            urs_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in urs_results.items()]))
            st.write("Test Case Comparison")
            st.dataframe(tc_df)
            st.write("URS Comparison")
            st.dataframe(urs_df)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                tc_df.to_excel(writer, sheet_name='Protocol COMP', index=False)
                urs_df.to_excel(writer, sheet_name='URS COMP', index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="VAL_TPxTM_T_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    run_tm_val_tp_comparison()