import streamlit as st
import pandas as pd
import re
import tempfile
import io
import re
import os
import zipfile
import numpy as np


def extract_paths(text):
    """
    Extracts file paths from a string. Looks for Windows and Unix-like paths.
    """
    # List of common video/file extensions you want to match
    exts = r"mp4|avi|mov|wmv|mkv|flv|webm|mpeg|mpg|pdf|docx|xlsx|txt|jpg|png|csv"
    # Regex for Windows and Unix paths (basic)
    pattern = rf'(\\\\[^\n\r,;]+?\.(?:{exts}))|([A-Za-z]:\\[^\n\r,;]+?\.(?:{exts}))|(/[^ \n\r,;]+?\.(?:{exts}))'
    matches = re.findall(pattern, text, re.IGNORECASE)
    # Flatten and filter empty
    paths = [m[0] or m[1] or m[2] for m in matches if any(m)]
    return [p.strip() for p in paths]

def extract_path_after_dot(text):
    """
    Extracts file paths after a dot (.) in a string, for verification check video.
    """
    # This is a simple heuristic; adjust as needed for your file path patterns
    pattern = r'([A-Za-z]:\\[^\s,;]+(?:\.[a-zA-Z0-9]+)?)|(/[^ \n\r\t,;]+(?:\.[a-zA-Z0-9]+)?)'
    matches = re.findall(pattern, text)
    paths = [m[0] if m[0] else m[1] for m in matches if m[0] or m[1]]
    return paths

def normalize_spaces(series):
    return series.astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def filter_invalid_paths(series):
    # This function checks for valid file paths in the series
    valid_paths = []
    for path in series:
        if re.match(r'^(?:[A-Za-z]:\\|/)', path):  # Basic check for Windows or Unix paths
            valid_paths.append(path)
    return valid_paths


def run_comparison():
    st.markdown("### Check Video Evidence Paths")
    st.info("Upload your Validation Test Records file to check and download a report of missing or invalid video/file paths in the 'Actual result / Evidence Path' column.")
    tr_file = st.file_uploader("Upload Validation Test Records (.xlsx)", type=["xlsx"], key="tr_val_video")
    if tr_file:
        if st.button("🔍 Run Video Analysis", key="start_video_analysis"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name
            xl_tr = pd.ExcelFile(tr_path)
            if "Test Case Report" not in xl_tr.sheet_names:
                st.error("Sheet 'Test Case Report' not found in the uploaded file.")
                st.stop()
            df = pd.read_excel(tr_path, sheet_name="Test Case Report")
            # Identify rows where the column is blank, NaN, or NaT
            blank_mask = df['Actual result / Evidence Path'].isnull() | \
                         df['Actual result / Evidence Path'].astype(str).str.strip().isin(['', 'nan', 'NaT'])

            blank_rows = df[blank_mask]

            if not blank_rows.empty:
                st.warning("The following rows have a blank, NaN, or NaT value in 'Actual result / Evidence Path':")
                for idx in blank_rows.index:
                    st.info(f"Row {idx + 2}: is blank in 'Actual result / Evidence Path'")  # +2 for Excel-like row number

            ignore_sentences = [
                "Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.",
                "Note: Defect/Enhancement are linked to test cases and their respective execution versions"
            ]
            if 'Actual result / Evidence Path' not in df.columns:
                st.error("Column 'Actual result / Evidence Path' not found in the file.")
                df = df[~df['Actual result / Evidence Path'].astype(str).str.strip().str.lower().isin(
                    [s.strip().lower() for s in ignore_sentences]
                )]
            else:
                video_exts = ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv', '.webm', '.mpeg', '.exe', '.mpg']
                analysis = []
                total_files_gt_zero = 0  # Counter for files > 0 bytes
                total_video_bytes = 0    # Counter for total size of video files not in zip

                for idx, row in df.iterrows():
                    text = str(row['Actual result / Evidence Path'])
                    file_paths = [text.strip()] if text.strip() else []
                        
                    row_exists = False
                    row_is_video = False
                    row_is_zip = False
                    row_zip_contains_video = False
                    row_zip_video_names = []
                    row_total_valid_videos = 0
                    row_total_files_gt_zero = 0

                    if not file_paths:
                        analysis.append({
                            'Row': idx + 2,
                            'Evidence Path': '',
                            'Exists': False,
                            'Is Video': False,
                            'Is Zip': False,
                            'Zip Contains Video': False,
                            'Opened': False,
                            'Zip Video Names': [],
                            'Total Valid Videos': 0,
                            'Total Files > 0 Bytes': 0,
                            'Total Video Bytes': 0
                        })
                        continue
                                              
                    for path in file_paths:
                        is_video = any(path.lower().endswith(ext) for ext in video_exts)
                        is_zip = path.lower().endswith('.zip')
                        exists = os.path.exists(path) and os.path.isfile(path) if not is_zip else True

                        zip_contains_video = False
                        zip_video_names = []
                        total_valid_videos = 0
                        total_files_gt_zero_in_path = 0

                        if is_zip:
                            try:
                                with zipfile.ZipFile(path, 'r') as z:
                                    for name in z.namelist():
                                        if any(name.lower().endswith(ext) for ext in video_exts):
                                            size = z.getinfo(name).file_size
                                            if size > 0:
                                                zip_contains_video = True
                                                zip_video_names.append(name)
                                                total_files_gt_zero_in_path += 1
                                total_valid_videos = len(zip_video_names)
                            except Exception as e:
                                st.error(f"Exception for ZIP {repr(path)}: {e}")
                                total_valid_videos = 0
                                zip_contains_video = False
                                zip_video_names = []
                                total_files_gt_zero_in_path = 0
                        else:
                            if exists and is_video:
                                try:
                                    with open(path, 'rb') as f:
                                        f.read(512)
                                    size = os.path.getsize(path)
                                    if size > 0:
                                        total_valid_videos = 1
                                        total_files_gt_zero = 1
                                        total_video_bytes += size
                                    else:
                                        total_valid_videos = 0
                                        total_files_gt_zero = 0
                                except Exception as e:
                                    st.error(f"Exception for video file {repr(path)}: {e}")
                                    total_valid_videos = 0
                                    total_files_gt_zero = 0
                            else:
                                total_valid_videos = 0
                                total_files_gt_zero = 0

                        # Aggregate for the row (if multiple paths in one cell)
                        row_exists = row_exists or exists
                        row_is_video = row_is_video or is_video
                        row_is_zip = row_is_zip or is_zip
                        row_zip_contains_video = row_zip_contains_video or zip_contains_video
                        row_zip_video_names.extend(zip_video_names)
                        row_total_valid_videos += total_valid_videos
                        row_total_files_gt_zero += total_files_gt_zero_in_path

                    total_files_gt_zero += row_total_files_gt_zero

                    analysis.append({
                        'Row': idx + 2,
                        'Evidence Path': ', '.join(file_paths),
                        'Exists': row_exists,
                        'Is Video': row_is_video,
                        'Is Zip': row_is_zip,
                        'Zip Contains Video': row_zip_contains_video,
                        'Opened': row_zip_contains_video or (row_exists and row_is_video),
                        'Zip Video Names': row_zip_video_names,
                        'Total Valid Videos': row_total_valid_videos,
                        'Total Files > 0 Bytes': row_total_files_gt_zero,
                        'Total Video Bytes': total_video_bytes
                    })

                valid_videos = [row for row in analysis if row['Total Valid Videos'] > 0]
                not_found = [row for row in analysis if row['Total Valid Videos'] == 0]
                total_rows_analyzed = len(df)
                total_valid_videos = sum(row['Total Valid Videos'] for row in analysis)
                not_found_count = len(not_found)
                st.info(
                    f"Total Rows Analyzed: {total_rows_analyzed} | "
                    f"Total Valid Videos: {total_valid_videos} | "
                    f"Files > 0 Bytes: {total_files_gt_zero} | "
                    f"Total Video Bytes (not in zip): {total_video_bytes:,} | "
                    f"Invalid/Not Found: {not_found_count}"
                )

                with st.container():
                    st.markdown("""
                    <style>
                        .scrollable-container {
                        max-height: 400px;
                        overflow-y: auto;
                        border: 1px solid #ddd;
                        padding: 10px;
                        background: #fafafa;
                    }
                    </style>
                    """, unsafe_allow_html=True)
                    st.markdown('<div class="scrollable-container">', unsafe_allow_html=True)
                    if not_found:
                        st.write("### Files Not Found or Not Valid Files")
                    for row in not_found:
                        st.error(f"❌ Row {row['Row']}: {row['Evidence Path']} (File not found, not a valid video file, or zip without video)")
                    if valid_videos:
                        st.write("### Valid Video Files")
                        for row in valid_videos:
                            if row['Is Zip'] and row['Zip Contains Video']:
                                st.success(f"✅ Row {row['Row']}: {row['Evidence Path']} (ZIP contains video(s): {', '.join(row['Zip Video Names'])})")
                            elif row['Exists'] and row['Is Video']:
                                st.success(f"✅ Row {row['Row']}: {row['Evidence Path']} (Valid video file)")
                    if not not_found and not valid_videos:
                        st.success("All files in 'Actual result / Evidence Path' were found and are valid video files.")
                    st.markdown('</div>', unsafe_allow_html=True)

                for row in analysis:
                    row['Opened'] = row['Total Valid Videos'] > 0

                summary_row = {
                    'Row': 'TOTAL',
                    'Evidence Path': '',
                    'Exists': '',
                    'Is Video': '',
                    'Is Zip': '',
                    'Zip Contains Video': '',
                    'Opened': '',
                    'Zip Video Names': '',
                    'Total Valid Videos': total_valid_videos,
                    'Total Files > 0 Bytes': total_files_gt_zero,
                    'Total Video Bytes': total_video_bytes
                }
                result_df = pd.DataFrame(analysis)
                result_df = pd.concat([result_df, pd.DataFrame([summary_row])], ignore_index=True)
                buffer = io.BytesIO()
                result_df.to_excel(buffer, index=False)
                buffer.seek(0)
                st.download_button(
                    "⬇️ Download Full Evidence Path Analysis (Excel)",
                    data=buffer,
                    file_name="evidence_path_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )