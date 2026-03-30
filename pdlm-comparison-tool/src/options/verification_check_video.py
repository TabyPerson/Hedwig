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
    exts = r"mp4|avi|mov|wmv|mkv|wmv|exe|flv|webm|mpeg|mpg|pdf|docx|xlsx|txt|jpg|png|csv"
    # Regex for Windows and Unix paths (basic)
    pattern = rf'(\\\\[^\n\r,;]+?\.(?:{exts}))|([A-Za-z]:\\[^\n\r,;]+?\.(?:{exts}))|(/[^ \n\r,;]+?\.(?:{exts}))'
    matches = re.findall(pattern, text, re.IGNORECASE)
    # Flatten and filter empty
    paths = [m[0] or m[1] or m[2] for m in matches if any(m)]
    return [p.strip() for p in paths]

def extract_path_after_dot(text):
    """
    Extracts file paths after a dot (.) in a string, for verification check video.
    Melhorado para capturar caminhos em textos após pontos finais.
    """
    # Padrão para caminhos Windows e Unix mais abrangente
    pattern = r'([A-Za-z]:\\[^\s,;"\'\[\]<>]+(?:\.[a-zA-Z0-9]+)?)|(/[^ \n\r\t,;"\'\[\]<>]+(?:\.[a-zA-Z0-9]+)?)'
    matches = re.findall(pattern, text)
    paths = [m[0] if m[0] else m[1] for m in matches if m[0] or m[1]]
    
    # Padrão adicional para caminhos UNC (\\server\share)
    unc_pattern = r'(\\\\[^ \n\r\t,;"\'\[\]<>]+(?:\.[a-zA-Z0-9]+)?)'
    unc_matches = re.findall(unc_pattern, text)
    paths.extend(unc_matches)
    
    # Filtra caminhos que pareçam válidos (contém uma extensão de arquivo)
    filtered_paths = []
    for path in paths:
        # Verifica se o caminho contém uma extensão comum de arquivo
        if re.search(r'\.(mp4|avi|mov|wmv|mkv|flv|webm|mpeg|mpg|exe|pdf|docx|xlsx|jpg|png)$', path.lower()):
            filtered_paths.append(path)
    
    return filtered_paths

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
    st.markdown("### Check Video Evidence Paths (PRS MD = 'S')")
    st.info("Upload your Verification Test Records file to check and download a report of missing or invalid video/file paths in the 'Actual Result (Description)' column for rows where PRS MD = 'S'. Caminhos de vídeo após um ponto final serão verificados.")
    # Generate a unique ID for this function execution to avoid duplicate keys
    session_id = id(run_comparison)
    tr_file = st.file_uploader("Upload Verification Test Records (.xlsx)", type=["xlsx"], key=f"tr_val_video_{session_id}")
    if tr_file:
        if st.button("🔍 Run Video Analysis", key=f"start_video_analysis_{session_id}"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tr:
                tmp_tr.write(tr_file.read())
                tr_path = tmp_tr.name
            xl_tr = pd.ExcelFile(tr_path)
            if "Test Case Report" not in xl_tr.sheet_names:
                st.error("Sheet 'Test Case Report' not found in the uploaded file.")
                st.stop()
            # Carregar o arquivo Excel
            df = pd.read_excel(tr_path, sheet_name="Test Case Report")
            
            # Verificar se a coluna PRS MD existe
            if 'PRS MD' not in df.columns:
                st.error("Coluna 'PRS MD' não encontrada no arquivo. Por favor, verifique se o arquivo contém esta coluna.")
                st.stop()
            
            # Filtrar apenas linhas onde PRS MD = 'S'
            df_filtered = df[df['PRS MD'].astype(str).str.strip().str.upper() == 'S']
            
            if len(df_filtered) == 0:
                st.warning("Nenhuma linha com PRS MD = 'S' foi encontrada no arquivo.")
                st.stop()
            
            st.success(f"Encontradas {len(df_filtered)} linhas com PRS MD = 'S'. Analisando apenas estas linhas.")
                
            # Verificar se a coluna 'Actual Result (Description)' existe
            if 'Actual Result (Description)' not in df_filtered.columns:
                st.error("Column 'Actual Result (Description)' not found in the file.")
                st.stop()
                
            # Identify rows where the column is blank, NaN, or NaT (somente nas linhas filtradas)
            blank_mask = df_filtered['Actual Result (Description)'].isnull() | \
                         df_filtered['Actual Result (Description)'].astype(str).str.strip().isin(['', 'nan', 'NaT'])

            blank_rows = df_filtered[blank_mask]

            if not blank_rows.empty:
                st.warning(f"As seguintes {len(blank_rows)} linhas com PRS MD = 'S' têm um valor vazio ou inválido em 'Actual Result (Description)':")
                for idx in blank_rows.index:
                    st.info(f"Row {idx + 2}: is blank in 'Actual Result (Description)'")  # +2 for Excel-like row number

            ignore_sentences = [
                "Note: * No Defect/Enhancement raised. Refers to test cases that were approved and for this reason, no defect was raised.",
                "Note: Defect/Enhancement are linked to test cases and their respective execution versions"
            ]
                
            # Filtrar sentenças de nota apenas no dataframe filtrado
            df_filtered = df_filtered[~df_filtered['Actual Result (Description)'].astype(str).str.strip().str.lower().isin(
                [s.strip().lower() for s in ignore_sentences]
            )]
                
            # Lista de extensões de vídeo para verificar
            video_exts = ['.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv', '.webm', '.mpeg', '.exe', '.mpg']
            analysis = []
            total_files_gt_zero = 0  # Counter for files > 0 bytes
            total_video_bytes = 0    # Counter for total size of video files not in zip

            # Usar df_filtered em vez de df para analisar apenas linhas onde PRS MD = 'S'
            for idx, row in df_filtered.iterrows():
                text = str(row['Actual Result (Description)'])
                
                # Extrair caminhos de vídeo após pontos finais
                # Dividir o texto em frases (separadas por pontos)
                sentences = text.split('.')
                
                # Extrair possíveis caminhos de cada frase (exceto a primeira)
                file_paths = []
                for sentence in sentences[1:]:  # Ignora a primeira parte antes do ponto
                    paths = extract_paths(sentence.strip())
                    if paths:
                        file_paths.extend(paths)
                    else:
                        # Se não encontrou com extract_paths, tenta com extract_path_after_dot
                        paths = extract_path_after_dot(sentence.strip())
                        file_paths.extend(paths)
                
                # Se não encontrou caminhos após pontos, tenta no texto completo
                if not file_paths:
                    paths = extract_paths(text)
                    if paths:
                        file_paths.extend(paths)
                    else:
                        paths = extract_path_after_dot(text)
                        file_paths.extend(paths)
                
                # Remover duplicatas e entradas vazias
                file_paths = [p for p in file_paths if p.strip()]
                file_paths = list(dict.fromkeys(file_paths))  # Remove duplicatas preservando a ordem
                    
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

                # Continuamos acumulando as informações em analysis para cada linha
                total_files_gt_zero += row_total_files_gt_zero

            # Após o loop por todas as linhas, exibimos os resultados consolidados
            for row in analysis:
                row['Opened'] = row['Total Valid Videos'] > 0

            # Calcular estatísticas finais
            valid_videos = [row for row in analysis if row['Total Valid Videos'] > 0]
            not_found = [row for row in analysis if row['Total Valid Videos'] == 0]
            total_rows_analyzed = len(df_filtered)
            total_valid_videos = sum(row['Total Valid Videos'] for row in analysis)
            not_found_count = len(not_found)
            
            # Conta quantas linhas têm múltiplos caminhos de vídeo
            multi_path_rows = sum(1 for row in analysis if len(str(row['Evidence Path']).split(',')) > 1)
            
            st.info(
                f"Total Rows com PRS MD = 'S': {total_rows_analyzed} | "
                f"Linhas com múltiplos caminhos: {multi_path_rows} | "
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
                    st.write("### Arquivos Não Encontrados ou Inválidos")
                    for row in not_found:
                        st.error(f"❌ Row {row['Row']}: {row['Evidence Path']} (Arquivo não encontrado, não é um arquivo de vídeo válido, ou é um ZIP sem vídeos)")
                
                if valid_videos:
                    st.write("### Arquivos de Vídeo Válidos")
                    for row in valid_videos:
                        if row['Is Zip'] and row['Zip Contains Video']:
                            st.success(f"✅ Row {row['Row']}: {row['Evidence Path']} (ZIP contém vídeo(s): {', '.join(row['Zip Video Names'])})")
                        elif row['Exists'] and row['Is Video']:
                            st.success(f"✅ Row {row['Row']}: {row['Evidence Path']} (Arquivo de vídeo válido)")
                
                if not not_found and not valid_videos:
                    st.success("Todos os arquivos em 'Actual Result (Description)' foram encontrados e são arquivos de vídeo válidos.")
                
                st.markdown('</div>', unsafe_allow_html=True)

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
            
            # Criamos um timestamp único para evitar duplicação de chaves
            timestamp = pd.Timestamp.now().strftime('%Y%m%d%H%M%S%f')
            
            buffer = io.BytesIO()
            result_df.to_excel(buffer, index=False)
            buffer.seek(0)
            
            st.download_button(
                "⬇️ Download Full Actual Result Video Analysis (Excel)",
                data=buffer,
                file_name="actual_result_video_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_actual_result_analysis_{timestamp}"  # Usando timestamp para garantir chave única
            )