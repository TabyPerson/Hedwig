import zipfile
import os
import mimetypes
from bs4 import BeautifulSoup
import streamlit as st

def is_valid_media(filename):
    mime, _ = mimetypes.guess_type(filename)
    return mime and (mime.startswith('image/') or mime.startswith('video/'))

def run_comparison():
    st.title("Comparação de Evidências")
    col1, col2 = st.columns(2)
    with col1:
        records_file = st.file_uploader("Upload records.html", type=["html"], key="records_html")
    with col2:
        evidencies_file = st.file_uploader("Upload evidencies.zip", type=["zip"], key="evidencies_zip")

    if records_file and evidencies_file:
        if st.button("🔍 Run Comparison", key="run_html_html"):
            # Salva arquivos temporários
            records_path = "temp_records.html"
            evidencies_path = "temp_evidencies.zip"
            with open(records_path, "wb") as f:
                f.write(records_file.read())
            with open(evidencies_path, "wb") as f:
                f.write(evidencies_file.read())

            # Lê o HTML e busca os arquivos de evidência
            with open(records_path, encoding='utf-8') as f:
                soup = BeautifulSoup(f, 'html.parser')
            evidence_files = []
            for tr in soup.find_all('tr'):
                th = tr.find('th')
                td = tr.find('td')
                if th and 'Objective Evidence:' in th.text and td:
                    value = td.text.strip().split(' ')[0]
                    if value != "N/A":
                        evidence_files.append(value)

            # Extrai o ZIP e verifica os arquivos
            with zipfile.ZipFile(evidencies_path, 'r') as zip_ref:
                zip_ref.extractall('temp_evidencies')
                results = []
                for evidence in evidence_files:
                    found = False
                    for root, _, files in os.walk('temp_evidencies'):
                        if evidence in files:
                            evidence_path = os.path.join(root, evidence)
                            found = True
                            if is_valid_media(evidence_path):
                                results.append(f'{evidence}: válido ✅')
                            elif evidence.lower().endswith('.zip'):
                                # Extrai o zip interno e procura imagens/vídeos válidos
                                inner_extract_path = os.path.join(root, evidence + "_extracted")
                                os.makedirs(inner_extract_path, exist_ok=True)
                                with zipfile.ZipFile(evidence_path, 'r') as inner_zip:
                                    inner_zip.extractall(inner_extract_path)
                                valid_found = False
                                for iroot, _, ifiles in os.walk(inner_extract_path):
                                    for ifile in ifiles:
                                        if is_valid_media(os.path.join(iroot, ifile)):
                                            results.append(f'{evidence}: contém arquivo válido ({ifile}) ✅')
                                            valid_found = True
                                if not valid_found:
                                    results.append(f'{evidence}: zip encontrado, mas não contém imagem/vídeo válido ❌')
                            else:
                                results.append(f'{evidence}: encontrado, mas não é imagem/vídeo válido ❌')
                            break
                    if not found:
                        results.append(f'{evidence}: não encontrado no ZIP ❌')
                st.write("## Resultados")
                for r in results:
                    st.write(r)

            # Limpeza
            os.remove(records_path)
            os.remove(evidencies_path)
            # Opcional: remover arquivos extraídos

# Para rodar: streamlit run src/options/cmdk_records_evidences_comparison.py