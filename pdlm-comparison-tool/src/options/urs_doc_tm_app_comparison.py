import streamlit as st
import pandas as pd
import tempfile
import io
import re
from docx import Document

   
def normalize_spaces(series):
    return pd.Series(series).astype(str).apply(lambda x: re.sub(r'\s+', ' ', x.strip()))

def get_feature_ids_from_word(doc_path, col_name='Feature ID'):
    doc = Document(doc_path)
    feature_ids = []
    for table in doc.tables[1:3]:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        try:
            idx = headers.index(col_name)
        except ValueError:
            st.error(f"Column '{col_name}' not found in URS DOC: {headers}")
            continue  # <-- Troque return [] por continue
        feature_ids += [row.cells[idx].text.strip() for row in table.rows[1:]]
    return pd.Series(feature_ids).drop_duplicates().dropna().astype(str).str.strip().tolist()  # <-- sempre retorna lista

def get_sheet_with_fallback(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return preferred_name
    else:
        st.warning(f"Worksheet named '{preferred_name}' not found. Please select the sheet that you want to analyze.")
        return st.selectbox("Select the worksheet that you want to analyze:", xl.sheet_names, key=preferred_name)

def get_urs_ids_from_excel(excel_path, sheet_name='Design Validation', col_name='URS Requirement ID'):
    xl = pd.ExcelFile(excel_path)
    # Use fallback para selecionar a worksheet
    sheet_to_use = get_sheet_with_fallback(xl, sheet_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_to_use, skiprows=1)
    if col_name not in df.columns:
        st.error(f"Column '{col_name}' not found in the sheet '{sheet_to_use}'.")
        return []
    return df[col_name].dropna().astype(str).str.strip().drop_duplicates().tolist()

def pad_list(lst, target_len):
    lst = lst if lst is not None else []
    if target_len is None or not isinstance(target_len, int) or target_len < 1:
        target_len = 1
    return lst + [""] * max(target_len - len(lst), 0)

def safe_str_list(lst):
    # Garante que cada elemento é string simples, nunca método, objeto estranho, lista ou slice
    out = []
    for x in lst:
        # Só aceita string, int ou float
        if isinstance(x, (str, int, float)):
            s = str(x)
            if (
                "method" in s
                or "descriptor" in s
                or s.startswith("<")
                or s.startswith("[")
                or s.startswith("(")
                or s.strip() == ""
            ):
                continue
            out.append(s)
        # Se não for, ignora
    return out

def run_comparison():
    st.markdown("### URS DOC x TM APP Comparison")
    st.info("Upload your URS DOC and TM APP files to compare requirements.")
    col1, col2 = st.columns(2)
    with col1:
        word_file = st.file_uploader("Upload URS DOC (.docx)", type=["docx"], key="ursdoc")
    with col2:
        excel_file = st.file_uploader("Upload TM APP (.xlsx)", type=["xlsx"], key="tmapp")
    if word_file and excel_file:
        if st.button("🔍 Run Comparison", key="run_urs_tm"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
                tmp_word.write(word_file.read())
                word_path = tmp_word.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                tmp_excel.write(excel_file.read())
                excel_path = tmp_excel.name
            feature_ids = set(normalize_spaces(get_feature_ids_from_word(word_path) or []).tolist())
            urs_ids = set(normalize_spaces(get_urs_ids_from_excel(excel_path) or []).tolist())

            only_in_word = sorted(list(feature_ids - urs_ids))
            only_in_excel = sorted(list(urs_ids - feature_ids))
            common = sorted(list(feature_ids & urs_ids))
            
            # Versão mais robusta para garantir apenas strings
            def force_to_strings(lst):
                result = []
                for item in lst:
                    try:
                        # Converter para string e verificar se é uma string válida
                        text = str(item)
                        # Filtrar strings que parecem objetos ou métodos
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
           
            # Garantir que temos apenas strings válidas
            only_in_word = [str(x) for x in only_in_word if x is not None]
            only_in_excel = [str(x) for x in only_in_excel if x is not None]
            # Garantir que temos pelo menos uma linha
            if not only_in_word:
                only_in_word = [""]
            if not only_in_excel:
                only_in_excel = [""]
            
            clean_common = []
            for x in common:
                try:
                    if isinstance(x, (str, int, float)):
                        text = str(x)
                        if (not "method" in text and 
                            not "descriptor" in text and
                            not text.startswith("<") and
                            not text.startswith("[") and
                            not text.startswith("(") and
                            text.strip() != ""):
                            clean_common.append(text)
                except:
                    pass

            common = clean_common
            if not common:
                common = [""]


            # Proteção extra: se todas as listas estiverem vazias, max_len = 1
            max_len = max(len(only_in_word), len(only_in_excel), len(common))
            
            # Preencher cada lista com strings vazias até o tamanho máximo
            only_in_word = only_in_word + [""] * (max_len - len(only_in_word))
            only_in_excel = only_in_excel + [""] * (max_len - len(only_in_excel))
            common = common + [""] * (max_len - len(common))            

            data = {
                'Only in URS DOC (URS)': only_in_word,
                'Only in TM APP (TM)': only_in_excel,
                'Common (URS ∩ TM)': common
            }
            
            summary_df = pd.DataFrame(data)
            st.dataframe(summary_df, use_container_width=True)
            buffer = io.BytesIO()
            summary_df.to_excel(buffer, index=False)
            buffer.seek(0)
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Result as Excel",
                data=buffer,
                file_name="URS_vs_TM_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )