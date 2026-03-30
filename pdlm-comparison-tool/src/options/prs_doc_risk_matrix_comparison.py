import streamlit as st
import pandas as pd
import tempfile
import re
import io

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
    st.markdown("### PRS DOC x Risk Management Matrix TASY_ Comparison")
    st.info("Upload your PRS DOC and Risk Management Matrix files to compare TASY_ requirements.")
    st.warning("This tool will extract all TASY_ values from column AC of the Risk Management Matrix and compare them with the PRS IDs.")
    
    col1, col2 = st.columns(2)
    with col1:
        prs_file = st.file_uploader("Upload PRS DOC (.xlsx)", type=["xlsx"], key="prsdoc_ver")
    with col2:
        rm_file = st.file_uploader("Upload Risk Management Matrix (.xlsx)", type=["xlsx"], key="tm_ver")
    if prs_file and rm_file:
        if st.button("🔍 Run Comparison", key="run_prs_rm_ver"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_prs:
                tmp_prs.write(prs_file.read())
                prs_path = tmp_prs.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_rm:
                tmp_rm.write(rm_file.read())
                rm_path = tmp_rm.name
            xl_prs = pd.ExcelFile(prs_path)
            xl_rm = pd.ExcelFile(rm_path)
            prs_sheet = get_sheet_with_fallback(xl_prs, 'Risk Management Matrix')
            rm_sheet = get_sheet_with_fallback(xl_rm, 'Risk Management Matrix')
            
            # Ler o PRS ID da primeira planilha
            PRScol = get_clean_cols(prs_path, prs_sheet, 'PRS ID')
            PRScol = normalize_spaces(filter_ignored(PRScol, IGNORE_SENTENCES_VAL_TP_TM))
            
            # Lendo a coluna AC da Risk Management Matrix e filtrando apenas valores com "TASY_"
            st.info("Lendo valores da coluna AC (TASY_) da Risk Management Matrix...")
            
            # Carregar a planilha Risk Management Matrix
            rm_df = pd.read_excel(rm_path, sheet_name=rm_sheet, skiprows=2)
            
            # Verificar se a coluna AC existe
            columns_list = list(rm_df.columns)
            if 'AC' not in columns_list:
                # Se não encontrar a coluna pelo nome, tentar pelo índice (AC = coluna 29)
                if len(columns_list) > 28:  # Verificar se há pelo menos 29 colunas (0-28)
                    col_ac = columns_list[28]  # Índice 28 corresponde à coluna AC (0-indexado)
                    st.info(f"Coluna AC não encontrada pelo nome, usando a coluna {col_ac} pelo índice.")
                else:
                    st.error("Coluna AC não encontrada na planilha e não há colunas suficientes.")
                    st.stop()
            else:
                col_ac = 'AC'
                
            # Extrair valores da coluna AC que contêm "TASY_"
            if col_ac in rm_df.columns:
                rm_tasy_values = rm_df[col_ac].astype(str).str.strip()
                # Filtrar apenas valores que contêm TASY_
                rm_tasy_values = rm_tasy_values[rm_tasy_values.str.contains('TASY_', case=False, na=False)]
                
                # Extrair todos os valores TASY_ de cada célula, pode haver múltiplos em uma célula
                all_tasy_values = []
                for val in rm_tasy_values:
                    # Encontrar todos os padrões "TASY_XXX" na string
                    tasy_matches = re.findall(r'TASY_\w+', val)
                    all_tasy_values.extend(tasy_matches)
                
                # Remover duplicatas e normalizar
                RMcol = pd.Series(list(set(all_tasy_values)))
                RMcol = normalize_spaces(RMcol)
                
                st.info(f"Encontrados {len(RMcol)} valores únicos com 'TASY_'.")
            # Adicionar informações sobre os resultados da comparação
            st.info("Comparando valores TASY_ da Risk Management Matrix com PRS IDs...")
            
            # Converter para conjuntos para operações de conjunto
            prs_set = set(PRScol)
            rm_set = set(RMcol)
            
            prs_results = {
                'PRS Only in PRS DOC': sorted(prs_set - rm_set),
                'TASY_ Only in RM': sorted(rm_set - prs_set),
                'Common to Both': sorted(prs_set & rm_set)
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
            for result_dict in [prs_results]:
                max_len = max(len(force_to_strings(v)) for v in result_dict.values())
                for k in result_dict:
                    clean_list = force_to_strings(result_dict[k])
                    if not clean_list:
                        clean_list = [""]
                # Preencher até o tamanho máximo
                    result_dict[k] = clean_list + [""] * (max_len - len(clean_list))

            # Agora pode criar os DataFrames sem risco de erro
            prs_df = pd.DataFrame(prs_results)
            st.write("TASY_ vs PRS ID Comparison")
            st.dataframe(prs_df, use_container_width=True)
            
            # Adicionar estatísticas
            st.markdown("### Statistics")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("PRS Only in PRS DOC", len(prs_set - rm_set))
            with col2:
                st.metric("TASY_ Only in RM", len(rm_set - prs_set))
            with col3:
                st.metric("Common to Both", len(prs_set & rm_set))
            
            # Preparar download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                prs_df.to_excel(writer, sheet_name='TASY_PRS_COMP', index=False)
            buffer.seek(0)
            
            st.success("Comparison complete! Download your results below.")
            st.download_button(
                "⬇️ Download Excel",
                data=buffer,
                file_name="TASY_PRS_comparison_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )