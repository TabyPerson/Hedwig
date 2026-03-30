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

def safe_display_dataframe(df, columns=None, max_rows=10, msg="Dados do DataFrame"):
    """
    Exibe um DataFrame de forma segura, lidando com casos onde o DataFrame pode ser None,
    estar vazio ou ter problemas com as colunas.
    
    Args:
        df (DataFrame): O DataFrame a ser exibido
        columns (list): Lista de colunas para mostrar. Se None, mostra todas as colunas
        max_rows (int): Número máximo de linhas para mostrar
        msg (str): Mensagem a ser exibida antes do DataFrame
    """
    try:
        # Verificar se a mensagem é válida
        if msg is None:
            msg = "Dados do DataFrame"
            
        if df is None:
            st.warning(f"{msg}: O DataFrame é None")
            # Exibir uma tabela vazia usando método mais seguro
            st.write(msg)
            info_df = pd.DataFrame({"Informação": ["Sem dados disponíveis"]})
            try:
                st.table(info_df)  # table é mais simples e tem menos parâmetros que podem causar erros
            except Exception as table_err:
                st.write(f"Não foi possível exibir a tabela: {str(table_err)}")
                st.write("Sem dados disponíveis")
            return
            
        if len(df) == 0:
            st.warning(f"{msg}: O DataFrame está vazio")
            st.write(msg)
            info_df = pd.DataFrame({"Informação": ["DataFrame vazio"]})
            try:
                st.table(info_df)  # table é mais simples e tem menos parâmetros que podem causar erros
            except Exception as table_err:
                st.write("DataFrame vazio")
            return
            
        # Criar uma cópia segura para não modificar o DataFrame original
        try:
            safe_df = df.copy()
        except Exception as copy_error:
            st.error(f"Erro ao copiar DataFrame: {str(copy_error)}")
            # Tentar criar um DataFrame novo com os dados disponíveis
            try:
                # Converter para dict e depois para DataFrame para "desconectar" do original
                safe_df = pd.DataFrame(df.to_dict())
            except Exception:
                # Se falhar completamente, criar um DataFrame com mensagem de erro
                st.error("Não foi possível processar o DataFrame")
                info_df = pd.DataFrame({"Erro": ["Falha ao processar o DataFrame"]})
                st.table(info_df)
                return
        
        # Determinar colunas para mostrar com tratamento de erro melhorado
        display_df = None
        try:
            if columns is not None:
                # Filtrar apenas colunas que existem no DataFrame
                valid_cols = [col for col in columns if col in safe_df.columns]
                if not valid_cols:
                    # Se nenhuma das colunas solicitadas existir, pegue as 5 primeiras
                    valid_cols = safe_df.columns[:min(5, len(safe_df.columns))].tolist()
                    
                # Criar um novo DataFrame apenas com as colunas desejadas    
                display_df = pd.DataFrame()
                for col in valid_cols:
                    if col in safe_df.columns:  # Verificação extra de segurança
                        display_df[col] = safe_df[col].copy()
            else:
                # Limitar para as primeiras 5 colunas para não sobrecarregar a visualização
                display_df = pd.DataFrame()
                cols_to_show = safe_df.columns[:min(5, len(safe_df.columns))].tolist()
                for col in cols_to_show:
                    if col in safe_df.columns:  # Verificação extra de segurança
                        display_df[col] = safe_df[col].copy()
        except Exception as cols_error:
            st.error(f"Erro ao processar colunas: {str(cols_error)}")
            # Fallback: criar DataFrame com todas as colunas
            try:
                display_df = safe_df.copy()
            except Exception:
                # Se ainda falhar, mostrar mensagem de erro
                st.error("Falha ao processar colunas do DataFrame")
                return
        
        # Verificar se o display_df foi criado com sucesso
        if display_df is None or len(display_df.columns) == 0:
            st.warning("Não foi possível criar um DataFrame para exibição")
            st.write(msg)
            info_df = pd.DataFrame({"Aviso": ["Não foi possível processar as colunas"]})
            st.table(info_df)
            return
        
        # Converter todas as colunas para string para evitar problemas de exibição
        for col in display_df.columns:
            try:
                display_df[col] = display_df[col].fillna("").astype(str)
            except Exception as convert_error:
                st.warning(f"Erro ao converter coluna '{col}': {str(convert_error)}")
                # Substituir a coluna problemática por uma mensagem de erro
                display_df[col] = "[Erro de conversão]"
        
        # Limitar o número de linhas de forma segura
        if len(display_df) > 0:
            try:
                if max_rows is not None and len(display_df) > max_rows:
                    # Usar iloc é mais seguro que head() para evitar problemas com índices
                    display_df = display_df.iloc[:max_rows].copy()
                
                # Resetar o índice para evitar problemas
                display_df = display_df.reset_index(drop=True)
            except Exception as rows_error:
                st.warning(f"Erro ao limitar linhas: {str(rows_error)}")
                # Ignorar o erro e tentar mostrar mesmo assim
            
            # Exibir o DataFrame com múltiplas tentativas de fallback
            st.write(msg)
            try:
                # Primeira tentativa: dataframe normal
                st.dataframe(display_df, use_container_width=True)
            except Exception as e1:
                st.error(f"Erro ao exibir como dataframe: {str(e1)}")
                try:
                    # Segunda tentativa: sem use_container_width
                    st.dataframe(display_df)
                except Exception as e2:
                    st.error(f"Segundo erro ao exibir: {str(e2)}")
                    try:
                        # Terceira tentativa: usar st.table que tem menos opções
                        st.table(display_df.head())
                    except Exception as e3:
                        st.error(f"Não foi possível exibir dados: {str(e3)}")
                        # Última tentativa: exibir como texto
                        st.write("Dados em formato texto:")
                        for idx, row in display_df.iterrows():
                            st.text(f"Linha {idx}: {dict(row)}")
        else:
            st.write(msg)
            try:
                st.dataframe(pd.DataFrame({"Informação": ["Sem linhas para exibir"]}))
            except Exception:
                st.write("Sem linhas para exibir")
        
    except Exception as e:
        st.error(f"Erro ao exibir DataFrame: {str(e)}")
        st.write(f"{msg}: Não foi possível exibir os dados devido a um erro")
        # Exibir uma tabela informativa sobre o erro - usando o método mais seguro
        try:
            info_df = pd.DataFrame({"Informação sobre o erro": [f"Erro: {str(e)}"]})
            st.table(info_df)  # Usando table em vez de dataframe
        except Exception:
            # Se até isso falhar, mostre como texto simples
            st.write(f"Erro: {str(e)}")
            st.exception(e)  # Mostrar o traceback completo para debug

def force_to_strings(item):
    """
    Converte dados para strings de forma segura, lidando com diversos tipos de dados.
    Pode processar listas, DataFrames ou valores individuais.
    Melhorado para tratar diversos casos de erro comuns.
    """
    # Se for None, retorna uma string vazia
    if item is None:
        return ""
        
    # Se for um DataFrame
    if isinstance(item, pd.DataFrame):
        try:
            if len(item) == 0:  # Se estiver vazio
                return pd.DataFrame({'Nenhum dado': ['Nenhum dado encontrado']})
            
            # Cria uma cópia para não modificar o original
            result_df = item.copy()
            
            # Converte todas as colunas para string com tratamento de erros por coluna
            for col in result_df.columns:
                try:
                    # Tratamento mais detalhado de valores nulos e tipos problemáticos
                    result_df[col] = result_df[col].apply(
                        lambda x: "" if pd.isna(x) else 
                                  str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) and float(x).is_integer() else
                                  str(x).strip()
                    )
                except Exception as e:
                    # Se falhar para uma coluna, tenta uma abordagem mais simples
                    try:
                        result_df[col] = result_df[col].fillna("").astype(str)
                    except:
                        # Se ainda falhar, pelo menos não quebra o processo
                        result_df[col] = ["[Erro de Conversão]"] * len(result_df)
            
            return result_df
        except Exception as e:
            # Fallback para caso de erro no processamento do DataFrame
            return pd.DataFrame({'Erro': [f'Erro ao processar DataFrame: {str(e)}']})
    
    # Se for uma lista, tupla ou conjunto
    if isinstance(item, (list, tuple, set)):
        result = []
        for subitem in item:
            try:
                # Tratar valores nulos primeiro
                if subitem is None or pd.isna(subitem):
                    continue
                
                # Converter para texto
                if isinstance(subitem, (int, float)) and not pd.isna(subitem):
                    # Preservar números inteiros sem decimal
                    if float(subitem).is_integer():
                        text = str(int(subitem))
                    else:
                        text = str(subitem)
                else:
                    text = str(subitem).strip()
                
                # Filtrar valores indesejados
                if (text and
                    not "method" in text.lower() and 
                    not "descriptor" in text.lower() and
                    not "function" in text.lower() and
                    not text.startswith("<") and
                    not text.startswith("[") and
                    not text.startswith("(") and
                    text.lower() not in ['nan', 'none', '', 'null'] and
                    text.strip() != ""):
                    result.append(text)
            except Exception:
                # Ignorar silenciosamente itens problemáticos
                pass
        return result
    
    # Para valores individuais (incluindo inteiros, floats, etc.)
    try:
        if pd.isna(item):  # Verificar se é NaN
            return ""
        elif isinstance(item, (int, float)) and float(item).is_integer():
            return str(int(item))  # Retornar inteiros sem decimal
        else:
            value = str(item).strip()
            # Filtrar strings indesejadas
            if value.lower() in ['nan', 'none', 'null', 'undefined']:
                return ""
            return value
    except Exception:
        return ""


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

            # Verificar se os DataFrames foram criados corretamente
            if protocol_df is None or records_df is None:
                st.error("Erro: Um ou mais arquivos não puderam ser carregados corretamente.")
                if protocol_df is None:
                    st.error("O arquivo de protocolo não foi carregado.")
                if records_df is None:
                    st.error("O arquivo de registros não foi carregado.")
                st.stop()
                return
                
            # Normalize columns
            try:
                for df in [protocol_df, records_df]:
                    df.columns = df.columns.str.replace(r'\s+', ' ', regex=True).str.strip()
            except Exception as e:
                st.error(f"Erro ao normalizar colunas: {str(e)}")
                st.stop()
                return

            # Use Test Case and Traceability (PRS) columns
            for df_name, df in [("Protocol", protocol_df), ("Records", records_df)]:
                if 'Test Case' not in df.columns or 'Traceability (PRS)' not in df.columns:
                    st.error(f"Required columns 'Test Case' and 'Traceability (PRS)' not found in {df_name} file.")
                    st.write(f"Colunas disponíveis em {df_name}: {df.columns.tolist()}")
                    st.stop()
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
            defect_df[status_col] = defect_df[status_col].astype(str).str.strip().str.lower()
            open_defects = defect_df[defect_df[status_col] == 'open'][defect_id_col].nunique()

            # Clean records_df for both columns using boolean masks
            def mask_test_case_column(series):
                return series.str.strip().str.startswith('TASY_', na=False)
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
                
                # Processamento dos registros sem exibir debug
                
                # Filtro menos restritivo
                valid_test_cases = records_all['Test Case'].str.strip().apply(
                    lambda x: len(x) > 3 and 'test case' not in x.lower() and 'note:' not in x.lower()
                )
                valid_prs = records_all['Traceability (PRS)'].str.strip().apply(
                    lambda x: len(x) > 1 and 'traceability' not in x.lower() and 'note:' not in x.lower()
                )
                
                records_all_filtered = records_all[valid_test_cases & valid_prs]
                protocols_executed = records_all_filtered['Test Case'].nunique()
                prs_executed = records_all_filtered['Traceability (PRS)'].nunique()
                
                # Calcula valores únicos sem debug
                protocols_executed = records_all_filtered['Test Case'].nunique()
                prs_executed = records_all_filtered['Traceability (PRS)'].nunique()

                # 2. Overall Not Executed (in protocol but not in records)
                # Aplicar os mesmos filtros de validação para protocolos e registros
                valid_protocol_tc = protocol_df['Test Case'].str.strip().apply(
                    lambda x: len(x) > 3 and 'test case' not in x.lower() and 'note:' not in x.lower()
                )
                valid_protocol_prs = protocol_df['Traceability (PRS)'].str.strip().apply(
                    lambda x: len(x) > 1 and 'traceability' not in x.lower() and 'note:' not in x.lower()
                )
                
                # Tratar valores NaN explicitamente sem exibir debug
                protocol_df['Test Case'] = protocol_df['Test Case'].fillna("").astype(str)
                protocol_df['Traceability (PRS)'] = protocol_df['Traceability (PRS)'].fillna("").astype(str)
                
                # Filtrar protocolos válidos com critérios ainda mais rigorosos
                try:
                    # Criar filtros separadamente para evitar problemas
                    not_na_test_case = ~protocol_df['Test Case'].isna()
                    not_empty_test_case = ~protocol_df['Test Case'].str.lower().isin(['nan', 'none', ''])
                    
                    # Verificar se o valor começa com TASY_ de forma mais segura
                    protocol_df['TC_Check'] = protocol_df['Test Case'].astype(str).apply(
                        lambda x: str(x).strip().startswith('TASY_') if isinstance(x, str) and x and x.lower() != 'nan' else False
                    )
                    tasy_vtc_filter = protocol_df['TC_Check']
                    
                    # Filtrar PRS
                    not_na_prs = ~protocol_df['Traceability (PRS)'].isna()
                    not_empty_prs = ~protocol_df['Traceability (PRS)'].str.lower().isin(['nan', 'none', ''])
                    
                    # Verificar se o valor começa com A de forma mais segura
                    protocol_df['PRS_Check'] = protocol_df['Traceability (PRS)'].astype(str).apply(
                        lambda x: str(x).strip().startswith('A') if isinstance(x, str) and x and x.lower() != 'nan' else False
                    )
                    a_prs_filter = protocol_df['PRS_Check']
                    
                    # Combinar os filtros
                    filter_a = not_na_test_case & not_empty_test_case & tasy_vtc_filter
                    filter_b = not_na_prs & not_empty_prs & a_prs_filter
                    filter_c = valid_protocol_tc & valid_protocol_prs
                    
                    final_filter = filter_a & filter_b & filter_c
                    
                    protocol_filtered = protocol_df[final_filter].copy()
                    
                except Exception as e:
                    # Filtro simplificado como fallback sem debug
                    test_case_filter = protocol_df['Test Case'].astype(str).str.contains('TASY_')
                    prs_filter = protocol_df['Traceability (PRS)'].astype(str).str.contains('^A', regex=True)
                    protocol_filtered = protocol_df[test_case_filter & prs_filter]
                
                # Verificar se o DataFrame filtrado tem alguma linha
                if len(protocol_filtered) == 0:
                    st.error("Nenhum protocolo válido encontrado após aplicar os filtros.")
                    
                    try:
                        # Aplicar filtros progressivos sem mensagens de debug
                        protocol_filtered = protocol_df.copy()
                        
                        # Passo 1: Filtrar Test Case por comprimento
                        protocol_filtered = protocol_filtered[
                            protocol_filtered['Test Case'].astype(str).str.len() > 3
                        ]
                        
                        # Passo 2: Excluir valores inválidos em Test Case
                        protocol_filtered = protocol_filtered[
                            ~protocol_filtered['Test Case'].astype(str).str.lower().isin(['nan', 'none', ''])
                        ]
                        
                        # Passo 3: Filtrar PRS por comprimento
                        protocol_filtered = protocol_filtered[
                            protocol_filtered['Traceability (PRS)'].astype(str).str.len() > 1
                        ]
                        
                        # Passo 4: Excluir valores inválidos em PRS
                        protocol_filtered = protocol_filtered[
                            ~protocol_filtered['Traceability (PRS)'].astype(str).str.lower().isin(['nan', 'none', ''])
                        ]
                        
                    except Exception as e:
                        # Fallbacks sem mensagens de debug
                        try:
                            protocol_filtered = protocol_df.copy()
                            mask = protocol_filtered['Test Case'].astype(str).str.contains('TASY_', na=False)
                            protocol_filtered = protocol_filtered[mask]
                        except Exception:
                            try:
                                protocol_filtered = protocol_df.head(100)
                            except Exception:
                                protocol_filtered = pd.DataFrame(columns=protocol_df.columns)
                
                # Obter casos de teste válidos em protocolo e registros (não vazios) de forma mais robusta
                try:
                    # Função para validar e limpar os casos de teste
                    def clean_test_cases(test_cases):
                        cleaned = []
                        for tc in test_cases:
                            if tc is None:
                                continue
                            if isinstance(tc, str):
                                tc_str = tc.strip()
                                if tc_str and tc_str.lower() != 'nan' and tc_str.startswith('TASY_'):
                                    cleaned.append(tc_str)
                            elif isinstance(tc, (int, float)) and not pd.isna(tc):
                                cleaned.append(str(int(tc) if tc.is_integer() else tc))
                        return set(cleaned)
                    
                    # Aplicar a função de limpeza aos dados
                    test_cases_in_protocol = clean_test_cases(protocol_filtered['Test Case'].tolist())
                    test_cases_in_records = clean_test_cases(records_all_filtered['Test Case'].tolist())
                    
                    # Calcular os que estão no protocolo mas não nos registros (não executados) sem debug
                    protocols_not_executed = test_cases_in_protocol - test_cases_in_records
                except Exception as e:
                    st.error(f"Erro ao processar casos de teste: {str(e)}")
                    st.exception(e)
                    protocols_not_executed = set()
                
                # Verificar se há protocolos não executados
                if len(protocols_not_executed) == 0:
                    st.warning("Não foram encontrados protocolos não executados.")
                
                # Obter os PRS únicos relacionados aos casos de teste não executados
                # Processar os PRS relacionados aos casos de teste não executados
                try:
                    if 'Traceability (PRS)' in protocol_df.columns and len(protocols_not_executed) > 0:
                        # Função para validar e limpar os PRS com critérios mais flexíveis
                        def clean_prs(prs_list):
                            cleaned = []
                            for prs in prs_list:
                                if prs is None:
                                    continue
                                if isinstance(prs, str):
                                    prs_str = prs.strip()
                                    if prs_str and prs_str.lower() != 'nan' and len(prs_str) > 1:
                                        # Verificar formatos comuns de PRS com critérios mais flexíveis
                                        if (prs_str.startswith('A') or      # Formato padrão
                                            prs_str.startswith('a') or      # Letras minúsculas
                                            prs_str.replace(' ', '').startswith('A')):  # Espaços antes do A
                                            cleaned.append(prs_str)
                                        # Se contém múltiplos PRSs separados por vírgula ou ponto-e-vírgula
                                        elif ',' in prs_str or ';' in prs_str:
                                            parts = re.split(r'[,;\s]+', prs_str)
                                            for part in parts:
                                                part = part.strip()
                                                if part and part.lower() != 'nan' and len(part) > 1:
                                                    if (part.startswith('A') or part.startswith('a')):
                                                        cleaned.append(part)
                                elif isinstance(prs, (int, float)) and not pd.isna(prs):
                                    # Converter números para string
                                    prs_str = str(int(prs) if float(prs).is_integer() else prs)
                                    if prs_str.startswith('A') or prs_str.startswith('a'):
                                        cleaned.append(prs_str)
                            return set(cleaned)
                        
                        # Usar o mesmo filtro rigoroso para os PRS
                        try:
                            # Primeiro verificamos se os casos de teste estão no formato adequado
                            if len(protocols_not_executed) == 0:
                                st.warning("Lista de protocolos não executados está vazia")
                                prs_not_executed = 0
                            else:
                                # Verificar se Test Case existe em protocol_filtered
                                if 'Test Case' not in protocol_filtered.columns:
                                    st.error("Coluna 'Test Case' não encontrada em protocol_filtered")
                                    prs_not_executed = 0
                                else:
                                    # Filtragem mais segura com tratamento de erros
                                    safe_protocols = [tc for tc in protocols_not_executed if tc and isinstance(tc, str)]
                                    
                                    if not safe_protocols:
                                        st.warning("Não há casos de teste válidos para filtrar")
                                        prs_related_df = pd.DataFrame()
                                    else:
                                        mask = protocol_filtered['Test Case'].isin(safe_protocols)
                                        prs_related_df = protocol_filtered.loc[mask]                                        # Verificar se temos linhas antes de continuar
                                    if len(prs_related_df) > 0:
                                        # Primeiro filtrar apenas os valores válidos de Traceability (PRS)
                                        valid_prs_values = prs_related_df['Traceability (PRS)'].dropna()
                                        
                                        # Aplicar filtros mais rigorosos para remover PRSs inválidos
                                        valid_prs_values = valid_prs_values[
                                            # Verificar se começa com 'A'
                                            valid_prs_values.str.strip().str.startswith('A') &
                                            # Remover strings vazias ou muito curtas
                                            (valid_prs_values.str.strip().str.len() > 1) &
                                            # Remover valores que não deveriam estar na lista
                                            ~valid_prs_values.str.lower().isin(['nan', 'none', '', 'traceability (prs)'])
                                        ]
                                        
                                        # Agora limpar os valores e obter valores únicos usando a função clean_prs
                                        all_prs = clean_prs(valid_prs_values.unique().tolist())
                                        
                                        # Contar apenas PRSs únicos e válidos
                                        prs_not_executed = len(all_prs)
                                    else:
                                        prs_not_executed = 0
                        except Exception as e:
                            st.error(f"Erro ao filtrar PRS relacionados: {str(e)}")
                            st.exception(e)
                            prs_not_executed = 0
                    else:
                        st.warning("Coluna 'Traceability (PRS)' não encontrada ou não há protocolos não executados")
                        prs_not_executed = 0
                except Exception as e:
                    st.error(f"Erro ao processar PRS não executados: {str(e)}")
                    st.exception(e)
                    prs_not_executed = 0
                
                # Recalcular prs_not_executed para garantir que está contando corretamente sem debug
                try:
                    # Garantir que temos protocolos não executados válidos
                    if protocols_not_executed and len(protocols_not_executed) > 0:
                        # Filtrar o DataFrame de protocolo para obter todos os registros dos protocolos não executados
                        master_prs_df = protocol_df.loc[
                            protocol_df['Test Case'].isin(protocols_not_executed)
                        ].copy()
                        
                        # Filtrar apenas PRSs válidos
                        if len(master_prs_df) > 0:
                            # Extrair PRSs dos protocolos não executados
                            all_prs = master_prs_df['Traceability (PRS)'].dropna().tolist()
                            
                            # Limpar e validar os PRSs com critérios flexíveis
                            valid_prs_values = []
                            
                            for prs in all_prs:
                                if isinstance(prs, str):
                                    prs = prs.strip()
                                    if (len(prs) > 1 and prs.lower() not in ['nan', 'none', '', 'traceability (prs)']):
                                        if (prs.startswith('A') or prs.startswith('a') or prs.replace(' ', '').startswith('A')):
                                            valid_prs_values.append(prs)
                                elif isinstance(prs, (int, float)) and not pd.isna(prs):
                                    valid_prs_values.append(str(int(prs) if float(prs).is_integer() else prs))
                            
                            # Remover duplicatas
                            unique_prs = set(valid_prs_values)
                            
                            # Atualizar prs_not_executed com o valor correto
                            prs_not_executed = len(unique_prs)
                            
                            # Se ainda não encontramos PRSs válidos, tentar uma abordagem diferente
                            if prs_not_executed == 0:
                                # Método alternativo: extrair qualquer texto que pareça um PRS
                                alt_prs_values = []
                                for prs in all_prs:
                                    if isinstance(prs, str):
                                        parts = re.split(r'[,;\s]+', prs)
                                        for part in parts:
                                            part = part.strip()
                                            if len(part) >= 2 and part.lower() not in ['nan', 'none', '', 'traceability']:
                                                alt_prs_values.append(part)
                                
                                # Remover duplicatas
                                alt_unique_prs = set(alt_prs_values)
                                
                                # Se encontramos PRSs com o método alternativo, usar esse valor
                                if len(alt_unique_prs) > 0:
                                    prs_not_executed = len(alt_unique_prs)
                except Exception:
                    pass
                
                # Verificar se os protocolos não executados estão realmente iniciando com TASY_
                if protocols_not_executed:
                    # Garantir que todos são strings e não NaN
                    protocols_not_executed_list = list(protocols_not_executed)
                    valid_protocols = []
                    
                    for tc in protocols_not_executed_list:
                        if isinstance(tc, str) and tc and tc.lower() != 'nan' and tc.startswith('TASY_'):
                            valid_protocols.append(tc)
                    
                    # Atualizar a lista de protocolos não executados se necessário
                    if len(valid_protocols) != len(protocols_not_executed_list):
                        protocols_not_executed = set(valid_protocols)

                # 3. Overall Passed Results (Result == Pass in records)
                if 'Conclusion (Pass / Fail)' in records_df.columns:
                    records_df['Conclusion (Pass / Fail)'] = records_df['Conclusion (Pass / Fail)'].astype(str).str.strip()
                    
                    # Aplicar o mesmo filtro usado no cálculo de protocols_executed
                    valid_tc = records_df['Test Case'].str.strip().apply(
                        lambda x: len(x) > 3 and 'test case' not in x.lower() and 'note:' not in x.lower()
                    )
                    valid_prs = records_df['Traceability (PRS)'].str.strip().apply(
                        lambda x: len(x) > 1 and 'traceability' not in x.lower() and 'note:' not in x.lower()
                    )
                    
                    passed_mask = (
                        records_df['Conclusion (Pass / Fail)'].str.lower().isin(['pass', 'passed']) &
                        valid_tc & valid_prs
                    )
                    passed_df = records_df[passed_mask]
                    protocols_passed = passed_df['Test Case'].nunique()
                    prs_passed = passed_df['Traceability (PRS)'].nunique()
                else:
                    protocols_passed = prs_passed = 0

                # 4. Overall Failed Results (not all passed, no round logic)
                if 'Conclusion (Pass / Fail)' in records_df.columns:
                    # Aplicar o mesmo filtro usado no cálculo de protocols_executed
                    valid_tc = records_df['Test Case'].str.strip().apply(
                        lambda x: len(x) > 3 and 'test case' not in x.lower() and 'note:' not in x.lower()
                    )
                    valid_prs = records_df['Traceability (PRS)'].str.strip().apply(
                        lambda x: len(x) > 1 and 'traceability' not in x.lower() and 'note:' not in x.lower()
                    )
                    
                    valid_mask = valid_tc & valid_prs
                    df_valid = records_df[valid_mask].copy()
                    
                    # Verificar se há entradas suficientes para fazer o groupby
                    if len(df_valid) > 0:
                        grouped = df_valid.groupby(['Traceability (PRS)', 'Test Case'])['Conclusion (Pass / Fail)'].apply(
                            lambda x: all(r.lower() == 'passed' for r in x)
                        )
                        failed_pairs = grouped[~grouped].reset_index()[['Traceability (PRS)', 'Test Case']]
                        protocols_failed = failed_pairs['Test Case'].nunique()
                        prs_failed = failed_pairs['Traceability (PRS)'].nunique()
                    else:
                        protocols_failed = prs_failed = 0
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
                try:
                    # Processar cada linha para garantir que todos os valores são strings
                    processed_data = []
                    for row in parameter_data:
                        processed_row = []
                        for item in row:
                            if isinstance(item, (list, tuple, set)):
                                processed_row.append(", ".join(force_to_strings(item)))
                            else:
                                processed_row.append(force_to_strings(item))
                        processed_data.append(processed_row)
                    
                    parameter_df = pd.DataFrame(processed_data, columns=["#", "Parameter", "Value"])
                    st.markdown("#### Parameter Table")
                    
                    # Exibir o DataFrame apenas uma vez usando o método padrão do Streamlit
                    try:
                        st.dataframe(parameter_df, use_container_width=True)
                    except Exception:
                        # Usar a função segura apenas como fallback
                        safe_display_dataframe(parameter_df, msg="Tabela de Parâmetros", max_rows=10)
                except Exception as e:
                    st.error(f"Erro ao criar tabela de parâmetros: {str(e)}")
                    st.write("Dados dos parâmetros:")
                    for row in parameter_data:
                        safe_data = force_to_strings(row)
                        st.write(safe_data)

                try:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        # Garantir que parameter_df existe e tem dados
                        try:
                            if 'parameter_df' in locals() and parameter_df is not None and len(parameter_df) > 0:
                                # Certificar que todas as colunas são strings
                                parameter_df_safe = parameter_df.copy()
                                for col in parameter_df_safe.columns:
                                    parameter_df_safe[col] = parameter_df_safe[col].astype(str)
                                parameter_df_safe.to_excel(writer, sheet_name='Parameter', index=False)
                                st.write("Tabela de parâmetros incluída no arquivo Excel.")
                            else:
                                # Criar um DataFrame de fallback se o original não estiver disponível
                                st.warning("Usando dados de fallback para o arquivo Excel")
                                fallback_data = [
                                    ["1", "Overall Test Case executed", str(protocols_executed) + " (related to " + str(prs_executed) + " PRS)"],
                                    ["2", "Overall Not Executed", str(len(protocols_not_executed)) + " (related to " + str(prs_not_executed) + " PRS)"],
                                    ["3", "Overall Passed Results", str(protocols_passed) + " (related to " + str(prs_passed) + " PRS)"],
                                    ["4", "Overall Failed Results", str(protocols_failed) + " (related to " + str(prs_failed) + " PRS)"],
                                    ["5", "Overall Number of Open Anomalies classified as Defects", str(open_defects)]
                                ]
                                fallback_df = pd.DataFrame(fallback_data, columns=["#", "Parameter", "Value"])
                                fallback_df.to_excel(writer, sheet_name='Parameter', index=False)
                                st.write("Dados de fallback incluídos no arquivo Excel.")
                        except Exception as e_excel:
                            st.error(f"Erro ao escrever dados na planilha: {str(e_excel)}")
                            # Último recurso: criar um DataFrame extremamente simples
                            pd.DataFrame({"Erro": ["Ocorreu um erro ao gerar o relatório."]}).to_excel(writer, sheet_name='Error', index=False)
                            
                    buffer.seek(0)
                    st.success("Product Validation Report generated! Download your results below.")
                    try:
                        st.download_button(
                            "⬇️ Download Parameter Table",
                            data=buffer,
                            file_name="product_verification_report_parameter.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as download_err:
                        st.error(f"Erro ao gerar botão de download: {str(download_err)}")
                        st.warning("Não foi possível criar o botão de download. Tente novamente após corrigir os erros acima.")
                except Exception as e:
                    st.error(f"Erro ao gerar arquivo para download: {str(e)}")
        else:
            st.warning("Please click on the 'Run Product Verification Report' button to generate the report.")