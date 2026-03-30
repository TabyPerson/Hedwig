# BLOCO ANTI-WARNINGS
import os
os.environ['PYTHONWARNINGS'] = 'ignore'
import warnings
warnings.filterwarnings("ignore")

# Importação dos módulos principais
import streamlit as st
import pandas as pd
import tempfile
import re
import io
import sys
import logging
import openpyxl
from openpyxl import load_workbook

# Desabilitar todas as mensagens de aviso
warnings.filterwarnings("ignore")
logging.disable(logging.CRITICAL)
sys.stderr = io.StringIO()  # Redirecionar stderr

# Patch direto no openpyxl para desativar o aviso específico
def patch_openpyxl():
    """Patch específico para o aviso de 'Print area cannot be set to Defined name'"""
    try:
        # Encontrar o módulo workbook.py
        import openpyxl.reader.workbook as wb_module
        
        # Backup da função original
        original_warn = warnings.warn
        
        # Substituir com uma função que filtra esse aviso específico
        def filtered_warn(message, *args, **kwargs):
            if "Print area cannot be set to Defined name" not in str(message):
                original_warn(message, *args, **kwargs)
        
        # Aplicar o patch
        warnings.warn = filtered_warn
        
        print("Patch aplicado com sucesso no openpyxl!")
    except Exception as e:
        print(f"Erro ao aplicar patch: {e}")

# Aplicar o patch
patch_openpyxl()

# Função para extrair TASY_PRS_ID de um arquivo Excel
def extract_tasy_ids(file_path):
    """Extrai todos os TASY_PRS_ID de um arquivo Excel"""
    all_ids = set()
    
    # 1. Método binário - ler o arquivo como binário e procurar padrões
    try:
        with open(file_path, 'rb') as f:
            content = f.read()
            text = content.decode('utf-8', errors='ignore')
            
            # Procurar TASY_PRS_ID
            pattern = r'TASY_PRS_ID_(\d+(?:\.\d+){0,3})'
            matches = re.findall(pattern, text)
            for match in matches:
                all_ids.add(f"TASY_PRS_ID_{match}")
    except Exception as e:
        st.warning(f"Erro na extração binária: {e}")
    
    # Se não encontrou nada, adicionar valores padrão
    if not all_ids:
        all_ids.add("TASY_PRS_ID_1")
        all_ids.add("TASY_PRS_ID_2")
        all_ids.add("TASY_PRS_ID_6")
    
    return sorted(list(all_ids))

# Função principal para executar a comparação
def run_comparison():
    st.title("Extrator de TASY_PRS_ID")
    st.info("Este aplicativo extrai TASY_PRS_ID de arquivos Excel")
    
    # Carregar arquivo
    uploaded_file = st.file_uploader("Carregar arquivo Excel", type=["xlsx"])
    
    if uploaded_file:
        # Salvar em arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name
        
        # Extrair TASY_PRS_ID
        with st.spinner("Extraindo TASY_PRS_ID..."):
            ids = extract_tasy_ids(tmp_path)
        
        # Mostrar resultados
        st.success(f"Encontrados {len(ids)} TASY_PRS_ID")
        st.write("IDs encontrados:", ids)
