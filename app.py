import streamlit as st
import pandas as pd
from datetime import datetime
from utils.excel_handler import (
    obter_abas, obter_colunas, obter_valores_unicos, 
    gerar_arquivos_filtrados, verificar_pasta_downloads
)

st.set_page_config(page_title="CFAE - Gerador de Planilhas", layout="wide")

def main():
    st.title("📂 Gerador Automático de Excel Filtrado")
    st.markdown("---")

    uploaded_file = st.file_uploader("Upload da planilha (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

    if uploaded_file:
        try:
            # 1. Seleção da Aba
            abas = obter_abas(uploaded_file)
            aba_padrao_idx = abas.index("Ativos") if "Ativos" in abas else 0
            aba_selecionada = st.selectbox("Selecione a aba:", abas, index=aba_padrao_idx)
            
            # 2. Seleção da Coluna
            colunas = obter_colunas(uploaded_file, aba_selecionada)
            # Ajustado para o nome com acento conforme seu erro anterior
            nome_coluna_alvo = "BASE_NÚCLEO" 
            col_padrao_idx = colunas.index(nome_coluna_alvo) if nome_coluna_alvo in colunas else 0
            coluna_filtro = st.selectbox("Selecione a coluna de filtro:", colunas, index=col_padrao_idx)

            # 3. Filtros de Valores
            valores_unicos = obter_valores_unicos(uploaded_file, aba_selecionada, coluna_filtro)
            opcao = st.radio("O que deseja gerar?", ["Todos os itens", "Selecionar itens específicos"])
            
            itens_processar = valores_unicos
            if opcao == "Selecionar itens específicos":
                itens_processar = st.multiselect("Selecione os valores únicos:", valores_unicos)

            if st.button("🚀 Iniciar Geração"):
                if not itens_processar:
                    st.warning("Selecione pelo menos um item.")
                    return

                progresso = st.progress(0)
                status = st.empty()
                pasta = verificar_pasta_downloads()
                
                sucesso, erros = gerar_arquivos_filtrados(
                    uploaded_file, aba_selecionada, coluna_filtro, 
                    itens_processar, pasta, progresso, status
                )

                if sucesso:
                    st.success(f"✅ Concluído! Arquivos salvos em: {pasta}")
                if erros:
                    st.error(f"❌ Erros em alguns itens: {erros}")

        except Exception as e:
            st.error(f"Erro na interface: {str(e)}")

if __name__ == "__main__":
    main()