import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from copy import copy
import tempfile
import os
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Gerador de Arquivos", layout="wide")

st.title("📁 Gerador de Arquivos Filtrados")

arquivo = st.file_uploader(
    "Selecione a planilha Excel",
    type=["xlsx"]
)

if arquivo:

    xls = pd.ExcelFile(arquivo)

    abas = xls.sheet_names

    aba_escolhida = st.selectbox(
        "Selecione a aba",
        abas,
        index=abas.index("BASE_NÚCLEO") if "BASE_NÚCLEO" in abas else 0
    )

    df = pd.read_excel(arquivo, sheet_name=aba_escolhida)

    coluna_filtro = st.selectbox(
        "Selecione a coluna para filtro",
        df.columns
    )

    valores = sorted(df[coluna_filtro].dropna().astype(str).unique())

    modo = st.radio(
        "Modo de geração",
        [
            "Gerar todos",
            "Gerar apenas selecionados"
        ]
    )

    selecionados = []

    if modo == "Gerar apenas selecionados":
        selecionados = st.multiselect(
            "Selecione os valores",
            valores
        )
    else:
        selecionados = valores

    nome_padrao = st.text_input(
        "Nome padrão dos arquivos",
        value="BASE_NÚCLEO"
    )

    if st.button("🚀 Gerar Arquivos"):

        pasta_downloads = str(Path.home() / "Downloads")

        progresso = st.progress(0)

        total = len(selecionados)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(arquivo.getbuffer())
        st.info(f"Arquivos salvos em: {pasta_downloads}")