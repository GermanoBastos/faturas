import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
import string
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import os
import msal
import requests
from datetime import datetime

# ================== Configuração da página ==================
st.set_page_config(page_title="Extrair Fatura para Excel e SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ================== Funções utilitárias ==================
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"

def valor_br_para_float(valor_str):
    if valor_str is None:
        return 0.0
    v = str(valor_str).strip().replace(".", "").replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return 0.0

def extract_text_from_pdf(file):
    texts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                texts.append(txt)
    if not texts:
        st.info("PDF sem texto detectável. Usando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img, lang="por"))
    return texts

def extract_tabela_transacoes(text):
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)
    if not matches:
        return pd.DataFrame()
    df = pd.DataFrame(matches, columns=["Data","Estabelecimento","Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

def extract_tabela_favorecidos(text):
    pattern = (
        r"(\d{2}/\d{2})\s+(\S+)\s+([A-Z0-9\s]+?)\s+"
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+(\d{8})\s+"
        r"(\d{3,5})\s+([\d\-]+)\s+([\d.,]+)"
    )
    matches = re.findall(pattern, text, re.MULTILINE)
    if not matches:
        return pd.DataFrame()
    df_full = pd.DataFrame(matches, columns=[
        "Data","Canal","Tipo","Favorecido","ISPB","Agência","Conta","Valor (raw)"
    ])
    df = pd.DataFrame()
    df["Data"] = df_full["Data"]
    df["Favorecido"] = df_full["Favorecido"].str.strip()
    df["Valor (R$)"] = df_full["Valor (raw)"].apply(valor_br_para_float)
    return df

def extrair_mes_ano(nome_arquivo):
    mes_ano = re.search(r"([A-Z]{3})\s*(\d{4})", nome_arquivo.upper())
    if mes_ano:
        mes_abrev, ano = mes_ano.groups()
        try:
            mes = datetime.strptime(mes_abrev, "%b").month
        except:
            meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
            mes = meses.index(mes_abrev)+1
        return datetime(int(ano), mes, 1)
    return datetime.now()

# ================== Processamento ==================
if uploaded_file:
    uploaded_file.seek(0)
    texts = extract_text_from_pdf(uploaded_file)

    if "df_transacoes" not in st.session_state:
        listas_transacoes, listas_favorecidos = [], []

        for t in texts:
            dt = extract_tabela_transacoes(t)
            if not dt.empty:
                listas_transacoes.append(dt)

            df = extract_tabela_favorecidos(t)
            if not df.empty:
                listas_favorecidos.append(df)

        st.session_state.df_transacoes = (
            pd.concat(listas_transacoes, ignore_index=True)
            if listas_transacoes else pd.DataFrame()
        )

        st.session_state.df_favorecidos = (
            pd.concat(listas_favorecidos, ignore_index=True)
            if listas_favorecidos else pd.DataFrame()
        )

    st.subheader("Pré-visualização (exclusão manual)")

    # ================== Débitos ==================
    if not st.session_state.df_transacoes.empty:
        st.write("Débitos:")

        for i, row in st.session_state.df_transacoes.iterrows():
            c1, c2, c3, c4 = st.columns([1,4,2,0.5])
            c1.write(row["Data"])
            c2.write(row["Estabelecimento"])
            c3.write(f"R$ {row['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"del_t_{i}"):
                st.session_state.df_transacoes.drop(i, inplace=True)
                st.session_state.df_transacoes.reset_index(drop=True, inplace=True)
                st.rerun()

        total = st.session_state.df_transacoes["Valor (R$)"].sum()
        st.info(f"💰 Total de Débitos: R$ {total:,.2f}")

    # ================== PIX ==================
    if not st.session_state.df_favorecidos.empty:
        st.write("Envios de PIX:")

        for i, row in st.session_state.df_favorecidos.iterrows():
            c1, c2, c3, c4 = st.columns([1,4,2,0.5])
            c1.write(row["Data"])
            c2.write(row["Favorecido"])
            c3.write(f"R$ {row['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"del_f_{i}"):
                st.session_state.df_favorecidos.drop(i, inplace=True)
                st.session_state.df_favorecidos.reset_index(drop=True, inplace=True)
                st.rerun()

        total_pix = st.session_state.df_favorecidos["Valor (R$)"].sum()
        st.info(f"💰 Total PIX: R$ {total_pix:,.2f}")
