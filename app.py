import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
import string

# ===============================
# CONFIGURAÇÃO DA PÁGINA
# ===============================
st.set_page_config(page_title="Extrair Fatura para Excel", layout="wide")
st.title("Extrair Débitos da Fatura")

# ===============================
# UPLOAD DO PDF
# ===============================
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ===============================
# FUNÇÕES
# ===============================
def extract_text_from_pdf(file):
    texts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                texts.append(txt)

    # OCR se não houver texto
    if not texts:
        st.info("PDF sem texto detectável. Usando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img, lang="por"))

    return texts


def valor_br_para_float(valor):
    """Converte 1.234,56 → 1234.56 com 2 casas"""
    valor = valor.replace(".", "").replace(",", ".")
    return round(float(valor), 2)


# ===============================
# TABELA 1 – TRANSAÇÕES
# ===============================
def extract_tabela_transacoes(text):
    """
    Data | Estabelecimento | Valor
    """
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df = pd.DataFrame(matches, columns=["Data", "Descrição", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df


# ===============================
# TABELA 2 – FAVORECIDOS
# ===============================
def extract_tabela_favorecidos(text):
    """
    Data | Canal | Tipo | Favorecido | ISPB | Agência | Conta | Valor
    (mas só vamos manter Data, Favorecido, Valor)
    """
    pattern = (
        r"(\d{2}/\d{2})\s+"           # Data
        r"\S+\s+"                     # Canal
        r"\S+\s+"                     # Tipo
        r"(.+?)\s+"                   # Favorecido
        r"\d+\s+"                     # ISPB
        r"\d+\s+"                     # Agência
        r"\d+\s+"                     # Conta
        r"([\d.,]+)"                  # Valor
    )

    matches = re.findall(pattern, text)

    if not matches:
        return pd.DataFrame()

    df = pd.DataFrame(matches, columns=["Data", "Favorecido", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df


def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"


# ===============================
# PROCESSAMENTO
# ===============================
if uploaded_file:
    uploaded_file.seek(0)
    texts = extract_text_from_pdf(uploaded_file)

    tabelas_transacoes = []
    tabelas_favorecidos = []

    for t in texts:
        df1 = extract_tabela_transacoes(t)
        if not df1.empty:
            tabelas_transacoes.append(df1)

        df2 = extract_tabela_favorecidos(t)
        if not df2.empty:
            tabelas_favorecidos.append(df2)

    if tabelas_transacoes or tabelas_favorecidos:
        file_name_input = st.text_input(
            "Nome do arquivo Excel (sem extensão)",
            "fatura_extraida"
        )

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if tabelas_transacoes:
                df_transacoes = pd.concat(tabelas_transacoes, ignore_index=True)
                df_transacoes.to_excel(
                    writer,
                    sheet_name="Transacoes",
                    index=False
                )

            if tabelas_favorecidos:
                df_fav = pd.concat(tabelas_favorecidos, ignore_index=True)
                df_fav.to_excel(
                    writer,
                    sheet_name="Favorecidos",
                    index=False
                )

        output.seek(0)

        st.download_button(
            "📥 Baixar Excel",
            data=output,
            file_name=sanitize_filename(file_name_input) + ".xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if tabelas_transacoes:
            st.subheader("Transações")
            st.dataframe(df_transacoes)

        if tabelas_favorecidos:
            st.subheader("Favorecidos (Data | Favorecido | Valor)")
            st.dataframe(df_fav)

    else:
        st.warning("Nenhuma tabela encontrada no PDF.")
