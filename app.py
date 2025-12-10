import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
import string

# Configurações da página
st.set_page_config(page_title="Extrair Fatura para Excel", layout="wide")
st.title("Extrair Débitos e Gerar Excel")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")


def extract_text_from_pdf(file):
    texts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                texts.append(page_text)

    if not texts:
        st.info("PDF sem texto detectável. Tentando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img))

    return texts


def extract_table_transacoes(text):
    """
    Tabela 1: Data | Estabelecimento | Valor
    """
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)
    return pd.DataFrame(matches, columns=["Data", "Estabelecimento", "Valor (R$)"])


def extract_table_favorecidos(text):
    """
    Tabela 2: Data | Favorecido | Valor
    Exemplo esperado:
    12/09 PAGAMENTO JOAO SILVA 1.250,00
    """
    pattern = r"(\d{2}/\d{2})\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)
    return pd.DataFrame(matches, columns=["Data", "Favorecido", "Valor (R$)"])


def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"


if uploaded_file:
    try:
        uploaded_file.seek(0)
        texts = extract_text_from_pdf(uploaded_file)

        tabelas_transacoes = []
        tabelas_favorecidos = []

        for t in texts:
            df1 = extract_table_transacoes(t)
            df2 = extract_table_favorecidos(t)

            if not df1.empty:
                tabelas_transacoes.append(df1)

            if not df2.empty:
                tabelas_favorecidos.append(df2)

        if tabelas_transacoes or tabelas_favorecidos:

            nome_arquivo = st.text_input(
                "Digite o nome do arquivo Excel (sem extensão)",
                "fatura_extraida"
            )

            if st.button("Gerar Excel"):
                output = BytesIO()

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    if tabelas_transacoes:
                        df_transacoes = pd.concat(tabelas_transacoes, ignore_index=True)
                        df_transacoes.to_excel(
                            writer,
                            sheet_name="Transações",
                            index=False
                        )

                    if tabelas_favorecidos:
                        df_favorecidos = pd.concat(tabelas_favorecidos, ignore_index=True)
                        df_favorecidos.to_excel(
                            writer,
                            sheet_name="Favorecidos",
                            index=False
                        )

                output.seek(0)

                st.download_button(
                    label="Baixar Excel",
                    data=output,
                    file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if tabelas_transacoes:
                st.subheader("Tabela de Transações")
                st.dataframe(pd.concat(tabelas_transacoes, ignore_index=True))

            if tabelas_favorecidos:
                st.subheader("Tabela de Favorecidos")
                st.dataframe(pd.concat(tabelas_favorecidos, ignore_index=True))

        else:
            st.warning("Nenhuma tabela encontrada no PDF.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
