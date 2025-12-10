import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
import string

# =========================
# Configurações da página
# =========================
st.set_page_config(
    page_title="Extrair Fatura para Excel",
    layout="wide"
)
st.title("Extrair Débitos e Gerar Excel")

uploaded_file = st.file_uploader(
    "Escolha o PDF da fatura",
    type="pdf"
)

# =========================
# Funções utilitárias
# =========================
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"


def extract_text_from_pdf(file):
    texts = []

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                texts.append(page_text)

    if not texts:
        st.info("PDF sem texto detectável. Usando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img))

    return texts


# =========================
# Tabela 1 – Transações
# =========================
def extract_table_transacoes(text):
    """
    Data | Estabelecimento | Valor
    """
    pattern = (
        r"(\d{2}/\d{2})\s+"
        r"[\d.]+\s+"
        r"(.+?)\s+"
        r"([\d.,]+)$"
    )

    matches = re.findall(pattern, text, re.MULTILINE)

    if matches:
        return pd.DataFrame(
            matches,
            columns=["Data", "Estabelecimento", "Valor (R$)"]
        )

    return pd.DataFrame()


# =========================
# Tabela 2 – Favorecidos
# =========================
def extract_table_favorecidos(text):
    """
    Data | Canal | Tipo | Favorecido | ISPB | Agência | Conta | Valor
    """

    pattern = (
        r"(\d{2}/\d{2})\s+"            # Data
        r"([A-Z]+)\s+"                 # Canal
        r"([A-Z\s]+?)\s+"              # Tipo
        r"([A-Z\s]+?)\s+"              # Favorecido
        r"(\d{8})\s+"                  # ISPB
        r"(\d{3,5})\s+"                # Agência
        r"([\d\-]+)\s+"                # Conta
        r"([\d.,]+)$"                  # Valor
    )

    matches = re.findall(pattern, text, re.MULTILINE)

    if matches:
        return pd.DataFrame(
            matches,
            columns=[
                "Data",
                "Favorecido",
                "Valor (R$)"
            ]
        )

    return pd.DataFrame()


# =========================
# Execução principal
# =========================
if uploaded_file:
    try:
        uploaded_file.seek(0)
        texts = extract_text_from_pdf(uploaded_file)

        transacoes = []
        favorecidos = []

        for t in texts:
            df_t = extract_table_transacoes(t)
            df_f = extract_table_favorecidos(t)

            if not df_t.empty:
                transacoes.append(df_t)

            if not df_f.empty:
                favorecidos.append(df_f)

        if not transacoes and not favorecidos:
            st.warning("Nenhuma tabela reconhecida no PDF.")
        else:
            nome_arquivo = st.text_input(
                "Nome do arquivo Excel (sem .xlsx)",
                "fatura_extraida"
            )

            if st.button("Gerar Excel"):
                output = BytesIO()

                with pd.ExcelWriter(output, engine="openpyxl") as writer:

                    if transacoes:
                        df_transacoes = pd.concat(
                            transacoes,
                            ignore_index=True
                        )
                        df_transacoes.to_excel(
                            writer,
                            sheet_name="Transações",
                            index=False
                        )

                    if favorecidos:
                        df_favorecidos = pd.concat(
                            favorecidos,
                            ignore_index=True
                        )
                        df_favorecidos.to_excel(
                            writer,
                            sheet_name="Favorecidos",
                            index=False
                        )

                output.seek(0)

                st.download_button(
                    "Baixar Excel",
                    data=output,
                    file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if transacoes:
                st.subheader("Transações")
                st.dataframe(pd.concat(transacoes, ignore_index=True))

            if favorecidos:
                st.subheader("Favorecidos")
                st.dataframe(pd.concat(favorecidos, ignore_index=True))

    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")

