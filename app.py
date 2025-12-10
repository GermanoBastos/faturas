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

# =========================
# Configuração da página
# =========================
st.set_page_config(page_title="Extrair Fatura para Excel", layout="wide")
st.title("Extrair Débitos e Gerar Excel")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

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
# (Validação completa, saída reduzida)
# =========================
def extract_table_favorecidos(text):
    pattern = (
        r"(\d{2}/\d{2})\s+"
        r"[A-Z]+\s+"
        r"[A-Z\s]+?\s+"
        r"([A-Z\s]+?)\s+"
        r"\d{8}\s+"
        r"\d{3,5}\s+"
        r"[\d\-]+\s+"
        r"([\d.,]+)$"
    )

    matches = re.findall(pattern, text, re.MULTILINE)

    if matches:
        return pd.DataFrame(
            matches,
            columns=["Data", "Favorecido", "Valor (R$)"]
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

                    style = TableStyleInfo(
                        name="TableStyleMedium9",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False
                    )

                    if transacoes:
                        df_transacoes = pd.concat(transacoes, ignore_index=True)
                        sheet = "Transações"
                        df_transacoes.to_excel(writer, sheet_name=sheet, index=False)

                        ws = writer.book[sheet]
                        max_col = ws.max_column
                        max_row = ws.max_row
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"

                        tabela = Table(
                            displayName="TabelaTransacoes",
                            ref=ref
                        )
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

                    if favorecidos:
                        df_favorecidos = pd.concat(favorecidos, ignore_index=True)
                        sheet = "Favorecidos"
                        df_favorecidos.to_excel(writer, sheet_name=sheet, index=False)

                        ws = writer.book[sheet]
                        max_col = ws.max_column
                        max_row = ws.max_row
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"

                        tabela = Table(
                            displayName="TabelaFavorecidos",
                            ref=ref
                        )
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

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
