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

# Upload do PDF
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

def extract_text_from_pdf(file):
    """Extrai texto do PDF; se não houver, usa OCR"""
    texts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                texts.append(page_text)
    # Se nenhuma página tiver texto, tenta OCR
    if not texts:
        st.info("PDF sem texto detectável. Tentando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            text = pytesseract.image_to_string(img)
            texts.append(text)
    return texts

def extract_table_from_text(text):
    """
    Extrai tabela de transações do texto usando regex.
    Colunas: Data, Estabelecimento, Valor (ignora Número do Cartão)
    """
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)
    if matches:
        df = pd.DataFrame(matches, columns=["Data", "Estabelecimento", "Valor (R$)"])
        return df
    else:
        return pd.DataFrame()

def sanitize_filename(name):
    """Remove caracteres inválidos para nome de arquivo"""
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"

if uploaded_file:
    try:
        uploaded_file.seek(0)
        texts = extract_text_from_pdf(uploaded_file)
        all_tables = [extract_table_from_text(t) for t in texts if not extract_table_from_text(t).empty]

        if all_tables:
            # Concatenar todas as tabelas extraídas
            final_df = pd.concat(all_tables, ignore_index=True)

            # Pedir nome do arquivo para o usuário
            file_name_input = st.text_input("Digite o nome do arquivo Excel (sem extensão)", "fatura_extraida")
            excel_file_name = sanitize_filename(file_name_input) + ".xlsx"

            # Gerar Excel
            output = BytesIO()
            final_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            # Botão de download
            st.download_button(
                label="Clique para baixar o Excel",
                data=output,
                file_name=excel_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Exibir tabela no app
            st.subheader("Tabela extraída:")
            st.dataframe(final_df)

        else:
            st.warning("Nenhuma tabela de transações encontrada no PDF.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
