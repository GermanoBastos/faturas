import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract

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

if uploaded_file:
    try:
        uploaded_file.seek(0)
        texts = extract_text_from_pdf(uploaded_file)
        all_tables = [extract_table_from_text(t) for t in texts if not extract_table_from_text(t).empty]

        if all_tables:
            # Concatenar todas as tabelas extraídas
            final_df = pd.concat(all_tables, ignore_index=True)

            # Gerar Excel imediatamente
            output = BytesIO()
            final_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            # Botão de download exibido logo no topo
            st.download_button(
                label="Clique para baixar o Excel",
                data=output,
                file_name="fatura_extraida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Exibir tabela no app
            st.subheader("Tabela extraída:")
            st.dataframe(final_df)

        else:
            st.warning("Nenhuma tabela de transações encontrada no PDF.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")

