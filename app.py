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

# Configuração da página
st.set_page_config(page_title="Extrair Fatura para Excel", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais e Excel)")

# Upload do PDF
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# Funções utilitárias
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"

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

def valor_br_para_float(valor_str):
    if valor_str is None:
        return 0.0
    v = str(valor_str).strip().replace(".", "").replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return 0.0

# Extrair Tabela 1 – Transações
def extract_tabela_transacoes(text):
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)
    if not matches:
        return pd.DataFrame()
    df = pd.DataFrame(matches, columns=["Data", "Estabelecimento", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

# Extrair Tabela 2 – Favorecidos
def extract_tabela_favorecidos(text):
    pattern = (
        r"(\d{2}/\d{2})\s+"           
        r"(\S+)\s+"                    
        r"([A-Z0-9\s]+?)\s+"           
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+" 
        r"(\d{8})\s+"                  
        r"(\d{3,5})\s+"                
        r"([\d\-]+)\s+"                
        r"([\d.,]+)"                    
    )
    matches = re.findall(pattern, text, re.MULTILINE)
    if not matches:
        return pd.DataFrame()
    df_full = pd.DataFrame(matches, columns=[
        "Data", "Canal", "Tipo", "Favorecido", "ISPB", "Agência", "Conta", "Valor (raw)"
    ])
    df = pd.DataFrame()
    df["Data"] = df_full["Data"]
    df["Favorecido"] = df_full["Favorecido"].str.strip()
    df["Valor (R$)"] = df_full["Valor (raw)"].apply(valor_br_para_float)
    return df

# Processamento principal
if uploaded_file:
    try:
        uploaded_file.seek(0)
        texts = extract_text_from_pdf(uploaded_file)

        listas_transacoes = []
        listas_favorecidos = []

        for t in texts:
            df_t = extract_tabela_transacoes(t)
            if not df_t.empty:
                listas_transacoes.append(df_t)

            df_f = extract_tabela_favorecidos(t)
            if not df_f.empty:
                listas_favorecidos.append(df_f)

        if not listas_transacoes and not listas_favorecidos:
            st.warning("Nenhuma tabela reconhecida no PDF.")
        else:
            # Pré-visualização
            st.subheader("Débitos e envios de PIX")

            if listas_transacoes:
                df_transacoes = pd.concat(listas_transacoes, ignore_index=True)
                st.write("Débitos:")
                st.dataframe(df_transacoes)
                # Soma no Streamlit
                total_transacoes = df_transacoes["Valor (R$)"].sum()
                st.info(f"💰 Total de Débitos: R$ {total_transacoes:,.2f}")

            if listas_favorecidos:
                df_favorecidos = pd.concat(listas_favorecidos, ignore_index=True)
                st.write("Envios de PIX:")
                st.dataframe(df_favorecidos)
                # Soma no Streamlit
                total_favorecidos = df_favorecidos["Valor (R$)"].sum()
                st.info(f"💰 Total de Envios de PIX: R$ {total_favorecidos:,.2f}")

            # Input para nome do Excel já preenchido com nome do PDF
            default_name = uploaded_file.name.rsplit(".", 1)[0]
            nome_arquivo = st.text_input(
                "Nome do arquivo Excel (sem .xlsx)",
                value=default_name
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

                    # --- Transações ---
                    if listas_transacoes:
                        sheet_name = "Transacoes"
                        df_transacoes.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]
                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"
                        tabela = Table(displayName="TabelaTransacoes", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)
                        # Coluna valor numérica
                        for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                            for cell in row:
                                cell.number_format = '#,##0.00'
                        # Linha TOTAL
                        ws.cell(row=max_row + 1, column=max_col - 1, value="TOTAL")
                        ws.cell(row=max_row + 1, column=max_col, value=f"=SUM({get_column_letter(max_col)}2:{get_column_letter(max_col)}{max_row})")
                        ws.cell(row=max_row + 1, column=max_col).number_format = '#,##0.00'

                    # --- Favorecidos ---
                    if listas_favorecidos:
                        sheet_name = "Favorecidos"
                        df_favorecidos.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]
                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"
                        tabela = Table(displayName="TabelaFavorecidos", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)
                        for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                            for cell in row:
                                cell.number_format = '#,##0.00'
                        ws.cell(row=max_row + 1, column=max_col - 1, value="TOTAL")
                        ws.cell(row=max_row + 1, column=max_col, value=f"=SUM({get_column_letter(max_col)}2:{get_column_letter(max_col)}{max_row})")
                        ws.cell(row=max_row + 1, column=max_col).number_format = '#,##0.00'

                output.seek(0)
                st.success("Excel gerado com sucesso — pronto para download.")
                st.download_button(
                    label="📥 Baixar Excel",
                    data=output,
                    file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")


