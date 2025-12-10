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

# ===============================
# Configuração da página
# ===============================
st.set_page_config(page_title="Extrair Fatura para Excel", layout="wide")
st.title("Extrair Débitos da Fatura (com Tabelas e valores numéricos)")

# ===============================
# Upload do PDF
# ===============================
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ===============================
# Utilitários
# ===============================
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"

def extract_text_from_pdf(file):
    texts = []
    # tenta extrair texto nativo
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                texts.append(txt)

    # se nada, tenta OCR
    if not texts:
        st.info("PDF sem texto detectável. Usando OCR...")
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img, lang="por"))

    return texts

def valor_br_para_float(valor_str):
    """
    Converte string no formato BR (1.234,56 ou 123,45) para float.
    Retorna float com duas casas (rounded).
    """
    if valor_str is None:
        return 0.0
    # remover espaços
    v = str(valor_str).strip()
    # remover pontos de milhar e trocar vírgula por ponto decimal
    v = v.replace(".", "").replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return 0.0

# ===============================
# Extrair Tabela 1 – Transações
# ===============================
def extract_tabela_transacoes(text):
    """
    Extrai linhas com: Data | Estabelecimento | Valor (formato BR)
    Retorna DataFrame com Valor como float (2 casas).
    """
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    matches = re.findall(pattern, text, re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df = pd.DataFrame(matches, columns=["Data", "Estabelecimento", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

# ===============================
# Extrair Tabela 2 – Favorecidos
# ===============================
def extract_tabela_favorecidos(text):
    """
    Extrai linhas com:
    Data | Canal | Tipo | Favorecido | ISPB* | Agência | Conta | Valor (R$)
    Para validação exigimos tais campos; retornamos apenas Data, Favorecido, Valor (float).
    """
    # Regex mais restritiva: Data + Canal (palavra), Tipo (palavras), Favorecido (captura não-gulosa), ISPB (8 dígitos),
    # Agência (3-5 dígitos), Conta (números e hífens), Valor (formato BR)
    pattern = (
        r"(\d{2}/\d{2})\s+"          # Data
        r"(\S+)\s+"                  # Canal (ex: PIX, TED)
        r"([A-Z0-9\s]+?)\s+"         # Tipo (ex: TRANSFERENCIA) - admite números e espaços
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+"  # Favorecido (mais permissivo com acentos, pontos, hífens)
        r"(\d{8})\s+"                # ISPB (8 dígitos)
        r"(\d{3,5})\s+"              # Agência
        r"([\d\-]+)\s+"              # Conta
        r"([\d.,]+)"                 # Valor (ex: 1.234,56 ou 123,45)
    )

    matches = re.findall(pattern, text, re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    # matches columns: Data, Canal, Tipo, Favorecido, ISPB, Agência, Conta, Valor
    df_full = pd.DataFrame(matches, columns=[
        "Data", "Canal", "Tipo", "Favorecido", "ISPB", "Agência", "Conta", "Valor (raw)"
    ])

    # converter Valor e manter só as colunas desejadas
    df = pd.DataFrame()
    df["Data"] = df_full["Data"]
    # limpar espaços extras em Favorecido
    df["Favorecido"] = df_full["Favorecido"].str.strip()
    df["Valor (R$)"] = df_full["Valor (raw)"].apply(valor_br_para_float)

    return df

# ===============================
# Processamento principal
# ===============================
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
            st.subheader("Pré-visualização das tabelas encontradas")
            if listas_transacoes:
                preview_t = pd.concat(listas_transacoes, ignore_index=True)
                st.write("Transações (amostra):")
                st.dataframe(preview_t)
            if listas_favorecidos:
                preview_f = pd.concat(listas_favorecidos, ignore_index=True)
                st.write("Favorecidos (Data | Favorecido | Valor):")
                st.dataframe(preview_f)

            nome_arquivo = st.text_input("Nome do arquivo Excel (sem .xlsx)", "fatura_extraida")

            if st.button("Gerar Excel"):
                # montar DataFrames finais (concatenar páginas)
                if listas_transacoes:
                    df_transacoes = pd.concat(listas_transacoes, ignore_index=True)
                else:
                    df_transacoes = pd.DataFrame(columns=["Data", "Estabelecimento", "Valor (R$)"])

                if listas_favorecidos:
                    df_favorecidos = pd.concat(listas_favorecidos, ignore_index=True)
                else:
                    df_favorecidos = pd.DataFrame(columns=["Data", "Favorecido", "Valor (R$)"])

                # preparar BytesIO
                output = BytesIO()

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    style = TableStyleInfo(
                        name="TableStyleMedium9",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False
                    )

                    # Escrever Transações e transformar em Table
                    if not df_transacoes.empty:
                        sheet_name = "Transacoes"
                        df_transacoes.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]

                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"

                        tabela = Table(displayName="TabelaTransacoes", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

                        # Formatar a última coluna (Valor) como número com 2 casas
                        col_letter = get_column_letter(max_col)
                        for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                            for cell in row:
                                cell.number_format = '#,##0.00'

                    # Escrever Favorecidos e transformar em Table
                    if not df_favorecidos.empty:
                        sheet_name = "Favorecidos"
                        df_favorecidos.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]

                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"

                        tabela = Table(displayName="TabelaFavorecidos", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

                        # Formatar a última coluna (Valor) como número com 2 casas
                        col_letter = get_column_letter(max_col)
                        for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                            for cell in row:
                                cell.number_format = '#,##0.00'

                    # salvar o arquivo no BytesIO (ExcelWriter faz isso no exit)

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
