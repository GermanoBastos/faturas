import os
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
import msal
import requests
from datetime import datetime

# ================= CONFIGURAÇÃO DA PÁGINA =================
st.set_page_config(page_title="Extrair Fatura para Excel/SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (Excel + SharePoint)")

# ================= VARIÁVEIS DO SHAREPOINT =================
SITE_ID = (
    "devgbsn.sharepoint.com,"
    "351e9978-140f-427e-a87d-332f6ce67a46,"
    "fc4e159a-5954-442f-a08f-28617bc84da1"
)
LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

if not CLIENT_SECRET:
    raise Exception("CLIENT_SECRET não encontrado como variável de ambiente")

# ================= FUNÇÕES AUXILIARES =================
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
    df = pd.DataFrame(matches, columns=["Data", "Estabelecimento", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

def extract_tabela_favorecidos(text):
    pattern = (
        r"(\d{2}/\d{2})\s+(\S+)\s+([A-Z0-9\s]+?)\s+([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+(\d{8})\s+(\d{3,5})\s+([\d\-]+)\s+([\d.,]+)"
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

def extrair_data_do_arquivo(nome_arquivo):
    """
    Extrai mês e ano do nome do arquivo (formato 'MMM YYYY' ex: 'JAN 2025')
    Retorna data ISO yyyy-mm-dd
    """
    meses = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
    }
    parts = nome_arquivo.upper().split()
    mes = 1
    ano = datetime.now().year
    for p in parts:
        if p in meses:
            mes = meses[p]
        elif p.isdigit() and len(p) == 4:
            ano = int(p)
    return f"{ano}-{mes:02d}-01"

# ================= UPLOAD DO PDF =================
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

if uploaded_file:
    uploaded_file.seek(0)
    texts = extract_text_from_pdf(uploaded_file)
    listas_transacoes, listas_favorecidos = [], []

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
        df_transacoes = pd.concat(listas_transacoes, ignore_index=True) if listas_transacoes else pd.DataFrame()
        df_favorecidos = pd.concat(listas_favorecidos, ignore_index=True) if listas_favorecidos else pd.DataFrame()

        # Soma dos totais
        total_transacoes = df_transacoes["Valor (R$)"].sum() if not df_transacoes.empty else 0
        total_favorecidos = df_favorecidos["Valor (R$)"].sum() if not df_favorecidos.empty else 0

        # Nome do arquivo
        default_name = uploaded_file.name.rsplit(".", 1)[0]
        nome_arquivo = st.text_input("Nome do arquivo Excel (sem .xlsx)", value=default_name)

        # ================= EXCEL =================
        if st.button("Gerar Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet_name = "Fatura"
                df_excel = pd.DataFrame()

                # Concatenar Débitos e PIX em uma única tabela para Excel
                if not df_transacoes.empty:
                    df_excel = pd.concat([df_excel, df_transacoes.rename(columns={"Estabelecimento":"Descrição", "Valor (R$)":"Valor"})], ignore_index=True)
                if not df_favorecidos.empty:
                    df_excel = pd.concat([df_excel, df_favorecidos.rename(columns={"Favorecido":"Descrição", "Valor (R$)":"Valor"})], ignore_index=True)

                # Adicionar linha TOTAL
                df_excel.loc[len(df_excel)] = ["TOTAL", total_transacoes + total_favorecidos]

                df_excel.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.book[sheet_name]
                max_row = ws.max_row
                max_col = ws.max_column
                ref = f"A1:{get_column_letter(max_col)}{max_row}"
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                tabela = Table(displayName="TabelaFatura", ref=ref)
                tabela.tableStyleInfo = style
                ws.add_table(tabela)
                # Formatar coluna Valor
                for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                    for cell in row:
                        cell.number_format = '#,##0.00'

            output.seek(0)
            st.success("Excel gerado com sucesso — pronto para download.")
            st.download_button(
                label="📥 Baixar Excel",
                data=output,
                file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ================= SHAREPOINT =================
        if st.button("Enviar para SharePoint"):
            # Conexão MSAL Confidential Client
            app = msal.ConfidentialClientApplication(
                client_id=CLIENT_ID,
                client_credential=CLIENT_SECRET,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}"
            )
            token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            access_token = token.get("access_token")
            if not access_token:
                st.error("Erro ao obter token de acesso")
            else:
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
                headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
                vencimento = extrair_data_do_arquivo(nome_arquivo)
                payload = {
                    "fields": {
                        "Title": nome_arquivo,
                        "Valor": round(total_transacoes + total_favorecidos, 2),
                        "Despesa": "Soma Fatura",
                        "Vencimento": vencimento
                    }
                }
                response = requests.post(url, headers=headers, json=payload)
                if response.status_code == 201:
                    st.success(f"✅ Item inserido no SharePoint (Vencimento: {vencimento})")
                else:
                    st.error(f"❌ Erro ao inserir item no SharePoint: {response.status_code}")
                    st.text(response.text)
