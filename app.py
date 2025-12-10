import streamlit as st
import pandas as pd
import pdfplumber
import re
import string
import requests
import msal
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# =====================================================
# CONFIGURAÇÕES SHAREPOINT / GRAPH
# =====================================================
CLIENT_ID = "4039ce1c-ee58-4b19-bf86-db64da445fe1"
TENANT_ID = "02543ce8-b773-43d0-9cf1-298729881b0d"

SITE_ID = (
    "devgbsn.sharepoint.com,"
    "351e9978-140f-427e-a87d-332f6ce67a46,"
    "fc4e159a-5954-442f-a08f-28617bc84da1"
)

LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Sites.ReadWrite.All"]

# =====================================================
# FUNÇÕES AUTH / SHAREPOINT
# =====================================================
def get_token():
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY
    )

    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" not in result:
        raise Exception(result)

    return result["access_token"]

def inserir_sharepoint(token, despesa, valor):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "fields": {
            "Title": despesa,
            "Despesa": despesa,
            "Valor": valor
        }
    }

    r = requests.post(url, headers=headers, json=payload)

    if r.status_code != 201:
        raise Exception(r.text)

# =====================================================
# STREAMLIT
# =====================================================
st.set_page_config(page_title="Extrair Fatura → SharePoint", layout="wide")
st.title("📄 Extrair Fatura e Enviar para SharePoint")

uploaded_file = st.file_uploader("Selecione o PDF da fatura", type="pdf")

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura_extraida"

def valor_br_para_float(v):
    if v is None:
        return 0.0
    v = str(v).replace(".", "").replace(",", ".")
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
        file.seek(0)
        images = convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img, lang="por"))

    return texts

# =====================================================
# EXTRAÇÃO TABELAS
# =====================================================
def extract_transacoes(text):
    pattern = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    m = re.findall(pattern, text, re.MULTILINE)
    if not m:
        return pd.DataFrame()
    df = pd.DataFrame(m, columns=["Data", "Estabelecimento", "Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

def extract_favorecidos(text):
    pattern = (
        r"(\d{2}/\d{2})\s+(\S+)\s+([A-Z0-9\s]+?)\s+"
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+\d+\s+\d+\s+[\d\-]+\s+([\d.,]+)"
    )
    m = re.findall(pattern, text, re.MULTILINE)
    if not m:
        return pd.DataFrame()

    df = pd.DataFrame(m, columns=["Data", "Canal", "Tipo", "Favorecido", "Valor"])
    df["Valor (R$)"] = df["Valor"].apply(valor_br_para_float)
    return df[["Data", "Favorecido", "Valor (R$)"]]

# =====================================================
# PROCESSAMENTO
# =====================================================
if uploaded_file:
    try:
        texts = extract_text_from_pdf(uploaded_file)

        dfs_transacoes = []
        dfs_favorecidos = []

        for t in texts:
            df_t = extract_transacoes(t)
            if not df_t.empty:
                dfs_transacoes.append(df_t)

            df_f = extract_favorecidos(t)
            if not df_f.empty:
                dfs_favorecidos.append(df_f)

        if not dfs_transacoes and not dfs_favorecidos:
            st.warning("Nenhuma informação encontrada.")
            st.stop()

        if dfs_transacoes:
            df_transacoes = pd.concat(dfs_transacoes, ignore_index=True)
            st.subheader("💳 Débitos")
            st.dataframe(df_transacoes)

        if dfs_favorecidos:
            df_favorecidos = pd.concat(dfs_favorecidos, ignore_index=True)
            st.subheader("🔁 PIX")
            st.dataframe(df_favorecidos)

        # =================================================
        # BOTÃO SHAREPOINT
        # =================================================
        if st.button("📤 Enviar dados para SharePoint"):
            token = get_token()
            total = 0

            if dfs_transacoes:
                for _, row in df_transacoes.iterrows():
                    inserir_sharepoint(
                        token,
                        despesa=row["Estabelecimento"],
                        valor=row["Valor (R$)"]
                    )
                    total += 1

            if dfs_favorecidos:
                for _, row in df_favorecidos.iterrows():
                    inserir_sharepoint(
                        token,
                        despesa=row["Favorecido"],
                        valor=row["Valor (R$)"]
                    )
                    total += 1

            st.success(f"✅ {total} registros inseridos no SharePoint com sucesso!")

    except Exception as e:
        st.error(f"Erro: {e}")
