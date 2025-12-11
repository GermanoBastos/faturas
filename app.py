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
import os
import msal
import requests
from datetime import datetime

# ================== CONFIG ==================
st.set_page_config(
    page_title="Extrair Fatura para Excel e SharePoint",
    layout="wide"
)

st.title("Extrair Débitos da Fatura (Excel + SharePoint)")

MAX_MB = 10
LIMITE_LINHAS_MOBILE = 25

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ================== FUNÇÕES ==================
def sanitize_filename(name):
    valid = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid).strip() or "fatura_extraida"

def valor_br_para_float(v):
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except:
        return 0.0

def extract_text_from_pdf(file, usar_ocr=False):
    textos = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                textos.append(t)

    if not textos and usar_ocr:
        file.seek(0)
        for img in convert_from_bytes(file.read()):
            textos.append(pytesseract.image_to_string(img, lang="por"))

    return textos

def extract_tabela_transacoes(text):
    p = r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"
    m = re.findall(p, text, re.MULTILINE)
    if not m:
        return pd.DataFrame()
    df = pd.DataFrame(m, columns=["Data","Estabelecimento","Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

def extract_tabela_favorecidos(text):
    p = (
        r"(\d{2}/\d{2})\s+\S+\s+[A-Z0-9\s]+?\s+"
        r"([A-ZÀ-ÿ0-9\.\- ]+?)\s+\d+\s+\d+\s+[\d\-]+\s+([\d.,]+)"
    )
    m = re.findall(p, text, re.MULTILINE)
    if not m:
        return pd.DataFrame()
    df = pd.DataFrame(m, columns=["Data","Favorecido","Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)
    return df

def extrair_mes_ano(nome):
    r = re.search(r"([A-Z]{3})\s*(\d{4})", nome.upper())
    meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
    if r:
        m, a = r.groups()
        mes = meses.index(m)+1 if m in meses else 1
        return datetime(int(a), mes, 1)
    return datetime.now()

# ================== EXEC ==================
if uploaded_file:

    if uploaded_file.size > MAX_MB * 1024 * 1024:
        st.error("📱 PDF muito grande para celular. Use computador.")
        st.stop()

    st.session_state["uploaded_filename"] = uploaded_file.name
    usar_ocr = st.checkbox("Usar OCR (mais lento – desktop recomendado)", value=False)

    if "df_transacoes" not in st.session_state:

        with st.spinner("🔎 Processando PDF..."):
            textos = extract_text_from_pdf(uploaded_file, usar_ocr)

        lt, lf = [], []
        for t in textos:
            if not extract_tabela_transacoes(t).empty:
                lt.append(extract_tabela_transacoes(t))
            if not extract_tabela_favorecidos(t).empty:
                lf.append(extract_tabela_favorecidos(t))

        st.session_state.df_transacoes = (
            pd.concat(lt, ignore_index=True) if lt else pd.DataFrame()
        )
        st.session_state.df_favorecidos = (
            pd.concat(lf, ignore_index=True) if lf else pd.DataFrame()
        )

    # ================== DÉBITOS ==================
    st.markdown("## Débitos")
    df = st.session_state.df_transacoes

    if len(df) <= LIMITE_LINHAS_MOBILE:
        for i, r in df.iterrows():
            c1,c2,c3,c4 = st.columns([1,4,2,0.5])
            c1.write(r["Data"])
            c2.write(r["Estabelecimento"])
            c3.write(f"R$ {r['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"dt{i}"):
                df.drop(i, inplace=True)
                df.reset_index(drop=True, inplace=True)
                st.rerun()
    else:
        st.warning("📱 Muitas linhas – modo otimizado")
        st.dataframe(df, use_container_width=True)
        sel = st.multiselect("Excluir linhas", df.index)
        if st.button("🗑️ Excluir selecionadas"):
            df.drop(sel, inplace=True)
            df.reset_index(drop=True, inplace=True)
            st.rerun()

    total_debitos = df["Valor (R$)"].sum()
    st.info(f"💰 Total Débitos: R$ {total_debitos:,.2f}")

    # ================== PIX ==================
    st.markdown("## PIX")
    dfp = st.session_state.df_favorecidos

    if len(dfp) <= LIMITE_LINHAS_MOBILE:
        for i, r in dfp.iterrows():
            c1,c2,c3,c4 = st.columns([1,4,2,0.5])
            c1.write(r["Data"])
            c2.write(r["Favorecido"])
            c3.write(f"R$ {r['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"pf{i}"):
                dfp.drop(i, inplace=True)
                dfp.reset_index(drop=True, inplace=True)
                st.rerun()
    else:
        st.warning("📱 Muitas linhas – modo otimizado")
        st.dataframe(dfp, use_container_width=True)
        sel = st.multiselect("Excluir linhas PIX", dfp.index)
        if st.button("🗑️ Excluir selecionadas"):
            dfp.drop(sel, inplace=True)
            dfp.reset_index(drop=True, inplace=True)
            st.rerun()

    total_pix = dfp["Valor (R$)"].sum()
    st.info(f"💰 Total PIX: R$ {total_pix:,.2f}")

    # ================== EXCEL ==================
    nome_base = st.session_state["uploaded_filename"].rsplit(".",1)[0]
    nome_arquivo = st.text_input("Nome do Excel", nome_base)
    vencimento = extrair_mes_ano(nome_arquivo)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_all = pd.concat([
            df.rename(columns={"Estabelecimento":"Descrição","Valor (R$)":"Valor"})[["Data","Descrição","Valor"]],
            dfp.rename(columns={"Favorecido":"Descrição","Valor (R$)":"Valor"})[["Data","Descrição","Valor"]]
        ], ignore_index=True)

        total_geral = df_all["Valor"].sum()
        df_all.loc[len(df_all)] = ["","TOTAL",total_geral]

        df_all.to_excel(writer, "Fatura", index=False)
        ws = writer.book["Fatura"]

        tabela = Table(
            displayName="TabelaFatura",
            ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        )
        tabela.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showRowStripes=True
        )
        ws.add_table(tabela)

    output.seek(0)

    st.download_button(
        "📥 Baixar Excel",
        output,
        sanitize_filename(nome_arquivo)+".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ================== SHAREPOINT ==================
    if st.button("Enviar total para SharePoint"):
        app = msal.ConfidentialClientApplication(
            os.getenv("AZURE_CLIENT_ID"),
            authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}",
            client_credential=os.getenv("AZURE_CLIENT_SECRET")
        )

        token = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        headers = {
            "Authorization": f"Bearer {token['access_token']}",
            "Content-Type": "application/json"
        }

        payload = {
            "fields": {
                "Despesa": f"Despesa Germano {nome_arquivo}",
                "Valor": float(total_geral),
                "Vencimento": vencimento.strftime("%m/%d/%Y"),
                "QuemPagou": "Germano",
                "pago": "sim"
            }
        }

        url = "https://graph.microsoft.com/v1.0/sites/SEU_SITE_ID/lists/SEU_LIST_ID/items"

        r = requests.post(url, headers=headers, json=payload)

        if r.status_code == 201:
            st.success("✅ Enviado ao SharePoint")
        else:
            st.error(f"❌ Erro SharePoint: {r.text}")
