
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

st.set_page_config(page_title="Extrair Fatura", layout="wide")
st.title("Extrair Débitos da Fatura")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ================= Funções =================

def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip()

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

    df = pd.DataFrame(matches, columns=["Data","Estabelecimento","Valor (R$)"])
    df["Valor (R$)"] = df["Valor (R$)"].apply(valor_br_para_float)

    return df

def extract_tabela_favorecidos(text):

    pattern = (
        r"(\d{2}/\d{2})\s+(\S+)\s+([A-Z0-9\s]+?)\s+"
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+(\d{8})\s+"
        r"(\d{3,5})\s+([\d\-]+)\s+([\d.,]+)"
    )

    matches = re.findall(pattern, text, re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df_full = pd.DataFrame(matches)

    df = pd.DataFrame()
    df["Data"] = df_full[0]
    df["Favorecido"] = df_full[3]
    df["Valor (R$)"] = df_full[7].apply(valor_br_para_float)

    return df

def extrair_mes_ano(nome):

    mes_ano = re.search(r"([A-Z]{3})\s*(\d{4})", nome.upper())

    if mes_ano:

        mes_abrev, ano = mes_ano.groups()

        meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]

        try:
            mes = meses.index(mes_abrev)+1
        except:
            mes = 1

        return datetime(int(ano), mes, 1)

    return datetime.now()

# ================= Processamento =================

if uploaded_file:

    if "df_transacoes" not in st.session_state:

        texts = extract_text_from_pdf(uploaded_file)

        listas_transacoes=[]
        listas_pix=[]

        for t in texts:

            d = extract_tabela_transacoes(t)
            if not d.empty:
                listas_transacoes.append(d)

            p = extract_tabela_favorecidos(t)
            if not p.empty:
                listas_pix.append(p)

        st.session_state.df_transacoes = (
            pd.concat(listas_transacoes,ignore_index=True)
            if listas_transacoes else pd.DataFrame(columns=["Data","Estabelecimento","Valor (R$)"])
        )

        st.session_state.df_pix = (
            pd.concat(listas_pix,ignore_index=True)
            if listas_pix else pd.DataFrame(columns=["Data","Favorecido","Valor (R$)"])
        )

    # ================= Inserção Manual Débito =================

    st.subheader("Inserir Débito Manual")

    c1,c2,c3,c4 = st.columns([1,4,2,1])

    data_manual = c1.text_input("Data", key="deb_data")
    desc_manual = c2.text_input("Descrição", key="deb_desc")
    valor_manual = c3.number_input("Valor", min_value=0.0, step=0.01, key="deb_valor")

    if c4.button("Adicionar", key="add_debito"):

        if desc_manual and valor_manual:

            nova = pd.DataFrame([{
                "Data":data_manual,
                "Estabelecimento":desc_manual,
                "Valor (R$)":float(valor_manual)
            }])

            st.session_state.df_transacoes = pd.concat(
                [st.session_state.df_transacoes,nova],
                ignore_index=True
            )

            st.session_state.deb_data=""
            st.session_state.deb_desc=""
            st.session_state.deb_valor=0.0

            st.rerun()

    # ================= Lista Débitos =================

    if not st.session_state.df_transacoes.empty:

        st.subheader("Débitos")

        for i,row in st.session_state.df_transacoes.iterrows():

            a,b,c,d = st.columns([1,4,2,0.5])

            a.write(row["Data"])
            b.write(row["Estabelecimento"])
            c.write(f"R$ {row['Valor (R$)']:,.2f}")

            if d.button("🗑️",key=f"del_t{i}"):

                st.session_state.df_transacoes.drop(i,inplace=True)
                st.session_state.df_transacoes.reset_index(drop=True,inplace=True)

                st.rerun()

        total_debito = st.session_state.df_transacoes["Valor (R$)"].sum()

        st.info(f"Total Débitos: R$ {total_debito:,.2f}")

    # ================= Inserir PIX =================

    st.subheader("Inserir PIX Manual")

    c1,c2,c3,c4 = st.columns([1,4,2,1])

    data_pix = c1.text_input("Data", key="pix_data")
    fav = c2.text_input("Favorecido", key="pix_desc")
    valor_pix = c3.number_input("Valor PIX", min_value=0.0, step=0.01, key="pix_valor")

    if c4.button("Adicionar PIX"):

        if fav and valor_pix:

            nova = pd.DataFrame([{
                "Data":data_pix,
                "Favorecido":fav,
                "Valor (R$)":float(valor_pix)
            }])

            st.session_state.df_pix = pd.concat(
                [st.session_state.df_pix,nova],
                ignore_index=True
            )

            st.session_state.pix_data=""
            st.session_state.pix_desc=""
            st.session_state.pix_valor=0.0

            st.rerun()

    # ================= Lista PIX =================

    if not st.session_state.df_pix.empty:

        st.subheader("Envios PIX")

        for i,row in st.session_state.df_pix.iterrows():

            a,b,c,d = st.columns([1,4,2,0.5])

            a.write(row["Data"])
            b.write(row["Favorecido"])
            c.write(f"R$ {row['Valor (R$)']:,.2f}")

            if d.button("🗑️",key=f"del_p{i}"):

                st.session_state.df_pix.drop(i,inplace=True)
                st.session_state.df_pix.reset_index(drop=True,inplace=True)

                st.rerun()

        total_pix = st.session_state.df_pix["Valor (R$)"].sum()

        st.info(f"Total PIX: R$ {total_pix:,.2f}")

    # ================= Excel =================

    nome_arquivo = st.text_input("Nome do Excel","fatura")

    vencimento = extrair_mes_ano(nome_arquivo)

    output = BytesIO()

    df_excel = pd.concat([
        st.session_state.df_transacoes.rename(columns={"Estabelecimento":"Descrição","Valor (R$)":"Valor"})[["Data","Descrição","Valor"]],
        st.session_state.df_pix.rename(columns={"Favorecido":"Descrição","Valor (R$)":"Valor"})[["Data","Descrição","Valor"]]
    ])

    total_geral = df_excel["Valor"].sum()

    df_excel.loc[len(df_excel)] = ["","TOTAL",total_geral]

    with pd.ExcelWriter(output,engine="openpyxl") as writer:

        df_excel.to_excel(writer,index=False)

    output.seek(0)

    st.download_button(
        "Baixar Excel",
        data=output,
        file_name=sanitize_filename(nome_arquivo)+".xlsx"
    )

    # ================= SharePoint =================

    if st.button("Enviar total para SharePoint"):

        try:

            app = msal.ConfidentialClientApplication(
                client_id=os.getenv("AZURE_CLIENT_ID"),
                client_credential=os.getenv("AZURE_CLIENT_SECRET"),
                authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}"
            )

            token = app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )

            access_token = token.get("access_token")

            SITE_ID="devgbsn.sharepoint.com,351e9978-140f-427e-a87d-332f6ce67a46,fc4e159a-5954-442f-a08f-28617bc84da1"
            LIST_ID="b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

            url=f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"

            payload={
                "fields":{
                    "Despesa":f"Despesa Germano {nome_arquivo}",
                    "Valor":float(total_geral),
                    "Vencimento":vencimento.strftime("%m/%d/%Y"),
                    "QuemPagou":"Germano",
                    "pago":"sim"
                }
            }

            response=requests.post(
                url,
                headers={
                    "Authorization":f"Bearer {access_token}",
                    "Content-Type":"application/json"
                },
                json=payload
            )

            if response.status_code==201:
                st.success("Enviado para SharePoint")
            else:
                st.error(response.text)

        except Exception as e:

            st.error(e)

