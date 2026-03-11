
import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
from datetime import datetime
import string
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import os
import msal
import requests

st.set_page_config(page_title="Extrair Fatura", layout="wide")
st.title("Extrair Débitos da Fatura")

# ---------------- SESSION STATE ----------------

if "df_transacoes" not in st.session_state:
    st.session_state.df_transacoes = pd.DataFrame(
        columns=["Data","Estabelecimento","Valor (R$)"]
    )

if "df_pix" not in st.session_state:
    st.session_state.df_pix = pd.DataFrame(
        columns=["Data","Favorecido","Valor (R$)"]
    )

if "desc_manual" not in st.session_state:
    st.session_state.desc_manual=""

if "valor_manual" not in st.session_state:
    st.session_state.valor_manual=""

if "fav_manual" not in st.session_state:
    st.session_state.fav_manual=""

if "valor_pix_manual" not in st.session_state:
    st.session_state.valor_pix_manual=""

# ---------------- FUNÇÕES ----------------

def valor_br_para_float(v):

    if v is None:
        return 0

    v=str(v).replace("R$","").replace(".","").replace(",",".").strip()

    try:
        return float(v)
    except:
        return 0

def parse_valor_input(v):

    if not v:
        return None

    v=v.replace(",",".").strip()

    try:
        return float(v)
    except:
        return None

def formatar_data(d):

    if d:
        return d.strftime("%d/%m")
    return ""

def sanitize_filename(name):

    valid=f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid)

# ---------------- EXTRAÇÃO PDF ----------------

def extract_text_from_pdf(file):

    texts=[]

    with pdfplumber.open(file) as pdf:

        for p in pdf.pages:
            t=p.extract_text()
            if t:
                texts.append(t)

    if not texts:

        file.seek(0)

        images=convert_from_bytes(file.read())

        for img in images:
            texts.append(pytesseract.image_to_string(img))

    return texts

def extract_tabela_transacoes(text):

    pattern=r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"

    matches=re.findall(pattern,text,re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df=pd.DataFrame(matches,columns=["Data","Estabelecimento","Valor (R$)"])

    df["Valor (R$)"]=df["Valor (R$)"].apply(valor_br_para_float)

    df["Estabelecimento"]=df["Estabelecimento"].str.upper()

    return df

def extract_tabela_pix(text):

    pattern=r"(\d{2}/\d{2}).+?([A-ZÀ-Ÿa-zà-ÿ ]+)\s+([\d.,]+)$"

    matches=re.findall(pattern,text,re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df=pd.DataFrame(matches,columns=["Data","Favorecido","Valor (R$)"])

    df["Valor (R$)"]=df["Valor (R$)"].apply(valor_br_para_float)

    df["Favorecido"]=df["Favorecido"].str.upper()

    return df

# ---------------- PDF ----------------

uploaded_file = st.file_uploader("Escolha o PDF",type="pdf")

if uploaded_file:

    texts=extract_text_from_pdf(uploaded_file)

    lista_t=[]
    lista_p=[]

    for t in texts:

        d=extract_tabela_transacoes(t)
        if not d.empty:
            lista_t.append(d)

        p=extract_tabela_pix(t)
        if not p.empty:
            lista_p.append(p)

    if lista_t:

        novo=pd.concat(lista_t,ignore_index=True)

        st.session_state.df_transacoes=pd.concat(
            [st.session_state.df_transacoes,novo],
            ignore_index=True
        )

    if lista_p:

        novo=pd.concat(lista_p,ignore_index=True)

        st.session_state.df_pix=pd.concat(
            [st.session_state.df_pix,novo],
            ignore_index=True
        )

# ---------------- INSERÇÃO MANUAL ----------------

st.subheader("Inserir Débito")

c1,c2,c3,c4=st.columns(4)

data_manual=c1.date_input("Data Débito")
desc=c2.text_input("Descrição",key="desc_manual").upper()
valor=c3.text_input("Valor",key="valor_manual")

if c4.button("Adicionar Débito"):

    v=parse_valor_input(valor)

    if desc and v is not None:

        nova=pd.DataFrame([{
            "Data":formatar_data(data_manual),
            "Estabelecimento":desc.strip().upper(),
            "Valor (R$)":v
        }])

        st.session_state.df_transacoes=pd.concat(
            [st.session_state.df_transacoes,nova],
            ignore_index=True
        )

        st.session_state.desc_manual=""
        st.session_state.valor_manual=""

        st.rerun()

# ---------------- INSERÇÃO PIX ----------------

st.subheader("Inserir PIX")

c1,c2,c3,c4=st.columns(4)

data_pix=c1.date_input("Data PIX")
fav=c2.text_input("Favorecido",key="fav_manual").upper()
valor_pix=c3.text_input("Valor PIX",key="valor_pix_manual")

if c4.button("Adicionar PIX"):

    v=parse_valor_input(valor_pix)

    if fav and v is not None:

        nova=pd.DataFrame([{
            "Data":formatar_data(data_pix),
            "Favorecido":fav.strip().upper(),
            "Valor (R$)":v
        }])

        st.session_state.df_pix=pd.concat(
            [st.session_state.df_pix,nova],
            ignore_index=True
        )

        st.session_state.fav_manual=""
        st.session_state.valor_pix_manual=""

        st.rerun()

# ---------------- TABELAS ----------------

st.subheader("Débitos")

st.session_state.df_transacoes=st.data_editor(
    st.session_state.df_transacoes,
    use_container_width=True,
    num_rows="dynamic"
)

total_debito=st.session_state.df_transacoes["Valor (R$)"].sum()

st.info(f"Total Débitos: R$ {total_debito:,.2f}")

st.subheader("PIX")

st.session_state.df_pix=st.data_editor(
    st.session_state.df_pix,
    use_container_width=True,
    num_rows="dynamic"
)

total_pix=st.session_state.df_pix["Valor (R$)"].sum()

st.info(f"Total PIX: R$ {total_pix:,.2f}")

# ---------------- EXCEL ----------------

nome=st.text_input("Nome do arquivo","fatura")

if st.button("Gerar Excel"):

    output=BytesIO()

    df_list=[]

    if not st.session_state.df_transacoes.empty:

        df_list.append(
            st.session_state.df_transacoes.rename(
                columns={"Estabelecimento":"Descrição","Valor (R$)":"Valor"}
            )
        )

    if not st.session_state.df_pix.empty:

        df_list.append(
            st.session_state.df_pix.rename(
                columns={"Favorecido":"Descrição","Valor (R$)":"Valor"}
            )
        )

    df_excel=pd.concat(df_list,ignore_index=True)

    total=df_excel["Valor"].sum()

    df_excel.loc[len(df_excel)]=["","TOTAL",total]

    with pd.ExcelWriter(output,engine="openpyxl") as writer:

        df_excel.to_excel(writer,sheet_name="Fatura",index=False)

        ws=writer.book["Fatura"]

        ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

        tabela=Table(displayName="TabelaFatura",ref=ref)

        tabela.tableStyleInfo=TableStyleInfo(
            name="TableStyleMedium9",
            showRowStripes=True
        )

        ws.add_table(tabela)

    output.seek(0)

    st.download_button(
        "Baixar Excel",
        data=output,
        file_name=sanitize_filename(nome)+".xlsx"
    )

# ---------------- SHAREPOINT ----------------

if st.button("Enviar total para SharePoint"):

    total=total_debito+total_pix

    try:

        app=msal.ConfidentialClientApplication(
            client_id=os.getenv("AZURE_CLIENT_ID"),
            client_credential=os.getenv("AZURE_CLIENT_SECRET"),
            authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}"
        )

        token=app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        access_token=token.get("access_token")

        url="https://graph.microsoft.com/v1.0/sites/SEU_SITE/lists/SUA_LISTA/items"

        payload={
            "fields":{
                "Despesa":f"Despesa Germano {nome}",
                "Valor":float(total),
                "QuemPagou":"Germano",
                "pago":"sim"
            }
        }

        r=requests.post(
            url,
            headers={
                "Authorization":f"Bearer {access_token}",
                "Content-Type":"application/json"
            },
            json=payload
        )

        if r.status_code==201:
            st.success("Enviado com sucesso")
        else:
            st.error(r.text)

    except Exception as e:
        st.error(e)

