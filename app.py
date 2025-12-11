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

# ================== Configuração da página ==================
st.set_page_config(page_title="Extrair Fatura para Excel e SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ================== Funções utilitárias ==================
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
    df_full = pd.DataFrame(matches, columns=[
        "Data","Canal","Tipo","Favorecido","ISPB","Agência","Conta","Valor (raw)"
    ])
    df = pd.DataFrame()
    df["Data"] = df_full["Data"]
    df["Favorecido"] = df_full["Favorecido"].str.strip()
    df["Valor (R$)"] = df_full["Valor (raw)"].apply(valor_br_para_float)
    return df

def extrair_mes_ano(nome_arquivo):
    mes_ano = re.search(r"([A-Z]{3})\s*(\d{4})", nome_arquivo.upper())
    if mes_ano:
        mes_abrev, ano = mes_ano.groups()
        try:
            mes = datetime.strptime(mes_abrev, "%b").month
        except:
            meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
            mes = meses.index(mes_abrev)+1
        return datetime(int(ano), mes, 1)
    return datetime.now()

# ================== Processamento ==================
if uploaded_file:
    uploaded_file.seek(0)
    texts = extract_text_from_pdf(uploaded_file)

    if "df_transacoes" not in st.session_state:
        listas_transacoes, listas_favorecidos = [], []

        for t in texts:
            dt = extract_tabela_transacoes(t)
            if not dt.empty:
                listas_transacoes.append(dt)

            df = extract_tabela_favorecidos(t)
            if not df.empty:
                listas_favorecidos.append(df)

        st.session_state.df_transacoes = (
            pd.concat(listas_transacoes, ignore_index=True)
            if listas_transacoes else pd.DataFrame()
        )

        st.session_state.df_favorecidos = (
            pd.concat(listas_favorecidos, ignore_index=True)
            if listas_favorecidos else pd.DataFrame()
        )

    st.subheader("Pré-visualização (exclusão manual)")

    # ================== Débitos ==================
    if not st.session_state.df_transacoes.empty:
        st.write("Débitos:")

        for i, row in st.session_state.df_transacoes.iterrows():
            c1, c2, c3, c4 = st.columns([1,4,2,0.5])
            c1.write(row["Data"])
            c2.write(row["Estabelecimento"])
            c3.write(f"R$ {row['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"del_t_{i}"):
                st.session_state.df_transacoes.drop(i, inplace=True)
                st.session_state.df_transacoes.reset_index(drop=True, inplace=True)
                st.rerun()

        total = st.session_state.df_transacoes["Valor (R$)"].sum()
        st.info(f"💰 Total de Débitos: R$ {total:,.2f}")

    # ================== PIX ==================
    if not st.session_state.df_favorecidos.empty:
        st.write("Envios de PIX:")

        for i, row in st.session_state.df_favorecidos.iterrows():
            c1, c2, c3, c4 = st.columns([1,4,2,0.5])
            c1.write(row["Data"])
            c2.write(row["Favorecido"])
            c3.write(f"R$ {row['Valor (R$)']:,.2f}")
            if c4.button("🗑️", key=f"del_f_{i}"):
                st.session_state.df_favorecidos.drop(i, inplace=True)
                st.session_state.df_favorecidos.reset_index(drop=True, inplace=True)
                st.rerun()

        total_pix = st.session_state.df_favorecidos["Valor (R$)"].sum()
        st.info(f"💰 Total PIX: R$ {total_pix:,.2f}")

# ================== Nome do arquivo ==================
default_name = uploaded_file.name.rsplit(".",1)[0]
nome_arquivo = st.text_input(
    "Nome do arquivo Excel (sem .xlsx)",
    value=default_name
)

vencimento = extrair_mes_ano(nome_arquivo)

# ================== Preparar Excel ==================
output = BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    df_excel_list = []

    if not st.session_state.df_transacoes.empty:
        df_trans_excel = (
            st.session_state.df_transacoes
            .rename(columns={
                "Estabelecimento":"Descrição",
                "Valor (R$)":"Valor"
            })[["Data","Descrição","Valor"]]
        )
        df_excel_list.append(df_trans_excel)

    if not st.session_state.df_favorecidos.empty:
        df_fav_excel = (
            st.session_state.df_favorecidos
            .rename(columns={
                "Favorecido":"Descrição",
                "Valor (R$)":"Valor"
            })[["Data","Descrição","Valor"]]
        )
        df_excel_list.append(df_fav_excel)

    if df_excel_list:
        df_excel = pd.concat(df_excel_list, ignore_index=True)
    else:
        df_excel = pd.DataFrame(columns=["Data","Descrição","Valor"])

    total_geral = df_excel["Valor"].sum()
    df_excel.loc[len(df_excel)] = ["", "TOTAL", total_geral]

    sheet_name = "Fatura"
    df_excel.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.book[sheet_name]

    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tabela = Table(displayName="TabelaFatura", ref=ref)
    tabela.tableStyleInfo = style
    ws.add_table(tabela)

    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = '#,##0.00'

output.seek(0)

st.success("✅ Excel gerado com sucesso — pronto para download.")

st.download_button(
    label="📥 Baixar Excel",
    data=output,
    file_name=sanitize_filename(nome_arquivo)+".xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ================== Enviar para SharePoint ==================
if st.button("Enviar total para SharePoint"):
    try:
        CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
        TENANT_ID = os.getenv("AZURE_TENANT_ID")
        CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

        app = msal.ConfidentialClientApplication(
            client_id=CLIENT_ID,
            client_credential=CLIENT_SECRET,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}"
        )

        token = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        access_token = token.get("access_token")
        if not access_token:
            raise Exception("Erro ao obter token do MS Graph")

        SITE_ID = "devgbsn.sharepoint.com,351e9978-140f-427e-a87d-332f6ce67a46,fc4e159a-5954-442f-a08f-28617bc84da1"
        LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"

        payload = {
            "fields": {
                "Despesa": f"Despesa Germano {nome_arquivo}",
                "Valor": float(total_geral),
                "Vencimento": vencimento.strftime("%m/%d/%Y"),
                "QuemPagou": "Germano",
                "pago": "sim"
            }
        }

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        response = requests.post(url, headers=headers, json=payload)

        if response.status_code == 201:
            st.success("✅ Total enviado com sucesso para SharePoint")
        else:
            st.error(
                f"❌ Erro ao enviar para SharePoint: "
                f"{response.status_code} {response.text}"
            )

    except Exception as e:
        st.error(f"Erro na integração SharePoint: {e}")
