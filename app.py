import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
import string
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import msal
import requests
from datetime import datetime

# ================= CONFIGURAÇÃO =================
st.set_page_config(page_title="Extrair Fatura para Excel e SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

# ================= FUNÇÕES =================
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

def get_month_year_from_filename(filename):
    # Espera formato "JAN 2025" ou algo semelhante no início do nome
    match = re.search(r"([A-Z]{3})\s*(\d{4})", filename.upper())
    if match:
        mes_abrev, ano = match.groups()
        try:
            mes = datetime.strptime(mes_abrev, "%b").month
        except ValueError:
            mes = 1
        return mes, int(ano)
    return 1, datetime.now().year

def connect_sharepoint():
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    if not CLIENT_SECRET:
        raise Exception("CLIENT_SECRET não encontrado como variável de ambiente")

    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    access_token = token.get("access_token")
    if not access_token:
        raise Exception("Falha ao obter token SharePoint")
    return access_token

def insert_item_sharepoint(access_token, site_id, list_id, descricao, valor, vencimento):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {
        "fields": {
            "Title": descricao,
            "Valor": valor,
            "Vencimento": vencimento.strftime("%Y-%m-%d")
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    return response

# ================= UPLOAD =================
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

if uploaded_file:
    try:
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
            # ================= CONSTRUIR DADOS PARA EXCEL =================
            df_excel_list = []

            total_transacoes = sum([df["Valor (R$)"].sum() for df in listas_transacoes])
            total_favorecidos = sum([df["Valor (R$)"].sum() for df in listas_favorecidos])

            # Débitos
            if listas_transacoes:
                df_t = pd.concat(listas_transacoes, ignore_index=True)
                df_t = df_t.rename(columns={"Estabelecimento":"Descrição", "Valor (R$)":"Valor"})
                df_excel_list.append(df_t[["Descrição","Valor"]])

            # PIX
            if listas_favorecidos:
                df_f = pd.concat(listas_favorecidos, ignore_index=True)
                df_f = df_f.rename(columns={"Favorecido":"Descrição", "Valor (R$)":"Valor"})
                df_excel_list.append(df_f[["Descrição","Valor"]])

            # Concat e TOTAL
            if df_excel_list:
                df_excel = pd.concat(df_excel_list, ignore_index=True)
            else:
                df_excel = pd.DataFrame(columns=["Descrição","Valor"])

            df_excel.loc[len(df_excel)] = ["TOTAL", total_transacoes + total_favorecidos]

            # ================= MOSTRAR NO STREAMLIT =================
            st.subheader("Pré-visualização do Excel")
            st.dataframe(df_excel)

            # ================= GERAR EXCEL =================
            default_name = uploaded_file.name.rsplit(".",1)[0]
            nome_arquivo = st.text_input("Nome do arquivo Excel (sem .xlsx)", value=default_name)
            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet_name = "Fatura"
                df_excel.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.book[sheet_name]
                max_row = ws.max_row
                max_col = ws.max_column
                ref = f"A1:{get_column_letter(max_col)}{max_row}"
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                                       showRowStripes=True, showColumnStripes=False)
                tabela = Table(displayName="TabelaFatura", ref=ref)
                tabela.tableStyleInfo = style
                ws.add_table(tabela)
                for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                    for cell in row:
                        cell.number_format = '#,##0.00'

            output.seek(0)
            st.download_button(
                label="📥 Baixar Excel",
                data=output,
                file_name=sanitize_filename(nome_arquivo)+".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ================= ENVIO PARA SHAREPOINT =================
            if st.button("Enviar total para SharePoint"):
                site_id = "devgbsn.sharepoint.com,351e9978-140f-427e-a87d-332f6ce67a46,fc4e159a-5954-442f-a08f-28617bc84da1"
                list_id = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"
                mes, ano = get_month_year_from_filename(uploaded_file.name)
                vencimento = datetime(ano, mes, 1)
                descricao = "Total Fatura"
                valor = total_transacoes + total_favorecidos

                try:
                    token = connect_sharepoint()
                    resp = insert_item_sharepoint(token, site_id, list_id, descricao, valor, vencimento)
                    if resp.status_code == 201:
                        st.success("✅ Total enviado para SharePoint com sucesso!")
                    else:
                        st.error(f"❌ Erro ao enviar SharePoint: {resp.status_code}\n{resp.text}")
                except Exception as e:
                    st.error(f"❌ Erro ao conectar com SharePoint: {e}")

    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
