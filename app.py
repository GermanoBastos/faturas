import streamlit as st
import pandas as pd
import pdfplumber
import re
import string
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import msal
import requests
import datetime

# ------------------------ CONFIGURAÇÃO DA PÁGINA ------------------------
st.set_page_config(page_title="Extrair Fatura e SharePoint", layout="wide")
st.title("Extrair Débitos e PIX (com Totais, Excel e SharePoint)")

# ------------------------ UPLOAD DO PDF ------------------------
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ------------------------ FUNÇÕES UTILITÁRIAS ------------------------
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

def extrair_vencimento(nome_arquivo):
    partes = nome_arquivo.upper().split()
    if len(partes) >= 2:
        mes_str = partes[0][:3]
        ano_str = partes[1]
        meses = {
            "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
            "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
        }
        mes = meses.get(mes_str, 1)
        try:
            ano = int(ano_str)
        except:
            ano = datetime.datetime.now().year
        return f"{ano}-{mes:02d}-01"
    else:
        hoje = datetime.datetime.now()
        return f"{hoje.year}-{hoje.month:02d}-01"

# ------------------------ PROCESSAMENTO PRINCIPAL ------------------------
if uploaded_file:
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
        st.subheader("Resumo da Fatura")
        df_transacoes = pd.concat(listas_transacoes, ignore_index=True) if listas_transacoes else pd.DataFrame()
        df_favorecidos = pd.concat(listas_favorecidos, ignore_index=True) if listas_favorecidos else pd.DataFrame()

        # Totais
        total_transacoes = df_transacoes["Valor (R$)"].sum() if not df_transacoes.empty else 0.0
        total_favorecidos = df_favorecidos["Valor (R$)"].sum() if not df_favorecidos.empty else 0.0
        total_geral = total_transacoes + total_favorecidos

        st.write("📊 Tabelas:")
        st.dataframe(df_transacoes)
        st.dataframe(df_favorecidos)
        st.info(f"💰 Total Geral (Débitos + PIX): R$ {total_geral:,.2f}")

        default_name = uploaded_file.name.rsplit(".", 1)[0]
        nome_arquivo = st.text_input("Nome do arquivo Excel (sem .xlsx)", value=default_name)

        # ------------------------ GERAR EXCEL ------------------------
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

                sheet_name = "Fatura"
                ws = writer.book.create_sheet(sheet_name)

                # Inserir Débitos
                if not df_transacoes.empty:
                    df_transacoes.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                    ws = writer.book[sheet_name]
                    max_row = ws.max_row
                    max_col = ws.max_column
                    ref = f"A1:{get_column_letter(max_col)}{max_row}"
                    tabela = Table(displayName="TabelaDebitos", ref=ref)
                    tabela.tableStyleInfo = style
                    ws.add_table(tabela)

                # Inserir PIX logo abaixo
                if not df_favorecidos.empty:
                    startrow = ws.max_row + 2
                    df_favorecidos.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)
                    ws = writer.book[sheet_name]
                    max_row = ws.max_row
                    max_col = ws.max_column
                    ref = f"A{startrow+1}:{get_column_letter(max_col)}{max_row}"
                    tabela = Table(displayName="TabelaPIX", ref=ref)
                    tabela.tableStyleInfo = style
                    ws.add_table(tabela)

                # Inserir total no final
                total_row = ws.max_row + 2
                ws[f"A{total_row}"] = "TOTAL GERAL"
                ws[f"B{total_row}"] = total_geral

            output.seek(0)
            st.success("Excel gerado com sucesso — pronto para download.")
            st.download_button(
                label="📥 Baixar Excel",
                data=output,
                file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ------------------------ ENVIAR PARA SHAREPOINT ------------------------
        if st.button("Enviar Total para SharePoint"):
            # Conexão via Secret
            CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
            TENANT_ID = os.getenv("AZURE_TENANT_ID")
            CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

            if not CLIENT_SECRET:
                st.error("CLIENT_SECRET não encontrado como variável de ambiente")
            else:
                app = msal.ConfidentialClientApplication(
                    client_id=CLIENT_ID,
                    client_credential=CLIENT_SECRET,
                    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
                )
                token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
                access_token = token.get("access_token")

                if not access_token:
                    st.error("❌ Erro ao obter token do Azure")
                else:
                    # IDs do site e lista
                    SITE_ID = (
                        "devgbsn.sharepoint.com,"
                        "351e9978-140f-427e-a87d-332f6ce67a46,"
                        "fc4e159a-5954-442f-a08f-28617bc84da1"
                    )
                    LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

                    vencimento = extrair_vencimento(nome_arquivo)

                    payload_total = {
                        "fields": {
                            "Title": "Total Geral Fatura",
                            "Despesa": "Total Geral",
                            "Valor": total_geral,
                            "Vencimento": vencimento
                        }
                    }

                    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
                    headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/json"
                    }

                    response = requests.post(url, headers=headers, json=payload_total)

                    if response.status_code == 201:
                        st.success("✅ Total enviado com sucesso para o SharePoint")
                    else:
                        st.error(f"❌ Erro ao enviar para SharePoint: {response.status_code}")
                        st.text(response.text)
