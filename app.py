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
# FUNÇÕES SHAREPOINT
# =====================================================
def get_token():
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)
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
# STREAMLIT CONFIG
# =====================================================
st.set_page_config(page_title="Extrair Fatura → SharePoint", layout="wide")
st.title("📄 Extrair Fatura e Enviar Totais para SharePoint")

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

        # Concatenar tabelas
        df_transacoes = pd.concat(dfs_transacoes, ignore_index=True) if dfs_transacoes else pd.DataFrame()
        df_favorecidos = pd.concat(dfs_favorecidos, ignore_index=True) if dfs_favorecidos else pd.DataFrame()

        # Mostrar no Streamlit
        if not df_transacoes.empty:
            st.subheader("💳 Débitos")
            st.dataframe(df_transacoes)
        if not df_favorecidos.empty:
            st.subheader("🔁 PIX")
            st.dataframe(df_favorecidos)

        # Totais
        total_debitos = df_transacoes["Valor (R$)"].sum() if not df_transacoes.empty else 0.0
        total_pix = df_favorecidos["Valor (R$)"].sum() if not df_favorecidos.empty else 0.0

        st.info(f"💰 Total Débitos: R$ {total_debitos:,.2f}")
        st.info(f"💰 Total PIX: R$ {total_pix:,.2f}")

        # =====================================================
        # Botão Excel
        # =====================================================
        default_name = uploaded_file.name.rsplit(".", 1)[0]
        nome_arquivo = st.text_input("Nome do arquivo Excel (sem .xlsx)", value=default_name)

        if st.button("💾 Gerar Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                style = TableStyleInfo(
                    name="TableStyleMedium9",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )

                # Transacoes
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
                    for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                        for cell in row:
                            cell.number_format = '#,##0.00'
                    ws.cell(row=max_row+1, column=max_col-1, value="TOTAL")
                    ws.cell(row=max_row+1, column=max_col, value=f"=SUM({get_column_letter(max_col)}2:{get_column_letter(max_col)}{max_row})")
                    ws.cell(row=max_row+1, column=max_col).number_format = '#,##0.00'

                # Favorecidos
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
                    for row in ws.iter_rows(min_row=2, min_col=max_col, max_col=max_col, max_row=max_row):
                        for cell in row:
                            cell.number_format = '#,##0.00'
                    ws.cell(row=max_row+1, column=max_col-1, value="TOTAL")
                    ws.cell(row=max_row+1, column=max_col, value=f"=SUM({get_column_letter(max_col)}2:{get_column_letter(max_col)}{max_row})")
                    ws.cell(row=max_row+1, column=max_col).number_format = '#,##0.00'

            output.seek(0)
            st.success("Excel gerado com sucesso — pronto para download.")
            st.download_button(
                label="📥 Baixar Excel",
                data=output,
                file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # =====================================================
        # Botão SharePoint (enviar somente totais)
        # =====================================================
        if st.button("📤 Enviar Totais para SharePoint"):
            token = get_token()
            total_registros = 0

            if total_debitos > 0:
                inserir_sharepoint(token, "Débitos", total_debitos)
                total_registros += 1
            if total_pix > 0:
                inserir_sharepoint(token, "PIX", total_pix)
                total_registros += 1

            st.success(f"✅ {total_registros} registros enviados ao SharePoint com sucesso!")

    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
