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

# ================== CONFIGURAÇÃO DA PÁGINA ==================
st.set_page_config(page_title="Extrair Fatura para Excel/SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

# ================== FUNÇÕES AUXILIARES ==================
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

# ================== UPLOAD DO PDF ==================
uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

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
            # ---------- Pré-visualização ----------
            st.subheader("Débitos e envios de PIX")

            total_debitos = 0.0
            total_pix = 0.0

            if listas_transacoes:
                df_transacoes = pd.concat(listas_transacoes, ignore_index=True)
                st.write("Débitos:")
                st.dataframe(df_transacoes)
                total_debitos = df_transacoes["Valor (R$)"].sum()
                st.info(f"💰 Total de Débitos: R$ {total_debitos:,.2f}")

            if listas_favorecidos:
                df_favorecidos = pd.concat(listas_favorecidos, ignore_index=True)
                st.write("Envios de PIX:")
                st.dataframe(df_favorecidos)
                total_pix = df_favorecidos["Valor (R$)"].sum()
                st.info(f"💰 Total de Envios de PIX: R$ {total_pix:,.2f}")

            total_geral = total_debitos + total_pix

            # ---------- Gerar Excel ----------
            default_name = uploaded_file.name.rsplit(".", 1)[0]
            nome_arquivo = st.text_input(
                "Nome do arquivo Excel (sem .xlsx)",
                value=default_name
            )

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

                    # Transações
                    if listas_transacoes:
                        sheet_name = "Transacoes"
                        df_transacoes.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]
                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"
                        tabela = Table(displayName="TabelaTransacoes", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

                    # Favorecidos
                    if listas_favorecidos:
                        sheet_name = "Favorecidos"
                        df_favorecidos.to_excel(writer, sheet_name=sheet_name, index=False)
                        ws = writer.book[sheet_name]
                        max_row = ws.max_row
                        max_col = ws.max_column
                        ref = f"A1:{get_column_letter(max_col)}{max_row}"
                        tabela = Table(displayName="TabelaFavorecidos", ref=ref)
                        tabela.tableStyleInfo = style
                        ws.add_table(tabela)

                    # Aba Totais
                    df_totais = pd.DataFrame({
                        "Tipo": ["Débitos", "PIX", "Total Geral"],
                        "Valor (R$)": [total_debitos, total_pix, total_geral]
                    })
                    df_totais.to_excel(writer, sheet_name="Totais", index=False)
                    ws = writer.book["Totais"]
                    max_row = ws.max_row
                    max_col = ws.max_column
                    ref = f"A1:{get_column_letter(max_col)}{max_row}"
                    tabela = Table(displayName="TabelaTotais", ref=ref)
                    tabela.tableStyleInfo = style
                    ws.add_table(tabela)

                output.seek(0)
                st.success("Excel gerado com sucesso — pronto para download.")
                st.download_button(
                    label="📥 Baixar Excel",
                    data=output,
                    file_name=sanitize_filename(nome_arquivo) + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # ---------- Enviar total geral para SharePoint ----------
            if st.button("Enviar total geral para SharePoint"):
                CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
                TENANT_ID = os.getenv("AZURE_TENANT_ID")
                CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

                if not CLIENT_SECRET:
                    st.error("CLIENT_SECRET não encontrado como variável de ambiente")
                else:
                    SITE_ID = "devgbsn.sharepoint.com,351e9978-140f-427e-a87d-332f6ce67a46,fc4e159a-5954-442f-a08f-28617bc84da1"
                    LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

                    app_msal = msal.ConfidentialClientApplication(
                        client_id=CLIENT_ID,
                        client_credential=CLIENT_SECRET,
                        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
                    )

                    token_response = app_msal.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
                    access_token = token_response.get("access_token")

                    if not access_token:
                        st.error("❌ Não foi possível obter token: " + str(token_response))
                    else:
                        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
                        headers = {
                            "Authorization": f"Bearer {access_token}",
                            "Content-Type": "application/json"
                        }

                        payload_total = {
                            "fields": {
                                "Title": "Total Geral Fatura",
                                "Despesa": "Total Geral",
                                "Valor": total_geral
                            }
                        }

                        response = requests.post(url, headers=headers, json=payload_total)
                        if response.status_code == 201:
                            st.success(f"✅ Total geral enviado com sucesso: R$ {total_geral:,.2f}")
                        else:
                            st.error("❌ Erro ao enviar para SharePoint")
                            st.text(f"Status: {response.status_code} {response.text}")

    except Exception as e:
        st.error(f"Erro ao processar PDF: {e}")
