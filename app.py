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

st.set_page_config(page_title="Extrair Fatura para Excel e SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf")

# ================= SESSION =================

if "edit_debito" not in st.session_state:
    st.session_state.edit_debito=None

if "edit_pix" not in st.session_state:
    st.session_state.edit_pix=None

# ================= Funções =================

def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars).strip() or "fatura"

def valor_br_para_float(valor_str):
    if valor_str is None:
        return 0.0
    v = str(valor_str).strip()
    v = v.replace("R$", "").replace(" ", "")
    v = v.replace(".", "").replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return 0.0

def parse_valor_input(valor_str):
    if not valor_str:
        return None
    v = valor_str.replace(" ", "")
    v = v.replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return None

def formatar_data(d):
    if d:
        return d.strftime("%d/%m")
    return ""

def extract_text_from_pdf(file):

    texts=[]

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt=page.extract_text()
            if txt:
                texts.append(txt)

    if not texts:
        file.seek(0)
        images=convert_from_bytes(file.read())
        for img in images:
            texts.append(pytesseract.image_to_string(img,lang="por"))

    return texts

def extract_tabela_transacoes(text):

    pattern=r"(\d{2}/\d{2})\s+[\d.]+\s+(.+?)\s+([\d.,]+)$"

    matches=re.findall(pattern,text,re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df=pd.DataFrame(matches,columns=["Data","Estabelecimento","Valor (R$)"])
    df["Valor (R$)"]=df["Valor (R$)"].apply(valor_br_para_float)

    return df

def extract_tabela_favorecidos(text):

    pattern=(
        r"(\d{2}/\d{2})\s+(\S+)\s+([A-Z0-9\s]+?)\s+"
        r"([A-ZÀ-Ÿa-zà-ÿ0-9\.\- ]+?)\s+(\d{8})\s+"
        r"(\d{3,5})\s+([\d\-]+)\s+([\d.,]+)"
    )

    matches=re.findall(pattern,text,re.MULTILINE)

    if not matches:
        return pd.DataFrame()

    df_full=pd.DataFrame(matches)

    df=pd.DataFrame()

    df["Data"]=df_full[0]
    df["Favorecido"]=df_full[3].str.strip()
    df["Valor (R$)"]=df_full[7].apply(valor_br_para_float)

    return df

def extrair_mes_ano(nome):

    mes_ano=re.search(r"([A-Z]{3})\s*(\d{4})",nome.upper())

    if mes_ano:

        mes_abrev,ano=mes_ano.groups()

        meses=["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]

        try:
            mes=meses.index(mes_abrev)+1
        except:
            mes=1

        return datetime(int(ano),mes,1)

    return datetime.now()

# ================= Processamento =================

if uploaded_file:

    uploaded_file.seek(0)

    if "df_transacoes" not in st.session_state:

        texts=extract_text_from_pdf(uploaded_file)

        lista_t=[]
        lista_p=[]

        for t in texts:

            d=extract_tabela_transacoes(t)
            if not d.empty:
                lista_t.append(d)

            p=extract_tabela_favorecidos(t)
            if not p.empty:
                lista_p.append(p)

        st.session_state.df_transacoes=(
            pd.concat(lista_t,ignore_index=True)
            if lista_t else pd.DataFrame(columns=["Data","Estabelecimento","Valor (R$)"])
        )

        st.session_state.df_pix=(
            pd.concat(lista_p,ignore_index=True)
            if lista_p else pd.DataFrame(columns=["Data","Favorecido","Valor (R$)"])
        )

# ================= Inserção Manual Débito =================

with st.form("form_debito", clear_on_submit=True):

    st.subheader("Inserir Débito Manual")

    c1,c2,c3,c4=st.columns([2,4,2,1])

    data_manual=c1.date_input("Data")
    desc_manual=c2.text_input("Descrição")
    valor_manual=c3.text_input("Valor")

    submitted=c4.form_submit_button("Adicionar")

    if submitted:

        valor=parse_valor_input(valor_manual)

        if desc_manual and valor is not None:

            nova=pd.DataFrame([{
                "Data":formatar_data(data_manual),
                "Estabelecimento":desc_manual.upper(),
                "Valor (R$)":valor
            }])

            st.session_state.df_transacoes=pd.concat(
                [st.session_state.df_transacoes,nova],
                ignore_index=True
            )

            st.session_state.df_transacoes.sort_values("Data",inplace=True)
            st.session_state.df_transacoes.reset_index(drop=True,inplace=True)

            st.rerun()

    

# ================= LISTA DÉBITOS =================

# ================= LISTA DÉBITOS =================

if "df_transacoes" in st.session_state and not st.session_state.df_transacoes.empty:

    st.subheader("Débitos")

    # ===== formulário edição (AGORA FICA ACIMA)

    if st.session_state.edit_debito is not None:

        idx = st.session_state.edit_debito
        row = st.session_state.df_transacoes.loc[idx]

        st.markdown("### Editar Débito")

        with st.form("editar_debito"):

            c1,c2,c3 = st.columns(3)

            data = c1.text_input("Data", row["Data"])
            desc = c2.text_input("Descrição", row["Estabelecimento"])
            valor = c3.text_input("Valor", str(row["Valor (R$)"]))

            salvar = st.form_submit_button("Salvar")
            cancelar = st.form_submit_button("Cancelar")

            if salvar:

                v = parse_valor_input(valor)

                if v is not None:

                    st.session_state.df_transacoes.loc[idx] = [
                        data,
                        desc.upper(),
                        v
                    ]

                    st.session_state.edit_debito = None
                    st.rerun()

            if cancelar:
                st.session_state.edit_debito = None
                st.rerun()

    # ===== lista de itens

    for i,row in st.session_state.df_transacoes.iterrows():

        a,b,c,d,e = st.columns([1,4,2,0.5,0.5])

        a.write(row["Data"])
        b.write(row["Estabelecimento"])
        c.write(f"R$ {row['Valor (R$)']:,.2f}")

        if d.button("✏️", key=f"edit_t_{i}"):
            st.session_state.edit_debito = i
            st.rerun()

        if e.button("🗑️", key=f"del_t_{i}"):

            st.session_state.df_transacoes.drop(i, inplace=True)
            st.session_state.df_transacoes.reset_index(drop=True, inplace=True)

            st.rerun()

    total_debito = st.session_state.df_transacoes["Valor (R$)"].sum()

    st.info(f"Total Débitos: R$ {total_debito:,.2f}")

# ================= LISTA PIX =================

if "df_pix" in st.session_state and not st.session_state.df_pix.empty:

    st.subheader("PIX")

    for i,row in st.session_state.df_pix.iterrows():

        a,b,c,d,e=st.columns([1,4,2,0.5,0.5])

        a.write(row["Data"])
        b.write(row["Favorecido"])
        c.write(f"R$ {row['Valor (R$)']:,.2f}")

        if d.button("✏️",key=f"edit_p_{i}"):
            st.session_state.edit_pix=i

        if e.button("🗑️",key=f"del_p_{i}"):

            st.session_state.df_pix.drop(i,inplace=True)
            st.session_state.df_pix.reset_index(drop=True,inplace=True)

            st.rerun()

    # ===== edição pix

    if st.session_state.edit_pix is not None:

        idx=st.session_state.edit_pix
        row=st.session_state.df_pix.loc[idx]

        st.markdown("### Editar PIX")

        with st.form("editar_pix"):

            c1,c2,c3=st.columns(3)

            data=c1.text_input("Data",row["Data"])
            fav=c2.text_input("Favorecido",row["Favorecido"])
            valor=c3.text_input("Valor",str(row["Valor (R$)"]))

            if st.form_submit_button("Salvar"):

                v=parse_valor_input(valor)

                if v is not None:

                    st.session_state.df_pix.loc[idx]=[
                        data,
                        fav.upper(),
                        v
                    ]

                    st.session_state.edit_pix=None
                    st.rerun()

    total_pix=st.session_state.df_pix["Valor (R$)"].sum()

    st.info(f"Total PIX: R$ {total_pix:,.2f}")
    # ================= Excel =================

nome_arquivo = st.text_input("Nome do arquivo Excel", "fatura")

vencimento = extrair_mes_ano(nome_arquivo)

if st.button("Gerar Excel"):

    output = BytesIO()

    df_excel_list = []

    if not st.session_state.df_transacoes.empty:
        df_excel_list.append(
            st.session_state.df_transacoes
            .rename(columns={"Estabelecimento":"Descrição","Valor (R$)":"Valor"})
            [["Data","Descrição","Valor"]]
        )

    if not st.session_state.df_pix.empty:
        df_excel_list.append(
            st.session_state.df_pix
            .rename(columns={"Favorecido":"Descrição","Valor (R$)":"Valor"})
            [["Data","Descrição","Valor"]]
        )

    df_excel = pd.concat(df_excel_list, ignore_index=True)

    total_geral = df_excel["Valor"].sum()

    df_excel.loc[len(df_excel)] = ["","TOTAL",total_geral]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        sheet = "Fatura"

        df_excel.to_excel(writer, sheet_name=sheet, index=False)

        ws = writer.book[sheet]

        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

        tabela = Table(displayName="TabelaFatura", ref=ref)

        tabela.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showRowStripes=True
        )

        ws.add_table(tabela)

        for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            for cell in row:
                cell.number_format='#,##0.00'

    output.seek(0)

    st.download_button(
        "Baixar Excel",
        data=output,
        file_name=sanitize_filename(nome_arquivo)+".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ================= SharePoint =================

if st.button("Enviar total para SharePoint"):

    total_geral = (
        st.session_state.df_transacoes["Valor (R$)"].sum() +
        st.session_state.df_pix["Valor (R$)"].sum()
    )

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
            st.success("Enviado para SharePoint com sucesso")
        else:
            st.error(response.text)

    except Exception as e:
        st.error(e)




