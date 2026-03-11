import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Extrair Fatura para Excel e SharePoint", layout="wide")
st.title("Extrair Débitos da Fatura (com Totais, Excel e SharePoint)")

# ================= SESSION =================

if "df_transacoes" not in st.session_state:
    st.session_state.df_transacoes = pd.DataFrame(
        columns=["Data", "Estabelecimento", "Valor (R$)"]
    )

if "df_pix" not in st.session_state:
    st.session_state.df_pix = pd.DataFrame(
        columns=["Data", "Favorecido", "Valor (R$)"]
    )

# ================= FUNÇÕES =================

def formatar_data(data):
    return data.strftime("%d/%m")

def ordenar_por_data(df):
    try:
        return df.sort_values(
            by="Data",
            key=lambda col: pd.to_datetime(col, format="%d/%m"),
            ignore_index=True
        )
    except:
        return df

# ================= UPLOAD PDF =================

uploaded_file = st.file_uploader("Escolha o PDF da fatura", type="pdf", key="upload_pdf")

if uploaded_file:

    texto = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

    linhas = texto.split("\n")

    dados = []

    for linha in linhas:

        match = re.search(r"(\d{2}/\d{2})\s+(.*?)\s+(-?\d+,\d{2})", linha)

        if match:

            data = match.group(1)
            desc = match.group(2).strip().upper()
            valor = float(match.group(3).replace(".", "").replace(",", "."))

            dados.append({
                "Data": data,
                "Estabelecimento": desc,
                "Valor (R$)": valor
            })

    if dados:

        df_pdf = pd.DataFrame(dados)

        st.session_state.df_transacoes = pd.concat(
            [st.session_state.df_transacoes, df_pdf],
            ignore_index=True
        )

        st.success("Itens do PDF adicionados!")

# ================= INSERÇÃO MANUAL DÉBITO =================

st.subheader("Inserir Débito Manual")

c1, c2, c3 = st.columns(3)

data_manual = c1.date_input("Data", key="data_debito")
desc_manual = c2.text_input("Descrição", key="desc_debito")
valor_manual = c3.number_input("Valor", step=0.01, key="valor_debito")

if st.button("Adicionar Débito", key="btn_add_debito"):

    if desc_manual and valor_manual:

        nova = pd.DataFrame([{
            "Data": formatar_data(data_manual),
            "Estabelecimento": desc_manual.strip().upper(),
            "Valor (R$)": valor_manual
        }])

        st.session_state.df_transacoes = pd.concat(
            [st.session_state.df_transacoes, nova],
            ignore_index=True
        )

        st.success("Item adicionado!")

# ================= INSERÇÃO MANUAL PIX =================

st.subheader("Inserir PIX")

c1, c2, c3 = st.columns(3)

data_pix = c1.date_input("Data PIX", key="data_pix")
fav = c2.text_input("Favorecido", key="fav_pix")
valor_pix = c3.number_input("Valor PIX", step=0.01, key="valor_pix")

if st.button("Adicionar PIX", key="btn_add_pix"):

    if fav and valor_pix:

        nova = pd.DataFrame([{
            "Data": formatar_data(data_pix),
            "Favorecido": fav.strip().upper(),
            "Valor (R$)": valor_pix
        }])

        st.session_state.df_pix = pd.concat(
            [st.session_state.df_pix, nova],
            ignore_index=True
        )

        st.success("PIX adicionado!")

# ================= LISTA DÉBITOS =================

st.subheader("Débitos")

if not st.session_state.df_transacoes.empty:

    df_deb = st.session_state.df_transacoes.copy()
    df_deb["Excluir"] = False

    edited_deb = st.data_editor(
        df_deb,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_debitos"
    )

    if st.button("Excluir Débitos Selecionados", key="btn_del_deb"):

        st.session_state.df_transacoes = edited_deb[
            edited_deb["Excluir"] == False
        ].drop(columns=["Excluir"])

        st.session_state.df_transacoes.reset_index(drop=True, inplace=True)

        st.rerun()

    st.session_state.df_transacoes = edited_deb.drop(columns=["Excluir"])
    st.session_state.df_transacoes = ordenar_por_data(st.session_state.df_transacoes)

    total = st.session_state.df_transacoes["Valor (R$)"].sum()

    st.info(f"Total Débitos: R$ {total:,.2f}")

# ================= LISTA PIX =================

st.subheader("PIX")

if not st.session_state.df_pix.empty:

    df_pix = st.session_state.df_pix.copy()
    df_pix["Excluir"] = False

    edited_pix = st.data_editor(
        df_pix,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_pix"
    )

    if st.button("Excluir PIX Selecionados", key="btn_del_pix"):

        st.session_state.df_pix = edited_pix[
            edited_pix["Excluir"] == False
        ].drop(columns=["Excluir"])

        st.session_state.df_pix.reset_index(drop=True, inplace=True)

        st.rerun()

    st.session_state.df_pix = edited_pix.drop(columns=["Excluir"])
    st.session_state.df_pix = ordenar_por_data(st.session_state.df_pix)

    total_pix = st.session_state.df_pix["Valor (R$)"].sum()

    st.info(f"Total PIX: R$ {total_pix:,.2f}")

# ================= GERAR EXCEL =================

st.subheader("Exportar")

def gerar_excel():

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        st.session_state.df_transacoes.to_excel(
            writer, index=False, sheet_name="Debitos"
        )

        st.session_state.df_pix.to_excel(
            writer, index=False, sheet_name="PIX"
        )

    return output.getvalue()

if st.button("Gerar Excel", key="btn_excel"):

    excel = gerar_excel()

    st.download_button(
        label="Baixar Excel",
        data=excel,
        file_name="fatura.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_excel"
    )

# ================= SHAREPOINT =================

st.subheader("Enviar para SharePoint")

if st.button("Enviar para SharePoint"):

    st.info("Aqui você conecta com Power Automate ou API do SharePoint.")


