import streamlit as st
import pandas as pd
import io
import requests

st.set_page_config(page_title="Processar CSV", page_icon="📄", layout="wide")

# ---------------------------------------------------------
# Inicializa o dataframe na sessão
# ---------------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = None

# ---------------------------------------------------------
# Função para obter token do Azure AD
# ---------------------------------------------------------
def get_token(client_id, client_secret, tenant_id):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

# ---------------------------------------------------------
# Função para enviar um item ao SharePoint
# ---------------------------------------------------------
def add_item_to_sharepoint(token, site_id, list_id, fields):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {"fields": fields}
    r = requests.post(url, json=payload, headers=headers)
    r.raise_for_status()

# ---------------------------------------------------------
# Upload do CSV
# ---------------------------------------------------------
st.header("📄 Upload do CSV")
arquivo = st.file_uploader("Envie seu CSV", type=["csv"])

if arquivo:
    df = pd.read_csv(arquivo)
    st.session_state.df = df.copy()
    st.success("CSV carregado com sucesso!")

# ---------------------------------------------------------
# Exibir tabela com opção de exclusão
# ---------------------------------------------------------
if st.session_state.df is not None:
    st.header("📌 Dados carregados")

    df = st.session_state.df

    # Criar uma coluna para excluir
    st.write("Clique para excluir uma linha:")

    # Criar botões individualmente
    for idx in df.index:
        cols = st.columns([10, 1])
        cols[0].write(df.loc[idx])

        if cols[1].button("❌", key=f"del_{idx}"):
            st.session_state.df = df.drop(idx).reset_index(drop=True)
            st.rerun()

    st.write("---")

    # ---------------------------------------------------------
    # Download do Excel atualizado
    # ---------------------------------------------------------
    st.subheader("⬇ Baixar Excel")

    output = io.BytesIO()
    st.session_state.df.to_excel(output, index=False)
    excel_bytes = output.getvalue()

    st.download_button(
        label="Baixar Excel",
        data=excel_bytes,
        file_name="dados_atualizados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write("---")

    # ---------------------------------------------------------
    # Enviar ao SharePoint
    # ---------------------------------------------------------
    st.header("📤 Enviar dados para SharePoint")

    client_id = st.text_input("Client ID")
    client_secret = st.text_input("Client Secret", type="password")
    tenant_id = st.text_input("Tenant ID")
    site_id = st.text_input("Site ID")
    list_id = st.text_input("List ID")

    if st.button("Enviar todos os dados"):
        try:
            token = get_token(client_id, client_secret, tenant_id)

            for _, row in st.session_state.df.iterrows():
                fields = row.to_dict()
                add_item_to_sharepoint(token, site_id, list_id, fields)

            st.success("Todos os dados foram enviados ao SharePoint!")

        except Exception as e:
            st.error(f"Erro ao enviar: {e}")
