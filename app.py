import os
import msal
import requests

# =====================================================
# 1) DADOS DO APP / TENANT (usando variáveis de ambiente)
# =====================================================
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")  # 4039ce1c-ee58-4b19-bf86-db64da445fe1
TENANT_ID = os.getenv("AZURE_TENANT_ID")  # 02543ce8-b773-43d0-9cf1-298729881b0d
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")  # Deve ser setado no Streamlit Secrets ou env local

if not CLIENT_SECRET:
    raise Exception("CLIENT_SECRET não encontrado como variável de ambiente")

# =====================================================
# 2) SITE / LIST SHAREPOINT
# =====================================================
SITE_ID = (
    "devgbsn.sharepoint.com,"
    "351e9978-140f-427e-a87d-332f6ce67a46,"
    "fc4e159a-5954-442f-a08f-28617bc84da1"
)
LIST_ID = "b7b00e6d-9ed0-492c-958f-f80f15bd8dce"

# =====================================================
# 3) CRIAR APP MSAL COM SECRET
# =====================================================
app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

# =====================================================
# 4) OBTER TOKEN
# =====================================================
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
access_token = token_response.get("access_token")

if not access_token:
    raise Exception("❌ Não foi possível obter token: " + str(token_response))

print("✅ Token obtido com sucesso")

# =====================================================
# 5) INSERIR ITEM DE TESTE NA LISTA
# =====================================================
url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

payload = {
    "fields": {
        "Title": "Teste Python",
        "Despesa": "Total Débitos",
        "Valor": 123.45
    }
}

response = requests.post(url, headers=headers, json=payload)

if response.status_code == 201:
    print("✅ Item inserido com sucesso no SharePoint")
else:
    print("❌ Erro ao inserir item")
    print("Status:", response.status_code)
    print(response.text)

