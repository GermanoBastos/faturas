# ================== SharePoint ==================

if st.button("Enviar total para SharePoint"):

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

        if not access_token:
            raise Exception("Erro ao obter token")

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

        response = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            },
            json=payload
        )

        if response.status_code == 201:

            st.success("✅ Enviado para SharePoint com sucesso")

        else:

            st.error(f"❌ Erro SharePoint: {response.status_code} - {response.text}")

    except Exception as e:

        st.error(f"Erro na integração SharePoint: {e}")
