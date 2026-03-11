import streamlit as st
import pandas as pd
import io
import requests
import numpy as np
import chardet  # Adicionado para detecção de encoding

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
# Função para normalizar string de valor para float
# Lida com formatos comuns: "R$ 1.234,56", "-1.234,56", "(1.234,56)", "1234.56"
# ---------------------------------------------------------
def parse_val_to_float(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()

    # remove símbolo de moeda e espaços não imprimíveis
    s = s.replace("R$", "").replace("r$", "").replace("\xa0", "").strip()

    # trata valor entre parênteses como negativo: (1.234,56) -> -1.234,56
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]

    # remover sinais extras
    s = s.replace("+", "")

    # se houver tanto '.' quanto ',' assumimos que '.' é milhares e ',' decimal (formato BR)
    if "." in s and "," in s:
        s = s.replace(".", "")      # remove milhares
        s = s.replace(",", ".")     # transforma decimal para ponto
    else:
        # se só tiver vírgula, é decimal no formato BR -> trocar por ponto
        if "," in s and "." not in s:
            s = s.replace(",", ".")
        # se só tiver ponto, pode já estar no formato EN (1234.56) -> manter

    # remover espaços remanescentes
    s = s.replace(" ", "")

    # tratar casos onde resta caracteres não-numéricos (ex: texto). tenta extrair número via filtro
    # se falhar, retorna NaN
    try:
        return float(s)
    except Exception:
        # remover tudo que não seja dígito, ponto ou sinal de menos e tentar de novo
        import re
        cleaned = re.sub(r"[^0-9\.\-]", "", s)
        try:
            if cleaned == "" or cleaned == "-" or cleaned == ".":
                return np.nan
            return float(cleaned)
        except Exception:
            return np.nan

# ---------------------------------------------------------
# Função para ler CSV com detecção automática de encoding
# ---------------------------------------------------------
def read_csv_with_encoding(file):
    # Detecta o encoding do arquivo
    raw_data = file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    
    # Volta ao início do arquivo
    file.seek(0)
    
    try:
        # Tenta ler com o encoding detectado
        df = pd.read_csv(file, encoding=encoding)
        return df, encoding
    except:
        # Se falhar, tenta encodings comuns
        file.seek(0)
        encodings_to_try = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
        for enc in encodings_to_try:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc)
                return df, enc
            except:
                continue
        raise Exception("Não foi possível ler o arquivo com nenhum encoding conhecido.")

# ---------------------------------------------------------
# Upload do CSV
# ---------------------------------------------------------
st.header("📄 Upload do CSV")
arquivo = st.file_uploader("Envie seu CSV", type=["csv"])

if arquivo:
    # Usa a função com detecção de encoding
    df, encoding_usado = read_csv_with_encoding(arquivo)
    st.info(f"Encoding detectado: {encoding_usado}")

    # Verifica se existe a coluna 'valor' (case-insensitive tentativa)
    col_candidates = [c for c in df.columns if c.lower() == "valor"]
    if not col_candidates:
        st.error("A coluna 'valor' não existe no CSV! Verifique o nome exato da coluna.")
    else:
        valor_col = col_candidates[0]  # pega a coluna com nome 'valor' (mesmo se tiver VARIAÇÃO de case)
        
        # Cria uma coluna numérica com o valor convertido
        df["_valor_num"] = df[valor_col].apply(parse_val_to_float)

        # Relatório de conversão
        total_linhas = len(df)
        num_nan = df["_valor_num"].isna().sum()
        num_neg = (df["_valor_num"] < 0).sum()
        num_pos = (df["_valor_num"] > 0).sum()
        num_zero = (df["_valor_num"] == 0).sum()

        st.info(f"Linhas: {total_linhas} • Negativos: {num_neg} • Positivos: {num_pos} • Zeros: {num_zero} • Não convertidos (NaN): {num_nan}")

        # Mostrar amostra das linhas onde conversão falhou, se houver
        if num_nan > 0:
            st.warning("Algumas linhas não foram convertidas para número (NaN). Exemplo:")
            st.dataframe(df[df["_valor_num"].isna()].head(10))

        # FILTRO: manter somente negativos (onde valor numérico < 0)
        df_filtrado = df[df["_valor_num"] < 0].copy().reset_index(drop=True)

        st.session_state.df = df_filtrado.copy()
        st.success("CSV carregado com sucesso! (filtrado: apenas valores negativos)")

# ---------------------------------------------------------
# Exibir tabela com opção de exclusão (apenas Data, Valor e Descrição)
# ---------------------------------------------------------
if st.session_state.df is not None:
    st.header("📌 Dados carregados (apenas valores negativos)")
    
    # Colunas que queremos exibir
    colunas_exibir = ["Data", "Valor", "Descrição"]
    
    # Verifica se as colunas existem no DataFrame
    colunas_existentes = [col for col in colunas_exibir if col in st.session_state.df.columns]
    
    if len(colunas_existentes) < len(colunas_exibir):
        st.warning(f"Algumas colunas não foram encontradas. Encontradas: {colunas_existentes}")
    
    if not colunas_existentes:
        st.error("Nenhuma das colunas desejadas foi encontrada no DataFrame!")
    else:
        # Cria DataFrame apenas com as colunas desejadas para exibição
        df_exibicao = st.session_state.df[colunas_existentes].copy()
        
        # Criar botões individualmente para exclusão
        st.write("Clique para excluir uma linha:")
        
        # Exibe apenas as colunas selecionadas com opção de exclusão
        for idx in df_exibicao.index:
            cols = st.columns([10, 1])
            
            # Formata a linha para exibição
            linha_dict = df_exibicao.loc[idx].to_dict()
            linha_formatada = " | ".join([f"{col}: {valor}" for col, valor in linha_dict.items()])
            cols[0].write(linha_formatada)
            
            if cols[1].button("❌", key=f"del_{idx}"):
                st.session_state.df = st.session_state.df.drop(idx).reset_index(drop=True)
                st.rerun()
        
        st.write("---")
        
        # ---------------------------------------------------------
        # Download do Excel com as colunas desejadas
        # ---------------------------------------------------------
        st.subheader("⬇ Baixar Excel")
        
        # Prepara DataFrame para exportação (apenas Data, Valor e Descrição)
        export_df = st.session_state.df[colunas_existentes].copy()
        
        output = io.BytesIO()
        export_df.to_excel(output, index=False)
        excel_bytes = output.getvalue()
        
        st.download_button(
            label="Baixar Excel",
            data=excel_bytes,
            file_name="dados_atualizados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.write("---")
        
        # ---------------------------------------------------------
        # Enviar ao SharePoint com as colunas desejadas
        # ---------------------------------------------------------
        st.header("📤 Enviar dados para SharePoint")
        
        st.info(f"Serão enviadas apenas as colunas: {', '.join(colunas_existentes)}")
        
        client_id = st.text_input("Client ID")
        client_secret = st.text_input("Client Secret", type="password")
        tenant_id = st.text_input("Tenant ID")
        site_id = st.text_input("Site ID")
        list_id = st.text_input("List ID")
        
        if st.button("Enviar todos os dados"):
            try:
                token = get_token(client_id, client_secret, tenant_id)
                
                # Envia apenas as colunas desejadas
                for _, row in st.session_state.df.iterrows():
                    payload_row = row[colunas_existentes].to_dict()
                    add_item_to_sharepoint(token, site_id, list_id, payload_row)
                
                st.success("Todos os dados foram enviados ao SharePoint!")
                
            except Exception as e:
                st.error(f"Erro ao enviar: {e}")
