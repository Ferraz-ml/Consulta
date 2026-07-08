import io
import re
import pandas as pd
import requests
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DO REPOSITÓRIO
# =========================================================================
USUARIO_GITHUB = "Ferraz-ml"
NOME_REPOSITORIO = "app-consulta-caixas"
NOME_ARQUIVO = "base_consulta.xlsx"

URL_RAW = f"https://raw.githubusercontent.com/{USUARIO_GITHUB}/{NOME_REPOSITORIO}/main/{NOME_ARQUIVO}"

st.set_page_config(
    page_title="Consulta de Cargas e Rotas", page_icon="📦", layout="wide"
)

st.title("📦 Consulta Rápida de Cargas por Rota e SKU")
st.markdown(
    "Busque o SKU e a rota para localizar onde o material está e realize a conferência na caixa."
)


def extrair_rota_limpa(valor_rota):
    """Extrai apenas o código da rota (ex: BRGP-BR BR0551285_BRGP -> BR0551285)"""
    if pd.isna(valor_rota):
        return "N/A"
    texto = str(valor_rota).strip()
    match = re.search(r"(BR\d+)", texto, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return texto


@st.cache_data(ttl=60)
def carregar_dados_direto(url):
    try:
        headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
        resposta = requests.get(url, headers=headers)
        
        if resposta.status_code != 200:
            st.error(f"Erro ao acessar o arquivo no GitHub. Status: {resposta.status_code}")
            return None

        conteudo = io.BytesIO(resposta.content)

        # --- ESTRATÉGIA DE LEITURA BLINDADA ---
        try:
            # 1. Tenta ler como Excel Real (aba Sheet1 ou padrão)
            df = pd.read_excel(conteudo, engine="openpyxl")
        except Exception:
            # 2. Se falhar (XLRDError), volta o ponteiro e lê como CSV (Tratando o disfarce do WMS)
            conteudo.seek(0)
            try:
                df = pd.read_csv(conteudo, sep=",", engine="python")
            except Exception:
                conteudo.seek(0)
                df = pd.read_csv(conteudo, sep=";", engine="python")

        # Força os nomes de colunas para letras maiúsculas e remove espaços
        df.columns = df.columns.str.strip().str.upper()

        # Localização dinâmica e inteligente das colunas essenciais
        col_orderkey = "ORDERKEY" if "ORDERKEY" in df.columns else df.columns[2]
        col_sku = "SKU" if "SKU" in df.columns else [c for c in df.columns if "SKU" in c][0]
        col_openqty = "OPENQTY" if "OPENQTY" in df.columns else ([c for c in df.columns if "QTY" in c][0] if [c for c in df.columns if "QTY" in c] else df.columns[10])
        col_route = "ROUTE" if "ROUTE" in df.columns else [c for c in df.columns if "ROUTE" in c][0]
        col_stop = "STOP" if "STOP" in df.columns else ([c for c in df.columns if "STOP" in c][0] if [c for c in df.columns if "STOP" in c] else None)

        # Tratamento dos dados para exibição na tela
        df["ROTA_LIMPA"] = df[col_route].apply(extrair_rota_limpa)
        
        if col_stop:
            df["PEDIDO_ROTA"] = df[col_stop].astype(str).str.replace(".0", "", regex=False)
        else:
            df["PEDIDO_ROTA"] = "1"

        # Consolidação da tabela estruturada
        df_final = pd.DataFrame({
            "ORDERKEY": df[col_orderkey],
            "SKU": df[col_sku],
            "OPENQTY": df[col_openqty],
            "ROTA_LIMPA": df["ROTA_LIMPA"],
            "PEDIDO_ROTA": df["PEDIDO_ROTA"]
        })
        
        return df_final

    except Exception as e:
        st.error(f"Erro ao processar as colunas da tabela: {e}")
        return None


# Botão de controle do cache lateral
if st.sidebar.button("🔄 Forçar Limpeza de Cache"):
    st.cache_data.clear()
    st.rerun()

df_base = carregar_dados_direto(URL_RAW)

if df_base is not None:
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 10226403")
    with col2:
        rota_busca = st.text_input("📍 Digite a Rota:", placeholder="Ex: BR0551285")

    if sku_busca or rota_busca:
        df_filtrado = df_base.copy()

        if sku_busca:
            df_filtrado = df_filtrado[df_filtrado["SKU"].astype(str).str.contains(sku_busca, case=False)]
        if rota_busca:
            df_filtrado = df_filtrado[df_filtrado["ROTA_LIMPA"].astype(str).str.contains(rota_busca, case=False)]

        if not df_filtrado.empty:
            df_exibicao = pd.DataFrame({
                "Conferido ✔": [False] * len(df_filtrado),
                "Ordem (OrderKey)": df_filtrado["ORDERKEY"],
                "Pedido na Rota": df_filtrado["PEDIDO_ROTA"],
                "SKU": df_filtrado["SKU"],
                "Quantia Solicitada": df_filtrado["OPENQTY"],
                "Rota": df_filtrado["ROTA_LIMPA"],
            })

            st.write(f"### 📋 {len(df_exibicao)} caixas encontradas:")

            st.data_editor(
                df_exibicao,
                hide_index=True,
                use_container_width=True,
                disabled=["Ordem (OrderKey)", "Pedido na Rota", "SKU", "Quantia Solicitada", "Rota"],
            )
        else:
            st.warning("Nenhum registro encontrado para os filtros aplicados.")
    else:
        st.info("💡 Digite um SKU ou Rota acima para listar as ordens de carregamento e quantias.")
else:
    st.info("Aguardando carregamento da base do GitHub...")
