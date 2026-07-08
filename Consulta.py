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


@st.cache_data(ttl=300)
def carregar_e_cruzar_dados(url):
    try:
        resposta = requests.get(url)
        if resposta.status_code != 200:
            st.error(
                f"Erro ao acessar o arquivo '{NOME_ARQUIVO}' no GitHub. Verifique o nome."
            )
            return None

        conteudo = io.BytesIO(resposta.content)

        # Força a leitura das duas abas usando openpyxl
        with pd.ExcelFile(conteudo, engine="openpyxl") as xls:
            df_detail = pd.read_excel(xls, sheet_name="Detail")
            df_data = pd.read_excel(xls, sheet_name="Data")

        # Força todas as colunas para Letras Maiúsculas e sem espaços laterais
        df_detail.columns = df_detail.columns.str.strip().str.upper()
        df_data.columns = df_data.columns.str.strip().str.upper()

        # --- MAPEAMENTO DA ABA DETAIL ---
        col_orderkey_det = "ORDERKEY" if "ORDERKEY" in df_detail.columns else df_detail.columns[2]
        col_sku = "SKU" if "SKU" in df_detail.columns else df_detail.columns[3]
        col_openqty = "OPENQTY" if "OPENQTY" in df_detail.columns else df_detail.columns[10]

        # --- MAPEAMENTO DA ABA DATA ---
        col_orderkey_dat = "ORDERKEY" if "ORDERKEY" in df_data.columns else df_data.columns[2]
        
        # Procura coluna de rota pelo nome, se não achar, pega a coluna índice 11 (ROUTE)
        col_route = "ROUTE" if "ROUTE" in df_data.columns else df_data.columns[11]
        
        # Procura coluna de stop pelo nome, se não achar, pega a coluna índice 12 (STOP)
        col_stop = "STOP" if "STOP" in df_data.columns else df_data.columns[12]

        # Tratamento seguro dos dados de rota e stop
        df_data["ROTA_LIMPA"] = df_data[col_route].apply(extrair_rota_limpa)
        df_data["PEDIDO_ROTA"] = (
            df_data[col_stop].astype(str).str.replace(".0", "", regex=False)
        )

        # Filtra apenas o essencial para o cruzamento usando os nomes identificados
        df_det_res = df_detail[[col_orderkey_det, col_sku, col_openqty]].rename(
            columns={col_orderkey_det: "ORDERKEY", col_sku: "SKU", col_openqty: "OPENQTY"}
        )
        
        df_dat_res = df_data[[col_orderkey_dat, "ROTA_LIMPA", "PEDIDO_ROTA"]].rename(
            columns={col_orderkey_dat: "ORDERKEY"}
        )

        # Faz o cruzamento (Merge / PROCV) pela ORDERKEY
        df_consolidado = pd.merge(df_det_res, df_dat_res, on="ORDERKEY", how="inner")
        return df_consolidado

    except Exception as e:
        st.error(f"Erro ao cruzar as abas Detail e Data: {e}")
        return None


# Botão de atualização manual do cache
if st.sidebar.button("🔄 Atualizar Dados do GitHub"):
    st.cache_data.clear()
    st.rerun()

# Carrega a base
df_base = carregar_e_cruzar_dados(URL_RAW)

if df_base is not None:
    # Área de Filtros
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 10226403")
    with col2:
        rota_busca = st.text_input("📍 Digite a Rota:", placeholder="Ex: BR0551285")

    # Executa o filtro dinâmico
    if sku_busca or rota_busca:
        df_filtrado = df_base.copy()

        if sku_busca:
            df_filtrado = df_filtrado[
                df_filtrado["SKU"].astype(str).str.contains(sku_busca, case=False)
            ]
        if rota_busca:
            df_filtrado = df_filtrado[
                df_filtrado["ROTA_LIMPA"]
                .astype(str)
                .str.contains(rota_busca, case=False)
            ]

        if not df_filtrado.empty:
            df_exibicao = pd.DataFrame(
                {
                    "Conferido ✔": [False] * len(df_filtrado),
                    "Ordem (OrderKey)": df_filtrado["ORDERKEY"],
                    "Pedido na Rota": df_filtrado["PEDIDO_ROTA"],
                    "SKU": df_filtrado["SKU"],
                    "Quantia Solicitada": df_filtrado["OPENQTY"],
                    "Rota": df_filtrado["ROTA_LIMPA"],
                }
            )

            st.write(f"### 📋 {len(df_exibicao)} caixas encontradas:")

            st.data_editor(
                df_exibicao,
                hide_index=True,
                use_container_width=True,
                disabled=[
                    "Ordem (OrderKey)",
                    "Pedido na Rota",
                    "SKU",
                    "Quantia Solicitada",
                    "Rota",
                ],
            )
        else:
            st.warning("Nenhum registro encontrado para os filtros aplicados.")
    else:
        st.info(
            "💡 Digite um SKU ou Rota acima para listar as ordens de carregamento e quantias."
        )
else:
    st.info("Aguardando carregamento da base do GitHub...")
