import io
import re
import pandas as pd
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================================
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
def carregar_dados_local():
    try:
        # Abre o Excel diretamente do repositório local
        with pd.ExcelFile("Export.xlsx", engine="openpyxl") as xls:
            df_detail = pd.read_excel(xls, sheet_name="Detail")
            df_data = pd.read_excel(xls, sheet_name="Data")

        # Limpa e padroniza os nomes das colunas para maiúsculo
        df_detail.columns = df_detail.columns.str.strip().str.upper()
        df_data.columns = df_data.columns.str.strip().str.upper()

        # =========================================================================
        # CORREÇÃO CRUCIAL: Força a chave ORDERKEY a ser String limpa nas duas abas
        # =========================================================================
        if "ORDERKEY" in df_detail.columns:
            df_detail["ORDERKEY"] = (
                df_detail["ORDERKEY"]
                .astype(str)
                .str.strip()
                .str.replace(".0", "", regex=False)
            )
        if "ORDERKEY" in df_data.columns:
            df_data["ORDERKEY"] = (
                df_data["ORDERKEY"]
                .astype(str)
                .str.strip()
                .str.replace(".0", "", regex=False)
            )

        # Tratamento das colunas de Rota e Parada (STOP)
        if "ROUTE" in df_data.columns:
            df_data["ROTA_LIMPA"] = df_data["ROUTE"].apply(extrair_rota_limpa)
        else:
            df_data["ROTA_LIMPA"] = "N/A"

        if "STOP" in df_data.columns:
            df_data["PEDIDO_ROTA"] = (
                df_data["STOP"].astype(str).str.replace(".0", "", regex=False)
            )
        else:
            df_data["PEDIDO_ROTA"] = "N/A"

        # Filtrando colunas essenciais de cada aba
        df_det_res = df_detail[["ORDERKEY", "SKU", "OPENQTY"]]
        df_dat_res = df_data[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # Faz o cruzamento (Merge) mapeando o Detail como tabela principal (left join)
        df_consolidado = pd.merge(
            df_det_res, df_dat_res, on="ORDERKEY", how="left"
        )
        return df_consolidado

    except FileNotFoundError:
        st.error(
            "Erro: O arquivo 'Export.xlsx' não foi encontrado na pasta do projeto. "
            "Garanta que ele foi enviado para o GitHub junto com este script."
        )
        return None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None


# Botão de controle do cache lateral
if st.sidebar.button("🔄 Forçar Limpeza de Cache"):
    st.cache_data.clear()
    st.rerun()

# Carrega a base localmente
df_base = carregar_dados_local()

if df_base is not None:
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 10226403")
    with col2:
        rota_busca = st.text_input("📍 Digite a Rota:", placeholder="Ex: BR0551285")

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
    st.info("Aguardando carregamento da base de dados...")
