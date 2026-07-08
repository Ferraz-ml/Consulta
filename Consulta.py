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


def ajustar_cabecalho_dinamico(df_cru):
    """Localiza a linha correta do cabeçalho que contém 'Order Number' e limpa o DataFrame"""
    for i in range(min(15, len(df_cru))):
        # Converte cada célula da linha para string antes de juntar tudo
        valores_linha = [str(val).strip().lower() for val in df_cru.iloc[i].tolist()]
        linha_texto = " ".join(valores_linha)

        if "order number" in linha_texto or "orderkey" in linha_texto:
            # Garante que as novas colunas sejam convertidas de forma limpa elemento por elemento
            novas_colunas = [str(col).strip().upper() for col in df_cru.iloc[i].tolist()]
            
            # Padroniza a chave se encontrar variações de "Order Number"
            novas_colunas = [
                "ORDERKEY" if "ORDER" in col and "NUM" in col else col
                for col in novas_colunas
            ]

            df_ajustado = df_cru.iloc[i + 1 :].copy()
            df_ajustado.columns = novas_colunas
            return df_ajustado

    # Fallback caso não encontre o termo-chave nas primeiras linhas
    df_cru.columns = [str(col).strip().upper() for col in df_cru.columns]
    return df_cru


@st.cache_data(ttl=60)
def carregar_dados_local():
    try:
        # Abre o Excel diretamente do repositório local sem definir cabeçalho fixo
        with pd.ExcelFile("Export.xlsx", engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)

        # Trata dinamicamente os cabeçalhos gerados pelo relatório do WMS
        df_detail = ajustar_cabecalho_dinamico(df_detail_cru)
        df_data = ajustar_cabecalho_dinamico(df_data_cru)

        # Padroniza a chave de cruzamento para texto limpo em ambas as abas
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

        # Mapeia dinamicamente as colunas de Rota e Parada por aproximação de nome
        coluna_rota = [c for c in df_data.columns if "ROUTE" in c]
        if coluna_rota:
            df_data["ROTA_LIMPA"] = df_data[coluna_rota[0]].apply(extrair_rota_limpa)
        else:
            df_data["ROTA_LIMPA"] = "N/A"

        coluna_stop = [c for c in df_data.columns if "STOP" in c]
        if coluna_stop:
            df_data["PEDIDO_ROTA"] = (
                df_data[coluna_stop[0]]
                .astype(str)
                .str.replace(".0", "", regex=False)
            )
        else:
            df_data["PEDIDO_ROTA"] = "N/A"

        # Mapeia as colunas de SKU e quantidade na aba de detalhes
        col_sku = [c for c in df_detail.columns if "SKU" in c]
        col_qty = [c for c in df_detail.columns if "QTY" in c or "QUANT" in c]

        sku_label = col_sku[0] if col_sku else "SKU"
        qty_label = col_qty[0] if col_qty else "OPENQTY"

        # Prepara os DataFrames finais para o cruzamento
        df_det_res = df_detail[["ORDERKEY", sku_label, qty_label]].rename(
            columns={sku_label: "SKU", qty_label: "OPENQTY"}
        )
        df_dat_res = df_data[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # Consolida via Left Join mantendo o foco na listagem das caixas
        df_consolidado = pd.merge(df_det_res, df_dat_res, on="ORDERKEY", how="left")
        return df_consolidado

    except FileNotFoundError:
        st.error("Erro: O arquivo 'Export.xlsx' não foi encontrado na pasta do projeto.")
        return None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None


# Controle de cache na barra lateral
if st.sidebar.button("🔄 Forçar Limpeza de Cache"):
    st.cache_data.clear()
    st.rerun()

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
                df_filtrado["ROTA_LIMPA"].astype(str).str.contains(rota_busca, case=False)
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
        st.info("💡 Digite um SKU ou Rota acima para listar as ordens de carregamento.")
else:
    st.info("Aguardando carregamento da base de dados...")
