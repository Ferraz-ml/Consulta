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


def limpar_chave_wms(serie):
    """
    Aplica uma limpeza profunda e agressiva para garantir que chaves de texto
    vistas como '6878076540' batam perfeitamente entre abas distintas.
    """
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )


def processar_aba_wms_coluna_c(df_bruto):
    """
    Usa a regra de que o Order Number está na Coluna C (índice 2)
    para achar a linha do cabeçalho e desduplica colunas para evitar erros de DataFrame.
    """
    linha_cabecalho_idx = 0
    
    # Procura especificamente na Coluna C (índice 2) onde está o termo 'order'
    for i in range(min(15, len(df_bruto))):
        if df_bruto.shape[1] > 2:
            celula_c = str(df_bruto.iloc[i, 2]).strip().lower()
            if "order" in celula_c:
                linha_cabecalho_idx = i
                break
                
    # Extrai os nomes das colunas daquela linha e os dados abaixo dela
    nomes_colunas = df_bruto.iloc[linha_cabecalho_idx].tolist()
    df_dados = df_bruto.iloc[linha_cabecalho_idx + 1 :].copy()
    
    # Limpa, padroniza e desduplica os nomes das colunas
    colunas_finais = []
    vistas = set()
    
    for idx, col in enumerate(nomes_colunas):
        nome_str = str(col).strip().upper()
        
        # Se for a Coluna C ou contiver ORDER e NUM, padroniza para ORDERKEY
        if idx == 2 or ("ORDER" in nome_str and "NUM" in nome_str):
            nome_base = "ORDERKEY"
        elif not nome_str or nome_str == "NAN":
            nome_base = f"COL_{idx}"
        else:
            nome_base = nome_str
            
        nome_final = nome_base
        contador = 1
        while nome_final in vistas:
            nome_final = f"{nome_base}_{contador}"
            contador += 1
            
        vistas.add(nome_final)
        colunas_finais.append(nome_final)
        
    df_dados.columns = colunas_finais
    return df_dados


@st.cache_data(ttl=60)
def carregar_dados_local():
    try:
        arquivo_excel = "Export.xlsx"
        
        # Carrega o arquivo de forma 100% bruta (sem cabeçalho inicial)
        with pd.ExcelFile(arquivo_excel, engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)

        # Processa as abas aplicando o filtro focado na Coluna C
        df_detail = processar_aba_wms_coluna_c(df_detail_cru)
        df_data = processar_aba_wms_coluna_c(df_data_cru)

        # Realiza a limpeza e padronização profunda da chave ORDERKEY
        if "ORDERKEY" in df_detail.columns:
            df_detail["ORDERKEY"] = limpar_chave_wms(df_detail["ORDERKEY"])
        if "ORDERKEY" in df_data.columns:
            df_data["ORDERKEY"] = limpar_chave_wms(df_data["ORDERKEY"])

        # Localiza dinamicamente a coluna de Rotas (ROUTE) na aba Data
        col_route = [c for c in df_data.columns if "ROUTE" in c]
        if col_route:
            df_data["ROTA_LIMPA"] = df_data[col_route[0]].apply(extrair_rota_limpa)
        else:
            df_data["ROTA_LIMPA"] = "N/A"

        # Localiza dinamicamente a coluna de Parada/Sequência (STOP) na aba Data
        col_stop = [c for c in df_data.columns if "STOP" in c]
        if col_stop:
            df_data["PEDIDO_ROTA"] = (
                df_data[col_stop[0]].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            )
        else:
            df_data["PEDIDO_ROTA"] = "N/A"

        # Localiza SKU e Quantidade na aba Detail
        col_sku = [c for c in df_detail.columns if "SKU" in c]
        col_qty = [c for c in df_detail.columns if "QTY" in c or "QUANT" in c]

        sku_lbl = col_sku[0] if col_sku else "SKU"
        qty_lbl = col_qty[0] if col_qty else "OPENQTY"

        # Isola as colunas limpas necessárias para o Procv (merge)
        df_det_res = df_detail[["ORDERKEY", sku_lbl, qty_lbl]].rename(
            columns={sku_lbl: "SKU", qty_lbl: "OPENQTY"}
        )
        df_dat_res = df_data[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # Faz o cruzamento perfeito entre Detail e Data garantindo compatibilidade de strings
        df_consolidado = pd.merge(df_det_res, df_dat_res, on="ORDERKEY", how="left")
        return df_consolidado

    except FileNotFoundError:
        st.error("Erro: O arquivo 'Export.xlsx' não foi encontrado na pasta do projeto.")
        return None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None


# Botão para limpar cache na barra lateral
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
