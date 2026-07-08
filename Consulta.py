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


def extrair_quatro_digitos_rota(valor_rota):
    """
    Extrai apenas os 4 números finais da rota (ex: BRGP-BR BR0551285_BRGP -> 1285)
    Busca o padrão BR055 + 4 dígitos e captura apenas os 4 dígitos finais.
    """
    if pd.isna(valor_rota):
        return "N/A"
    texto = str(valor_rota).strip()
    # Procura por BR seguido de números e captura os 4 últimos antes do underline ou espaço
    match = re.search(r"BR\d{3}(\d{4})", texto, re.IGNORECASE)
    if match:
        return match.group(1)
    
    # Caso o padrão mude um pouco, tenta pegar qualquer sequência de 4 dígitos isolados
    match_fallback = re.search(r"(\d{4})", texto)
    if match_fallback:
        return match_fallback.group(1)
        
    return texto


def limpar_chave_wms(serie):
    """Transforma a chave em texto puro, remove decimais e espaços zerados."""
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )


def processar_aba_detail(df_bruto):
    """Processa a aba Detail: Order Number está na Coluna C (índice 2)"""
    linha_cabecalho_idx = 0
    for i in range(min(15, len(df_bruto))):
        if df_bruto.shape[1] > 2:
            if "order" in str(df_bruto.iloc[i, 2]).strip().lower():
                linha_cabecalho_idx = i
                break
                
    nomes_colunas = df_bruto.iloc[linha_cabecalho_idx].tolist()
    df_dados = df_bruto.iloc[linha_cabecalho_idx + 1 :].copy()
    
    colunas_finais = []
    for idx, col in enumerate(nomes_colunas):
        nome_str = str(col).strip().upper()
        if idx == 2 or ("ORDER" in nome_str and "NUM" in nome_str):
            colunas_finais.append("ORDERKEY")
        else:
            colunas_finais.append(nome_str if nome_str and nome_str != "NAN" else f"COL_{idx}")
            
    df_dados.columns = colunas_finais
    return df_dados


def processar_aba_data_mapeada(df_bruto):
    """
    Processa a aba Data indo direto nas colunas físicas passadas:
    Coluna GY (índice 206) -> Rota Completa (EXT_UDF_STR3)
    Coluna GZ (índice 207) -> Número do Pedido/Ordem da Rota (EXT_UDF_STR4)
    Coluna C (índice 2)   -> Chave OrderKey original para o Merge
    """
    # Encontra a linha do cabeçalho real
    linha_cabecalho_idx = 0
    for i in range(min(15, len(df_bruto))):
        if df_bruto.shape[1] > 2:
            if "order" in str(df_bruto.iloc[i, 2]).strip().lower():
                linha_cabecalho_idx = i
                break
                
    # Recorta os dados
    df_dados = df_bruto.iloc[linha_cabecalho_idx + 1 :].copy()
    
    # Criamos um DataFrame estruturado mapeando diretamente os índices das colunas no Excel
    # Coluna C = índice 2 | Coluna GY = índice 206 | Coluna GZ = índice 207
    df_mapeado = pd.DataFrame()
    
    if df_dados.shape[1] > 2:
        df_mapeado["ORDERKEY"] = limpar_chave_wms(df_dados.iloc[:, 2])
    else:
        df_mapeado["ORDERKEY"] = "N/A"
        
    if df_dados.shape[1] > 206:
        df_mapeado["ROTA_RAW"] = df_dados.iloc[:, 206]
        df_mapeado["ROTA_LIMPA"] = df_mapeado["ROTA_RAW"].apply(extrair_quatro_digitos_rota)
    else:
        df_mapeado["ROTA_LIMPA"] = "N/A"
        
    if df_dados.shape[1] > 207:
        df_mapeado["PEDIDO_ROTA"] = df_dados.iloc[:, 207].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    else:
        df_mapeado["PEDIDO_ROTA"] = "N/A"
        
    return df_mapeado[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]


@st.cache_data(ttl=60)
def carregar_dados_local():
    try:
        arquivo_excel = "Export.xlsx"
        
        with pd.ExcelFile(arquivo_excel, engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)

        # Processa a aba Detail normalmente buscando o SKU e Qtd
        df_detail = processar_aba_detail(df_detail_cru)
        if "ORDERKEY" in df_detail.columns:
            df_detail["ORDERKEY"] = limpar_chave_wms(df_detail["ORDERKEY"])

        # Processa a aba Data usando o mapeamento exato das colunas GY e GZ
        df_data_limpo = processar_aba_data_mapeada(df_data_cru)

        # Localiza colunas de SKU e Quantidade na aba Detail
        col_sku = [c for c in df_detail.columns if "SKU" in c]
        col_qty = [c for c in df_detail.columns if "QTY" in c or "QUANT" in c]

        sku_lbl = col_sku[0] if col_sku else "SKU"
        qty_lbl = col_qty[0] if col_qty else "OPENQTY"

        df_det_res = df_detail[["ORDERKEY", sku_lbl, qty_lbl]].rename(
            columns={sku_lbl: "SKU", qty_lbl: "OPENQTY"}
        )

        # Faz o cruzamento usando as chaves limpas como Texto puro
        df_consolidado = pd.merge(df_det_res, df_data_limpo, on="ORDERKEY", how="left")
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
        rota_busca = st.text_input("📍 Digite a Rota (4 dígitos):", placeholder="Ex: 1285")

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
