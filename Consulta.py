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


def processar_aba_wms(df_bruto):
    """
    Varre o DataFrame bruto para encontrar a linha do cabeçalho real.
    Garante o retorno de um DataFrame com colunas limpas e padronizadas,
    sem usar métodos que gerem conflitos de tipo ('str' ou 'DataFrame').
    """
    # Converte tudo para string de forma segura para análise de metadados
    df_strings = df_bruto.astype(str).copy()
    
    linha_cabecalho_idx = None
    
    # Procura qual linha contém as palavras-chave do WMS
    for i in range(min(20, len(df_strings))):
        valores_linha = df_strings.iloc[i].tolist()
        linha_texto_unida = " ".join(valores_linha).lower()
        
        if "order number" in linha_texto_unida or "orderkey" in linha_texto_unida:
            linha_cabecalho_idx = i
            break
            
    if linha_cabecalho_idx is not None:
        # Extrai os nomes das colunas diretamente dessa linha
        nomes_colunas = df_bruto.iloc[linha_cabecalho_idx].tolist()
        # Corta o dataframe para conter apenas os dados abaixo do cabeçalho
        df_dados = df_bruto.iloc[linha_cabecalho_idx + 1 :].copy()
    else:
        # Se não achar nada, assume a primeira linha como cabeçalho provisório
        nomes_colunas = df_bruto.iloc[0].tolist()
        df_dados = df_bruto.iloc[1:].copy()

    # Limpa e padroniza os nomes das colunas (reovendo espaços e em maiúsculo)
    colunas_limpas = []
    for col in nomes_colunas:
        nome_str = str(col).strip().upper()
        # Padroniza variações da chave principal do WMS
        if "ORDER" in nome_str and "NUM" in nome_str:
            colunas_limpas.append("ORDERKEY")
        else:
            colunas_limpas.append(nome_str)
            
    df_dados.columns = colunas_limpas
    return df_dados


@st.cache_data(ttl=60)
def carregar_dados_local():
    try:
        arquivo_excel = "Export.xlsx"
        
        # Lê as abas sem assumir nenhuma linha fixa como cabeçalho (header=None)
        with pd.ExcelFile(arquivo_excel, engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)

        # Processa dinamicamente a estrutura de cada aba do relatório
        df_detail = processar_aba_wms(df_detail_cru)
        df_data = processar_aba_wms(df_data_cru)

        # Garante que a chave ORDERKEY seja tratada estritamente como texto limpo
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

        # Identifica e trata a coluna de Rotas (Procura por ROUTE)
        col_route = [c for c in df_data.columns if "ROUTE" in c]
        if col_route:
            df_data["ROTA_LIMPA"] = df_data[col_route[0]].apply(extrair_rota_limpa)
        else:
            df_data["ROTA_LIMPA"] = "N/A"

        # Identifica e trata a coluna de Sequência/Pedido na rota (Procura por STOP)
        col_stop = [c for c in df_data.columns if "STOP" in c]
        if col_stop:
            df_data["PEDIDO_ROTA"] = (
                df_data[col_stop[0]].astype(str).str.replace(".0", "", regex=False)
            )
        else:
            df_data["PEDIDO_ROTA"] = "N/A"

        # Identifica dinamicamente as colunas de SKU e quantidade na aba Detail
        col_sku = [c for c in df_detail.columns if "SKU" in c]
        col_qty = [c for c in df_detail.columns if "QTY" in c or "QUANT" in c]

        sku_lbl = col_sku[0] if col_sku else "SKU"
        qty_lbl = col_qty[0] if col_qty else "OPENQTY"

        # Garante que as colunas críticas existem e as isola
        df_det_res = df_detail[["ORDERKEY", sku_lbl, qty_lbl]].rename(
            columns={sku_lbl: "SKU", qty_lbl: "OPENQTY"}
        )
        df_dat_res = df_data[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # Faz o cruzamento das duas tabelas através da chave ORDERKEY (Left Join)
        df_consolidado = pd.merge(df_det_res, df_dat_res, on="ORDERKEY", how="left")
        return df_consolidado

    except FileNotFoundError:
        st.error("Erro: O arquivo 'Export.xlsx' não foi encontrado na pasta do projeto.")
        return None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None


# Botão de controle de cache na barra lateral
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
