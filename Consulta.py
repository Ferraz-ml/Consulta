import io
import re
import pandas as pd
import requests
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DO REPOSITÓRIO (Ajustado exatamente para o seu GitHub)
# =========================================================================
USUARIO_GITHUB = "Ferraz-ml"
NOME_REPOSITORIO = "app-consulta-caixas"
NOME_ARQUIVO = "Export.xlsx"  # Seu arquivo no GitHub

URL_RAW = f"https://raw.githubusercontent.com/{USUARIO_GITHUB}/{NOME_REPOSITORIO}/main/{NOME_ARQUIVO}"

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Consulta de Cargas e Rotas", page_icon="📦", layout="wide"
)

st.title("📦 Consulta Rápida de Cargas por Rota e SKU")
st.markdown(
    "Busque o SKU e a rota para localizar onde o material está e realize a conferência na caixa."
)


def extrair_rota_limpa(valor_rota):
    """Extrai apenas o código da rota (ex: BR0551285)"""
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
                f"Erro ao acessar o GitHub (Status {resposta.status_code})."
            )
            return None

        conteudo = io.BytesIO(resposta.content)

        # --- AJUSTE DE SEGURANÇA OPERACIONAL ---
        # Tenta ler como Excel padrão. Se falhar por ser um CSV renomeado, lê como CSV.
        try:
            df_detail = pd.read_excel(conteudo, sheet_name="Detail")
            df_data = pd.read_excel(conteudo, sheet_name="Data")
        except Exception:
            # Caso o arquivo seja um CSV plano
            conteudo.seek(0)
            df_geral = pd.read_csv(conteudo, sep=None, engine="python")
            # Se for uma tabela única, usamos ela mesma para mapear
            df_geral.columns = df_geral.columns.str.strip()
            df_geral["OrderKey"] = df_geral.get(
                "OrderKey", df_geral.iloc[:, 2] if len(df_geral.columns) > 2 else "N/A"
            )
            df_geral["SKU"] = df_geral.get(
                "SKU", df_geral.iloc[:, 3] if len(df_geral.columns) > 3 else "N/A"
            )
            df_geral["OpenQty"] = df_geral.get(
                "OpenQty",
                df_geral.iloc[:, 10] if len(df_geral.columns) > 10 else 0,
            )
            col_gy = (
                "GY"
                if "GY" in df_geral.columns
                else df_geral.columns[206]
                if len(df_geral.columns) > 206
                else df_geral.columns[-1]
            )
            df_geral["Rota_Limpa"] = df_geral[col_gy].apply(extrair_rota_limpa)
            return df_geral[[ "OrderKey", "SKU", "OpenQty", "Rota_Limpa" ]]

        # Limpeza de colunas das duas abas
        df_detail.columns = df_detail.columns.str.strip()
        df_data.columns = df_data.columns.str.strip()

        # Isola a rota limpa na tabela Data (Coluna GY ou índice 206)
        coluna_rota = "GY" if "GY" in df_data.columns else df_data.columns[206]
        df_data["Rota_Limpa"] = df_data[coluna_rota].apply(extrair_rota_limpa)

        # Seleciona as colunas essenciais e cruza (Merge)
        df_det_res = df_detail[["OrderKey", "SKU", "OpenQty"]]
        df_dat_res = df_data[["OrderKey", "Rota_Limpa"]]

        df_consolidado = pd.merge(
            df_det_res, df_dat_res, on="OrderKey", how="inner"
        )
        return df_consolidado
    except Exception as e:
        st.error(f"Erro ao processar os dados do WMS: {e}")
        return None


# Botão de atualização manual do cache
if st.sidebar.button("🔄 Atualizar Dados do GitHub"):
    st.cache_data.clear()
    st.rerun()

# Carrega a base filtrada
df_base = carregar_e_cruzar_dados(URL_RAW)

if df_base is not None:
    # Área de Filtros
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 123456")
    with col2:
        rota_busca = st.text_input(
            "📍 Digite a Rota:", placeholder="Ex: BR0551285"
        )

    # Executa o filtro dinâmico
    if sku_busca or rota_busca:
        df_filtrado = df_base.copy()

        if sku_busca:
            df_filtrado = df_filtrado[
                df_filtrado["SKU"].astype(str).str.contains(sku_busca, case=False)
            ]
        if rota_busca:
            df_filtrado = df_filtrado[
                df_filtrado["Rota_Limpa"]
                .astype(str)
                .str.contains(rota_busca, case=False)
            ]

        if not df_filtrado.empty:
            df_exibicao = pd.DataFrame(
                {
                    "Conferido ✔": [False] * len(df_filtrado),
                    "Ordem de Carregamento": df_filtrado["OrderKey"],
                    "SKU": df_filtrado["SKU"],
                    "Quantia Solicitada": df_filtrado["OpenQty"],
                    "Rota": df_filtrado["Rota_Limpa"],
                }
            )

            st.write(f"### 📋 {len(df_exibicao)} caixas encontradas:")

            st.data_editor(
                df_exibicao,
                hide_index=True,
                use_container_width=True,
                disabled=[
                    "Ordem de Carregamento",
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
