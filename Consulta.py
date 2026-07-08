import io
import re
import pandas as pd
import requests
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DO REPOSITÓRIO (Ajuste com seus dados do GitHub)
# =========================================================================
USUARIO_GITHUB = "Ferraz-ml"  # Seu usuário do GitHub
NOME_REPOSITORIO = "app-consulta-caixas"  # Nome do repositório atual
NOME_ARQUIVO = "Export.xlsx"  # Nome do arquivo que você subiu lá

# URL do arquivo cru (Raw) no GitHub
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


# Função com cache para não ficar baixando o arquivo pesado do GitHub a cada clique
@st.cache_data(ttl=300)  # Limpa o cache a cada 5 minutos automaticamente
def carregar_e_cruzar_dados(url):
    try:
        resposta = requests.get(url)
        if resposta.status_code != 200:
            st.error(
                f"Erro ao acessar o GitHub (Status {resposta.status_code}). Verifique se o arquivo está na raiz do repositório."
            )
            return None

        # Carrega as abas direto da memória
        conteudo = io.BytesIO(resposta.content)
        df_detail = pd.read_excel(conteudo, sheet_name="Detail")
        df_data = pd.read_excel(conteudo, sheet_name="Data")

        # Limpeza de colunas
        df_detail.columns = df_detail.columns.str.strip()
        df_data.columns = df_data.columns.str.strip()

        # Isola a rota limpa na tabela Data (Coluna GY / Índice 206)
        coluna_rota = "GY" if "GY" in df_data.columns else df_data.columns[206]
        df_data["Rota_Limpa"] = df_data[coluna_rota].apply(extrair_rota_limpa)

        # Seleciona apenas o que importa e faz o PROCV (Merge) pela OrderKey
        df_det_res = df_detail[["OrderKey", "SKU", "OpenQty"]]
        df_dat_res = df_data[["OrderKey", "Rota_Limpa"]]

        df_consolidado = pd.merge(
            df_det_res, df_dat_res, on="OrderKey", how="inner"
        )
        return df_consolidado
    except Exception as e:
        st.error(f"Erro ao processar as abas do WMS: {e}")
        return None


# Botão manual para forçar a atualização caso você mude o Excel no GitHub
if st.sidebar.button("🔄 Atualizar Dados do GitHub"):
    st.cache_data.clear()
    st.rerun()

# Carrega a base
df_base = carregar_e_cruzar_dados(URL_RAW)

if df_base is not None:
    # Área de Filtros lado a lado
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 123456")
    with col2:
        rota_busca = st.text_input(
            "📍 Digite a Rota:", placeholder="Ex: BR0551285"
        )

    # Executa o filtro dinâmico se houver digitação
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
            # Prepara a tabela estruturada
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

            # O Data Editor renderiza a caixinha interativa (Checkbox) no 'Conferido ✔'
            st.data_editor(
                df_exibicao,
                hide_index=True,
                use_container_width=True,
                disabled=[
                    "Ordem de Carregamento",
                    "SKU",
                    "Quantia Solicitada",
                    "Rota",
                ],  # Bloqueia edição do resto
            )
        else:
            st.warning("Nenhum registro encontrado para os filtros aplicados.")
    else:
        st.info(
            "💡 Digite um SKU ou Rota acima para listar as ordens de carregamento e quantias."
        )
else:
    st.info("Aguardando carregamento da base do GitHub...")
