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
NOME_ARQUIVO = "Export.xlsx"

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
def carregar_dados_direto(url):
    try:
        resposta = requests.get(url)
        if resposta.status_code != 200:
            st.error(
                f"Erro ao acessar o GitHub (Status {resposta.status_code})."
            )
            return None

        conteudo = io.BytesIO(resposta.content)

        # Lê a planilha única (sem especificar abas para evitar o erro)
        try:
            df = pd.read_excel(conteudo)
        except Exception:
            conteudo.seek(0)
            df = pd.read_csv(conteudo, sep=None, engine="python")

        # Limpa espaços em branco dos nomes das colunas
        df.columns = df.columns.str.strip()

        # Mapeamento dinâmico baseado nos nomes exatos que estão no seu WMS
        col_ordem = "ORDERKEY" if "ORDERKEY" in df.columns else "OrderKey"
        col_sku = "SKU" if "SKU" in df.columns else "Sku"
        col_qtd = "OPENQTY" if "OPENQTY" in df.columns else "OpenQty"

        # Identifica a coluna da rota (procura por colunas que contenham "ROUTE" ou assume a posição padrão)
        colunas_rota_possiveis = [c for c in df.columns if "ROUTE" in c.upper()]
        if colunas_rota_possiveis:
            col_rota = colunas_rota_possiveis[0]
        else:
            # Caso não ache por nome, tenta pegar pelo índice 206 (GY) se houver colunas suficientes, ou a última
            col_rota = (
                df.columns[206] if len(df.columns) > 206 else df.columns[-1]
            )

        # Cria a coluna de rota limpa de forma segura
        df["Rota_Limpa"] = df[col_rota].apply(extrair_rota_limpa)

        # Retorna apenas o feijão com arroz que precisamos pra tela
        df_resumido = pd.DataFrame(
            {
                "OrderKey": df[col_ordem],
                "SKU": df[col_sku],
                "OpenQty": df[col_qtd],
                "Rota_Limpa": df["Rota_Limpa"],
            }
        )

        return df_resumido

    except Exception as e:
        st.error(f"Erro ao processar os dados do WMS: {e}")
        return None


# Botão de atualização manual do cache
if st.sidebar.button("🔄 Atualizar Dados do GitHub"):
    st.cache_data.clear()
    st.rerun()

# Carrega a base limpa
df_base = carregar_dados_direto(URL_RAW)

if df_base is not None:
    # Área de Filtros
    col1, col2 = st.columns(2)
    with col1:
        sku_busca = st.text_input("🔍 Digite o SKU:", placeholder="Ex: 10226403")
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
