import io
import re
import pandas as pd
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================================
st.set_page_config(
    page_title="Consulta de Cargas", page_icon="📦", layout="wide"
)

# INJEÇÃO DE DESIGN 100% AZUL ESCURO COM BANNER 3D AZUL-CÉU
st.markdown(
    """
    <style>
    /* 1. Altera o fundo de toda a aplicação (Área principal) */
    .stApp {
        background-color: #0f172a !important; 
    }
    
    /* 2. Altera o fundo da barra lateral (Sidebar) */
    [data-testid="stSidebar"] {
        background-color: #1e293b !important; 
    }

    /* 3. Ajusta a cor padrão de todos os textos informativos */
    .stMarkdown, p, span, label, h3 {
        color: #f1f5f9 !important; 
    }

    /* 4. Estilização do Banner Centralizado com Efeito 3D e Degradê Azul-Céu */
    .custom-header {
        /* Degradê marcante que vai do azul-céu vivo ao azul escuro profundo */
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 40%, #0f172a 100%);
        padding: 35px 20px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 30px;
        
        /* Combinação de sombras internas para criar o efeito 3D (relevo e luz no topo) */
        box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.4), 
                    inset 0 10px 20px rgba(255, 255, 255, 0.1),
                    inset 0 -5px 15px rgba(0, 0, 0, 0.3),
                    0 10px 25px rgba(0,0,0,0.5);
                    
        border: 1px solid #0284c7;
        border-bottom: 5px solid #0369a1;
    }
    
    .custom-title {
        color: #ffffff !important;
        font-family: 'Helvetica Neue', Arial, sans-serif;
        font-size: 3.2rem;
        font-weight: 800;
        letter-spacing: 4px;
        margin: 0;
        text-transform: uppercase;
        text-shadow: 0px 4px 8px rgba(0,0,0,0.5);
    }
    
    .custom-subtitle {
        color: #e0f2fe !important; /* Azul muito claro para harmonizar com o azul-céu */
        font-size: 1.05rem;
        margin-top: 12px;
        margin-bottom: 0;
        font-weight: 500;
        letter-spacing: 2px;
        text-transform: uppercase;
        text-shadow: 0px 2px 4px rgba(0,0,0,0.3);
    }
    
    /* Customização dos inputs */
    div[data-baseweb="input"] {
        background-color: #1e293b !important;
        border-color: rgba(2, 132, 199, 0.4) !important;
    }
    input {
        color: #ffffff !important;
    }
    </style>
    
    <div class="custom-header">
        <h1 class="custom-title">📦 CONSULTA</h1>
        <p class="custom-subtitle">Controle de Fluxo Last-Mile & Validação de Rotas</p>
    </div>
    """,
    unsafe_allow_html=True
)

def extrair_quatro_digitos_rota(valor_rota):
    """
    Extrai apenas os 4 números finais da rota (ex: BRGP-BR BR0551285_BRGP -> 1285)
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


def limpar_serie_texto(serie):
    """
    Força a conversão para string elemento por elemento, removendo espaços
    e decimais residuais de forma totalmente segura contra erros de DataFrame.
    """
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )


@st.cache_data(ttl=60)
def carregar_dados_local():
    try:
        arquivo_excel = "Export.xlsx"
        
        # Carrega as abas de forma 100% crua (sem processar cabeçalhos automáticos)
        with pd.ExcelFile(arquivo_excel, engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)

        # ---------------------------------------------------------------------
        # 1. PROCESSANDO A ABA "DETAIL" (FOCADO EM ÍNDICES DE COLUNA)
        # ---------------------------------------------------------------------
        # Descobre onde começa o cabeçalho real
        linha_cab_det = 0
        for i in range(min(15, len(df_detail_cru))):
            if df_detail_cru.shape[1] > 2 and "order" in str(df_detail_cru.iloc[i, 2]).strip().lower():
                linha_cab_det = i
                break
        
        # Coleta os dados puramente abaixo do cabeçalho
        dados_detail = df_detail_cru.iloc[linha_cab_det + 1 :].copy()
        
        # Identifica as colunas dinamicamente com base nos nomes da linha de cabeçalho
        nomes_det = [str(x).strip().upper() for x in df_detail_cru.iloc[linha_cab_det].tolist()]
        
        idx_sku = next((i for i, col in enumerate(nomes_det) if "SKU" in col), 3) # Padrão coluna 3 se falhar
        idx_qty = next((i for i, col in enumerate(nomes_det) if "QTY" in col or "QUANT" in col), 4)

        # Monta um dataframe limpo e isolado de Detail usando índices exatos
        df_detail_limpo = pd.DataFrame()
        df_detail_limpo["ORDERKEY"] = limpar_serie_texto(dados_detail.iloc[:, 2]) # Coluna C (índice 2)
        df_detail_limpo["SKU"] = dados_detail.iloc[:, idx_sku].astype(str).str.strip()
        
        # --- ALTERAÇÃO AQUI: Conversão numérica robusta que evita o 0.00000 ---
        df_detail_limpo["OPENQTY"] = pd.to_numeric(dados_detail.iloc[:, idx_qty], errors='coerce').fillna(0)

        # ---------------------------------------------------------------------
        # 2. PROCESSANDO A ABA "DATA" (USANDO ÍNDICES BIUNÍVOCOS C, GY, GZ)
        # ---------------------------------------------------------------------
        linha_cab_dat = 0
        for i in range(min(15, len(df_data_cru))):
            if df_data_cru.shape[1] > 2 and "order" in str(df_data_cru.iloc[i, 2]).strip().lower():
                linha_cab_dat = i
                break
                
        dados_data = df_data_cru.iloc[linha_cab_dat + 1 :].copy()
        
        # Mapeamento cirúrgico de Data:
        df_data_limpo = pd.DataFrame()
        df_data_limpo["ORDERKEY"] = limpar_serie_texto(dados_data.iloc[:, 2])   # Coluna C (índice 2)
        df_data_limpo["ROTA_RAW"] = dados_data.iloc[:, 206]                     # Coluna GY (índice 206)
        df_data_limpo["PEDIDO_ROTA"] = dados_data.iloc[:, 207].astype(str).str.strip().str.replace(r"\.0$", "", regex=True) # Coluna GZ (índice 207)
        
        # Aplica a extração dos 4 dígitos na rota gerada
        df_data_limpo["ROTA_LIMPA"] = df_data_limpo["ROTA_RAW"].apply(extrair_quatro_digitos_rota)
        
        # Filtra apenas o necessário para o PROCV
        df_data_final = df_data_limpo[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # ---------------------------------------------------------------------
        # 3. MERGE (PROCV) GARANTIDO ENTRE STRINGS LIMPAS
        # ---------------------------------------------------------------------
        df_consolidado = pd.merge(df_detail_limpo, df_data_final, on="ORDERKEY", how="left")
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
