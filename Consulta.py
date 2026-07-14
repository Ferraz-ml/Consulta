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
    .stApp { background-color: #0f172a !important; }
    [data-testid="stSidebar"] { background-color: #1e293b !important; }
    .stMarkdown, p, span, label, h3 { color: #f1f5f9 !important; }
    .custom-header {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 40%, #0f172a 100%);
        padding: 35px 20px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 30px;
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
        color: #e0f2fe !important;
        font-size: 1.05rem;
        margin-top: 12px;
        margin-bottom: 0;
        font-weight: 500;
        letter-spacing: 2px;
        text-transform: uppercase;
        text-shadow: 0px 2px 4px rgba(0,0,0,0.3);
    }
    div[data-baseweb="input"] {
        background-color: #1e293b !important;
        border-color: rgba(2, 132, 199, 0.4) !important;
    }
    input { color: #ffffff !important; }
    </style>
    
    <div class="custom-header">
        <h1 class="custom-title">📦 CONSULTA</h1>
        <p class="custom-subtitle">Controle de Fluxo Last-Mile & Validação de Rotas</p>
    </div>
    """,
    unsafe_allow_html=True
)

def extrair_quatro_digitos_rota(valor_rota):
    if pd.isna(valor_rota):
        return "N/A"
    texto = str(valor_rota).strip()
    match = re.search(r"BR\d{3}(\d{4})", texto, re.IGNORECASE)
    if match:
        return match.group(1)
    match_fallback = re.search(r"(\d{4})", texto)
    if match_fallback:
        return match_fallback.group(1)
    return texto

def limpar_serie_texto(serie):
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )

@st.cache_data(ttl=60)
def carregar_dados_local():
    arquivo_excel = "Export.xlsx"
    
    df_detail_cru = None
    df_data_cru = None
    
    try:
        # TENTATIVA 1: Tenta ler o formato padrão moderno (.xlsx)
        with pd.ExcelFile(arquivo_excel, engine="openpyxl") as xls:
            df_detail_cru = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data_cru = pd.read_excel(xls, sheet_name="Data", header=None)
    except Exception as e:
        erro_str = str(e).lower()
        
        # TENTATIVA 2: Se for um ZIP inválido ou formato não suportado, pode ser HTML ou XML
        if "zip" in erro_str or "unsupported" in erro_str or "format" in erro_str or "bad" in erro_str:
            try:
                # Força a leitura interpretando a estrutura de tabelas HTML (comum em exports de ERPs)
                df_detail_cru = pd.read_html(arquivo_excel, match="Detail")[0]
                df_data_cru = pd.read_html(arquivo_excel, match="Data")[0]
            except Exception:
                try:
                    # TENTATIVA 3: Tenta o motor antigo xlrd (arquivos .xls reais de 97-2003)
                    df_detail_cru = pd.read_excel(arquivo_excel, sheet_name="Detail", header=None, engine="xlrd")
                    df_data_cru = pd.read_excel(arquivo_excel, sheet_name="Data", header=None, engine="xlrd")
                except Exception:
                    try:
                        # TENTATIVA 4: E se for apenas um arquivo CSV disfarçado com tabulação ou vírgula?
                        df_completo = pd.read_csv(arquivo_excel, sep=None, engine="python", header=None)
                        df_detail_cru = df_completo
                        df_data_cru = df_completo
                        st.warning("⚠️ O arquivo foi detectado como CSV/Texto. O mapeamento de abas 'Detail'/'Data' pode não funcionar como esperado.")
                    except Exception as e_final:
                        st.error(f"Não foi possível decodificar o arquivo. Formato totalmente incompatível. Erro interno: {e_final}")
                        return None
        else:
            st.error(f"Erro ao abrir arquivo: {e}")
            return None

    # Validação de segurança se os dataframes foram carregados de forma bem-sucedida
    if df_detail_cru is None or df_data_cru is None:
        st.error("Falha crítica: Os dados de origem não puderam ser carregados em memória.")
        return None

    try:
        # ---------------------------------------------------------------------
        # 1. PROCESSANDO A ABA "DETAIL" (MAPEAMENTO FIXO COLUNA L)
        # ---------------------------------------------------------------------
        linha_cab_det = 0
        for i in range(min(15, len(df_detail_cru))):
            if df_detail_cru.shape[1] > 2 and "order" in str(df_detail_cru.iloc[i, 2]).strip().lower():
                linha_cab_det = i
                break
        
        dados_detail = df_detail_cru.iloc[linha_cab_det + 1 :].copy()
        nomes_det = [str(x).strip().upper() for x in df_detail_cru.iloc[linha_cab_det].tolist()]
        idx_sku = next((i for i, col in enumerate(nomes_det) if "SKU" in col), 3) 

        # Coluna L fixa (Índice 11)
        idx_qty = 11 

        df_detail_limpo = pd.DataFrame()
        df_detail_limpo["ORDERKEY"] = limpar_serie_texto(dados_detail.iloc[:, 2])
        df_detail_limpo["SKU"] = dados_detail.iloc[:, idx_sku].astype(str).str.strip()
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
        
        df_data_limpo = pd.DataFrame()
        df_data_limpo["ORDERKEY"] = limpar_serie_texto(dados_data.iloc[:, 2])
        df_data_limpo["ROTA_RAW"] = dados_data.iloc[:, 206]
        df_data_limpo["PEDIDO_ROTA"] = dados_data.iloc[:, 207].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        
        df_data_limpo["ROTA_LIMPA"] = df_data_limpo["ROTA_RAW"].apply(extrair_quatro_digitos_rota)
        df_data_final = df_data_limpo[["ORDERKEY", "ROTA_LIMPA", "PEDIDO_ROTA"]]

        # ---------------------------------------------------------------------
        # 3. MERGE (PROCV)
        # ---------------------------------------------------------------------
        df_consolidado = pd.merge(df_detail_limpo, df_data_final, on="ORDERKEY", how="left")
        return df_consolidado

    except Exception as e_proc:
        st.error(f"Erro ao processar as colunas do arquivo de hoje: {e_proc}")
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
