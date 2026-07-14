import io
import re
import pandas as pd
import streamlit as st

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================================
st.set_page_config(
    page_title="Consulta de Cargas", 
    page_icon="📦", 
    layout="wide"
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

def processar_dataframes(df_detail_cru, df_data_cru):
    try:
        # 1. PROCESSANDO A ABA "DETAIL" (Mapeamento na Coluna L fixa)
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
        df_detail_limpo["OPENQTY"] = pd.to_numeric(dados_detail.iloc[:, idx_qty], errors="coerce").fillna(0)

        # 2. PROCESSANDO A ABA "DATA" (Usando índices de coluna fixas)
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

        # 3. MERGE (PROCV)
        df_consolidado = pd.merge(df_detail_limpo, df_data_final, on="ORDERKEY", how="left")
        return df_consolidado
    except Exception as e:
        st.error(f"Erro ao processar as colunas do arquivo: {e}")
        return None

def ler_arquivo_excel(caminho_ou_buffer):
    try:
        # Tentativa 1: Excel padrão .xlsx
        with pd.ExcelFile(caminho_ou_buffer, engine="openpyxl") as xls:
            df_detail = pd.read_excel(xls, sheet_name="Detail", header=None)
            df_data = pd.read_excel(xls, sheet_name="Data", header=None)
            return df_detail, df_data
    except Exception as e1:
        erro_str = str(e1).lower()
        if "zip" in erro_str or "unsupported" in erro_str or "format" in erro_str or "bad" in erro_str:
            try:
                # Tentativa 2: HTML/XML disfarçado
                df_detail = pd.read_html(caminho_ou_buffer, match="Detail")[0]
                df_data = pd.read_html(caminho_ou_buffer, match="Data")[0]
                return df_detail, df_data
            except Exception:
                try:
                    # Tentativa 3: XLS antigo
                    df_detail = pd.read_excel(caminho_ou_buffer, sheet_name="Detail", header=None, engine="xlrd")
                    df_data = pd.read_excel(caminho_ou_buffer, sheet_name="Data", header=None, engine="xlrd")
                    return df_detail, df_data
                except Exception:
                    pass
    return None, None

# --- CARREGAMENTO INTEGRALMENTE MANUAL ---
df_base = None

st.markdown("### 📂 Envie seu Relatório de Carga")
arquivo_enviado = st.file_uploader(
    "Arraste o arquivo exportado do ERP aqui (.xlsx ou .xls):", 
    type=["xlsx", "xls"]
)

if arquivo_enviado is not None:
    # Mostra o tamanho real do arquivo que você acabou de subir para provar que o upload deu certo!
    tamanho_kb = len(arquivo_enviado.getvalue()) / 1024
    st.info(f"📁 Arquivo recebido com sucesso! Tamanho carregado na memória: {tamanho_kb:.2f} KB")
    
    df_det, df_dat = ler_arquivo_excel(arquivo_enviado)
    if df_det is not None and df_dat is not None:
        df_base = processar_dataframes(df_det, df_dat)
        if df_base is not None:
            st.success("✅ Banco de dados processado com sucesso!")
    else:
        st.error("❌ O arquivo enviado não pôde ser lido. Certifique-se de que ele possui as abas 'Detail' e 'Data'.")
else:
    st.warning("⚠️ Aguardando você arrastar o arquivo de cargas no campo acima para iniciar.")

# --- INTERFACE DE BUSCA ---
if df_base is not None:
    st.write("---")
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
