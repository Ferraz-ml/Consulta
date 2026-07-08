import re
import pandas as pd
from nicegui import ui

# Estado global para armazenar o DataFrame unificado
state = {"df_consolidado": None}


def extrair_rota_limpa(valor_rota):
    """
    Remove os prefixos e sufixos da rota.
    Exemplo: 'BRGP-BR BR0551285_BRGP' -> 'BR0551285'
    """
    if pd.isna(valor_rota):
        return "N/A"

    texto = str(valor_rota).strip()
    # Expressão regular para capturar o padrão BR seguido de números (ex: BR0551285)
    match = re.search(r"(BR\d+)", texto, re.IGNORECASE)
    if match:
        return match.group(1).upper()

    return texto  # Retorna o original caso não encontre o padrão


def handle_upload(e):
    """Processa o upload do Excel e faz o cruzamento das abas Data e Detail"""
    try:
        # Lendo o arquivo Excel diretamente da memória (com as duas abas)
        conteudo_arquivo = e.content

        ui.notify(
            "Processando planilhas... Aguarde.", type="ongoing", duration=2
        )

        # 1. Carrega as duas abas necessárias
        df_detail = pd.read_excel(conteudo_arquivo, sheet_name="Detail")
        df_data = pd.read_excel(conteudo_arquivo, sheet_name="Data")

        # 2. Garante que os nomes das colunas estão limpos (sem espaços extras)
        df_detail.columns = df_detail.columns.str.strip()
        df_data.columns = df_data.columns.str.strip()

        # 3. Limpa e isola a rota na tabela Data (Coluna GY)
        # Se a coluna se chamar exatamente 'GY', usamos ela, senão tratamos pelo índice ou nome real
        coluna_rota_original = (
            "GY" if "GY" in df_data.columns else df_data.columns[206]
        )  # Ajuste preventivo

        df_data["Rota_Limpa"] = df_data[coluna_rota_original].apply(
            extrair_rota_limpa
        )

        # 4. Seleciona apenas as colunas que importam para o cruzamento
        df_detail_resumido = df_detail[["OrderKey", "SKU", "OpenQty"]]
        df_data_resumido = df_data[["OrderKey", "Rota_Limpa"]]

        # 5. Faz o "PROCV" (Merge) unindo as duas tabelas pela chave OrderKey
        df_final = pd.merge(
            df_detail_resumido, df_data_resumido, on="OrderKey", how="inner"
        )

        # Armazena o resultado final no estado do app
        state["df_consolidado"] = df_final
        ui.notify(
            "Base WMS (Data + Detail) consolidada com sucesso!", type="positive"
        )

    except Exception as ex:
        ui.notify(
            f"Erro ao processar as abas do Excel: {ex}",
            type="negative",
            duration=5,
        )


def realizar_consulta(sku_input, rota_input, grid):
    """Filtra os dados consolidados e joga no AG Grid"""
    df = state["df_consolidado"]
    if df is None:
        ui.notify(
            "Por favor, carregue o arquivo Excel com as abas Data e Detail primeiro.",
            type="warning",
        )
        return

    sku_alvo = sku_input.value.strip() if sku_input.value else ""
    rota_alvo = rota_input.value.strip() if rota_input.value else ""

    if not sku_alvo and not rota_alvo:
        ui.notify(
            "Insira um SKU ou uma Rota para pesquisar.", type="warning"
        )
        return

    dados_filtrados = df.copy()

    # Aplica os filtros digitados pelo usuário
    if sku_alvo:
        dados_filtrados = dados_filtrados[
            dados_filtrados["SKU"].astype(str).str.contains(sku_alvo, case=False)
        ]
    if rota_alvo:
        dados_filtrados = dados_filtrados[
            dados_filtrados["Rota_Limpa"]
            .astype(str)
            .str.contains(rota_alvo, case=False)
        ]

    # Monta as linhas para a tabela visual
    linhas_tabela = []
    for _, row in dados_filtrados.iterrows():
        linhas_tabela.append(
            {
                "verificado": False,
                "ordem": row.get("OrderKey", "N/A"),
                "sku": row.get("SKU", "N/A"),
                "quantidade": row.get("OpenQty", 0),
                "rota": row.get("Rota_Limpa", "N/A"),
            }
        )

    # Atualiza o componente na tela
    grid.options["rowData"] = linhas_tabela
    grid.update()
    ui.notify(f"{len(linhas_tabela)} registros encontrados.", type="info")


# --- INTERFACE GRÁFICA (NiceGUI) ---
ui.dark_mode(True)

with ui.card().classes("w-full max-w-4xl mx-auto q-pa-md mt-4"):
    ui.label("Consulta de Cargas por Rota e SKU").classes(
        "text-h5 font-bold text-primary"
    )
    ui.markdown(
        "Carregue a planilha de exportação do WMS. O sistema fará o cruzamento automático entre as abas **Data** e **Detail**."
    )

    # Área de Upload
    ui.upload(
        label="Arraste o arquivo Excel completo aqui",
        on_upload=handle_upload,
        auto_upload=True,
    ).classes("w-full mb-4")

    # Filtros de Busca
    with ui.row().classes("w-full gap-4 items-center mb-4"):
        sku_field = ui.input(
            label="SKU (Coluna D)", placeholder="Digite o SKU..."
        ).classes("flex-1")
        rota_field = ui.input(
            label="Rota (Filtra pelo código limpo)",
            placeholder="Ex: BR0551285...",
        ).classes("flex-1")
        btn_consultar = ui.button("Consultar", icon="search").classes(
            "q-px-lg h-14"
        )

    ui.separator().classes("my-2")

    # Definição das colunas com o quadradinho de check
    defs_colunas = [
        {
            "headerName": "Ok",
            "field": "verificado",
            "checkboxSelection": True,
            "headerCheckboxSelection": True,
            "width": 80,
        },
        {"headerName": "Ordem (OrderKey)", "field": "ordem", "flex": 1},
        {"headerName": "SKU", "field": "sku", "width": 150},
        {"headerName": "Quantidade (OpenQty)", "field": "quantidade", "width": 160},
        {"headerName": "Rota (Limpa)", "field": "rota", "width": 150},
    ]

    grid_resultados = ui.aggrid(
        {
            "columnDefs": defs_colunas,
            "rowData": [],
            "rowSelection": "multiple",
            "animateRows": True,
        }
    ).classes("w-full h-96")

    btn_consultar.on(
        "click",
        lambda: realizar_consulta(sku_field, rota_field, grid_resultados),
    )

ui.run(title="Consulta e Picking WMS", port=8083)
