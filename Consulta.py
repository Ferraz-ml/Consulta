import pandas as pd
from nicegui import ui

# Estado global do app
state = {"df": None, "filtered_data": []}


def handle_upload(e):
    """Lê a planilha enviada pelo usuário"""
    try:
        # Lê o conteúdo do arquivo (suporta Excel)
        state["df"] = pd.read_excel(e.content)
        ui.notify("Planilha carregada com sucesso!", type="positive")
    except Exception as ex:
        ui.notify(f"Erro ao ler arquivo: {ex}", type="negative")


def filtrar_dados(sku_input, rota_input, grid):
    """Filtra a planilha por SKU e Rota e atualiza a tabela"""
    df = state["df"]
    if df is None:
        ui.notify("Por favor, carregue uma planilha primeiro.", type="warning")
        return

    sku = sku_input.value.strip() if sku_input.value else ""
    rota = rota_input.value.strip() if rota_input.value else ""

    # Ajuste os nomes das colunas ('SKU', 'Rota') conforme a sua planilha real
    query = df.copy()

    if sku:
        query = query[query["SKU"].astype(str).str.contains(sku, case=False)]
    if rota:
        query = query[query["Rota"].astype(str).str.contains(rota, case=False)]

    # Transforma em dicionário para o grid do NiceGUI e adiciona o campo do checkbox
    rows = []
    for _, row in query.iterrows():
        rows.append(
            {
                "verificado": False,
                "ordem_carregamento": row.get(
                    "Ordem", "N/A"
                ),  # Altere para o nome real da coluna
                "quantidade": row.get(
                    "Quantidade", 0
                ),  # Altere para o nome real da coluna
                "sku": row.get("SKU", ""),
                "rota": row.get("Rota", ""),
            }
        )

    grid.options["rowData"] = rows
    grid.update()
    ui.notify(f"{len(rows)} registros encontrados.", type="info")


# --- INTERFACE CORRIGIDA ---
ui.dark_mode(True)  # Mantendo o visual escuro que poupa as vistas na operação

with ui.card().classes("w-full max-w-4xl mx-auto q-pa-md mt-4"):
    ui.label("Consulta de Caixa por SKU e Rota").classes(
        "text-h5 font-bold text-primary"
    )
    ui.markdown(
        "Upload da planilha base e busca rápida de ordens de carregamento."
    )

    # 1. Área de Upload
    ui.upload(
        label="Arraste a planilha de estoque/rotas aqui",
        on_upload=handle_upload,
        auto_upload=True,
    ).classes("w-full mb-4")

    # 2. Filtros de Busca
    with ui.row().classes("w-full gap-4 items-center mb-4"):
        sku_search = ui.input(label="Buscar SKU", placeholder="Digite o SKU...")
        rota_search = ui.input(
            label="Buscar Rota", placeholder="Digite a rota..."
        )

        # Botão de gatilho
        btn_buscar = ui.button("Consultar").classes("q-px-lg")

    ui.separator().classes("my-4")

    # 3. Tabela de Resultados (AG Grid)
    # Definimos as colunas. A primeira coluna usa o 'checkboxSelection' do AG Grid
    column_defs = [
        {
            "headerName": "Ok",
            "field": "verificado",
            "checkboxSelection": True,
            "headerCheckboxSelection": True,
            "width": 90,
        },
        {"headerName": "Ordem Carregamento", "field": "ordem_carregamento"},
        {"headerName": "Qtd Solicitada", "field": "quantidade", "width": 130},
        {"headerName": "SKU", "field": "sku", "width": 150},
        {"headerName": "Rota", "field": "rota", "width": 120},
    ]

    grid_resultados = ui.aggrid(
        {
            "columnDefs": column_defs,
            "rowData": [],
            "rowSelection": "multiple",  # Permite marcar vários quadradinhos
            "animateRows": True,
        }
    ).classes("w-full h-96")

    # Vincula o clique do botão à função de filtro (passando os elementos corretos)
    btn_buscar.on(
        "click", lambda: filtrar_dados(sku_search, rota_search, grid_resultados)
    )

ui.run(title="Consulta de Cargas", port=8081)
