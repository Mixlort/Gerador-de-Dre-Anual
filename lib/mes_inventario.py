from ano_main import carregar_pagina


def get_inventario(mes_arquivo_diretorio):
    pagina_inventario = carregar_pagina(mes_arquivo_diretorio, "Inventário")
    inventario = {}

    for indice, valor in enumerate(
        pagina_inventario.iter_rows(min_row=4, max_col=3, values_only=True), start=4
    ):
        if valor[0]:
            if valor[0] == "Estoque de Insumos" or valor[0] == "Produto":
                continue
            inventario[valor[0]] = (
                pagina_inventario.cell(row=indice, column=3).value or 0
            )

    return inventario
