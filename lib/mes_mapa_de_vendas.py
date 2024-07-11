from lib.ano_main import carregar_pagina


def get_mapa_de_vendas(mes_arquivo_diretorio):
    pagina_mapa_de_vendas = carregar_pagina(mes_arquivo_diretorio, "Mapa de Vendas")
    dados_mapa_de_vendas = []
    mapa_de_vendas = {}

    for linha in pagina_mapa_de_vendas.iter_rows(
        min_row=2, min_col=3, values_only=True
    ):
        for i in range(0, len(linha) - 1, 3):
            item = {
                "quantidade": linha[i],
                "preco": linha[i + 1],
                "total": linha[i + 2],
            }
            dados_mapa_de_vendas.append(
                (pagina_mapa_de_vendas.cell(row=3, column=i + 3).value, item)
            )

    for nome, item in dados_mapa_de_vendas:
        if isinstance(item["quantidade"], (int, float)):
            if item["quantidade"] and item["preco"]:
                mapa_de_vendas[nome] = [
                    item["quantidade"],
                    round(item["preco"], 2),
                    round(item["total"], 2),
                ]
            elif item["quantidade"] and not item["preco"]:
                mapa_de_vendas[nome] = [round(item["quantidade"], 2)]
                break

    return mapa_de_vendas
