from ano_main import carregar_pagina


def get_producao_mensal(mes_arquivo_diretorio):
    pagina_producao_mensal = carregar_pagina(mes_arquivo_diretorio, "Produção Mensal")
    producao_mensal = {}

    for indice, valor in enumerate(
        pagina_producao_mensal.iter_cols(
            min_row=4, max_row=4, min_col=3, values_only=True
        ),
        start=3,
    ):
        if valor[0]:
            producao_mensal[valor[0]] = pagina_producao_mensal.cell(
                row=37, column=indice
            ).value

    return producao_mensal


# print(producao_mensal)
# {
# 'Telha Cumeeira': 0,
# 'Telha Plan': 36284,
# 'Bloco de Vedação': 12200
# }

# print(producao_mensal["Telha Cumeeira"])
# 0
