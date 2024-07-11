from lib.ano_main import carregar_pagina
from datetime import datetime

mes_atual = None
deu_erro = False


def obter_mes(data, indice):
    global deu_erro
    if isinstance(data, datetime):
        return data.month
    else:
        print(f"[ERRO]: Data inválida: {data}. Na linha {indice}.")
        deu_erro = True
        return 0


def verificar_mes_atual(data, indice):
    global mes_atual, deu_erro
    if not data:
        print(f"[ERRO]: Data ausente na linha {indice}. Informação não adicionada.")
        deu_erro = True
        return
    if indice < 7:
        if not mes_atual:
            mes_atual = obter_mes(data, indice)
        elif mes_atual != obter_mes(data, indice):
            print(mes_atual, obter_mes(data, indice))
            print(
                "[ERRO]: Os 3 primeiros meses não coincidem. Erro ao gerar o mês atual."
            )
            deu_erro = True
            return
    elif mes_atual != obter_mes(data, indice):
        print(f"[ERRO]: O mês da linha {indice} não coincide com mês atual.")
        deu_erro = True
        return


def get_custo_mensal(mes_arquivo_diretorio):
    pagina_custo_mensal = carregar_pagina(mes_arquivo_diretorio, "Custo Mensal")
    pagina_lista_geral = carregar_pagina(mes_arquivo_diretorio, 'Lista de Centro de Custo')
    custo_mensal = {}
    global mes_atual, deu_erro
    mes_atual = None
    for indice, linha in enumerate(
        pagina_custo_mensal.iter_rows(min_row=4, values_only=True), start=4
    ):
        if all(valor is None for valor in linha):
            continue

        mes_linha = linha[0]
        linha_modificada = (linha[2], linha[4], linha[5])
        verificar_mes_atual(mes_linha, indice)

        if not linha_modificada[1] and not linha_modificada[2]:
            print(
                f"[ERRO]: Valores ausentes na linha {indice}. Informação não adicionada."
            )
            deu_erro = True
            continue
        else:
            chave = linha_modificada[0]
            valor = ((linha_modificada[1] or 0), (linha_modificada[2] or 0))
            if chave in custo_mensal:
                custo_mensal[chave] = [
                    custo_mensal[chave][0] + valor[0],
                    custo_mensal[chave][1] + valor[1],
                ]
            else:
                custo_mensal[chave] = valor
                
                
    for valor in pagina_lista_geral.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
        val = valor[0]
        if val and not val in custo_mensal:
            custo_mensal[val] = [0, 0]

    return custo_mensal


def get_mes_atual():
    month_by_number = [
        "janeiro",
        "fevereiro",
        "março",
        "abril",
        "maio",
        "junho",
        "julho",
        "agosto",
        "setembro",
        "outubro",
        "novembro",
        "dezembro",
    ]
    return month_by_number[mes_atual - 1]


def func_deu_erro():
    return deu_erro
