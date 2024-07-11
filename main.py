from lib.ano_main import *
import os

pasta_dres_mensais = "arquivos/dre_mes/"
arquivos = os.listdir(pasta_dres_mensais)


def main(arquivo):
    print(f'\n----------------- [ {arquivo} ] -----------------\n')
    mes_arquivo_diretorio = f"{pasta_dres_mensais}/{arquivo}"
    substituir_valores_custo_mensal(mes_arquivo_diretorio)
    substituir_valores_mapa_de_vendas(mes_arquivo_diretorio)
    if not func_deu_erro():
        planilha_ano_write.save(caminho_arquivo_ano)
        print("\nSucesso!")
        input("\nPressione Enter para sair...\n")
        return True
    else:
        print(f'\n[LOG]: \nOcorreram alguns erros relacionados ao arquivo "{arquivo}", arrume-os para que o programa possa ser executado corretamente.')
        input("\nPressione Enter para sair...\n")
        return False


for arquivo in arquivos:
    caminho_arquivo_mes = os.path.join(pasta_dres_mensais, arquivo)
    if os.path.isfile(caminho_arquivo_mes):
        if not main(arquivo):
            break
