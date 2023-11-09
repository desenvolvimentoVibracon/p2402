# > importando bibliotecas
import pandas as pd

# -----------------------------------------
# função para abrir planilha e ler todas suas abas, retornando em lista de dict
def f_abrePlanilha(caminhoArquivo):
    # lista de abas lidas da planilha
    abasLidas = []

    # lendo arquivo
    arquivoSelecionado = pd.ExcelFile(caminhoArquivo)

    # percorrendo cada aba
    for abaDaVez in arquivoSelecionado.sheet_names:
        # pegando aba da vez
        df = arquivoSelecionado.parse(
            abaDaVez,

            # pegando apenas colunas A e B
            usecols = [0, 1],

            # nomeando abas no df
            names = ['tarefas', 'datas']
        )

        # salvando no df de acordo com títulos
        abasLidas.append(
            {
                'nomeGestor': abaDaVez.upper(),
                'tarefas': df['tarefas'].tolist(),
                'datas': df['datas'].tolist(),
            }
        )

    return abasLidas