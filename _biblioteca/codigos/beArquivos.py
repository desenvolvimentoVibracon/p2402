# importando bibliotecas
import locale

import pandas as pd


# -----------------------------------------------------------------------------------------------
# definições globais
# região (necessário para por datas em português)
locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')

# -----------------------------------------
# função para abrir planilha e ler todas suas abas, retornando em lista de dict
def f_abrePlanilha(caminhoArquivo):
    # lendo arquivo
    arquivoSelecionado = pd.ExcelFile(caminhoArquivo)
    nomesAbas = arquivoSelecionado.sheet_names

    # percorrendo cada aba
    tabelaLida = []
    for indiceAbas, tabelaDaVez in enumerate(arquivoSelecionado.sheet_names):
        # pegando aba da vez
        tabelaDaVez = arquivoSelecionado.parse(
            tabelaDaVez,

            # pegando apenas colunas A e B
            usecols = [0, 1],

            # nomeando abas no df
            names = ['tarefas', 'datas']
        )

        # passando datas de timestamp para string personalizada
        tabelaDaVez['datas'] = tabelaDaVez['datas'].apply(lambda dataDaVez: dataDaVez if not pd.isnull(dataDaVez) else None)

        # adicionando colunas de status e cor do status
        nLinhas = len(tabelaDaVez)
        tabelaDaVez.insert(0, 'cor', ['f7f7f7'] * nLinhas) 
        tabelaDaVez.insert(1, 'status', ['ND'] * nLinhas)

        tabelaLida.append({'gestor': nomesAbas[indiceAbas].upper(), 'dados': tabelaDaVez})

    return tabelaLida