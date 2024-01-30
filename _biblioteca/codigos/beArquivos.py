# Importando as bibliotecas necessárias
import locale
import pandas as pd

# Definindo a configuração regional para datas em português
locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')

# Função para abrir planilha e ler todas suas abas, retornando em lista de dict
def f_abrePlanilha(caminhoArquivo):
    # Lendo arquivo
    arquivoSelecionado = pd.ExcelFile(caminhoArquivo)
    nomesAbas = arquivoSelecionado.sheet_names

    # Lista para armazenar os dados de todas as abas
    tabelaLida = []

    # Obtendo a data atual como um objeto Timestamp
    hoje = pd.Timestamp.now().normalize()  # Normaliza para começar no início do dia

    # Percorrendo cada aba
    for indiceAbas, tabelaDaVez in enumerate(arquivoSelecionado.sheet_names):
        # Lendo a aba atual
        tabelaDaVez = arquivoSelecionado.parse(
            tabelaDaVez,

            # Pegando apenas colunas A e B
            usecols=[0, 1],

            # Nomeando abas no DataFrame
            names=['tarefas', 'datas']
        )

        # Convertendo datas de timestamp para string personalizada
        tabelaDaVez['datas'] = tabelaDaVez['datas'].apply(lambda dataDaVez: dataDaVez if not pd.isnull(dataDaVez) else None)

        # Adicionando colunas de status e cor do status
        nLinhas = len(tabelaDaVez)
        tabelaDaVez.insert(0, 'cor', ['f7f7f7'] * nLinhas)
        
        # Adicionando coluna de status com valores padrão 'OK'
        tabelaDaVez.insert(1, 'status', ['OK'] * nLinhas) 
        # Verificando se a tarefa já está atrasada ou em risco
        for i, data in enumerate(tabelaDaVez['datas']):
            if not pd.isnull(data):
                diferenca_dias = (data - hoje).days 
                if diferenca_dias < 0:
                    tabelaDaVez.at[i, 'status'] = 'ATRASADO'
                elif 0 <= diferenca_dias <= 3:
                    tabelaDaVez.at[i, 'status'] = 'RISCO'

        # Adicionando os dados da aba atual à lista
        tabelaLida.append({'gestor': nomesAbas[indiceAbas].upper(), 'dados': tabelaDaVez, 'planoDeAcao': []})

    return tabelaLida
