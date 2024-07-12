import win32com.client
from datetime import datetime
from openpyxl import Workbook

# Documento do MS Project
doc = 'p2402sw/_aux/MASTER PLAN_V.2.0_08-07-24.mpp'
# O documento deve ser atualizado semanalmente e o código tb, removido as linhas 'aprovação do cliente' e depois organizar 
# o excel, pois os projetos ficam duplicados devido ao prazo 

# Data atual como offset-naive datetime
agora = datetime.now()

# Lista de cabeçalhos principais
cabecalhosPrincipais = [
    "COORDENAÇÃO - ROGÉRIO",
    "COORDENAÇÃO - ANA",
    "COORDENAÇÃO - MAURO",
    "COORDENAÇÃO - MARCOS"
]

# Função para verificar se um nome de tarefa corresponde a um dos cabeçalhos principais
def f_retornaCabecalhoPrincipal(nomeTarefa, cabecalhos):
    return any(nomeTarefa.lower().startswith(cabecalho.lower()) for cabecalho in cabecalhos)

# Função para encontrar o pai correspondente na lista de cabeçalhos principais
def f_encontrarCabecalhoPrincipalCorrespondente(tarefa, cabecalhos):
    tarefaAtual = tarefa
    maxIteracoes = 1000  # Definindo um limite máximo de iterações (não sei pq, mas antes ficava em loop)
    iteracoes = 0
    while tarefaAtual.OutlineParent and iteracoes < maxIteracoes:
        nomePai = tarefaAtual.OutlineParent.Name
        if f_retornaCabecalhoPrincipal(nomePai, cabecalhos):
            return tarefaAtual.OutlineParent
        tarefaAtual = tarefaAtual.OutlineParent
        iteracoes += 1
    return None

# Função para encontrar a sub-tarefa correspondente em uma lista de projetos
def f_encontrarSubtarefaCorrespondente(tarefa, projetos):
    tarefaAtual = tarefa
    maxIteracoes = 1000  # Definindo um limite máximo de iterações
    iteracoes = 0
    while tarefaAtual.OutlineParent and iteracoes < maxIteracoes:
        nomePai = tarefaAtual.OutlineParent.Name
        if nomePai in projetos:
            return tarefaAtual.OutlineParent
        tarefaAtual = tarefaAtual.OutlineParent
        iteracoes += 1
    return None

try:
    mpp = win32com.client.Dispatch("MSProject.Application")
    try:
        mpp.FileOpen(doc)
        proj = mpp.ActiveProject
        tarefas = proj.Tasks

        # Lista para armazenar todos os projetos de todos os coordenadores
        todosProjetos = []

        # Iterar pelas tarefas para determinar todos os projetos de todos os coordenadores
        for tarefa in tarefas:
            if tarefa is not None:
                cabecalhoPrincipalCorrespondente = f_encontrarCabecalhoPrincipalCorrespondente(tarefa, cabecalhosPrincipais)
                if cabecalhoPrincipalCorrespondente:
                    projetos = [subsubtarefa for subtarefa in cabecalhoPrincipalCorrespondente.OutlineChildren for subsubtarefa in subtarefa.OutlineChildren]
                    todosProjetos.extend([(proj.Name, proj.Finish, proj.PercentComplete) for proj in projetos if 'HOLD' not in proj.Name.upper()])

        # Remover duplicatas da lista geral de projetos
        todosProjetos = list(set(todosProjetos))

        # Dicionário para armazenar os projetos por coordenação, contrato e status
        coordenacaoContratoProjetosStatus = {cabecalho: {'atrasado': {}, 'no prazo': {}} for cabecalho in cabecalhosPrincipais}

        # Variável para armazenar o nome da última sub-tarefa correspondente
        ultimoNomeSubtarefaCorrespondente = None

        # Iterar pelas tarefas novamente para encontrar marcos atrasados
        for tarefa in tarefas:
            if tarefa.Finish and tarefa.PercentComplete < 100:
                # Convert task.Finish to offset-naive datetime
                dataTermino = tarefa.Finish.replace(tzinfo=None)
                if dataTermino < agora:
                    if "MARCO - " in tarefa.Name.upper():  # Verifica se 'MARCO' está no nome da tarefa
                        cabecalhoPrincipalCorrespondente = f_encontrarCabecalhoPrincipalCorrespondente(tarefa, cabecalhosPrincipais)
                        if cabecalhoPrincipalCorrespondente:
                            # Criar uma lista das sub-tarefas do pai correspondente (segundo nível abaixo de cabecalhoPrincipalCorrespondente)
                            projetos = [subsubtarefa for subtarefa in cabecalhoPrincipalCorrespondente.OutlineChildren for subsubtarefa in subtarefa.OutlineChildren]
                            # Encontrar a sub-tarefa correspondente ao marco
                            subtarefaCorrespondente = f_encontrarSubtarefaCorrespondente(tarefa, [proj.Name for proj in projetos])
                            if subtarefaCorrespondente and subtarefaCorrespondente.PercentComplete < 100:
                                # Verifica se a sub-tarefa correspondente é a mesma da iteração anterior
                                if subtarefaCorrespondente.Name == ultimoNomeSubtarefaCorrespondente:
                                    continue

                                # Nome da coordenação, contrato e do projeto
                                nomeCoordenacao = cabecalhoPrincipalCorrespondente.Name
                                nomeContrato = subtarefaCorrespondente.OutlineParent.Name
                                nomeProjeto = subtarefaCorrespondente.Name
                                dataTermino = subtarefaCorrespondente.Finish
                                percentualConclusao = subtarefaCorrespondente.PercentComplete

                                # Inicializar o dicionário para o contrato se não existir
                                if nomeContrato not in coordenacaoContratoProjetosStatus[nomeCoordenacao]['atrasado']:
                                    coordenacaoContratoProjetosStatus[nomeCoordenacao]['atrasado'][nomeContrato] = []

                                # Adicionar o projeto ao contrato correspondente no dicionário
                                if (nomeProjeto, dataTermino, percentualConclusao) not in coordenacaoContratoProjetosStatus[nomeCoordenacao]['atrasado'][nomeContrato]:
                                    coordenacaoContratoProjetosStatus[nomeCoordenacao]['atrasado'][nomeContrato].append((nomeProjeto, dataTermino, percentualConclusao))

                                # Atualiza a variável com o nome da sub-tarefa correspondente atual
                                ultimoNomeSubtarefaCorrespondente = subtarefaCorrespondente.Name

        # Determinar os projetos 'no prazo' comparando com todos os projetos
        for nomeCoordenacao in coordenacaoContratoProjetosStatus.keys():
            for nomeContrato, projetos in coordenacaoContratoProjetosStatus[nomeCoordenacao]['atrasado'].items():
                for nomeProjeto, dataTermino, percentualConclusao in projetos:
                    todosProjetos = [(proj, data, pc) for proj, data, pc in todosProjetos if proj != nomeProjeto]

        for nomeProjeto, dataTermino, percentualConclusao in todosProjetos:
            for nomeCoordenacao in coordenacaoContratoProjetosStatus.keys():
                for tarefa in tarefas:
                    if tarefa is not None and tarefa.Name == nomeProjeto:
                        cabecalhoPrincipalCorrespondente = f_encontrarCabecalhoPrincipalCorrespondente(tarefa, cabecalhosPrincipais)
                        if cabecalhoPrincipalCorrespondente and cabecalhoPrincipalCorrespondente.Name == nomeCoordenacao:
                            nomeContrato = tarefa.OutlineParent.Name
                            if nomeContrato not in coordenacaoContratoProjetosStatus[nomeCoordenacao]['no prazo']:
                                coordenacaoContratoProjetosStatus[nomeCoordenacao]['no prazo'][nomeContrato] = []
                            coordenacaoContratoProjetosStatus[nomeCoordenacao]['no prazo'][nomeContrato].append((nomeProjeto, dataTermino, percentualConclusao))

        # Exibir os projetos por coordenação, contrato e status
        for coordenacao, statusDict in coordenacaoContratoProjetosStatus.items():
            print(f"Coordenação: {coordenacao}")
            for status, contratos in statusDict.items():
                print(f"  Status: {status}")
                for contrato, projetos in contratos.items():
                    print(f"    Contrato: {contrato}")
                    for projeto, dataTermino, percentualConclusao in projetos:
                        print(f"      Projeto: {projeto} - {dataTermino} - {percentualConclusao}% completo")
            print()

        # Criação de arquivos Excel separados para cada coordenação
        for coordenacao, statusDict in coordenacaoContratoProjetosStatus.items():
            wb = Workbook()
            ws = wb.active
            ws.title = coordenacao
            ws.append(['Projeto/Contrato', 'Prazo', '% Conclusão', 'Status'])
            for status, contratos in statusDict.items():
                for contrato, projetos in contratos.items():
                    ws.append([contrato, '', '', status])
                    for projeto, dataTermino, percentualConclusao in projetos:
                        ws.append([f'  {projeto}', dataTermino.strftime('%d/%m/%Y'), percentualConclusao, status])

            # Salvar o arquivo Excel
            nomeArquivoExcel = f'projetos{coordenacao}.xlsx'
            wb.save(nomeArquivoExcel)

            nomeCoordenacao = coordenacao.split(" - ")[1]
            nomeArquivoExcel = f'projetos{nomeCoordenacao}.xlsx'
            wb.save(nomeArquivoExcel)

    except Exception as e:
        print(f"Erro ao abrir o arquivo MPP: {e}")
    finally:
        mpp.FileClose()
except Exception as e:
    print(f"Erro ao iniciar o MS Project: {e}")
