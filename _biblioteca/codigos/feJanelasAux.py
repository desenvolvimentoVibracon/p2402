
# importando bibliotecas
from datetime import datetime

import numpy as np

import plotly.offline as pyo
import plotly.graph_objs as go
from plotly.subplots import make_subplots

import pandas as pd

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import(QHBoxLayout, QVBoxLayout, QDialog, QLabel, QDialogButtonBox, QComboBox, QComboBox)

# -----------------------------------------------------------------------------------------------
# classe para criação de janela de 
class JanelaSelecionaGestor(QDialog):
    def __init__(self, parent = None):
        super().__init__(parent)
        # propriedades da janela parent
        self.janelaReferencia = parent

        # inicializando gui
        self.f_inicializaGui(parent)

    # -----------------------------------------
    # função para inicializar interface
    def f_inicializaGui(self, parent):
        
        # título da janela
        self.setWindowTitle('LIDERANÇAS')

        # ocultando botão default de ajuda
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        # criação de layout
        layout = QVBoxLayout()
        linha = QHBoxLayout()

        coluna = QVBoxLayout()
        self.entradaGestorDaVez = self.f_geraEntradaListaSuspensaDistribuicoes(coluna, '→ LIDERANÇA', parent.propriedadesGerais)
        self.entradaGestorDaVez.setCurrentText(parent.propriedadesGerais['gestorDaVez'])
        coluna.addWidget(self.entradaGestorDaVez)
        linha.addLayout(coluna)
        
        # pulando para próxima linha
        layout.addLayout(linha)
        linha = QHBoxLayout()
        
        # adicionando espaço forçadamente # FIXME: entender pq precisei forçar com um qlabel, antes não estava dando espaço
        self.saidaTexto = QLabel('')
        layout.addWidget(self.saidaTexto)  

        # > botões
        self.dialogoOpcoes = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.dialogoOpcoes.accepted.connect(self.accept)
        self.dialogoOpcoes.rejected.connect(self.reject)
        self.dialogoOpcoes.button(QDialogButtonBox.Ok).setText('Confirmar')
        self.dialogoOpcoes.button(QDialogButtonBox.Cancel).setText('Cancelar')
        linha.addWidget(self.dialogoOpcoes)
        layout.addLayout(linha)
        # definindo layout
        self.setLayout(layout)

    # -----------------------------------------
    # função para gerar entrada lista suspensa
    def f_geraEntradaListaSuspensaDistribuicoes(self, coluna, texto, parentDadoLido):
        entradaTextoGerada = QLabel(texto)
        coluna.addWidget(entradaTextoGerada)
        entradaListaSuspensaGerada = QComboBox()
        entradaListaSuspensaGerada.addItems(parentDadoLido['gestoesPossiveis'])
        return entradaListaSuspensaGerada

    # -----------------------------------------
    # função para retornar propriedades para a classe de janela principal
    def f_obtemPropriedades(self):
        return self.entradaGestorDaVez.currentText()

# -----------------------------------------
# função para geração dos gráficos de aderências
def f_plotaAderencias(self, datasCompletas):
    # obtendo número de dias por semana
    diasSemanaVez = 0
    diasPorSemana = []
    for indiceData in range(len(datasCompletas)-2):
        diferencaEmDias = abs((datetime.strptime(datasCompletas[indiceData+1], '%d-%b') - datetime.strptime(datasCompletas[indiceData], '%d-%b')).days)
        if diferencaEmDias <= 1:
            diasSemanaVez += 1
            pass
        else:
            diasPorSemana.append(diasSemanaVez+1)
            diasSemanaVez = 0
    diasPorSemana.append(diasSemanaVez+1)

    # verificando quantidade de ok e de entregue por dia
    contaOk = 0
    totalOk = []
    contaEntregue = 0
    totalEntregue = []
    contaEfprazo = 0
    totalEfprazo = []
    for colunaDaVez in range(self.quadroTarefas.columnCount()):
        for linhaDaVez in range(self.quadroTarefas.rowCount()):
            try:
                if self.quadroTarefas.item(linhaDaVez, colunaDaVez+3).text() != '':
                    contaOk += 1

                    # verificando quantidade de entregue
                    if self.quadroTarefas.cellWidget(linhaDaVez, 0).currentText() ==  'ENTREGUE':
                        contaEntregue += 1

                    # verificando quantidade de efprazo
                    if self.quadroTarefas.cellWidget(linhaDaVez, 0).currentText() ==  'EFPRAZO':
                        contaEfprazo += 1

            except: pass
        totalOk.append(contaOk)
        contaOk = 0
        totalEntregue.append(contaEntregue)
        contaEntregue = 0
        totalEfprazo.append(contaEfprazo)
        contaEfprazo = 0

    # verificando quantidade de ok e de entregue por semana
    auxOk = 0
    indiceOk = 0
    okSemanal = []
    auxEntregue = 0
    indiceEntregue = 0
    entregueSemanal = []
    auxEfprazo = 0
    indiceEfprazo = 0
    efprazoSemanal = []
    for diasSemanaVez in diasPorSemana:
        while diasSemanaVez > 0:
            auxOk += totalOk[indiceOk]
            indiceOk += 1
            auxEntregue += totalEntregue[indiceEntregue]
            indiceEntregue += 1
            auxEfprazo += totalEfprazo[indiceEfprazo]
            indiceEfprazo += 1
            diasSemanaVez -= 1
        okSemanal.append(auxOk)
        auxOk = 0
        entregueSemanal.append(auxEntregue)
        auxEntregue = 0
        efprazoSemanal.append(auxEfprazo)
        auxEfprazo = 0


    # cálculos
    okSemanal = np.array(okSemanal)
    entregueSemanal = np.array(entregueSemanal)
    efprazoSemanal = np.array(efprazoSemanal)

    # aderência semanal
    aderenciaSemanal = np.round(100 * entregueSemanal/okSemanal, 0)

    # aderência acumulada
    from itertools import accumulate
    okAcumulado = np.array(list(accumulate(okSemanal)))
    entregues = entregueSemanal + efprazoSemanal
    entregueComEfprazoAcumulado = np.array(list(accumulate(entregues)))
    aderenciaAcumulada = np.round(100 * entregueComEfprazoAcumulado/okAcumulado)

    # gerando lista de semanas
    nomesSemana = []
    for nSemana in range(len(aderenciaSemanal)):
        nomesSemana.append('Semana ' + str(nSemana+1))

    # plotando
    # curva de meta
    curvaMeta = go.Scatter(
        x = [nomesSemana[0], nomesSemana[-1]],
        y = [80, 80],
        name = 'Meta',
        marker = dict(color = 'rgba(93, 145, 69, .5)', size = 0),
        mode = 'lines+markers',
    )

        # criando figura com subplots
    graficoAderencia = make_subplots(
        rows = 2, cols = 1,
        shared_xaxes = True,
        subplot_titles = ['Semanal [%]', 'Acumulada [%]'],
        # vertical_spacing=  0.1,
    )

    aderenciaSemanal = go.Bar(
    x = nomesSemana,
    y = aderenciaSemanal,
    name = 'Semanal',
    marker = dict(color = 'rgba(44, 69, 148, 1)'),
    )
    graficoAderencia.add_trace(aderenciaSemanal, row = 1, col = 1)
    graficoAderencia.add_trace(curvaMeta, row = 1, col = 1)

    aderenciaAcumulada = go.Bar(
        x = nomesSemana,
        y = aderenciaAcumulada,
        name = 'Acumulada',
        marker = dict(color = 'rgba(44, 69, 148, 1)'),
    )
    graficoAderencia.add_trace(aderenciaAcumulada, row = 2, col = 1)
    graficoAderencia.add_trace(curvaMeta, row = 2, col = 1)

    # ajustando layout
    graficoAderencia.update_layout(
        barmode='group',  # Define o modo de agrupamento para barras
        showlegend = False,
        title_text = 'Aderências: ' + self.propriedadesGerais['gestorDaVez'].upper(),
    )

    # Exibindo figura
    graficoAderencia.show()


# -----------------------------------------
# função para plotar o gráfico pizza
def f_plotaTarefas(self): 
    planilha = 'W:\PLANEJAMENTO\GESTÃO À VISTA\VIBRACON VibraPlan\_arquivos\informativo.xlsx'  
    df = pd.read_excel(planilha, usecols=[1])
    valores = df.iloc[0:, 0].tolist()
    dfe = pd.read_excel(planilha, usecols=[2])
    valoresTotais = dfe.iloc[4:, 0].tolist()
    nomesLideres = ['MAURO', 'ROGÉRIO', 'MARCOS', 'ANA']
    cores = ['#5D9145', '#2C4594', '#EE964B', '#942c79']

    # Criando o gráfico de pizza
    grafico = go.Figure(data=[go.Pie(labels=nomesLideres, values=valores, textposition='inside', marker=dict(colors=cores))])
    # Adicionando título ao gráfico
    grafico.update_layout(title_text=f'RELAÇÃO DE TAREFAS')
    # Aumentando o tamanho da fonte 
    grafico.update_layout(legend_font_size=25, title_font_size=40)
    grafico.update_layout(height=850,  width=1700)
    grafico.update_layout(legend=dict(x=0.8, y=1))  

    
    fontConfig =dict(
                size=20,  # Tamanho da fonte
                color="black",  # Cor do texto
                family="Arial",  # Estilo da fonte
            )

    # Adicionando anotações
    anotacoes = [
        dict(
            x=0,
            y=0.15,
            text=f"TOTAL EM ANDAMENTO: {int(valoresTotais[0])}",
            showarrow=False,
            font=fontConfig
        ),
        dict(
            x=0,
            y=0.1,
            text=f"TOTAL JANEIRO: {int(valoresTotais[1])}",
            showarrow=False,
            font=fontConfig
        ),
        dict(
            x=0,
            y=0.05,
            text=f"TOTAL FEVEREIRO: {int(valoresTotais[2])}",
            showarrow=False,
            font=fontConfig
        ),
        dict(
            x=0,
            y=0,
            text=f"TOTAL MARÇO: {int(valoresTotais[3])}",
            showarrow=False,
            font=fontConfig
        ),
    ]

    # Adicionando as anotações ao layout do gráfico
    for anotacao in anotacoes:
        grafico.add_annotation(anotacao)

    # Exibindo o gráfico
    grafico.show()
