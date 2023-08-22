#
#  > [p2303] visualizador masterplan
#
#  > histórico de revisões
#      - 20230822
#        - autor: cádmo dias
#        - observações: 
#          - criação
#

# -----------------------------------------------------------------------------------------------
# > importando bibliotecas
import sys
import plotly.express as px
from PyQt5.QtWidgets import QMessageBox, QHBoxLayout, QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QFileDialog, QComboBox, QPushButton
from PyQt5.QtGui import QIcon, QPixmap, QFont
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QSize
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objs as go



# -----------------------------------------------------------------------------------------------
# > variáveis globais
# opções de exibição de gráficos
opcoesExibicao = ['Período de atividades', 'Recursos: Itens', 'Recursos: Pessoas']



# -----------------------------------------------------------------------------------------------
# > classes
# janela principal
class JanelaPrincipal(QMainWindow):
    # -----------------------------------------
    # função para inicializar janela principal
    def __init__(self):
        super().__init__()
        self.f_inicializaGui()



    # -----------------------------------------
    # função para inicializar interface
    def f_inicializaGui(self):
        # título
        self.setWindowTitle('VIBRACON Visualizador MasterPlan (v.2023.08.22)')

        # ícone
        icone = QIcon('_biblioteca/arte/logos/logoVibracon1.png')
        QApplication.setWindowIcon(icone)
        
        # tamanho e posição iniciais da janela
        self.setGeometry(100, 100, 800, 600)

        # bg
        self.bg = QLabel(self)
        self.bg.setGeometry(0, 0, 2500, 2500)
        pixmap = QPixmap('_biblioteca/arte/bg/bgUm.png')
        self.bg.setPixmap(pixmap)
        self.bg.setScaledContents(True)

        # layout
        layout = QVBoxLayout()
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))
        
        # cabeçalho/logo vibracon
        self.logoVibracon = QLabel(self)
        pixmap = QPixmap('_biblioteca/arte/logos/logoPrograma.png')
        self.logoVibracon.setPixmap(pixmap)
        #self.logoVibracon.setScaledContents(True)
        self.logoVibracon.setMaximumHeight(90)
        self.logoVibracon.setMaximumWidth(1500)

        # gerando coluna com botões de salvar e carregar projeto
        colunaBotoes = QVBoxLayout()

        # gerando linha de cabeçalho
        linhaCabecalho = QHBoxLayout()
        linhaCabecalho.addWidget(self.logoVibracon)
        linhaCabecalho.addLayout(colunaBotoes)

        # botao para selecionar masterplan
        self.botaoSelecionaMasterplan = self.f_criaBotao('', '_biblioteca/arte/botoes/botaoAbrirMasterplan.png', self.f_janelaSelecaoArquivo)
        colunaBotoes.addWidget(self.botaoSelecionaMasterplan)

        # lista suspensa para selecionar exibição
        self.listaOpcoes = QComboBox(self)
        self.listaOpcoes.addItems(opcoesExibicao)
        self.listaOpcoes.currentIndexChanged.connect(self.f_atualizaVisualizacao)
        colunaBotoes.addWidget(self.listaOpcoes)
        
        #
        layout.addLayout(linhaCabecalho)
    
        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # canvas para gráfico
        self.canvasParaGraficos = QWebEngineView(self)
        layout.addWidget(self.canvasParaGraficos)
        self.canvasParaGraficos.setStyleSheet('background-color: #2f419e;')



    # -----------------------------------------
    # função para criação de botões
    def f_criaBotao(self, textoBotao, imagemBotao, funcaoBotao):
        # personalização dos botões
        personalizacaoBotoes =  '''
            QPushButton {
                border: none;
            }
            QPushButton:hover {
                background-color: lightGray;
            }
        '''

        botaoDaVez = QPushButton(textoBotao)
        botaoDaVez.setIcon(QIcon(QPixmap(imagemBotao)))
        botaoDaVez.setIconSize(QSize(700, 70))
        botaoDaVez.setStyleSheet(personalizacaoBotoes)
        botaoDaVez.clicked.connect(funcaoBotao)
        return botaoDaVez



    # -----------------------------------------
    # janela para abertura do arquivo masterplan
    def f_janelaSelecaoArquivo(self):
        opcoes = QFileDialog.Options()
        opcoes |= QFileDialog.ReadOnly

        caminhoArquivo, _ = QFileDialog.getOpenFileName(
            self,
            'Selecione a planilha do MasterPlan',
            '',
            'Excel Files (*.xlsx *.xls)',
            options = opcoes
        )

        if caminhoArquivo:
            QMessageBox.warning(self, 'Atenção!', 'O seguinte MasterPlan será carregado:\n' + caminhoArquivo)
            #
            try:
                global dadosLidos
                dadosLidos = pd.read_excel(caminhoArquivo, sheet_name = 'MasterPlan')

                #
                self.f_atualizaVisualizacao()
            except:
                QMessageBox.warning(self, 'Atenção!', 'Favor verificar se a planilha Excel foi devidamente exportada do Project. Falha ao carregar o arquivo:\n' + caminhoArquivo)

        else:
            QMessageBox.warning(self, 'Atenção!', 'Nenhum MasterPlan foi selecionado')



    # -----------------------------------------
    # função para atualizar exibição do gráfico
    def f_atualizaVisualizacao(self):
        global dadosLidos
        
        # executando caso masterplan tenha sido selecionado
        try:
            # cor das barras
            corBarras = '#2f419e'

            # pegando valores da planilha
            nomesDosRecursos = dadosLidos['Nomes dos recursos'].tolist()
            inicioPlanejado = dadosLidos['Início Planejado'].tolist()
            terminoPlanejado = dadosLidos['Término Planejado'].tolist()

            # percorrendo todos os recursos
            diasPlanejado = []
            for indice in range(len(nomesDosRecursos)):
                try:
                    # passando datas de início e fim da vez para datetime
                    inicioDaVez = datetime.strptime(inicioPlanejado[indice].split(' ')[1], '%d/%m/%y')
                    fimDaVez = datetime.strptime(terminoPlanejado[indice].split(' ')[1], '%d/%m/%y')
                    # calculando quantidade de dias e salvando
                    diasPlanejado.append((fimDaVez - inicioDaVez).days + 1)
                except Exception as erro:
                    print('>> atenção! erro na linha ', str(indice), ' do masterplan!\n', erro, '\n\n')

            # pegando recursos únicos
            recursosUnicos = np.unique(nomesDosRecursos)

            # verificando quantidade de dias de todos os recursos
            recursosPessoas = []
            recursosItens = []
            for recursoDaVez in (recursosUnicos):
                nDiasRecursoDaVez = 0
                iniciosRecursoDaVez = []
                finaisRecursoDaVez = []
                for indice in range(len(nomesDosRecursos)):
                    # verificando se é recurso da vez e, caso sim, somando dias desse recurso
                    if nomesDosRecursos[indice] == recursoDaVez:
                        nDiasRecursoDaVez += diasPlanejado[indice]
                        iniciosRecursoDaVez.append(datetime.strptime(inicioPlanejado[indice].split(' ')[1], '%d/%m/%y'))
                        finaisRecursoDaVez.append(datetime.strptime(terminoPlanejado[indice].split(' ')[1], '%d/%m/%y'))

                # depois que passar pelo masterplan, salvando dias de utilização do recurso
                if 'Buscar Recurso_' in recursoDaVez or recursoDaVez == 'Arquivo Técnico':
                    recursosItens.append(
                        {
                            'nomeRecurso': recursoDaVez,
                            'diasRecurso': nDiasRecursoDaVez,
                            'iniciosRecurso': iniciosRecursoDaVez,
                            'finaisRecurso': finaisRecursoDaVez,
                        }
                    )
                else:
                    recursosPessoas.append(
                        {
                            'nomeRecurso': recursoDaVez,
                            'diasRecurso': nDiasRecursoDaVez,
                            'iniciosRecurso': iniciosRecursoDaVez,
                            'finaisRecurso': finaisRecursoDaVez
                        }
                    )

            # pegando opção desejada da lista suspensa
            exibicaoEscolhida = self.listaOpcoes.currentText()

            # exibindo gráfico de acordo com modo escolhido
            if exibicaoEscolhida == 'Recursos: Itens':
                fig = px.bar(recursosItens, x = 'nomeRecurso', y = 'diasRecurso', color_discrete_sequence = [corBarras])
                fig.update_layout(
                    title = '',
                    yaxis_title = 'Quantidade de dias',
                    xaxis_title = '',
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )
            elif exibicaoEscolhida == 'Recursos: Pessoas':
                fig = px.bar(recursosPessoas, x = 'nomeRecurso', y = 'diasRecurso', color_discrete_sequence = [corBarras])
                fig.update_layout(
                    title = '',
                    yaxis_title = 'Quantidade de dias',
                    xaxis_title = '',
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )
            elif exibicaoEscolhida == 'Período de atividades':
                fig = go.Figure()
                for recurso in recursosPessoas:
                    fig.add_trace(
                        go.Scatter(
                            x = recurso['iniciosRecurso'],
                            y = [recurso['nomeRecurso']] * len(recurso['iniciosRecurso']),
                            mode = 'markers',
                            name = recurso['nomeRecurso']
                        )
                    )
                fig.update_layout(
                    title = '',
                    yaxis_title = '',
                    xaxis_title = '',
                    showlegend = True, #FIXME: ENTENDER PQ NÃO ESTÁ APARECENDO LEGENDA
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )

            # atualizando figura
            fig.update_layout(showlegend = False)
            graficoNoCanvas = fig.to_html(full_html = False, include_plotlyjs = 'cdn')
            self.canvasParaGraficos.setHtml(graficoNoCanvas)

        # caso masterplan não tenha sido selecionado antes
        except Exception as e:
            QMessageBox.warning(self, 'Atenção!', 'Nenhum MasterPlan foi selecionado\n')
            None



# -----------------------------------------------------------------------------------------------
# principal, executando programa
if __name__ == '__main__':
    # criação do app
    app = QApplication(sys.argv)

    # definindo fonte e cor de fonte padrão
    fontePadrao = QFont('Gill Sans MT')
    app.setFont(fontePadrao)
    app.setStyleSheet('color: #2C4594; font: {0} 8pt;'.format(fontePadrao.family()))

    
    # incializando com janela principal
    janelaInicial = JanelaPrincipal()
    janelaInicial.show()
    janelaInicial.setStyleSheet('color: #2C4594;')  

    # executando
    sys.exit(app.exec_())