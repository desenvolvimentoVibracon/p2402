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


from PyQt5.QtCore import Qt, QSize, QTimer
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import(QApplication, QMainWindow, QPushButton, QComboBox, QVBoxLayout, QTextEdit, QWidget, QDialog, QLabel, QFileDialog, QHBoxLayout)


import numpy as np
from datetime import datetime
import plotly.graph_objs as go
import _biblioteca.codigos.beArquivos as beArquivos
import _biblioteca.codigos.feJanelasAux as feJanelasAux
import pandas as pd


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


        # variáveis globais
        self.propriedadesGerais
        self.abasLidas


    # -----------------------------------------
    # função para inicializar interface
    def f_inicializaGui(self):
        # > be
        # variáveis
        self.propriedadesGerais = {
            'gestorDaVez': '',
            'gestoesPossiveis': []
        }
        self.abasLidas = [ ]

        # > fe
        # título da janela
        self.setWindowTitle('VIBRACON VibraPlan (v.dev)')

        # ícone
        icone = QIcon('_biblioteca/arte/logos/logoVibracon1.png')
        QApplication.setWindowIcon(icone)

        # posição e tamanho da janela
        self.setGeometry(100, 100, 1200, 900) 
                
        # bg
        self.bg = QLabel(self)
        self.bg.setGeometry(0, 0, 2500, 2500)
        pixmap = QPixmap('_biblioteca/arte/bg/bgUm.png')
        self.bg.setPixmap(pixmap)
        self.bg.setScaledContents(True)

        # criação de layout
        layout = QVBoxLayout()

        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # logo vibracon
        linha = QHBoxLayout()
        self.logoVibracon = QLabel(self)
        pixmap = QPixmap('_biblioteca/arte/logos/logoPrograma.png')
        self.logoVibracon.setPixmap(pixmap)
        self.logoVibracon.setMaximumHeight(90)
        self.logoVibracon.setMaximumWidth(1000)
        linha.addWidget(self.logoVibracon)
        linha.setAlignment(Qt.AlignRight)

        # pulando para próxima linha
        layout.addLayout(linha)
        linha = QHBoxLayout()
        
        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # canvas para gráfico
        self.canvasParaGraficos = QWebEngineView(self)
        layout.addWidget(self.canvasParaGraficos)
        self.canvasParaGraficos.setStyleSheet('background-color: #2f419e;')

        # pulando para próxima linha
        layout.addLayout(linha)
        linha = QHBoxLayout()

        # entrada de arquivos
        self.saidaTexto = QLabel('→ ENTRADA DE DADOS:')

        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # pulando para próxima linha
        layout.addLayout(linha)
        linha = QHBoxLayout()

        # botao para selecionar masterplan
        self.botaoSelecionaMasterplan = self.f_criaBotao('', '_biblioteca/arte/botoes/botaoAbrirMasterplan.png', self.f_janelaSelecaoArquivo)
        linha.addWidget(self.botaoSelecionaMasterplan)

        # botao para selecionar masterplan
        self.botaoSelecionaGestor = self.f_criaBotao('', '_biblioteca/arte/botoes/botaoAbrirMasterplan.png', self.f_janelaSelecaoGestor)
        linha.addWidget(self.botaoSelecionaGestor)

        #
        layout.addLayout(linha)
    
        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # container para centralização do layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)



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
    # janela para 
    def f_janelaSelecaoGestor(self):
        janelaSelecionaGestor = feJanelasAux.JanelaSelecionaGestor(self)
        if janelaSelecionaGestor.exec_() == QDialog.Accepted:
            self.propriedadesGerais['gestorDaVez'] = janelaSelecionaGestor.f_obtemPropriedades()

            # atualizando visualização com o gestor escolhido
            if self.propriedadesGerais['gestorDaVez'] != '': self.f_atualizaVisualizacao()

        else:
            pass



    # -----------------------------------------
    # janela para abertura do arquivo masterplan
    def f_janelaSelecaoArquivo(self):
        opcoes = QFileDialog.Options()
        opcoes |= QFileDialog.ReadOnly

        caminhoArquivo, _ = QFileDialog.getOpenFileName(
            self,
            'Selecione a planilha exportada do MasterPlan',
            '',
            'Excel Files (*.xlsx *.xls)',
            options = opcoes
        )

        # tentando abrir arquivo caso selecionado
        if caminhoArquivo:
            try:
                # abrindo arquivo e lendo cada aba presente
                self.abasLidas = beArquivos.f_abrePlanilha(caminhoArquivo)

                # atualizando lista de abas (gestores)
                gestoresUnicos = []
                for abaDaVez in self.abasLidas: gestoresUnicos.append(abaDaVez['nomeGestor'])
                self.propriedadesGerais['gestorDaVez'] = gestoresUnicos[-2]
                self.propriedadesGerais['gestoesPossiveis'] = gestoresUnicos

                # atualizando exibição
                self.f_atualizaVisualizacao()

                #
                datasUnicas = self.f_listaDatasUnicas()



            except Exception as erro:
                QMessageBox.warning(self, 'Atenção!', 'Erro ao ler os dados do MasterPlan: \n' + str(erro))

        else:
            QMessageBox.warning(self, 'Atenção!', 'Nenhum MasterPlan foi selecionado')




    # -----------------------------------------
    # 
    def f_listaDatasUnicas(self):
        datasUnicas = []

        # concatenando todas as datas presentes para cada gestor
        datasConcatenadas = []
        for abaDaVez in self.abasLidas:
            datasConcatenadas += abaDaVez['datas']

        # filtrando valores únicos
        datasUnicas = list(set(datasConcatenadas))
        
        # ordenando
        datasUnicas = sorted(datasUnicas)
        
        # removendo primeiro valor (NaT das linhas de título) e retornando datas únicas
        return datasUnicas[1:] 



    # -----------------------------------------
    # função para atualizar exibição do gráfico
    def f_atualizaVisualizacao(self):
        # percorrendo todas as abas até encontrar a desejada
        for abaDaVez in self.abasLidas:
            # quando encontrar a desejada
            if abaDaVez['nomeGestor'] == self.propriedadesGerais['gestorDaVez']:
                print('nomeGestor', abaDaVez['nomeGestor'])
                print('tarefas', abaDaVez['tarefas'])
                print('datas', abaDaVez['datas'])
                break



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