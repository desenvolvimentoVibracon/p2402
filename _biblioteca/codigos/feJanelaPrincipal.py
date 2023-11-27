# importando bibliotecas
import pickle

import pandas as pd

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QBrush, QColor
from PyQt5.QtWidgets import QMessageBox, QHBoxLayout, QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QFileDialog, QComboBox, QTableWidgetItem, QDialog
from PyQt5 import QtGui

import _biblioteca.codigos.beArquivos as beArquivos
import _biblioteca.codigos.feJanelasAux as feJanelasAux
import _biblioteca.codigos.feComponentes as feComponentes

# -----------------------------------------------------------------------------------------------
# classe da janela principal
class JanelaPrincipal(QMainWindow):
    # -----------------------------------------
    # função para inicializar janela principal
    def __init__(self):
        super().__init__()
        self.f_inicializaGui()

        # variáveis globais
        # armazena liderança selecionada e lideranças possíveis
        self.propriedadesGerais

        # armazena dados importados do masterplan
        self.tabelaLida
        
        # armazena status das tarefas
        global statusSelecionados
        statusSelecionados = []
        
        # armazena datas, dia a dia, da primeira à última data lida do masterplan
        global datasCompletas
        datasCompletas = []



    # -----------------------------------------
    # função para inicializar interface
    def f_inicializaGui(self):
        # > be
        self.propriedadesGerais = {
            'gestorDaVez': '',
            'gestoesPossiveis': []
        }
        self.tabelaLida = []

        # > fe
        # título da janela
        self.setWindowTitle('VIBRACON VibraPlan (v.2023.11.21)')

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

        # canvas para tabela
        self.quadroTarefas = feComponentes.f_criaTabela(self.f_atualizouCelula)
        layout.addWidget(self.quadroTarefas)

        # pulando para próxima linha
        layout.addLayout(linha)
        linha = QHBoxLayout()

        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # botao para carregar projeto
        self.botaoCarregaProjeto = feComponentes.f_criaBotao('', '_biblioteca/arte/botoes/botaoAbrirProjeto.png', self.f_janelaCarregaProjeto)
        linha.addWidget(self.botaoCarregaProjeto)

        # botao para salvar projeto
        self.botaoSalvaProjeto = feComponentes.f_criaBotao('', '_biblioteca/arte/botoes/botaoSalvarProjeto.png', self.f_janelaSalvaProjeto)
        linha.addWidget(self.botaoSalvaProjeto)

        # botao para selecionar gestor
        self.botaoSelecionaGestor = feComponentes.f_criaBotao('', '_biblioteca/arte/botoes/botaoMudarLideranca.png', self.f_abreJanelaSelecaoGestor)
        linha.addWidget(self.botaoSelecionaGestor)

        # botao para selecionar masterplan
        self.botaoSelecionaMasterplan = feComponentes.f_criaBotao('', '_biblioteca/arte/botoes/botaoImportarMasterplan.png', self.f_janelaImportaMasterplan)
        linha.addWidget(self.botaoSelecionaMasterplan)

        # botao para mostrar gráficos
        self.botaoMostraGraficos = feComponentes.f_criaBotao('', '_biblioteca/arte/botoes/botaoAderencias.png', self.f_abreJanelaGraficos)
        linha.addWidget(self.botaoMostraGraficos)

        # adicionando linha
        layout.addLayout(linha)
    
        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        # container para centralização do layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)



    # -----------------------------------------
    # função para atualizar exibição do gráfico
    def f_atualizaVisualizacao(self):        
        # limpando tabela
        self.quadroTarefas.clearContents()
        self.quadroTarefas.setRowCount(0)
        self.quadroTarefas.setColumnCount(0)
        
        # variável global que irá armazenar dia a dia do período lido
        global datasCompletas
        datasCompletas = []
        
        # percorrendo todas as abas até encontrar a desejada
        for indiceTabela, gestaoPossivel in enumerate(self.propriedadesGerais['gestoesPossiveis']):
            # quando encontrar a desejada
            if gestaoPossivel == self.propriedadesGerais['gestorDaVez']:
                # pegando datas mais antiga e nova e gerando delta de 1 dia entre datas # pegando datas mais antiga e nova e gerando delta de 1 dia entre datas
                for dataDaVez in self.tabelaLida[indiceTabela]['dados']['datas']:
                    if not pd.isna(dataDaVez):
                        datasCompletas.append(dataDaVez)

                # passando strings para datas
                datasCompletas = sorted([pd.to_datetime(date) for date in datasCompletas], key=lambda x: (pd.isnull(x), x))

                # definindo um intervalo de dadas completas, do primeiro ao último dia dispostos na aba da liderança atual, dia a dia
                primeiraData = datasCompletas[0]
                ultimaData = datasCompletas[-1]
                datasCompletas = pd.date_range(start = primeiraData, end = ultimaData, freq = 'b').strftime('%d-%b-%Y').tolist()

                # definindo número de linhas e colunas de acordo com quantidade de tarefas e datas, respectivamente
                self.quadroTarefas.setRowCount(self.tabelaLida[indiceTabela]['dados'].shape[0])
                self.quadroTarefas.setColumnCount(self.tabelaLida[indiceTabela]['dados'].shape[1]+len(datasCompletas)-1)
                self.quadroTarefas.setHorizontalHeaderLabels(['', 'STATUS', 'TAREFAS'] + datasCompletas)

                # listando tarefas
                colunaTarefas = self.tabelaLida[indiceTabela]['dados'].iloc[:, 2]
                for indiceTarefa in range(colunaTarefas.shape[0]):
                    item = QTableWidgetItem(str(colunaTarefas.iloc[indiceTarefa]))
                    self.quadroTarefas.setItem(indiceTarefa, 2, item)

                # identificando datas de realização de cada tarefa
                datasTarefas = self.tabelaLida[indiceTabela]['dados'].iloc[:, 3]
                for linhaDaVez, dataTarefaDaVez in enumerate(datasTarefas):
                    try:
                        # selecionando coluna: adicionando 3 devido às colunas de cabeçalho
                        nomeColunaDaVez = dataTarefaDaVez.strftime('%d-%b-%Y')
                        indiceColunaDaVez = datasCompletas.index(nomeColunaDaVez) + 3
                        self.quadroTarefas.setItem(
                            linhaDaVez, indiceColunaDaVez,
                            QTableWidgetItem(str('OK'))
                        )
                
                        # adicionando lista suspensa na coluna de status quando houver tarefa
                        dropdown = QComboBox()
                        dropdown.setStyleSheet("QComboBox { background-color: transparent; border: 0px}")
                        dropdown.addItems(['ND', 'ENTREGUE', 'RISCO', 'ATRASADO', 'EFPRAZO'])
                        dropdown.currentIndexChanged.connect(lambda index, indiceTarefas = linhaDaVez: self.f_atualizouListaSuspensa(indiceTarefas, indiceTabela))
                        dropdown.setCurrentText(self.tabelaLida[indiceTabela]['dados']['status'][linhaDaVez])
                        self.quadroTarefas.setCellWidget(linhaDaVez, 1, dropdown)

                    except Exception: pass

                break

        # personalizando linhas cujo tarefas são cabeçalhos
        self.f_identificaCabecalho()

        # definindo tamanho das primeiras três colunas: cor do status, status e tarefa
        self.quadroTarefas.setColumnWidth(0, 50)
        self.quadroTarefas.setColumnWidth(1, 160)
        self.quadroTarefas.setColumnWidth(2, 500)



    # -----------------------------------------
    # função para inicializar janela de gráficos
    def f_abreJanelaGraficos(self):
        # chamando classe com janela para seleção do gesto
        try: feJanelasAux.f_plotaAderencias(self, datasCompletas)
        except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro ao plotar as aderências.\n\n {str(erro)}')



    # -----------------------------------------
    # função para inicializar janela de seleção de gestor
    def f_abreJanelaSelecaoGestor(self):
        # chamando classe com janela para seleção do gesto
        janelaSelecionaGestor = feJanelasAux.JanelaSelecionaGestor(self)
        if janelaSelecionaGestor.exec_() == QDialog.Accepted:
            self.propriedadesGerais['gestorDaVez'] = janelaSelecionaGestor.f_obtemPropriedades()

            # atualizando visualização com o gestor escolhido
            if self.propriedadesGerais['gestorDaVez'] != '':
                QMessageBox.information(self, 'AVISO', f"Aguarde enquanto a visualização é atualizada!\n\n Liderança selecionada: {self.propriedadesGerais['gestorDaVez']}")
                self.f_atualizaVisualizacao()



    # -----------------------------------------
    # função para atualizar lista de gestores
    def f_atualizaGestores(self):
        # atualizando lista de abas (gestores)
        self.propriedadesGerais['gestoesPossiveis'] = []
        for tabelaDaVez in self.tabelaLida: self.propriedadesGerais['gestoesPossiveis'].append(tabelaDaVez['gestor'])
        self.propriedadesGerais['gestorDaVez'] = self.propriedadesGerais['gestoesPossiveis'][-2]



    # -----------------------------------------
    # função trigada quando célula é atualizada
    def f_atualizouCelula(self):
        # cores de acordo com textos
        textoCor = { 'OK': '#2C4594', 'NOK': '#5D9145', }
        
        # percorrendo linhas e colunas da tabela
        for linhaDaVez in range(self.quadroTarefas.rowCount()):
            for colunaDaVez in range(self.quadroTarefas.columnCount()):
                # obtendo valor da célula caso ela exista
                celulaDaVez = self.quadroTarefas.item(linhaDaVez, colunaDaVez)
                if celulaDaVez is not None:
                    celulaDaVezValor = celulaDaVez.text().upper()
                    
                    # se texto da célula atual tiver na lista de cores de acordo com texto, atualizo cores de fundo e da fonte
                    if celulaDaVezValor in textoCor:
                        corDaVez = textoCor[celulaDaVezValor]
                        self.quadroTarefas.item(linhaDaVez, colunaDaVez).setBackground(QBrush(QColor(corDaVez)))
                        self.quadroTarefas.item(linhaDaVez, colunaDaVez).setForeground(QBrush(QColor(corDaVez)))



    # -----------------------------------------
    # função trigada quando lista suspensa é atualizada
    def f_atualizouListaSuspensa(self, linhaDaVez, indiceTabela):
            # identificando emissor do sinal de atualização de lista suspensa
            campoListaSuspensa = self.sender()

            # pegando texto atual da lista suspensa
            valorDaVez = campoListaSuspensa.currentText()

            # caso de fato tenha lista suspensa, atualizo
            if isinstance(campoListaSuspensa, QComboBox):
                # cor da linha na primeira coluna
                self.f_coloreStatus(linhaDaVez, valorDaVez)
                
                # texto de status na variável
                self.tabelaLida[indiceTabela]['dados'].iloc[linhaDaVez, 1] = valorDaVez



    # -----------------------------------------
    # função para colorir status de acordo com o definido
    def f_coloreStatus(self, linhaDaVez, valorDaVez):
            
            # definindo a cor de acordo com o texto
            corDaVez = '#f7f7f7'
            if valorDaVez == 'ENTREGUE': corDaVez = '#5D9145'
            elif valorDaVez == 'RISCO': corDaVez = '#ee964b'
            elif valorDaVez == 'ATRASADO': corDaVez = '#DF6158'
            elif valorDaVez == 'EFPRAZO': corDaVez = '#942c79'
            
            # colorindo célula da primeira coluna e linha selecionada: caso não exista (célula vazia), uma nova célula é criada
            celulaDaVez = self.quadroTarefas.item(linhaDaVez, 0)
            if celulaDaVez:
                celulaDaVez.setBackground(QColor(corDaVez))
            else:
                novaCelula = QTableWidgetItem()
                novaCelula.setBackground(QColor(corDaVez))
                self.quadroTarefas.setItem(linhaDaVez, 0, novaCelula)



    # -----------------------------------------
    # função para identificar linhas de cabeçalho
    def f_identificaCabecalho(self):
        # percorrendo as linhas da tabela
        for linhaDaVez in range(self.quadroTarefas.rowCount()):

            # caso não tenha lista suspensa de status, defino como cabeçalho e coloro
            linhaListaSuspensa = self.quadroTarefas.cellWidget(linhaDaVez, 1)
            if not isinstance(linhaListaSuspensa, QComboBox):
                self.f_personalizaCabecalho(linhaDaVez)



    # -----------------------------------------
    # função para personalização dos cabeçalhos da tabela
    def f_personalizaCabecalho(self, linhaDaVez):
        # percorrendo três primeiras colunas de cabeçalho
        for colunaDaVez in range(0, 3):
            # pegando célula: caso não exista (células vazias), crio
            celulaDaVez = self.quadroTarefas.item(linhaDaVez, colunaDaVez)
            if celulaDaVez is None:
                celulaDaVez = QTableWidgetItem()
                self.quadroTarefas.setItem(linhaDaVez, colunaDaVez, celulaDaVez)

            # personalizando: definindo bg, fg e fonte
            celulaDaVez.setBackground(QtGui.QColor('#e9e9e9'))
            celulaDaVez.setForeground(QtGui.QColor('#2C4594'))
            fonteDaVez = celulaDaVez.font()
            fonteDaVez.setBold(True)
            celulaDaVez.setFont(fonteDaVez)



    # -----------------------------------------
    # janela para carregar projeto
    def f_janelaCarregaProjeto(self):
        # abrindo janela de diálogo
        caminhoArquivo, _ = QFileDialog.getOpenFileName(self, 'CARREGAR', '', 'Arquivos VibraPlan (*.projVP);;All Files (*)')

        # caso usuário defina nome e caminho, lendo os dados do projeto
        if caminhoArquivo:
            try:
                with open(caminhoArquivo, 'rb') as file:
                    self.tabelaLida = pickle.load(file)

                # atualizando lista de gestores e visualização da tabela
                self.f_atualizaGestores()
                self.f_atualizaVisualizacao()

                QMessageBox.information(self, 'AVISO', f'Projeto carregado com sucesso!\n\n {caminhoArquivo}')
            except Exception as erro:
                QMessageBox.critical(self, 'AVISO', f'Falha ao carregar o projeto.\n\n {str(erro)}')
        else:
            QMessageBox.warning(self, 'AVISO', 'Nenhum projeto foi selecionado!')



    # -----------------------------------------
    # janela para salvamento do projeto
    def f_janelaSalvaProjeto(self):
        # abrindo janela de diálogo
        caminhoArquivo, _ = QFileDialog.getSaveFileName(self, 'SALVAR', '', 'Arquivos VibraPlan (*.projVP);;All Files (*)')

        # caso usuário defina nome e caminho, salvando os dados do projeto
        if caminhoArquivo:
            try:
                with open(caminhoArquivo, 'wb') as file:
                    pickle.dump(self.tabelaLida, file)

                QMessageBox.information(self, 'AVISO', f'Sucesso ao salvar o projeto!\n\n {caminhoArquivo}')
            except Exception as erro:
                QMessageBox.critical(self, 'AVISO', f'Falha ao carregar o projeto.\n\n {str(erro)}')
        else:
            QMessageBox.warning(self, 'AVISO', 'Projeto sem salvar!')



    # -----------------------------------------
    # janela para abertura do arquivo masterplan
    def f_janelaImportaMasterplan(self):
        opcoes = QFileDialog.Options()
        opcoes |= QFileDialog.ReadOnly

        # abrindo janela de diálogo
        caminhoArquivo, _ = QFileDialog.getOpenFileName(
            self,
            'IMPORTAR PROJETO',
            '',
            'Arquivos exportados do Masterplan (*.xlsx *.xls)',
            options = opcoes
        )

        # tentando abrir arquivo caso selecionado
        if caminhoArquivo:
            try:
                # abrindo arquivo e lendo cada aba presente
                self.tabelaLida = beArquivos.f_abrePlanilha(caminhoArquivo)

                self.f_atualizaGestores()

                # atualizando exibição
                self.f_atualizaVisualizacao()

                QMessageBox.information(self, 'AVISO', f'Dados importados com sucesso!\n\n {caminhoArquivo}')
            except Exception as erro:
                self.propriedadesGerais['gestoesPossiveis'] = []
                self.propriedadesGerais['gestorDaVez'] = ''
                QMessageBox.warning(self, 'AVISO', 'Erro ao abrir o MasterPlan\n\n' + str(erro))

        else:
            QMessageBox.warning(self, 'AVISO', 'Nenhum MasterPlan foi selecionado!')