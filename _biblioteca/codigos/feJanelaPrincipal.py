# importando bibliotecas
import pickle
import pandas as pd
import sys
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QBrush, QColor, QFont
from PyQt5.QtWidgets import QScrollBar, QHeaderView, QMenu, QTextEdit, QAction, QApplication, QSizePolicy, QMessageBox, QLineEdit, QHBoxLayout, QMainWindow, QLabel, QVBoxLayout, QWidget, QFileDialog, QComboBox, QTableWidgetItem, QDialog, QInputDialog
from PyQt5 import QtGui
from datetime import datetime

import _biblioteca.codigos.beArquivos as beArquivos
import _biblioteca.codigos.feJanelasAux as feJanelasAux
import _biblioteca.codigos.feComponentes as feComponentes

# -----------------------------------------------------------------------------------------------
# classe da janela principal
class JanelaPrincipal(QMainWindow):
    # -----------------------------------------
    # função para inicializar janela 
    def __init__(self, indice):
        super().__init__()
        self.setStyleSheet('color: #2C4594; font: 15pt') 
        self.f_inicializaGui()
        # Ocultando o botão padrão de ajuda
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        # variáveis globais
        self.propriedadesGerais
        self.quantidadeColunas = 0
        self.linhasCabecalho = []
        self.diasPorSemana = []
        self.semanaAtual = 0
        # armazena dados importados do masterplan
        self.tabelaLida
        self.dictDropdown = {}
        self.indiceTabela = 0
        self.opcoesSelecionadas = []
        self.contadorDesvio = {
            'Aprovação do cliente': 0,
            'Prazo': 0,
            'Falta de recurso': 0,
            'Qualidade de entrega': 0,
            'Desenvolvimento da tarefa': 0,
            'Mobilização': 0,
            'Elaboração e verificação': 0,
            'Falta de prioridade': 0,
            'Arquivo técnico': 0,
            'Falha no planejamento': 0, 
            'Solicitação do cliente': 0,
            'Solicitação de exclusão pelo cliente': 0,
            'Falta de informação': 0,
            'Efeito climático/operacional': 0,
            '': 0
        }
        self.tarefasAcumuladas = {}
        self.salvaLinhasPlano = []
        # armazena status das tarefas
        global statusSelecionados
        statusSelecionados = [] 
        self.cabecalhos = []
        self.gestorAtual = None
        self.gestorAntigo = None
        # armazena datas, dia a dia, da primeira à última data lida do masterplan
        global datasCompletas
        datasCompletas = []
        self.indice = indice
        self.f_janelaCarregaProjeto()
    # -----------------------------------------
    # função para inicializar interface
    def f_inicializaGui(self):
        # > be
        self.propriedadesGerais = {
            'gestorDaVez': '',
            'gestoesPossiveis': []
        }
        self.tabelaLida = []
        data = datetime.now()
        dataParaExibir = data.strftime("%Y-%m-%d")
        # > fe
        # título da janela
        self.setWindowTitle(f'VIBRACON VibraPlan (v.{dataParaExibir})')

        # ícone
        self.setWindowIcon(QIcon('_biblioteca/arte/logos/logoVibracon1.png'))

        # posição e tamanho da janela
        self.setGeometry(100, 100, 1200, 900) 
                
        # bg
        self.bg = QLabel(self)
        self.bg.setGeometry(0, 0, 2500, 2500)
        pixmap = QPixmap('_biblioteca/arte/bg/bgUm.png')
        self.bg.setPixmap(pixmap)
        self.bg.setScaledContents(True)
        
        linha = QHBoxLayout()
        # criação de layout
        layout = QVBoxLayout() # vertical

        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.05))

        linha = QHBoxLayout() # horizontal

        self.labelGestor = QLabel(self)
        self.labelGestor.setStyleSheet("font-size: 20pt;")
        linha.addWidget(self.labelGestor)
        
        linha.addStretch()

        # logo vibracon 
        self.logoVibracon = QLabel(self)
        pixmap = QPixmap('_biblioteca/arte/logos/logoPrograma.png')
        self.logoVibracon.setPixmap(pixmap)
        self.logoVibracon.setMaximumHeight(80)
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
        font = QFont("Arial", 25)  
        self.quadroTarefas.setFont(font)

        # pulando para próxima linha
        # espaçamento vertical
        layout.addSpacing(int(self.height() * 0.01))

        # criando legenda para status
        linha = QHBoxLayout()
        linha.addStretch()

        
        legenda = {'PROGRAMADO': '#2C4594', 'ENTREGUE': '#5D9145', 'ATRASADO': '#CF0E0E', 'RISCO': '#FFCC33', 'EFPRAZO': '#942c79', 'INATIVO': '#6A6E70'}
        # Adicionando QLabel para cada entrada na legenda
        for texto, cor in legenda.items():
            # Criando o ponto com a cor do status
            ponto = QLabel()
            ponto.setStyleSheet(f"background-color: {cor}; border-radius: 5%;")
            linha.addWidget(ponto)

            # Criando QLabel com o texto
            textoLabel = QLabel(f"{texto}   ")
            textoLabel.setStyleSheet("font-size: 12pt;")
            textoLabel.setFont(QFont('Gill Sans MT'))
            linha.addWidget(textoLabel)
            
        layout.addLayout(linha)
        layout.addSpacing(int(self.height() * 0.01))

        # container para centralização do layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.showMaximized()

    # -----------------------------------------
    # função para atualizar exibição do gráfico
    def f_atualizaVisualizacao(self): 
        # chamando classe com janela para seleção do gestor
        janelaSelecionaGestor = feJanelasAux.JanelaSelecionaGestor(self)
        # salvando datas de tarefas atuais
        self.f_salvaNovasDatas()
        self.gestores = ['MAURO', 'ROGÉRIO', 'MARCOS', 'ANA']
        self.gestorAtual = self.gestores[self.indice]
        # atualizando
        self.propriedadesGerais['gestorDaVez'] = self.gestorAtual
        
        # limpando tabela
        self.quadroTarefas.clearContents()
        self.quadroTarefas.setRowCount(0)
        self.quadroTarefas.setColumnCount(0)
        
        # variável global que irá armazenar dia a dia do período lido
        global datasCompletas
        datasCompletas = []
        dataAtualSemHora = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
        riscoEmDias = 2

        # percorrendo todas as abas até encontrar a desejada
        for indiceTabela, gestaoPossivel in enumerate(self.propriedadesGerais['gestoesPossiveis']):
            # quando encontrar a desejada
            if gestaoPossivel == self.propriedadesGerais['gestorDaVez']:
                self.labelGestor.setText('COORDENADOR(A):  '+ self.propriedadesGerais['gestorDaVez'])
                # pegando datas mais antiga e nova e gerando delta de 1 dia entre datas
                for dataDaVez in self.tabelaLida[indiceTabela]['dados']['datas']:
                    if not pd.isna(dataDaVez):
                        if isinstance(dataDaVez, pd.Timestamp):
                            dataDaVez = dataDaVez.strftime('%d-%b-%y')
                        datasCompletas.append(dataDaVez) # adicionando a data da vez à variável datasCompletas
                # passando strings para datas
                datasCompletas = [pd.to_datetime(date, format='%d-%b-%y') for date in datasCompletas if isinstance(date, str)]   # Filtra apenas strings para conversão
                if datasCompletas:
                    primeiraData = min(datasCompletas)  # encontra a primeira data válida
                    ultimaData = max(datasCompletas)    # encontra a última data válida
                    datasCompletas = pd.date_range(start=primeiraData, end=ultimaData, freq='b').strftime('%d-%b').tolist()
                # definindo número de linhas e colunas de acordo com quantidade de tarefas e datas, respectivamente
                self.quantidadeColunas = self.tabelaLida[indiceTabela]['dados'].shape[1] + len(datasCompletas) - 3
                self.quadroTarefas.setRowCount(self.tabelaLida[indiceTabela]['dados'].shape[0])
                self.quadroTarefas.setColumnCount(self.quantidadeColunas)
                header = self.quadroTarefas.horizontalHeader()
                font = header.font()
                font.setPointSize(15)
                header.setFont(font)
                self.quadroTarefas.setHorizontalHeaderLabels(['STATUS', 'TAREFAS', ] + datasCompletas + ['CAUSA DO DESVIO', 'PLANO DE AÇÃO', 'DATA', 'RESPONSÁVEL', 'STATUS DO PLANO'])
                
                # chamando scroll 
                self.scrollTimer = QTimer(self)
                self.scrollTimer.timeout.connect(self.f_scrollDown)
                self.scrollTimer.start(20000)
                self.scrollCont = 1
                
                
                # largura das colunas, em ordem 
                larguras = [110, 400] + [30] * (self.quantidadeColunas - 7) + [120, 200, 100, 50, 30]
                for indiceColuna, largura in enumerate(larguras):
                    self.quadroTarefas.setColumnWidth(indiceColuna, largura)
                for linha in range(self.quadroTarefas.rowCount()):
                    self.quadroTarefas.setRowHeight(linha, 40)
                # listando tarefas
                colunaTarefas = self.tabelaLida[indiceTabela]['dados'].iloc[:, 2]
                colunaPlanoAcao = self.tabelaLida[indiceTabela]['dados']['planoDeAcao']
                colunaDataPlano = self.tabelaLida[indiceTabela]['dados']['dataDoPlano']
                colunaNomePlano = self.tabelaLida[indiceTabela]['dados']['coordenador']
                self.indiceTabela = indiceTabela
                
                for indiceTarefa in range(colunaTarefas.shape[0]):
                    # tarefa
                    item = QTableWidgetItem(str(colunaTarefas[indiceTarefa]))
                    self.quadroTarefas.setItem(indiceTarefa, 1, item)
                    # plano de ação     
                    try:
                        item = QTableWidgetItem(str(colunaPlanoAcao[indiceTarefa]))
                        self.quadroTarefas.setItem(indiceTarefa, self.quantidadeColunas-4, item)
                    except: pass
                    # data do plano de ação 
                    try:
                        item = QTableWidgetItem(str(colunaDataPlano[indiceTarefa]))
                        self.quadroTarefas.setItem(indiceTarefa, self.quantidadeColunas-3, item)
                    except: pass
                    # nome do coordenador do plano 
                    try:
                        item = QTableWidgetItem(str(colunaNomePlano[indiceTarefa]))
                        self.quadroTarefas.setItem(indiceTarefa, self.quantidadeColunas-2, item)
                    except: pass
                try:
                    for okDaVez in self.tabelaLida[indiceTabela]['registroDatas']:
                        QApplication.processEvents()
                        if dataTarefaDaVez < dataAtualSemHora:
                            # Marcar como atrasado
                            self.f_coloreStatus(indiceTabela, okDaVez['linha'], 'ATRASADO', okDaVez['coluna'])
                        elif (dataTarefaDaVez - dataAtualSemHora).days <= riscoEmDias:
                            # Marcar como "RISCO"
                            self.f_coloreStatus(indiceTabela,okDaVez['linha'], 'RISCO', okDaVez['coluna'])
                        else:
                            #Marcar como "OK"
                            self.f_coloreStatus(indiceTabela,okDaVez['linha'], 'OK', okDaVez['coluna'])

                        # adicionando lista suspensa na coluna de status quando houver tarefa
                        self.f_adicionaListaStatus(indiceTabela, okDaVez['linha'], okDaVez['coluna'])
                        self.f_adicionaListaDesvio(indiceTabela, okDaVez['linha'])

                except Exception:
                    # identificando datas de realização de cada tarefa
                    datasTarefas = self.tabelaLida[indiceTabela]['dados'].iloc[:, 3]
                    for linhaDaVez, dataTarefaDaVez in enumerate(datasTarefas):
                        QApplication.processEvents()
                        try:
                            # selecionando coluna: adicionando 3 devido às colunas de cabeçalho
                            nomeColunaDaVez = dataTarefaDaVez.strftime('%d-%b')
                            indiceColunaDaVez = datasCompletas.index(nomeColunaDaVez) + 2

                            if (dataTarefaDaVez < dataAtualSemHora):
                                # Marcar como "ATRASADO"
                                self.f_coloreStatus(indiceTabela, linhaDaVez,  'ATRASADO', indiceColunaDaVez)
                            elif (dataTarefaDaVez - dataAtualSemHora).days <= riscoEmDias:
                                # Marcar como "RISCO"
                                self.f_coloreStatus(indiceTabela, linhaDaVez,  'RISCO', indiceColunaDaVez)
                            else:   
                                #Marcar como "OK"
                                self.f_coloreStatus(indiceTabela, linhaDaVez,  'OK', indiceColunaDaVez)

                            # adicionando lista suspensa na coluna de status quando houver tarefa                              
                            self.f_adicionaListaStatus(indiceTabela, linhaDaVez, indiceColunaDaVez)
                            self.f_adicionaListaDesvio(indiceTabela, linhaDaVez)
                        except Exception: pass
                    break
        # chamada da função que salva o plano de ação
        self.quadroTarefas.itemChanged.connect(lambda item, indiceTabela=indiceTabela: self.f_atualizouPlanoDeAcao(item, indiceTabela))
        #criando menu suspenso com o clique direito do mouse
        self.quadroTarefas.setContextMenuPolicy(Qt.CustomContextMenu)
        self.quadroTarefas.customContextMenuRequested.connect(lambda pos: self.f_abrirMenu(indiceTabela, pos))
        
        # adicionando dropdown de status do plano
        for linha in range(self.quadroTarefas.rowCount()):
            if self.quadroTarefas.item(linha, self.quantidadeColunas-4).text() != ' ':
                self.f_adicionaListaStatusPlano(indiceTabela, linha)
                self.salvaLinhasPlano.append(linha)

        # removendo colunas
        for i in range(1, 6):
            self.quadroTarefas.removeColumn(self.quantidadeColunas - i)
        self.quadroTarefas.removeColumn(0)
        
        # Ajustando o modo de redimensionamento para que as colunas se estiquem para preencher o espaço
        self.quadroTarefas.setColumnWidth(0, 875)
        for column in range(1, self.quadroTarefas.columnCount()):
            self.quadroTarefas.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)
            self.quadroTarefas.verticalHeader().hide()
        # chamada da função que colore as semanas 
        #self.f_contaDiasPorSemana()
        # chamada da função que identifica cabeçalho
        self.f_identificaCabecalho() 

    # ------------------------------------------
    # função do scroll automático 
    def f_scrollDown(self):
            # definindo automação do scroll
            scroll = self.quadroTarefas.verticalScrollBar()
            valorMax = scroll.maximum()
            valorAtual = scroll.value()
            self.scrollTimer.setInterval(2000)
            novoValor = min(valorAtual + self.scrollCont, valorMax)
            scroll.setValue(novoValor)

            #condição para parada do scroll
            if novoValor == valorMax:
                self.scrollTimer.timeout.disconnect(self.f_scrollDown)
                self.scrollTimer.stop()

                # inicializando gráficos
                graficoAderencia, graficoEngenharia, graficoDesvio = feJanelasAux.f_plotaAderencias(self, datasCompletas, self.contadorDesvio)
                graficoAderenciaHtml = graficoAderencia.to_html(full_html=False, include_plotlyjs='cdn')
                graficoEngenhariaHtml = graficoEngenharia.to_html(full_html=False, include_plotlyjs='cdn')
                graficoDesvioHtml = graficoDesvio.to_html(full_html=False, include_plotlyjs='cdn')
                graficoCoordenadores, graficoProjetos = feJanelasAux.f_plotaTarefas(self)
                graficoCoordenadoresHtml = graficoCoordenadores.to_html(full_html=False, include_plotlyjs='cdn')
                graficoProjetosHtml = graficoProjetos.to_html(full_html=False, include_plotlyjs='cdn')
                # chamando função que exibe gráficos
                janelaGraficos = feJanelasAux.JanelaGraficos(graficoAderenciaHtml, graficoEngenhariaHtml, graficoDesvioHtml, graficoCoordenadoresHtml, graficoProjetosHtml, self.gestorAtual, self.indice)
                janelaGraficos.exec_()
                self.close()

    # funções do vibraplan
    # ---------------------------------
    # função que exibe o plano de ação 
    def f_pegaValorDropdown(self, linha):
        # identificando emissor do sinal de atualização de lista suspensa
        campoListaSuspensa = self.sender()
        valorDaVez = campoListaSuspensa.currentText()

        if isinstance(campoListaSuspensa, QComboBox):
            # texto de status na variável
            self.tabelaLida[self.indiceTabela]['dados'].iloc[linha, -1] = valorDaVez

    # -----------------------------------------
    # função que atualiza as colunas do plano
    def f_retornaValorStatusPlano(self, linha):
        try:
            return self.tabelaLida[self.indiceTabela]['dados'].iloc[linha, -1]
        except Exception as erro: 
            pass
    # -----------------------------------------
    # função que atualiza as colunas do plano
    def f_atualizaStatusPlano(self):
        for linha, valor in self.dictDropdown.items():
            self.tabelaLida[self.indiceTabela]['dados'].at[linha, 'statusPlano'] = valor
    # -----------------------------------------
    # função que atualiza as colunas do plano
    def f_atualizouPlanoDeAcao(self, item, indiceTabela):
        # atualizando as colunas do plano, data e responsável pelo plano de ação 
        if item.column() == self.quantidadeColunas-4:
            novoValor = item.text()
            linha = item.row()
            self.tabelaLida[indiceTabela]['dados'].at[linha, 'planoDeAcao'] = novoValor

        elif item.column() == self.quantidadeColunas-3:  
            novoValor = item.text()
            linha = item.row()
            self.tabelaLida[indiceTabela]['dados'].at[linha, 'dataDoPlano'] = novoValor

        elif item.column() == self.quantidadeColunas-2:  
            novoValor = item.text()
            linha = item.row()
            self.tabelaLida[indiceTabela]['dados'].at[linha, 'coordenador'] = novoValor       
            
    # ---------------------------------------------
    # Função para o usuário selecionar o mês
    def f_selecionaMes(self):
        meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        mes, ok = QInputDialog.getItem(self, "Selecionar Mês", "Escolha o mês a ser aberto:", meses, 0, False)
        if ok and mes:
            return meses.index(mes) + 1 

    # -------------------------------------------
    # função para adicionar lista suspensa de status
    def f_adicionaListaStatusPlano(self, indiceTabela, linhaDaVez):
        dropdown = QComboBox()
        dropdown.setStyleSheet('QComboBox { background-color: transparent; border: 0px}')
        dropdown.addItems(['NÃO OK', 'OK'])
        dropdown.currentIndexChanged.connect(lambda index: self.f_atualizouListaPlano(linhaDaVez, indiceTabela))
        if isinstance(self.tabelaLida[indiceTabela]['dados']['statusPlano'][linhaDaVez], str):
            pass
        else:
            self.tabelaLida[indiceTabela]['dados']['statusPlano'][linhaDaVez] = ''
        dropdown.setCurrentText(self.tabelaLida[indiceTabela]['dados']['statusPlano'][linhaDaVez])
        self.quadroTarefas.setCellWidget(linhaDaVez, self.quantidadeColunas-1, dropdown)
        for linha in range(self.quadroTarefas.rowCount()):
            if self.quadroTarefas.item(linha, self.quantidadeColunas-4).text() == '':
                self.quadroTarefas.removeCellWidget(linha, self.quantidadeColunas-1)
        
    # -----------------------------------------
    # função trigada quando lista suspensa é atualizada
    def f_atualizouListaPlano(self, linhaDaVez, indiceTabela):
        # identificando emissor do sinal de atualização de lista suspensa
        campoListaSuspensa = self.sender()
        valorDaVez = campoListaSuspensa.currentText()

        if isinstance(campoListaSuspensa, QComboBox):
            # texto de status na variável
            self.tabelaLida[indiceTabela]['dados'].iloc[linhaDaVez, -1] = valorDaVez

    # -------------------------------------------
    # função para adicionar lista suspensa de status
    def f_adicionaListaStatus(self, indiceTabela, linhaDaVez, colunaDaVez):
        dropdown = QComboBox()
        dropdown.setStyleSheet('QComboBox { background-color: transparent; border: 0px}')
        dropdown.addItems(['OK', 'ENTREGUE', 'RISCO', 'ATRASADO', 'EFPRAZO', 'INATIVO'])
        dropdown.currentIndexChanged.connect(lambda index, indiceTarefas=linhaDaVez: self.f_atualizouListaStatus(indiceTarefas, indiceTabela, colunaDaVez))
        dropdown.setCurrentText(self.tabelaLida[indiceTabela]['dados']['status'][linhaDaVez]) 
        self.quadroTarefas.setCellWidget(linhaDaVez, 0, dropdown)

    # -----------------------------------------
    # função trigada quando lista suspensa é atualizada
    def f_atualizouListaStatus(self, linhaDaVez, indiceTabela, colunaDaVez):
        try:
            # identificando emissor do sinal de atualização de lista suspensa
            campoListaSuspensa = self.sender()
            valorDaVez = campoListaSuspensa.currentText()
            try:
                # Inicializando o dict com os valores salvos
                self.tarefasAcumuladas = self.tabelaLida[indiceTabela]['dictAcumulado']
            except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro! Arquivo fora do padrão.\n\n {str(erro)}')
            if valorDaVez == 'EFPRAZO': 
                # Salvar a data da alteração
                dataAtual = datetime.now().day
                tarefa = f'linha {linhaDaVez}' 
                if tarefa not in self.tarefasAcumuladas:
                    if self.tabelaLida[indiceTabela][linhaDaVez]['desvios'] == 'Aprovação do cliente' or self.tabelaLida[indiceTabela][linhaDaVez]['desvios'] == 'Solicitação de exclusão pelo cliente' or self.tabelaLida[indiceTabela][linhaDaVez]['desvios'] == 'Falta de informação':
                        pass
                    else: 
                        self.tarefasAcumuladas[tarefa] = dataAtual # Salvando o dia que a alteração foi feita
                        self.tabelaLida[indiceTabela]['dictAcumulado'] = self.tarefasAcumuladas # Salvando dict na tabela lida
            else:
                # removendo a opção que foi alterada para outro status 
                tarefa = f'linha {linhaDaVez}'
                if tarefa in self.tarefasAcumuladas:
                    del self.tarefasAcumuladas[tarefa]
            if isinstance(campoListaSuspensa, QComboBox):
                self.f_coloreStatus(indiceTabela, linhaDaVez, valorDaVez, colunaDaVez)
                # texto de status na variável
                self.tabelaLida[indiceTabela]['dados'].iloc[linhaDaVez, 1] = valorDaVez
        except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro ao atualizar status.\n\n {str(erro)}')

    # -----------------------------------------
    # função para adicionar lista suspensa de causa de desvio
    def f_adicionaListaDesvio(self, indiceTabela, linhaDaVez): 
        # Caso contrário, cria um novo widget e o adiciona à célula
        dropdown = QComboBox()
        dropdown.setStyleSheet('QComboBox { background-color: transparent; border: 0px}')
        dropdown.addItems(['', 'Aprovação do cliente', 'Prazo', 'Falta de recurso', 'Qualidade de entrega', 'Desenvolvimento da tarefa', 
                        'Mobilização', 'Elaboração e verificação', 'Falta de prioridade', 'Arquivo técnico', 'Falha no planejamento', 
                        'Solicitação do cliente', 'Solicitação de exclusão pelo cliente', 'Falta de informação', 'Efeito climático/operacional'])
        dropdown.currentIndexChanged.connect(lambda index, linha=linhaDaVez: self.f_atualizouListaDesvio(linha, indiceTabela))
        dropdown.setCurrentText(self.tabelaLida[indiceTabela]['dados']['desvios'][linhaDaVez])
        self.quadroTarefas.setCellWidget(linhaDaVez, self.quantidadeColunas-5, dropdown)

    #----------------------------------------------------------------
    # função para salvar lista de desvio
    def f_atualizouListaDesvio(self, linha, indiceTabela): 
        # identificando emissor do sinal de atualização de lista suspensa
        campoListaSuspensa = self.sender()
        valorDaVez = campoListaSuspensa.currentText()
        if isinstance(campoListaSuspensa, QComboBox):
            # texto de status na variável
            self.tabelaLida[indiceTabela]['dados'].iloc[linha, -6] = valorDaVez
        
        # Verifica se a lista de opções selecionadas já foi inicializada
        if 'opcoesSelecionadas' not in self.__dict__:
            self.opcoesSelecionadas = []

        # Verifica se já existe uma opção selecionada para a linha atual
        for item in self.opcoesSelecionadas:
            if item['linha'] == linha:
                # Diminui o contador da opção anterior em 1
                self.contadorDesvio[item['opcao']] -= 1
                break
        
        # Atualiza a opção selecionada e seu respectivo contador
        self.opcoesSelecionadas.append({
            'linha': linha,
            'opcao': valorDaVez 
        })
        self.contadorDesvio[valorDaVez] += 1

    # -------------------------------------------
    # função para adicionar linhas a um projeto em andamento 
    def f_adicionaLinha(self, indiceTabela):
            janelaAdicionaLinha = feJanelasAux.janelaAdicionaLinha(self)
            if janelaAdicionaLinha.exec_() == QDialog.Accepted:     

                dir = 'W:/PLANEJAMENTO/GESTÃO À VISTA/Vibraplan/_arquivos/Documentos/novasLinhas.xlsx'
                dfNovasLinhas = pd.read_excel(dir)
                tarefas = dfNovasLinhas.iloc[:, 0].tolist()
                datas = dfNovasLinhas.iloc[:, 1].tolist()
                indiceTarefa = int(janelaAdicionaLinha.retornaLinha())
                linhas = []

                # criando dict para cada linha a ser adicionada
                for textoData, nomeTarefa in zip(datas, tarefas):   
                    dia = textoData.day
                    dataFinal = textoData.replace(day=dia, hour=0, minute=0, second=0, microsecond=0)
                    novaLinhaDict = {
                        'cor': 'f7f7f7',
                        'status': 'OK', 
                        'tarefas': nomeTarefa,  
                        'datas': dataFinal,
                        'desvios': '', 
                        'planoDeAcao': '',  
                        'dataDoPlano': '',  
                        'coordenador': ''  
                    }
                    linhas.append(novaLinhaDict)

                # separando o df atual na linha desejada e adicionando as novas linhas
                dfTabelaLida = pd.DataFrame(self.tabelaLida[indiceTabela]['dados'])
                parteSuperior = dfTabelaLida.iloc[:indiceTarefa]
                parteInferior = dfTabelaLida.iloc[indiceTarefa:]
                cont=1
                for linha in linhas:
                    linha = pd.DataFrame(linha, index=[0])
                    dfTabelaLida = pd.concat([parteSuperior, linha, parteInferior], ignore_index=True)
                    parteSuperior = dfTabelaLida.iloc[:indiceTarefa+cont]
                    cont+=1

                self.tabelaLida[indiceTabela]['dados'] = []
                self.tabelaLida[indiceTabela]['dados'] = dfTabelaLida
                self.f_atualizaVisualizacao()

            
    # -------------------------------------------
    # função que mostra o contador de desvios
    def f_exibirDesvios(self): #FIXME
        try:
            # Exibe a mensagem com os contadores
            mensagem = ''
            for indice, (opcao, contador) in enumerate(self.contadorDesvio.items()):
                # Verifica se não é a última linha da lista
                if indice != len(self.contadorDesvio) - 1:
                    mensagem += f'{opcao}: {contador}\n'
            QMessageBox.information(self, 'Contadores de Desvio', mensagem)   

        except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro ao exibir desvios!\n {str(erro)}')          
    # ----------------------------------------
    # função que conta os dias da semana para separá-las (copiada de feJanelasAux)
    def f_contaDiasPorSemana(self):
        # obtendo número de dias por semana
        diasSemanaVez = 0
        contaDias = 0 
        diasPorSemana = []
        for indiceData in range(len(datasCompletas)-1):
            diferencaEmDias = abs((datetime.strptime(datasCompletas[indiceData+1], '%d-%b') - datetime.strptime(datasCompletas[indiceData], '%d-%b')).days)
            if diferencaEmDias <= 1:
                diasSemanaVez += 1
                pass
            else:
                diasPorSemana.append(diasSemanaVez+1)           
                diasSemanaVez = 0
        diasPorSemana.append(diasSemanaVez+1)

        for semana in range(len(diasPorSemana)):
            coluna = contaDias + 2
            # separando semanas pares e ímpares
            if semana %2 != 0:
                for i in range(diasPorSemana[semana]):
                    for linha in range(self.quadroTarefas.rowCount()):
                        atual = self.quadroTarefas.item(linha, coluna)    
                        if atual is None or atual == '':
                            self.quadroTarefas.setItem(
                                linha, coluna, QTableWidgetItem(str('SD')))
                            self.quadroTarefas.item(linha, coluna).setBackground(QBrush(QColor('#f5f5f5')))
                            self.quadroTarefas.item(linha, coluna).setForeground(QBrush(QColor('#f5f5f5')))
                    coluna +=1
            contaDias += diasPorSemana[semana]
            coluna = 0
        self.diasPorSemana = diasPorSemana

    # ------------------------------------------
    # função para inicializar janela de gráficos
    def f_abreJanelaGraficoAderencias(self):
        try:
            feJanelasAux.f_plotaAderencias(self, datasCompletas, self.contadorDesvio)
        except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro ao plotar gráficos.\n\n {str(erro)}')
    # ------------------------------------------
    # função para inicializar janela de gráficos
    def f_abreJanelaGraficoProjetos(self):
        try:
            feJanelasAux.f_plotaTarefas(self)
        except Exception as erro: QMessageBox.critical(self, 'AVISO', f'Erro ao plotar gráficos.\n\n {str(erro)}')

    # -----------------------------------------
    # função para salvar novas datas modificadas já no vibraplan
    def f_salvaNovasDatas(self):
        registrosDatas = []
        for linhaDaVez in range(self.quadroTarefas.rowCount()):
            for colunaDaVez in range(2, self.quadroTarefas.columnCount() - 4):
                # obtendo valor da célula caso ela exista
                celulaDaVez = self.quadroTarefas.item(linhaDaVez, colunaDaVez)
                if celulaDaVez is not None:
                    if celulaDaVez.text().upper() != ' ' :
                        registrosDatas.append({'linha': linhaDaVez, 'coluna': colunaDaVez})
        
        # percorrendo todas as abas até encontrar a da liderança atual
        for indiceTabela, gestaoPossivel in enumerate(self.propriedadesGerais['gestoesPossiveis']):
            # quando encontrar a desejada
            if gestaoPossivel == self.propriedadesGerais['gestorDaVez']:
                self.tabelaLida[indiceTabela]['registroDatas'] = registrosDatas
            break

    # -----------------------------------------
    # função do botão que envia emails 
    def f_enviaEmail(self):
        try:
            janelaEmail = feJanelasAux.janelaOpcoesEmail(self)
            janelaEmail.exec_()

        except Exception as erro:
             QMessageBox.information(self, 'AVISO', f'Falha ao enviar e-mail!\n {str(erro)}')
    # -----------------------------------------
    # função para atualizar lista de gestores
    def f_atualizaGestores(self):
        # atualizando lista de abas (gestores)
        self.propriedadesGerais['gestoesPossiveis'] = []
        for tabelaDaVez in self.tabelaLida: 
            self.propriedadesGerais['gestoesPossiveis'].append(tabelaDaVez['gestor'])

    # ----------------------------------------
    # função trigada quando célula é atualizada       
    def f_atualizouCelula(self):
        # cores de acordo com textos
        textoCor = {'OK': '#2C4594', 'ENTREGUE': '#5D9145', 'ATRASADO': '#CF0E0E' , 'RISCO': '#FFCC33', 'EFPRAZO': '#942c79', 'INATIVO': '#6A6E70'}
        
        # percorrendo linhas e colunas da tabela
        for linhaDaVez in range(self.quadroTarefas.rowCount()):
            for colunaDaVez in range(self.quadroTarefas.columnCount()):
                # obtendo valor da célula caso ela exista
                celulaDaVez = self.quadroTarefas.item(linhaDaVez, colunaDaVez)
                if celulaDaVez is not None:
                    celulaDaVezValor = celulaDaVez.text().upper()
                    
                    # se texto da célula atual estiver na lista de cores de acordo com texto, atualizo cores de fundo e da fonte
                    if celulaDaVezValor in textoCor:
                        corDaVez = textoCor[celulaDaVezValor]
                        self.quadroTarefas.item(linhaDaVez, colunaDaVez).setBackground(QBrush(QColor(corDaVez)))
                        self.quadroTarefas.item(linhaDaVez, colunaDaVez).setForeground(QBrush(QColor(corDaVez)))
      
    # -----------------------------------------  
    # função para colorir status de acordo com o definido
    def f_coloreStatus(self, indiceTabela, linhaDaVez, valorDaVez, colunaDaVez):
        if self.tabelaLida[indiceTabela]['dados']['status'][linhaDaVez] == 'OK':
            if valorDaVez == 'ENTREGUE':
                pass
            else:
                valorDaVez = 'OK'
        # colorindo a data da tarefa da vez
        self.quadroTarefas.setItem(
            linhaDaVez, colunaDaVez,
            QTableWidgetItem(str(valorDaVez))
            )      
    
    # --------------------------------------------
    # função para identificar linhas de cabeçalho
    def f_identificaCabecalho(self):
    # percorrendo as linhas da tabela
        for linhaDaVez in range(self.quadroTarefas.rowCount()):
            # Verificar se todas as letras na linha estão em caixa alta
            linhaTexto = ""
            for coluna in range(self.quadroTarefas.columnCount() - 1):
                try:
                    linhaTexto += str(self.quadroTarefas.item(linhaDaVez, coluna).text())
                except AttributeError:
                    pass
            # chamando função que colore o background
            if linhaTexto.isupper():
                self.cabecalhos.append(linhaDaVez)
                self.f_personalizaCabecalho(linhaDaVez)
        
    # -----------------------------------------        
    # função para personalização dos cabeçalhos da tabela
    def f_personalizaCabecalho(self, linhaDaVez):
    # percorrendo três primeiras colunas de cabeçalho
        for colunaDaVez in range(0, 1):
            # pegando célula: caso não exista (células vazias), crio
            celulaDaVez = self.quadroTarefas.item(linhaDaVez, colunaDaVez)
            if celulaDaVez is None:
                celulaDaVez = QTableWidgetItem()
                self.quadroTarefas.setItem(linhaDaVez, colunaDaVez, celulaDaVez)

            celulaDaVez.setBackground(QtGui.QColor('#2C4594'))
            
            celulaDaVez.setForeground(QtGui.QColor('white'))
            fonteDaVez = celulaDaVez.font()
            fonteDaVez.setBold(True)
            celulaDaVez.setFont(fonteDaVez)
            
    # -----------------------------------------
    # janela para carregar projeto
    def f_janelaCarregaProjeto(self):
        # abrindo janela de diálogo
        # atualizar 1 vez no mês
        if self.indice == 0:
            caminhoArquivo = 'W:/PLANEJAMENTO/GESTÃO À VISTA/Vibraplan/_arquivos/Projetos/Jul_2024/Mauro_jul24.projVibraPlan'
        elif self.indice == 1:
            caminhoArquivo = 'W:/PLANEJAMENTO/GESTÃO À VISTA/Vibraplan/_arquivos/Projetos/Jul_2024/Rogerio_jul24.projVibraPlan'
        elif self.indice == 2:
            caminhoArquivo = 'W:/PLANEJAMENTO/GESTÃO À VISTA/Vibraplan/_arquivos/Projetos/Jul_2024/Marcos_jul24.projVibraPlan'
        else:
            caminhoArquivo = 'W:/PLANEJAMENTO/GESTÃO À VISTA/Vibraplan/_arquivos/Projetos/Jul_2024/Ana_jul24.projVibraPlan'

        # caso usuário defina nome e caminho, lendo os dados do projeto
        if caminhoArquivo:
            try:
                with open(caminhoArquivo, 'rb') as file:
                    self.tabelaLida = pickle.load(file)

                # atualizando lista de gestores e visualização da tabela
                self.f_atualizaGestores()
                self.f_atualizaVisualizacao()

            except Exception as erro:
                QMessageBox.critical(self, 'AVISO', f'Falha ao carregar o projeto.\n\n {str(erro)}')
        else:
            QMessageBox.warning(self, 'AVISO', 'Nenhum projeto foi selecionado!')

    # -----------------------------------------
    # janela para salvamento do projeto
    def f_janelaSalvaProjeto(self):
        # salvando datas de tarefas atuais
        self.f_salvaNovasDatas()
        # abrindo janela de diálogo
        caminhoArquivo, _ = QFileDialog.getSaveFileName(self, 'SALVAR', '', 'Arquivos VibraPlan (*.projVibraPlan);;All Files (*)')

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
    def f_chamaExcel(self):
        try:
            #chamando a função que exporta o plano para excel          
            dados = self.tabelaLida[self.indiceTabela]['dados']
            gestor = self.propriedadesGerais['gestorDaVez']
            beArquivos.f_salvaExcel(dados, gestor)
            QMessageBox.information(self, 'AVISO', f'E-mail enviado!')
        except Exception as erro:
            QMessageBox.critical(self, 'AVISO', f'Não foi possível enviar o e-mail. \n {str(erro)}')
    # -----------------------------------------
    # janela para abertura do arquivo masterplan
    def f_janelaImportaMasterplan(self):
        try:
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
                
                    # abrindo arquivo e lendo cada aba presente
                    self.tabelaLida = beArquivos.f_abrePlanilha(caminhoArquivo)

                    # 
                    self.f_atualizaGestores()

                    # atualizando exibição
                    self.f_atualizaVisualizacao()

                

            else:
                QMessageBox.warning(self, 'AVISO', 'Nenhum MasterPlan foi selecionado!')
        except Exception: 
            QMessageBox.critical(self, 'AVISO', f'Arquivo fora do padrão!')