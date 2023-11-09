
# > importando bibliotecas
from PyQt5.QtWidgets import(QHBoxLayout, QVBoxLayout, QDialog, QLabel, QDialogButtonBox, QComboBox, QComboBox)
from PyQt5.QtCore import Qt

# -----------------------------------------
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
        self.setWindowTitle('Seleção de gestor')

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