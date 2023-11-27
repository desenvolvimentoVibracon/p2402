# importando bibliotecas
from PyQt5.QtCore import QSize
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QPushButton, QTableWidget

# -----------------------------------------------------------------------------------------------
# função para criação de botões
def f_criaBotao(textoBotao, imagemBotao, funcaoBotao):
    # personalização
    personalizacao =  '''
        QPushButton {
            border: none;
        }
        QPushButton:hover {
            background-color: lightGray;
        }
    '''

    # configurações
    botaoCriado = QPushButton(textoBotao)
    botaoCriado.setIcon(QIcon(QPixmap(imagemBotao)))
    botaoCriado.setIconSize(QSize(350, 70))
    botaoCriado.setStyleSheet(personalizacao)
    botaoCriado.clicked.connect(funcaoBotao)
    
    return botaoCriado



# -----------------------------------------
# função para criação de tabelas
def f_criaTabela(funcaoTabela):
    # personalização
    personalizacao =  '''
        background-color: white;
        QTableWidget { font-family: 'Gill Sans MT', sans-serif; }
        QTableWidget QTableCornerButton::section { background-color: lightGray; border: none; }
        QTableWidget::item { font-family: 'Gill Sans MT', sans-serif; border: none; }
        QTableWidget::item:selected { background-color: lightGray; }
    '''

    # configurações
    tabelaCriada = QTableWidget()
    tabelaCriada.itemChanged.connect(funcaoTabela)
    tabelaCriada.setStyleSheet(personalizacao)

    return tabelaCriada