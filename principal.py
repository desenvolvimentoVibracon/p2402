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
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication

import _biblioteca.codigos.feJanelaPrincipal as feJanelaPrincipal

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
    janelaInicial = feJanelaPrincipal.JanelaPrincipal()
    janelaInicial.show()
    janelaInicial.setStyleSheet('color: #2C4594;')  

    # executando
    sys.exit(app.exec_())