#
#  > [p2402] visualizador vibraplan
#      - autor: lucas teixeira
#      - observações:
#           - primeiro commit
#

# -----------------------------------------------------------------------------------------------
# > importando bibliotecas
import sys
from PyQt5.QtWidgets import QApplication
import _biblioteca.codigos.feJanelasAux as feJanelasAux

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # inicializando projeto com índice 0
    janelapts = feJanelasAux.JanelaPts(0)
    sys.exit(app.exec_())