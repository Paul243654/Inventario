import sys
import os
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtCore import *
from PyQt5.QtCore import QTimer, QTime, Qt

class Acerca_de(QDialog):
    def __init__(self):
        super().__init__()

        # NUMERO 1/2
        #nombre_archivo3=self.resolver_ruta3("acercade.ui")
        #uic.loadUi(nombre_archivo3, self)
        uic.loadUi("C:/Users/Paul/Desktop/PROYECTOS INFORMATICOS/FICHERO EXCEL_PYTHON/acercade.ui", self)

        self.setFixedSize(QSize(446, 168))
        self.setStyleSheet("background-color: azure;")

    # NUMERO 2/2
    def resolver_ruta3(self,ruta_relativa3):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, ruta_relativa3)
        return os.path.join(os.path.abspath('.'), ruta_relativa3)
    
if __name__ == "__main__":
    app=QApplication(sys.argv)
    QApplication.setAttribute(Qt.AA_DisableWindowContextHelpButton)  
    GUIde=Acerca_de()
    GUIde.show()
    sys.exit(app.exec_())

               