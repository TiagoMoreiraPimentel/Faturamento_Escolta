import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence, QPixmap, QBrush, QPalette
from PyQt5.QtWidgets import QMainWindow, QApplication, QShortcut

from tela_menu import Ui_MainWindow

# pyuic5 -x tela_menu.ui -o tela_menu.py

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.showFullScreen()
        self.set_background()

        # Atalho para fechar o formulário com a tecla Esc
        QShortcut(QKeySequence("Esc"), self).activated.connect(self.fechar_tela_cadastro_usuario)

    def fechar_tela_cadastro_usuario(self):
        self.close()

    def set_background(self):
        # Usando uma imagem como fundo
        pixmap = QPixmap('ico/fundo_security.jpg')

        if pixmap.isNull():
            print("Erro ao carregar a imagem. Verifique o caminho do arquivo.")
            return

        # Ajusta a imagem para o tamanho da tela sem distorção
        self.pixmap_scaled = pixmap.scaled(self.size(), Qt.IgnoreAspectRatio, Qt.SmoothTransformation)

        # Cria o fundo com a imagem redimensionada
        brush = QBrush(self.pixmap_scaled)
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setBrush(QPalette.Background, brush)
        self.setPalette(palette)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())