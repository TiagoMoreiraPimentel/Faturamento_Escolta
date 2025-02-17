import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence, QPixmap, QBrush, QPalette
from PyQt5.QtWidgets import QMainWindow, QApplication, QShortcut, QAction, QMessageBox

from tela_menu import Ui_MainWindow

# pyuic5 -x tela_menu.ui -o tela_menu.py


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.showFullScreen()
        self.set_background()

        # Inicializa a variável pixmap_scaled
        self.pixmap_scaled = None

        # Atalho para fechar o formulário com a tecla Esc
        QShortcut(QKeySequence("Esc"), self).activated.connect(self.fechar_tela_menu)

        # Criando a barra de menu
        menubar = self.menuBar()
        # Adicionando o menu "CADASTRAR"
        self.sistema_menu = menubar.addMenu("SISTEMA")
        # Adicionando ações ao menu
        self.fechar_action = QAction("FECHAR SISTEMA", self)
        self.sistema_menu.addAction(self.fechar_action)
        # Estilo do menu - fazendo os itens parecerem botões com borda arredondada
        self.setStyleSheet(
            """
                            QMenuBar {
                                background-color: #fecb00;
                                color: black;
                                border: 2px solid #ccc;
                                border-radius: 20px;
                            }
                            QMenuBar::item {
                                background-color: #d0d3db;
                                color: black;
                                padding: 5px 20px;
                                border: 1px solid black; /* Adiciona borda preta */
                                margin: 1px; /* Adiciona margem para espaçamento */
                            }
                            QMenuBar::item:selected {
                                background-color: #d0d3db;
                                color: white;
                            }
                            QMenuBar::item:pressed {
                                background-color: #5d605d;
                                color: white;
                            }
                            QMenu {
                                background-color: #d0d3db;
                                border: 1px solid #ccc;
                                border-radius: 5px;
                            }
                            QMenu::item {
                                background-color: #d0d3db;
                                border: 1px solid #ccc;
                                border-radius: 5px;
                                padding: 5px 20px;
                            }
                            QMenu::item:selected {
                                background-color: #5d605d;
                                color: white;
                            }
                            QAction {
                                padding: 10px;
                                background-color: #5d605d;
                                border-radius: 5px;
                            }
                            QAction:hover {
                                background-color: #5d605d;
                            }
                        """
        )
        # Conectando eventos específicos para cada ação
        self.fechar_action.triggered.connect(self.fechar_tela_menu)

    def fechar_tela_menu(self):
        """Exibe uma caixa de diálogo de confirmação antes de fechar o formulário."""
        resposta = QMessageBox.question(
            self,
            "Confirmação",
            "Tem certeza de que deseja fechar o sistema?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resposta == QMessageBox.Yes:
            QApplication.quit()

    def set_background(self):
        # Usando uma imagem como fundo
        pixmap = QPixmap("ico/fundo_preto.jpg")

        if pixmap.isNull():
            print("Erro ao carregar a imagem. Verifique o caminho do arquivo.")
            return

        # Ajusta a imagem para o tamanho da tela sem distorção
        self.pixmap_scaled = pixmap.scaled(
            self.size(), Qt.IgnoreAspectRatio, Qt.SmoothTransformation
        )

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
