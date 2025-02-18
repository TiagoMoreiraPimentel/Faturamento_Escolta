import os
import sys
import oracledb
import pandas as pd

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence, QPixmap, QBrush, QPalette
from PyQt5.QtWidgets import QMainWindow, QApplication, QShortcut, QAction, QMessageBox
from tkinter import Tk, filedialog
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
        # Adicionando o menu "SISTEMA"
        self.sistema_menu = menubar.addMenu("SISTEMA")
        # Adicionando ações ao menu
        self.fechar_action = QAction("FECHAR SISTEMA", self)
        self.sistema_menu.addAction(self.fechar_action)
        # Adicionando o menu "OPERAÇÃO"
        self.operacao_menu = menubar.addMenu("OPERAÇÃO")
        # Adicionando ações ao menu
        self.importar_excel_action = QAction("IMPORTAR EXCEL", self)
        self.operacao_menu.addAction(self.importar_excel_action)
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
        self.importar_excel_action.triggered.connect(self.chamar_funcao_importar_ler_excel)

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

    def chamar_funcao_importar_ler_excel(self):
        # Utilização da função que realiza o upload de arquivos para o sistema
        data = self.importar_ler_excel()
        if data:
            for sheet, df in data.items():
                print(f"Aba: {sheet}")
                print(df.head())  # Exibe as primeiras linhas de cada aba

    def importar_ler_excel(self):
        # Abre um diálogo para selecionar o arquivo
        Tk().withdraw()  # Esconde a janela principal do Tkinter
        file_path = filedialog.askopenfilename(title="Selecione a planilha Excel",
                                               filetypes=[("Excel Files", "*.xlsx;*.xls")])

        if not file_path:
            print("Nenhum arquivo selecionado.")
            return None

        try:
            # Lê a planilha inteira
            df = pd.read_excel(file_path, sheet_name=None)  # Lê todas as abas

            self.registrar_excel_banco(file_path)

            return df
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None

    def get_db_credentials(self):
        return {
            "user": os.getenv("DB_USER", "DSCSECTOOLS"),
            "password": os.getenv("DB_PASSWORD", "DSCs3cT00Ls2022"),
            "dsn": os.getenv("DB_DSN", "QASBRRACT2-SCAN.PHX-DC.DHL.COM:1521/DSCSECTOOLS"),
        }

    # Função que realizar o registro das informações da planilha no banco de dados oracle.
    def registrar_excel_banco(self, file_path):
        try:
            creds = self.get_db_credentials()

            # Lê os dados do Excel
            df = pd.read_excel(file_path, dtype=str)  # Lê tudo como string para evitar erros

            with oracledb.connect(**creds) as connection:
                print("Conexão estabelecida.")
                with connection.cursor() as cursor:
                    print("Executando a inserção...")

                    query = """INSERT INTO FATURAMENTO_ESCOLTA_BASE (
                        N_SE, CLIENTES, DESCRICAO_SE, EMPRESA_ESCOLTA, SITUACAO, COBERTURA, 
                        AGENDAMENTO_DATA_HORA, CHEGADA_ORIGEM_DATA_HORA, CHEGADA_DESTINO_DATA_HORA, 
                        FIM_EMISSAO_REAL, FRANQUIA_HORAS, VALOR_HORA_EXCEDENTE, DISTANCIA_REAL, 
                        FRANQUIA_KM, VALOR_KM_EXCEDENTE, PRECO_FRANQUIA_BASE, VALOR_TOTAL_EMISSAO, 
                        STATUS_PAGAMENTO, VIAGEM_CONSIDERADA, STATUS
                    ) VALUES (
                        :1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14, :15, :16, :17, :18, :19, :20
                    )"""

                    # Contador para exibir o progresso
                    total_registros = len(df)
                    contador = 0

                    # Inserir cada linha do DataFrame
                    for _, row in df.iterrows():
                        try:
                            # Assegura que os valores são strings e lidando com valores nulos
                            row = row.fillna("")  # Substitui valores nulos por uma string vazia
                            values = tuple(row)  # Converte a linha em uma tupla

                            print(f"Inserindo registro: {row.to_dict()}")  # Exibe o registro a ser inserido

                            cursor.execute(query, values)  # Executa a inserção

                            # Incrementa o contador e exibe o progresso no terminal
                            contador += 1
                            print(f"Registros inseridos: {contador}/{total_registros}", end='\r')

                        except Exception as e:
                            print(f"Erro ao inserir linha {row.to_dict()}: {e}")

                    # Commit das inserções
                    connection.commit()
                    print(f"\nDados inseridos com sucesso. Total de registros: {contador}")
                    QtWidgets.QMessageBox.information(self, "Sucesso", "Dados registrados com sucesso!")

        except oracledb.DatabaseError as e:
            print(f"Erro de banco de dados: {e}")
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha na conexão ou na inserção: {e}")
        except Exception as e:
            print(f"Erro inesperado: {e}")
            QtWidgets.QMessageBox.critical(self, "Erro inesperado", f"Ocorreu um erro inesperado: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
