import os
import sys
import time
import traceback

import oracledb
import pandas as pd
import threading

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QMetaObject, pyqtSignal
from PyQt5.QtGui import QKeySequence, QPixmap, QBrush, QPalette, QMovie
from PyQt5.QtWidgets import QMainWindow, QApplication, QShortcut, QAction, QMessageBox, QWidget, QVBoxLayout, QLabel
from tkinter import Tk, filedialog
from tela_menu import Ui_MainWindow


class MainWindow(QMainWindow):

    sinal_finalizacao = pyqtSignal(int)  # Cria um sinal que recebe um inteiro

    def __init__(self):

        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.showFullScreen()
        self.set_background()

        # sinalizar processamento para fechar a thread
        self.sinal_finalizacao.connect(self.finalizarProcessamento)

        # Inicializa o QMovie e associa à label_loading (se necessário)
        self.gif = QMovie("ico/loading.gif")  # Substitua pelo caminho correto do GIF
        self.ui.label_loading.setMovie(self.gif)

        # ocultar
        self.ui.label_loading.setVisible(False)

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

    def get_db_credentials(self):
        return {
            "user": os.getenv("DB_USER", "DSCSECTOOLS"),
            "password": os.getenv("DB_PASSWORD", "DSCs3cT00Ls2022"),
            "dsn": os.getenv("DB_DSN", "QASBRRACT2-SCAN.PHX-DC.DHL.COM:1521/DSCSECTOOLS"),
        }

    def chamar_funcao_importar_ler_excel(self):
        # Utilização da função que realiza o upload de arquivos para o sistema
        data = self.importar_ler_excel()
        if data:
            for sheet, df in data.items():
                print(f"Aba: {sheet}")
                print(df.head())  # Exibe as primeiras linhas de cada aba

    def importar_ler_excel(self):
        try:
            resposta = QMessageBox.question(
                self,
                "Confirmação",
                "Tem certeza de que deseja deletar os dados da tabela de faturamento anterior?",
                QMessageBox.Yes | QMessageBox.No,
            )

            if resposta == QMessageBox.Yes:
                # Query para limpar a tabela de faturamento antes de registrar
                creds = self.get_db_credentials()

                # Conectar ao banco de dados
                with oracledb.connect(**creds) as connection:
                    print("Conexão estabelecida.")
                    with connection.cursor() as cursor:
                        print("Executando a limpeza...")

                        # Comando para verificar se a tabela tem dados
                        check_query = "SELECT COUNT(*) FROM FATURAMENTO_ESCOLTA_BASE"
                        cursor.execute(check_query)
                        row_count = cursor.fetchone()[0]

                        print(row_count)

                        if row_count > 0:
                            # Comando DELETE
                            query = """
                                DELETE FROM FATURAMENTO_ESCOLTA_BASE
                            """
                            try:
                                cursor.execute(query)  # Executa a query para deletar os dados
                                connection.commit()  # Faz commit para garantir que as alterações sejam salvas
                                print("Query executada com sucesso.")

                                # Usar cursor.rowcount para contar o número de linhas afetadas
                                deletador_contador = cursor.rowcount
                                print(f"Dados deletados da tabela de faturamento: {deletador_contador}")

                                # Exibe a quantidade de registros deletados
                                QMessageBox.information(
                                    self,  # widget pai
                                    "Confirmação",  # Título da janela
                                    f"Dados deletados da tabela de faturamento: {deletador_contador}"  # Mensagem
                                )
                            except oracledb.Error as e:
                                print(f"Erro ao executar a query: {e}")
                                QMessageBox.warning(
                                    self,
                                    "Erro",
                                    f"Erro ao executar a query: {e}",
                                )
                                return None
                        else:
                            print("Tabela já está vazia. Nenhum dado foi deletado.")

                        # Abre o diálogo para selecionar o arquivo
                        Tk().withdraw()  # Esconde a janela principal do Tkinter
                        file_path = filedialog.askopenfilename(title="Selecione a planilha Excel",
                                                               filetypes=[("Excel Files", "*.xlsx;*.xls")])

                        if not file_path:
                            print("Nenhum arquivo selecionado.")
                            return None

                # Lê a planilha inteira
                df = pd.read_excel(file_path, sheet_name=None)  # Lê todas as abas da planilha

                # Chama a função para registrar os dados no banco após a limpeza
                self.chamar_registrar_excel_banco(file_path)

                return df

        except Exception as e:
            print(f"Erro ao processar a requisição: {e}")
            return None

    def chamar_registrar_excel_banco(self, file_path):
        """Inicia a execução da consulta em uma thread separada."""

        # Mostrar a label e iniciar o GIF
        self.ui.label_loading.setVisible(True)
        self.gif.start()  # Inicia o GIF

        # Criar uma thread para executar a consulta
        self.thread = threading.Thread(target=self.registrar_excel_banco, args=(file_path,))
        self.thread.start()

    # Função que realizar o registro das informações da planilha no banco de dados oracle.
    def registrar_excel_banco(self, file_path):

        try:
            creds = self.get_db_credentials()
            df = pd.read_excel(file_path, dtype=str)  # Lê como string para evitar problemas

            required_columns = [
                "N_SE", "CLIENTES", "DESCRICAO_SE", "EMPRESA_ESCOLTA", "SITUACAO", "COBERTURA",
                "AGENDAMENTO_DATA_HORA", "CHEGADA_ORIGEM_DATA_HORA", "CHEGADA_DESTINO_DATA_HORA",
                "FIM_EMISSAO_REAL", "FRANQUIA_HORAS", "VALOR_HORA_EXCEDENTE", "DISTANCIA_REAL",
                "FRANQUIA_KM", "VALOR_KM_EXCEDENTE", "PRECO_FRANQUIA_BASE", "VALOR_TOTAL_EMISSAO",
                "STATUS_PAGAMENTO", "VIAGEM_CONSIDERADA", "STATUS"
            ]

            # Verifica se todas as colunas necessárias estão presentes
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Colunas ausentes na planilha: {missing_cols}")

            with oracledb.connect(**creds) as connection:
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

                    total_registros = len(df)
                    if total_registros == 0:
                        raise ValueError("A planilha está vazia!")

                    # Substitui valores nulos por string vazia e converte para lista de tuplas
                    data = [tuple(row.fillna("").values) for _, row in df.iterrows()]

                    # Inserção em lote
                    cursor.executemany(query, data)
                    connection.commit()

                    print(f"Registros inseridos: {total_registros}")

        except oracledb.DatabaseError as e:
            print(f"Erro de banco de dados: {e}")
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha na conexão ou na inserção: {e}")

        except Exception as e:
            print(f"Erro inesperado: {e}")
            print(traceback.format_exc())  # Exibe o erro completo no console
            QtWidgets.QMessageBox.critical(self, "Erro inesperado", f"Ocorreu um erro inesperado: {e}")

        finally:
            if hasattr(self, 'gif') and self.gif:
                try:
                    self.sinal_finalizacao.emit(total_registros)
                except RuntimeError:
                    pass  # Ignora erro caso o objeto já tenha sido destruído

    # Função que finaliza a thread para usar os parametros sem erro
    def finalizarProcessamento(self, total_registros):
        """Atualiza a UI e exibe a mensagem na thread principal."""
        self.gif.stop()
        self.ui.label_loading.setVisible(False)
        self.thread = None  # Limpa a referência da thread

        QMessageBox.information(self, "Sucesso", f"Dados registrados com sucesso! Total: {total_registros}")

        self.consultar_viagem_benner()
    def consultar_viagem_benner(self):
        try:
            creds = self.get_db_credentials()

            with oracledb.connect(**creds) as connection:
                with connection.cursor() as cursor:
                    print("Executando a consulta...")

                    query = """SELECT VIAGEM_CONSIDERADA 
                               FROM FATURAMENTO_ESCOLTA_BASE 
                               WHERE VIAGEM_CONSIDERADA IS NOT NULL 
                               AND REGEXP_LIKE(VIAGEM_CONSIDERADA, '^\\d{4}/')"""

                    # Executa a consulta
                    cursor.execute(query)

                    # Recupera os resultados
                    resultados = cursor.fetchall()

                    # Exibir os resultados no console
                    for linha in resultados:
                        print(linha[0])  # Exibe cada resultado

                    # Contar quantas linhas foram retornadas
                    num_linhas = len(resultados)
                    print(f"Número de linhas retornadas: {num_linhas}")

                    return resultados, num_linhas  # Retorna os dados e a contagem

        except oracledb.DatabaseError as e:
            print(f"Erro de banco de dados: {e}")
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha na conexão ou 2 consulta dos numeros de viagem benner: {e}")

        except Exception as e:
            print(f"Erro inesperado: {e}")
            print(traceback.format_exc())  # Exibe o erro completo no console
            QtWidgets.QMessageBox.critical(self, "Erro inesperado", f"Ocorreu um erro inesperado  consulta dos numeros de viagem benner: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
1