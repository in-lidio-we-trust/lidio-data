import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QVBoxLayout,
    QInputDialog,
    QMessageBox,
)


class ExcelEditor(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Define nome da janela
        self.setWindowTitle("LIDIO: Lovingly Improved Data Input Organizer.")
        self.setGeometry(100, 100, 800, 600)

        # Botão para selecionar o arquivo XLSX
        self.btn_select_file = QPushButton("Selecionar arquivo", self)
        self.btn_select_file.clicked.connect(self.select_file)

        # Tabela para mostrar os dados do arquivo XLSX
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)

        # Botão para salvar o arquivo XLSX com o novo nome de coluna
        self.btn_save_file = QPushButton("Salvar arquivo", self)
        self.btn_save_file.clicked.connect(self.save_file)
        self.btn_save_file.setEnabled(False)

        # Botão para renomear as colunas
        self.btn_rename_columns = QPushButton("Renomear colunas", self)
        self.btn_rename_columns.clicked.connect(self.rename_columns)
        self.btn_rename_columns.setEnabled(False)

        # Adicionando os widgets à janela principal
        layout = QVBoxLayout(self)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.table_widget)
        layout.addWidget(self.btn_rename_columns)
        layout.addWidget(self.btn_save_file)

    def select_file(self):
        # Abrir a janela de seleção de arquivo XLSX
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            # Carregar o arquivo XLSX na tabela
            df = pd.read_excel(file_path)
            self.table_widget.setRowCount(len(df.index))
            self.table_widget.setColumnCount(len(df.columns))
            self.table_widget.setHorizontalHeaderLabels(df.columns)

            for row in range(len(df.index)):
                for col in range(len(df.columns)):
                    item = QTableWidgetItem(str(df.iat[row, col]))
                    self.table_widget.setItem(row, col, item)

            # Habilitar o botão de salvar
            self.btn_save_file.setEnabled(True)
        self.btn_rename_columns.setEnabled(True)

    def save_file(self):
        # Abrir a janela de seleção de local para salvar o arquivo
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar arquivo", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            # Salvar o arquivo XLSX com as novas colunas renomeadas
            df = pd.read_excel(self.file_path)
            new_column_names = [
                self.table_widget.horizontalHeaderItem(col).text()
                for col in range(self.table_widget.columnCount())
            ]
            df.columns = new_column_names
            df.to_excel(file_path, index=False)

    def rename_columns(self):
        num_cols = self.table_widget.columnCount()

        # Cria um diálogo com opções de renomeação
        options_dialog = QMessageBox(self)
        options_dialog.setIcon(QMessageBox.Information)
        options_dialog.setWindowTitle("Opções de Renomeação")
        options_dialog.setText("Selecione um laboratório:")
        options_dialog.addButton("Unithal", QMessageBox.AcceptRole)
        options_dialog.addButton("Lab forte", QMessageBox.AcceptRole)
        options_dialog.addButton("Agrisolum", QMessageBox.AcceptRole)

        # Exibe o diálogo e obtém a opção selecionada
        option = options_dialog.exec_()
        if option == 0:
            # Renomear para Nome Padrão
            new_names = [f"Coluna {i}" for i in range(1, num_cols + 1)]
        elif option == 1:
            # Renomear com Sufixo
            suffix, ok = QInputDialog.getText(
                self,
                "Renomear colunas",
                "Digite o sufixo para renomear as colunas:",
            )
            if not ok:
                # O usuário cancelou o diálogo, nada a fazer
                return
            new_names = [
                f"{self.table_widget.horizontalHeaderItem(col).text()}_{suffix}"
                for col in range(num_cols)
            ]
        else:
            # Renomear Manualmente
            new_names, ok = QInputDialog.getText(
                self,
                "Renomear colunas",
                "Digite os novos nomes de coluna, separados por vírgula:",
            )
            if not ok:
                # O usuário cancelou o diálogo, nada a fazer
                return
            new_names = new_names.split(",")

        if len(new_names) != num_cols:
            # O número de novos nomes não corresponde ao número de colunas na tabela
            error_dialog = QMessageBox(self)
            error_dialog.setIcon(QMessageBox.Warning)
            error_dialog.setWindowTitle("Erro")
            error_dialog.setText(
                f"O número de novos nomes ({len(new_names)}) não corresponde ao número de colunas na tabela ({num_cols})."
            )
            error_dialog.exec_()
            return

        # Renomeia as colunas na tabela
        for i, new_name in enumerate(new_names):
            self.table_widget.setHorizontalHeaderItem(i, QTableWidgetItem(new_name))

        # Desabilita o botão de renomear colunas (será habilitado novamente quando o usuário abrir outro arquivo)
        self.btn_rename_columns.setEnabled(False)

        # Habilita o botão de salvar
        self.btn_save_file.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    excel_editor = ExcelEditor()
    excel_editor.show()
    sys.exit(app.exec_())
