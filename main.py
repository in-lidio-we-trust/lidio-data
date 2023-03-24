import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QFileDialog,
    QMessageBox,
    QAction,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QHBoxLayout,
    QPushButton,
)
import pandas as pd


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("LIDIO: Lovingly Improved Data Input Organizer.")

        # Criação da barra de menu
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Arquivo")

        # Adição do item de abrir arquivo ao menu
        open_file = QAction("Abrir arquivo", self)
        open_file.setShortcut("Ctrl+O")
        open_file.setStatusTip("Abrir arquivo")
        open_file.triggered.connect(self.openFile)
        file_menu.addAction(open_file)

        # Adição da tabela à janela principal
        self.tableWidget = QTableWidget()
        self.setCentralWidget(self.tableWidget)

        # Adição do botão de padronizar dados
        padronizar_btn = QPushButton("Padronizar dados")
        padronizar_btn.clicked.connect(self.showOptions)
        self.padronizar_options = QHBoxLayout()
        self.padronizar_options.addWidget(padronizar_btn)

        # Adição do botão de download
        download_btn = QPushButton("Salvar tabela")
        download_btn.clicked.connect(self.downloadTable)
        self.download_layout = QHBoxLayout()
        self.download_layout.addWidget(download_btn)

        # Adição do layout vertical para organizar os widgets
        layout = QVBoxLayout()
        layout.addLayout(self.padronizar_options)
        layout.addWidget(self.tableWidget)
        layout.addLayout(self.download_layout)

        # Criação do widget principal e definição do layout
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        # Ajuste do tamanho da janela
        self.resize(800, 600)
        self.show()

    def openFile(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Abrir arquivo", "", "CSV (*.csv);;XLSX (*.xlsx);;XLS (*.xls)")
        if filename:
            try:
                if filename.endswith(".csv"):
                    self.df = pd.read_csv(filename)
                elif filename.endswith(".xlsx"):
                    self.df = pd.read_excel(filename)
                elif filename.endswith(".xls"):
                    self.df = pd.read_excel(filename)
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))
            else:
                self.loadTable()

    def loadTable(self):
        self.tableWidget.setRowCount(len(self.df.index))
        self.tableWidget.setColumnCount(len(self.df.columns))
        self.tableWidget.setHorizontalHeaderLabels(list(self.df.columns))

        for row in range(len(self.df.index)):
            for col in range(len(self.df.columns)):
                self.tableWidget.setItem(row, col, QTableWidgetItem(str(self.df.iloc[row, col]))) 

    def showOptions(self):
        self.padronizar_options.itemAt(0).widget().deleteLater()  # Remove o botão "Padronizar Dados"
        for i in reversed(range(self.padronizar_options.count())):  # Remove as opções antigas
            self.padronizar_options.itemAt(i).widget().setParent(None)

        self.padronizar_options.addWidget(QPushButton("Opção 1"))
        self.padronizar_options.addWidget(QPushButton("Unithal"))
        self.padronizar_options.addWidget(QPushButton("Opção 3"))
        
        self.padronizar_options.itemAt(1).widget().clicked.connect(self.option1)

    def downloadTable(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Salvar tabela", "", "XLSX (*.xlsx)")

        if filename:
            try:
                self.df.to_excel(filename, index=False, engine='xlsxwriter')
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))
            else:
                QMessageBox.information(self, "Sucesso", "Tabela salva com sucesso!")
                
    def option1(self):
        # Renomeando colunas
        self.df = self.df.rename(columns={'IDENTIF_1': 'id', 'IDENTIF_2': 'prof', 'AL': 'Al', 'AR_FINA': 'Areia fina',
                        'AR_GROSSA': 'Areia grossa', 'ARGILA': 'Argila', 'CA': 'Ca', 'CA_MG': 'Ca+Mg',
                        'CL': 'Cl', 'CU': 'Cu', 'FE': 'Fe', 'H_AL': 'H/Al', 'MG': 'Mg', 'MN': 'Mn',
                        'M_ORG': 'MOS', 'NA': 'Na', 'PH_H2O': 'pH Agua', 'PH_CACL2': 'pH CaCl2',
                        'PH_SMP': 'ph_smp', 'P_MEL': 'P mehl', 'P_REM': 'prem', 'P_RES': 'P res',
                        'P_TOTAL': 'P_Total', 'AL_CTC': 'Al%', 'CA_CTC': 'Ca%', 'H_CTC': 'H%',
                        'MG_CTC': 'Mg%', 'SILTE': 'Silte', 'V': 'V%', 'ZN': 'Zn', 'K_CTC': 'K%'})
        
        # Adicionando colunas
        self.df['AlS'] = ''
        self.df['Altimetria SRTM'] = ''
        self.df['Areia total'] = ''
        self.df['C'] = ''
        self.df['CaS'] = ''
        self.df['CEa'] = ''
        self.df['CO3'] = ''
        self.df['Ds'] = ''
        self.df['Fe/Mn'] = ''
        self.df['FeO'] = ''
        self.df['H'] = ''
        self.df['HCO3'] = ''
        self.df['K mg'] = ''
        self.df['K/Na'] = ''
        self.df['KS'] = ''
        self.df['m%'] = ''
        self.df['MgS'] = ''
        self.df['NaS'] = ''
        self.df['NH4'] = ''
        self.df['NO3'] = ''
        self.df['P'] = ''
        self.df['PB'] = ''
        self.df['pH'] = ''
        self.df['ph_kcl'] = ''
        self.df['PO'] = ''
        self.df['P/Zn'] = ''
        self.df['RAS'] = ''
        self.df['Ca/K'] = ''
        self.df['Ca/Mg'] = ''
        self.df['Ca+Mg/K'] = ''
        self.df['Mg/K'] = ''
        self.df['H/Al%'] = ''
        self.df['Na%'] = ''
        self.df['Si'] = ''
        self.df['SO4'] = ''
        self.df['S/P'] = ''
        self.df['t'] = ''
        
        # Padronizando posicoes das colunas
        self.df = self.df.loc[:, ['id', 'prof', 'Al', 'AlS', 'Altimetria SRTM', 'Areia fina', 'Areia grossa',
                'Areia total', 'Argila', 'B', 'C', 'Ca', 'Ca+Mg', 'CaS', 'CEa', 'Cl', 'CO3',
                'CTC', 'Cu', 'Ds', 'Fe', 'Fe/Mn', 'FeO', 'H', 'H/Al', 'HCO3', 'K', 'K mg', 'K/Na',
                'KS', 'm%', 'Mg', 'MgS', 'Mn', 'MOS', 'N', 'Na', 'NaS', 'NH4', 'NO3', 'P', 'PB', 'pH', 
                'pH Agua', 'pH CaCl2', 'ph_kcl', 'ph_smp', 'P mehl', 'PO', 'prem', 'P res', 'P_Total',
                'P/Zn', 'RAS', 'Ca/K', 'Ca/Mg', 'Ca+Mg/K', 'Mg/K', 'S', 'Al%', 'Ca%', 'H%', 'H/Al%',
                'K%', 'Mg%', 'Na%', 'SB', 'Si', 'Silte', 'SO4', 'S/P', 't', 'V%', 'Zn']]
        
        # Removendo items indesejaveis
        self.df['prof'] = self.df['prof'].str.replace('PONTO ', '')
        self.df['id'] = self.df['id'].str.replace('(', '', regex=True)
        self.df['id'] = self.df['id'].str.replace(')', '', regex=True)

        # Atualize a tabela
        self.loadTable()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())
