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

        # Creation of the menu bar
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Arquivo")

        # Adding the file open item to the menu
        open_file = QAction("Abrir arquivo", self)
        open_file.setShortcut("Ctrl+O")
        open_file.setStatusTip("Abrir arquivo")
        open_file.triggered.connect(self.openFile)
        file_menu.addAction(open_file)

        # Adding the table to the main window
        self.tableWidget = QTableWidget()
        self.setCentralWidget(self.tableWidget)

        # Adding the standardize data button 
        padronizar_btn = QPushButton("Padronizar dados")
        padronizar_btn.clicked.connect(self.showOptions)
        self.padronizar_options = QHBoxLayout()
        self.padronizar_options.addWidget(padronizar_btn)

        # Adding the download button 
        download_btn = QPushButton("Salvar tabela")
        download_btn.clicked.connect(self.downloadTable)
        self.download_layout = QHBoxLayout()
        self.download_layout.addWidget(download_btn)

        # Creating the main widget and setting the layout 
        layout = QVBoxLayout()
        layout.addLayout(self.padronizar_options)
        layout.addWidget(self.tableWidget)
        layout.addLayout(self.download_layout)

        # Creating the main widget and setting the layout. 
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        # Adjusting the window size. 
        self.resize(800, 600)
        self.show()
        
        # Opening the file when the window opens. 
        self.openFile()

    def openFile(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Abrir arquivo", "", "XLSX (*.xlsx);;CSV (*.csv);;XLS (*.xls)")
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
        self.padronizar_options.itemAt(0).widget().deleteLater()  # Removing the "Standardize Data" button. 
        for i in reversed(range(self.padronizar_options.count())):  # Removing old options. 
            self.padronizar_options.itemAt(i).widget().setParent(None)

        self.padronizar_options.addWidget(QPushButton("A fazer"))
        self.padronizar_options.addWidget(QPushButton("Unithal"))
        self.padronizar_options.addWidget(QPushButton("A fazer 2"))
        
        self.padronizar_options.itemAt(1).widget().clicked.connect(self.option1)

    def downloadTable(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Salvar tabela", "", "XLSX (*.xlsx)")

        if filename:
            try:
                if not filename.endswith('.xlsx'):
                    filename += '.xlsx'  # Adds the .xlsx extension if it is not present. 
                    self.df.to_excel(filename, index=False, engine='xlsxwriter')
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))
        else:
            QMessageBox.information(self, "Sucesso", "Tabela salva com sucesso!")

    def option1(self):
        # Renomeando colunas
        self.df = self.df.rename(columns={'IDENTIF_2': 'id', 'IDENTIF_1': 'prof', 'AL': 'Al', 'AR_FINA': 'Areia fina',
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
        
        # Standardizing the positions of the columns. 
        self.df = self.df.loc[:, ['id', 'prof', 'Al', 'AlS', 'Altimetria SRTM', 'Areia fina', 'Areia grossa',
                'Areia total', 'Argila', 'B', 'C', 'Ca', 'Ca+Mg', 'CaS', 'CEa', 'Cl', 'CO3',
                'CTC', 'Cu', 'Ds', 'Fe', 'Fe/Mn', 'FeO', 'H', 'H/Al', 'HCO3', 'K', 'K mg', 'K/Na',
                'KS', 'm%', 'Mg', 'MgS', 'Mn', 'MOS', 'N', 'Na', 'NaS', 'NH4', 'NO3', 'P', 'PB', 'pH', 
                'pH Agua', 'pH CaCl2', 'ph_kcl', 'ph_smp', 'P mehl', 'PO', 'prem', 'P res', 'P_Total',
                'P/Zn', 'RAS', 'Ca/K', 'Ca/Mg', 'Ca+Mg/K', 'Mg/K', 'S', 'Al%', 'Ca%', 'H%', 'H/Al%',
                'K%', 'Mg%', 'Na%', 'SB', 'Si', 'Silte', 'SO4', 'S/P', 't', 'V%', 'Zn']]
        
        # Removing unwanted items. 
        self.df['id'] = self.df['id'].str.replace('PONTO ', '')
        self.df['prof'] = self.df['prof'].str.replace('(', '')
        self.df['prof'] = self.df['prof'].str.replace(')', '')
        self.df = self.insert_row(0, self.df, ['', '', 'cmolc/dm³', 'meq/L', 'metros', '%', 
                                               '%', '%', '%', 'ppm', 'cmolc/dm³', 'cmolc/dm³', 'cmolc/dm³',
                                               'meq/L', 'dS/m', 'ppm', 'meq/L', 'cmolc/dm³', 'ppm', 'g/dm³',
                                               'ppm', 'Sem Unidade', '%', 'cmolc/dm³', 'cmolc/dm³', 'meq/L',
                                               'cmolc/dm³', 'ppm', 'Sem Unidade', 'meq/L', '%',	'cmolc/dm³',
                                               'meq/L',	'ppm', '%',	'cmolc/dm³', 'cmolc/dm³', 'meq/L', 
                                               'meq/L',	'meq/L', 'ppm', 'ppm', 'Sem Unidade', 'Sem Unidade',
                                               'Sem Unidade', 'Sem Unidade', 'Sem Unidade', 'ppm','ppm','ppm'
                                               ,'ppm','ppm', 'Sem Unidade', 'Sem Unidade', 'Sem Unidade', 
                                               'Sem Unidade', 'Sem Unidade', 'Sem Unidade', 'ppm', '%', '%',
                                               '%', '%', '%', '%', '%', 'cmolc/dm³', '%', '%', 'meq/L', 'Sem Unidade',
                                               'cmolc/dm³', '%', 'ppm'])
        # Updating the table. 
        self.loadTable()

# Function to insert row in the dataframe
    def insert_row(self, row_number, df, row_value):
        # Starting value of upper half
        start_upper = 0

        # End value of upper half
        end_upper = row_number

        # Start value of lower half
        start_lower = row_number

        # End value of lower half
        end_lower = df.shape[0]

        # Create a list of upper_half index
        upper_half = [*range(start_upper, end_upper, 1)]

        # Create a list of lower_half index
        lower_half = [*range(start_lower, end_lower, 1)]

        # Increment the value of lower half by 1
        lower_half = [x.__add__(1) for x in lower_half]

        # Combine the two lists
        index_ = upper_half + lower_half

        # Update the index of the dataframe
        df.index = index_

        # Insert a row at the end
        df.loc[row_number] = row_value

        # Sort the index labels
        df = df.sort_index()

        # return the dataframe
        return df


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())
