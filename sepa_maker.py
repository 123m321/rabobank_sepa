import sys
import datetime
import pandas as pd
import random
import json
import os
from PyQt6.QtWidgets import QFrame, QApplication, QComboBox, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QDialog, QTableView, QHeaderView
from PyQt6.QtCore import QAbstractTableModel, Qt

'''
nog te doen, hoe om te gaan als de excel file niet zo strak is als ik verwacht. Dus nan waardes, en zero bedragen. En extra regels of kolommen waar ik 
helemaal niets mee doet. Daarna is het langzaam tijd om het te uploaden op github. Oh ja, ook nog bic check of partij die betaalt laten draaien.
'''

class DataFrameModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid() and role == Qt.ItemDataRole.DisplayRole:
            return str(self._df.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self._df.columns[section]
            if orientation == Qt.Orientation.Vertical:
                return str(self._df.index[section])
        return None

from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QComboBox, 
                           QLineEdit, QPushButton, QMessageBox, QFrame)
import json


def iban_check(iban_input):
    iban = iban_input.replace(' ', '')
    if not iban.isalnum():
        raise ValueError("Invalid characters inside")
    elif len(iban) < 15:
        raise ValueError("IBAN too short")
    elif len(iban) > 31:
        raise ValueError("IBAN too long")
    else:
        iban = (iban[4:] + iban[0:4]).upper()
        iban2 = ''
        for char in iban:
            if char.isdigit():
                iban2 += char
            else:
                iban2 += str(10 + ord(char) - ord('A'))
        iban3 = int(iban2)
        if not iban3 % 97 == 1:
            raise ValueError(f"Invalid IBAN checksum for {iban_input}")

class SEPAApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('SEPA Maker')
        self.main_layout = QVBoxLayout()

        # Laad configuratiebestand
        self.load_config()

        self.betalende_gegevens = self.config['betalende_gegevens']

        # Betalende naam dropdown
        self.betalende_naam_label = QLabel('Betalende Naam:')
        self.betalende_naam_input = QComboBox()
        self.update_dropdown()
        self.main_layout.addWidget(self.betalende_naam_label)
        self.main_layout.addWidget(self.betalende_naam_input)

        # Betalende IBAN veld
        self.betalende_iban_label = QLabel('Betalende IBAN:')
        self.betalende_iban_input = QLineEdit()
        self.betalende_iban_input.setReadOnly(True)
        self.main_layout.addWidget(self.betalende_iban_label)
        self.main_layout.addWidget(self.betalende_iban_input)

        # Update IBAN wanneer naam verandert
        self.betalende_naam_input.currentIndexChanged.connect(self.update_iban)
        self.update_iban()  # Initialize with first value

        # Config bewerken knop
        self.edit_config_button = QPushButton('Config Bewerken')
        self.edit_config_button.clicked.connect(self.toggle_config_editor)
        self.main_layout.addWidget(self.edit_config_button)

        # Container voor config editor (initially hidden)
        self.config_editor_container = QWidget()
        self.config_editor_layout = QVBoxLayout()
        self.config_editor_container.setLayout(self.config_editor_layout)
        self.config_editor_container.hide()

        # Config editor widgets voor toevoegen
        self.add_section_label = QLabel('Nieuwe gegevens toevoegen:')
        self.config_editor_layout.addWidget(self.add_section_label)

        # Nieuwe naam input
        self.new_naam_label = QLabel('Nieuwe naam:')
        self.new_naam_input = QLineEdit()
        self.config_editor_layout.addWidget(self.new_naam_label)
        self.config_editor_layout.addWidget(self.new_naam_input)

        # Nieuwe IBAN input
        self.new_iban_label = QLabel('Nieuwe IBAN:')
        self.new_iban_input = QLineEdit()
        self.config_editor_layout.addWidget(self.new_iban_label)
        self.config_editor_layout.addWidget(self.new_iban_input)

        # Toevoegen knop
        self.add_button = QPushButton('Toevoegen aan config')
        self.add_button.clicked.connect(self.add_to_config)
        self.config_editor_layout.addWidget(self.add_button)

        # Horizontale lijn voor visuele scheiding
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        self.config_editor_layout.addWidget(separator)

        # Config editor widgets voor verwijderen
        self.delete_section_label = QLabel('Gegevens verwijderen:')
        self.config_editor_layout.addWidget(self.delete_section_label)

        # Dropdown voor te verwijderen item
        self.delete_combo = QComboBox()
        self.update_delete_dropdown()
        self.config_editor_layout.addWidget(QLabel('Selecteer te verwijderen naam:'))
        self.config_editor_layout.addWidget(self.delete_combo)

        # Verwijder knop
        self.delete_button = QPushButton('Verwijder geselecteerde naam')
        self.delete_button.clicked.connect(self.delete_from_config)
        self.config_editor_layout.addWidget(self.delete_button)

        # Voeg config editor container toe aan main layout
        self.main_layout.addWidget(self.config_editor_container)

        # Start knop
        self.start_button = QPushButton('Start')
        self.start_button.clicked.connect(self.start_process)
        self.main_layout.addWidget(self.start_button)

        self.setLayout(self.main_layout)

    def toggle_config_editor(self):
        if self.config_editor_container.isHidden():
            self.config_editor_container.show()
            self.edit_config_button.setText('Verberg Config Editor')
            # Update delete dropdown when showing
            self.update_delete_dropdown()
        else:
            self.config_editor_container.hide()
            self.edit_config_button.setText('Config Bewerken')
            # Clear input fields when hiding
            self.new_naam_input.clear()
            self.new_iban_input.clear()

    def load_config(self):
        with open('config.json', 'r') as f:
            self.config = json.load(f)

    def save_config(self):
        with open('config.json', 'w') as f:
            json.dump(self.config, f, indent=4)

    def update_dropdown(self):
        self.betalende_naam_input.clear()
        for item in self.betalende_gegevens:
            self.betalende_naam_input.addItem(item['naam'])

    def update_delete_dropdown(self):
        self.delete_combo.clear()
        for item in self.betalende_gegevens:
            self.delete_combo.addItem(item['naam'])

    def update_iban(self):
        current_index = self.betalende_naam_input.currentIndex()
        if current_index >= 0:
            self.betalende_iban_input.setText(self.betalende_gegevens[current_index]['iban'])

    def add_to_config(self):
        nieuwe_naam = self.new_naam_input.text().strip()
        nieuwe_iban = self.new_iban_input.text().strip()

        if not nieuwe_naam or not nieuwe_iban:
            QMessageBox.warning(self, 'Fout', 'Vul beide velden in.')
            return

        # Controleer of de naam al bestaat
        if any(item['naam'] == nieuwe_naam for item in self.betalende_gegevens):
            QMessageBox.warning(self, 'Fout', 'Deze naam bestaat al.')
            return

        # Voeg nieuwe gegevens toe
        new_entry = {
            "naam": nieuwe_naam,
            "iban": nieuwe_iban
        }
        self.betalende_gegevens.append(new_entry)
        
        # Update config file
        self.save_config()
        
        # Update dropdowns
        self.update_dropdown()
        self.update_delete_dropdown()
        
        # Clear input fields
        self.new_naam_input.clear()
        self.new_iban_input.clear()
        
        QMessageBox.information(self, 'Succes', 'Nieuwe gegevens toegevoegd.')

    def delete_from_config(self):
        te_verwijderen_naam = self.delete_combo.currentText()
        
        # Bevestiging vragen
        reply = QMessageBox.question(self, 'Bevestiging', 
                                   f'Weet je zeker dat je "{te_verwijderen_naam}" wilt verwijderen?',
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                   QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # Verwijder het geselecteerde item
            self.betalende_gegevens = [item for item in self.betalende_gegevens 
                                     if item['naam'] != te_verwijderen_naam]
            self.config['betalende_gegevens'] = self.betalende_gegevens
            
            # Update config file
            self.save_config()
            
            # Update dropdowns
            self.update_dropdown()
            self.update_delete_dropdown()
            
            # Update IBAN display
            self.update_iban()
            
            QMessageBox.information(self, 'Succes', f'"{te_verwijderen_naam}" is verwijderd.')

    def bic_vinden(self, iban):
        bic_dict = {
            'INGB': 'INGBNL2A',
            'RABO': 'RABONL2U',
            'KNAB': 'KNABNL2H',
            'ABNA': 'ABNANL2A',
            'BUNQ': 'BUNQNL2A',
            'SNSB': 'SNSBNL2A',
            'TRIO': 'TRIONL2U',
            'ASNB': 'ASNBNL21',
            'RBRB': 'RBRBNL21'
        }   
        bankcode = iban[4:8]
        return bic_dict.get(bankcode, 'UNKNOWN')

    def betaling_toevoegen(self, datum_met_streepje, batchnaam, bedrag, bestemming_naam, iban, omschrijving):
        with open("betaling.xml", "r") as f:
            betaling = f.read()
            betaling = betaling.replace("{datumyyy-mm-dd}", datum_met_streepje)
            betaling = betaling.replace("{naam}", batchnaam)
            betaling = betaling.replace("{bedrag}", bedrag)
            betaling = betaling.replace("{bestemming_naam}", bestemming_naam)
            try:
                iban_check(iban)
            except ValueError as e:
                QMessageBox.critical(self, f"Ongeldige IBAN: {iban}", str(e))
                return
            #iban_check(iban)
            betaling = betaling.replace("{iban}", iban)
            bic = self.bic_vinden(iban)
            betaling = betaling.replace("{bic}", bic)
            betaling = betaling.replace("{omschrijving}", omschrijving)
        return betaling

    def random_letters(self):
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        return random.choice(letters) + random.choice(letters)

    def start_process(self):
        betaalbestand, _ = QFileDialog.getOpenFileName(self, "Selecteer XLSX bestand", "", "Excel files (*.xlsx)")
        if not betaalbestand:
            return

        try:
            df = pd.read_excel(betaalbestand)
            
            # Validate all IBANs before processing
            for idx, iban in enumerate(df['IBAN']):
                try:
                    iban_check(str(iban))
                except ValueError as e:
                    QMessageBox.critical(self, "Error", f"Ongeldige IBAN op rij {idx + 1}: {iban}\nFout: {str(e)}")
                    return
            
            # Continue with rest of the processing only if all IBANs are valid
            betalende_naam = self.betalende_naam_input.currentText()
            betalende_iban = self.betalende_iban_input.text()
            
            try:
                iban_check(betalende_iban)
            except ValueError as e:
                QMessageBox.critical(self, "Error", f"Ongeldige IBAN van betalende partij: {betalende_iban}\nFout: {str(e)}")
                return

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Kon het Excel bestand niet verwerken: {str(e)}")
            return

        code = self.random_letters()

        datum_met_streepje = datetime.datetime.now().strftime("%Y-%m-%d")
        datum_vast = datetime.datetime.now().strftime("%Y%m%d")
        batchnaam = f"MM-{datetime.datetime.now().day}-{code}"

        aantal = str(df.shape[0])
        totaal_bedrag = str(round(df['Bedrag'].sum(), 2))     

        try:
            with open("example.xml", "r", encoding="utf-8") as f:
                text = f.read()
        except UnicodeDecodeError:
            with open("example.xml", "r", encoding="latin-1") as f:
                text = f.read()

        # Basis vervanging
        replaced_text = text.replace("{datumyyy-mm-dd}", datum_met_streepje)
        replaced_text = replaced_text.replace("{datum_yyyymmdd}", datum_vast)
        replaced_text = replaced_text.replace("{naam}", batchnaam)
        replaced_text = replaced_text.replace("{aantal}", aantal)
        replaced_text = replaced_text.replace("{totaal_bedrag}", totaal_bedrag)
        replaced_text = replaced_text.replace("{betalende_naam}", betalende_naam)
        replaced_text = replaced_text.replace("{betalende_iban}", betalende_iban)

        # Eerste betaling
        if len(df) > 0:
            replaced_text = replaced_text.replace("{bedrag}", str(df['Bedrag'][0]))
            replaced_text = replaced_text.replace("{bestemming_naam}", str(df['Naam'][0]))
            replaced_text = replaced_text.replace("{iban}", str(df['IBAN'][0]))
            replaced_text = replaced_text.replace("{bic}", self.bic_vinden(str(df['IBAN'][0])))
            replaced_text = replaced_text.replace("{omschrijving}", str(df['Omschrijving'][0]))

        # Extra betalingen
        betalingen = ""
        for index in range(1, len(df)):
            betaling = self.betaling_toevoegen(
                datum_met_streepje,
                batchnaam,
                str(df['Bedrag'][index]),
                str(df['Naam'][index]),
                str(df['IBAN'][index]),
                str(df['Omschrijving'][index])
            )
            betalingen += betaling

        replaced_text = replaced_text.replace("</CdtTrfTxInf>", "</CdtTrfTxInf>\n" + betalingen)
        
        #show dataframe only with the columns we need : naam, bedrag, iban, omschrijving
        df = df[['Naam', 'Bedrag', 'IBAN', 'Omschrijving']]

        self.show_dataframe_popup(df.iloc[:, 0:4], betalende_naam)

        #ask for user to check the df and give them an option to abort the process
        reply = QMessageBox.question(self, 'Check Dataframe', 'Check the dataframe and press OK to continue', QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Abort) 
        if reply == QMessageBox.StandardButton.Abort:
            sys.exit()
        
        os.makedirs("data", exist_ok=True)

        save_file = f"data/sepa_{code}{datum_vast}.xml"
        with open(save_file, "w", encoding="utf-8") as f:
            f.write(replaced_text)

        QMessageBox.information(self, "Succes", f"{save_file}")
        #sys.exit()

    def show_dataframe_popup(self, df, betalende_naam):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{betalende_naam}")
        layout = QVBoxLayout()

        table_view = QTableView()
        model = DataFrameModel(df)
        table_view.setModel(model)
        table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(table_view)
        dialog.setLayout(layout)
        dialog.resize(800, 600)  # Maak het popup-venster groter
        dialog.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SEPAApp()
    ex.show()
    app.exec()
    #sys.exit(app.exec())
