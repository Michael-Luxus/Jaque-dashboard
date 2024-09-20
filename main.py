import datetime
from datetime import datetime as dtweek
import json
import os
import smtplib
import sys

from email.message import EmailMessage

import qdarkstyle

from PyQt5.QtWidgets import (
    QWidget,
    QPushButton,
    QApplication,
    QListWidget,
    QLabel,
    QHBoxLayout,
    QLineEdit,
    QComboBox,
    QTableWidget,
    QVBoxLayout,
    QTableWidgetItem, QMessageBox, QTabWidget,
)
from PyQt5.QtCore import QTimer, QDateTime, Qt, QTime

from query import connexionSQlServer, getDataLink
from utils import get_today, write_log

def is_open_day():
    # Get the current day of the week
    today = dtweek.today().weekday()
    # Monday is 0 and Sunday is 6
    if today < 5:  # 0-4 are Monday to Friday
        return True
    else:
        return False

def load_json(file_source):
    with open(file_source, 'r', encoding='utf-8') as fileR:
        fileX = json.load(fileR)
    return fileX


# def load_config():
#     with open('config.json', 'r', encoding='utf-8') as file:
#         json_file = json.load(file)
#     return json_file


def modifier_objet(no, nouveau_statut):
    try:
        with open('historique.json', "r") as f:
            data = json.load(f)

        for obj in data:
            if int(obj["No"]) == int(no):
                obj["DateTime"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                obj["Status"] = nouveau_statut

        with open('historique.json', "w") as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        print(f"Erreur de Modification : {str(e)}")
        return False


def custom_send_email(dates, email_adress, attachment_filename, soc_name):
    json_file = load_json('config.json')

    message_text = json_file['CONFIGURATION']['MESSAGE']
    objet_text = json_file['CONFIGURATION']['OBJET']

    message_text = message_text.replace("{date}", dates.strftime('%d/%m/%Y %H:%M:%S'))

    message = EmailMessage()
    message["Subject"] = objet_text.replace("{name_societe}", soc_name).replace("{date}", dates.strftime('%d/%m/%Y'))
    message["From"] = json_file['CONFIGURATION']['FROM']
    message["To"] = ', '.join(email_adress.split(';'))
    # if cc:
    #     message["Cc"] = ', '.join(cc.split(';'))
    # else:
    message["Cc"] = json_file['CONFIGURATION']['CC']
    message.set_content(message_text)

    with open(attachment_filename, "rb") as file:
        content = file.read()
        attached_filename = os.path.basename(attachment_filename)
        message.add_attachment(content, maintype="application", subtype="octet-stream",
                               filename=attached_filename)
    smtp_conf = json_file['CONFIGURATION']

    with smtplib.SMTP(smtp_conf['HOST'], smtp_conf['PORT']) as server:
        server.starttls()
        server.login(smtp_conf['LOGIN'], smtp_conf['PASSWORD'])
        server.send_message(message)




class WinForm(QWidget):
    def __init__(self, parent=None):
        super(WinForm, self).__init__(parent)
        self.executeBtn = None
        self.timer = None
        self.table_historique = None
        self.validate_button = None
        self.add_row_button = None
        self.table_destinataire = None
        self.tab_historique = None
        self.tab_destinataire = None
        self.tab_widget = None
        self.endBtn = None
        self.startBtn = None
        self.label = None
        self.listFile = None
        self.day_combo = None
        self.day_label = None
        self.target_second_edit = None
        self.target_minute_edit = None
        self.target_hour_edit = None
        self.target_time_label = None
        self.setFixedSize(787, 700)
        self.setWindowTitle('JK Board')
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Recherche dans l'historique")  # Search in history

        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        self.setupUi()

        self.load_destination_from_json()
        self.load_historique_from_json()

        self.target_time = None
        # self.target_day = None
        self.data_destination = []
        self.data_historique = []

    def setupUi(self):
        self.target_time_label = QLabel('Heure:')
        self.target_hour_edit = QLineEdit('17')
        self.target_hour_edit.setFixedWidth(50)
        self.target_minute_edit = QLineEdit('55')
        self.target_minute_edit.setFixedWidth(50)
        self.target_second_edit = QLineEdit('0')
        self.target_second_edit.setFixedWidth(50)

        # self.day_label = QLabel('Jour:')
        # self.day_combo = QComboBox()
        #  self.day_combo.addItems(['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'])

        self.listFile = QListWidget()
        self.label = QLabel('Etat automatique de la liste des articles')
        self.label.setAlignment(Qt.AlignCenter)  # Align label text to center
        self.startBtn = QPushButton('Start')
        self.endBtn = QPushButton('Stop')
        self.executeBtn = QPushButton('Execute')

        # Création d'un QTabWidget pour gérer les onglets
        self.tab_widget = QTabWidget()

        # Création des onglets "Destinataire" et "Historique"
        self.tab_destinataire = QWidget()
        self.tab_historique = QWidget()

        self.tab_widget.addTab(self.tab_destinataire, "Destinataire")
        self.tab_widget.addTab(self.tab_historique, "Historique")

        # Tableau pour l'onglet "Destinataire"
        self.table_destinataire = QTableWidget()
        self.table_destinataire.setColumnCount(6)  # Increased to 5 for the new 'Status' column
        self.table_destinataire.setHorizontalHeaderLabels(['Id', 'Nom', 'Email', 'Status', 'Societe', 'Action'])
        self.table_destinataire.setColumnWidth(0, 20)
        self.table_destinataire.setColumnWidth(1, 150)
        self.table_destinataire.setColumnWidth(2, 250)
        self.table_destinataire.setColumnWidth(3, 100)
        self.table_destinataire.setColumnWidth(4, 150)
        self.table_destinataire.setColumnWidth(5, 60)

        self.add_row_button = QPushButton('Ajouter')
        self.add_row_button.clicked.connect(self.add_row_to_table)

        self.validate_button = QPushButton('Valider')
        self.validate_button.clicked.connect(self.save_destination_to_json)

        # Layout pour l'onglet "Destinataire"
        layout_destinataire = QVBoxLayout(self.tab_destinataire)
        layout_destinataire.addWidget(self.table_destinataire)  # Ajoutez le tableau à cet onglet
        layout_destinataire.addWidget(self.add_row_button)
        layout_destinataire.addWidget(self.validate_button)

        # Création d'un deuxième tableau pour l'onglet "Historique"
        self.table_historique = QTableWidget()
        self.table_historique.setColumnCount(10)  # Nombre de colonnes pour l'historique
        self.table_historique.setHorizontalHeaderLabels(['N°', 'Date', 'Destinataire', 'Email', 'Status', 'Action', 'Server', 'Base', 'Societe', 'filename'])
        self.table_historique.setColumnWidth(0, 5)
        self.table_historique.setColumnWidth(1, 115)
        self.table_historique.setColumnWidth(2, 100)
        self.table_historique.setColumnWidth(3, 225)
        self.table_historique.setColumnWidth(4, 50)
        self.table_historique.setColumnWidth(5, 60)

        self.table_historique.setColumnWidth(6, 60)
        self.table_historique.setColumnWidth(7, 60)
        self.table_historique.setColumnWidth(8, 60)
        self.table_historique.setColumnWidth(9, 60)
        layout_historique = QVBoxLayout(self.tab_historique)
        # layout_historique.addWidget(self.search_bar)
        layout_historique.addWidget(self.table_historique)

        layout = QVBoxLayout()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_countdown)

        # Target time input layout
        target_time_layout = QHBoxLayout()
        target_time_layout.addWidget(self.target_time_label)
        target_time_layout.addWidget(self.target_hour_edit)
        target_time_layout.addWidget(QLabel('Minute: '))
        target_time_layout.addWidget(self.target_minute_edit)
        target_time_layout.addWidget(QLabel('Second: '))
        target_time_layout.addWidget(self.target_second_edit)
        # target_time_layout.addWidget(self.day_label)
        # target_time_layout.addWidget(self.day_combo)
        target_time_layout.addWidget(self.startBtn)
        target_time_layout.addWidget(self.endBtn)
        target_time_layout.addWidget(self.executeBtn)

        layout.addLayout(target_time_layout)
        layout.addWidget(self.label)
        layout.addWidget(self.tab_widget)

        self.startBtn.clicked.connect(self.start_countdown)
        self.endBtn.clicked.connect(self.end_timer)
        self.executeBtn.clicked.connect(self.iter_destination_json)

        self.table_destinataire.setColumnHidden(0, True)
        self.table_historique.setColumnHidden(0, True)
        # self.search_bar.textChanged.connect(self.filter_historique_table)

        self.setLayout(layout)

    def filter_historique_table(self, search_text):
        self.table_historique.setRowCount(0)  # Clear the table before filtering

        for item in self.data_historique:
            if search_text.lower() in item['Destinataire'].lower() or search_text.lower() in item['Email'].lower():
                row_position = self.table_historique.rowCount()
                self.table_historique.insertRow(row_position)

                self.table_historique.setItem(row_position, 0, QTableWidgetItem(str(item['No'])))
                self.table_historique.setItem(row_position, 1, QTableWidgetItem(str(item['DateTime'])))
                self.table_historique.setItem(row_position, 2, QTableWidgetItem(str(item['Destinataire'])))
                self.table_historique.setItem(row_position, 3, QTableWidgetItem(str(item['Email'])))
                self.table_historique.setItem(row_position, 4,
                                              QTableWidgetItem("Success" if item['Status'] else "Echec"))
                if not item['Status']:
                    resend_button = QPushButton('^')
                    resend_button.clicked.connect(lambda _, row=row_position: self.resend_row_from_table(row))
                    self.table_historique.setCellWidget(row_position, 5, resend_button)

                self.table_historique.setItem(row_position, 6, QTableWidgetItem(str(item['server'])))
                self.table_historique.setItem(row_position, 7, QTableWidgetItem(str(item['base_name'])))
                self.table_historique.setItem(row_position, 8, QTableWidgetItem(str(item['soc_name'])))
                self.table_historique.setItem(row_position, 9, QTableWidgetItem(str(item['filename'])))
                # self.table_historique.setItem(row_position, 9, QTableWidgetItem(str(item['adress_cc'])))

    def filter_historique_table_realtime(self):
        current_search_text = self.search_bar.text()
        self.filter_historique_table(current_search_text)

    def add_row_to_table(self):
        row_position = self.table_destinataire.rowCount()
        self.table_destinataire.insertRow(row_position)

        id_item = QTableWidgetItem(str(row_position + 1))  # Définition de la clé 'Id'
        self.table_destinataire.setItem(row_position, 0, id_item)

        name_item = QTableWidgetItem('')
        self.table_destinataire.setItem(row_position, 1, name_item)

        email_item = QTableWidgetItem('')
        self.table_destinataire.setItem(row_position, 2, email_item)

        # Create a combo box for the Status column
        status_combo = QComboBox()
        status_combo.addItems(['Active', 'Inactive'])
        status_combo.setCurrentIndex(0)  # Set default to Active

        self.table_destinataire.setCellWidget(row_position, 3, status_combo)

        societe_json = load_json('societe.json')
        society_names = []
        for soc in societe_json:
            society_names.append(soc['nom'])
        society_combo = QComboBox()
        society_combo.addItems(society_names)  # Assuming self.society_names is defined elsewhere
        self.table_destinataire.setCellWidget(row_position, 4, society_combo)

        delete_button = QPushButton('X')
        delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
        self.table_destinataire.setCellWidget(row_position, 5, delete_button)

        # Update data_destination dictionary with the selected status
        self.data_destination.append({'Id': str(row_position + 1), 'Status': 'Active'})  # Default to Active

    # def add_row_to_table(self):
    #     row_position = self.table_destinataire.rowCount()
    #     self.table_destinataire.insertRow(row_position)

    #     id_item = QTableWidgetItem(str(row_position + 1))  # Définition de la clé 'Id'
    #     self.table_destinataire.setItem(row_position, 0, id_item)

    #     name_item = QTableWidgetItem('')
    #     self.table_destinataire.setItem(row_position, 1, name_item)

    #     email_item = QTableWidgetItem('')
    #     self.table_destinataire.setItem(row_position, 2, email_item)

    #     # Create a combo box for the Status column
    #     status_combo = QComboBox()
    #     status_combo.addItems(['Active', 'Inactive'])
    #     status_combo.setCurrentIndex(0)  # Set default to Active
    #     self.table_destinataire.setCellWidget(row_position, 3, status_combo)

    #     # Create a combo box for society names
    #     societe_json = load_json('societe.json')
    #     society_names = []
    #     for soc in societe_json:
    #         society_names.append(soc['nom'])
    #     society_combo = QComboBox()
    #     society_combo.addItems(self.society_names)  # Assuming self.society_names is defined elsewhere
    #     self.table_destinataire.setCellWidget(row_position, 4, society_combo)

    #     delete_button = QPushButton('X')
    #     delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
    #     self.table_destinataire.setCellWidget(row_position, 5, delete_button)

    #     # Update data_destination dictionary with the default selected status and empty society
    #     self.data_destination.append({'Id': str(row_position + 1), 'Status': 'Active', 'Societe': ''})  # Default to Active and empty society



    def delete_row_from_table(self, row):
        confirm_delete = QMessageBox.question(self, 'Confirmation', 'Êtes-vous sûr de vouloir supprimer cette entrée ?',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm_delete == QMessageBox.Yes:
            id_item = self.table_destinataire.item(row, 0)
            if id_item is not None:
                id_to_delete = id_item.text()
                try:
                    self.data_destination = [item for item in self.data_destination if item['Id'] != id_to_delete]
                    self.table_destinataire.removeRow(row)
                    for i in range(row, self.table_destinataire.rowCount()):
                        self.table_destinataire.item(i, 0).setText(str(i + 1))
                    self.save_destination_to_json()
                except Exception as e:
                    print("Erreur lors de la suppression:", str(e))
                    pass
                finally:
                    self.load_destination_from_json()

    def resend_row_from_table(self, row):
        confirm_delete = QMessageBox.question(self, 'Confirmation', 'Êtes-vous sûr de vouloir renvoyer cette mail ?',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm_delete == QMessageBox.Yes:
            id_item = self.table_historique.item(row, 0).text()
            date_item = self.table_historique.item(row, 1).text()
            email_item = self.table_historique.item(row, 3).text()
            # name_item = self.table_historique.item(row, 3).text()
            # print("On a : ",id_item.text(), date_item.text(), email_item.text(), name_item.text())

            server = self.table_historique.item(row, 6).text()
            base_name = self.table_historique.item(row, 7).text()
            soc_name = self.table_historique.item(row, 8).text()
            filename = self.table_historique.item(row, 9).text()
            # adress_cc = self.table_historique.item(row, 10).text()
            
            connexion = None
            if id_item is not None:
                try:
                    id_to_resend = id_item
                    dates = datetime.datetime.strptime(date_item, "%d/%m/%Y %H:%M:%S")
                    new_status = False
                    try:
                        # json_file = load_json('config.json')
                        connexion = connexionSQlServer(server=server, base=base_name)
                        
                        # security break 2024
                        if connexion:
                            print("Resending...")

                            # filename = getDataLink(connexion, dates)
                            # print(filename)
                            # print(dates)
                            # print(email_item)
                            # print(filename)
                            # print(soc_name)
                            # print(adress_cc)

                            custom_send_email(dates=dates, email_adress=email_item, attachment_filename=filename, soc_name=soc_name)
                            new_status = True
                            print(f"Mail envoyer vers : {email_item}")

                    except Exception as e:
                        write_log(str(e))

                    modifier_objet(id_to_resend, new_status)
                    print(f"Modification Faite sur N° {id_to_resend}")

                except Exception as e:
                    print(f"Erreur : {str(e)}")
                finally:
                    if connexion is not None:
                        connexion.close()
                    self.load_historique_from_json()



                    

    def save_destination_to_json(self):
        data = []
        for row in range(self.table_destinataire.rowCount()):
            id_item = self.table_destinataire.item(row, 0)
            nom_item = self.table_destinataire.item(row, 1)
            email_item = self.table_destinataire.item(row, 2)

            # Get the status from the combo box and convert it to boolean
            status_combo = self.table_destinataire.cellWidget(row, 3)
            status = status_combo.currentText() == "Active"  # True if "Active", False otherwise

            # Get the selected society from the combo box
            society_combo = self.table_destinataire.cellWidget(row, 4)
            society = society_combo.currentText()

            if id_item is not None and nom_item is not None and email_item is not None:
                data.append({
                    'Id': id_item.text(),
                    'Nom': nom_item.text(),
                    'Email': email_item.text(),
                    'Status': status,  # Add the converted boolean status
                    'Societe': society  # Add the selected society
                })

        with open('destination.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, indent=4)

        self.load_destination_from_json()


    def load_destination_from_json(self):
        try:
            # with open('destination.json', 'r', encoding='utf-8') as file:
            #     data = json.load(file)

            destination_json = load_json('destination.json')

            societe_json = load_json('societe.json')

            society_names = []
            for soc in societe_json:
                society_names.append(soc['nom'])


                            
            # Effacer le contenu actuel du tableau
            self.table_destinataire.setRowCount(0)

            # Ajouter les données du fichier JSON au tableau
            # for item in data:
            #     row_position = self.table_destinataire.rowCount()
            #     self.table_destinataire.insertRow(row_position)
            #     self.table_destinataire.setItem(row_position, 0, QTableWidgetItem(str(item['Id'])))
            #     self.table_destinataire.setItem(row_position, 1, QTableWidgetItem(item['Nom']))
            #     self.table_destinataire.setItem(row_position, 2, QTableWidgetItem(item['Email']))
            #     self.table_destinataire.setItem(row_position, 3, QTableWidgetItem(item['Societe']))

            #     # Create a combo box with the appropriate status selected
            #     status_combo = QComboBox()
            #     status_combo.addItems(['Active', 'Inactive'])
            #     status_combo.setCurrentText(
            #         "Active" if item['Status'] else "Inactive")  # Set based on loaded boolean status
            #     self.table_destinataire.setCellWidget(row_position, 3, status_combo)

            #     delete_button = QPushButton('X')
            #     delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
            #     self.table_destinataire.setCellWidget(row_position, 4, delete_button)

            for item in destination_json:
                row_position = self.table_destinataire.rowCount()
                self.table_destinataire.insertRow(row_position)
                self.table_destinataire.setItem(row_position, 0, QTableWidgetItem(str(item['Id'])))
                self.table_destinataire.setItem(row_position, 1, QTableWidgetItem(item['Nom']))
                self.table_destinataire.setItem(row_position, 2, QTableWidgetItem(item['Email']))

                # Create a combo box with the appropriate status selected
                status_combo = QComboBox()
                status_combo.addItems(['Active', 'Inactive'])
                status_combo.setCurrentText("Active" if item['Status'] else "Inactive")  # Set based on loaded boolean status
                self.table_destinataire.setCellWidget(row_position, 3, status_combo)

                # Create a combo box for society names
                society_combo = QComboBox()
                society_combo.addItems(society_names)
                society_combo.setCurrentText(item['Societe'])  # Set current value based on loaded destination_json
                self.table_destinataire.setCellWidget(row_position, 4, society_combo)

                delete_button = QPushButton('X')
                delete_button.clicked.connect(lambda _, row=row_position: self.delete_row_from_table(row))
                self.table_destinataire.setCellWidget(row_position, 5, delete_button)

                # self.table_destinataire.setItem(row_position, 6, QTableWidgetItem(item['Cc']))


        except FileNotFoundError:
            print("Fichier Introuvable !")
            pass

    def load_historique_from_json(self):
        try:
            # with open('historique.json', 'r', encoding='utf-8') as file:
            #     data = json.load(file)
                
            try:
                with open('historique.json', 'r', encoding='utf-8') as file:
                    data = json.load(file)
            except json.JSONDecodeError:
                raise ValueError("The history.json file is empty or malformed. Add [] as a defaut")
                # raise ValueError("The JSON file is empty or malformed.")
            

            if data:
                self.table_historique.setRowCount(0)
                for item in data:
                    row_position = self.table_historique.rowCount()
                    self.table_historique.insertRow(row_position)
                    self.table_historique.setItem(row_position, 0, QTableWidgetItem(str(item['No'])))
                    self.table_historique.setItem(row_position, 1, QTableWidgetItem(str(item['DateTime'])))
                    self.table_historique.setItem(row_position, 2, QTableWidgetItem(str(item['Destinataire'])))
                    self.table_historique.setItem(row_position, 3, QTableWidgetItem(str(item['Email'])))
                    self.table_historique.setItem(row_position, 4, QTableWidgetItem("Success" if item['Status'] else "Echec"))

                    if not item['Status']:
                        resend_button = QPushButton('^')
                        resend_button.clicked.connect(lambda _, row=row_position: self.resend_row_from_table(row))
                        self.table_historique.setCellWidget(row_position, 5, resend_button)

                    self.table_historique.setItem(row_position, 6, QTableWidgetItem(str(item['server'])))
                    self.table_historique.setItem(row_position, 7, QTableWidgetItem(str(item['base_name'])))
                    self.table_historique.setItem(row_position, 8, QTableWidgetItem(str(item['soc_name'])))
                    self.table_historique.setItem(row_position, 9, QTableWidgetItem(str(item['filename'])))
                    # self.table_historique.setItem(row_position, 9, QTableWidgetItem(str(item['adress_cc'])))
                        


        except FileNotFoundError:
            print("Fichier Introuvable !")
            pass

    # affichage de la minuterie de compte à rebours en fonction du temps actuel et du temps cible
    def update_countdown(self):
        current_time = QDateTime.currentDateTime()
        if self.target_time is None:
            self.label.setText("Please set target time")
            return

        # Calculer la prochaine date et heure d'exécution en ajoutant un jour à la fois
        target_datetime = self.target_time
        while current_time >= target_datetime:
            target_datetime = target_datetime.addDays(1)

        diff = current_time.secsTo(target_datetime)

        if diff <= 0:
            self.iter_destination_json()
            self.recalculate_target_time()
            return

        days, remainder = divmod(diff, 86400)
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)

        time_left = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        self.label.setText(time_left)

        if time_left == "00:00:00":
            return "Envoyé avec succès"
        return time_left

    def start_countdown(self):
        try:
            target_hour = int(self.target_hour_edit.text())
            target_minute = int(self.target_minute_edit.text())
            target_second = int(self.target_second_edit.text())

            self.target_time = QDateTime.currentDateTime()
            self.target_time.setTime(QTime(target_hour, target_minute, target_second))
            # self.target_day = self.day_combo.currentText()

            self.timer.start(1000)  # Update every second
            self.label.setText(self.get_time_remaining_text())
            self.startBtn.setEnabled(False)
            self.endBtn.setEnabled(True)
        except Exception as e:
            print(str(e))
            pass

    def end_timer(self):
        try:
            self.timer.stop()
            self.label.setText("Countdown Stopped")
            self.startBtn.setEnabled(True)
            self.endBtn.setEnabled(False)
        except Exception as e:
            print(str(e))
            pass

    def recalculate_target_time(self):
        self.target_time = self.target_time.addDays(1)  # Ajoutez un jour au lieu de 7 jours
        self.label.setText(self.get_time_remaining_text())

    def get_time_remaining_text(self):
        return "Recalculating for Next day" if self.target_time is None else self.calculate_time_remaining()

    def calculate_time_remaining(self):
        current_time = QDateTime.currentDateTime()
        target_datetime = self.target_time.addDays(1)

        diff_seconds = current_time.secsTo(target_datetime)

        if diff_seconds < 0:
            diff_days = current_time.daysTo(target_datetime)  # Calculate days from current to target
            if diff_days >= 0:
                diff_seconds += (diff_days * 86400)  # Adjust diff_seconds based on days difference
            else:
                diff_seconds = -diff_seconds  # Handle negative diff_seconds appropriately

        diff_seconds %= 86400  # Get remaining seconds in the day

        hours, remainder = divmod(diff_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        time_left = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        return time_left


  
        
    def iter_destination_json(self):

        # only send if is on open days
        if is_open_day():

            # Load existing df from historique.json (or create an empty list if not found)
            try:
                # with open('historique.json', 'r', encoding='utf-8') as file:
                #     historique_data = json.load(file)
                historique_data = load_json('historique.json')
            except FileNotFoundError:
                historique_data = []

            next_no = len(historique_data) + 1  # Start with the next available N°

            # Process destinations and create new entries
            save = []

            # with open('destination.json', 'r', encoding='utf-8') as file:
            #     destination = json.load(file)

            destination_json = load_json('destination.json')
            # config_json = load_json('config.json')
            societe_json = load_json('societe.json')



            # should_not_send = False

            if destination_json:
                try:
                    for soc in societe_json:

                        base_name = soc['base']
                        base_type = soc['type']
                        soc_name = soc['nom']


                        server = soc['server']


                        # connexion = connexionSQlServer(server=server, base=base_name, soc_name=soc_name)
                        # filename = getDataLink(connexion, None, soc_name, base_type)

                        for dest in destination_json:
                            active = dest['Status']
                            adress_e = dest['Email']
                            # adress_cc = dest['Cc']
                            # adress_cc = None
                            # print(">>>>>>>>>>>>>>>>>> ", active, adress_e)

                            if active and dest['Societe'] == soc_name:

                                # print("ssssssssssssssssssssssssss")

                                connexion = connexionSQlServer(server=server, base=base_name, soc_name=soc_name)
                                filename = getDataLink(connexion, None, soc_name, base_type)

                                status = False

                                if connexion:
                                    try:
                                        custom_send_email(dates=get_today(), email_adress=adress_e, attachment_filename=filename, soc_name=soc_name)
                                        status = True
                                        # print(f"Mail envoyer vers : {dest['Email']}")
                                    except Exception as e:
                                        write_log(str(e))


                                    # for history (modif 2024)
                                    save.append({
                                        'No': next_no,
                                        'Destinataire': dest['Nom'],
                                        'Email': dest['Email'],
                                        'DateTime': get_today().strftime("%d/%m/%Y %H:%M:%S"),
                                        'Status': status,

                                        'server': server,
                                        'base_name': base_name,
                                        'soc_name': soc_name,
                                        'filename': filename
                                        # 'adress_cc': adress_cc,
                                    })
                                    next_no += 1  # Increment N° for the next entry



                        # Combine existing and new df (historique_data + save)
                        combined_data = historique_data + save

                        # Save the combined df to historique.json
                        with open('historique.json', 'w', encoding='utf-8') as file:
                            json.dump(combined_data, file, indent=4)

                        self.load_historique_from_json()  # Reload df to update table

                except Exception as e:
                    print(f"Error encountered concerning: {str(e)}, {soc_name}")
                    # should_not_send = True




if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = WinForm()
    win.show()
    sys.exit(app.exec_())
