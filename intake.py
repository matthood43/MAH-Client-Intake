import os
import json
import pyodbc
import dropbox
import base64
import re
import requests
from cryptography.fernet import Fernet
from docx import Document
from io import BytesIO
from dotenv import load_dotenv

# PyQt5 imports for GUI
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QFormLayout, QMessageBox, QCheckBox)

load_dotenv()

# Dropbox Access Token (ensure this is set in your .env file)
DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN")
if not DROPBOX_ACCESS_TOKEN:
    raise ValueError("Dropbox access token not set.")

# Harvest Credentials
# Replace with your actual Harvest Account ID and Personal Access Token.
HARVEST_ACCOUNT_ID = "905300"
HARVEST_TOKEN = "1546177.pt.oMJ-PjYazgmqnrIvRHYA9J8aC2Hah0eIFnYpJzjbKBSGixLEN9F2uqopPqdRzfgckt64JYwOWAwmQ0ZZhY3THQ"

# Bosch Client ID in Harvest (replace with the actual client_id for Bosch)
HARVEST_BOSCH_CLIENT_ID = 6685250

try:
    db_client = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
except dropbox.exceptions.AuthError as e:
    raise ValueError(f"Dropbox Authentication failed: {e}")

def generate_key():
    key = Fernet.generate_key()
    with open("encryption_key.key", "wb") as key_file:
        key_file.write(key)
    return key

def load_key():
    if not os.path.exists("encryption_key.key"):
        generate_key()
    return open("encryption_key.key", "rb").read()

def encrypt_data(data):
    key = load_key()
    f = Fernet(key)
    encrypted_data = f.encrypt(data.encode())
    return base64.b64encode(encrypted_data).decode('utf-8')

def is_valid_email(email):
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return re.match(email_regex, email)

def upload_to_dropbox(local_path, dropbox_path):
    try:
        with open(local_path, "rb") as file:
            db_client.files_upload(file.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
            print(f"Uploaded to Dropbox: {dropbox_path}")
    except Exception as e:
        print(f"Error uploading to Dropbox: {e}")

def initialize_database(local_db_path):
    """Ensure the database exists locally and in Dropbox."""
    db_dropbox_path = "/Client Intake Program/Clients1.accdb"

    try:
        db_client.files_get_metadata(db_dropbox_path)
        print("Database exists in Dropbox.")
    except dropbox.exceptions.ApiError:
        print("Database not found in Dropbox. Uploading local database...")
        upload_to_dropbox(local_db_path, db_dropbox_path)

    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + local_db_path + r';PWD=your_password')
    return conn

def insert_client_data(conn, encrypted_data_map):
    """Insert encrypted client data into the database.
       Fields:
       Client_First_Name, Client_Last_Name, Opposing_Party, Client_Email, Client_Address, Client_Phone,
       Fee_Arrangement, Fee_Amount, Contingency_Amount, Hourly_Amount, Primary_Attorney, EncryptedData, IsBosch
    """
    try:
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT INTO Clients (
                Client_First_Name, Client_Last_Name, Opposing_Party, Client_Email, Client_Address, Client_Phone,
                Fee_Arrangement, Fee_Amount, Contingency_Amount, Hourly_Amount, Primary_Attorney, EncryptedData, IsBosch
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (
                encrypted_data_map['Client_First_Name'],
                encrypted_data_map['Client_Last_Name'],
                encrypted_data_map['Opposing_Party'],
                encrypted_data_map['Client_Email'],
                encrypted_data_map['Client_Address'],
                encrypted_data_map['Client_Phone'],
                encrypted_data_map['Fee_Arrangement'],
                encrypted_data_map['Fee_Amount'],
                encrypted_data_map['Contingency_Amount'],
                encrypted_data_map['Hourly_Amount'],
                encrypted_data_map['Primary_Attorney'],
                encrypted_data_map['EncryptedData'],
                encrypted_data_map['IsBosch']
            )
        )
        conn.commit()
        print("Client data stored successfully.")
    except Exception as e:
        print(f"Error inserting data into the database: {e}")

def create_dropbox_folders(client_name):
    folder_path = f"/{client_name}"
    subfolders = ["Pleadings", "Correspondence", "Prelitigation", "Case Notes"]

    try:
        db_client.files_create_folder_v2(folder_path)
        print(f"Main folder created: {folder_path}")
    except dropbox.exceptions.ApiError as e:
        if e.error.is_path() and e.error.get_path().is_conflict():
            print(f"Folder already exists: {folder_path}")
        else:
            raise

    for subfolder in subfolders:
        try:
            db_client.files_create_folder_v2(f"{folder_path}/{subfolder}")
            print(f"Subfolder created: {subfolder}")
        except dropbox.exceptions.ApiError as e:
            if e.error.is_path() and e.error.get_path().is_conflict():
                print(f"Subfolder already exists: {subfolder}")
            else:
                raise

def process_fee_agreement(template_path, client_data, output_path):
    """Replace placeholders in the fee agreement template with client data."""
    try:
        metadata, response = db_client.files_download(template_path)
        document = Document(BytesIO(response.content))

        for paragraph in document.paragraphs:
            for key, value in client_data.items():
                placeholder = f"{{{key}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, value)

        document.save(output_path)
        print(f"Fee agreement generated: {output_path}")
    except Exception as e:
        print(f"Error processing fee agreement: {e}")

def create_harvest_project(client_data, harvest_bosch_client_id):
    headers = {
        "User-Agent": "ClientIntakeApp (support@example.com)",
        "Authorization": f"Bearer {HARVEST_TOKEN}",
        "Harvest-Account-Id": HARVEST_ACCOUNT_ID,
        "Content-Type": "application/json"
    }

    project_data = {
        "client_id": harvest_bosch_client_id,
        "name": f"{client_data['Client_First_Name']} {client_data['Client_Last_Name']}",
        "is_billable": True,
        "bill_by": "Project",
        "hourly_rate": 450.0
    }

    response = requests.post("https://api.harvestapp.com/v2/projects", json=project_data, headers=headers)
    if response.status_code == 201:
        print("Harvest project created successfully with an hourly rate of 450.")
    else:
        print(f"Failed to create Harvest project: {response.status_code}, {response.text}")
        raise Exception("Harvest project creation failed.")

def process_client_data(client_data):
    if not is_valid_email(client_data['Client_Email']):
        raise ValueError("Invalid email address.")

    for field in ['Fee_Amount', 'Contingency_Amount', 'Hourly_Amount']:
        val = client_data[field]
        if val and not val.replace('.', '').isdigit():
            raise ValueError(f"{field} must be numeric.")

    local_db_path = "C:/Users/Matt/Dropbox/Client Intake Program/Clients1.accdb"
    conn = initialize_database(local_db_path)

    raw_json = json.dumps(client_data)
    encrypted_data_map = {key: encrypt_data(value) for key, value in client_data.items()}
    encrypted_data_map['EncryptedData'] = encrypt_data(raw_json)

    insert_client_data(conn, encrypted_data_map)

    client_name = f"{client_data['Client_First_Name']}_{client_data['Client_Last_Name']}"
    create_dropbox_folders(client_name)

    template_path = "/Client Intake Program/Fee_Agreement.docx"
    output_path = f"{client_name}_fee_agreement.docx"
    process_fee_agreement(template_path, client_data, output_path)

    upload_to_dropbox(output_path, f"/{client_name}/Correspondence/{output_path}")
    if os.path.exists(output_path):
        os.remove(output_path)

    if client_data['IsBosch'] == "True":
        create_harvest_project(client_data, HARVEST_BOSCH_CLIENT_ID)

    return client_name

class ClientIntakeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Client Intake System")
        self.setGeometry(100, 100, 400, 500)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QFormLayout(central_widget)

        self.first_name_input = QLineEdit()
        self.last_name_input = QLineEdit()
        self.opposing_party_input = QLineEdit()
        self.email_input = QLineEdit()
        self.address_input = QLineEdit()
        self.phone_input = QLineEdit()
        self.fee_arrangement_input = QLineEdit()
        self.fee_amount_input = QLineEdit()
        self.contingency_amount_input = QLineEdit()
        self.hourly_amount_input = QLineEdit()
        self.primary_attorney_input = QLineEdit()
        self.bosch_checkbox = QCheckBox("Is this a Bosch file?")

        layout.addRow("Client First Name:", self.first_name_input)
        layout.addRow("Client Last Name:", self.last_name_input)
        layout.addRow("Opposing Party:", self.opposing_party_input)
        layout.addRow("Client Email:", self.email_input)
        layout.addRow("Client Address:", self.address_input)
        layout.addRow("Client Phone:", self.phone_input)
        layout.addRow("Fee Arrangement:", self.fee_arrangement_input)
        layout.addRow("Fee Amount:", self.fee_amount_input)
        layout.addRow("Contingency Amount:", self.contingency_amount_input)
        layout.addRow("Hourly Amount:", self.hourly_amount_input)
        layout.addRow("Primary Attorney:", self.primary_attorney_input)
        layout.addRow(self.bosch_checkbox)

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.on_submit)
        layout.addRow(self.submit_button)

    def on_submit(self):
        client_data = {
            "Client_First_Name": self.first_name_input.text().strip(),
            "Client_Last_Name": self.last_name_input.text().strip(),
            "Opposing_Party": self.opposing_party_input.text().strip(),
            "Client_Email": self.email_input.text().strip(),
            "Client_Address": self.address_input.text().strip(),
            "Client_Phone": self.phone_input.text().strip(),
            "Fee_Arrangement": self.fee_arrangement_input.text().strip(),
            "Fee_Amount": self.fee_amount_input.text().strip(),
            "Contingency_Amount": self.contingency_amount_input.text().strip(),
            "Hourly_Amount": self.hourly_amount_input.text().strip(),
            "Primary_Attorney": self.primary_attorney_input.text().strip(),
            "IsBosch": "True" if self.bosch_checkbox.isChecked() else "False"
        }

        try:
            client_name = process_client_data(client_data)
            QMessageBox.information(self, "Success", f"Client intake process completed successfully for {client_name}")
        except ValueError as ve:
            QMessageBox.warning(self, "Validation Error", str(ve))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

def main():
    app = QApplication(sys.argv)
    window = ClientIntakeWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
