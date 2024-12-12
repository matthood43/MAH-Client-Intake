import os
import pyodbc
import dropbox
import platform
import base64
import re
from cryptography.fernet import Fernet
from docx import Document
from io import BytesIO
from dotenv import load_dotenv

# PyQt5 imports for the GUI
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel,
                             QLineEdit, QPushButton, QFormLayout, QMessageBox)

load_dotenv()

# Dropbox Access Token (use environment variables for security)
DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN")

if not DROPBOX_ACCESS_TOKEN:
    raise ValueError("Error: Dropbox access token not set. Use environment variables for security.")

# Initialize Dropbox Client
try:
    db_client = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
except dropbox.exceptions.AuthError as e:
    raise ValueError(f"Dropbox Authentication failed: {e}")

# Encryption Key Management
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

# Utility Functions
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

def download_from_dropbox(dropbox_path, local_path):
    try:
        metadata, response = db_client.files_download(dropbox_path)
        with open(local_path, "wb") as file:
            file.write(response.content)
            print(f"Downloaded from Dropbox: {dropbox_path} to {local_path}")
    except Exception as e:
        print(f"Error downloading from Dropbox: {e}")

def initialize_database(local_db_path):
    """Ensure the database exists locally and in Dropbox."""
    db_dropbox_path = "/Client Intake Program/Clients1.accdb"

    try:
        # Check if the database exists in Dropbox
        db_client.files_get_metadata(db_dropbox_path)
        print("Database exists in Dropbox.")
    except dropbox.exceptions.ApiError:
        # If not, upload the local database to Dropbox
        print("Database not found in Dropbox. Uploading local database...")
        upload_to_dropbox(local_db_path, db_dropbox_path)

    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + local_db_path + r';PWD=your_password')
    return conn

def insert_client_data(conn, encrypted_data):
    """Insert encrypted client data into the database."""
    try:
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT INTO Clients (Client_First_Name, Client_Last_Name, Client_Email, Client_Phone, Client_Address, 
            Opposing_Party, Primary_Attorney, Fee_Arrangement, Fee_Amount) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (encrypted_data['Client_First_Name'], encrypted_data['Client_Last_Name'], encrypted_data['Client_Email'],
             encrypted_data['Client_Phone'], encrypted_data['Client_Address'], encrypted_data['Opposing_Party'],
             encrypted_data['Primary_Attorney'], encrypted_data['Fee_Arrangement'], encrypted_data['Fee_Amount'])
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
        # Download the template from Dropbox
        metadata, response = db_client.files_download(template_path)
        document = Document(BytesIO(response.content))

        # Replace placeholders with client data
        for paragraph in document.paragraphs:
            for key, value in client_data.items():
                placeholder = f"{{{key}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, value)

        # Save the completed document locally
        document.save(output_path)
        print(f"Fee agreement generated: {output_path}")
    except Exception as e:
        print(f"Error processing fee agreement: {e}")

def process_client_data(client_data):
    """
    This function encapsulates the logic previously found in `main()`:
    - Validate data
    - Encrypt data
    - Store in DB
    - Create Dropbox folders
    - Generate and upload fee agreement
    """

    # Validate Email
    if not is_valid_email(client_data['Client_Email']):
        raise ValueError("Invalid email address.")

    # Validate Fee Amount
    if not client_data['Fee_Amount'].replace('.', '').isdigit():
        raise ValueError("Fee amount must be numeric.")

    # Encrypt client data
    encrypted_data = {key: encrypt_data(value) for key, value in client_data.items()}

    # Initialize database
    local_db_path = "C:/Users/Matt/Dropbox/Client Intake Program/Clients1.accdb"
    conn = initialize_database(local_db_path)

    # Store client data in the database
    insert_client_data(conn, encrypted_data)

    # Create Dropbox folders
    client_name = f"{client_data['Client_First_Name']}_{client_data['Client_Last_Name']}"
    create_dropbox_folders(client_name)

    # Generate fee agreement
    template_path = "/Client Intake Program/Fee_Agreement.docx"
    output_path = f"{client_name}_fee_agreement.docx"
    process_fee_agreement(template_path, client_data, output_path)

    # Upload fee agreement to Dropbox
    upload_to_dropbox(output_path, f"/{client_name}/Correspondence/{output_path}")
    if os.path.exists(output_path):
        os.remove(output_path)

    return client_name


# -------------------------
# GUI Code Integration
# -------------------------
class ClientIntakeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Client Intake System")
        self.setGeometry(100, 100, 400, 300)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QFormLayout(central_widget)

        self.first_name_input = QLineEdit()
        self.last_name_input = QLineEdit()
        self.email_input = QLineEdit()
        self.phone_input = QLineEdit()
        self.address_input = QLineEdit()
        self.opposing_party_input = QLineEdit()
        self.primary_attorney_input = QLineEdit()
        self.fee_arrangement_input = QLineEdit()
        self.fee_amount_input = QLineEdit()

        layout.addRow("Client First Name:", self.first_name_input)
        layout.addRow("Client Last Name:", self.last_name_input)
        layout.addRow("Client Email:", self.email_input)
        layout.addRow("Client Phone:", self.phone_input)
        layout.addRow("Client Address:", self.address_input)
        layout.addRow("Opposing Party:", self.opposing_party_input)
        layout.addRow("Primary Attorney:", self.primary_attorney_input)
        layout.addRow("Fee Arrangement (Hourly/Contingency):", self.fee_arrangement_input)
        layout.addRow("Fee Amount (Rate or Percentage):", self.fee_amount_input)

        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.on_submit)
        layout.addRow(self.submit_button)

    def on_submit(self):
        client_data = {
            "Client_First_Name": self.first_name_input.text().strip(),
            "Client_Last_Name": self.last_name_input.text().strip(),
            "Client_Email": self.email_input.text().strip(),
            "Client_Phone": self.phone_input.text().strip(),
            "Client_Address": self.address_input.text().strip(),
            "Opposing_Party": self.opposing_party_input.text().strip(),
            "Primary_Attorney": self.primary_attorney_input.text().strip(),
            "Fee_Arrangement": self.fee_arrangement_input.text().strip(),
            "Fee_Amount": self.fee_amount_input.text().strip()
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
