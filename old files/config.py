# config.py
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Database Configuration using pyodbc
ACCESS_DB_PATH = os.getenv("ACCESS_DB_PATH", r"C:\Path\To\Your\Database.accdb")
ACCESS_DB_PASSWORD = os.getenv("ACCESS_DB_PASSWORD", "")  # If your DB is password protected

# Construct the ODBC connection string
if ACCESS_DB_PATH.endswith('.accdb'):
    # For Access 2007 and later
    DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
elif ACCESS_DB_PATH.endswith('.mdb'):
    # For Access 2003 and earlier
    DRIVER = "{Microsoft Access Driver (*.mdb)}"
else:
    raise ValueError("Unsupported database file extension. Use .accdb or .mdb")

# Build the connection string
if ACCESS_DB_PASSWORD:
    CONNECTION_STRING = (
        f"DRIVER={DRIVER};"
        f"DBQ={ACCESS_DB_PATH};"
        f"PWD={ACCESS_DB_PASSWORD};"
    )
else:
    CONNECTION_STRING = (
        f"DRIVER={DRIVER};"
        f"DBQ={ACCESS_DB_PATH};"
    )

# Dropbox OAuth2 Credentials
DROPBOX_CLIENT_ID = os.getenv("DROPBOX_CLIENT_ID")
DROPBOX_CLIENT_SECRET = os.getenv("DROPBOX_CLIENT_SECRET")
DROPBOX_REDIRECT_URI = os.getenv("DROPBOX_REDIRECT_URI")

if not all([DROPBOX_CLIENT_ID, DROPBOX_CLIENT_SECRET, DROPBOX_REDIRECT_URI]):
    raise ValueError("Dropbox OAuth2 credentials are not fully set in the .env file.")

# Dropbox Token Storage
DROPBOX_TOKEN_FILE = os.getenv("DROPBOX_TOKEN_FILE", "../dropbox_token.enc")

# Harvest Credentials
HARVEST_ACCOUNT_ID = os.getenv("HARVEST_ACCOUNT_ID", "905300")
HARVEST_TOKEN = os.getenv("HARVEST_TOKEN")
if not HARVEST_TOKEN:
    raise ValueError("Harvest token not set.")

# Bosch Client ID in Harvest
HARVEST_BOSCH_CLIENT_ID = int(os.getenv("HARVEST_BOSCH_CLIENT_ID", "6685250"))

# Dropbox Paths (if still applicable)
DB_DROPBOX_PATH = "/Client Intake Program/Clients1.accdb"  # May be obsolete if moving away from Access
FEE_AGREEMENT_TEMPLATE_PATH = "/Client Intake Program/Fee_Agreement.docx"
