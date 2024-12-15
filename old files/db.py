# db.py
import pyodbc
from encryption import encrypt_data, decrypt_data
import logging
import os

# Configure logging with rotation
from logging.handlers import RotatingFileHandler

logger = logging.getLogger()
logger.setLevel(logging.INFO)

handler = RotatingFileHandler('../db.log', maxBytes=1000000, backupCount=5)
formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class Database:
    def __init__(self):
        try:
            # Define the path to your Access database
            db_path = r"C:/Users/Matt/Dropbox/Client Intake Program/Clients1.accdb"

            # Alternatively, use os.path to ensure cross-platform compatibility
            # db_path = os.path.join("C:", "Users", "Matt", "Dropbox", "Client Intake Program", "Clients1.accdb")

            # Ensure the path exists
            if not os.path.exists(db_path):
                logger.error(f"Database file does not exist at path: {db_path}")
                raise FileNotFoundError(f"Database file not found at {db_path}")

            # Define the connection string
            connection_string = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={db_path};'
            )

            # Establish the connection
            self.conn = pyodbc.connect(connection_string)
            self.cursor = self.conn.cursor()
            logger.info("Connected to Access database successfully.")
        except FileNotFoundError as fnf_error:
            logger.error(fnf_error)
            raise
        except pyodbc.Error as db_error:
            logger.error(f"Failed to connect to Access database: {db_error}")
            raise ConnectionError(f"Failed to connect to Access database: {db_error}")
        except Exception as e:
            logger.error(f"An unexpected error occurred: {e}")
            raise ConnectionError(f"An unexpected error occurred: {e}")

    def insert_client_data(self, client_data):
        """
        Inserts a new client record into the Clients table.
        """
        try:
            # Prepare the SQL statement
            fields = ', '.join(client_data.keys())
            placeholders = ', '.join(['?'] * len(client_data))
            sql = f"INSERT INTO Clients ({fields}) VALUES ({placeholders})"
            values = list(client_data.values())
            self.cursor.execute(sql, values)
            self.conn.commit()
            logger.info("Inserted client data successfully.")
        except pyodbc.Error as db_error:
            self.conn.rollback()
            logger.error(f"Error inserting client data: {db_error}")
            raise
        except Exception as e:
            self.conn.rollback()
            logger.error(f"Unexpected error inserting client data: {e}")
            raise

    def get_client_by_email(self, email):
        """
        Retrieves a client record based on email.
        """
        try:
            sql = "SELECT * FROM Clients WHERE Client_Email = ?"
            self.cursor.execute(sql, (email,))
            row = self.cursor.fetchone()
            if row:
                # Convert the row to a dictionary
                columns = [column[0] for column in self.cursor.description]
                client = dict(zip(columns, row))
                logger.info(f"Retrieved client data for email: {email}")
                return client
            else:
                logger.info(f"No client found with email: {email}")
                return None
        except pyodbc.Error as db_error:
            logger.error(f"Error retrieving client data: {db_error}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error retrieving client data: {e}")
            raise

    def update_client_data(self, email, update_fields):
        """
        Updates client data based on email.
        """
        try:
            set_clause = ', '.join([f"{field} = ?" for field in update_fields.keys()])
            sql = f"UPDATE Clients SET {set_clause} WHERE Client_Email = ?"
            values = list(update_fields.values()) + [email]
            self.cursor.execute(sql, values)
            self.conn.commit()
            logger.info(f"Updated client data for email: {email}")
        except pyodbc.Error as db_error:
            self.conn.rollback()
            logger.error(f"Error updating client data: {db_error}")
            raise
        except Exception as e:
            self.conn.rollback()
            logger.error(f"Unexpected error updating client data: {e}")
            raise

    def delete_client(self, email):
        """
        Deletes a client record based on email.
        """
        try:
            sql = "DELETE FROM Clients WHERE Client_Email = ?"
            self.cursor.execute(sql, (email,))
            self.conn.commit()
            logger.info(f"Deleted client data for email: {email}")
        except pyodbc.Error as db_error:
            self.conn.rollback()
            logger.error(f"Error deleting client data: {db_error}")
            raise
        except Exception as e:
            self.conn.rollback()
            logger.error(f"Unexpected error deleting client data: {e}")
            raise

    def close_connection(self):
        """
        Closes the database connection.
        """
        try:
            self.cursor.close()
            self.conn.close()
            logger.info("Closed database connection successfully.")
        except pyodbc.Error as db_error:
            logger.error(f"Error closing database connection: {db_error}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error closing database connection: {e}")
            raise

    ```


### **c. Key Points in the Configuration:**

1. ** Raw
Strings
for Paths: **
- Prefixing
the
string
with `r`(e.g., `r"Path"`) ensures that backslashes are treated as literal characters, preventing escape sequence issues.
- Since
you
're using forward slashes (`/`) in your path, it'
s
already
safe.However, using
raw
strings is a
good
practice.

2. ** Validating
the
Database
Path: **
- Before
attempting
to
connect, the
script
checks if the
database
file
exists
at
the
specified
path.
- If
the
file
doesn
't exist, it logs an error and raises a `FileNotFoundError`.

3. ** Using
`RotatingFileHandler`
for Logging: **
- This
prevents
log
files
from growing indefinitely

by
rotating
them
after
they
reach
a
specified
size.
- In
this
configuration, each
log
file
can
grow
up
to ** 1, 000, 000
bytes(approximately
1
MB) **, and up
to ** 5
backup
logs ** are
kept.

4. ** Error
Handling: **
- The
`
try-except` blocks catch and log both `pyodbc` specific errors and general exceptions.
- On
encountering
an
error
during
database
operations, the
script
rolls
back
the
transaction
to
maintain
database
integrity.

---

## **2. Ensuring ODBC Driver Installation**

To
connect
to
an
Access
database
using
`pyodbc`, you
must
have
the
appropriate
ODBC
drivers
installed
on
your
system.

### **a. Check for Installed Drivers**

Run
the
following
script
to
list
all
installed
ODBC
drivers:

```python
import pyodbc


def list_odbc_drivers():
    drivers = pyodbc.drivers()
    print("Installed ODBC Drivers:")
    for driver in drivers:
        print(driver)


if __name__ == "__main__":
    list_odbc_drivers()
