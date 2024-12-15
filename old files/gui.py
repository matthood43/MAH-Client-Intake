# gui.py
import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from utils import is_valid_email
from db import Database
from dropbox_utils import upload_to_dropbox, create_dropbox_folders, process_fee_agreement, get_dropbox_client
from encryption import encrypt_data
from harvest import create_harvest_project
import logging

class AuthDialog(tk.Toplevel):
    """
    A dialog to prompt the user to authenticate with Dropbox.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Dropbox Authentication")
        self.geometry("400x200")
        self.resizable(False, False)
        self.parent = parent

        self.label = ttk.Label(self, text="Dropbox access is not authenticated.\nPlease authenticate to continue.", anchor="center")
        self.label.pack(pady=20)

        self.auth_button = ttk.Button(self, text="Authenticate with Dropbox", command=self.authenticate)
        self.auth_button.pack(pady=10)

    def authenticate(self):
        try:
            # Pass the parent window to get_dropbox_client for dialogs
            access_token = get_dropbox_client(self.parent)
            messagebox.showinfo("Authentication Successful", "Dropbox has been authenticated successfully.")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Authentication Failed", f"An error occurred during authentication:\n{e}")
            self.destroy()

class ClientIntakeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Client Intake System")
        self.geometry("600x700")
        self.resizable(False, False)

        # Initialize Database
        try:
            self.db = Database()
        except ConnectionError as ce:
            messagebox.showerror("Database Connection Error", str(ce))
            self.destroy()
            return  # Exit initialization if database connection fails

        # Initialize Dropbox Client
        try:
            self.db_client = get_dropbox_client(self)
        except ValueError as ve:
            self.db_client = None
            messagebox.showerror("Dropbox Authentication Error", str(ve))
            self.prompt_authentication()

        # Create Widgets
        self.create_widgets()

    def prompt_authentication(self):
        """
        Prompts the user to authenticate with Dropbox via a dialog.
        """
        auth_dialog = AuthDialog(self)
        self.wait_window(auth_dialog)
        try:
            self.db_client = get_dropbox_client(self)
            messagebox.showinfo("Dropbox Authenticated", "Dropbox has been authenticated successfully.")
        except Exception as e:
            messagebox.showerror("Authentication Failed", f"Failed to authenticate with Dropbox:\n{e}")
            self.destroy()

    def create_widgets(self):
        # Create a frame for the form
        form_frame = ttk.Frame(self, padding="20 20 20 20")
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Form Fields
        self.entries = {}

        fields = [
            ("Client First Name:", "first_name"),
            ("Client Last Name:", "last_name"),
            ("Opposing Party:", "opposing_party"),
            ("Client Email:", "email"),
            ("Client Address:", "address"),
            ("Client Phone:", "phone"),
            ("Fee Arrangement:", "fee_arrangement"),
            ("Fee Amount:", "fee_amount"),
            ("Contingency Amount:", "contingency_amount"),
            ("Hourly Amount:", "hourly_amount"),
            ("Primary Attorney:", "primary_attorney")
        ]

        for idx, (label_text, key) in enumerate(fields):
            label = ttk.Label(form_frame, text=label_text)
            label.grid(row=idx, column=0, sticky=tk.W, pady=5)

            entry = ttk.Entry(form_frame, width=40)
            entry.grid(row=idx, column=1, pady=5, padx=10)
            self.entries[key] = entry

        # Bosch Checkbox
        self.bosch_var = tk.BooleanVar()
        self.bosch_checkbox = ttk.Checkbutton(form_frame, text="Is this a Bosch file?", variable=self.bosch_var)
        self.bosch_checkbox.grid(row=len(fields), column=0, columnspan=2, pady=10)

        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=20)

        submit_button = ttk.Button(button_frame, text="Submit", command=self.on_submit)
        submit_button.grid(row=0, column=0, padx=10)

        clear_button = ttk.Button(button_frame, text="Clear", command=self.on_clear)
        clear_button.grid(row=0, column=1, padx=10)

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.destroy)
        cancel_button.grid(row=0, column=2, padx=10)

        # Progress Bar
        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=20, pady=10)
        self.progress.pack_forget()

        # Status Bar
        self.status = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(fill=tk.X, side=tk.BOTTOM, ipady=2)

        # Email Validation Binding
        self.entries['email'].bind("<KeyRelease>", self.validate_email)

    def validate_email(self, event):
        email = self.entries['email'].get().strip()
        if is_valid_email(email):
            self.entries['email'].config(foreground="green")
        else:
            self.entries['email'].config(foreground="red")

    def on_submit(self):
        if not self.db_client:
            messagebox.showwarning("Authentication Required", "Please authenticate with Dropbox before submitting.")
            self.prompt_authentication()
            return

        # Gather client data from the form
        client_data = {
            "Client_First_Name": self.entries['first_name'].get().strip(),
            "Client_Last_Name": self.entries['last_name'].get().strip(),
            "Opposing_Party": self.entries['opposing_party'].get().strip(),
            "Client_Email": self.entries['email'].get().strip(),
            "Client_Address": self.entries['address'].get().strip(),
            "Client_Phone": self.entries['phone'].get().strip(),
            "Fee_Arrangement": self.entries['fee_arrangement'].get().strip(),
            "Fee_Amount": float(self.entries['fee_amount'].get().strip()) if self.entries['fee_amount'].get().strip() else None,
            "Contingency_Amount": float(self.entries['contingency_amount'].get().strip()) if self.entries['contingency_amount'].get().strip() else None,
            "Hourly_Amount": float(self.entries['hourly_amount'].get().strip()) if self.entries['hourly_amount'].get().strip() else None,
            "Primary_Attorney": self.entries['primary_attorney'].get().strip(),
            "IsBosch": self.bosch_var.get(),
            "EncryptedData": encrypt_data(json.dumps({
                "Client_First_Name": self.entries['first_name'].get().strip(),
                "Client_Last_Name": self.entries['last_name'].get().strip(),
                "Opposing_Party": self.entries['opposing_party'].get().strip(),
                "Client_Email": self.entries['email'].get().strip(),
                "Client_Address": self.entries['address'].get().strip(),
                "Client_Phone": self.entries['phone'].get().strip(),
                "Fee_Arrangement": self.entries['fee_arrangement'].get().strip(),
                "Fee_Amount": self.entries['fee_amount'].get().strip(),
                "Contingency_Amount": self.entries['contingency_amount'].get().strip(),
                "Hourly_Amount": self.entries['hourly_amount'].get().strip(),
                "Primary_Attorney": self.entries['primary_attorney'].get().strip(),
                "IsBosch": self.bosch_var.get()
            }))
        }

        # Show progress bar and update status
        self.progress.pack()
        self.progress.start(10)
        self.status.config(text="Processing client data...")
        self.update_idletasks()

        try:
            # Validate email
            if not is_valid_email(client_data['Client_Email']):
                raise ValueError("Invalid email address.")

            # Validate numeric fields
            for field in ['Fee_Amount', 'Contingency_Amount', 'Hourly_Amount']:
                val = client_data[field]
                if val is not None and val < 0:
                    raise ValueError(f"{field} must be a positive number.")

            # Insert into Access Database
            self.db.insert_client_data(client_data)

            # Create Dropbox folders
            client_name = f"{client_data['Client_First_Name']}_{client_data['Client_Last_Name']}"
            create_dropbox_folders(client_name, self)  # Pass the root window

            # Process fee agreement
            output_path = f"{client_name}_fee_agreement.docx"
            process_fee_agreement(client_data, output_path)

            # Upload fee agreement to Dropbox
            dropbox_fee_agreement_path = f"/{client_name}/Correspondence/{output_path}"
            upload_to_dropbox(output_path, dropbox_fee_agreement_path, self)  # Pass the root window
            if os.path.exists(output_path):
                os.remove(output_path)

            # Create Harvest project if Bosch
            if client_data['IsBosch']:
                from config import HARVEST_BOSCH_CLIENT_ID  # Import here to avoid circular imports
                create_harvest_project(client_data, HARVEST_BOSCH_CLIENT_ID)

            messagebox.showinfo("Success", f"Client intake process completed successfully for {client_name}")
            self.status.config(text=f"Client {client_name} processed successfully.")
        except ValueError as ve:
            messagebox.showwarning("Validation Error", str(ve))
            self.status.config(text="Validation error occurred.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status.config(text="An error occurred during processing.")
        finally:
            # Hide progress bar and reset
            self.progress.stop()
            self.progress.pack_forget()

    def on_clear(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.bosch_var.set(False)
        self.status.config(text="Form cleared.")

    def on_close(self):
        """
        Ensures that the database connection is closed before exiting.
        """
        try:
            self.db.close_connection()
        except Exception as e:
            logging.error(f"Error closing database connection: {e}")
        self.destroy()

    # Override the default close behavior to ensure proper cleanup
    def run(self):
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.mainloop()

if __name__ == "__main__":
    app = ClientIntakeApp()
    app.run()
