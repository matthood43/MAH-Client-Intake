# harvest.py
import requests
from config import HARVEST_ACCOUNT_ID, HARVEST_TOKEN

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
