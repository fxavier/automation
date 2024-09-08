import requests
import os
from requests.auth import HTTPBasicAuth

dhis2_url = "https://dhis2sand.echomoz.org/api/dataValueSets"
username = "xnhagumbe"  
password = "Go$btgo1" 

# Path to the CSV file
file_path = os.path.join(os.path.dirname(__file__), 'Merged', 'file_to_import.csv')

# Headers for the request
headers = {
    'Content-Type': 'application/csv'
}

# Read the CSV file and prepare it for upload
with open(file_path, 'rb') as file:
    # Make the POST request to upload the file
    response = requests.post(
        dhis2_url,
        headers=headers,
        data=file,
        auth=HTTPBasicAuth(username, password)
    )

# Check the response status
if response.status_code == 200:
    print("Data imported successfully!")
else:
    print(f"Failed to import data. Status code: {response.status_code}")
    print(f"Response: {response.text}")