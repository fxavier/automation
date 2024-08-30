import requests
import pandas as pd

# Configuration for DHIS2 API
dhis2_url = 'https://dhis2.echomoz.org' 
username = 'xnhagumbe' 
password = 'Go$btgo1'  

# Set up authentication
session = requests.Session()
session.auth = (username, password)

# Fetch the list of analytics tables
def get_analytics_tables():
    try:
        response = session.get(f'{dhis2_url}/api/analytics')
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching analytics tables: {e}")
        return None

# Fetch data from a specific analytics table
def fetch_analytics_data(table_name):
    try:
        response = session.get(f'{dhis2_url}/api/{table_name}')
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for table {table_name}: {e}")
        return None

# Main function to write analytics tables to CSV
def write_analytics_to_csv():
    analytics_tables = get_analytics_tables()
    
    if not analytics_tables:
        print("No analytics tables found.")
        return

    all_data = []

    for table in analytics_tables:
        table_name = table.get('name')
        print(f"Fetching data for table: {table_name}")
        data = fetch_analytics_data(table_name)

        if data:
            # Convert the data to DataFrame and append to the list
            df = pd.json_normalize(data)  # Flatten nested JSON
            all_data.append(df)

    # Concatenate all DataFrames and write to CSV
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df.to_csv('dhis2_analytics_data.csv', index=False)
        print("Analytics data written to 'dhis2_analytics_data.csv'.")
    else:
        print("No data to write.")

# Run the script
if __name__ == "__main__":
    write_analytics_to_csv()
