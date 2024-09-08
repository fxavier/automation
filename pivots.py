# import requests
# from requests.auth import HTTPBasicAuth

# # DHIS2 credentials and URL
# url = "https://dhis2.echomoz.org/api/29/visualizations.json"
# username = "xnhagumbe"
# password = "Go$btgo1"

# def list_pivot_tables():
#     try:
#         # Make a GET request to the DHIS2 API to fetch pivot tables
#         response = requests.get(url, auth=HTTPBasicAuth(username, password), params={'filter': 'type:eq:PIVOT_TABLE', 'fields': 'id,name', 'paging': 'false'})
#         response.raise_for_status()  # Raise an HTTPError for bad responses

#         # Parse the JSON response
#         pivot_tables = response.json().get('visualizations', [])

#         # Print the list of pivot tables
#         if (pivot_tables):
#             print("Pivot Tables:")
#             for pt in pivot_tables:
#                 print(f"ID: {pt['id']}, Name: {pt['name']}")
#         else:
#             print("No pivot tables found.")

#     except requests.exceptions.HTTPError as http_err:
#         print(f"HTTP error occurred: {http_err}")
#     except requests.exceptions.RequestException as req_err:
#         print(f"Request error occurred: {req_err}")
#     except Exception as e:
#         print(f"An error occurred: {e}")

# if __name__ == "__main__":
#     list_pivot_tables()

import requests
from requests.auth import HTTPBasicAuth
import csv

# DHIS2 credentials and URL
url = "https://dhis2.echomoz.org/api/29/visualizations.json"
username = "xnhagumbe"
password = "Go$btgo1"

def list_pivot_tables():
    try:
        response = requests.get(url, auth=HTTPBasicAuth(username, password), params={'filter': 'type:eq:PIVOT_TABLE', 'fields': 'id,name', 'paging': 'false'})
        response.raise_for_status()  
        pivot_tables = response.json().get('visualizations', [])

        if (pivot_tables):
            print("Pivot Tables:")
            for pt in pivot_tables:
                print(f"ID: {pt['id']}, Name: {pt['name']}")
        else:
            print("No pivot tables found.")

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
    except requests.exceptions.RequestException as req_err:
        print(f"Request error occurred: {req_err}")
    except Exception as e:
        print(f"An error occurred: {e}")

def list_indicators_and_save_to_csv():
    try:
        indicators_url = "https://dhis2.echomoz.org/api/29/analytics.json"
        params = {
            'dimension': [
                'dx:dkZHx1STjIG;ODw4ILhNQjc;DPelk9N6lkO;cIUIr28czEp;G0dQJXrmivL;WLg2OTJuuAL;KbF7faPyL3O;bedEKj2MvAp;yhgoaX6I5SI;M6radndNMcI;jVQfAT1i6s9;WMZn1kseuAQ;wn0nkF3Qy4Z;Y9AnAmH63gR;JDUIlfZsvNR;uSgb62avNGS;T3IjjSH3Qoi;oaTFQGtXVUE;prfHdt6Iwfm;G6qCaD1HZB5;DpFwGPrnM5b;UjfH29cP2nO;Pr5xXF4u6Po;TuecmWZs13b;QgbgKIU8ha6;bD6uStxG20G;EhvHDjy4K3i;Wn0Dl0HNkmL;g3dx0K9C4v9;xlIpRdZDuVD;JRZFbJX8n8e;AGVOZPlKVoY;RzZd4yEq6QU;pS0OjaYEg1J;wW1hykhFepQ;sbcw1HgmdFz;ZXDkDfHug0l;isBiAEvhk18;gRe79HOGjEy;oat5DroeXt5;xPnsEjFx1BM;pMpiFnDRQvV;RHG3dzhAULq;Ewz8cNU1YMb;CQQ6OvfEwX4;bWK3Qvg7oeJ;Fe1x5nNpZ20;c3emJS6Ivsx;wFCYshbrI19;Z5o7he63Us4;HxXEZp5pM1t;lqwRTBbmf7r;oNH46nmDgPO;TLQBZazjlyP;TiVspxdO2Nv;QvMMyrUY6Sg;Gfb7iLhjkvk;eV1Jk45moHc;VkobvrJtE8K;z6WZZFeeWt7;AxecPb6xcCB;RjhuJBWMRmz;oMDsrhbICit;wYCXwqqlmIX',
                'ou:zQUKoh5WmJt;OU_GROUP-fwkewapqBD3',
                'pe:202406'
            ],
            'displayProperty': 'NAME',
            'hierarchyMeta': 'true'
        }

        response = requests.get(indicators_url, auth=HTTPBasicAuth(username, password), params=params)
        response.raise_for_status() 

        data = response.json().get('rows', [])

        with open('indicators_values.csv', mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Indicator', 'Value']) 
            for row in data:
                writer.writerow(row) 

        print("Indicators and values saved to indicators_values.csv")

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
        if response.status_code == 409:
            print(f"Conflict error for URL: {response.url}")
            print(f"Response content: {response.content}")
    except requests.exceptions.RequestException as req_err:
        print(f"Request error occurred: {req_err}")
    except Exception as e:
        print(f"An error occurred: {e}")

def fetch_data(url):
    try:
        response = requests.get(url)
        response.raise_for_status() 
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        if response.status_code == 409:
            print(f"HTTP error occurred: {response.status_code} Conflict for url: {url}")
        else:
            print(f"HTTP error occurred: {http_err}")
    except Exception as err:
        print(f"Other error occurred: {err}")

if __name__ == "__main__":
    list_pivot_tables()
    list_indicators_and_save_to_csv()