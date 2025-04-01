import json
import pandas as pd
import requests
import openpyxl
from openpyxl import Workbook
from pprint import pprint
import urllib.parse

# ______________________________________________________________________________________________________

##### Para resolver o limit itens 1000, para uma necessidade de 1500 itens, foram configuradas duas chamadas, uma com "start_at_index":0 e outra "start_at_index":1000
##### "start_at_index":0 #####
raw_data = {"query":{"models":["Alert"],"type":"object_set","with":{"operator":"and","type":"operation","values":[{"key":"Status","values":["open","in_progress"],"type":"str","operator":"in"},{"key":"RiskLevel","values":["critical","high"],"type":"str","operator":"in"}]}},"standard_format":False,"with_model_names":True,"get_results_and_count":True,"order_by_pk":False,"additional_models[]":["CloudAccount","CodeOrigins","Inventory"],"enable_pagination":True,"limit":1000,"start_at_index":0,"order_by[]":["-OrcaScore"],"max_tier":2}
headers = {"Authorization": "Token "}
#response = requests.get("https://api.orcasecurity.io/api/alerts?state.risk_level=critical,high", headers=headers)
response = requests.post("https://api.orcasecurity.io/api/sonar/query", headers=headers, json=raw_data)

if response.status_code == 200:
    new_data = response.json()

    try:
        with open("data1.json", "r") as json_file:
            existing_data = json.load(json_file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        existing_data = []

    existing_data.append(new_data)

    with open("data1.json", "w") as json_file:
        json.dump(existing_data, json_file, indent=4)
        print("Arquivo data.json criado com sucesso!")
else:
    print("Failed to retrieve data from the API. Status code:", response.status_code)
# ________________________________________________________________________________________________________

json_file = "data1.json"
xlsx_file = "data1.xlsx"

with open(json_file, "r", encoding="utf-8") as file:
    json_data = json.load(file)

records = json_data[0].get("data", [])  # Pegamos a lista de registros
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Alertas"
headers = ["Nome_Conta", "Titulo", "Descricao", "Severidade"]
ws.append(headers)

for record in records:
    ws.append([
        record.get("data", {}).get("CloudAccount", {}).get("name", "N/A"),
        record.get("data", {}).get("AlertType", {}).get("value", "N/A"),
        record.get("data", {}).get("Inventory", {}).get("name", "N/A"),
        record.get("data", {}).get("RiskLevel", {}).get("value", "N/A")
    ])

#for record in records:
#    ws.append([
#        record.get("account_name", "N/A"),
#        record.get("description", "N/A"),
#        record.get("details", "N/A"),
#        record.get("state", {}).get("risk_level", "N/A")  # Lidando com JSON aninhado
#    ])

wb.save(xlsx_file)
print(f"Arquivo {xlsx_file} criado com sucesso!")

#____________________________________________________________________________________________________________

##### "start_at_index":1000 #####
raw_data = {"query":{"models":["Alert"],"type":"object_set","with":{"operator":"and","type":"operation","values":[{"key":"Status","values":["open","in_progress"],"type":"str","operator":"in"},{"key":"RiskLevel","values":["critical","high"],"type":"str","operator":"in"}]}},"standard_format":False,"with_model_names":True,"get_results_and_count":True,"order_by_pk":False,"additional_models[]":["CloudAccount","CodeOrigins","Inventory"],"enable_pagination":True,"limit":1000,"start_at_index":1000,"order_by[]":["-OrcaScore"],"max_tier":2}
#response = requests.get("https://api.orcasecurity.io/api/alerts?state.risk_level=critical,high", headers=headers)
response = requests.post("https://api.orcasecurity.io/api/sonar/query", headers=headers, json=raw_data)

if response.status_code == 200:
    new_data = response.json()

    try:
        with open("data2.json", "r") as json_file:
            existing_data = json.load(json_file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        existing_data = []

    existing_data.append(new_data)

    with open("data2.json", "w") as json_file:
        json.dump(existing_data, json_file, indent=4)
        print("Arquivo data.json criado com sucesso!")
else:
    print("Failed to retrieve data from the API. Status code:", response.status_code)

# ________________________________________________________________________________________________________

json_file = "data2.json"
xlsx_file = "data2.xlsx"

with open(json_file, "r", encoding="utf-8") as file:
    json_data = json.load(file)

records = json_data[0].get("data", [])  # Pegamos a lista de registros
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Alertas"
headers = ["Nome_Conta", "Titulo", "Descricao", "Severidade"]
ws.append(headers)

for record in records:
    ws.append([
        record.get("data", {}).get("CloudAccount", {}).get("name", "N/A"),
        record.get("data", {}).get("AlertType", {}).get("value", "N/A"),
        record.get("data", {}).get("Inventory", {}).get("name", "N/A"),
        record.get("data", {}).get("RiskLevel", {}).get("value", "N/A")
    ])

wb.save(xlsx_file)
print(f"Arquivo {xlsx_file} criado com sucesso!")

df1 = pd.read_excel("data1.xlsx")
df2 = pd.read_excel("data2.xlsx")
df_concatenado = pd.concat([df1, df2], ignore_index=True)
df_concatenado.to_excel("data3.xlsx", index=False)
