import json
import pandas as pd
import requests
import openpyxl
from openpyxl import Workbook
from pprint import pprint
import urllib.parse
import os

# start_at_index": 0 ________________________________________________________________________________________

raw_data = {
    "query": {
    "models": [
      "AliCloudEcsInstance",
      "AwsEc2Instance",
      "AzureComputeVm",
      "GcpVmInstance",
      "OciComputeVmInstance",
      "OnPremVm",
      "Vm",
      "VmwareVmInstance",
      "TencentCloudCvmInstance",
      "Aci",
      "CloudRun",
      "Container",
      "AwsEc2Image",
      "AzureAcrImage",
      "AzureComputeImage",
      "GcpVmImage",
      "JFrogArtifactoryContainerImage",
      "VmImage",
      "AwsSagemakerImage",
      "DockerHubContainerImage",
      "AwsLambdaFunction",
      "AwsLambdaLayer",
      "AwsServerlessApplicationRepositoryApplication",
      "AwsStepFunctionsStateMachine",
      "AzureEventGridSubscription",
      "AzureEventGridTopic",
      "AzureEventHub",
      "AzureEventHubNamespace",
      "AzureFunction",
      "AzureFunctionApp",
      "Function",
      "GcpCloudFunction",
      "ContainerImageSpec",
      "ImageRegistrySpec",
      "ImageRepositorySpec"
    ],
    "type": "object_set"
    },
        "limit": 10000,
        "start_at_index": 0,
        "order_by[]": [
            "Vm.Type"
        ],
        "select": [
            "Name",
            "CloudAccount.Name",
            "CloudAccount.CloudProvider",    
            "Targets.Name",
            "ComputeVms.Name",
            "Tags",
            "NewCategory",
            "NewSubCategory",    
            "ConsoleUrlLink"
        ],
        "get_results_and_count": False,
        "full_graph_fetch": {
        "enabled": True
        },
    "max_tier": 2
    }

headers = {"Authorization": "Token "}
response = requests.post("https://api.orcasecurity.io/api/serving-layer/query", headers=headers, json=raw_data)

if response.status_code == 200:
    new_data = response.json()

    try:
        with open("inventario1.json", "r") as json_file:
            existing_data = json.load(json_file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        existing_data = []

    existing_data.append(new_data)

    with open("inventario1.json", "w") as json_file:
        json.dump(existing_data, json_file, indent=4)
        print("Arquivo inventario1.json criado com sucesso!")
else:
    print("Failed to retrieve data from the API. Status code:", response.status_code)

json_file = "inventario1.json"
xlsx_file = "inventario1.xlsx"

with open(json_file, "r", encoding="utf-8") as file:
    json_data = json.load(file)

records = json_data[0].get("data", [])  # Pegamos a lista de registros
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Inventario"
headers = ["CloudAccount_Name", "CloudAccount_ID", "Asset_Name", "Asset_Type"]
ws.append(headers)

for record in records:
    ws.append([
        record.get("data", {}).get("CloudAccount", {}).get("name", "N/A"),
        record.get("data", {}).get("CloudAccount", {}).get("id", "N/A"),
        record.get("data", {}).get("Name", {}).get("value", "N/A"),
        record.get("data", {}).get("Type", {}).get("value", "N/A")
    ])

wb.save(xlsx_file)
print(f"Arquivo {xlsx_file} criado com sucesso!")

# start_at_index": 10000 ________________________________________________________________________________________

raw_data = {
    "query": {
    "models": [
      "AliCloudEcsInstance",
      "AwsEc2Instance",
      "AzureComputeVm",
      "GcpVmInstance",
      "OciComputeVmInstance",
      "OnPremVm",
      "Vm",
      "VmwareVmInstance",
      "TencentCloudCvmInstance",
      "Aci",
      "CloudRun",
      "Container",
      "AwsEc2Image",
      "AzureAcrImage",
      "AzureComputeImage",
      "GcpVmImage",
      "JFrogArtifactoryContainerImage",
      "VmImage",
      "AwsSagemakerImage",
      "DockerHubContainerImage",
      "AwsLambdaFunction",
      "AwsLambdaLayer",
      "AwsServerlessApplicationRepositoryApplication",
      "AwsStepFunctionsStateMachine",
      "AzureEventGridSubscription",
      "AzureEventGridTopic",
      "AzureEventHub",
      "AzureEventHubNamespace",
      "AzureFunction",
      "AzureFunctionApp",
      "Function",
      "GcpCloudFunction",
      "ContainerImageSpec",
      "ImageRegistrySpec",
      "ImageRepositorySpec"
    ],
    "type": "object_set"
    },
        "limit": 10000,
        "start_at_index": 10000,
        "order_by[]": [
            "Vm.Type"
        ],
        "select": [
            "Name",
            "CloudAccount.Name",
            "CloudAccount.CloudProvider",    
            "Targets.Name",
            "ComputeVms.Name",
            "Tags",
            "NewCategory",
            "NewSubCategory",    
            "ConsoleUrlLink"
        ],
        "get_results_and_count": False,
        "full_graph_fetch": {
        "enabled": True
        },
    "max_tier": 2
    }

headers = {"Authorization": "Token "}
response = requests.post("https://api.orcasecurity.io/api/serving-layer/query", headers=headers, json=raw_data)

if response.status_code == 200:
    new_data = response.json()

    try:
        with open("inventario2.json", "r") as json_file:
            existing_data = json.load(json_file)
    except (FileNotFoundError, json.decoder.JSONDecodeError):
        existing_data = []

    existing_data.append(new_data)

    with open("inventario2.json", "w") as json_file:
        json.dump(existing_data, json_file, indent=4)
        print("Arquivo inventario2.json criado com sucesso!")
else:
    print("Failed to retrieve data from the API. Status code:", response.status_code)

json_file = "inventario2.json"
xlsx_file = "inventario2.xlsx"

with open(json_file, "r", encoding="utf-8") as file:
    json_data = json.load(file)

records = json_data[0].get("data", [])  # Pegamos a lista de registros
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Inventario"
headers = ["CloudAccount_Name", "CloudAccount_ID", "Asset_Name", "Asset_Type"]
ws.append(headers)

for record in records:
    ws.append([
        record.get("data", {}).get("CloudAccount", {}).get("name", "N/A"),
        record.get("data", {}).get("CloudAccount", {}).get("id", "N/A"),
        record.get("data", {}).get("Name", {}).get("value", "N/A"),
        record.get("data", {}).get("Type", {}).get("value", "N/A")
    ])

wb.save(xlsx_file)
print(f"Arquivo {xlsx_file} criado com sucesso!")

# CONCATENA _____________________________________________________________________________

df1 = pd.read_excel("inventario1.xlsx")
df2 = pd.read_excel("inventario2.xlsx")
df_concatenado = pd.concat([df1, df2], ignore_index=True)
df_concatenado.to_excel("inventario.xlsx", index=False)
print(f"Arquivo inventario.xlsx criado com sucesso!")

# EXCLUI TEMP _____________________________________________________________________________

xlsx_file = "inventario1.xlsx"

if os.path.exists(xlsx_file):
    os.remove(xlsx_file)
    print(f"Arquivo {xlsx_file} foi excluído com sucesso.")
else:
    print(f"O arquivo {xlsx_file} não existe.")

xlsx_file = "inventario2.xlsx"

if os.path.exists(xlsx_file):
    os.remove(xlsx_file)
    print(f"Arquivo {xlsx_file} foi excluído com sucesso.")
else:
    print(f"O arquivo {xlsx_file} não existe.")
