import requests
import os

# Configuration
access_token = "<Paste_Your_Access_Token>"
output_path = "C:\\Path\\To\\Save\\Reports"
os.makedirs(output_path, exist_ok=True)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# Get workspaces
workspaces_url = "https://api.powerbi.com/v1.0/myorg/groups"
response = requests.get(workspaces_url, headers=headers)
if response.status_code != 200:
    print(f"Failed to get workspaces: {response.status_code} - {response.text}")
    exit()
workspaces = response.json().get("value", [])

for workspace in workspaces:
    workspace_id = workspace["id"]
    workspace_name = workspace["name"]
    print(f"Workspace: {workspace_name}")

    # Get datasets in workspace
    datasets_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
    response = requests.get(datasets_url, headers=headers)
    if response.status_code != 200:
        print(f"  Failed to get datasets for {workspace_name}: {response.status_code}")
        continue
    datasets = response.json().get("value", [])

    for dataset in datasets:
        dataset_id = dataset["id"]
        dataset_name = dataset["name"]
        print(f"  Semantic Model: {dataset_name}")

        # Get reports for the dataset
        reports_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports"
        response = requests.get(reports_url, headers=headers)
        if response.status_code != 200:
            print(f"    Failed to get reports for {dataset_name}: {response.status_code}")
            continue
        reports = response.json().get("value", [])
        reports = [r for r in reports if r["datasetId"] == dataset_id]

        for report in reports:
            report_id = report["id"]
            report_name = report["name"]
            export_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{report_id}/Export"
            try:
                response = requests.get(export_url, headers=headers, stream=True)
                if response.status_code == 200:
                    file_path = os.path.join(output_path, f"{report_name}.pbix")
                    with open(file_path, "wb") as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    print(f"    Exported: {report_name}.pbix")
                else:
                    print(f"    Failed to export {report_name}: {response.status_code} - {response.text}")
            except Exception as e:
                print(f"    Error exporting {report_name}: {str(e)}")