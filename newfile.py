import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import urllib3

# === DISABLE SSL WARNINGS ===
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# === CONFIGURATION START ===
BASE_URL = "https://your-jira-datacenter-instance"
USERNAME = "your-jira-username"
API_TOKEN = "your-api-token-or-password"
EXCEL_INPUT_FILE = "projects.xlsx"
EXCEL_OUTPUT_FILE = "jira_issue_counts_output.xlsx"
# === CONFIGURATION END ===

def get_issue_count(base_url, headers, auth, project_key):
    url = f"{base_url}/rest/api/2/search?jql=project={project_key}&maxResults=0"
    response = requests.get(url, headers=headers, auth=auth, verify=False)
    response.raise_for_status()
    return response.json().get("total", 0)

def main():
    headers = {"Accept": "application/json"}
    auth = HTTPBasicAuth(USERNAME, API_TOKEN)

    # Read project keys from Excel
    try:
        df_input = pd.read_excel(EXCEL_INPUT_FILE)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if "project_key" not in df_input.columns:
        print("Error: Excel sheet must have a 'project_key' column.")
        return

    results = []
    for project_key in df_input["project_key"]:
        print(f"Fetching issue count for project: {project_key}")
        try:
            total = get_issue_count(BASE_URL, headers, auth, project_key)
        except Exception as e:
            print(f"Failed to fetch issues for {project_key}: {e}")
            total = "Error"
        results.append({"project_key": project_key, "issue_count": total})

    # Save to output Excel
    df_output = pd.DataFrame(results)
    try:
        df_output.to_excel(EXCEL_OUTPUT_FILE, index=False)
        print(f"Issue counts saved to {EXCEL_OUTPUT_FILE}")
    except Exception as e:
        print(f"Error writing to output file: {e}")

if __name__ == "__main__":
    main()
