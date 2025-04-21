import requests
import pandas as pd
import urllib3

# === DISABLE SSL WARNINGS ===
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# === CONFIGURATION START ===
BASE_URL = "https://your-jira-datacenter-instance"
API_TOKEN = "your-api-token-here"
EXCEL_INPUT_FILE = "projects.xlsx"           # Input Excel file with project keys
EXCEL_OUTPUT_FILE = "jira_issue_counts.xlsx" # Output Excel file
# === CONFIGURATION END ===

# === AUTH HEADERS ===
headers = {
    "Accept": "application/json",
    "Authorization": f"Bearer {API_TOKEN}"
}

def get_issue_count(base_url, headers, project_key):
    url = f"{base_url}/rest/api/2/search?jql=project={project_key}&maxResults=0"
    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()
    return response.json().get("total", 0)

def main():
    try:
        df_input = pd.read_excel(EXCEL_INPUT_FILE)
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return

    if "project_key" not in df_input.columns:
        print("❌ Input file must contain a column named 'project_key'")
        return

    results = []
    for project_key in df_input["project_key"]:
        print(f"🔍 Fetching issue count for project: {project_key}")
        try:
            count = get_issue_count(BASE_URL, headers, project_key)
        except Exception as e:
            print(f"⚠️ Error fetching issues for {project_key}: {e}")
            count = "Error"
        results.append({
            "project_key": project_key,
            "issue_count": count
        })

    df_output = pd.DataFrame(results)
    try:
        df_output.to_excel(EXCEL_OUTPUT_FILE, index=False)
        print(f"✅ Results saved to {EXCEL_OUTPUT_FILE}")
    except Exception as e:
        print(f"❌ Error writing to output file: {e}")

if __name__ == "__main__":
    main()
