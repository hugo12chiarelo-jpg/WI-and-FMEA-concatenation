import glob
import os
import sys
import json
import requests
import pandas as pd
from openai import OpenAI

# --- Configuration ---
API_KEY_DS = os.environ.get("API_KEY_DS")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")  # Automatically provided in GitHub Actions
REPO = os.environ.get("GITHUB_REPOSITORY")     # Format: "owner/repo"

if not API_KEY_DS:
    print("Error: API_KEY_DS environment variable not set.")
    sys.exit(1)

client = OpenAI(api_key=API_KEY_DS, base_url="https://api.deepseek.com")

# File paths
# Input files must be placed in the data/ folder of this repository:
#   - Work Instructions: data/work instruction/*.xlsx  (all .xlsx files in the folder)
#   - FMEA:              data/FMEA_COCE CENTRIFUGAL COMPRESSOR.xlsx
WI_DIR = "data/work instruction"
FMEA_PATH = "data/FMEA_COCE CENTRIFUGAL COMPRESSOR.xlsx"
OUTPUT_PATH = "results/analysis.json"
ISSUES_LABEL = "failure-mechanism-gap"  # Label for auto-created issues


# --- Helper: Check if issue already exists ---
def issue_exists(title):
    """Return True if an open issue with the given title already exists.

    Paginates through all open issues with the label to avoid false negatives
    when there are more than 100 existing issues.
    """
    if not GITHUB_TOKEN or not REPO:
        return False
    page = 1
    headers = {"Authorization": f"token {GITHUB_TOKEN}"}
    while True:
        url = (
            f"https://api.github.com/repos/{REPO}/issues"
            f"?state=open&labels={ISSUES_LABEL}&per_page=100&page={page}"
        )
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return False
        issues = response.json()
        if not issues:
            break
        if any(issue["title"] == title for issue in issues):
            return True
        page += 1
    return False


# --- Helper: Create GitHub Issue ---
def create_github_issue(title, body, labels=None):
    """Create an issue in the repository using the GitHub API."""
    if not GITHUB_TOKEN or not REPO:
        print("Skipping issue creation: missing GITHUB_TOKEN or REPO")
        return

    if issue_exists(title):
        print(f"↷ Issue already exists, skipping: {title}")
        return

    url = f"https://api.github.com/repos/{REPO}/issues"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }
    data = {
        "title": title,
        "body": body,
        "labels": labels or [ISSUES_LABEL],
    }
    response = requests.post(url, json=data, headers=headers)
    if response.status_code == 201:
        print(f"✓ Issue created: {title}")
    else:
        print(f"✗ Failed to create issue: {response.status_code} - {response.text}")


# --- Load data ---
def load_data():
    try:
        wi_files = [
            f for f in glob.glob(os.path.join(WI_DIR, "*.xlsx"))
            if not os.path.basename(f).startswith("~$")
        ]
        if not wi_files:
            print(f"Error loading data: no .xlsx files found in '{WI_DIR}'")
            sys.exit(1)
        df_wi = pd.concat([pd.read_excel(f) for f in wi_files], ignore_index=True)
        df_fmea = pd.read_excel(FMEA_PATH)
        print(f"✓ Data loaded successfully ({len(wi_files)} Work Instruction file(s)).")
        return df_wi, df_fmea
    except Exception as e:
        print(f"Error loading data: {e}")
        sys.exit(1)


# --- Build prompt for deepseek-reasoner ---
def build_prompt(df_wi, df_fmea):
    wi_sample = df_wi.head(100).to_string()
    fmea_sample = df_fmea.head(100).to_string()

    prompt = f"""
You are a maintenance engineering expert specialized in rotating equipment and ISO 14224 failure classification.

You have two datasets:

1. **Current Maintenance Strategy (Reporting Questions):**
```
{wi_sample}
```

2. **FMEA Failure Mode Tree (from ISO 14224):**
```
{fmea_sample}
```

**Task:**
For each failure mechanism listed in the FMEA, determine whether it is covered by any reporting question in the strategy.

**Coverage classification:**
- "Covered": A question directly or functionally detects or prevents that failure mechanism.
- "Partial": A question indirectly indicates the mechanism but not reliably or not directly.
- "Not covered": No question addresses that mechanism.

**Output format:** Return a JSON object with an array called "analysis". Each element must have:
- "component": the maintainable item from FMEA
- "symptom": the symptom (if available)
- "failure_mechanism": the mechanism name
- "coverage": one of "Covered", "Partial", "Not covered"
- "justification": a short explanation

Only output valid JSON, no extra text.
"""
    return prompt


# --- Call DeepSeek Reasoner API ---
def call_deepseek(prompt):
    try:
        response = client.chat.completions.create(
            model="deepseek-reasoner",
            messages=[
                {
                    "role": "system",
                    "content": "You are a maintenance reliability expert. Output only valid JSON.",
                },
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
        )
        print("✓ DeepSeek Reasoner analysis complete.")
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error calling DeepSeek API: {e}")
        sys.exit(1)


# --- Save results and create issues ---
def process_results(json_str):
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    try:
        data = json.loads(json_str)
        with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        print(f"✓ Results saved to {OUTPUT_PATH}")

        # Identify uncovered mechanisms
        uncovered = [
            item
            for item in data.get("analysis", [])
            if item.get("coverage") == "Not covered"
        ]
        print(f"Found {len(uncovered)} uncovered failure mechanisms.")

        # Create an issue for each uncovered mechanism (skip duplicates)
        for item in uncovered:
            title = f"[GAP] {item['component']} - {item['failure_mechanism']}"
            body = f"""## Uncovered Failure Mechanism

**Component:** {item['component']}
**Symptom:** {item.get('symptom', 'N/A')}
**Failure Mechanism:** {item['failure_mechanism']}

**Justification from analysis:**
{item['justification']}

**Recommendation:** Consider adding a reporting question or inspection to detect this mechanism.
"""
            create_github_issue(title, body)

    except json.JSONDecodeError:
        print("Error: API response is not valid JSON.")
        print("Response received:")
        print(json_str)
        sys.exit(1)


# --- Main ---
def main():
    print("--- FMEA Coverage Analysis with DeepSeek Reasoner ---")
    df_wi, df_fmea = load_data()
    prompt = build_prompt(df_wi, df_fmea)
    result_json = call_deepseek(prompt)
    process_results(result_json)
    print("--- Analysis finished ---")


if __name__ == "__main__":
    main()
