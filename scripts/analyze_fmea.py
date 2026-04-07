import glob
import os
import sys
import json
import time
import pandas as pd
from openai import OpenAI, APIConnectionError, APIStatusError, RateLimitError
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Configuration ---
API_KEY_DS = os.environ.get("API_KEY_DS")

if not API_KEY_DS:
    print("Error: API_KEY_DS environment variable not set.")
    sys.exit(1)

client = OpenAI(api_key=API_KEY_DS, base_url="https://api.deepseek.com", timeout=120.0)

# File paths
# Input files must be placed in the data/ folder of this repository:
#   - Work Instructions: data/work instruction/*.xlsx  (all .xlsx files in the folder)
#   - FMEA:              any single .xlsx file in the data/ folder (not inside a subdirectory)
WI_DIR = "data/work instruction"
DATA_DIR = "data"
OUTPUT_PATH = "results/analysis.xlsx"

# FMEA spreadsheet column names that hold the maintainable item and failure mechanism
FMEA_COL_ITEM = "Unnamed: 6"
FMEA_COL_MECHANISM = "Unnamed: 7"

# Coverage colour fills
FILL_COVERED = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
FILL_PARTIAL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
FILL_NOT_COVERED = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")

# Header colour fill (teal, matching the example image)
FILL_HEADER = PatternFill(start_color="1F6B75", end_color="1F6B75", fill_type="solid")
FONT_HEADER = Font(bold=True, color="FFFFFF")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# --- Locate FMEA file dynamically ---
def find_fmea_file():
    """Return the path to the single FMEA .xlsx file located directly inside DATA_DIR."""
    candidates = [
        f for f in glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        print(f"Error: no .xlsx FMEA file found in '{DATA_DIR}'")
        sys.exit(1)
    if len(candidates) > 1:
        print(f"Warning: multiple .xlsx files found in '{DATA_DIR}', using: {candidates[0]}")
    return candidates[0]


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

        fmea_path = find_fmea_file()
        df_fmea = pd.read_excel(fmea_path)
        print(
            f"✓ Data loaded: FMEA '{os.path.basename(fmea_path)}', "
            f"{len(wi_files)} Work Instruction file(s)."
        )
        return df_wi, df_fmea
    except SystemExit:
        raise
    except Exception as e:
        print(f"Error loading data: {e}")
        sys.exit(1)


# --- Extract FMEA component/mechanism pairs ---
def extract_fmea_rows(df_fmea):
    """Parse the FMEA spreadsheet and return a list of dicts with
    'maintainable_item' and 'failure_mechanism', skipping header/total rows."""
    rows = []
    current_item = None
    for _, row in df_fmea.iterrows():
        item_val = row.get(FMEA_COL_ITEM)
        mech_val = row.get(FMEA_COL_MECHANISM)

        # Skip pure header row
        if str(item_val).strip() == "Maintainable Item":
            continue

        # A non-null item cell starts a new component group (ignore "* Total" rows)
        if pd.notna(item_val):
            label = str(item_val).strip()
            if "Total" in label or label == "Grand Total":
                current_item = None
                continue
            current_item = label

        # A mechanism cell with a valid current component is a data row
        if current_item and pd.notna(mech_val):
            rows.append(
                {
                    "maintainable_item": current_item,
                    "failure_mechanism": str(mech_val).strip(),
                }
            )
    return rows


# --- Build prompt for deepseek-reasoner ---
def build_prompt(df_wi, fmea_rows):
    wi_text = df_wi.to_string()

    fmea_lines = "\n".join(
        f"- {r['maintainable_item']} | {r['failure_mechanism']}"
        for r in fmea_rows
    )

    prompt = f"""
You are a maintenance engineering expert specialized in rotating equipment and ISO 14224 failure classification.

You have two datasets:

1. **Work Instructions – Reporting Questions (all rows):**
```
{wi_text}
```

2. **FMEA – All Maintainable Items and Failure Mechanisms to analyse:**
```
{fmea_lines}
```

**Task:**
For EVERY failure mechanism listed above, determine whether it is covered by any reporting question
in the Work Instructions. You MUST return one entry per line in the FMEA list — do not skip any.

**Coverage classification rules:**
- "Covered": A reporting question directly or functionally identifies, detects, or prevents that
  failure mechanism. This includes: a question that explicitly mentions the phenomenon (e.g.,
  "corrosion" → covers Corrosion), a question linked to a component that is functionally
  responsible for the mechanism, or a question that monitors a symptom that is a direct
  consequence of the mechanism.
- "Partial": A question may indicate the mechanism indirectly or unreliably — e.g., a generic
  vibration question that could relate to multiple mechanisms, or a question on a nearby component
  rather than the exact component.
- "Not covered": No reporting question, even indirectly, can detect or prevent the mechanism.

**Output format:** Return a JSON object with a top-level array called "analysis".
Each element MUST have exactly these fields:
- "maintainable_item": the component name (from the FMEA list)
- "failure_mechanism": the failure mechanism name (from the FMEA list)
- "coverage": one of "Covered", "Partial", "Not covered"
- "note": a brief justification in the same language as the Work Instructions data
- "question_ids": a list of WTT IDs or question task numbers that support the coverage
  (empty list [] if Not covered)

Only output valid JSON, no extra text, no markdown fences.
"""
    return prompt


# --- Call DeepSeek Reasoner API ---
def call_deepseek(prompt, max_retries=3):
    """Call the DeepSeek Reasoner API with retry/backoff on transient errors.

    Note: ``deepseek-reasoner`` does NOT support the ``response_format``
    parameter, so JSON is requested via the prompt only.
    """
    for attempt in range(1, max_retries + 1):
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
            )
            print("✓ DeepSeek Reasoner analysis complete.")
            return response.choices[0].message.content
        except (APIConnectionError, RateLimitError) as e:
            if attempt < max_retries:
                wait = 2 ** attempt  # exponential backoff: 2 s, 4 s, 8 s (attempts 1–3)
                print(f"⚠ DeepSeek API error (attempt {attempt}/{max_retries}): {e}. Retrying in {wait}s…")
                time.sleep(wait)
            else:
                print(f"Error calling DeepSeek API: {e}")
                sys.exit(1)
        except APIStatusError as e:
            print(f"Error calling DeepSeek API: HTTP {e.status_code} - {e.message}")
            sys.exit(1)
        except Exception as e:
            print(f"Error calling DeepSeek API: {e}")
            sys.exit(1)


# --- Save results as Excel workbook ---
def save_excel(analysis_items):
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "FMEA Coverage Analysis"

    headers = [
        "Maintainable Item (FMEA)",
        "Failure Mechanism",
        "Coverage",
        "Note",
        "Reporting Question ID identification",
    ]

    # Write and style header row
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 30

    # Write data rows
    for row_idx, item in enumerate(analysis_items, start=2):
        coverage = item.get("coverage", "")
        question_ids = item.get("question_ids", [])
        ids_str = ", ".join(str(q) for q in question_ids) if question_ids else ""

        values = [
            item.get("maintainable_item", ""),
            item.get("failure_mechanism", ""),
            coverage,
            item.get("note", ""),
            ids_str,
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

        # Apply coverage colour to the Coverage cell (column C)
        cov_cell = ws.cell(row=row_idx, column=3)
        if coverage == "Covered":
            cov_cell.fill = FILL_COVERED
            cov_cell.font = Font(bold=True, color="FFFFFF")
        elif coverage == "Partial":
            cov_cell.fill = FILL_PARTIAL
            cov_cell.font = Font(bold=True, color="000000")
        elif coverage == "Not covered":
            cov_cell.fill = FILL_NOT_COVERED
            cov_cell.font = Font(bold=True, color="FFFFFF")

        cov_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Set column widths
    col_widths = [30, 35, 15, 60, 35]
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze the header row
    ws.freeze_panes = "A2"

    # Auto-filter on all columns
    ws.auto_filter.ref = ws.dimensions

    wb.save(OUTPUT_PATH)
    print(f"✓ Results saved to {OUTPUT_PATH}")


# --- Parse API response and save ---
def process_results(json_str, fmea_rows):
    try:
        # Strip optional markdown code fences the model may add (e.g. ```json … ```)
        cleaned = json_str.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("\n", 1)[-1]
            if cleaned.endswith("```"):
                cleaned = cleaned.rsplit("```", 1)[0]

        data = json.loads(cleaned)
        analysis = data.get("analysis", [])

        # Ensure every FMEA row is represented; add placeholder if AI missed any
        ai_keys = {
            (str(e.get("maintainable_item", "")).strip(),
             str(e.get("failure_mechanism", "")).strip())
            for e in analysis
        }
        for fmea_row in fmea_rows:
            key = (str(fmea_row["maintainable_item"]).strip(),
                   str(fmea_row["failure_mechanism"]).strip())
            if key not in ai_keys:
                analysis.append(
                    {
                        "maintainable_item": fmea_row["maintainable_item"],
                        "failure_mechanism": fmea_row["failure_mechanism"],
                        "coverage": "Not covered",
                        "note": "No analysis returned by AI for this mechanism.",
                        "question_ids": [],
                    }
                )

        covered = sum(1 for e in analysis if e.get("coverage") == "Covered")
        partial = sum(1 for e in analysis if e.get("coverage") == "Partial")
        not_covered = sum(1 for e in analysis if e.get("coverage") == "Not covered")
        print(
            f"Summary: {len(analysis)} mechanisms — "
            f"{covered} Covered, {partial} Partial, {not_covered} Not covered."
        )

        save_excel(analysis)

    except json.JSONDecodeError:
        print("Error: API response is not valid JSON.")
        print("Response received:")
        print(json_str)
        sys.exit(1)


# --- Main ---
def main():
    print("--- FMEA Coverage Analysis with DeepSeek Reasoner ---")
    df_wi, df_fmea = load_data()
    fmea_rows = extract_fmea_rows(df_fmea)
    print(f"✓ Extracted {len(fmea_rows)} FMEA component/mechanism pairs.")
    prompt = build_prompt(df_wi, fmea_rows)
    result_json = call_deepseek(prompt)
    process_results(result_json, fmea_rows)
    print("--- Analysis finished ---")


if __name__ == "__main__":
    main()
