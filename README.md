# WI-and-FMEA-concatenation
This agent will merge the informations from FMEA and Work Instruction

## Input files

Place the input files in the `data/` folder of this repository before running the analysis:

| File | Description |
|------|-------------|
| `data/work_instruction.xlsx` | Work Instruction (reporting questions / maintenance strategy) |
| `data/fmea.xlsx` | FMEA Failure Mode Tree (ISO 14224) |

## Configuration

> ⚠️ **The workflow will fail if `API_KEY_DS` is not configured.**

The workflow uses the repository secret **`API_KEY_DS`** as the DeepSeek API key.  
Follow these steps to add it:

1. Go to your repository on GitHub.
2. Click **Settings** → **Secrets and variables** → **Actions**.
3. Click **New repository secret**.
4. Set the name to `API_KEY_DS` and paste your DeepSeek API key as the value.
5. Click **Add secret**.

Once the secret is saved, re-run the workflow — the `Error: API_KEY_DS environment variable not set.` error will no longer appear.
