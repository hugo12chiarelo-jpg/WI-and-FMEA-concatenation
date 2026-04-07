# WI-and-FMEA-concatenation
This agent will merge the informations from FMEA and Work Instruction

## Input files

Place the input files in the `data/` folder of this repository before running the analysis:

| File | Description |
|------|-------------|
| `data/work_instruction.xlsx` | Work Instruction (reporting questions / maintenance strategy) |
| `data/fmea.xlsx` | FMEA Failure Mode Tree (ISO 14224) |

## Configuration

The workflow uses the repository secret **`API_KEY_DS`** as the DeepSeek API key.  
Make sure to add this secret under *Settings → Secrets and variables → Actions* in your repository.
