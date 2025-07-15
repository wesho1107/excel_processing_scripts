# Excel to Markdown Converter - AI-Assisted Workflow

A command-line tool designed to work with AI assistants for analyzing and converting Excel files. This tool implements a two-step workflow where AI handles the analysis and decision-making, while Python scripts handle the data processing.

## Workflow Overview

### Step 1: User provides Excel file
User provides an Excel file (like the examples provided).

### Step 2: Convert sheets to markdown (Python Script)
The AI assistant uses the Python script to convert all Excel sheets into markdown files for analysis.

### Step 3: AI analyzes structure (AI Direct)
The AI assistant reads the markdown files and understands the structure without scripting.

### Step 4: AI identifies data elements (AI Direct)
The AI identifies:
- Column headers
- Starting cell of the actual data table
- Unique primary key of each row
- Content and data types for summary

### Step 5: AI generates filtered data (Python Script)
The AI assistant generates markdown files that filter data by rows and columns using the Python script with parameters it provides based on its analysis.

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Step 2: Convert All Sheets to Markdown

```bash
python scripts/excel/excel_processor.py input.xlsx --convert-all
```

This will:
- Convert all sheets in the Excel file to markdown format
- Save each sheet as a separate `.md` file in the `sheets_markdown/` directory
- Include cell references for precise AI analysis
- Provide metadata about the structure

### Step 5: Filter Data Based on AI Analysis

```bash
python scripts/excel/excel_processor.py input.xlsx -s "A3" -c "Name" "Job" -pc "Department" -k "Engineering" -o "engineering.md"
```

Parameters:
- `-s "A3"`: Starting cell of the data table (AI-identified)
- `-c "Name" "Job"`: Columns to extract (AI-selected)
- `-pc "Department"`: Primary column to filter by (AI-determined)
- `-k "Engineering"`: Key value to filter for (AI-specified)
- `-o "engineering.md"`: Output file path

## Command-Line Options

### For Step 2 (Convert All Sheets):
- `--convert-all`: Convert all sheets to markdown files for AI analysis
- `-o OUTPUT_DIR`: Specify output directory (default: "sheets_markdown")

### For Step 5 (Filter Data):
- `-s START_CELL`: Starting cell of data table (e.g., "A6") **[Required]**
- `-c COLUMNS`: List of column names to extract (space-separated) **[Optional - defaults to all columns]**
- `-pc PRIMARY_COLUMN`: Primary column to filter by **[Required]**
- `-k KEY_VALUE`: Value to filter for in primary column **[Required]**
- `-o OUTPUT`: Output markdown file path **[Required]**
- `--sheet SHEET_NAME`: Specific sheet to process (optional)

**Note**: If `-c` is not specified, all available columns will be included in the output.

## AI-Assisted Workflow Examples

### Complete Workflow Example:

1. **Step 2: Convert to markdown**
```bash
python scripts/excel/excel_processor.py employee_data.xlsx --convert-all
```

2. **Steps 3-4: AI analyzes the generated markdown files** (done by AI directly)

3. **Step 5: AI calls the script with identified parameters**
```bash
# AI determines these parameters from analysis:
python scripts/excel/excel_processor.py employee_data.xlsx -s "A6" -c "Name" "Job" "Email" -pc "Department" -k "Engineering" -o "engineering_team.md"
```

### Multiple Filter Example:
```bash
# Filter for specific columns
python scripts/excel/excel_processor.py data.xlsx -s "A6" -c "Name" "Job" -pc "Department" -k "Marketing" -o "marketing_team.md"

# Filter for all columns (omit -c parameter)
python scripts/excel/excel_processor.py data.xlsx -s "A6" -pc "Department" -k "Sales" -o "sales_team_all_columns.md"
```

## Output Examples

### Step 2 Output (Sheet Analysis):
```markdown
# Sheet: Employee Data

## Raw Data with Cell References

| Col 1 | Col 2 | Col 3 | Col 4 |
| --- | --- | --- | --- |
| Company Report (A1) | | | |
| Employee Data (A2) | | | |
| | | | |
| Name (A6) | Age (B6) | Job (C6) | Department (D6) |
| John Doe (A7) | 30 (B7) | Software Engineer (C7) | Engineering (D7) |
| Jane Smith (A8) | 28 (B8) | Data Analyst (C8) | Analytics (D8) |

## Summary for AI Analysis
- Total rows with data: 5
- Total columns: 4
- Sheet name: Employee Data
- Cell references are included in parentheses for precise identification
```

### Step 5 Output (Filtered Data):
```markdown
# Filtered Data: Department = 'Engineering'

## Metadata
- Filter criteria: Department = 'Engineering'
- Data source starting cell: A6
- Total matching rows: 3
- Columns included: Name, Job, Email

## Data Table

| Name | Job | Email |
| --- | --- | --- |
| John Doe | Software Engineer | john@company.com |
| Alice Brown | UX Designer | alice@company.com |
| Charlie Wilson | DevOps Engineer | charlie@company.com |
```

## AI Assistant Instructions

When using this tool, the AI assistant should:

1. **For Step 2**: Call the script with `--convert-all` to generate markdown files
2. **For Steps 3-4**: Analyze the generated markdown files to identify:
   - Column headers and their positions
   - Starting cell of actual data (e.g., "A6")
   - Primary key columns for filtering
   - Data types and structure
3. **For Step 5**: Call the script with specific parameters based on analysis:
   - Use exact cell references identified in Step 4
   - Select appropriate columns for the user's needs
   - Choose meaningful filter criteria

## Legacy Mode (Deprecated)

The tool still supports the old workflow for backward compatibility:

```bash
python scripts/excel/excel_processor.py input.xlsx -c "Name" "Job" -o output.md --legacy
```

However, the AI-assisted workflow is recommended for better results.

## Requirements

- Python 3.6+
- pandas>=1.3.0
- openpyxl>=3.0.0
