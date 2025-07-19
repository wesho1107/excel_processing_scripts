# ID-Based Data Dictionary Processor - AI Assistant Guide

A specialized command-line tool designed for AI assistants to analyze and process Excel files containing data dictionaries with ID-based table and column mappings. This tool implements an AI-assisted workflow optimized for handling complex Excel structures commonly found in data dictionaries.

## Tool Overview

The `id_based_data_dictionary_processor.py` is specifically designed for Excel files that contain:
- Multiple sheets with table definitions and column mappings
- ID-based relationships between tables and columns
- Complex data structures requiring AI analysis before processing
- Data dictionaries with metadata and actual data in separate sections

## Core Capabilities

### 1. **Sheet-to-Markdown Conversion**
Convert all Excel sheets to markdown format for AI analysis with cell references preserved.

### 2. **Batch Filtering with JSON Mapping**
Filter data using JSON configuration files for multiple filter criteria simultaneously.

### 3. **ID-Based Relationship Processing**
Handle Excel files where tables and columns are referenced by IDs rather than direct names.

## AI-Assisted Workflow

### Step 1: Convert All Sheets to Markdown (AI Tool Call)

```bash
python id_based_data_dictionary_processor.py input.xlsx --convert-all
```

**Purpose**: Generate markdown files from all Excel sheets for AI analysis.

**Parameters**:
- `input.xlsx`: Path to the Excel file
- `--convert-all`: Converts all sheets to markdown
- `-o OUTPUT_DIR`: (Optional) Specify output directory (default: "sheets_markdown")

**Output**: Creates individual `.md` files for each Excel sheet with:
- Cell references in parentheses (e.g., "Table Name (A1)")
- Complete data structure preservation
- Metadata for AI analysis

### Step 2: AI Analysis (AI Direct Processing)

The AI assistant analyzes the generated markdown files to identify:
- **Table ID mappings**: Relationships between table IDs and table names
- **Column ID mappings**: Relationships between column IDs and column names  
- **Data start locations**: Exact cell where actual data begins (headers)
- **Primary key columns**: Unique identifiers for filtering
- **Data relationships**: How tables and columns relate to each other

### Step 3: Create JSON Mapping File (AI Tool Call)

Based on analysis, create a JSON file containing the mapping for batch processing:

```json
{
    "table_engineering": "ENG001",
    "table_marketing": "MKT001", 
    "table_sales": "SLS001",
    "user_profile": "USR123"
}
```

**JSON Structure**:
- **Key**: Output filename (without .md extension), this is usually the table name
- **Value**: The ID or value to filter for in the primary column, this is usually the table id

### Step 4: Batch Filter Data (AI Tool Call)

```bash
python id_based_data_dictionary_processor.py input.xlsx -s "A6" -c "Name" "Department" "Role" -pc "TableID" --key-value-mapping "mapping.json" -o "output_dir/"
```

**Required Parameters**:
- `input.xlsx`: Path to the Excel file
- `-s "A6"`: Starting cell of data table (AI-identified from Step 2)
- `-pc "TableID"`: Primary column to filter by (AI-identified)
- `--key-value-mapping "mapping.json"`: Path to JSON mapping file
- `-o "output_dir/"`: Output directory for generated markdown files

**Optional Parameters**:
- `-c "Name" "Department" "Role"`: Specific columns to extract (omit for all columns)
- `--sheet "SheetName"`: Specific sheet to process (omit for default sheet)

## Command-Line Reference

### Convert All Sheets
```bash
python id_based_data_dictionary_processor.py <input_file> --convert-all [-o OUTPUT_DIR]
```

### Batch Filter with JSON Mapping
```bash
python id_based_data_dictionary_processor.py <input_file> -s START_CELL -pc PRIMARY_COLUMN --key-value-mapping JSON_FILE -o OUTPUT_DIR [-c COLUMNS] [--sheet SHEET_NAME]
```

## Parameters Explained

| Parameter | Description | Required | Example |
|-----------|-------------|----------|---------|
| `input_file` | Path to Excel file | ✓ | `data_dictionary.xlsx` |
| `--convert-all` | Convert all sheets to markdown | - | `--convert-all` |
| `-s, --start-cell` | Starting cell of data table | ✓ (for filtering) | `-s "A6"` |
| `-c, --columns` | Columns to extract | - | `-c "Name" "ID" "Type"` |
| `-pc, --primary-column` | Column to filter by | ✓ (for filtering) | `-pc "TableID"` |
| `--key-value-mapping` | JSON mapping file path | ✓ (for filtering) | `--key-value-mapping "map.json"` |
| `-o, --output` | Output directory/file | ✓ | `-o "results/"` |
| `--sheet` | Specific sheet name | - | `--sheet "Tables"` |

## Output Examples

### Step 1 Output (Sheet Conversion):
```markdown
# Sheet: Table Definitions

## Raw Data with Cell References

| Col 1 | Col 2 | Col 3 | Col 4 |
|-------|-------|-------|-------|
| Data Dictionary (A1) | | | |
| Version 2.1 (A2) | | | |
| | | | |
| Table ID (A6) | Table Name (B6) | Description (C6) | Status (D6) |
| ENG001 (A7) | Engineering (B7) | Engineering Department Data (C7) | Active (D7) |
| MKT001 (A8) | Marketing (B8) | Marketing Department Data (C8) | Active (D8) |

## Summary for AI Analysis
- Total rows with data: 5
- Total columns: 4
- Sheet name: Table Definitions
- Cell references are included in parentheses for precise identification
```

### Step 4 Output (Filtered Data):
```markdown
# Filtered Data: TableID = 'ENG001'

## Metadata
- Filter criteria: TableID = 'ENG001'
- Data source starting cell: A6
- Total matching rows: 15
- Columns included: Name, Department, Role, Email

## Data Table

| Name | Department | Role | Email |
|------|------------|------|-------|
| John Smith | Engineering | Senior Developer | john@company.com |
| Sarah Johnson | Engineering | Tech Lead | sarah@company.com |
| Mike Wilson | Engineering | DevOps Engineer | mike@company.com |
```

## JSON Mapping File Examples

### Simple Table ID Mapping:
```json
{
    "engineering_team": "ENG001",
    "marketing_team": "MKT001",
    "sales_team": "SLS001"
}
```

### User Profile Mapping:
```json
{
    "admin_users": "ADMIN",
    "standard_users": "USER", 
    "guest_users": "GUEST"
}
```

### Department Code Mapping:
```json
{
    "hr_department": "HR",
    "it_department": "IT",
    "finance_department": "FIN"
}
```

## AI Assistant Best Practices

### 1. **Always Start with Sheet Conversion**
```bash
python id_based_data_dictionary_processor.py data_dict.xlsx --convert-all
```
Never skip this step - it provides essential structure analysis.

### 2. **Thoroughly Analyze Generated Markdown**
Look for:
- ID patterns (e.g., "ENG001", "USR123")
- Table relationships
- Column header locations
- Data start positions

### 3. **Create Accurate JSON Mappings**
Ensure JSON keys are valid filenames and values match exactly what appears in the Excel data.

### 4. **Use Precise Cell References**
Always use the exact cell reference identified in Step 1 analysis (e.g., "A6", not "A5" or "A7").

### 5. **Validate Column Names**
Use exact column names as they appear in the Excel headers, including spaces and special characters.

## Error Handling

### Common Issues and Solutions:

**"Primary column not found"**
- Check column name spelling and case
- Verify the start cell is correct
- Ensure the sheet contains the expected data structure

**"No data found for filter criteria"**
- Verify the filter value exists in the primary column
- Check for leading/trailing spaces in the data
- Ensure case-sensitive matching if needed

**"JSON mapping file empty"**
- Verify JSON file syntax
- Ensure file path is correct
- Check that JSON contains at least one key-value pair

## Integration with Other Tools

This tool is designed to work alongside:
- `excel_advanced_processor.py` for general Excel processing
- Standard pandas/openpyxl workflows
- Data analysis pipelines requiring structured input

## Requirements

- Python 3.6+
- pandas>=1.3.0
- openpyxl>=3.0.0

## Advanced Usage Scenarios

### Multi-Sheet Processing
```bash
# Process specific sheet
python id_based_data_dictionary_processor.py data.xlsx -s "B3" -pc "ID" --key-value-mapping "map.json" -o "results/" --sheet "UserData"
```

### Column-Specific Extraction
```bash
# Extract only specific columns
python id_based_data_dictionary_processor.py data.xlsx -s "A1" -c "UserID" "Username" "Email" -pc "Status" --key-value-mapping "active_users.json" -o "active/"
```

### Full Workflow Example
```bash
# Step 1: Convert all sheets
python id_based_data_dictionary_processor.py employee_data.xlsx --convert-all -o "analysis/"

# AI analyzes markdown files in analysis/ directory

# Step 2: Create JSON mapping based on analysis
echo '{"engineering": "ENG", "marketing": "MKT"}' > dept_mapping.json

# Step 3: Extract filtered data
python id_based_data_dictionary_processor.py employee_data.xlsx -s "A4" -c "Name" "Role" "Email" -pc "DeptCode" --key-value-mapping "dept_mapping.json" -o "departments/"
```

This tool is specifically optimized for AI-assisted workflows where the AI handles analysis and decision-making while the Python script handles the data processing efficiently.
