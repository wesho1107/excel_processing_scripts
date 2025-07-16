# Advanced Excel Processor - AI Agent Guide

This guide provides comprehensive instructions for AI agents on how to use the Advanced Excel Processor CLI tool.

## Quick Reference

The Advanced Excel Processor is a command-line tool located at `scripts/excel/excel_advanced_processor.py` that provides comprehensive Excel manipulation capabilities.

### Basic Command Structure
```bash
python scripts/excel/excel_advanced_processor.py [input_file] [operation] [parameters]
```

## Core Operations

### 1. Workbook Operations

#### Create New Workbook
```bash
python scripts/excel/excel_advanced_processor.py --create-workbook "new_file.xlsx"
```

#### Get Workbook Metadata
```bash
python scripts/excel/excel_advanced_processor.py "input.xlsx" --metadata
python scripts/excel/excel_advanced_processor.py "input.xlsx" --metadata --include-ranges
```

### 2. Data Operations

#### Read Data from Range
```bash
# Read specific range
python scripts/excel/excel_advanced_processor.py "input.xlsx" --read-data "Sheet1" "A1:D10"

# Read with preview (first 10 rows only)
python scripts/excel/excel_advanced_processor.py "input.xlsx" --read-data "Sheet1" "A1:D100" --preview-only

# Save output to file
python scripts/excel/excel_advanced_processor.py "input.xlsx" --read-data "Sheet1" "A1:D10" -o "output.json"
```

#### Write Data to Range
```bash
# Requires JSON input file
python scripts/excel/excel_advanced_processor.py "input.xlsx" --write-data "Sheet1" "A1" --input-data "data.json"
```

**JSON Format for Input Data:**
```json
[
  ["Name", "Age", "Department"],
  ["John Doe", 30, "Engineering"],
  ["Jane Smith", 25, "Marketing"]
]
```

### 3. Formatting Operations

#### Format Cell Ranges
```bash
# Basic formatting
python scripts/excel/excel_advanced_processor.py "input.xlsx" --format-range "Sheet1" "A1" "D1" --bold

# Advanced formatting
python scripts/excel/excel_advanced_processor.py "input.xlsx" --format-range "Sheet1" "A1" "D1" --bold --italic --font-size 14 --font-color "FF0000" --bg-color "FFFF00" --border-style "thin"
```

**Available Formatting Options:**
- `--bold`: Apply bold text
- `--italic`: Apply italic text
- `--font-size N`: Set font size (number)
- `--font-color RRGGBB`: Set font color (hex)
- `--bg-color RRGGBB`: Set background color (hex)
- `--border-style STYLE`: Set border style (thin, medium, thick)

### 4. Chart Operations

#### Create Charts
```bash
# Basic chart
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-chart "Sheet1" "A1:D10" "line" "F1"

# Chart with labels
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-chart "Sheet1" "A1:D10" "bar" "F1" --title "Sales Data" --x-axis "Months" --y-axis "Revenue"
```

**Supported Chart Types:**
- `line`: Line chart
- `bar`: Bar chart
- `pie`: Pie chart
- `scatter`: Scatter plot
- `area`: Area chart

### 5. Pivot Table Operations

#### Create Pivot Tables
```bash
# Basic pivot table
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-pivot "Sheet1" "A1:D100" --rows "Department" --values "Salary"

# Advanced pivot table
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-pivot "Sheet1" "A1:D100" --rows "Department" "Location" --values "Salary" "Bonus" --columns "Year" --agg-func "mean"
```

**Aggregation Functions:**
- `sum`: Sum values
- `mean`: Average values
- `count`: Count values
- `max`: Maximum value
- `min`: Minimum value

### 6. Table Operations

#### Create Excel Tables
```bash
# Basic table
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-table "Sheet1" "A1:D10"

# Named table with style
python scripts/excel/excel_advanced_processor.py "input.xlsx" --create-table "Sheet1" "A1:D10" --table-name "EmployeeData" --table-style "TableStyleMedium9"
```

### 7. Formula Operations

#### Apply Formulas
```bash
python scripts/excel/excel_advanced_processor.py "input.xlsx" --apply-formula "Sheet1" "E1" "=SUM(A1:D1)"
python scripts/excel/excel_advanced_processor.py "input.xlsx" --apply-formula "Sheet1" "E2" "=AVERAGE(A1:A10)"
```

### 8. Worksheet Operations

#### Copy Worksheets
```bash
python scripts/excel/excel_advanced_processor.py "input.xlsx" --copy-sheet "Sheet1" "Sheet1_Backup"
```

#### Delete Worksheets
```bash
python scripts/excel/excel_advanced_processor.py "input.xlsx" --delete-sheet "UnwantedSheet"
```

## Common AI Agent Workflows

### Workflow 1: Data Analysis Report
```bash
# 1. Read data
python scripts/excel/excel_advanced_processor.py "data.xlsx" --read-data "Sheet1" "A1:Z1000" -o "data.json"

# 2. Create pivot table for analysis
python scripts/excel/excel_advanced_processor.py "data.xlsx" --create-pivot "Sheet1" "A1:Z1000" --rows "Department" --values "Sales" --agg-func "sum"

# 3. Create chart from pivot
python scripts/excel/excel_advanced_processor.py "data.xlsx" --create-chart "Sheet1_pivot" "A1:B10" "bar" "D1" --title "Sales by Department"

# 4. Format headers
python scripts/excel/excel_advanced_processor.py "data.xlsx" --format-range "Sheet1" "A1" "Z1" --bold --bg-color "366092" --font-color "FFFFFF"
```

### Workflow 2: Report Generation
```bash
# 1. Create new workbook
python scripts/excel/excel_advanced_processor.py --create-workbook "report.xlsx"

# 2. Write data from JSON
python scripts/excel/excel_advanced_processor.py "report.xlsx" --write-data "Sheet" "A1" --input-data "report_data.json"

# 3. Format as table
python scripts/excel/excel_advanced_processor.py "report.xlsx" --create-table "Sheet" "A1:E100" --table-name "ReportData"

# 4. Create summary chart
python scripts/excel/excel_advanced_processor.py "report.xlsx" --create-chart "Sheet" "A1:E100" "line" "G1" --title "Trend Analysis"
```

### Workflow 3: Data Exploration
```bash
# 1. Get metadata first
python scripts/excel/excel_advanced_processor.py "unknown.xlsx" --metadata --include-ranges

# 2. Preview data
python scripts/excel/excel_advanced_processor.py "unknown.xlsx" --read-data "Sheet1" "A1:Z100" --preview-only

# 3. Create exploratory pivot
python scripts/excel/excel_advanced_processor.py "unknown.xlsx" --create-pivot "Sheet1" "A1:Z100" --rows "Category" --values "Value"
```

## Error Handling

The tool provides clear error messages for common issues:
- File not found
- Invalid sheet names
- Invalid cell ranges
- Missing required parameters
- Formatting errors

## Output Formats

- **JSON**: Data operations return JSON format
- **Console**: Status messages and errors
- **Excel Files**: Direct modifications to Excel files

## Best Practices for AI Agents

1. **Always check file existence** before operations
2. **Use metadata** to understand file structure
3. **Preview data** before full operations
4. **Create backups** using copy-sheet before major changes
5. **Use descriptive names** for charts and tables
6. **Handle errors gracefully** by checking return status

## Integration with Other Tools

The Advanced Excel Processor can be easily integrated with:
- Data processing pipelines
- Report generation systems
- Analysis workflows
- Automation scripts

## Performance Notes

- Large datasets: Use `--preview-only` for initial exploration
- Memory usage: Consider processing in chunks for very large files
- File locking: Ensure Excel files are not open in other applications

This tool provides enterprise-level Excel manipulation capabilities optimized for AI agent usage with clear, predictable command-line interfaces and comprehensive error handling.
