# Excel Processing Commands for AI Agents

## Tool Overview
Use the Advanced Excel Processor located at `scripts/excel/excel_advanced_processor.py` for comprehensive Excel operations.

## Command Templates

### 1. Explore Excel File
```bash
# Get file structure and metadata
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --metadata --include-ranges

# Preview data from a sheet
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --read-data "Sheet1" "A1:Z100" --preview-only
```

### 2. Data Operations
```bash
# Read data to JSON
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --read-data "Sheet1" "A1:D10" -o "data.json"

# Write data from JSON
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --write-data "Sheet1" "A1" --input-data "data.json"
```

### 3. Analysis Operations
```bash
# Create pivot table for analysis
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --create-pivot "Sheet1" "A1:D100" --rows "Department" --values "Salary" --agg-func "sum"

# Create chart for visualization
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --create-chart "Sheet1" "A1:D10" "bar" "F1" --title "Analysis Chart"
```

### 4. Formatting Operations
```bash
# Format headers
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --format-range "Sheet1" "A1" "D1" --bold --bg-color "366092" --font-color "FFFFFF"

# Create professional table
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --create-table "Sheet1" "A1:D10" --table-name "DataTable" --table-style "TableStyleMedium9"
```

### 5. Worksheet Management
```bash
# Create backup
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --copy-sheet "Sheet1" "Sheet1_Backup"

# Apply calculations
python scripts/excel/excel_advanced_processor.py "filename.xlsx" --apply-formula "Sheet1" "E1" "=SUM(A1:D1)"
```

## Common Workflows

### Data Analysis Workflow
1. **Explore**: Get metadata and preview data
2. **Analyze**: Create pivot tables for insights
3. **Visualize**: Generate charts from analysis
4. **Format**: Apply professional formatting

### Report Generation Workflow
1. **Create**: New workbook for report
2. **Populate**: Write data from external sources
3. **Format**: Apply headers and styling
4. **Enhance**: Add charts and tables

### Data Processing Workflow
1. **Backup**: Copy sheets before modifications
2. **Extract**: Read data for processing
3. **Transform**: Process data externally
4. **Update**: Write processed data back

## Error Handling
- Always check if file exists before operations
- Use metadata to verify sheet names and ranges
- Handle permission errors (close file in Excel)
- Validate data formats before writing

## Best Practices
- Use meaningful names for charts and tables
- Create backups before major operations
- Use preview mode for large datasets
- Apply formatting consistently
- Document operations in comments

## Integration with Python
```python
# Use the helper module for programmatic access
from scripts.excel.excel_agent_helper import ExcelProcessorAgent

agent = ExcelProcessorAgent()
result = agent.get_metadata("filename.xlsx", include_ranges=True)
```

This tool provides enterprise-level Excel manipulation with AI-friendly command-line interface.
