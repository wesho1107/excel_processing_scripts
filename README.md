# Excel to Markdown Converter

A command-line tool to extract data tables from Excel files and convert them to Markdown format. This tool is designed to handle Excel files where data doesn't start at the first cell and can extract only specified columns of interest.

## Features

- Automatically finds the data table within Excel files, even if there are title cells before the actual data
- Extracts only the specified columns from the data table
- Converts the filtered data to clean Markdown table format
- Supports different Excel sheets
- Case-insensitive column name matching
- Handles empty cells gracefully

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python scripts/excel/excel_processor.py input.xlsx -c "Name" "Job" -o output.md
```

### Command-line Options

- `input_file`: Path to the Excel file (required)
- `-c, --columns`: List of column names to extract (required, space-separated)
- `-o, --output`: Output Markdown file path (optional, defaults to input filename with .md extension)
- `-s, --sheet`: Sheet name to process (optional, defaults to first sheet)
- `--print-only`: Print markdown to stdout instead of saving to file

### Examples

1. Extract "Name" and "Job" columns and save to a file:
```bash
python scripts/excel/excel_processor.py data.xlsx -c "Name" "Job" -o output.md
```

2. Extract multiple columns from a specific sheet:
```bash
python scripts/excel/excel_processor.py data.xlsx -c "Name" "Age" "Job" -s "Sheet1"
```

3. Print output to console instead of saving:
```bash
python scripts/excel/excel_processor.py data.xlsx -c "Name" "Job" --print-only
```

4. Use default output filename (input filename with .md extension):
```bash
python scripts/excel/excel_processor.py data.xlsx -c "Name" "Job"
```

## How It Works

1. **Data Table Detection**: The script automatically scans the Excel file to find the first row that contains multiple non-empty cells, which is assumed to be the column headers.

2. **Column Extraction**: It identifies the specified columns from the headers using case-insensitive matching.

3. **Data Processing**: Extracts all data rows under the identified headers until it reaches an empty row (indicating the end of the data table).

4. **Markdown Generation**: Converts the filtered data into a properly formatted Markdown table.

## Example Output

Given an Excel file with columns "Name", "Age", "Job", "Birthday" and requesting only "Name" and "Job":

```markdown
| Name | Job |
| --- | --- |
| John Doe | Software Engineer |
| Jane Smith | Data Analyst |
| Bob Johnson | Product Manager |
```

## Error Handling

The script includes comprehensive error handling for:
- Missing or invalid Excel files
- Sheet names that don't exist
- Missing required columns
- Empty data tables
- File I/O errors

## Requirements

- Python 3.6+
- pandas>=1.3.0
- openpyxl>=3.0.0
