#!/usr/bin/env python3
"""
Excel to Markdown Converter

A command-line tool to extract data tables from Excel files and convert them to Markdown format.
Handles Excel files where data doesn't start at the first cell and extracts only specified columns.
"""

import argparse
import sys
import os
from pathlib import Path
import pandas as pd
import openpyxl
from typing import List, Optional, Tuple, Dict, Any


class ExcelProcessor:
    """Process Excel files and convert data tables to Markdown format."""
    
    def __init__(self, file_path: str):
        """Initialize the Excel processor with a file path."""
        self.file_path = file_path
        self.workbook = None
        self.worksheet = None
    
    def load_excel(self, sheet_name: Optional[str] = None) -> bool:
        """Load Excel file and select worksheet."""
        try:
            self.workbook = openpyxl.load_workbook(self.file_path)
            if sheet_name:
                if sheet_name in self.workbook.sheetnames:
                    self.worksheet = self.workbook[sheet_name]
                else:
                    print(f"Warning: Sheet '{sheet_name}' not found. Using first sheet.")
                    self.worksheet = self.workbook.active
            else:
                self.worksheet = self.workbook.active
            return True
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def find_data_table_start(self) -> Optional[Tuple[int, int]]:
        """
        Find the starting position of the data table by looking for the first row
        that contains multiple non-empty cells (likely column headers).
        """
        max_row = self.worksheet.max_row
        max_col = self.worksheet.max_column
        
        for row in range(1, max_row + 1):
            non_empty_cells = 0
            consecutive_cells = 0
            max_consecutive = 0
            
            for col in range(1, max_col + 1):
                cell_value = self.worksheet.cell(row=row, column=col).value
                if cell_value is not None and str(cell_value).strip():
                    non_empty_cells += 1
                    consecutive_cells += 1
                    max_consecutive = max(max_consecutive, consecutive_cells)
                else:
                    consecutive_cells = 0
            
            # Consider it a header row if it has at least 2 non-empty cells
            # and at least 2 consecutive non-empty cells
            if non_empty_cells >= 2 and max_consecutive >= 2:
                return (row, 1)
        
        return None
    
    def extract_column_headers(self, start_row: int, start_col: int) -> List[str]:
        """Extract column headers from the identified header row."""
        headers = []
        max_col = self.worksheet.max_column
        
        for col in range(start_col, max_col + 1):
            cell_value = self.worksheet.cell(row=start_row, column=col).value
            if cell_value is not None:
                headers.append(str(cell_value).strip())
            else:
                # Stop when we hit an empty cell after finding some headers
                if headers:
                    break
                headers.append("")
        
        return headers
    
    def extract_data_table(self, required_columns: List[str]) -> Optional[pd.DataFrame]:
        """
        Extract the data table from Excel file, filtering for required columns only.
        """
        # Find the start of the data table
        table_start = self.find_data_table_start()
        if not table_start:
            print("Error: Could not find data table in the Excel file.")
            return None
        
        start_row, start_col = table_start
        print(f"Found data table starting at row {start_row}, column {start_col}")
        
        # Extract column headers
        headers = self.extract_column_headers(start_row, start_col)
        print(f"Found headers: {headers}")
        
        # Find indices of required columns
        required_indices = []
        missing_columns = []
        
        for req_col in required_columns:
            try:
                # Case-insensitive matching
                index = next(i for i, h in enumerate(headers) if h.lower() == req_col.lower())
                required_indices.append(index)
            except StopIteration:
                missing_columns.append(req_col)
        
        if missing_columns:
            print(f"Warning: The following required columns were not found: {missing_columns}")
            print(f"Available columns: {[h for h in headers if h]}")
        
        if not required_indices:
            print("Error: None of the required columns were found in the Excel file.")
            return None
        
        # Extract data rows
        data_rows = []
        max_row = self.worksheet.max_row
        
        for row in range(start_row + 1, max_row + 1):
            row_data = []
            has_data = False
            
            for col_index in required_indices:
                cell_value = self.worksheet.cell(row=row, column=start_col + col_index).value
                if cell_value is not None:
                    row_data.append(str(cell_value).strip())
                    has_data = True
                else:
                    row_data.append("")
            
            # Only add row if it has at least one non-empty cell
            if has_data:
                data_rows.append(row_data)
            else:
                # Stop when we hit an empty row (end of data)
                if data_rows:
                    break
        
        # Create DataFrame with filtered columns
        filtered_headers = [headers[i] for i in required_indices]
        df = pd.DataFrame(data_rows, columns=filtered_headers)
        
        return df
    
    def dataframe_to_markdown(self, df: pd.DataFrame) -> str:
        """Convert DataFrame to Markdown table format."""
        if df.empty:
            return "No data found."
        
        # Create markdown table
        markdown_lines = []
        
        # Header row
        header_row = "| " + " | ".join(df.columns) + " |"
        markdown_lines.append(header_row)
        
        # Separator row
        separator_row = "| " + " | ".join(["---"] * len(df.columns)) + " |"
        markdown_lines.append(separator_row)
        
        # Data rows
        for _, row in df.iterrows():
            data_row = "| " + " | ".join([str(val) if val else "" for val in row]) + " |"
            markdown_lines.append(data_row)
        
        return "\n".join(markdown_lines)
    
    def process_excel_to_markdown(self, required_columns: List[str], 
                                 output_path: Optional[str] = None,
                                 sheet_name: Optional[str] = None) -> str:
        """
        Main processing function to convert Excel to Markdown.
        """
        if not self.load_excel(sheet_name):
            return ""
        
        df = self.extract_data_table(required_columns)
        if df is None:
            return ""
        
        print(f"Extracted {len(df)} rows and {len(df.columns)} columns")
        
        markdown_content = self.dataframe_to_markdown(df)
        
        if output_path:
            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                print(f"Markdown file saved to: {output_path}")
            except Exception as e:
                print(f"Error saving file: {e}")
        
        return markdown_content


def main():
    """Main function for command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert Excel data tables to Markdown format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_processor.py input.xlsx -c "Name" "Job" -o output.md
  python excel_processor.py data.xlsx -c "Name" "Age" "Job" -s "Sheet1"
  python excel_processor.py file.xlsx -c "Name" "Job" --print-only
        """
    )
    
    parser.add_argument("input_file", help="Path to the Excel file")
    parser.add_argument("-c", "--columns", nargs="+", required=True,
                       help="List of column names to extract (space-separated)")
    parser.add_argument("-o", "--output", help="Output Markdown file path")
    parser.add_argument("-s", "--sheet", help="Sheet name to process (default: first sheet)")
    parser.add_argument("--print-only", action="store_true",
                       help="Print markdown to stdout instead of saving to file")
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' does not exist.")
        sys.exit(1)
    
    # Determine output path
    output_path = None
    if not args.print_only:
        if args.output:
            output_path = args.output
        else:
            # Generate default output filename
            input_path = Path(args.input_file)
            output_path = input_path.with_suffix('.md')
    
    # Process the Excel file
    processor = ExcelProcessor(args.input_file)
    markdown_content = processor.process_excel_to_markdown(
        args.columns, output_path, args.sheet
    )
    
    if args.print_only and markdown_content:
        print("\n" + "="*50)
        print("MARKDOWN OUTPUT:")
        print("="*50)
        print(markdown_content)
    
    if not markdown_content:
        print("Failed to process the Excel file.")
        sys.exit(1)


if __name__ == "__main__":
    main()