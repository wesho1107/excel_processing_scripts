#!/usr/bin/env python3
"""
Example usage of the Excel processor script.
This demonstrates how to use the ExcelProcessor class programmatically.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts', 'excel'))

from excel_processor import ExcelProcessor

def example_usage():
    """Example of how to use the ExcelProcessor class."""
    
    # Example 1: Process an Excel file programmatically
    print("Example 1: Processing Excel file programmatically")
    print("=" * 50)
    
    # Replace with your actual Excel file path
    excel_file = "sample_data.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"Note: Example file '{excel_file}' not found.")
        print("To test this script, create an Excel file with some data.")
        return
    
    # Initialize the processor
    processor = ExcelProcessor(excel_file)
    
    # Define the columns you want to extract
    required_columns = ["Name", "Job"]
    
    # Process and convert to markdown
    markdown_content = processor.process_excel_to_markdown(
        required_columns=required_columns,
        output_path="output_example.md"
    )
    
    if markdown_content:
        print("Successfully processed Excel file!")
        print("Markdown content:")
        print("-" * 30)
        print(markdown_content)
    else:
        print("Failed to process Excel file.")

if __name__ == "__main__":
    example_usage()
