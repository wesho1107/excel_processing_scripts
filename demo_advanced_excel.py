#!/usr/bin/env python3
"""
Demo script for the Advanced Excel Processor

This script demonstrates the comprehensive Excel operations available
in the advanced processor, converted from the Excel MCP Server.
"""

import os
import json
import tempfile
from scripts.excel.excel_advanced_processor import AdvancedExcelProcessor


def demo_advanced_excel_operations():
    """Demonstrate advanced Excel operations."""
    
    print("=== Advanced Excel Processor Demo ===")
    print("Based on Excel MCP Server functionality\n")
    
    # Create a temporary Excel file for demonstration
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        demo_file = tmp_file.name
    
    try:
        # 1. Create a new workbook
        print("1. Creating new workbook...")
        processor = AdvancedExcelProcessor(demo_file)
        processor.create_workbook(demo_file)
        
        # 2. Get workbook metadata
        print("\n2. Getting workbook metadata...")
        metadata = processor.get_workbook_metadata(include_ranges=True)
        print(json.dumps(metadata, indent=2))
        
        # 3. Write sample data
        print("\n3. Writing sample data...")
        sample_data = [
            ["Name", "Department", "Salary", "Years"],
            ["John Doe", "Engineering", 75000, 3],
            ["Jane Smith", "Marketing", 65000, 2],
            ["Bob Johnson", "Engineering", 80000, 5],
            ["Alice Brown", "Sales", 70000, 4],
            ["Charlie Davis", "Engineering", 85000, 6],
            ["Diana Wilson", "Marketing", 60000, 1]
        ]
        
        processor.write_data_to_range("Sheet", sample_data, "A1")
        
        # 4. Read data back
        print("\n4. Reading data back...")
        read_data = processor.read_data_range("Sheet", "A1", "D7")
        print("Data read from Excel:")
        for row in read_data:
            print(row)
        
        # 5. Format header row
        print("\n5. Formatting header row...")
        processor.format_range("Sheet", "A1", "D1", 
                             bold=True, bg_color="366092", font_color="FFFFFF")
        
        # 6. Create a chart
        print("\n6. Creating a chart...")
        processor.create_chart("Sheet", "A1:D7", "bar", "F1", 
                             title="Employee Salary Chart", 
                             x_axis="Employees", y_axis="Salary")
        
        # 7. Create a pivot table
        print("\n7. Creating a pivot table...")
        processor.create_pivot_table("Sheet", "A1:D7", 
                                   rows=["Department"], 
                                   values=["Salary"], 
                                   agg_func="sum")
        
        # 8. Create an Excel table
        print("\n8. Creating an Excel table...")
        processor.create_table("Sheet", "A1:D7", 
                             table_name="EmployeeTable",
                             table_style="TableStyleMedium9")
        
        # 9. Apply formulas
        print("\n9. Applying formulas...")
        processor.apply_formula("Sheet", "E1", "=AVERAGE(C2:C7)")
        processor.apply_formula("Sheet", "E2", "=MAX(C2:C7)")
        processor.apply_formula("Sheet", "E3", "=MIN(C2:C7)")
        
        # 10. Copy worksheet
        print("\n10. Copying worksheet...")
        processor.copy_worksheet("Sheet", "Sheet_Copy")
        
        # 11. Final metadata
        print("\n11. Final workbook metadata:")
        final_metadata = processor.get_workbook_metadata(include_ranges=True)
        print(json.dumps(final_metadata, indent=2))
        
        print(f"\n=== Demo completed successfully! ===")
        print(f"Demo file saved as: {demo_file}")
        print(f"You can open this file in Excel to see the results.")
        
        # Show command-line examples
        print("\n=== Command-line Examples ===")
        print("You can also use the advanced processor from the command line:")
        print(f"python scripts/excel/excel_advanced_processor.py {demo_file} --metadata")
        print(f"python scripts/excel/excel_advanced_processor.py {demo_file} --read-data Sheet A1:D7")
        print(f"python scripts/excel/excel_advanced_processor.py --create-workbook new_workbook.xlsx")
        print(f"python scripts/excel/excel_advanced_processor.py {demo_file} --create-chart Sheet A1:D7 line F10 --title 'New Chart'")
        
    except Exception as e:
        print(f"Error during demo: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Clean up (comment out if you want to keep the file)
        # os.unlink(demo_file)
        pass


def demonstrate_cli_usage():
    """Demonstrate command-line usage patterns."""
    
    print("\n=== CLI Usage Examples ===")
    
    examples = [
        {
            "description": "Create a new workbook",
            "command": "python scripts/excel/excel_advanced_processor.py --create-workbook new_file.xlsx"
        },
        {
            "description": "Get workbook metadata",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --metadata --include-ranges"
        },
        {
            "description": "Read data from a range",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --read-data Sheet1 A1:D10 --preview-only"
        },
        {
            "description": "Format a range with bold and background color",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --format-range Sheet1 A1 D1 --bold --bg-color FF0000"
        },
        {
            "description": "Create a line chart",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --create-chart Sheet1 A1:D10 line E1 --title 'Sales Chart' --x-axis 'Month' --y-axis 'Sales'"
        },
        {
            "description": "Create a pivot table",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --create-pivot Sheet1 A1:D100 --rows Department --values Salary --agg-func sum"
        },
        {
            "description": "Create an Excel table",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --create-table Sheet1 A1:D10 --table-name MyTable --table-style TableStyleLight1"
        },
        {
            "description": "Apply a formula to a cell",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --apply-formula Sheet1 E1 '=SUM(A1:D1)'"
        },
        {
            "description": "Copy a worksheet",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --copy-sheet Sheet1 Sheet1_Backup"
        },
        {
            "description": "Delete a worksheet",
            "command": "python scripts/excel/excel_advanced_processor.py input.xlsx --delete-sheet UnwantedSheet"
        }
    ]
    
    for i, example in enumerate(examples, 1):
        print(f"{i}. {example['description']}:")
        print(f"   {example['command']}")
        print()


if __name__ == "__main__":
    demo_advanced_excel_operations()
    demonstrate_cli_usage()
