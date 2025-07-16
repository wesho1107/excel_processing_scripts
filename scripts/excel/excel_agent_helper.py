#!/usr/bin/env python3
"""
AI Agent Helper Module for Advanced Excel Processor

This module provides a Python interface for AI agents to interact with
the Advanced Excel Processor, making it easier to call functions
programmatically while maintaining the CLI interface.
"""

import subprocess
import json
import os
import tempfile
from typing import List, Dict, Any, Optional, Union
from pathlib import Path


class ExcelProcessorAgent:
    """
    AI Agent interface for the Advanced Excel Processor.
    
    This class provides methods that AI agents can call to interact with
    Excel files using the Advanced Excel Processor CLI tool.
    """
    
    def __init__(self, processor_path: str = "scripts/excel/excel_advanced_processor.py"):
        """Initialize the Excel Processor Agent."""
        self.processor_path = processor_path
        self.python_cmd = "python"
    
    def _run_command(self, args: List[str]) -> Dict[str, Any]:
        """Run a command and return the result."""
        cmd = [self.python_cmd, self.processor_path] + args
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False
            )
            
            return {
                "success": result.returncode == 0,
                "stdout": result.stdout.strip(),
                "stderr": result.stderr.strip(),
                "return_code": result.returncode
            }
        except Exception as e:
            return {
                "success": False,
                "stdout": "",
                "stderr": str(e),
                "return_code": -1
            }
    
    def create_workbook(self, filepath: str) -> Dict[str, Any]:
        """
        Create a new Excel workbook.
        
        Args:
            filepath: Path where to create the workbook
            
        Returns:
            Dictionary with operation result
        """
        result = self._run_command(["--create-workbook", filepath])
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"],
            "filepath": filepath if result["success"] else None
        }
    
    def get_metadata(self, filepath: str, include_ranges: bool = False) -> Dict[str, Any]:
        """
        Get comprehensive metadata about an Excel workbook.
        
        Args:
            filepath: Path to Excel file
            include_ranges: Whether to include detailed range information
            
        Returns:
            Dictionary with workbook metadata
        """
        args = [filepath, "--metadata"]
        if include_ranges:
            args.append("--include-ranges")
        
        result = self._run_command(args)
        
        if result["success"]:
            try:
                metadata = json.loads(result["stdout"])
                return {
                    "success": True,
                    "metadata": metadata
                }
            except json.JSONDecodeError:
                return {
                    "success": False,
                    "message": "Failed to parse metadata JSON"
                }
        else:
            return {
                "success": False,
                "message": result["stderr"]
            }
    
    def read_data(self, filepath: str, sheet_name: str, range_str: str, 
                  preview_only: bool = False, output_file: Optional[str] = None) -> Dict[str, Any]:
        """
        Read data from an Excel range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            range_str: Cell range (e.g., 'A1:D10')
            preview_only: Whether to limit to first 10 rows
            output_file: Optional file to save JSON output
            
        Returns:
            Dictionary with data or file path
        """
        args = [filepath, "--read-data", sheet_name, range_str]
        
        if preview_only:
            args.append("--preview-only")
        
        if output_file:
            args.extend(["-o", output_file])
        
        result = self._run_command(args)
        
        if result["success"]:
            if output_file:
                return {
                    "success": True,
                    "output_file": output_file,
                    "message": result["stdout"]
                }
            else:
                try:
                    data = json.loads(result["stdout"])
                    return {
                        "success": True,
                        "data": data
                    }
                except json.JSONDecodeError:
                    return {
                        "success": False,
                        "message": "Failed to parse data JSON"
                    }
        else:
            return {
                "success": False,
                "message": result["stderr"]
            }
    
    def write_data(self, filepath: str, sheet_name: str, start_cell: str, 
                   data: List[List[Any]]) -> Dict[str, Any]:
        """
        Write data to an Excel range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            start_cell: Starting cell (e.g., 'A1')
            data: List of lists containing data to write
            
        Returns:
            Dictionary with operation result
        """
        # Create temporary JSON file for data
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f, indent=2, default=str)
            temp_file = f.name
        
        try:
            args = [filepath, "--write-data", sheet_name, start_cell, "--input-data", temp_file]
            result = self._run_command(args)
            
            return {
                "success": result["success"],
                "message": result["stdout"] if result["success"] else result["stderr"]
            }
        finally:
            # Clean up temporary file
            os.unlink(temp_file)
    
    def format_range(self, filepath: str, sheet_name: str, start_cell: str, end_cell: str,
                     bold: bool = False, italic: bool = False, font_size: Optional[int] = None,
                     font_color: Optional[str] = None, bg_color: Optional[str] = None,
                     border_style: Optional[str] = None) -> Dict[str, Any]:
        """
        Apply formatting to a cell range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            start_cell: Starting cell
            end_cell: Ending cell
            bold: Apply bold formatting
            italic: Apply italic formatting
            font_size: Font size in points
            font_color: Font color in hex (e.g., 'FF0000')
            bg_color: Background color in hex (e.g., 'FFFF00')
            border_style: Border style (thin, medium, thick)
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--format-range", sheet_name, start_cell, end_cell]
        
        if bold:
            args.append("--bold")
        if italic:
            args.append("--italic")
        if font_size:
            args.extend(["--font-size", str(font_size)])
        if font_color:
            args.extend(["--font-color", font_color])
        if bg_color:
            args.extend(["--bg-color", bg_color])
        if border_style:
            args.extend(["--border-style", border_style])
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def create_chart(self, filepath: str, sheet_name: str, data_range: str, 
                     chart_type: str, target_cell: str, title: str = "",
                     x_axis: str = "", y_axis: str = "") -> Dict[str, Any]:
        """
        Create a chart from data range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            data_range: Data range for chart (e.g., 'A1:D10')
            chart_type: Chart type (line, bar, pie, scatter, area)
            target_cell: Cell where to place chart
            title: Chart title
            x_axis: X-axis label
            y_axis: Y-axis label
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--create-chart", sheet_name, data_range, chart_type, target_cell]
        
        if title:
            args.extend(["--title", title])
        if x_axis:
            args.extend(["--x-axis", x_axis])
        if y_axis:
            args.extend(["--y-axis", y_axis])
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def create_pivot_table(self, filepath: str, sheet_name: str, data_range: str,
                          rows: List[str], values: List[str], 
                          columns: Optional[List[str]] = None,
                          agg_func: str = "sum") -> Dict[str, Any]:
        """
        Create a pivot table from data.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            data_range: Data range for pivot (e.g., 'A1:D100')
            rows: Column names for pivot rows
            values: Column names for pivot values
            columns: Column names for pivot columns (optional)
            agg_func: Aggregation function (sum, mean, count, max, min)
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--create-pivot", sheet_name, data_range, 
                "--rows"] + rows + ["--values"] + values + ["--agg-func", agg_func]
        
        if columns:
            args.extend(["--columns"] + columns)
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def create_table(self, filepath: str, sheet_name: str, data_range: str,
                     table_name: Optional[str] = None, 
                     table_style: str = "TableStyleMedium9") -> Dict[str, Any]:
        """
        Create a native Excel table.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            data_range: Data range for table (e.g., 'A1:D10')
            table_name: Name for the table (optional)
            table_style: Table style
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--create-table", sheet_name, data_range, "--table-style", table_style]
        
        if table_name:
            args.extend(["--table-name", table_name])
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def apply_formula(self, filepath: str, sheet_name: str, cell: str, formula: str) -> Dict[str, Any]:
        """
        Apply an Excel formula to a cell.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of the worksheet
            cell: Target cell (e.g., 'E1')
            formula: Excel formula (e.g., '=SUM(A1:D1)')
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--apply-formula", sheet_name, cell, formula]
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def copy_worksheet(self, filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
        """
        Copy a worksheet within the workbook.
        
        Args:
            filepath: Path to Excel file
            source_sheet: Name of source worksheet
            target_sheet: Name for copied worksheet
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--copy-sheet", source_sheet, target_sheet]
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }
    
    def delete_worksheet(self, filepath: str, sheet_name: str) -> Dict[str, Any]:
        """
        Delete a worksheet from the workbook.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet to delete
            
        Returns:
            Dictionary with operation result
        """
        args = [filepath, "--delete-sheet", sheet_name]
        
        result = self._run_command(args)
        
        return {
            "success": result["success"],
            "message": result["stdout"] if result["success"] else result["stderr"]
        }


# Convenience functions for AI agents
def quick_analysis(filepath: str, sheet_name: str = None) -> Dict[str, Any]:
    """
    Quick analysis of an Excel file for AI agents.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Optional specific sheet name
        
    Returns:
        Dictionary with analysis results
    """
    agent = ExcelProcessorAgent()
    
    # Get metadata
    metadata_result = agent.get_metadata(filepath, include_ranges=True)
    if not metadata_result["success"]:
        return metadata_result
    
    metadata = metadata_result["metadata"]
    
    # If no sheet specified, use first sheet
    if not sheet_name:
        sheet_name = metadata["sheets"][0]["name"]
    
    # Get preview of data
    sheet_info = next((s for s in metadata["sheets"] if s["name"] == sheet_name), None)
    if not sheet_info:
        return {"success": False, "message": f"Sheet '{sheet_name}' not found"}
    
    data_range = sheet_info["data_range"]
    preview_result = agent.read_data(filepath, sheet_name, data_range, preview_only=True)
    
    return {
        "success": True,
        "metadata": metadata,
        "sheet_name": sheet_name,
        "data_preview": preview_result.get("data", []),
        "analysis": {
            "total_sheets": len(metadata["sheets"]),
            "sheet_names": [s["name"] for s in metadata["sheets"]],
            "max_row": sheet_info["max_row"],
            "max_column": sheet_info["max_column"],
            "has_data": sheet_info.get("has_data", False)
        }
    }


def create_report(data: List[List[Any]], filepath: str, sheet_name: str = "Report",
                  format_headers: bool = True, create_chart: bool = True,
                  chart_type: str = "bar") -> Dict[str, Any]:
    """
    Create a formatted Excel report from data.
    
    Args:
        data: List of lists containing data (first row should be headers)
        filepath: Output file path
        sheet_name: Name for the worksheet
        format_headers: Whether to format header row
        create_chart: Whether to create a chart
        chart_type: Type of chart to create
        
    Returns:
        Dictionary with operation result
    """
    agent = ExcelProcessorAgent()
    
    # Create workbook
    create_result = agent.create_workbook(filepath)
    if not create_result["success"]:
        return create_result
    
    # Write data
    write_result = agent.write_data(filepath, sheet_name, "A1", data)
    if not write_result["success"]:
        return write_result
    
    results = [create_result, write_result]
    
    # Format headers if requested
    if format_headers and len(data) > 0:
        num_cols = len(data[0])
        end_col = chr(ord('A') + num_cols - 1)
        format_result = agent.format_range(
            filepath, sheet_name, "A1", f"{end_col}1",
            bold=True, bg_color="366092", font_color="FFFFFF"
        )
        results.append(format_result)
    
    # Create chart if requested
    if create_chart and len(data) > 1:
        num_rows = len(data)
        num_cols = len(data[0])
        end_col = chr(ord('A') + num_cols - 1)
        data_range = f"A1:{end_col}{num_rows}"
        
        chart_result = agent.create_chart(
            filepath, sheet_name, data_range, chart_type, "G1",
            title="Data Visualization"
        )
        results.append(chart_result)
    
    # Create table
    if len(data) > 0:
        num_rows = len(data)
        num_cols = len(data[0])
        end_col = chr(ord('A') + num_cols - 1)
        table_range = f"A1:{end_col}{num_rows}"
        
        table_result = agent.create_table(
            filepath, sheet_name, table_range, "ReportData"
        )
        results.append(table_result)
    
    return {
        "success": all(r["success"] for r in results),
        "results": results,
        "filepath": filepath
    }


# Example usage for AI agents
if __name__ == "__main__":
    # Demonstrate how AI agents can use this module
    agent = ExcelProcessorAgent()
    
    # Example 1: Quick analysis
    print("=== Quick Analysis Example ===")
    analysis = quick_analysis("sample_data.xlsx")
    print(json.dumps(analysis, indent=2, default=str))
    
    # Example 2: Create report
    print("\n=== Create Report Example ===")
    sample_data = [
        ["Name", "Department", "Salary"],
        ["John", "Engineering", 75000],
        ["Jane", "Marketing", 65000],
        ["Bob", "Sales", 70000]
    ]
    
    report_result = create_report(sample_data, "ai_report.xlsx")
    print(json.dumps(report_result, indent=2, default=str))
