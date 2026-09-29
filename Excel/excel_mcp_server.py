"""
Excel MCP Server
----------------
Model Context Protocol (MCP) server providing high-level tools to interact with Microsoft Excel.
Supports both live desktop Excel automation (via COM/pywin32) and headless file manipulation (via openpyxl).
"""

import os
import json
from typing import List, Optional, Dict, Any

try:
    from mcp.server.fastmcp import FastMCP
except Exception:
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except Exception:
        from mcp.server import MCPServer as FastMCP

from excel_engine import ExcelEngine

# Initialize MCP Server
mcp = FastMCP("excel-mcp-server")


@mcp.tool()
def excel_status() -> str:
    """
    Check the current status of Microsoft Excel and active workbooks.
    Returns whether Excel is running, COM availability, version, open workbooks, and active sheet.
    """
    status = ExcelEngine.get_status()
    if status.get("excel_running"):
        wb_names = [wb["name"] for wb in status.get("open_workbooks", [])]
        return (
            f"✅ Microsoft Excel (v{status.get('version')}) is RUNNING.\n"
            f"- Open Workbooks ({len(wb_names)}): {', '.join(wb_names) if wb_names else 'None'}\n"
            f"- Active Workbook: {status.get('active_workbook', 'None')}\n"
            f"- Active Sheet: {status.get('active_sheet', 'None')}"
        )
    else:
        com_status = "Available (Excel will be launched if needed)." if status.get("com_available") else "Unavailable (pywin32 not installed or non-Windows OS)."
        return (
            "ℹ️ Microsoft Excel is currently NOT running.\n"
            f"- COM automation: {com_status}\n"
            "- Headless file mode: Active (openpyxl can read/write files directly)."
        )


@mcp.tool()
def excel_create_workbook(
    filepath: Optional[str] = None,
    sheet_name: str = "Sheet1",
    headers: Optional[List[str]] = None,
    data: Optional[List[List[Any]]] = None,
    visible: bool = True,
    table_style: Optional[str] = None
) -> str:
    """
    Create a new Excel workbook with optional initial data, headers, sheet name, and table style.
    
    Args:
        filepath: Optional path where the workbook will be saved (e.g., 'reports/sales.xlsx'). If omitted, creates an unsaved workbook in live Excel.
        sheet_name: Name of the initial sheet (default: 'Sheet1').
        headers: List of column header names (e.g. ['Product', 'Price', 'Quantity', 'Total']).
        data: 2D list of row values/formulas (e.g. [['Apple', 1.5, 10, '=B2*C2'], ['Orange', 2.0, 5, '=B3*C3']]).
        visible: Whether to make Excel visible on screen when opened via COM (default: True).
        table_style: Optional Excel table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1', 'TableStyleDark1').
    """
    try:
        res = ExcelEngine.create_workbook(
            filepath=filepath,
            sheet_name=sheet_name,
            headers=headers,
            data=data,
            visible=visible,
            table_style=table_style
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error creating workbook: {str(e)}"


@mcp.tool()
def excel_open_workbook(
    filepath: str,
    read_only: bool = False,
    visible: bool = True
) -> str:
    """
    Open an existing Excel workbook file (.xlsx, .xlsm, .xls, .csv).
    
    Args:
        filepath: Full path to the workbook file.
        read_only: Open the file in read-only mode (default: False).
        visible: Make Excel visible on screen (default: True).
    """
    try:
        res = ExcelEngine.open_workbook(filepath=filepath, read_only=read_only, visible=visible)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error opening workbook: {str(e)}"


@mcp.tool()
def excel_read_range(
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    range_address: Optional[str] = None,
    include_formulas: bool = False,
    output_format: str = "markdown_table"
) -> str:
    """
    Read data and formulas from a worksheet or cell range.
    
    Args:
        filepath: Optional file path. If omitted, reads from active workbook in live Excel.
        sheet_name: Worksheet name (e.g. 'Sheet1', 'Sales'). Defaults to active sheet.
        range_address: Range to read (e.g. 'A1:D10', 'B2', or omit to read entire used range).
        include_formulas: If True, returns formulas (e.g. '=SUM(A1:A5)') instead of evaluated values (default: False).
        output_format: 'markdown_table' (default, easiest for LLM to view), 'json', or 'records'.
    """
    try:
        res = ExcelEngine.read_range(
            filepath=filepath,
            sheet_name=sheet_name,
            range_address=range_address,
            include_formulas=include_formulas,
            output_format=output_format
        )
        if output_format == "markdown_table" and "markdown" in res:
            header = f"**Sheet:** `{res['sheet']}` | **Range:** `{res['range']}` ({res['row_count']} rows x {res['col_count']} cols)\n\n"
            return header + res["markdown"]
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error reading range: {str(e)}"


@mcp.tool()
def excel_write_range(
    data: List[List[Any]],
    start_cell: str = "A1",
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    auto_fit_columns: bool = True
) -> str:
    """
    Write a 2D matrix of values and formulas into a worksheet.
    
    Args:
        data: 2D array of rows and columns (e.g. [['Revenue', 1000], ['Cost', 600], ['Profit', '=B1-B2']]).
        start_cell: Top-left cell address to start writing (default: 'A1').
        filepath: Optional file path. If omitted, writes to active workbook in live Excel.
        sheet_name: Optional sheet name (defaults to active sheet).
        auto_fit_columns: Auto-fit column widths after writing (default: True).
    """
    try:
        res = ExcelEngine.write_range(
            data=data,
            start_cell=start_cell,
            filepath=filepath,
            sheet_name=sheet_name,
            auto_fit_columns=auto_fit_columns
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error writing range: {str(e)}"


@mcp.tool()
def excel_format_range(
    range_address: str,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    number_format: Optional[str] = None,
    horizontal_align: Optional[str] = None,
    border: Optional[str] = None,
    auto_fit: Optional[bool] = None
) -> str:
    """
    Format a cell range with colors, font styling, borders, alignment, and number formats.
    
    Args:
        range_address: Cell range (e.g. 'A1:D1', 'B2:B20').
        filepath: Optional file path. If omitted, formats active workbook in live Excel.
        sheet_name: Optional worksheet name.
        bold: Set font bold (True / False).
        italic: Set font italic (True / False).
        font_size: Font size in points (e.g. 11, 14, 16).
        font_name: Font family name (e.g. 'Segoe UI', 'Calibri', 'Arial').
        font_color: Hex color string for text (e.g. '#FFFFFF', '#002060', '#C00000').
        bg_color: Hex color string for cell fill (e.g. '#1F497D', '#D9E1F2', '#E2EFDA').
        number_format: Excel format string (e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd', '#,##0').
        horizontal_align: Text alignment ('left', 'center', 'right').
        border: Border style ('all_thin', 'outline_thick', 'bottom_double', 'none').
        auto_fit: Auto-fit column widths (True / False).
    """
    try:
        res = ExcelEngine.format_range(
            range_address=range_address,
            filepath=filepath,
            sheet_name=sheet_name,
            bold=bold,
            italic=italic,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            bg_color=bg_color,
            number_format=number_format,
            horizontal_align=horizontal_align,
            border=border,
            auto_fit=auto_fit
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error formatting range: {str(e)}"


@mcp.tool()
def excel_add_chart(
    chart_type: str,
    data_range: str,
    title: Optional[str] = None,
    x_axis_title: Optional[str] = None,
    y_axis_title: Optional[str] = None,
    position_cell: str = "E2",
    width: int = 450,
    height: int = 280,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> str:
    """
    Insert a chart into the worksheet based on a data range.
    
    Args:
        chart_type: Type of chart: 'column_clustered', 'bar_clustered', 'line', 'pie', 'area', 'scatter', 'doughnut'.
        data_range: Range containing headers and data for the chart (e.g. 'A1:B10', 'A1:D6').
        title: Title of the chart.
        x_axis_title: Optional X-axis category label.
        y_axis_title: Optional Y-axis value label.
        position_cell: Top-left cell where the chart will be placed (e.g. 'E2', 'G1').
        width: Chart width in points (default: 450).
        height: Chart height in points (default: 280).
        filepath: Optional file path. If omitted, adds chart to active workbook in live Excel.
        sheet_name: Optional worksheet name.
    """
    try:
        res = ExcelEngine.add_chart(
            chart_type=chart_type,
            data_range=data_range,
            title=title,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            position_cell=position_cell,
            width=width,
            height=height,
            filepath=filepath,
            sheet_name=sheet_name
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding chart: {str(e)}"


@mcp.tool()
def excel_add_sheet(
    sheet_name: str,
    filepath: Optional[str] = None
) -> str:
    """
    Add a new worksheet to the workbook.
    
    Args:
        sheet_name: Name for the new worksheet.
        filepath: Optional file path. If omitted, adds sheet to active workbook in live Excel.
    """
    try:
        res = ExcelEngine.add_sheet(sheet_name=sheet_name, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding sheet: {str(e)}"


@mcp.tool()
def excel_list_sheets(
    filepath: Optional[str] = None
) -> str:
    """
    List all worksheets in the workbook with their dimensions and row/column counts.
    
    Args:
        filepath: Optional file path. If omitted, lists sheets from active workbook in live Excel.
    """
    try:
        res = ExcelEngine.list_sheets(filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error listing sheets: {str(e)}"


@mcp.tool()
def excel_export_pdf(
    output_pdf_path: str,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    orientation: str = "landscape"
) -> str:
    """
    Export the worksheet or entire workbook to a PDF file.
    
    Args:
        output_pdf_path: Destination file path for the generated PDF (e.g. 'E:/reports/report.pdf').
        filepath: Optional source Excel file path. If omitted, exports active workbook from live Excel.
        sheet_name: Optional specific sheet to export. If omitted, exports the active sheet or workbook.
        orientation: Page orientation ('landscape' or 'portrait'). Default: 'landscape'.
    """
    try:
        res = ExcelEngine.export_pdf(
            output_pdf_path=output_pdf_path,
            filepath=filepath,
            sheet_name=sheet_name,
            orientation=orientation
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error exporting PDF: {str(e)}"


@mcp.tool()
def excel_save_workbook(
    save_as_path: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Save the currently active workbook or save it to a new file path.
    
    Args:
        save_as_path: Optional new destination path to save as.
        filepath: Optional specific workbook file to save.
    """
    try:
        res = ExcelEngine.save_workbook(save_as_path=save_as_path, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error saving workbook: {str(e)}"


@mcp.tool()
def excel_analyze_sheet(
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    has_headers: bool = True
) -> str:
    """
    Generate an in-depth statistical analysis, column schema profile, and data distribution of a worksheet.
    
    Args:
        filepath: Optional file path. If omitted, analyzes active worksheet in live Excel.
        sheet_name: Optional sheet name.
        has_headers: Whether the first row contains column headers (default: True).
    """
    try:
        res = ExcelEngine.analyze_sheet(
            filepath=filepath,
            sheet_name=sheet_name,
            has_headers=has_headers
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error analyzing sheet: {str(e)}"


@mcp.tool()
def excel_execute_script(
    code: str,
    filepath: Optional[str] = None
) -> str:
    """
    Execute arbitrary Python automation code with direct access to Excel and OpenPyXL objects.
    
    Injected variables available in code:
      - `app`: Excel.Application COM instance (win32com.client)
      - `wb`: Active Workbook COM instance
      - `ws`: Active Worksheet COM instance
      - `win32com`: win32com.client module
      - `openpyxl`: openpyxl module
      - `engine`: ExcelEngine class
      - `filepath`: target filepath if provided
      
    Args:
        code: Python script string to execute.
        filepath: Optional target workbook file path.
    """
    try:
        res = ExcelEngine.execute_script(code=code, filepath=filepath)
        if res.get("success"):
            return f"✅ Execution Successful:\n{res.get('output')}"
        else:
            return f"❌ Execution Failed:\nError: {res.get('error')}\nOutput:\n{res.get('output')}"
    except Exception as e:
        return f"❌ Error executing script: {str(e)}"


def main():
    mcp.run()


if __name__ == "__main__":
    main()
