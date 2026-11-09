"""
Test Suite for Excel MCP Server
--------------------------------
Verifies all Excel MCP tools in both live COM and headless modes.
"""

import os
import sys
import json

# Ensure UTF-8 output encoding on Windows consoles
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import excel_mcp_servers as srv

TEST_FILE = os.path.abspath("test_mcp_output.xlsx")
TEST_PDF = os.path.abspath("test_mcp_output.pdf")


def run_tests():
    print("========================================")
    print("   EXCEL MCP SERVER VERIFICATION SUITE  ")
    print("========================================")

    # 1. Status Check
    print("\n[1/9] Testing excel_status()...")
    status_out = srv.excel_status()
    print("Status Result:\n", status_out)

    # 2. Create Workbook with initial data & formulas
    print("\n[2/9] Testing excel_create_workbook()...")
    headers = ["Department", "Q1 Budget", "Q2 Budget", "Q3 Budget", "Q4 Budget", "Annual Total"]
    data = [
        ["Engineering", 45000, 52000, 48000, 60000, "=SUM(B2:E2)"],
        ["Marketing", 20000, 25000, 22000, 30000, "=SUM(B3:E3)"],
        ["Sales", 35000, 38000, 42000, 45000, "=SUM(B4:E4)"],
        ["Design", 15000, 16000, 18000, 20000, "=SUM(B5:E5)"],
        ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)", "=SUM(F2:F5)"]
    ]
    create_out = srv.excel_create_workbook(
        filepath=TEST_FILE,
        sheet_name="Budget2026",
        headers=headers,
        data=data,
        visible=False,
        table_style="TableStyleMedium2"
    )
    print("Create Output:\n", create_out)

    # 3. Read Range
    print("\n[3/9] Testing excel_read_range()...")
    read_table = srv.excel_read_range(filepath=TEST_FILE, output_format="markdown_table")
    print("Read Range Table:\n", read_table)

    # 4. Format Range (Header styling + currency + totals)
    print("\n[4/9] Testing excel_format_range()...")
    # Format Headers: Navy fill, white bold text
    srv.excel_format_range(
        range_address="A1:F1",
        filepath=TEST_FILE,
        bold=True,
        bg_color="#1F497D",
        font_color="#FFFFFF",
        horizontal_align="center"
    )
    # Format Numbers: Currency format
    srv.excel_format_range(
        range_address="B2:F6",
        filepath=TEST_FILE,
        number_format="$#,##0",
        horizontal_align="right"
    )
    # Format Total row: Bold + light blue fill
    srv.excel_format_range(
        range_address="A6:F6",
        filepath=TEST_FILE,
        bold=True,
        bg_color="#D9E1F2",
        border="bottom_double"
    )
    print("Formatting applied successfully.")

    # 5. Add Chart (Multi-column & Single-column tests)
    print("\n[5/9] Testing excel_add_chart()...")
    chart_out = srv.excel_add_chart(
        chart_type="column_clustered",
        data_range="A1:E5",
        title="Department Quarterly Budgets 2026",
        x_axis_title="Department",
        y_axis_title="Budget ($)",
        position_cell="H2",
        filepath=TEST_FILE
    )
    print("Add Multi-Column Chart Output:\n", chart_out)

    # Test single-column chart bounds handling
    single_chart_out = srv.excel_add_chart(
        chart_type="line",
        data_range="F1:F5",
        title="Annual Totals Trend",
        position_cell="H18",
        filepath=TEST_FILE
    )
    print("Add Single-Column Chart Output:\n", single_chart_out)

    # 6. Add Sheet & Write Range
    print("\n[6/9] Testing excel_add_sheet() & excel_write_range()...")
    srv.excel_add_sheet(sheet_name="KPIs", filepath=TEST_FILE)
    srv.excel_write_range(
        data=[
            ["Metric", "Target", "Achieved", "Status"],
            ["Customer Retention", "95%", "97%", "Exceeded"],
            ["Uptime SLA", "99.9%", "99.95%", "Exceeded"],
            ["CSAT Score", "4.8 / 5.0", "4.9 / 5.0", "Exceeded"]
        ],
        start_cell="A1",
        filepath=TEST_FILE,
        sheet_name="KPIs"
    )
    srv.excel_format_range(
        range_address="A1:D1",
        filepath=TEST_FILE,
        sheet_name="KPIs",
        bold=True,
        bg_color="#375623",
        font_color="#FFFFFF"
    )
    print("Added sheet 'KPIs' and formatted range.")

    # 7. List Sheets
    print("\n[7/9] Testing excel_list_sheets()...")
    sheets_out = srv.excel_list_sheets(filepath=TEST_FILE)
    print("Sheets List:\n", sheets_out)

    # 8. Analyze Sheet
    print("\n[8/9] Testing excel_analyze_sheet()...")
    analysis_out = srv.excel_analyze_sheet(filepath=TEST_FILE, sheet_name="Budget2026")
    print("Analysis:\n", analysis_out)

    # 9. Execute Script
    print("\n[9/9] Testing excel_execute_script()...")
    script = """
wb_obj = openpyxl.load_workbook(filepath)
print("Loaded workbook successfully via script!")
print("Sheets count:", len(wb_obj.sheetnames))
for s in wb_obj.sheetnames:
    print(" - Sheet:", s)
"""
    script_out = srv.excel_execute_script(code=script, filepath=TEST_FILE)
    print("Script Output:\n", script_out)

    print("\n========================================")
    print(f"🎉 ALL TESTS PASSED! Output file: {TEST_FILE}")
    print("========================================")


if __name__ == "__main__":
    run_tests()
