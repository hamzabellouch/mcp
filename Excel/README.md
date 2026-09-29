# 📊 Excel MCP

<h3 align="center">Dual-Engine Spreadsheet Automation, Statistical Data Profiling & Interactive Charting.</h3>

<p align="center">
Empower AI assistants (Claude Desktop, Cursor, Antigravity, VS Code, Cline) to read, write, format, analyze spreadsheets, and generate formulas and charts across Windows, macOS, and Linux.
</p>

## Overview

**Microsoft Excel MCP Server** is a production-ready Model Context Protocol (MCP) server providing AI assistants with high-level tools to interact with Microsoft Excel and manipulate spreadsheet files (`.xlsx`, `.xlsm`, `.xls`, `.csv`).

It features a hybrid dual-engine design combining live interactive desktop control via Windows COM with cross-platform headless processing via `openpyxl`, allowing agents to create styled financial workbooks, profile column data distributions, and build native charts seamlessly.

## ⚡ Dual-Engine Architecture

This server features a hybrid dual-engine design tailored for both interactive desktop workflows and automated headless pipelines:

```
                      ┌──────────────────────────────┐
                      │    AI Client / MCP Host      │
                      └──────────────┬───────────────┘
                                     │ (JSON-RPC / stdio)
                      ┌──────────────▼───────────────┐
                      │    excel-mcp-server          │
                      └──────────────┬───────────────┘
                                     │
                 ┌───────────────────┴───────────────────┐
                 │                                       │
      ┌──────────▼──────────┐                 ┌──────────▼──────────┐
      │  Live Windows COM   │                 │  Headless OpenPyXL  │
      │  (Interactive UI)   │                 │  (Fast File Engine) │
      └──────────┬──────────┘                 └──────────┬──────────┘
                 │                                       │
      ┌──────────▼──────────┐                 ┌──────────▼──────────┐
      │ Microsoft Excel App │                 │ Direct .xlsx Files  │
      │  (Windows Desktop)  │                 │  (Win / Mac / Linux)│
      └─────────────────────┘                 └─────────────────────┘
```

1. **Live Windows COM Automation (`pywin32`)**:
   - Real-time interactive control over the Microsoft Excel desktop application.
   - Ideal for visual walkthroughs, live formula evaluations, interactive charting, and presentation generation on Windows.
2. **Headless Cross-Platform Engine (`openpyxl`)**:
   - Direct file reading and writing without requiring Microsoft Excel to be installed or running.
   - 100% compatible with Windows, macOS, and Linux server environments.

## 🛠️ Complete MCP Tools Catalog (13 Tools)

| Tool Name | Engine Support | Description | Key Arguments |
| :--- | :--- | :--- | :--- |
| `excel_status` | COM / Headless | Checks whether Excel is running, active workbooks, sheets, and engine availability. | *None* |
| `excel_create_workbook` | COM + openpyxl | Creates a new workbook with optional sheet name, headers, 2D data matrix, and table styles. | `filepath`, `sheet_name`, `headers`, `data`, `visible`, `table_style` |
| `excel_open_workbook` | COM + openpyxl | Opens an existing workbook file (`.xlsx`, `.xlsm`, `.xls`, `.csv`). | `filepath`, `read_only`, `visible` |
| `excel_read_range` | COM + openpyxl | Reads values or formulas from a range or sheet in Markdown table, JSON, or record formats. | `filepath`, `sheet_name`, `range_address`, `include_formulas`, `output_format` |
| `excel_write_range` | COM + openpyxl | Writes a 2D matrix of values and formulas into cells with optional column AutoFit. | `data`, `start_cell`, `filepath`, `sheet_name`, `auto_fit_columns` |
| `excel_format_range` | COM + openpyxl | Formats cells (colors, fonts, bold/italic, borders, alignments, and number formats). | `range_address`, `filepath`, `sheet_name`, `bold`, `font_size`, `font_color`, `bg_color`, `number_format`, `border` |
| `excel_add_chart` | COM + openpyxl | Inserts charts (Column, Bar, Line, Pie, Area, Scatter, Doughnut) linked to a data range. | `chart_type`, `data_range`, `title`, `x_axis_title`, `y_axis_title`, `position_cell`, `filepath` |
| `excel_add_sheet` | COM + openpyxl | Adds a new worksheet to an active or specified workbook. | `sheet_name`, `filepath` |
| `excel_list_sheets` | COM + openpyxl | Lists all worksheets, used range dimensions, row counts, and column counts. | `filepath` |
| `excel_export_pdf` | COM (Excel) | Exports a worksheet or entire workbook to a professional PDF document. | `output_pdf_path`, `filepath`, `sheet_name`, `orientation` |
| `excel_save_workbook` | COM + openpyxl | Saves the active workbook or saves a copy to a new file path. | `save_as_path`, `filepath` |
| `excel_analyze_sheet` | Analysis Engine | Generates statistical data profiles, schema detection, null counts, min/max/mean/sum aggregates. | `filepath`, `sheet_name`, `has_headers` |
| `excel_execute_script` | Custom Automation | Executes arbitrary Python automation code with pre-injected COM and openpyxl objects. | `code`, `filepath` |

## 🔒 Security Notice

The `excel_execute_script` tool allows arbitrary Python code execution to provide unlimited flexibility for custom macros and low-level Excel COM manipulations.

> [!WARNING]
> **Execute Scripts with Caution**: Always verify the generated Python script before granting execution permissions in untrusted or production environments.

## 📦 Installation & Setup

### 1. Prerequisites
- **Python**: Version 3.10 or higher.
- **Operating System**: Windows (for full COM automation + Headless) or macOS / Linux (Headless mode via openpyxl).

### 2. Install Dependencies
Clone this repository and install dependencies:

```bash
git clone https://github.com/example/excel-mcp-server.git
cd excel-mcp-server

# Option A: Install from requirements.txt
pip install -r requirements.txt

# Option B: Install package in editable mode
pip install -e .
```

## 🔌 Client Configurations

Add the Excel MCP server to your favorite MCP client configuration file:

### 1. Claude Desktop (`claude_desktop_config.json`)
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "/path/to/excel-mcp-server/excel_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

*If installed via `pip install -e .`, you can alternatively use:*
```json
{
  "mcpServers": {
    "excel": {
      "command": "excel-mcp"
    }
  }
}
```

### 2. Cursor (`mcp.json` / Cursor Settings)
Navigate to **Cursor Settings > Features > MCP > Add New MCP Server**:
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "/path/to/excel-mcp-server/excel_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

### 3. Antigravity / Gemini CLI (`mcp_config.json`)
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "/path/to/excel-mcp-server/excel_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

### 4. VS Code (Cline / Roo Code / Continue)
Add to your extension's MCP configuration settings:
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "/path/to/excel-mcp-server/excel_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

## 🧪 Verification & Testing

Run the included automated verification suite to test all 13 tools and engines:

```bash
python test_mcp.py
```

Expected output:
```text
========================================
   EXCEL MCP SERVER VERIFICATION SUITE  
========================================

[1/9] Testing excel_status()...
[2/9] Testing excel_create_workbook()...
[3/9] Testing excel_read_range()...
[4/9] Testing excel_format_range()...
[5/9] Testing excel_add_chart()...
[6/9] Testing excel_add_sheet() & excel_write_range()...
[7/9] Testing excel_list_sheets()...
[8/9] Testing excel_analyze_sheet()...
[9/9] Testing excel_execute_script()...

========================================
🎉 ALL TESTS PASSED! Output file: test_mcp_output.xlsx
========================================
```

> [!WARNING]
> Automated spreadsheet manipulations and formula computations can modify existing workbooks. There is always a possibility of error, so we assume no responsibility for any inaccuracies.

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

