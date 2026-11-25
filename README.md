# MCP Suite

<h3 align="center">Modern, Production-Ready MCP Servers for Office Automation, 3D Graphics & Engineering CAD.</h3>

<p align="center">
Empower AI assistants (Claude Desktop, Google Antigravity, Cursor, VS Code) to automate Word, Excel, PowerPoint, FreeCAD, and Blender seamlessly with over 95 high-level tools.
</p>

<img width="1280" height="720" alt="MCP" src="https://github.com/user-attachments/assets/225d7a3a-9e8e-48bb-8d2f-946a3638874b" />


## Overview

The **Enterprise Model Context Protocol (MCP) Suite** is a unified collection of production-grade MCP servers designed to bridge Large Language Models with desktop productivity applications and engineering software.

By combining interactive live desktop automation (via Windows COM and GUI RPC bridges) with cross-platform headless engines (`python-docx`, `python-pptx`, `openpyxl`, `FreeCADCmd`), this suite enables AI agents to generate structured documents, analyze financial spreadsheets, build executive slide presentations, execute finite-element CAD simulations, and render photorealistic 3D scenes through natural language.

## 🌟 Included MCP Servers

| MCP Server | Directory | Tools Count | Architecture / Engine | Primary Capabilities |
| :--- | :--- | :---: | :--- | :--- |
| **[Word MCP](Word/)** | [`./Word`](Word/) | **24 Tools** | Live Windows COM + Headless `python-docx` + `lxml` | Document creation, typography, inline rich formatting, tables/merges, TOC, watermarks, RTL/Arabic, comments, `.docx`/`.dotx` templates, PDF export. |
| **[PowerPoint MCP](Powerpoint/)** | [`./Powerpoint`](Powerpoint/) | **26 Tools** | Live Windows COM + Headless `python-pptx` | Presentation lifecycle, slide layout, KPI cards, charts, roadmaps/timelines, themes, regex replace, slide PNG/PDF export. |
| **[Excel MCP](Excel/)** | [`./Excel`](Excel/) | **14 Tools** | Live Windows COM + Headless `openpyxl` | Workbook creation, range read/write, styling, formula execution, column/pie/line/scatter/doughnut charts, statistical data profiling. |
| **[FreeCAD MCP](FreeCAD/)** | [`./FreeCAD`](FreeCAD/) | **33 Tools** | Live Qt GUI RPC (`2607`) + Headless `FreeCADCmd` | 3D CAD modeling, boolean operations, sketch extrusion, TechDraw 2D sheets, CalculiX FEM structural simulations, mass analysis. |
| **[Blender MCP](Blender/)** | [`./Blender`](Blender/) | **8 Tools** | FastMCP stdio + Blender Main-Thread Bridge (`9876`) | 3D primitives, PBR material creation, model importing (GLTF/OBJ/FBX), camera & lighting setup, Cycles/Eevee rendering. |

## 🚀 Quick Setup & Client Configuration

You can connect all 5 MCP servers to your favorite MCP-compatible AI client using standard JSON configuration.

### 1. Claude Desktop (`claude_desktop_config.json`)
Located at:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["-m", "word_mcp_servers"],
      "cwd": "/path/to/mcp/Word"
    },
    "powerpoint": {
      "command": "python",
      "args": ["-m", "powerpoint_mcp_servers"],
      "cwd": "/path/to/mcp/Powerpoint"
    },
    "excel": {
      "command": "python",
      "args": ["-m", "excel_mcp_servers"],
      "cwd": "/path/to/mcp/Excel"
    },
    "freecad": {
      "command": "python",
      "args": ["-m", "freecad_mcp_server"],
      "cwd": "/path/to/mcp/FreeCAD"
    },
    "blender": {
      "command": "python",
      "args": ["-m", "blender_mcp_server"],
      "cwd": "/path/to/mcp/Blender"
    }
  }
}
```

### 2. Cursor (`.cursor/mcp.json`)
```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["/path/to/mcp/Word/word_mcp_servers.py"]
    },
    "powerpoint": {
      "command": "python",
      "args": ["/path/to/mcp/Powerpoint/powerpoint_mcp_servers.py"]
    },
    "excel": {
      "command": "python",
      "args": ["/path/to/mcp/Excel/excel_mcp_servers.py"]
    },
    "freecad": {
      "command": "python",
      "args": ["/path/to/mcp/FreeCAD/freecad_mcp_server.py"]
    },
    "blender": {
      "command": "python",
      "args": ["/path/to/mcp/Blender/blender_mcp_server.py"]
    }
  }
}
```

### 3. Google Antigravity / Gemini CLI (`mcp_config.json`)
```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["/path/to/mcp/Word/word_mcp_servers.py"]
    },
    "powerpoint": {
      "command": "python",
      "args": ["/path/to/mcp/Powerpoint/powerpoint_mcp_servers.py"]
    },
    "excel": {
      "command": "python",
      "args": ["/path/to/mcp/Excel/excel_mcp_servers.py"]
    },
    "freecad": {
      "command": "python",
      "args": ["/path/to/mcp/FreeCAD/freecad_mcp_server.py"]
    },
    "blender": {
      "command": "python",
      "args": ["/path/to/mcp/Blender/blender_mcp_server.py"]
    }
  }
}
```

## 📦 Monorepo Directory Structure

```text
mcp/
├── .gitignore               # Unified monorepo ignore rules
├── README.md                # Main suite documentation & overview
│
├── Word/                    # Microsoft Word MCP Server (24 tools)
│   ├── word_mcp_servers.py  # Unified COM, python-docx & FastMCP server
│   ├── requirements.txt     # Python dependencies
│   ├── pyproject.toml       # Package metadata
│   └── README.md            # Word server documentation
│
├── Powerpoint/                # PowerPoint MCP Server (26 tools)
│   ├── powerpoint_mcp_servers.py # Unified COM, python-pptx & FastMCP server
│   ├── requirements.txt       # Python dependencies
│   ├── pyproject.toml         # Package metadata
│   └── README.md              # PowerPoint server documentation
│
├── Excel/                     # Microsoft Excel MCP Server (14 tools)
│   ├── excel_mcp_servers.py   # Unified COM, openpyxl & FastMCP server
│   ├── requirements.txt       # Python dependencies
│   ├── pyproject.toml         # Package metadata
│   └── README.md              # Excel server documentation
│
├── FreeCAD/                   # FreeCAD Engineering MCP Server (33 tools)
│   ├── add-on/                # FreeCAD GUI XML-RPC Addon & Bridge
│   ├── freecad_mcp_server.py  # Unified Headless, RPC & FastMCP server
│   ├── requirements.txt       # Python dependencies
│   ├── pyproject.toml         # Package metadata
│   └── README.md              # FreeCAD server documentation
│
├── Blender/                   # Blender 3D MCP Server (8 tools)
│   ├── Add-on/blender_addon.py# Blender main-thread HTTP bridge addon
│   ├── blender_mcp_server.py  # FastMCP server definition
│   ├── pyproject.toml         # Package metadata
│   ├── assets/                # Showcase and render previews
│   └── README.md              # Blender server documentation
│
└── Tests/                     # Centralized verification test suite
    ├── run_all.py             # Unified test runner for all 5 MCPs
    ├── test_word.py
    ├── test_powerpoint.py
    ├── test_excel.py
    ├── test_freecad.py
    └── test_blender.py
```

## 🛠️ Global Installation & Dependencies

To install all dependencies across the entire suite:

```bash
# Core MCP SDK
pip install mcp requests

# Office Automation Suite
pip install python-docx python-pptx openpyxl pillow lxml
pip install pywin32; sys_platform == 'win32'
```

## 🔒 Security & Architecture Standards

- **Safe Script Execution**: Tools executing dynamic Python code (`execute_script`) are isolated within dedicated execution namespaces and capture standard error buffers.
- **Cross-Origin & Local Loopback Protection**: Blender and FreeCAD local bridges enforce localhost origin checks and loopback bindings (`127.0.0.1`), eliminating remote CSRF and DNS rebinding risks.
- **Process & Resource Cleanup**: Headless PDF and image export workers guarantee process termination and tempfile removal in structured `finally` blocks.

> [!WARNING]
> Automation and script execution tools interact directly with local files and desktop applications. There is always a possibility of error, so we assume no responsibility for any inaccuracies or unintended changes.

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

