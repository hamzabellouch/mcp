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
| **[Word MCP](Word/)** | [`./Word`](Word/) | **22 Tools** | Live Windows COM + Headless `python-docx` + `lxml` | Document creation, typography, tables, TOC, watermarks, RTL/Arabic, comments, template filling, PDF export. |
| **[PowerPoint MCP](Powerpoint/)** | [`./Powerpoint`](Powerpoint/) | **26 Tools** | Live Windows COM + Headless `python-pptx` | Presentation lifecycle, slide layout, KPI cards, charts, roadmaps/timelines, themes, regex replace, slide PNG/PDF export. |
| **[Excel MCP](Excel/)** | [`./Excel`](Excel/) | **13 Tools** | Live Windows COM + Headless `openpyxl` | Workbook creation, range read/write, styling, formula execution, column/pie/line charts, statistical data profiling. |
| **[FreeCAD MCP](FreeCAD/)** | [`./FreeCAD`](FreeCAD/) | **29 Tools** | Live Qt GUI RPC (`9875`) + Headless `FreeCADCmd` | 3D CAD modeling, boolean operations, sketch extrusion, TechDraw 2D sheets, CalculiX FEM structural simulations, mass analysis. |
| **[Blender MCP](Blender/)** | [`./Blender`](Blender/) | **7 Tools** | FastMCP stdio + Blender Main-Thread Bridge (`9876`) | 3D primitives, PBR material creation, model importing (GLTF/OBJ/FBX), camera & lighting setup, Cycles/Eevee rendering. |

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
      "args": ["-m", "word_mcp_server"],
      "cwd": "/path/to/mcp/Word"
    },
    "powerpoint": {
      "command": "python",
      "args": ["-m", "powerpoint_mcp_server"],
      "cwd": "/path/to/mcp/Powerpoint"
    },
    "excel": {
      "command": "python",
      "args": ["-m", "excel_mcp_server"],
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
      "args": ["/path/to/mcp/Word/word_mcp_server.py"]
    },
    "powerpoint": {
      "command": "python",
      "args": ["/path/to/mcp/Powerpoint/powerpoint_mcp_server.py"]
    },
    "excel": {
      "command": "python",
      "args": ["/path/to/mcp/Excel/excel_mcp_server.py"]
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
      "args": ["/path/to/mcp/Word/word_mcp_server.py"]
    },
    "powerpoint": {
      "command": "python",
      "args": ["/path/to/mcp/Powerpoint/powerpoint_mcp_server.py"]
    },
    "excel": {
      "command": "python",
      "args": ["/path/to/mcp/Excel/excel_mcp_server.py"]
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
├── Word/                    # Microsoft Word MCP Server (22 tools)
│   ├── word_engine.py       # Dual COM & python-docx engine
│   ├── word_mcp_server.py   # FastMCP server definition
│   ├── test_mcp.py          # Verification test suite (16 tests)
│   ├── pyproject.toml       # Package metadata
│   └── README.md            # Word server documentation
│
├── Powerpoint/              # PowerPoint MCP Server (26 tools)
│   ├── powerpoint_engine.py # Dual COM & python-pptx engine
│   ├── powerpoint_mcp_server.py # FastMCP server definition
│   ├── test_mcp.py          # Verification test suite (26 tools)
│   ├── pyproject.toml       # Package metadata
│   └── README.md            # PowerPoint server documentation
│
├── Excel/                   # Microsoft Excel MCP Server (13 tools)
│   ├── excel_engine.py      # Dual COM & openpyxl engine
│   ├── excel_mcp_server.py  # FastMCP server definition
│   ├── test_mcp.py          # Verification test suite (9 suites)
│   ├── pyproject.toml       # Package metadata
│   └── README.md            # Excel server documentation
│
├── FreeCAD/                 # FreeCAD Engineering MCP Server (29 tools)
│   ├── addon/               # FreeCAD GUI XML-RPC Addon & Bridge
│   ├── freecad_engine.py    # Headless & RPC execution coordinator
│   ├── freecad_mcp_server.py# FastMCP server definition
│   ├── test_mcp.py          # Integration verification suite
│   ├── pyproject.toml       # Package metadata
│   └── README.md            # FreeCAD server documentation
│
└── Blender/                 # Blender 3D MCP Server (7 tools)
    ├── blender_addon.py     # Blender main-thread HTTP bridge addon
    ├── blender_mcp_server.py# FastMCP server definition
    ├── test_mcp.py          # Unit & mocked integration suite (20 tests)
    ├── pyproject.toml       # Package metadata
    ├── assets/              # Showcase and render previews
    └── README.md            # Blender server documentation
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

