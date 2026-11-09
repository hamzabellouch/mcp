# MCP Unified Test Suite

This folder consolidates all automated test suites for the 5 MCP servers:

| MCP Server | Test File | Description |
| :--- | :--- | :--- |
| **Blender** | `test_blender.py` | FastMCP schemas, tool dispatch, offline fallback, mock HTTP bridge |
| **Excel** | `test_excel.py` | Workbook creation, range formats, formulas, charts, analysis, scripts |
| **FreeCAD** | `test_freecad.py` | All 29 tools, headless CAD pipeline, RPC commands, bridge security |
| **PowerPoint** | `test_powerpoint.py` | 26 tools, slides, metric cards, styled tables, charts, timelines, themes |
| **Word** | `test_word.py` | 22 tools, headings, formatting, tables, callouts, TOC, comments, templates |

---

## Running Tests

### 1. Run All Tests
```bash
python Tests/run_all.py
```

### 2. Run an Individual Suite
```bash
python Tests/run_all.py blender
python Tests/run_all.py excel
python Tests/run_all.py freecad
python Tests/run_all.py powerpoint
python Tests/run_all.py word
```

Or run the specific test file directly:
```bash
python Tests/test_excel.py
python Tests/test_word.py
```
