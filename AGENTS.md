# MCP Suite Agent Guidelines & Architecture Rules

This document provides operational rules and architectural standards for AI agents (Claude, Gemini, Cursor, ChatGPT, Copilot, etc.) interacting with this repository.

---

## 🏛️ Monorepo Structure

```
MCP/
├── Blender/       # Blender 3D MCP Server & Blender Add-on
├── Excel/         # Microsoft Excel MCP Server (COM & openpyxl)
├── FreeCAD/       # FreeCAD Parametric CAD MCP Server & RPC Addon
├── Powerpoint/    # Microsoft PowerPoint MCP Server (COM & python-pptx)
├── Word/          # Microsoft Word MCP Server (COM & python-docx)
├── Tests/         # Centralized verification test suite for all 5 MCPs
├── .gitignore     # Unified global Git ignore rules
├── AGENTS.md      # AI Agent operational guidelines (this file)
└── README.md      # Suite overview and setup guide
```

---

## 🛠️ MCP Operational Directives by Tool

### 1. FreeCAD MCP
- **In-Memory Execution First**: Execute CAD scripts and parametric operations in-memory via the live RPC bridge. Do **not** dump scratch `.py` files onto disk unless explicitly requested by the user.
- **Fast Mode & No Default Screenshots**: Do not capture or return 3D viewport screenshots on every tool call. Capture screenshots **only** when the user explicitly requests visual output.
- **Progressive Step-by-Step Construction**:
  1. *Stage 1 (2D Wireframe/Sketch)*: Draw contours, reference points, and constraints first.
  2. *Stage 2 (Planar Face)*: Form closed planar faces.
  3. *Stage 3 (3D Extrusion/Pocketing)*: Pad or pocket in observable depth stages.
  4. *Stage 4 (Finishing & Styling)*: Apply fillets, chamfers, boolean operations, and materials.
- **Live Visual Feedback**: Include `FreeCADGui.updateGui()` with short delays in interactive scripts so the user can follow the construction live in the FreeCAD viewport.

### 2. Word MCP
- **Dual Engine Architecture**: Automatically attaches to the user's running Word instance via COM. If Word is not running, falls back safely to headless `python-docx` file manipulation.
- **Window Safety**: Never silently overwrite a file on disk behind an open Word window. Never force `Quit()` on an instance containing user documents.
- **Alert Restoration**: Always restore `DisplayAlerts` after tool operations.
- **Lock Files**: Respect `~$*.docx` lock files created by Word; do not attempt to delete them while Word is running.

### 3. Excel MCP
- **Dual Engine Architecture**: Supports live COM automation for open workbooks and high-speed `openpyxl` processing for headless batch operations.
- **Data Integrity**: Use proper uppercase formula syntax (`=SUM(...)`, `=AVERAGE(...)`) and avoid hardcoding values where dynamic formulas are expected.
- **Chart Formatting**: Ensure data ranges for charts cover valid series and category headers with sensible coordinate positioning.

### 4. PowerPoint MCP
- **Slide Composition**: Default to modern 16:9 widescreen presentations (`widescreen=True`).
- **Visual Design**: Leverage built-in executive cards, styled tables with alternating row colors, structured bullet hierarchy, and native charts.
- **Safe Fallback**: Use `python-pptx` for fast headless deck building; use live COM for high-resolution PDF and slide image exports.

### 5. Blender MCP
- **Bridge Connectivity**: Interacts with Blender via the lightweight socket bridge (`blender_addon.py`).
- **Scene Hygiene**: Check scene summary before clearing objects; respect existing user cameras and lights unless explicitly asked to rebuild studio lighting.
- **Material Nodes**: When generating procedural PBR shaders, use the Principled BSDF node setup with normalized color tuples `[R, G, B, A]` in `0.0 - 1.0` range.

---

## 🧪 Testing Guidelines

All tests are centralized in the [`Tests/`](Tests) directory:
- **Runner**: Use `python Tests/run_all.py` to run all suites, or specify an individual suite (e.g. `python Tests/run_all.py word`).
- **Artifacts**: Tests must clean up their own temporary test files upon completion.
- **No Test Files in Subprojects**: Do not place new test files inside individual MCP project folders; keep all test logic within `Tests/`.
