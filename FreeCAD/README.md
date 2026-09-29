# 📐 FreeCAD MCP

<h3 align="center">Modern, High-Performance FreeCAD Automation, Parametric 3D CAD, FEM Simulation & Technical Drawing MCP Server.</h3>

<p align="center">
Connect AI assistants directly to FreeCAD for automated parametric 3D modeling, technical drawings (TechDraw), finite element analysis (FEM), live viewport screenshots, and background code execution.
</p>

## Overview

**FreeCAD MCP Server** is a robust, production-ready Model Context Protocol (MCP) server that empowers AI assistants (Antigravity, Claude Desktop, Cursor, VS Code, Roo Code, Cline) to interact directly with **FreeCAD**.

It provides an end-to-end bridge between Large Language Models and FreeCAD's parametric C++ kernel and Python APIs. With support for dual execution modes (live GUI desktop session and headless command-line execution), AI agents can construct 3D CAD models step-by-step, inspect geometries, run structural FEA simulations with CalculiX, generate engineering projection sheets, and export/import standard CAD formats (`STEP`, `STL`, `OBJ`, `IGES`, `BREP`, `DXF`, `SVG`, `PDF`).

FreeCAD MCP is built with thread-safe **Qt Queued Signals**, isolated subprocess crash guards for OpenCASCADE (OCCT), non-blocking asynchronous execution with `commit()` helpers, and live high-resolution viewport image feedback.

## 🏗️ Architecture Overview

```
 ┌────────────────────────────────────────────────────────┐
 │     AI Assistant (Antigravity / Claude / Cursor)       │
 └──────────────────────────┬─────────────────────────────┘
                            │ Model Context Protocol (MCP)
                            ▼
 ┌────────────────────────────────────────────────────────┐
 │                FreeCAD MCP Server                      │
 │     FastMCP Server · Parameter Validation · Auth       │
 └─────────────┬──────────────────────────┬───────────────┘
               │ XML-RPC (Port 9875)      │ Subprocess / CLI
               ▼                          ▼
 ┌───────────────────────────┐ ┌──────────────────────────┐
 │  FreeCAD GUI Addon/Bridge │ │  Headless FreeCADCmd     │
 │  • Qt Signal Dispatcher   │ │  • Isolated Execution    │
 │  • Main Thread Safety     │ │  • Native Crash Guard    │
 │  • Viewport Screenshots   │ │  • Batch CAD Generation  │
 │  • CalculiX FEM Runner    │ │  • OpenCASCADE Workers   │
 └───────────────────────────┘ └──────────────────────────┘
```

### Dual-Engine Execution Pipeline
1. **Live GUI Mode (XML-RPC Port `9875` & HTTP Bridge Port `9877`)**: Communicates with an active FreeCAD GUI session. Commands are dispatched to the Qt main thread via queued signals (`QTimer` / custom event waker), ensuring responsive UI, instant 3D viewport rendering, and live PNG screenshot capture.
2. **Headless Execution (`FreeCADCmd`)**: Heavy OpenCASCADE boolean operations, mesh tessellations, or batch exports can run in isolated subprocesses. Any native C++ crashes (segfaults) are safely trapped without risking unsaved GUI work.

## ⭐ Key Features & Capabilities

| Feature | API / Method | Performance | Safety / Isolation | Description |
| :--- | :--- | :--- | :--- | :--- |
| **Live GUI Interaction** | XML-RPC (`Port 9875`) & Qt Queued Signals | **Real-Time** (< 50ms) | **Thread-Safe** | Executes code directly on FreeCAD's main GUI thread without UI freezing or race conditions. |
| **Headless Isolated Execution** | `FreeCADCmd` Subprocess Runner | **High-Speed** | **Crash-Isolated** | Runs heavy OCCT scripts in separate processes; native C++ crashes never affect open GUI documents. |
| **Asynchronous Computation** | Background Threading & `commit()` Helper | **Non-Blocking** | **Scenegraph Protected** | Performs heavy OCCT booleans/lofts in the background while safely updating documents via GUI commits. |
| **Live 3D Viewport Screenshots** | Offscreen Render & Coin3D Camera Views | **Instant** (< 200ms) | **Base64 Encoded** | Returns visual PNG images after each operation from Isometric, Front, Top, Right, or Custom angles. |
| **Parametric Modeling Suite** | `Part`, `PartDesign`, `Sketcher`, `Draft` | **Instant** | **Parametric History** | Full primitive creation, boolean operations (Fuse/Cut/Common), Fillets, Chamfers, Sketches, and Revolutions. |
| **Document Management** | Open, Save, Create, Reload, List `.FCStd` | **Instant** | **Disk & State Synced** | Open existing `.FCStd` files, save models with custom paths, or reload after headless edits. |
| **FEM Structural Simulation** | `CalculiX` Solver & `Gmsh` Meshing | **Automated** | **Full Engineering FEA** | Solves finite element analyses, computing max von Mises stress (MPa), node displacements, and mesh metrics. |
| **Technical Drawings (TechDraw)** | `TechDraw` Workbench & SVG Templates | **Fast** | **Vector Accurate** | Generates multi-view projection sheets (Front, Top, Right, Isometric) and exports to SVG, DXF, and PDF. |
| **Universal File Import & Export** | `Import`, `Mesh`, & OCCT Exporters | **High-Speed** | **Format Compliant** | Bi-directional export/import for STEP, STL, OBJ, IGES, BREP, DXF, SVG, and PDF files. |
| **Parts Library Integration** | FreeCAD Parts Library Addon | **Instant** | **Standardized** | Inspects and inserts standard mechanical components (fasteners, bearings, structural profiles). |

## 🛠️ Installation & Setup Guide

### 1. Prerequisites
- **FreeCAD**: Version 0.20, 0.21, 1.0, or 1.1+ ([Download from freecad.org](https://www.freecad.org/downloads.php)).
- **Python 3.10+** installed on your system.
- **Python Dependencies**:
  ```bash
  pip install -r requirements.txt
  ```
  *(or `pip install mcp requests`)*

### 2. FreeCAD Integration Methods

#### Method A: FreeCAD Addon Installation (Recommended) 🏆
Installing the Addon provides a dedicated **MCP Workbench**, GUI menu bar, and status bar controls inside FreeCAD.

1. Locate your FreeCAD `Mod` directory:
   - **Windows**: `%APPDATA%\FreeCAD\Mod\` (`C:\Users\<YourUser>\AppData\Roaming\FreeCAD\Mod\`)
   - **Linux**: `~/.local/share/FreeCAD/Mod/`
   - **macOS**: `~/Library/Application Support/FreeCAD/Mod/`

2. Copy the `addon/FreeCADMCP` folder into the `Mod` directory:
   - **PowerShell (Windows)**:
     ```powershell
     Copy-Item -Path ".\addon\FreeCADMCP" -Destination "$env:APPDATA\FreeCAD\Mod\FreeCADMCP" -Recurse -Force
     ```
   - **Bash (Linux/macOS)**:
     ```bash
     mkdir -p ~/.local/share/FreeCAD/Mod/
     cp -r addon/FreeCADMCP ~/.local/share/FreeCAD/Mod/
     ```

3. Launch FreeCAD. The **FreeCAD MCP** menu will appear in the top menu bar. Click **Start RPC Server** (or enable Auto-Start).

#### Method B: Standalone Macro Bridge (Zero-Install Option) ⚡
If you prefer not to install an Addon, you can run the bridge as a macro directly in FreeCAD:

1. Open FreeCAD.
2. Go to **Macro** -> **Macros...**.
3. Select or browse to `freecad_macro_bridge.py`.
4. Click **Execute**. The XML-RPC server (Port `9875`) and HTTP bridge (Port `9877`) will start immediately.

#### Method C: Headless & Environment Variables Setup 🖥️
For automated CI/CD pipelines or headless server environments without a GUI:

1. FreeCAD MCP will automatically detect `FreeCADCmd` in your system `PATH` and common installation directories.
2. If FreeCAD is installed in a custom location, configure any of the following environment variables:
   - `FREECAD_CMD`: Full path to `FreeCADCmd` executable (e.g. `C:\Program Files\FreeCAD 1.0\bin\FreeCADCmd.exe` or `/usr/bin/freecadcmd`).
   - `FREECAD_PATH`: Root directory of FreeCAD installation.
   - `FREECAD_HOST`: Host address for RPC server (default: `127.0.0.1`).
   - `FREECAD_RPC_PORT`: Port number for RPC server (default: `9875`).
   - `FREECAD_MCP_TOKEN`: Optional authentication token for secured remote connections.

## 💻 AI Client Configuration

### 1. Claude Desktop
Add to `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "freecad": {
      "command": "python",
      "args": [
        "/path/to/freecad-mcp/freecad_mcp_server.py"
      ],
      "env": {
        "FREECAD_HOST": "localhost",
        "FREECAD_RPC_PORT": "9875"
      }
    }
  }
}
```

### 2. Antigravity IDE
Add to your Antigravity MCP server configuration (`mcp_config.json`):
```json
{
  "mcpServers": {
    "freecad": {
      "command": "python",
      "args": [
        "/path/to/freecad-mcp/freecad_mcp_server.py"
      ],
      "env": {
        "FREECAD_HOST": "localhost",
        "FREECAD_RPC_PORT": "9875"
      }
    }
  }
}
```

### 3. Cursor / VS Code (Roo Code / Cline)
In Cursor MCP Settings or `.vscode/settings.json`:
- **Server Name**: `freecad`
- **Transport**: `stdio`
- **Command**: `python`
- **Args**: `["/path/to/freecad-mcp/freecad_mcp_server.py"]`
- **Environment**:
  - `FREECAD_HOST`: `localhost`
  - `FREECAD_RPC_PORT`: `9875`

## 🚀 Running the MCP Server Directly

You can start the MCP server manually from the terminal:

```bash
# Standard stdio server mode
python freecad_mcp_server.py

# Optional CLI flags:
# --only-text-feedback   : Suppress inline viewport screenshots to save LLM tokens
# --host <host>          : RPC host address (default: localhost)
# --port <port>          : RPC port number (default: 9875)
# --auth-token <token>   : Require authentication bearer token
# --freecadcmd <path>    : Explicit FreeCADCmd executable path for headless execution
python freecad_mcp_server.py --host 127.0.0.1 --port 9875
```

## 🛠️ Complete MCP Tools Catalog

### 1. Document & Session Management
- `create_document(name)`: Create a new FreeCAD `.FCStd` document.
- `list_documents()`: Return list of all currently open documents.
- `reload_document(doc_name)`: Reload document from disk to synchronize changes made by external or headless processes.
- `freecad_open_document(filepath)`: Open an existing `.FCStd` document file from disk.
- `freecad_save_document(filepath, document_name)`: Save an open document to disk as `.FCStd` (either active document or specified by name).
- `freecad_status()`: Comprehensive status check of XML-RPC, HTTP bridge, and headless `FreeCADCmd`.
- `get_rpc_status()`: Inspect RPC server health and GUI dispatch queue status.

### 2. Object Creation & Geometry Inspection
- `create_object(doc_name, obj_type, obj_name, analysis_name, obj_properties, include_screenshot, view_name)`: Create any registered FreeCAD C++ or Python object (`Part::Box`, `Part::Cylinder`, `PartDesign::Body`, `Fem::MaterialCommon`, etc.).
- `edit_object(doc_name, obj_name, obj_properties, include_screenshot, view_name)`: Update properties, dimensions, placement vectors, and visual appearance.
- `delete_object(doc_name, obj_name, include_screenshot, view_name)`: Remove an object from the active document tree.
- `get_objects(doc_name, include_screenshot, view_name)`: List all objects with bounding boxes, volumes, and placement data.
- `get_object(doc_name, obj_name, include_screenshot, view_name)`: Retrieve detailed parameters, geometric properties, and focused screenshot of a single object.
- `get_view(view_name, width, height, focus_object)`: Capture a high-resolution screenshot from any camera orientation (`Isometric`, `Front`, `Top`, `Right`, `Back`, `Left`, `Bottom`, `Dimetric`, `Trimetric`).
- `freecad_analyze_geometry(object_name, document_name)`: Calculate volume ($\text{mm}^3$), surface area ($\text{mm}^2$), center of mass, bounding box, face/edge/vertex counts, and solid validity.

### 3. Parametric CAD Modeling & Helpers
- `freecad_create_primitive(primitive_type, name, dimensions, position, rotation_axis_angle, document_name, include_screenshot, view_name)`: Create parametric 3D primitives (`box`, `cylinder`, `sphere`, `cone`, `torus`, `tube`, `wedge`, `plane`).
- `freecad_boolean_operation(operation, base_object, tool_object, result_name, document_name, include_screenshot, view_name)`: Perform solid boolean union (`fuse`), difference (`cut`), intersection (`common`), or section.
- `freecad_apply_fillet_chamfer(operation_type, object_name, radius_or_size, edge_indices, result_name, document_name, include_screenshot, view_name)`: Apply edge fillets or chamfers to solids.
- `freecad_create_sketch_and_extrude(plane, extrude_length, elements, sketch_name, body_name, document_name, include_screenshot, view_name)`: Create 2D sketches with circles, rectangles, and lines on XY/XZ/YZ planes, then extrude into a 3D solid Pad.
- `freecad_revolve_shape(shape_name, axis_vector, axis_base, angle_degrees, result_name, document_name, include_screenshot, view_name)`: Revolve a 2D profile around a 3D axis to create rotational solids.
- `insert_part_from_library(relative_path)`: Insert standard hardware from the FreeCAD parts library.
- `get_parts_list()`: List all standard components available in the parts library.

### 4. FEM Structural Analysis, TechDraw & File I/O
- `run_fem_analysis(doc_name, analysis_name, timeout, include_screenshot, view_name)`: Execute CalculiX solver on a FEM container, returning max von Mises stress (MPa), node displacements, and mesh summary.
- `freecad_create_techdraw(object_name, page_name, template_name, scale, projection_views, export_filepath, document_name)`: Generate multi-view 2D engineering projection drawings and export to SVG, DXF, or PDF.
- `freecad_export_file(export_format, filepath, object_names, document_name)`: Export 3D solids to STEP, STL, OBJ, IGES, BREP, DXF, SVG, or PDF.
- `freecad_import_file(filepath, document_name)`: Import external 3D CAD models into FreeCAD.

### 5. Code Execution & Background Computation
- `execute_code(code, include_screenshot, view_name, timeout)`: Execute Python code directly on FreeCAD's main GUI thread with stdout capture and visual feedback.
- `execute_code_async(code)`: Run heavy OpenCASCADE computations in a background thread; use `commit(fn)` for document writes.
- `get_async_status(job_id)`: Inspect async background jobs, execution states, and error traces.
- `execute_code_headless(code, timeout)`: Run Python scripts in an isolated `FreeCADCmd` subprocess with native crash detection.

## 🔒 Thread Safety & Best Practices

> [!WARNING]
> FreeCAD document trees and Coin3D scenegraphs are **not thread-safe**. When using `execute_code_async`, never modify document objects or scenegraph nodes directly in the worker thread. Always use the injected `commit(fn)` function:
>
> ```python
> # 1. Compute heavy OCCT geometry in background
> box = Part.makeBox(100, 50, 20)
> cyl = Part.makeCylinder(15, 25)
> result_shape = box.cut(cyl).removeSplitter()
> 
> # 2. Safely commit modifications to the GUI document
> def apply_changes():
>     obj = doc.addObject("Part::Feature", "CutResult")
>     obj.Shape = result_shape
>     doc.recompute()
> 
> commit(apply_changes)
> ```

> [!WARNING]
> CAD geometry modifications and FEM simulations can alter design documents. There is always a possibility of error, so we assume no responsibility for any inaccuracies.

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

