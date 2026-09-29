"""
FreeCAD Model Context Protocol (MCP) Server 📐⚙️
------------------------------------------------
Comprehensive MCP server providing low-level and high-level parametric CAD modeling,
FEM simulation, Technical Drawings (TechDraw), live GUI screenshots, and automated scripting in FreeCAD.
"""

import os
import sys
import json
import logging
import argparse
import functools
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Literal, Optional, Union

try:
    from mcp.server.fastmcp import Context, FastMCP
except (ImportError, ModuleNotFoundError, Exception):
    try:
        from mcp.server.mcpserver import Context
        from mcp.server.mcpserver import MCPServer as FastMCP
    except ImportError:
        raise ImportError("The 'mcp' package is required. Install with: pip install 'mcp>=1.0.0'")

from mcp.types import ImageContent, TextContent

from freecad_client import FreeCADConnection
from freecad_engine import FreeCADEngine
from freecad_headless import run_headless, detect_freecadcmd, parse_command

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("FreeCADMCPServer")

ViewName = Literal[
    "Isometric", "Front", "Top", "Right", "Back", "Left", "Bottom", "Dimetric", "Trimetric"
]

ToolResponse = List[Union[TextContent, ImageContent]]


@dataclass
class ServerState:
    only_text_feedback: bool = False
    rpc_host: str = os.environ.get("FREECAD_HOST", "localhost")
    rpc_port: int = int(os.environ.get("FREECAD_RPC_PORT", "9875"))
    auth_token: Optional[str] = os.environ.get("FREECAD_MCP_TOKEN", None)
    freecad_connection: Optional[FreeCADConnection] = None
    freecadcmd: Optional[List[str]] = None
    version_notice: Optional[str] = None


state = ServerState()


def get_freecad_connection() -> FreeCADConnection:
    if state.freecad_connection is None:
        conn = FreeCADConnection(
            host=state.rpc_host,
            port=state.rpc_port,
            token=state.auth_token,
        )
        try:
            reachable = conn.ping()
        except Exception as e:
            conn.disconnect()
            raise Exception(f"Failed to connect to FreeCAD RPC ({e}). Make sure FreeCAD is running with the Addon or Macro bridge.") from e
        
        if not reachable:
            conn.disconnect()
            raise Exception("Failed to ping FreeCAD RPC. Ensure FreeCAD is active.")
        
        state.freecad_connection = conn
    return state.freecad_connection


def text_response(message: str) -> ToolResponse:
    return [TextContent(type="text", text=str(message))]


def json_response(data: Any) -> ToolResponse:
    return [TextContent(type="text", text=json.dumps(data, ensure_ascii=False, indent=2, default=str))]


def add_screenshot_if_available(
    response: ToolResponse,
    screenshot_b64: Optional[str],
    skip_screenshot: bool = False,
) -> ToolResponse:
    if skip_screenshot or state.only_text_feedback or not screenshot_b64:
        return response
    return [*response, ImageContent(type="image", data=screenshot_b64, mimeType="image/png")]


# Initialize FastMCP Server
mcp = FastMCP(
    "FreeCADMCP",
    instructions="Comprehensive FreeCAD Model Context Protocol (MCP) server for 3D parametric CAD modeling, TechDraw, FEM analysis, and automation."
)


# ==============================================================================
# SECTION 1: Standard Reference FreeCAD MCP Tools (Core RPC & Geometry Suite)
# ==============================================================================

@mcp.tool()
def create_document(name: str = "Unnamed") -> ToolResponse:
    """
    Create a new FreeCAD document.
    
    Args:
        name: Name of the document to create.
    """
    try:
        conn = get_freecad_connection()
        res = conn.create_document(name)
        if res.get("success"):
            return text_response(f"Document '{res.get('document_name', name)}' created successfully.")
        return text_response(f"Failed to create document: {res.get('error', 'Unknown error')}")
    except Exception:
        # Fallback to engine
        engine_res = FreeCADEngine.execute_code(f"import FreeCAD\ndoc = FreeCAD.newDocument('{name}')\nFreeCAD.setActiveDocument(doc.Name)")
        if engine_res.get("success"):
            return text_response(f"Document '{name}' created successfully.")
        return text_response(f"Error creating document: {engine_res.get('error')}")


@mcp.tool()
def list_documents() -> ToolResponse:
    """
    Get the list of all open documents in FreeCAD.
    """
    try:
        conn = get_freecad_connection()
        docs = conn.list_documents()
        return json_response(docs)
    except Exception:
        engine_res = FreeCADEngine.execute_code("import FreeCAD, json\nprint('__JSON__:' + json.dumps(list(FreeCAD.listDocuments().keys())))")
        stdout = engine_res.get("stdout", "")
        if "__JSON__:" in stdout:
            return text_response(stdout.split("__JSON__:")[1].strip())
        return text_response(f"Open documents: {stdout or 'None'}")


@mcp.tool()
def reload_document(doc_name: str) -> ToolResponse:
    """
    Close and re-open a document to pick up external file changes (such as from headless scripts).
    
    Args:
        doc_name: The name of the open document to reload.
    """
    try:
        conn = get_freecad_connection()
        res = conn.reload_document(doc_name)
        if res.get("success"):
            return text_response(f"Document '{doc_name}' reloaded from disk successfully.")
        return text_response(f"Failed to reload document: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error reloading document: {str(e)}")


@mcp.tool()
def create_object(
    doc_name: str,
    obj_type: str,
    obj_name: str,
    analysis_name: Optional[str] = None,
    obj_properties: Optional[Dict[str, Any]] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Create a new object in FreeCAD.
    
    Supports registered C++ types (e.g. 'Part::Box', 'Part::Cylinder', 'Part::Cut', 'PartDesign::Body', 'Fem::*')
    and Python-implemented types ('Part::Tube', 'Draft::Circle', 'Draft::Rectangle', 'Draft::Polygon', 'Draft::Wire').
    
    Args:
        doc_name: Name of the document to create the object in.
        obj_type: Type of object (e.g. 'Part::Box', 'Part::Cylinder', 'Draft::Rectangle', 'Fem::MaterialCommon').
        obj_name: Desired name for the object.
        analysis_name: Container name if creating a FEM sub-object.
        obj_properties: Dictionary of properties (Placement, Radius, Length, Height, Material, ShapeColor, etc.).
        include_screenshot: Whether to return a 3D viewport screenshot after creation (default True).
        view_name: View orientation for screenshot ('Isometric', 'Front', 'Top', 'Right', etc.).
    """
    try:
        conn = get_freecad_connection()
        obj_data = {
            "Name": obj_name,
            "Type": obj_type,
            "Analysis": analysis_name,
            "Properties": obj_properties or {},
        }
        res = conn.create_object(doc_name, obj_data)
        if res.get("success"):
            actual_name = res.get("object_name", obj_name)
            resp = text_response(f"Object '{actual_name}' ({obj_type}) created successfully in document '{doc_name}'.")
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return text_response(f"Failed to create object: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error creating object: {str(e)}")


@mcp.tool()
def edit_object(
    doc_name: str,
    obj_name: str,
    obj_properties: Dict[str, Any],
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Edit properties of an existing object in FreeCAD.
    
    Args:
        doc_name: Name of the document.
        obj_name: Name of the object to edit.
        obj_properties: Properties to update (e.g. {'Length': 50, 'Placement': {'Base': {'x': 10, 'y': 0, 'z': 0}}}).
        include_screenshot: Whether to return a screenshot (default True).
        view_name: Screenshot view orientation.
    """
    try:
        conn = get_freecad_connection()
        res = conn.edit_object(doc_name, obj_name, {"Properties": obj_properties})
        if res.get("success"):
            resp = text_response(f"Object '{obj_name}' updated successfully.")
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return text_response(f"Failed to edit object: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error editing object: {str(e)}")


@mcp.tool()
def delete_object(
    doc_name: str,
    obj_name: str,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Delete an object from a FreeCAD document.
    
    Args:
        doc_name: Name of the document.
        obj_name: Name of the object to delete.
        include_screenshot: Whether to return screenshot.
        view_name: Screenshot view orientation.
    """
    try:
        conn = get_freecad_connection()
        res = conn.delete_object(doc_name, obj_name)
        if res.get("success"):
            resp = text_response(f"Object '{obj_name}' deleted successfully from '{doc_name}'.")
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return text_response(f"Failed to delete object: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error deleting object: {str(e)}")


@mcp.tool()
def get_objects(
    doc_name: str,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Get all objects and their metadata from a FreeCAD document.
    
    Args:
        doc_name: Name of the document.
        include_screenshot: Whether to return viewport screenshot.
        view_name: View orientation for screenshot.
    """
    try:
        conn = get_freecad_connection()
        objs = conn.get_objects(doc_name)
        resp = json_response(objs)
        screenshot = None
        if include_screenshot and not state.only_text_feedback:
            screenshot = conn.get_active_screenshot(view_name)
        return add_screenshot_if_available(resp, screenshot, not include_screenshot)
    except Exception as e:
        return text_response(f"Error getting objects: {str(e)}")


@mcp.tool()
def get_object(
    doc_name: str,
    obj_name: str,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Get detailed properties and geometry information of a single object.
    
    Args:
        doc_name: Name of the document.
        obj_name: Name of the object.
        include_screenshot: Whether to return screenshot.
        view_name: Screenshot view orientation.
    """
    try:
        conn = get_freecad_connection()
        obj = conn.get_object(doc_name, obj_name)
        if obj is None:
            return text_response(f"Object '{obj_name}' not found in document '{doc_name}'.")
        resp = json_response(obj)
        screenshot = None
        if include_screenshot and not state.only_text_feedback:
            screenshot = conn.get_active_screenshot(view_name, focus_object=obj_name)
        return add_screenshot_if_available(resp, screenshot, not include_screenshot)
    except Exception as e:
        return text_response(f"Error getting object: {str(e)}")


@mcp.tool()
def get_view(
    view_name: ViewName = "Isometric",
    width: Optional[int] = None,
    height: Optional[int] = None,
    focus_object: Optional[str] = None,
) -> ToolResponse:
    """
    Get a high-resolution screenshot of the active 3D view.
    
    Args:
        view_name: Orientation ('Isometric', 'Front', 'Top', 'Right', 'Back', 'Left', 'Bottom', 'Dimetric', 'Trimetric').
        width: Optional image width in pixels (e.g. 1280).
        height: Optional image height in pixels (e.g. 720).
        focus_object: Optional object name to center/fit the camera on.
    """
    try:
        conn = get_freecad_connection()
        b64 = conn.get_active_screenshot(view_name, width, height, focus_object)
        if b64:
            return [ImageContent(type="image", data=b64, mimeType="image/png")]
        return text_response("Cannot capture screenshot in the current view or no active document.")
    except Exception as e:
        return text_response(f"Error capturing view: {str(e)}")


@mcp.tool()
def execute_code(
    code: str,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
    timeout: Optional[float] = None,
) -> ToolResponse:
    """
    Execute arbitrary Python code directly on FreeCAD's GUI main thread.
    
    Args:
        code: Python script to execute in FreeCAD.
        include_screenshot: Whether to return screenshot after execution (default True).
        view_name: View orientation for screenshot.
        timeout: Optional timeout in seconds (default 90).
    """
    try:
        conn = get_freecad_connection()
        res = conn.execute_code(code, timeout=timeout)
        if res.get("success"):
            resp = text_response(res.get("message", "Python code executed successfully."))
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return text_response(f"Execution failed: {res.get('error')}")
    except Exception:
        # Fallback to engine
        engine_res = FreeCADEngine.execute_code(code, timeout=timeout or 90.0)
        if engine_res.get("success"):
            return text_response(f"Code executed successfully via {engine_res.get('mode')}:\n{engine_res.get('stdout', '')}")
        return text_response(f"Code execution error: {engine_res.get('error')}")


@mcp.tool()
def execute_code_async(code: str) -> ToolResponse:
    """
    Execute Python code in a background thread without blocking the FreeCAD GUI.
    
    Use for heavy OpenCASCADE (OCCT) computations (fusing hundreds of solids, lofts, heavy meshes).
    Use the injected `commit(fn)` helper to safely write the computed shapes back to the GUI document.
    
    Args:
        code: Background-safe Python code.
    """
    try:
        conn = get_freecad_connection()
        res = conn.execute_code_async(code)
        if res.get("success"):
            job_id = res.get("job_id", "")
            return text_response(
                f"✅ Async background computation started (Job ID: '{job_id}').\n"
                f"Use `get_async_status(job_id='{job_id}')` to check completion status and tracebacks."
            )
        return text_response(f"Failed to start async execution: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error starting async execution: {str(e)}")


@mcp.tool()
def get_async_status(job_id: str = "") -> ToolResponse:
    """
    Report the status of background jobs started by `execute_code_async`.
    Does not touch the GUI thread, so it responds instantly even while computation is running.
    
    Args:
        job_id: The ID of the background job. If empty, lists all running and recently completed jobs.
    """
    try:
        conn = get_freecad_connection()
        res = conn.get_async_status(job_id)
        if not res.get("success"):
            return text_response(f"Failed to get async status: {res.get('error')}")
        job = res.get("job")
        if job is not None:
            text = f"Async Job '{job.get('id')}': {job.get('state', 'unknown')}"
            if job.get("error"):
                text += f"\nError: {job['error']}"
            if job.get("traceback"):
                text += f"\nTraceback:\n{job['traceback']}"
            return text_response(text)
        return json_response(res.get("jobs", []))
    except Exception as e:
        return text_response(f"Error getting async status: {str(e)}")


@mcp.tool()
def execute_code_headless(code: str, timeout: float = 600.0) -> ToolResponse:
    """
    Run a FreeCAD Python script in a separate headless `freecadcmd` process with crash isolation.
    If OpenCASCADE segfaults on complex geometry, only the helper process dies while the GUI session survives intact.
    
    Args:
        code: Complete Python script.
        timeout: Seconds before terminating process (default 600).
    """
    res = run_headless(code, timeout=timeout, command=state.freecadcmd)
    if res.get("success"):
        out = f"✅ Headless FreeCAD script completed successfully (Exit 0)."
        if res.get("output"):
            out += f"\nOutput:\n{res['output']}"
        out += "\nTip: If the script saved a document that is open in FreeCAD GUI, call `reload_document(doc_name)` to show results."
        return text_response(out)
    else:
        out = f"❌ Headless FreeCAD script FAILED: {res.get('error')}"
        if res.get("output"):
            out += f"\nOutput:\n{res['output']}"
        return text_response(out)


@mcp.tool()
def get_rpc_status() -> ToolResponse:
    """
    Report RPC server and GUI dispatch health without using the GUI thread.
    """
    try:
        conn = get_freecad_connection()
        return json_response(conn.get_rpc_status())
    except Exception as e:
        return text_response(f"RPC status unavailable: {str(e)}")


@mcp.tool()
def insert_part_from_library(
    relative_path: str,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Insert a standard part from the FreeCAD parts library.
    
    Args:
        relative_path: Relative path of the part in the library (e.g. 'Mechanical/Fasteners/ISO4762_M6x20.step').
        include_screenshot: Whether to return a screenshot.
        view_name: Screenshot view orientation.
    """
    try:
        conn = get_freecad_connection()
        res = conn.insert_part_from_library(relative_path)
        if res.get("success"):
            resp = text_response(f"Part '{relative_path}' inserted from library.")
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return text_response(f"Failed to insert part: {res.get('error')}")
    except Exception as e:
        return text_response(f"Error inserting part from library: {str(e)}")


@mcp.tool()
def get_parts_list() -> ToolResponse:
    """
    Get the list of available mechanical parts in the FreeCAD parts library.
    """
    try:
        conn = get_freecad_connection()
        parts = conn.get_parts_list()
        if parts:
            return json_response(parts)
        return text_response("No parts found. Ensure the FreeCAD Parts Library addon is installed.")
    except Exception as e:
        return text_response(f"Error retrieving parts list: {str(e)}")


@mcp.tool()
def run_fem_analysis(
    doc_name: str,
    analysis_name: str,
    timeout: int = 600,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric",
) -> ToolResponse:
    """
    Run CalculiX FEM solver on an existing Fem::FemAnalysis container and return stress and displacement results.
    
    Args:
        doc_name: Name of the FreeCAD document.
        analysis_name: Name of the Fem::AnalysisPython object.
        timeout: Max seconds to wait for CalculiX solver (default 600).
        include_screenshot: Whether to return visual screenshot.
        view_name: Screenshot view orientation.
    """
    try:
        conn = get_freecad_connection()
        res = conn.run_fem_analysis(doc_name, analysis_name, timeout=timeout)
        if res.get("success"):
            resp = json_response({
                "summary": (
                    f"FEM analysis '{analysis_name}' solved successfully! "
                    f"Max von Mises = {res.get('max_von_mises_MPa', 'N/A')} MPa, "
                    f"Max displacement = {res.get('max_displacement_mm', 'N/A')} mm "
                    f"({res.get('node_count', 0)} nodes)."
                ),
                **res
            })
            screenshot = None
            if include_screenshot and not state.only_text_feedback:
                screenshot = conn.get_active_screenshot(view_name)
            return add_screenshot_if_available(resp, screenshot, not include_screenshot)
        return json_response({"summary": f"FEM analysis failed: {res.get('error')}", **res})
    except Exception as e:
        return text_response(f"Error running FEM analysis: {str(e)}")


# ==============================================================================
# SECTION 2: High-Level Parametric CAD & Automation Convenience Tools
# ==============================================================================

@mcp.tool()
def freecad_status() -> str:
    """
    Check the comprehensive status of FreeCAD, RPC server, HTTP bridge, and headless engine.
    """
    status = FreeCADEngine.get_status()
    lines = ["📐 FreeCAD Status Report:"]
    if status.get("rpc_connected"):
        details = status.get("rpc_details", {})
        lines.append(f"✅ XML-RPC Server: CONNECTED (Port {FreeCADEngine.RPC_PORT})")
        lines.append(f"   - Addon Version: {details.get('addon_version', 'Unknown')}")
        lines.append(f"   - GUI Dispatch: {details.get('gui_dispatch', {}).get('status', 'Active')}")
    else:
        lines.append(f"❌ XML-RPC Server: Not connected on port {FreeCADEngine.RPC_PORT}")

    if status.get("http_bridge_connected"):
        lines.append(f"✅ HTTP Macro Bridge: CONNECTED (Port {FreeCADEngine.HTTP_PORT})")
    else:
        lines.append(f"❌ HTTP Macro Bridge: Disconnected")

    if status.get("headless_available"):
        lines.append(f"✅ Headless FreeCADCmd: FOUND ({status.get('freecadcmd_path')})")
    else:
        lines.append(f"⚠️ Headless FreeCADCmd: Not detected in PATH or standard install directories.")

    return "\n".join(lines)


@mcp.tool()
def freecad_create_primitive(
    primitive_type: str,
    name: Optional[str] = None,
    dimensions: Optional[Dict[str, float]] = None,
    position: Optional[List[float]] = None,
    rotation_axis_angle: Optional[List[float]] = None,
    document_name: Optional[str] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric"
) -> ToolResponse:
    """
    Create a parametric 3D primitive (box, cylinder, sphere, cone, torus, tube, wedge, plane).
    
    Args:
        primitive_type: 'box', 'cylinder', 'sphere', 'cone', 'torus', 'tube', 'wedge', or 'plane'.
        name: Custom label/name for the object.
        dimensions: Dimension dictionary (e.g. {'length': 100, 'width': 50, 'height': 20} or {'radius': 25, 'height': 50}).
        position: [x, y, z] placement coordinates in mm.
        rotation_axis_angle: [axis_x, axis_y, axis_z, angle_degrees] (e.g. [0, 0, 1, 45]).
        document_name: Target document name (creates new if omitted).
        include_screenshot: Return screenshot after creation (default True).
        view_name: View orientation for screenshot.
    """
    ptype = primitive_type.lower().strip()
    dims = dimensions or {}
    pos = position or [0.0, 0.0, 0.0]
    rot = rotation_axis_angle or [0.0, 0.0, 1.0, 0.0]
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    
    code = f"""
import FreeCAD, Part
doc = {doc_lookup}
if not doc:
    doc = FreeCAD.newDocument("CAD_Model")
    FreeCAD.setActiveDocument(doc.Name)

p_type = "{ptype}"
obj_name = "{name or ''}"

if p_type == "box":
    obj = doc.addObject("Part::Box", obj_name or "Box")
    obj.Length = {dims.get('length', 10.0)}
    obj.Width = {dims.get('width', 10.0)}
    obj.Height = {dims.get('height', 10.0)}

elif p_type == "cylinder":
    obj = doc.addObject("Part::Cylinder", obj_name or "Cylinder")
    obj.Radius = {dims.get('radius', 5.0)}
    obj.Height = {dims.get('height', 10.0)}
    obj.Angle = {dims.get('angle', 360.0)}

elif p_type == "sphere":
    obj = doc.addObject("Part::Sphere", obj_name or "Sphere")
    obj.Radius = {dims.get('radius', 5.0)}

elif p_type == "cone":
    obj = doc.addObject("Part::Cone", obj_name or "Cone")
    obj.Radius1 = {dims.get('radius1', 10.0)}
    obj.Radius2 = {dims.get('radius2', 0.0)}
    obj.Height = {dims.get('height', 20.0)}

elif p_type == "torus":
    obj = doc.addObject("Part::Torus", obj_name or "Torus")
    obj.Radius1 = {dims.get('radius1', 20.0)}
    obj.Radius2 = {dims.get('radius2', 5.0)}

elif p_type == "tube":
    outer = Part.makeCylinder({dims.get('outer_radius', 10.0)}, {dims.get('height', 20.0)})
    inner = Part.makeCylinder({dims.get('inner_radius', 7.0)}, {dims.get('height', 20.0)})
    obj = doc.addObject("Part::Feature", obj_name or "Tube")
    obj.Shape = outer.cut(inner)

elif p_type == "wedge":
    obj = doc.addObject("Part::Wedge", obj_name or "Wedge")
    obj.Xmin = {dims.get('xmin', 0.0)}
    obj.Ymin = {dims.get('ymin', 0.0)}
    obj.Zmin = {dims.get('zmin', 0.0)}
    obj.Xmax = {dims.get('xmax', 20.0)}
    obj.Ymax = {dims.get('ymax', 20.0)}
    obj.Zmax = {dims.get('zmax', 20.0)}

elif p_type == "plane":
    obj = doc.addObject("Part::Plane", obj_name or "Plane")
    obj.Length = {dims.get('length', 100.0)}
    obj.Width = {dims.get('width', 100.0)}

else:
    raise ValueError(f"Unknown primitive type: {{p_type}}")

obj.Placement.Base = FreeCAD.Vector({pos[0]}, {pos[1]}, {pos[2]})
obj.Placement.Rotation = FreeCAD.Rotation(FreeCAD.Vector({rot[0]}, {rot[1]}, {rot[2]}), {rot[3]})
doc.recompute()
print(f"Created {{obj.Name}} ({{p_type}}) at {{obj.Placement.Base}}")
"""
    return execute_code(code, include_screenshot=include_screenshot, view_name=view_name)


@mcp.tool()
def freecad_boolean_operation(
    operation: str,
    base_object: str,
    tool_object: str,
    result_name: Optional[str] = None,
    document_name: Optional[str] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric"
) -> ToolResponse:
    """
    Perform a Boolean operation (fuse, cut, common, section) between two objects.
    
    Args:
        operation: 'fuse' (Union), 'cut' (Subtract), 'common' (Intersect), or 'section'.
        base_object: Name of primary object.
        tool_object: Name of tool object.
        result_name: Optional name for resulting boolean object.
        document_name: Optional document name.
        include_screenshot: Return screenshot after operation.
        view_name: Screenshot view orientation.
    """
    op = operation.lower().strip()
    type_map = {
        "fuse": "Part::Fuse",
        "union": "Part::Fuse",
        "cut": "Part::Cut",
        "difference": "Part::Cut",
        "subtract": "Part::Cut",
        "common": "Part::Common",
        "intersect": "Part::Common",
        "intersection": "Part::Common",
        "section": "Part::Section",
    }
    obj_type = type_map.get(op)
    if not obj_type:
        return text_response(f"❌ Unknown boolean operation '{operation}'. Use 'fuse', 'cut', 'common', or 'section'.")

    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    code = f"""
import FreeCAD
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

base = doc.getObject("{base_object}")
tool = doc.getObject("{tool_object}")
if not base:
    raise ValueError("Base object '{base_object}' not found.")
if not tool:
    raise ValueError("Tool object '{tool_object}' not found.")

bool_obj = doc.addObject("{obj_type}", "{result_name or op.capitalize()}")
bool_obj.Base = base
bool_obj.Tool = tool
doc.recompute()

if hasattr(base, 'ViewObject') and base.ViewObject:
    base.ViewObject.Visibility = False
if hasattr(tool, 'ViewObject') and tool.ViewObject:
    tool.ViewObject.Visibility = False

doc.recompute()
print(f"Boolean operation '{op}' completed: created {{bool_obj.Name}}")
"""
    return execute_code(code, include_screenshot=include_screenshot, view_name=view_name)


@mcp.tool()
def freecad_apply_fillet_chamfer(
    operation_type: str,
    object_name: str,
    radius_or_size: float,
    edge_indices: Optional[List[int]] = None,
    result_name: Optional[str] = None,
    document_name: Optional[str] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric"
) -> ToolResponse:
    """
    Apply rounded fillets or beveled chamfers to 3D solid edges.
    
    Args:
        operation_type: 'fillet' or 'chamfer'.
        object_name: Name of solid object.
        radius_or_size: Fillet radius or chamfer size in mm.
        edge_indices: List of 1-based edge indices (e.g. [1, 2, 3]). If omitted, applies to all edges.
        result_name: Optional name for new object.
        document_name: Optional document name.
        include_screenshot: Return screenshot after operation.
        view_name: View orientation for screenshot.
    """
    op = operation_type.lower().strip()
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    edges_json = json.dumps(edge_indices or [])
    
    code = f"""
import FreeCAD, Part
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

target = doc.getObject("{object_name}")
if not target:
    raise ValueError("Target object '{object_name}' not found.")

edges_in = {edges_json}
if not edges_in:
    edges_in = list(range(1, len(target.Shape.Edges) + 1))

if "{op}" == "fillet":
    obj = doc.addObject("Part::Fillet", "{result_name or 'Fillet'}")
    obj.Base = target
    obj.Edges = [(idx, {radius_or_size}, {radius_or_size}) for idx in edges_in]
elif "{op}" == "chamfer":
    obj = doc.addObject("Part::Chamfer", "{result_name or 'Chamfer'}")
    obj.Base = target
    obj.Edges = [(idx, {radius_or_size}, {radius_or_size}) for idx in edges_in]
else:
    raise ValueError(f"Unknown operation type: {op}")

doc.recompute()
if hasattr(target, 'ViewObject') and target.ViewObject:
    target.ViewObject.Visibility = False
doc.recompute()
print(f"{op.capitalize()} applied to {{len(edges_in)}} edges on {object_name}.")
"""
    return execute_code(code, include_screenshot=include_screenshot, view_name=view_name)


@mcp.tool()
def freecad_create_sketch_and_extrude(
    plane: str = "XY",
    extrude_length: float = 10.0,
    elements: Optional[List[Dict[str, Any]]] = None,
    sketch_name: Optional[str] = None,
    body_name: Optional[str] = None,
    document_name: Optional[str] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric"
) -> ToolResponse:
    """
    Create a 2D parametric Sketch on a plane (XY, XZ, YZ) and extrude (Pad) into a 3D solid.
    
    Args:
        plane: 'XY', 'XZ', or 'YZ'.
        extrude_length: Extrusion length in mm (default 10).
        elements: List of geometric elements in sketch:
            - Circle: {'type': 'circle', 'center': [x, y], 'radius': r}
            - Rectangle: {'type': 'rectangle', 'corner1': [x1, y1], 'corner2': [x2, y2]}
            - Line: {'type': 'line', 'start': [x1, y1], 'end': [x2, y2]}
        sketch_name: Optional sketch label.
        body_name: Optional PartDesign body name.
        document_name: Optional document name.
        include_screenshot: Return screenshot after extrusion.
        view_name: View orientation for screenshot.
    """
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    elements_json = json.dumps(elements or [
        {"type": "rectangle", "corner1": [-20, -10], "corner2": [20, 10]}
    ])

    code = f"""
import FreeCAD, Sketcher, Part, PartDesign
doc = {doc_lookup}
if not doc:
    doc = FreeCAD.newDocument("CAD_Model")
    FreeCAD.setActiveDocument(doc.Name)

body = doc.addObject("PartDesign::Body", "{body_name or 'Body'}")
sketch = doc.addObject("Sketcher::SketchObject", "{sketch_name or 'Sketch'}")
body.addObject(sketch)

pln = "{plane.upper().strip()}"
if pln == "XY":
    sketch.Support = [(doc.addObject("PartDesign::Plane", "XY_Plane"), '')]
    sketch.MapMode = "FlatFace"
elif pln == "XZ":
    sketch.Placement = FreeCAD.Placement(FreeCAD.Vector(0,0,0), FreeCAD.Rotation(FreeCAD.Vector(1,0,0), 90))
elif pln == "YZ":
    sketch.Placement = FreeCAD.Placement(FreeCAD.Vector(0,0,0), FreeCAD.Rotation(FreeCAD.Vector(0,1,0), -90))

elems = {elements_json}
for el in elems:
    t = el.get("type", "").lower()
    if t == "circle":
        c = el.get("center", [0, 0])
        r = el.get("radius", 10.0)
        sketch.addGeometry(Part.Circle(FreeCAD.Vector(c[0], c[1], 0), FreeCAD.Vector(0, 0, 1), r), False)
    elif t == "rectangle":
        c1 = el.get("corner1", [0, 0])
        c2 = el.get("corner2", [10, 10])
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(c1[0], c1[1], 0), FreeCAD.Vector(c2[0], c1[1], 0)), False)
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(c2[0], c1[1], 0), FreeCAD.Vector(c2[0], c2[1], 0)), False)
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(c2[0], c2[1], 0), FreeCAD.Vector(c1[0], c2[1], 0)), False)
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(c1[0], c2[1], 0), FreeCAD.Vector(c1[0], c1[1], 0)), False)
    elif t == "line":
        s = el.get("start", [0, 0])
        e = el.get("end", [10, 10])
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(s[0], s[1], 0), FreeCAD.Vector(e[0], e[1], 0)), False)

doc.recompute()
pad = doc.addObject("PartDesign::Pad", "Pad")
pad.Profile = sketch
pad.Length = {extrude_length}
body.addObject(pad)
doc.recompute()
print(f"Created Sketch on {{pln}} and Extruded Pad with length {extrude_length}mm.")
"""
    return execute_code(code, include_screenshot=include_screenshot, view_name=view_name)


@mcp.tool()
def freecad_revolve_shape(
    shape_name: str,
    axis_vector: Optional[List[float]] = None,
    axis_base: Optional[List[float]] = None,
    angle_degrees: float = 360.0,
    result_name: Optional[str] = None,
    document_name: Optional[str] = None,
    include_screenshot: bool = True,
    view_name: ViewName = "Isometric"
) -> ToolResponse:
    """
    Revolve a 2D wire/shape/sketch around an axis to create a 3D rotational solid.
    
    Args:
        shape_name: Name of source sketch or shape.
        axis_vector: [x, y, z] direction vector of rotation axis (default [0, 0, 1] Z-axis).
        axis_base: [x, y, z] origin coordinates for rotation axis (default [0, 0, 0]).
        angle_degrees: Rotation angle in degrees (default 360.0).
        result_name: Optional name for new solid.
        document_name: Optional document name.
        include_screenshot: Return screenshot after revolution.
        view_name: View orientation for screenshot.
    """
    ax = axis_vector or [0.0, 0.0, 1.0]
    base = axis_base or [0.0, 0.0, 0.0]
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    
    code = f"""
import FreeCAD, Part
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

src = doc.getObject("{shape_name}")
if not src:
    raise ValueError("Source object '{shape_name}' not found.")

rev = doc.addObject("Part::Revolution", "{result_name or 'Revolution'}")
rev.Source = src
rev.Axis = FreeCAD.Vector({ax[0]}, {ax[1]}, {ax[2]})
rev.Base = FreeCAD.Vector({base[0]}, {base[1]}, {base[2]})
rev.Angle = {angle_degrees}
doc.recompute()
if hasattr(src, 'ViewObject') and src.ViewObject:
    src.ViewObject.Visibility = False
doc.recompute()
print(f"Revolved {{src.Name}} by {angle_degrees} degrees around axis {{rev.Axis}}.")
"""
    return execute_code(code, include_screenshot=include_screenshot, view_name=view_name)


@mcp.tool()
def freecad_create_techdraw(
    object_name: str,
    page_name: Optional[str] = None,
    template_name: str = "A4_LandscapeTD.svg",
    scale: float = 1.0,
    projection_views: Optional[List[str]] = None,
    export_filepath: Optional[str] = None,
    document_name: Optional[str] = None
) -> str:
    """
    Generate a 2D technical engineering drawing sheet (TechDraw) with multi-view projections.
    
    Args:
        object_name: 3D model object to project.
        page_name: Optional name for the TechDraw page.
        template_name: SVG Template filename (e.g. 'A4_LandscapeTD.svg').
        scale: Scale factor (1.0 = 1:1, 0.5 = 1:2).
        projection_views: Views to project (e.g. ['Front', 'Top', 'Right', 'Isometric']).
        export_filepath: Optional path to export drawing as SVG, DXF, or PDF.
        document_name: Optional document name.
    """
    views_list = projection_views or ["Front", "Top", "Right", "Isometric"]
    views_json = json.dumps(views_list)
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    export_stmt = ""
    if export_filepath:
        norm_exp = export_filepath.replace("\\", "/")
        export_stmt = f"""
import TechDrawGui
page.exportPage("{norm_exp}")
print(f"TechDraw exported to: {norm_exp}")
"""

    code = f"""
import FreeCAD, TechDraw
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

target = doc.getObject("{object_name}")
if not target:
    raise ValueError("Target object '{object_name}' not found.")

page = doc.addObject("TechDraw::DrawPage", "{page_name or 'TechDraw_Sheet'}")
template = doc.addObject("TechDraw::DrawSVGTemplate", "Template")
template.Template = TechDraw.findTemplate("{template_name}")
page.Template = template
doc.recompute()

views = {views_json}
positions = {{
    "Front": (100, 100),
    "Top": (100, 200),
    "Right": (200, 100),
    "Isometric": (200, 200),
    "Left": (50, 100),
    "Bottom": (100, 50),
    "Rear": (200, 50)
}}

for v in views:
    dv = doc.addObject("TechDraw::DrawViewPart", f"View_{{v}}")
    dv.Source = [target]
    dv.Scale = {scale}
    pos = positions.get(v, (100, 100))
    dv.X = pos[0]
    dv.Y = pos[1]
    
    if v == "Front":
        dv.Direction = FreeCAD.Vector(0, -1, 0)
    elif v == "Top":
        dv.Direction = FreeCAD.Vector(0, 0, 1)
    elif v == "Right":
        dv.Direction = FreeCAD.Vector(1, 0, 0)
    elif v == "Isometric":
        dv.Direction = FreeCAD.Vector(1, -1, 1)
    page.addView(dv)

doc.recompute()
print(f"TechDraw sheet created with views: {{', '.join(views)}}")
{export_stmt}
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Failed to create TechDraw: {res.get('error')}"
    return f"✅ Technical Drawing created successfully!\n{res.get('stdout', '')}"


@mcp.tool()
def freecad_analyze_geometry(
    object_name: str,
    document_name: Optional[str] = None
) -> str:
    """
    Calculate physical and geometric mass properties: volume, surface area, center of mass, bounding box.
    
    Args:
        object_name: Name of solid object to inspect.
        document_name: Optional document name.
    """
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    code = f"""
import FreeCAD, json
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

obj = doc.getObject("{object_name}")
if not obj or not hasattr(obj, "Shape") or obj.Shape.isNull():
    raise ValueError("Object '{object_name}' not found or has no solid shape.")

s = obj.Shape
bb = s.BoundBox
com = s.CenterOfMass
data = {{
    "object_name": obj.Name,
    "label": obj.Label,
    "volume_mm3": round(s.Volume, 4),
    "surface_area_mm2": round(s.Area, 4),
    "center_of_mass": [round(com.x, 4), round(com.y, 4), round(com.z, 4)],
    "bounding_box_mm": {{
        "length_x": round(bb.XLength, 4),
        "width_y": round(bb.YLength, 4),
        "height_z": round(bb.ZLength, 4),
        "min": [round(bb.XMin, 4), round(bb.YMin, 4), round(bb.ZMin, 4)],
        "max": [round(bb.XMax, 4), round(bb.YMax, 4), round(bb.ZMax, 4)]
    }},
    "faces_count": len(s.Faces),
    "edges_count": len(s.Edges),
    "vertex_count": len(s.Vertexes),
    "is_valid": s.isValid(),
    "is_closed_solid": s.isClosed()
}}
print("__JSON_ANALYSIS__:" + json.dumps(data))
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Geometry analysis failed: {res.get('error')}"
    stdout = res.get("stdout", "")
    if "__JSON_ANALYSIS__:" in stdout:
        return stdout.split("__JSON_ANALYSIS__:")[1].strip()
    return stdout


@mcp.tool()
def freecad_export_file(
    export_format: str,
    filepath: str,
    object_names: Optional[List[str]] = None,
    document_name: Optional[str] = None
) -> str:
    """
    Export FreeCAD 3D objects to STEP, STL, OBJ, IGES, BREP, DXF, SVG, or PDF.
    
    Args:
        export_format: 'step', 'stp', 'stl', 'obj', 'iges', 'brep', 'dxf', 'svg', or 'pdf'.
        filepath: Destination file path (e.g. 'E:/exports/bracket.step').
        object_names: List of object names to export. Defaults to all objects.
        document_name: Optional document name.
    """
    fmt = export_format.lower().replace(".", "").strip()
    norm_path = filepath.replace("\\", "/")
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    objs_json = json.dumps(object_names or [])

    code = f"""
import FreeCAD, Part, Mesh
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found.")

req_objs = {objs_json}
if req_objs:
    export_list = [doc.getObject(name) for name in req_objs if doc.getObject(name)]
else:
    export_list = [o for o in doc.Objects if hasattr(o, 'Shape') and o.Shape and not o.Shape.isNull()]

if not export_list:
    raise RuntimeError("No valid solid objects found to export.")

fmt = "{fmt}"
dest = "{norm_path}"

if fmt in ["step", "stp"]:
    import Import
    Import.export(export_list, dest)
elif fmt == "stl":
    import Mesh
    mesh_objs = []
    for o in export_list:
        m = doc.addObject("Mesh::Feature", f"Mesh_{{o.Name}}")
        m.Mesh = Mesh.Mesh(o.Shape.tessellate(0.1))
        mesh_objs.append(m)
    Mesh.export(mesh_objs, dest)
    for m in mesh_objs:
        doc.removeObject(m.Name)
elif fmt == "obj":
    import ImportGui
    ImportGui.export(export_list, dest)
elif fmt in ["iges", "igs"]:
    import Import
    Import.export(export_list, dest)
elif fmt == "brep":
    export_list[0].Shape.exportBrep(dest)
else:
    import Import
    Import.export(export_list, dest)

print(f"Exported {{len(export_list)}} objects to {{dest}} in {{fmt.upper()}} format.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Export failed: {res.get('error')}"
    return f"✅ File exported successfully!\n{res.get('stdout', '')}"


@mcp.tool()
def freecad_import_file(
    filepath: str,
    document_name: Optional[str] = None
) -> str:
    """
    Import an external 3D CAD or mesh file (.step, .stl, .obj, .iges, .dxf) into FreeCAD.
    
    Args:
        filepath: Full path to the CAD file.
        document_name: Optional target document name.
    """
    norm_path = filepath.replace("\\", "/")
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    
    code = f"""
import FreeCAD, Import
doc = {doc_lookup}
if not doc:
    doc = FreeCAD.newDocument("Imported_CAD")
    FreeCAD.setActiveDocument(doc.Name)

Import.insert("{norm_path}", doc.Name)
doc.recompute()
print(f"Imported '{norm_path}' into document '{{doc.Name}}'. Total objects: {{len(doc.Objects)}}")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Import failed: {res.get('error')}"
    return f"✅ File imported successfully!\n{res.get('stdout', '')}"


@mcp.tool()
def freecad_open_document(filepath: str) -> str:
    """
    Open an existing FreeCAD document file (.FCStd).
    
    Args:
        filepath: Full path to the .FCStd document file.
    """
    norm_path = filepath.replace("\\", "/")
    code = f"""
import FreeCAD
doc = FreeCAD.openDocument("{norm_path}")
if doc:
    FreeCAD.setActiveDocument(doc.Name)
    print(f"Opened document '{{doc.Name}}' from '{norm_path}'. Objects count: {{len(doc.Objects)}}")
else:
    raise RuntimeError("Failed to open document: 'FreeCAD.openDocument' returned None.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Failed to open document: {res.get('error')}"
    return f"✅ Document opened successfully!\n{res.get('stdout', '')}"


@mcp.tool()
def freecad_save_document(
    filepath: Optional[str] = None,
    document_name: Optional[str] = None
) -> str:
    """
    Save an open FreeCAD document (.FCStd) to disk.
    
    Args:
        filepath: Optional destination file path to save as (e.g. 'C:/models/part.FCStd'). If omitted, saves to the document's current file path.
        document_name: Optional name of the document to save. If omitted, saves the active document.
    """
    doc_lookup = f'FreeCAD.getDocument("{document_name}")' if document_name else "FreeCAD.ActiveDocument"
    if filepath:
        norm_path = filepath.replace("\\", "/")
        save_call = f'''doc.saveAs("{norm_path}")
print(f"Document '{{doc.Name}}' saved as '{norm_path}'.")'''
    else:
        save_call = f'''if not getattr(doc, "FileName", None):
    raise RuntimeError("Document has no existing file path on disk. Please specify a 'filepath' parameter to save.")
doc.save()
print(f"Document '{{doc.Name}}' saved to '{{doc.FileName}}'.")'''

    code = f"""
import FreeCAD
doc = {doc_lookup}
if not doc:
    raise RuntimeError("No active document found to save.")
{save_call}
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return f"❌ Failed to save document: {res.get('error')}"
    return f"✅ Document saved successfully!\n{res.get('stdout', '')}"


# ==============================================================================
# SECTION 3: Strategy Prompts
# ==============================================================================

@mcp.prompt()
def asset_creation_strategy() -> str:
    """Standard asset creation strategy for AI CAD engineers."""
    return """
Asset Creation Strategy for FreeCAD MCP:

When creating 3D models in FreeCAD:
1. Always check existing document objects using `get_objects(doc_name)` or `freecad_status()`.
2. For standard hardware (bolts, nuts, bearings), check `get_parts_list()` and insert using `insert_part_from_library()`.
3. For custom geometry:
   - Create basic primitives (`create_object` or `freecad_create_primitive`).
   - For complex profiles, use 2D Sketches + Extrusions (`freecad_create_sketch_and_extrude`) or Revolutions (`freecad_revolve_shape`).
   - Combine solids using Boolean operations (`freecad_boolean_operation` - fuse/cut/common).
   - Add fillets and chamfers (`freecad_apply_fillet_chamfer`).
4. Keep screenshots enabled (`include_screenshot=True`) for visual verification, or set `include_screenshot=False` during intermediate batch steps.
5. For heavy computation or complex boolean chains, use `execute_code_async` with `commit()` or `execute_code_headless` to prevent GUI lockups.
"""


def main():
    parser = argparse.ArgumentParser(description="FreeCAD MCP Server")
    parser.add_argument("--only-text-feedback", action="store_true", help="Only return text feedback (suppress inline screenshots)")
    parser.add_argument("--host", default=os.environ.get("FREECAD_HOST", "localhost"), help="FreeCAD RPC Host (default: localhost)")
    parser.add_argument("--port", type=int, default=int(os.environ.get("FREECAD_RPC_PORT", "9875")), help="FreeCAD RPC Port (default: 9875)")
    parser.add_argument("--auth-token", default=os.environ.get("FREECAD_MCP_TOKEN", None), help="FreeCAD RPC Auth Token")
    parser.add_argument("--freecadcmd", default=None, help="Path or command for headless FreeCADCmd")
    args, _ = parser.parse_known_args()

    state.only_text_feedback = args.only_text_feedback
    state.rpc_host = args.host
    state.rpc_port = args.port
    state.auth_token = args.auth_token
    state.freecadcmd = parse_command(args.freecadcmd)

    logger.info(f"FreeCAD MCP Server initialized. Target RPC: {state.rpc_host}:{state.rpc_port}")
    mcp.run()


if __name__ == "__main__":
    main()
