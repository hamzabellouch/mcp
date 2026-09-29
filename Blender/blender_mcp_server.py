"""
Blender MCP Server
------------------
Model Context Protocol (MCP) server providing high-level tools to interact with Blender.
Communicates with the Blender MCP Bridge Addon running inside Blender on http://127.0.0.1:9876.
"""

import os
import sys
import json
import time
import tempfile
import urllib.parse
import pathlib
from typing import List, Optional, Dict, Any
import requests

try:
    from mcp.server.fastmcp import FastMCP
except (ImportError, ModuleNotFoundError):
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except ImportError:
        from mcp.server import MCPServer as FastMCP

# Initialize FastMCP Server
mcp = FastMCP("blender-mcp-server")

BLENDER_BRIDGE_URL = os.environ.get("BLENDER_BRIDGE_URL", "http://127.0.0.1:9876")


def _send_request(endpoint: str, method: str = "GET", payload: Optional[dict] = None, timeout: float = 20.0) -> dict:
    """Helper to send requests to Blender Bridge with error handling."""
    url = f"{BLENDER_BRIDGE_URL.rstrip('/')}/{endpoint.lstrip('/')}"
    headers = {"Host": "127.0.0.1:9876"}
    try:
        if method == "GET":
            response = requests.get(url, headers=headers, timeout=timeout)
        else:
            response = requests.post(url, headers=headers, json=payload or {}, timeout=timeout)
            
        if response.status_code == 200:
            return response.json()
        else:
            try:
                err_data = response.json()
                return {"error": err_data.get("error", response.text), "status_code": response.status_code}
            except Exception:
                return {"error": response.text, "status_code": response.status_code}
    except requests.exceptions.ConnectionError:
        return {
            "error": "Cannot connect to Blender. Please ensure Blender is running and the 'MCP Bridge' addon/script has been started (on port 9876).",
            "connected": False
        }
    except requests.exceptions.Timeout:
        return {"error": f"Request to Blender timed out after {timeout} seconds.", "timeout": True}
    except Exception as e:
        return {"error": f"Unexpected error: {str(e)}"}


@mcp.tool()
def blender_status() -> str:
    """
    Check if Blender is currently open and connected to the MCP Bridge.
    Returns the connection status, Blender version, and active scene info.
    """
    res = _send_request("/status", method="GET", timeout=5.0)
    if "error" in res and not res.get("status"):
        return f"❌ Disconnected: {res['error']}"
    return f"✅ Connected to Blender {res.get('blender_version', 'Unknown')}! Status: {res.get('status', 'online')}"


@mcp.tool()
def blender_get_scene_summary() -> str:
    """
    Get a complete summary of the active Blender scene, including all 3D objects,
    cameras, lights, materials, active object, and render configuration.
    """
    res = _send_request("/scene_info", method="GET", timeout=10.0)
    if "error" in res and not res.get("scene_name"):
        return f"Error fetching scene info: {res['error']}"
    
    return json.dumps(res, indent=2, ensure_ascii=False)


@mcp.tool()
def blender_execute_script(code: str) -> str:
    """
    Execute arbitrary Python / bpy code directly inside the running Blender instance.
    This allows full control over Blender's API (modeling, modifiers, geometry nodes, animation, rendering, etc.).
    
    Args:
        code: Python script string using bpy (e.g. 'import bpy\\nbpy.ops.mesh.primitive_cube_add()')
    """
    res = _send_request("/execute", method="POST", payload={"code": code}, timeout=30.0)
    if "error" in res and not res.get("success"):
        return f"❌ Execution Failed / Error:\n{res['error']}"
    
    output_parts = []
    if res.get("stdout"):
        output_parts.append(f"Output:\n{res['stdout']}")
    if res.get("stderr"):
        output_parts.append(f"Warnings/Stderr:\n{res['stderr']}")
    if not output_parts:
        output_parts.append("✅ Executed successfully with no console output.")
        
    return "\n\n".join(output_parts)


@mcp.tool()
def blender_create_primitive(
    primitive_type: str,
    name: str = "",
    location: Optional[List[float]] = None,
    rotation: Optional[List[float]] = None,
    scale: Optional[List[float]] = None,
    size: float = 2.0
) -> str:
    """
    Create a 3D primitive object in Blender.
    
    Args:
        primitive_type: Type of primitive ('cube', 'uv_sphere', 'ico_sphere', 'cylinder', 'cone', 'torus', 'plane', 'monkey', 'light', 'camera')
        name: Optional custom name for the object
        location: [X, Y, Z] world location (defaults to [0.0, 0.0, 0.0])
        rotation: [X, Y, Z] Euler rotation in radians (defaults to [0.0, 0.0, 0.0])
        scale: [X, Y, Z] scale multipliers (defaults to [1.0, 1.0, 1.0])
        size: Size/radius dimension for the primitive (defaults to 2.0)
    """
    loc = location if location is not None else [0.0, 0.0, 0.0]
    rot = rotation if rotation is not None else [0.0, 0.0, 0.0]
    scl = scale if scale is not None else [1.0, 1.0, 1.0]

    ptype = primitive_type.lower().strip()
    
    # Safe representations
    loc_repr = json.dumps(loc)
    rot_repr = json.dumps(rot)
    scl_repr = json.dumps(scl)
    name_repr = json.dumps(name)
    size_val = float(size)
    
    script = "import bpy, math\n"
    
    if ptype == "cube":
        script += f"bpy.ops.mesh.primitive_cube_add(size={size_val}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype in ("uv_sphere", "sphere"):
        script += f"bpy.ops.mesh.primitive_uv_sphere_add(radius={size_val/2.0}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype == "ico_sphere":
        script += f"bpy.ops.mesh.primitive_ico_sphere_add(radius={size_val/2.0}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype == "cylinder":
        script += f"bpy.ops.mesh.primitive_cylinder_add(radius={size_val/2.0}, depth={size_val}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype == "cone":
        script += f"bpy.ops.mesh.primitive_cone_add(radius1={size_val/2.0}, depth={size_val}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype == "torus":
        script += f"bpy.ops.mesh.primitive_torus_add(major_radius={size_val/2.0}, minor_radius={size_val/8.0}, location={loc_repr}, rotation={rot_repr})\n"
    elif ptype == "plane":
        script += f"bpy.ops.mesh.primitive_plane_add(size={size_val}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype in ("monkey", "suzanne"):
        script += f"bpy.ops.mesh.primitive_monkey_add(size={size_val}, location={loc_repr}, rotation={rot_repr}, scale={scl_repr})\n"
    elif ptype == "light":
        script += f"bpy.ops.object.light_add(type='POINT', location={loc_repr})\n"
    elif ptype == "camera":
        script += f"bpy.ops.object.camera_add(location={loc_repr}, rotation={rot_repr})\n"
    else:
        return f"❌ Unsupported primitive type: '{primitive_type}'. Supported types: cube, uv_sphere, ico_sphere, cylinder, cone, torus, plane, monkey, light, camera."

    if name:
        script += f"if bpy.context.active_object:\n    bpy.context.active_object.name = {name_repr}\n"
        
    script += f"print(f'Created {primitive_type} named ' + (bpy.context.active_object.name if bpy.context.active_object else 'unknown'))\n"
    
    return blender_execute_script(script)


@mcp.tool()
def blender_assign_material(
    object_name: str,
    material_name: str = "",
    base_color: Optional[List[float]] = None,
    metallic: float = 0.0,
    roughness: float = 0.4,
    emission_color: Optional[List[float]] = None,
    emission_strength: float = 0.0
) -> str:
    """
    Create a PBR material with Principled BSDF shader and assign it to an object.
    
    Args:
        object_name: Name of the target 3D object in Blender
        material_name: Name for the material (defaults to object_name + '_Mat')
        base_color: [R, G, B, A] color values between 0.0 and 1.0 (defaults to [0.8, 0.2, 0.2, 1.0])
        metallic: Metallic property between 0.0 (non-metal) and 1.0 (pure metal)
        roughness: Surface roughness between 0.0 (mirror smooth) and 1.0 (rough/matte)
        emission_color: [R, G, B, A] emission glow color (defaults to [0.0, 0.0, 0.0, 1.0])
        emission_strength: Emission light strength (defaults to 0.0)
    """
    col = base_color if base_color is not None else [0.8, 0.2, 0.2, 1.0]
    emit_col = emission_color if emission_color is not None else [0.0, 0.0, 0.0, 1.0]
    mat_name = material_name if material_name else f"{object_name}_Mat"

    obj_name_repr = json.dumps(object_name)
    mat_name_repr = json.dumps(mat_name)
    base_color_repr = json.dumps(col)
    metallic_repr = float(metallic)
    roughness_repr = float(roughness)
    emission_color_repr = json.dumps(emit_col)
    emission_strength_repr = float(emission_strength)

    script = f"""import bpy

obj_name = {obj_name_repr}
mat_name = {mat_name_repr}

obj = bpy.data.objects.get(obj_name)
if not obj:
    raise ValueError(f"Object '{{obj_name}}' not found in scene.")

# Create or get material
mat = bpy.data.materials.get(mat_name)
if not mat:
    mat = bpy.data.materials.new(name=mat_name)
    mat.use_nodes = True

nodes = mat.node_tree.nodes
principled = None
for node in nodes:
    if node.type == 'BSDF_PRINCIPLED':
        principled = node
        break

if principled:
    principled.inputs['Base Color'].default_value = {base_color_repr}
    principled.inputs['Metallic'].default_value = {metallic_repr}
    principled.inputs['Roughness'].default_value = {roughness_repr}
    
    # Handle emission across different Blender versions (3.x vs 4.x)
    if 'Emission Color' in principled.inputs:
        principled.inputs['Emission Color'].default_value = {emission_color_repr}
    elif 'Emission' in principled.inputs:
        principled.inputs['Emission'].default_value = {emission_color_repr}
        
    if 'Emission Strength' in principled.inputs:
        principled.inputs['Emission Strength'].default_value = {emission_strength_repr}

# Assign to object
if obj.data and hasattr(obj.data, 'materials'):
    if not obj.data.materials:
        obj.data.materials.append(mat)
    else:
        obj.data.materials[0] = mat

print(f"Assigned material '{{mat_name}}' to '{{obj_name}}' successfully.")
"""
    return blender_execute_script(script)


@mcp.tool()
def blender_clear_scene(keep_camera: bool = True, keep_lights: bool = True) -> str:
    """
    Clear objects and meshes from the active scene.
    
    Args:
        keep_camera: Whether to preserve cameras in the scene (default: True)
        keep_lights: Whether to preserve lights in the scene (default: True)
    """
    res = _send_request("/clear", method="POST", payload={
        "keep_camera": keep_camera,
        "keep_lights": keep_lights
    })
    if "error" in res and not res.get("success"):
        return f"Error clearing scene: {res['error']}"
    return f"🧹 Cleared scene successfully. Removed {res.get('deleted_count', 0)} objects."


@mcp.tool()
def blender_render_image(
    output_path: str = "",
    resolution_x: int = 1920,
    resolution_y: int = 1080
) -> str:
    """
    Render the current camera view of the Blender scene to an image file.
    
    Args:
        output_path: Optional absolute file path to save the rendered PNG image.
        resolution_x: Horizontal pixel resolution (default: 1920)
        resolution_y: Vertical pixel resolution (default: 1080)
    """
    res = _send_request("/render", method="POST", payload={
        "output_path": output_path,
        "resolution_x": resolution_x,
        "resolution_y": resolution_y
    }, timeout=60.0)
    
    if "error" in res and not res.get("success"):
        return f"Render failed: {res['error']}"
        
    return f"🖼️ Render completed successfully! Saved to: {res.get('filepath')}"


@mcp.tool()
def blender_import_3d_model(
    source_path_or_url: str,
    name: str = "",
    location: Optional[List[float]] = None,
    rotation: Optional[List[float]] = None,
    scale: Optional[List[float]] = None
) -> str:
    """
    Download and import a 3D model (.glb, .gltf, .obj, .fbx) directly into Blender.
    Supports both local file paths and remote URLs.
    
    Args:
        source_path_or_url: Local file path OR web URL to a 3D model (.glb, .gltf, .obj, .fbx)
        name: Custom name for the imported object in the scene
        location: [X, Y, Z] world position in meters (defaults to [0.0, 0.0, 0.0])
        rotation: [X, Y, Z] Euler rotation in radians (defaults to [0.0, 0.0, 0.0])
        scale: [X, Y, Z] scale multipliers (defaults to [1.0, 1.0, 1.0])
    """
    loc = location if location is not None else [0.0, 0.0, 0.0]
    rot = rotation if rotation is not None else [0.0, 0.0, 0.0]
    scl = scale if scale is not None else [1.0, 1.0, 1.0]

    local_path = source_path_or_url.strip()
    if local_path.startswith("http://") or local_path.startswith("https://"):
        try:
            parsed = urllib.parse.urlparse(local_path)
            raw_filename = os.path.basename(parsed.path) or "downloaded_model.glb"
            clean_filename = "".join(c for c in raw_filename if c.isalnum() or c in "._-")
            temp_dir = tempfile.gettempdir()
            target_dest = os.path.join(temp_dir, f"mcp_model_{int(time.time())}_{clean_filename}")
            
            resp = requests.get(local_path, stream=True, timeout=60.0)
            resp.raise_for_status()
            with open(target_dest, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            local_path = target_dest
        except Exception as e:
            return f"❌ Failed to download 3D model from URL: {str(e)}"

    local_path = os.path.abspath(local_path)
    if not os.path.exists(local_path):
        return f"❌ Model file not found at: {local_path}"

    ext = pathlib.Path(local_path).suffix.lower()
    filepath_repr = json.dumps(local_path)
    ext_repr = json.dumps(ext)
    name_repr = json.dumps(name)
    loc_repr = json.dumps(loc)
    rot_repr = json.dumps(rot)
    scl_repr = json.dumps(scl)

    script = f"""import bpy
import os

filepath = {filepath_repr}
ext = {ext_repr}

before_objs = set(bpy.context.scene.objects)

if ext in ('.glb', '.gltf'):
    bpy.ops.import_scene.gltf(filepath=filepath)
elif ext == '.obj':
    try:
        bpy.ops.wm.obj_import(filepath=filepath)
    except Exception:
        bpy.ops.import_scene.obj(filepath=filepath)
elif ext == '.fbx':
    bpy.ops.import_scene.fbx(filepath=filepath)
else:
    raise ValueError(f"Unsupported 3D file extension: {{ext}}")

after_objs = set(bpy.context.scene.objects)
new_objs = list(after_objs - before_objs)

if not new_objs:
    imported_obj = bpy.context.active_object
else:
    # Find root or primary mesh object
    imported_obj = new_objs[0]
    for o in new_objs:
        if not o.parent:
            imported_obj = o
            break

custom_name = {name_repr}
if imported_obj:
    imported_obj.location = {loc_repr}
    imported_obj.rotation_euler = {rot_repr}
    imported_obj.scale = {scl_repr}
    if custom_name:
        imported_obj.name = custom_name
    print(f"Successfully imported 3D model '{{imported_obj.name}}' at {{imported_obj.location}}")
else:
    print("Import finished, but no active object found.")
"""
    return blender_execute_script(script)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
