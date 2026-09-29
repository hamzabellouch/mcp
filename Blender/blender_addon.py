"""
Blender MCP Bridge Add-on / Script
----------------------------------
This script runs a lightweight, secure HTTP server inside Blender on port 9876.
It allows external MCP (Model Context Protocol) servers to communicate with Blender safely,
executing Python/bpy commands on the main thread, querying scene data, and rendering.

Usage:
1. In Blender, go to Scripting tab -> New -> Paste this code -> Click "Run Script" (Alt+P).
   OR install as an Add-on in Edit -> Preferences -> Add-ons.
2. In 3D Viewport, press 'N' to open the Sidebar and go to the "MCP Bridge" tab to Start/Stop the server.
"""

bl_info = {
    "name": "Blender MCP Bridge",
    "author": "Antigravity MCP Team",
    "version": (1, 0, 0),
    "blender": (3, 0, 0),
    "location": "View3D > Sidebar > MCP Bridge",
    "description": "Bridge server enabling AI models via MCP to interact with Blender safely in real time.",
    "category": "Development",
}

import bpy
import http.server
import socketserver
import threading
import json
import traceback
import io
import sys
import queue
import time
import os
import urllib.parse

# Server Configuration
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 9876

# Global state
server_thread = None
httpd = None
is_server_running = False
task_queue = queue.Queue()


class BridgeTask:
    def __init__(self, action, payload):
        self.action = action
        self.payload = payload
        self.done_event = threading.Event()
        self.result = None
        self.error = None


def process_queue_timer():
    """Timer running on Blender's main thread to process tasks safely."""
    global task_queue
    start_time = time.time()
    
    # Process pending tasks within a small time budget (e.g. 50ms) to keep UI responsive
    while not task_queue.empty():
        try:
            task = task_queue.get_nowait()
        except queue.Empty:
            break
            
        try:
            if task.action == "exec":
                task.result = execute_code_in_main_thread(task.payload.get("code", ""))
            elif task.action == "scene_info":
                task.result = get_scene_info_main_thread()
            elif task.action == "render":
                task.result = render_scene_main_thread(task.payload)
            elif task.action == "clear":
                task.result = clear_scene_main_thread(task.payload)
            else:
                task.error = f"Unknown action: {task.action}"
        except Exception as e:
            task.error = f"{type(e).__name__}: {str(e)}\n{traceback.format_exc()}"
        finally:
            task.done_event.set()
            
        if time.time() - start_time > 0.05:
            break
            
    # Keep timer running while server is active
    return 0.02 if is_server_running else None


def execute_code_in_main_thread(code_str):
    """Executes arbitrary python code within Blender context and captures output safely."""
    stdout_capture = io.StringIO()
    stderr_capture = io.StringIO()
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    
    exec_globals = {
        "bpy": bpy,
        "context": bpy.context,
        "data": bpy.data,
        "ops": bpy.ops,
        "math": __import__("math"),
        "os": os,
    }
    
    sys.stdout = stdout_capture
    sys.stderr = stderr_capture
    
    success = False
    error_msg = None
    
    try:
        compiled = compile(code_str, "<mcp_exec>", "exec")
        exec(compiled, exec_globals)
        success = True
    except Exception as e:
        error_msg = f"{type(e).__name__}: {str(e)}\n{traceback.format_exc()}"
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        
    return {
        "success": success,
        "stdout": stdout_capture.getvalue(),
        "stderr": stderr_capture.getvalue(),
        "error": error_msg
    }


def get_scene_info_main_thread():
    """Gathers comprehensive scene metadata."""
    scene = bpy.context.scene
    objects = []
    
    for obj in scene.objects:
        obj_data = {
            "name": obj.name,
            "type": obj.type,
            "location": [round(v, 4) for v in obj.location],
            "rotation_euler": [round(v, 4) for v in obj.rotation_euler],
            "scale": [round(v, 4) for v in obj.scale],
            "visible": obj.visible_get(),
            "materials": [slot.material.name for slot in obj.material_slots if slot.material] if hasattr(obj, "material_slots") else []
        }
        objects.append(obj_data)
        
    return {
        "scene_name": scene.name,
        "blender_version": bpy.app.version_string,
        "render_engine": scene.render.engine,
        "frame_current": scene.frame_current,
        "resolution": [scene.render.resolution_x, scene.render.resolution_y, scene.render.resolution_percentage],
        "active_camera": scene.camera.name if scene.camera else None,
        "active_object": bpy.context.active_object.name if bpy.context.active_object else None,
        "total_objects": len(objects),
        "objects": objects
    }


def render_scene_main_thread(payload):
    """Renders the current scene or frame to disk."""
    scene = bpy.context.scene
    output_path = payload.get("output_path", "")
    
    if not output_path:
        import tempfile
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, f"mcp_render_{int(time.time())}.png")
        
    res_x = payload.get("resolution_x")
    res_y = payload.get("resolution_y")
    if res_x:
        scene.render.resolution_x = int(res_x)
    if res_y:
        scene.render.resolution_y = int(res_y)
        
    original_filepath = scene.render.filepath
    scene.render.filepath = output_path
    
    try:
        bpy.ops.render.render(write_still=True)
        return {
            "success": True,
            "filepath": output_path,
            "resolution": [scene.render.resolution_x, scene.render.resolution_y],
            "engine": scene.render.engine
        }
    finally:
        scene.render.filepath = original_filepath


def clear_scene_main_thread(payload):
    """Cleans objects from the scene with options."""
    keep_camera = payload.get("keep_camera", True)
    keep_lights = payload.get("keep_lights", True)
    
    deleted_count = 0
    bpy.ops.object.select_all(action='DESELECT')
    
    for obj in list(bpy.context.scene.objects):
        if obj.type == 'CAMERA' and keep_camera:
            continue
        if obj.type == 'LIGHT' and keep_lights:
            continue
        obj.select_set(True)
        deleted_count += 1
        
    if deleted_count > 0:
        bpy.ops.object.delete()
        
    return {
        "success": True,
        "deleted_count": deleted_count
    }


class MCPBridgeHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Suppress noisy HTTP request logging in terminal
        pass

    def _is_authorized(self):
        """
        Validate incoming request security:
        1. Restrict connections exclusively to local loopback interface.
        2. Validate Host header to defend against DNS rebinding attacks.
        3. Block cross-origin requests originating from untrusted web browser origins.
        """
        client_ip = self.client_address[0] if self.client_address else ""
        if client_ip not in ("127.0.0.1", "::1", "localhost", "::ffff:127.0.0.1"):
            return False, 403, "Access denied: Only local connections are allowed."

        host_header = self.headers.get("Host", "").split(":")[0].strip().lower()
        if host_header not in ("localhost", "127.0.0.1", "::1", "[::1]"):
            return False, 400, "Access denied: Invalid Host header."

        origin = self.headers.get("Origin")
        if origin:
            try:
                parsed_origin = urllib.parse.urlparse(origin)
                if parsed_origin.hostname not in ("localhost", "127.0.0.1", "::1"):
                    return False, 403, "Access denied: Cross-origin requests from external web origins are prohibited."
            except Exception:
                return False, 400, "Malformed Origin header."

        return True, 200, ""

    def _send_json_response(self, status_code, data):
        try:
            response_bytes = json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(status_code)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(response_bytes)))
            self.end_headers()
            self.wfile.write(response_bytes)
        except Exception:
            pass

    def do_OPTIONS(self):
        auth_ok, status_code, err_msg = self._is_authorized()
        if not auth_ok:
            self._send_json_response(status_code, {"error": err_msg})
            return
        self._send_json_response(200, {"status": "ok"})

    def do_GET(self):
        auth_ok, status_code, err_msg = self._is_authorized()
        if not auth_ok:
            self._send_json_response(status_code, {"error": err_msg})
            return

        try:
            if self.path in ("/status", "/"):
                self._send_json_response(200, {
                    "status": "online",
                    "service": "Blender MCP Bridge",
                    "blender_version": bpy.app.version_string,
                    "version": "1.0.0"
                })
            elif self.path == "/scene_info":
                task = BridgeTask("scene_info", {})
                task_queue.put(task)
                if task.done_event.wait(timeout=10.0):
                    if task.error:
                        self._send_json_response(500, {"error": task.error})
                    else:
                        self._send_json_response(200, task.result)
                else:
                    self._send_json_response(504, {"error": "Blender main thread timed out"})
            else:
                self._send_json_response(404, {"error": "Not Found"})
        except Exception as e:
            self._send_json_response(500, {"error": f"Internal Server Error: {str(e)}"})

    def do_POST(self):
        auth_ok, status_code, err_msg = self._is_authorized()
        if not auth_ok:
            self._send_json_response(status_code, {"error": err_msg})
            return

        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b""
            
            payload = {}
            if post_data:
                try:
                    payload = json.loads(post_data.decode('utf-8'))
                except Exception as e:
                    self._send_json_response(400, {"error": f"Invalid JSON payload: {str(e)}"})
                    return

            action = None
            if self.path == "/execute":
                action = "exec"
            elif self.path == "/render":
                action = "render"
            elif self.path == "/clear":
                action = "clear"
                
            if action:
                task = BridgeTask(action, payload)
                task_queue.put(task)
                
                # Timeout of 60 seconds for long tasks like rendering
                timeout = 60.0 if action == "render" else 20.0
                if task.done_event.wait(timeout=timeout):
                    if task.error:
                        self._send_json_response(500, {"error": task.error})
                    else:
                        self._send_json_response(200, task.result)
                else:
                    self._send_json_response(504, {"error": "Operation timed out on Blender main thread"})
            else:
                self._send_json_response(404, {"error": "Endpoint not found"})
        except Exception as e:
            self._send_json_response(500, {"error": f"Internal Server Error: {str(e)}"})


class ReusableTCPServer(socketserver.TCPServer):
    allow_reuse_address = True


def start_server(host=DEFAULT_HOST, port=DEFAULT_PORT):
    global server_thread, httpd, is_server_running
    if is_server_running:
        return True, "Server is already running"
        
    try:
        httpd = ReusableTCPServer((host, port), MCPBridgeHandler)
        is_server_running = True
        server_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        server_thread.start()
        
        # Register main thread timer
        if not bpy.app.timers.is_registered(process_queue_timer):
            bpy.app.timers.register(process_queue_timer, persistent=True)
            
        print(f"[Blender MCP Bridge] Server started on http://{host}:{port}")
        return True, f"Server started on http://{host}:{port}"
    except Exception as e:
        is_server_running = False
        return False, f"Failed to start server: {str(e)}"


def stop_server():
    global server_thread, httpd, is_server_running
    if not is_server_running:
        return True, "Server is not running"
        
    is_server_running = False
    if httpd:
        httpd.shutdown()
        httpd.server_close()
        httpd = None
    if server_thread:
        server_thread.join(timeout=1.0)
        server_thread = None
        
    print("[Blender MCP Bridge] Server stopped")
    return True, "Server stopped"


# -------------------------------------------------------------
# Blender UI Panel and Operators
# -------------------------------------------------------------

class MCP_OT_StartServer(bpy.types.Operator):
    bl_idname = "mcp.start_server"
    bl_label = "Start MCP Bridge"
    bl_description = "Start the background server for AI MCP integration"
    
    def execute(self, context):
        scene = context.scene
        success, msg = start_server(scene.mcp_host, scene.mcp_port)
        if success:
            self.report({'INFO'}, msg)
        else:
            self.report({'ERROR'}, msg)
        return {'FINISHED'}


class MCP_OT_StopServer(bpy.types.Operator):
    bl_idname = "mcp.stop_server"
    bl_label = "Stop MCP Bridge"
    bl_description = "Stop the background MCP bridge server"
    
    def execute(self, context):
        success, msg = stop_server()
        self.report({'INFO'}, msg)
        return {'FINISHED'}


class MCP_PT_SidebarPanel(bpy.types.Panel):
    bl_label = "MCP AI Bridge"
    bl_idname = "MCP_PT_SidebarPanel"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = 'MCP Bridge'

    def draw(self, context):
        layout = self.layout
        scene = context.scene

        box = layout.box()
        box.label(text="Server Status:", icon='NETWORK_DRIVE')
        if is_server_running:
            row = box.row()
            row.label(text=f"Online (Port {scene.mcp_port})", icon='CHECKMARK')
            box.operator("mcp.stop_server", text="Stop Server", icon='CANCEL')
        else:
            row = box.row()
            row.label(text="Offline", icon='X')
            box.prop(scene, "mcp_port", text="Port")
            box.operator("mcp.start_server", text="Start Server", icon='PLAY')

        layout.separator()
        layout.label(text="Ready to receive AI commands via MCP.", icon='INFO')


classes = (
    MCP_OT_StartServer,
    MCP_OT_StopServer,
    MCP_PT_SidebarPanel,
)


def register():
    for cls in classes:
        bpy.utils.register_class(cls)
        
    bpy.types.Scene.mcp_host = bpy.props.StringProperty(
        name="Host",
        default=DEFAULT_HOST,
        description="Local host address"
    )
    bpy.types.Scene.mcp_port = bpy.props.IntProperty(
        name="Port",
        default=DEFAULT_PORT,
        min=1024,
        max=65535,
        description="Local port for the MCP Bridge server"
    )
    # Auto-start when script is run
    start_server(DEFAULT_HOST, DEFAULT_PORT)


def unregister():
    stop_server()
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
    del bpy.types.Scene.mcp_host
    del bpy.types.Scene.mcp_port


if __name__ == "__main__":
    register()
