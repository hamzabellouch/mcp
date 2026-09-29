"""
FreeCAD MCP Live Bridge Macro & Launcher
-----------------------------------------
Execute this Macro inside FreeCAD (Macro -> Macros -> Select & Execute)
to enable instant AI control over FreeCAD via both:
1. XML-RPC Engine (Port 9875 - Full fast event loop with Qt Signals & async jobs)
2. HTTP REST Bridge (Port 9877 - Live viewport rendering & status inspection)

Works with FreeCAD 0.20, 0.21, 1.0, 1.1+ on Windows, Linux, and macOS.
"""

import sys
import os
import io
import json
import time
import queue
import threading
import traceback
from http.server import HTTPServer, BaseHTTPRequestHandler

# FreeCAD & GUI environment imports
try:
    import FreeCAD
    import FreeCADGui
    import Part
except ImportError:
    pass

# Ensure the addon directory is in sys.path
_current_dir = os.path.dirname(os.path.abspath(__file__))
_addon_dir = os.path.join(_current_dir, "addon", "FreeCADMCP")
if _addon_dir not in sys.path:
    sys.path.insert(0, _addon_dir)

# Try starting the official XML-RPC Addon server on port 9875
_rpc_started = False
try:
    from rpc_server import rpc_server
    msg = rpc_server.start_rpc_server(port=9875)
    FreeCAD.Console.PrintMessage(f"🚀 [FreeCAD MCP] {msg}\n")
    _rpc_started = True
except Exception as e:
    try:
        FreeCAD.Console.PrintWarning(f"⚠️ [FreeCAD MCP] Could not start XML-RPC server: {e}\n")
    except Exception:
        print(f"Warning: Could not start XML-RPC server: {e}")

# HTTP REST Bridge configuration (Port 9877)
HTTP_PORT = 9877
HTTP_HOST = "127.0.0.1"

task_queue = queue.Queue()
httpd_server = None
httpd_thread = None
is_running = False


class BridgeTask:
    def __init__(self, action: str, payload: dict):
        self.action = action
        self.payload = payload
        self.done_event = threading.Event()
        self.result = None
        self.error = None


def process_qt_queue():
    """Executed on FreeCAD main GUI thread via QTimer to process HTTP bridge tasks."""
    global task_queue
    start = time.time()
    while not task_queue.empty():
        try:
            task = task_queue.get_nowait()
        except queue.Empty:
            break

        try:
            if task.action == "execute":
                task.result = _main_thread_execute(task.payload.get("code", ""))
            elif task.action == "status":
                task.result = _main_thread_get_status()
            elif task.action == "document_tree":
                task.result = _main_thread_get_document_tree(task.payload.get("document_name"))
            elif task.action == "render_viewport":
                task.result = _main_thread_render_viewport(task.payload)
            else:
                task.error = f"Unknown action: {task.action}"
        except Exception as e:
            task.error = f"{type(e).__name__}: {str(e)}\n{traceback.format_exc()}"
        finally:
            task.done_event.set()

        if time.time() - start > 0.05:
            break


def _live_step_update(delay: float = 0.3):
    """Refreshes FreeCAD GUI viewport and processes Qt events for live viewing."""
    if FreeCAD.ActiveDocument:
        FreeCAD.ActiveDocument.recompute()
    if "FreeCADGui" in sys.modules and FreeCADGui.ActiveDocument:
        FreeCADGui.updateGui()
    
    try:
        if "PySide2" in sys.modules or "PySide" in sys.modules:
            from PySide2.QtWidgets import QApplication
        else:
            from PySide.QtGui import QApplication
        app = QApplication.instance()
        end_time = time.time() + delay
        while time.time() < end_time:
            if app:
                app.processEvents()
            time.sleep(0.02)
    except Exception:
        time.sleep(delay)


def _main_thread_get_status() -> dict:
    docs = []
    active_doc_name = FreeCAD.ActiveDocument.Name if FreeCAD.ActiveDocument else None

    for name, doc in FreeCAD.listDocuments().items():
        docs.append({
            "name": name,
            "label": doc.Label,
            "file_path": getattr(doc, "FileName", ""),
            "object_count": len(doc.Objects),
            "is_active": (name == active_doc_name)
        })

    return {
        "status": "online",
        "freecad_version": str(FreeCAD.Version()),
        "rpc_server_active": _rpc_started,
        "active_document": active_doc_name,
        "documents": docs
    }


def _main_thread_get_document_tree(doc_name: str = None) -> dict:
    doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
    if not doc:
        return {"error": "No active or specified FreeCAD document found."}

    objects_list = []
    for obj in doc.Objects:
        obj_info = {
            "name": obj.Name,
            "label": obj.Label,
            "type_id": obj.TypeId,
            "visibility": obj.ViewObject.Visibility if hasattr(obj, "ViewObject") and obj.ViewObject else True
        }

        if hasattr(obj, "Shape") and obj.Shape and not obj.Shape.isNull():
            s = obj.Shape
            bb = s.BoundBox
            obj_info["shape"] = {
                "shape_type": s.ShapeType,
                "volume": s.Volume,
                "area": s.Area,
                "length": s.Length,
                "bound_box": {
                    "xmin": bb.XMin, "xmax": bb.XMax,
                    "ymin": bb.YMin, "ymax": bb.YMax,
                    "zmin": bb.ZMin, "zmax": bb.ZMax,
                    "xlength": bb.XLength, "ylength": bb.YLength, "zlength": bb.ZLength
                },
                "num_faces": len(s.Faces),
                "num_edges": len(s.Edges),
                "num_vertexes": len(s.Vertexes)
            }

        if hasattr(obj, "Placement"):
            p = obj.Placement
            pos = p.Base
            rot = p.Rotation
            obj_info["placement"] = {
                "position": [pos.x, pos.y, pos.z],
                "rotation_quaternion": [rot.Q[0], rot.Q[1], rot.Q[2], rot.Q[3]],
                "yaw_pitch_roll": [rot.getYawPitchRoll()[0], rot.getYawPitchRoll()[1], rot.getYawPitchRoll()[2]]
            }

        objects_list.append(obj_info)

    return {
        "document_name": doc.Name,
        "document_label": doc.Label,
        "file_path": getattr(doc, "FileName", ""),
        "total_objects": len(objects_list),
        "objects": objects_list
    }


def _main_thread_execute(code_str: str) -> dict:
    stdout_buf = io.StringIO()
    stderr_buf = io.StringIO()
    old_stdout = sys.stdout
    old_stderr = sys.stderr

    exec_globals = {
        "FreeCAD": FreeCAD,
        "App": FreeCAD,
        "Gui": FreeCADGui if "FreeCADGui" in sys.modules else None,
        "FreeCADGui": FreeCADGui if "FreeCADGui" in sys.modules else None,
        "Part": sys.modules.get("Part"),
        "Sketcher": sys.modules.get("Sketcher"),
        "Draft": sys.modules.get("Draft"),
        "TechDraw": sys.modules.get("TechDraw"),
        "Mesh": sys.modules.get("Mesh"),
        "live_step": _live_step_update,
        "update_gui": _live_step_update,
        "time": time,
    }

    try:
        sys.stdout = stdout_buf
        sys.stderr = stderr_buf
        exec(code_str, exec_globals)
        
        if FreeCAD.ActiveDocument:
            FreeCAD.ActiveDocument.recompute()
            if "FreeCADGui" in sys.modules and FreeCADGui.ActiveDocument:
                FreeCADGui.updateGui()
                FreeCADGui.SendMsgToActiveView("ViewFit")

        return {
            "success": True,
            "stdout": stdout_buf.getvalue(),
            "stderr": stderr_buf.getvalue(),
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"{type(e).__name__}: {str(e)}",
            "traceback": traceback.format_exc(),
            "stdout": stdout_buf.getvalue(),
            "stderr": stderr_buf.getvalue(),
        }
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr


def _main_thread_render_viewport(payload: dict) -> dict:
    if "FreeCADGui" not in sys.modules or not FreeCADGui.ActiveDocument:
        return {"error": "FreeCAD 3D Viewport is not active or GUI is unavailable."}

    import tempfile
    filepath = payload.get("filepath")
    if not filepath:
        filepath = os.path.join(tempfile.gettempdir(), f"freecad_viewport_{int(time.time())}.png")

    width = payload.get("width", 1280)
    height = payload.get("height", 720)
    background = payload.get("background", "Current")

    view = FreeCADGui.ActiveDocument.ActiveView
    if not view:
        return {"error": "No active 3D view available."}

    view.saveImage(filepath, width, height, background)
    return {
        "success": True,
        "image_path": filepath,
        "resolution": f"{width}x{height}"
    }


class FreeCADBridgeHTTPHandler(BaseHTTPRequestHandler):
    def _set_headers(self, status=200, content_type="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_OPTIONS(self):
        self._set_headers(200)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/status":
            task = BridgeTask("status", {})
            task_queue.put(task)
            if task.done_event.wait(timeout=5.0):
                self._set_headers(200)
                self.wfile.write(json.dumps(task.result or {"error": task.error}).encode("utf-8"))
            else:
                self._set_headers(504)
                self.wfile.write(json.dumps({"error": "Status request timed out"}).encode("utf-8"))

        elif path == "/document_tree":
            task = BridgeTask("document_tree", {})
            task_queue.put(task)
            if task.done_event.wait(timeout=10.0):
                self._set_headers(200)
                self.wfile.write(json.dumps(task.result or {"error": task.error}).encode("utf-8"))
            else:
                self._set_headers(504)
                self.wfile.write(json.dumps({"error": "Document tree request timed out"}).encode("utf-8"))
        else:
            self._set_headers(404)
            self.wfile.write(b'{"error": "Endpoint not found"}')

    def do_POST(self):
        path = self.path.split("?")[0]
        content_length = int(self.headers.get("Content-Length", 0))
        post_data = self.rfile.read(content_length).decode("utf-8") if content_length > 0 else "{}"
        
        try:
            payload = json.loads(post_data)
        except Exception:
            payload = {}

        if path == "/execute":
            task = BridgeTask("execute", payload)
            task_queue.put(task)
            if task.done_event.wait(timeout=60.0):
                self._set_headers(200)
                self.wfile.write(json.dumps(task.result or {"error": task.error}).encode("utf-8"))
            else:
                self._set_headers(504)
                self.wfile.write(json.dumps({"error": "Execution timed out on GUI thread"}).encode("utf-8"))

        elif path == "/render_viewport":
            task = BridgeTask("render_viewport", payload)
            task_queue.put(task)
            if task.done_event.wait(timeout=15.0):
                self._set_headers(200)
                self.wfile.write(json.dumps(task.result or {"error": task.error}).encode("utf-8"))
            else:
                self._set_headers(504)
                self.wfile.write(json.dumps({"error": "Viewport render timed out"}).encode("utf-8"))
        else:
            self._set_headers(404)
            self.wfile.write(b'{"error": "Endpoint not found"}')

    def log_message(self, format, *args):
        pass


def start_bridge_macro():
    """Starts both the XML-RPC and HTTP bridges in FreeCAD."""
    global httpd_server, httpd_thread, is_running

    try:
        from PySide import QtCore
    except ImportError:
        try:
            from PySide2 import QtCore
        except ImportError:
            QtCore = None

    if QtCore:
        timer = QtCore.QTimer()
        timer.timeout.connect(process_qt_queue)
        timer.start(50)
        # Keep timer reference alive
        global _qt_timer_ref
        _qt_timer_ref = timer

    try:
        httpd_server = HTTPServer((HTTP_HOST, HTTP_PORT), FreeCADBridgeHTTPHandler)
        httpd_thread = threading.Thread(target=httpd_server.serve_forever, daemon=True)
        httpd_thread.start()
        is_running = True
        msg = f"🌟 FreeCAD MCP Bridge is ONLINE!\n- XML-RPC Server: port 9875\n- HTTP Bridge: http://{HTTP_HOST}:{HTTP_PORT}"
        FreeCAD.Console.PrintMessage(msg + "\n")
        try:
            FreeCADGui.getMainWindow().statusBar().showMessage("FreeCAD MCP Bridge Connected (Ports 9875 & 9877)", 10000)
        except Exception:
            pass
    except Exception as e:
        FreeCAD.Console.PrintError(f"❌ Failed to start HTTP Bridge on port {HTTP_PORT}: {e}\n")


# Automatically start when executed as a Macro in FreeCAD
if __name__ == "__main__" or "FreeCAD" in sys.modules:
    start_bridge_macro()
