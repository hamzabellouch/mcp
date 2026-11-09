"""
FreeCAD Model Context Protocol (MCP) Server 📐⚙️
------------------------------------------------
Comprehensive MCP server providing low-level and high-level parametric CAD modeling,
FEM simulation, Technical Drawings (TechDraw), live GUI screenshots, and automated scripting in FreeCAD.
"""

import os
import sys
import math
import re
import shlex
import shutil
import signal
import subprocess
import tempfile
import json
import logging
import argparse
import functools
import xmlrpc.client
from pathlib import Path
from dataclasses import dataclass
from http.server import HTTPServer, BaseHTTPRequestHandler
from typing import Any, Callable, Dict, List, Literal, Optional, Tuple, Union

import requests

try:
    from mcp.server.fastmcp import Context, FastMCP
except (ImportError, ModuleNotFoundError, Exception):
    try:
        from mcp.server.mcpserver import Context
        from mcp.server.mcpserver import MCPServer as FastMCP
    except ImportError:
        raise ImportError("The 'mcp' package is required. Install with: pip install 'mcp>=1.0.0'")

from mcp.types import ImageContent, TextContent

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("FreeCADMCPServer")


# ==============================================================================
# 1. XML-RPC Client Layer
# ==============================================================================

class _TimeoutTransport(xmlrpc.client.Transport):
    """XML-RPC transport with configurable socket timeout and headers."""
    def __init__(self, timeout: float = 30, headers: Optional[List[Tuple[str, str]]] = None, **kwargs):
        super().__init__(**kwargs)
        self._timeout = timeout
        self._custom_headers = headers or []

    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = self._timeout
        return conn

    def send_headers(self, connection, headers):
        for header, value in self._custom_headers:
            connection.putheader(header, value)
        super().send_headers(connection, headers)


def _is_budget(value: Any, ceiling: float) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and 0 < value <= ceiling


class FreeCADConnection:
    """Client connection to the live FreeCAD RPC Server."""
    
    EXECUTE_CODE_TIMEOUT = 90
    MAX_EXECUTE_CODE_TIMEOUT = 1800
    RPC_TIMEOUT_MARGIN = 30
    VERSION_CHECK_TIMEOUT = 5

    def __init__(
        self,
        host: str = "localhost",
        port: int = 9875,
        timeout: float = 150,
        token: Optional[str] = None,
    ):
        self._host = host
        self._port = port
        self._uri = f"http://{host}:{port}"
        self._token = token
        self._headers = [("Authorization", f"Bearer {token}")] if token else []
        self._timeout = timeout
        self.server = self._make_proxy(timeout)

    def _make_proxy(self, timeout: float) -> xmlrpc.client.ServerProxy:
        return xmlrpc.client.ServerProxy(
            self._uri,
            allow_none=True,
            transport=_TimeoutTransport(timeout=timeout, headers=self._headers),
        )

    def reconnect(self) -> bool:
        """Close existing connection and re-initialize proxy."""
        self.disconnect()
        self.server = self._make_proxy(self._timeout)
        return self.ping()

    def disconnect(self) -> None:
        transport = getattr(self.server, "_ServerProxy__transport", None)
        close = getattr(transport, "close", None)
        if callable(close):
            close()

    def ping(self) -> bool:
        try:
            with self._make_proxy(self.VERSION_CHECK_TIMEOUT) as proxy:
                return bool(proxy.ping())
        except Exception:
            return False

    def get_rpc_status(self) -> Dict[str, Any]:
        with self._make_proxy(self.VERSION_CHECK_TIMEOUT) as proxy:
            return proxy.get_rpc_status()

    def check_addon_version(self) -> Optional[str]:
        try:
            status = self.get_rpc_status()
            if isinstance(status, dict):
                for key, attr in (
                    ("execute_code_timeout", "EXECUTE_CODE_TIMEOUT"),
                    ("max_execute_code_timeout", "MAX_EXECUTE_CODE_TIMEOUT"),
                ):
                    val = status.get(key)
                    if _is_budget(val, self.MAX_EXECUTE_CODE_TIMEOUT):
                        setattr(self, attr, val)
                return None
        except Exception as e:
            logger.debug(f"Version check note: {e}")
        return None

    def _call(self, func_name: str, *args, **kwargs):
        try:
            return getattr(self.server, func_name)(*args, **kwargs)
        except (ConnectionRefusedError, OSError, TimeoutError, xmlrpc.client.ProtocolError, Exception):
            if self.reconnect():
                return getattr(self.server, func_name)(*args, **kwargs)
            raise

    def create_document(self, name: str = "New_Document") -> Dict[str, Any]:
        return self._call("create_document", name)

    def list_documents(self) -> List[str]:
        return self._call("list_documents")

    def reload_document(self, doc_name: str) -> Dict[str, Any]:
        return self._call("reload_document", doc_name)

    def create_object(self, doc_name: str, obj_data: Dict[str, Any]) -> Dict[str, Any]:
        return self._call("create_object", doc_name, obj_data)

    def edit_object(self, doc_name: str, obj_name: str, obj_data: Dict[str, Any]) -> Dict[str, Any]:
        return self._call("edit_object", doc_name, obj_name, obj_data)

    def delete_object(self, doc_name: str, obj_name: str) -> Dict[str, Any]:
        return self._call("delete_object", doc_name, obj_name)

    def get_objects(self, doc_name: str) -> List[Dict[str, Any]]:
        return self._call("get_objects", doc_name)

    def get_object(self, doc_name: str, obj_name: str) -> Optional[Dict[str, Any]]:
        return self._call("get_object", doc_name, obj_name)

    def execute_code(self, code: str, timeout: Optional[float] = None) -> Dict[str, Any]:
        run_budget = self.EXECUTE_CODE_TIMEOUT
        if timeout is not None:
            try:
                run_budget = float(timeout)
            except (TypeError, ValueError, OverflowError) as exc:
                raise ValueError("timeout must be a positive finite number") from exc
            if isinstance(timeout, bool) or not math.isfinite(run_budget) or run_budget <= 0:
                raise ValueError("timeout must be a positive finite number")
            run_budget = min(run_budget, self.MAX_EXECUTE_CODE_TIMEOUT)

        socket_timeout = max(self._timeout, 2 * run_budget + self.RPC_TIMEOUT_MARGIN)
        with self._make_proxy(socket_timeout) as proxy:
            if timeout is None:
                return proxy.execute_code(code)
            return proxy.execute_code(code, run_budget)

    def execute_code_async(self, code: str) -> Dict[str, Any]:
        return self._call("execute_code_async", code)

    def get_async_status(self, job_id: str = "") -> Dict[str, Any]:
        with self._make_proxy(self._timeout) as proxy:
            return proxy.get_async_status(job_id)

    def get_active_screenshot(
        self,
        view_name: str = "Isometric",
        width: Optional[int] = None,
        height: Optional[int] = None,
        focus_object: Optional[str] = None,
    ) -> Optional[str]:
        return self._call("get_active_screenshot", view_name, width, height, focus_object)

    def run_fem_analysis(self, doc_name: str, analysis_name: str, timeout: int = 600) -> Dict[str, Any]:
        return self._call("run_fem_analysis", doc_name, analysis_name, timeout)

    def insert_part_from_library(self, relative_path: str) -> Dict[str, Any]:
        return self._call("insert_part_from_library", relative_path)

    def get_parts_list(self) -> Any:
        return self._call("get_parts_list")


# ==============================================================================
# 2. Headless FreeCADCmd Runner Layer
# ==============================================================================

FLATPAK_APP = "org.freecad.FreeCAD"
_NOISE = ("%)", "Importing project files", "Postprocessing", "FreeCAD 1.", "(C) 2001", "LGPL")

COMMON_WIN_PATHS = [
    os.environ.get("FREECAD_CMD", ""),
    os.environ.get("FREECAD_PATH", ""),
    os.environ.get("FREECAD_BIN", ""),
    r"C:\Program Files\FreeCAD",
    r"C:\Program Files\FreeCAD 1.1",
    r"C:\Program Files\FreeCAD 1.0",
    r"C:\Program Files\FreeCAD 0.22",
    r"C:\Program Files\FreeCAD 0.21",
    r"C:\Program Files\FreeCAD 0.20",
    r"C:\Program Files (x86)\FreeCAD",
    r"C:\Program Files (x86)\FreeCAD 1.1",
    r"C:\Program Files (x86)\FreeCAD 1.0",
    r"C:\Program Files (x86)\FreeCAD 0.21",
    r"D:\Program Files\FreeCAD",
    r"D:\Program Files\FreeCAD 1.1",
    r"D:\Program Files\FreeCAD 1.0",
    r"D:\Program Files\FreeCAD 0.21",
    r"E:\Program Files\FreeCAD",
    r"E:\Program Files\FreeCAD 1.1",
    r"E:\Program Files\FreeCAD 1.0",
    r"E:\Program Files\FreeCAD 0.21",
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 1.1"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 1.0"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD"),
    os.path.expanduser(r"~\AppData\Local\Programs\FreeCAD 0.21"),
]


def detect_freecadcmd() -> Optional[List[str]]:
    """Locate a headless FreeCAD executable: PATH first, known Windows directories, then Flatpak."""
    env_cmd = os.environ.get("FREECAD_CMD")
    if env_cmd and os.path.exists(env_cmd):
        return [env_cmd]

    for name in ("freecadcmd", "FreeCADCmd", "freecadcmd.exe", "FreeCADCmd.exe", "freecad.cmd", "FreeCAD.exe", "freecad.exe"):
        path = shutil.which(name)
        if path:
            return [path]

    for base in COMMON_WIN_PATHS:
        if not base or not os.path.exists(base):
            continue
        if os.path.isfile(base) and base.lower().endswith(".exe"):
            return [base]
        
        candidates = [
            os.path.join(base, "bin", "FreeCADCmd.exe"),
            os.path.join(base, "bin", "freecadcmd.exe"),
            os.path.join(base, "bin", "FreeCAD.exe"),
            os.path.join(base, "bin", "freecad.exe"),
            os.path.join(base, "FreeCADCmd.exe"),
            os.path.join(base, "FreeCAD.exe"),
        ]
        for c in candidates:
            if os.path.isfile(c):
                return [c]

    flatpak = shutil.which("flatpak")
    if flatpak:
        try:
            probe = subprocess.run(
                [flatpak, "info", FLATPAK_APP], capture_output=True, timeout=10
            )
            if probe.returncode == 0:
                return [flatpak, "run", "--command=freecadcmd", FLATPAK_APP]
        except (OSError, subprocess.TimeoutExpired):
            pass

    return None


def parse_command(value: Optional[str]) -> Optional[List[str]]:
    return shlex.split(value) if value else None


def _script_dir() -> Path:
    d = Path.home() / ".cache" / "freecad-mcp" / "headless"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _clean(output: Optional[Any]) -> str:
    if output is None:
        return ""
    if isinstance(output, bytes):
        output = output.decode("utf-8", errors="replace")
    lines = [ln for ln in str(output).splitlines() if ln.strip() and not any(n in ln for n in _NOISE)]
    return "\n".join(lines)


def run_headless(code: str, timeout: float = 600.0, command: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Execute python code headlessly with FreeCADCmd in an isolated process.
    Returns success, returncode, output, crashed, and timed_out flags.
    """
    if not math.isfinite(timeout) or timeout <= 0:
        return {"success": False, "error": "timeout must be a positive finite number"}

    cmd = command or detect_freecadcmd()
    if not cmd:
        return {
            "success": False,
            "error": "FreeCADCmd not found: install FreeCAD (https://www.freecad.org/) or set FREECAD_CMD / FREECAD_PATH",
        }

    script_path = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".py", dir=_script_dir(), delete=False, encoding="utf-8") as f:
            f.write(code)
            script_path = f.name

        argv = cmd + ["-c", f"exec(open({script_path!r}, encoding='utf-8').read())"]
        proc = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )

        output = _clean(proc.stdout + "\n" + proc.stderr)
        result: Dict[str, Any] = {
            "success": proc.returncode == 0,
            "returncode": proc.returncode,
            "output": output,
        }

        if proc.returncode < 0:
            sig = -proc.returncode
            try:
                name = signal.Signals(sig).name
            except ValueError:
                name = str(sig)
            result["crashed"] = True
            result["error"] = f"Headless FreeCAD crashed with signal {name} (OCCT native crash; GUI unaffected)"
        elif proc.returncode != 0:
            crash = re.search(r"(?m)^Program received signal (SIG[A-Z0-9]+),", output)
            if crash:
                result["crashed"] = True
                result["error"] = (
                    f"Headless FreeCAD reported a crash with {crash.group(1)} "
                    f"(exit code {proc.returncode}; GUI unaffected)"
                )
            else:
                result["error"] = f"Script failed with exit code {proc.returncode}: {output}"

        return result

    except subprocess.TimeoutExpired as e:
        out = "\n".join(part for part in (_clean(e.stdout), _clean(e.stderr)) if part)
        return {
            "success": False,
            "timed_out": True,
            "error": f"Headless FreeCAD script did not finish within {timeout:g} seconds",
            "output": out,
        }
    except OSError as e:
        return {"success": False, "error": f"Could not start headless FreeCAD: {e}"}
    finally:
        if script_path and os.path.exists(script_path):
            try:
                os.unlink(script_path)
            except OSError:
                pass


# ==============================================================================
# 3. FreeCAD Unified Engine Layer
# ==============================================================================

class FreeCADEngine:
    """Unified FreeCAD engine coordinating RPC, HTTP bridge, and headless execution."""
    
    RPC_PORT = int(os.environ.get("FREECAD_RPC_PORT", "9875"))
    HTTP_PORT = int(os.environ.get("FREECAD_BRIDGE_PORT", "9877"))
    DEFAULT_HOST = os.environ.get("FREECAD_HOST", "127.0.0.1")
    DEFAULT_TOKEN = os.environ.get("FREECAD_MCP_TOKEN", None)
    
    _rpc_conn: Optional[FreeCADConnection] = None

    @classmethod
    def get_rpc_connection(cls, host: Optional[str] = None, port: Optional[int] = None, token: Optional[str] = None) -> FreeCADConnection:
        target_host = host or cls.DEFAULT_HOST
        target_port = port or cls.RPC_PORT
        target_token = token if token is not None else cls.DEFAULT_TOKEN

        if cls._rpc_conn is not None:
            if (cls._rpc_conn._host != target_host or 
                cls._rpc_conn._port != target_port or 
                cls._rpc_conn._token != target_token):
                cls._rpc_conn.disconnect()
                cls._rpc_conn = None
            elif not cls._rpc_conn.ping():
                cls._rpc_conn.reconnect()

        if cls._rpc_conn is None:
            cls._rpc_conn = FreeCADConnection(
                host=target_host,
                port=target_port,
                token=target_token
            )
        return cls._rpc_conn

    @classmethod
    def _http_headers(cls) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if cls.DEFAULT_TOKEN:
            headers["Authorization"] = f"Bearer {cls.DEFAULT_TOKEN}"
        return headers

    @classmethod
    def check_rpc_connection(cls, timeout: float = 2.0) -> Dict[str, Any]:
        """Check if FreeCAD XML-RPC Server (Port 9875) is active."""
        try:
            conn = cls.get_rpc_connection()
            status = conn.get_rpc_status()
            if isinstance(status, dict) and status.get("success"):
                status["connected"] = True
                return status
            return {"connected": False, "error": "RPC responded but reported error"}
        except Exception as e:
            return {"connected": False, "error": str(e)}

    @classmethod
    def check_http_bridge(cls, timeout: float = 1.5) -> Dict[str, Any]:
        """Check if FreeCAD HTTP Bridge (Port 9877) is active."""
        url = f"http://{cls.DEFAULT_HOST}:{cls.HTTP_PORT}/status"
        try:
            res = requests.get(url, headers=cls._http_headers(), timeout=timeout)
            if res.status_code == 200:
                data = res.json()
                data["connected"] = True
                return data
            return {"connected": False, "error": f"HTTP status {res.status_code}"}
        except Exception as e:
            return {"connected": False, "error": str(e)}

    @classmethod
    def get_status(cls) -> Dict[str, Any]:
        """Check comprehensive status across RPC, HTTP bridge, and Headless FreeCAD."""
        rpc_stat = cls.check_rpc_connection(timeout=1.0)
        http_stat = cls.check_http_bridge(timeout=1.0)
        cmd_path = detect_freecadcmd()

        return {
            "rpc_connected": rpc_stat.get("connected", False),
            "rpc_details": rpc_stat,
            "http_bridge_connected": http_stat.get("connected", False),
            "http_bridge_details": http_stat,
            "headless_available": bool(cmd_path),
            "freecadcmd_path": cmd_path[0] if cmd_path else None,
        }

    @classmethod
    def execute_code(cls, script_code: str, timeout: float = 90.0) -> Dict[str, Any]:
        """
        Executes a Python script using the best available execution method:
        1. Live GUI via XML-RPC (port 9875)
        2. Live GUI via HTTP Bridge (port 9877)
        3. Headless FreeCADCmd subprocess
        """
        # 1. Try Live XML-RPC
        if cls.check_rpc_connection(timeout=0.8).get("connected"):
            conn = cls.get_rpc_connection()
            res = conn.execute_code(script_code, timeout=timeout)
            if isinstance(res, dict) and res.get("success"):
                msg = res.get("message", "")
                output = ""
                if "Output:" in msg:
                    output = msg.split("Output:", 1)[1].strip()
                return {
                    "success": True,
                    "mode": "rpc_gui",
                    "stdout": output or msg,
                    "raw": res
                }
            elif isinstance(res, dict):
                return {
                    "success": False,
                    "mode": "rpc_gui",
                    "error": res.get("error", "Execution failed")
                }

        # 2. Try HTTP Bridge
        if cls.check_http_bridge(timeout=0.8).get("connected"):
            url = f"http://{cls.DEFAULT_HOST}:{cls.HTTP_PORT}/execute"
            try:
                r = requests.post(url, json={"code": script_code}, headers=cls._http_headers(), timeout=timeout)
                if r.status_code == 200:
                    data = r.json()
                    data["mode"] = "http_bridge"
                    return data
            except Exception as e:
                logger.warning(f"HTTP bridge request error: {e}")

        # 3. Try Headless FreeCADCmd
        cmd_path = detect_freecadcmd()
        if cmd_path:
            res = run_headless(script_code, timeout=timeout, command=cmd_path)
            res["mode"] = "headless_cmd"
            res["stdout"] = res.get("output", "")
            return res

        return {
            "success": False,
            "error": (
                "No FreeCAD runtime available. FreeCAD is neither running with the Addon/Macro bridge "
                "nor found on disk for headless execution. Please start FreeCAD with the MCP Addon or install FreeCAD."
            )
        }


# ==============================================================================
# 4. HTTP Bridge Security & Request Verification Layer
# ==============================================================================

def _get_auth_token() -> Optional[str]:
    token = os.environ.get("FREECAD_MCP_TOKEN", "")
    if not token:
        try:
            from rpc_server.settings import load_settings
            token = load_settings().get("auth_token", "")
        except Exception:
            pass
    return token or None


class FreeCADBridgeHTTPHandler(BaseHTTPRequestHandler):
    """HTTP Request handler with security verification for DNS rebinding and CSRF."""
    
    def _set_headers(self, status=200, content_type="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()

    def _send_error(self, status: int, message: str):
        self._set_headers(status)
        self.wfile.write(json.dumps({"error": message}).encode("utf-8"))

    @classmethod
    def _verify_request(cls, handler_instance: Any):
        # 1. CSRF / Browser guard: refuse requests from web browsers
        if handler_instance.headers.get("Origin") is not None:
            return 403, "Requests from web browsers/origins are not accepted"

        # 2. DNS rebinding check
        host = handler_instance.headers.get("Host", "")
        host_clean = host.split(":")[0].strip().lower()
        if host_clean not in ("127.0.0.1", "localhost", "::1"):
            return 403, "Requests must be addressed to localhost or 127.0.0.1"

        # 3. Auth Token check if configured
        expected_token = _get_auth_token()
        if expected_token:
            auth_hdr = handler_instance.headers.get("Authorization", "")
            token_provided = ""
            if auth_hdr.startswith("Bearer "):
                token_provided = auth_hdr[7:].strip()
            elif auth_hdr.startswith("Basic "):
                import base64
                try:
                    decoded = base64.b64decode(auth_hdr[6:].strip()).decode("utf-8")
                    token_provided = decoded.split(":", 1)[-1]
                except Exception:
                    pass

            import hmac
            if not token_provided or not hmac.compare_digest(token_provided, expected_token):
                return 401, "Unauthorized: Valid Bearer token required"

        return None

ViewName = Literal[
    "Isometric", "Front", "Top", "Right", "Back", "Left", "Bottom", "Dimetric", "Trimetric"
]

ToolResponse = List[Union[TextContent, ImageContent]]


@dataclass
class ServerState:
    only_text_feedback: bool = True
    rpc_host: str = os.environ.get("FREECAD_HOST", "localhost")
    rpc_port: int = int(os.environ.get("FREECAD_RPC_PORT", "9875"))
    auth_token: Optional[str] = os.environ.get("FREECAD_MCP_TOKEN", None)
    freecad_connection: Optional[FreeCADConnection] = None
    freecadcmd: Optional[List[str]] = None
    version_notice: Optional[str] = None


state = ServerState()


def get_freecad_connection() -> FreeCADConnection:
    conn = state.freecad_connection
    if conn is None or not conn.ping():
        if conn is not None:
            conn.disconnect()
        conn = FreeCADConnection(
            host=state.rpc_host,
            port=state.rpc_port,
            token=state.auth_token,
        )
        try:
            reachable = conn.ping()
        except Exception as e:
            conn.disconnect()
            state.freecad_connection = None
            raise Exception(f"Failed to connect to FreeCAD RPC ({e}). Make sure FreeCAD is running with the Addon or Macro bridge.") from e
        
        if not reachable:
            conn.disconnect()
            state.freecad_connection = None
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


# Initialize FastMCP Server with Comprehensive Operational Rules
mcp = FastMCP(
    "FreeCADMCP",
    instructions="""Comprehensive FreeCAD Model Context Protocol (MCP) server for 3D parametric CAD modeling, TechDraw, FEM analysis, and automation.

⭐ MANDATORY CORE RULES & OPERATIONAL CONSTRAINTS:

1. NO STANDALONE .PY SCRIPT FILES ON DISK:
   - Execute all CAD scripts, parametric commands, and geometry operations directly in-memory via the live RPC bridge.
   - Never write, generate, or save temporary or scratch .py files to disk unless the user explicitly requests a script file.

2. FAST MODE & NO SCREENSHOTS BY DEFAULT:
   - Do NOT capture, request, or return viewport screenshots by default.
   - Run in high-speed text/geometry verification mode.
   - Capture and return a viewport image ONLY when the user explicitly commands it (e.g. "take a picture", "show screenshot", "أرني لقطة").

3. VISUAL STEP-BY-STEP LIVE ANIMATION:
   - Never apply geometric changes or modifications in a sudden, single unobservable jump.
   - Break construction into observable stages:
     * Stage 1 (Wireframe/Contour): Trace wireframes/edges progressively.
     * Stage 2 (Planar Face): Form closed faces.
     * Stage 3 (3D Extrusion/Pocketing): Extrude or cut in depth stages.
     * Stage 4 (Finishing): Apply fillets, chamfers, or colors.
   - In live GUI scripts, use `FreeCADGui.updateGui()` with brief pacing pauses (`time.sleep(0.04)` to `0.15`) between segments so the user sees the CAD model constructing live in the viewport.

4. LIVE THINKING & VISIBLE ITERATION (TRIAL-AND-ERROR ON SCREEN):
   - Never compute all design choices, color alternatives, and dimensional refinements internally in silence and then output a final polished block all at once.
   - Externalize your iterative thinking onto the live 3D canvas:
     * When experimenting with color palettes or styling, apply the initial idea live to the model first, let the user see it, then refine or correct it live.
     * When tuning dimensions, positions, or features, place the initial draft geometry on the canvas first, then adjust and perfect it in progressive observable steps.
   - Structure actions as live, interactive micro-actions so the user experiences the model evolving through visible exploration and continuous refinement.

5. USER OVERRIDE:
   - If the user explicitly asks for screenshots or for .py files, fulfill that specific request. Otherwise, adhere strictly to the in-memory, no-screenshot, step-by-step, visible-iteration default rules."""
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
def freecad_status() -> ToolResponse:
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

    return json_response({"status_summary": "\n".join(lines), **status})


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
    payload = {
        "primitive_type": primitive_type.lower().strip(),
        "name": name or "",
        "dimensions": dimensions or {},
        "position": position or [0.0, 0.0, 0.0],
        "rotation_axis_angle": rotation_axis_angle or [0.0, 0.0, 1.0, 0.0],
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, Part, json
payload = json.loads({json.dumps(payload_json)})
p_type = payload["primitive_type"]
obj_name = payload["name"]
dims = payload["dimensions"]
pos = payload["position"]
rot = payload["rotation_axis_angle"]
doc_name = payload["document_name"]

doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    doc = FreeCAD.newDocument("CAD_Model")
    FreeCAD.setActiveDocument(doc.Name)

if p_type == "box":
    obj = doc.addObject("Part::Box", obj_name or "Box")
    obj.Length = float(dims.get('length', 10.0))
    obj.Width = float(dims.get('width', 10.0))
    obj.Height = float(dims.get('height', 10.0))

elif p_type == "cylinder":
    obj = doc.addObject("Part::Cylinder", obj_name or "Cylinder")
    obj.Radius = float(dims.get('radius', 5.0))
    obj.Height = float(dims.get('height', 10.0))
    obj.Angle = float(dims.get('angle', 360.0))

elif p_type == "sphere":
    obj = doc.addObject("Part::Sphere", obj_name or "Sphere")
    obj.Radius = float(dims.get('radius', 5.0))

elif p_type == "cone":
    obj = doc.addObject("Part::Cone", obj_name or "Cone")
    obj.Radius1 = float(dims.get('radius1', 10.0))
    obj.Radius2 = float(dims.get('radius2', 0.0))
    obj.Height = float(dims.get('height', 20.0))

elif p_type == "torus":
    obj = doc.addObject("Part::Torus", obj_name or "Torus")
    obj.Radius1 = float(dims.get('radius1', 20.0))
    obj.Radius2 = float(dims.get('radius2', 5.0))

elif p_type == "tube":
    try:
        from BasicShapes import Shapes
        obj = Shapes.addTube(doc, obj_name or "Tube")
        obj.InnerRadius = float(dims.get('inner_radius', 7.0))
        obj.OuterRadius = float(dims.get('outer_radius', 10.0))
        obj.Height = float(dims.get('height', 20.0))
    except Exception:
        obj = doc.addObject("Part::Tube", obj_name or "Tube")
        if hasattr(obj, "InnerRadius"):
            obj.InnerRadius = float(dims.get('inner_radius', 7.0))
            obj.OuterRadius = float(dims.get('outer_radius', 10.0))
            obj.Height = float(dims.get('height', 20.0))
        else:
            outer = Part.makeCylinder(float(dims.get('outer_radius', 10.0)), float(dims.get('height', 20.0)))
            inner = Part.makeCylinder(float(dims.get('inner_radius', 7.0)), float(dims.get('height', 20.0)))
            obj = doc.addObject("Part::Feature", obj_name or "Tube")
            obj.Shape = outer.cut(inner)

elif p_type == "wedge":
    obj = doc.addObject("Part::Wedge", obj_name or "Wedge")
    obj.Xmin = float(dims.get('xmin', 0.0))
    obj.Ymin = float(dims.get('ymin', 0.0))
    obj.Zmin = float(dims.get('zmin', 0.0))
    obj.Xmax = float(dims.get('xmax', 20.0))
    obj.Ymax = float(dims.get('ymax', 20.0))
    obj.Zmax = float(dims.get('zmax', 20.0))
    if hasattr(obj, "X2min"):
        obj.X2min = float(dims.get('x2min', dims.get('xmin', 0.0)))
        obj.Z2min = float(dims.get('z2min', dims.get('zmin', 0.0)))
        obj.X2max = float(dims.get('x2max', dims.get('xmax', 20.0)))
        obj.Z2max = float(dims.get('z2max', dims.get('zmax', 20.0)))

elif p_type == "plane":
    obj = doc.addObject("Part::Plane", obj_name or "Plane")
    obj.Length = float(dims.get('length', 100.0))
    obj.Width = float(dims.get('width', 100.0))

else:
    raise ValueError(f"Unknown primitive type: {{p_type}}")

obj.Placement.Base = FreeCAD.Vector(float(pos[0]), float(pos[1]), float(pos[2]))
obj.Placement.Rotation = FreeCAD.Rotation(FreeCAD.Vector(float(rot[0]), float(rot[1]), float(rot[2])), float(rot[3]))
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

    payload = {
        "op": op,
        "obj_type": obj_type,
        "base_object": base_object,
        "tool_object": tool_object,
        "result_name": result_name,
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

base = doc.getObject(payload["base_object"])
tool = doc.getObject(payload["tool_object"])
if not base:
    raise ValueError(f"Base object '{payload['base_object']}' not found.")
if not tool:
    raise ValueError(f"Tool object '{payload['tool_object']}' not found.")

bool_obj = doc.addObject(payload["obj_type"], payload["result_name"] or payload["op"].capitalize())
bool_obj.Base = base
bool_obj.Tool = tool
doc.recompute()

if hasattr(base, 'ViewObject') and base.ViewObject:
    base.ViewObject.Visibility = False
if hasattr(tool, 'ViewObject') and tool.ViewObject:
    tool.ViewObject.Visibility = False

doc.recompute()
print(f"Boolean operation '{payload['op']}' completed: created {{bool_obj.Name}}")
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
    if op not in ("fillet", "chamfer"):
        return text_response(f"❌ Unknown operation type '{operation_type}'. Use 'fillet' or 'chamfer'.")

    payload = {
        "op": op,
        "object_name": object_name,
        "radius_or_size": float(radius_or_size),
        "edge_indices": edge_indices or [],
        "result_name": result_name,
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, Part, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

target = doc.getObject(payload["object_name"])
if not target:
    raise ValueError(f"Target object '{payload['object_name']}' not found.")

edges_in = payload["edge_indices"]
if not edges_in:
    edges_in = list(range(1, len(target.Shape.Edges) + 1))

r = payload["radius_or_size"]
op = payload["op"]
if op == "fillet":
    obj = doc.addObject("Part::Fillet", payload["result_name"] or "Fillet")
    obj.Base = target
    obj.Edges = [(int(idx), r, r) for idx in edges_in]
elif op == "chamfer":
    obj = doc.addObject("Part::Chamfer", payload["result_name"] or "Chamfer")
    obj.Base = target
    obj.Edges = [(int(idx), r, r) for idx in edges_in]
else:
    raise ValueError(f"Unknown operation type: {{op}}")

doc.recompute()
if hasattr(target, 'ViewObject') and target.ViewObject:
    target.ViewObject.Visibility = False
doc.recompute()
print(f"{{op.capitalize()}} applied to {{len(edges_in)}} edges on {{target.Name}}.")
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
    payload = {
        "plane": plane.upper().strip(),
        "extrude_length": float(extrude_length),
        "elements": elements or [{"type": "rectangle", "corner1": [-20, -10], "corner2": [20, 10]}],
        "sketch_name": sketch_name or "Sketch",
        "body_name": body_name or "Body",
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, Sketcher, Part, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    doc = FreeCAD.newDocument("CAD_Model")
    FreeCAD.setActiveDocument(doc.Name)

body = doc.addObject("PartDesign::Body", payload["body_name"] or "Body")
sketch = body.newObject("Sketcher::SketchObject", payload["sketch_name"] or "Sketch")

pln = payload["plane"]
plane_map = {{"XY": "XY_Plane", "XZ": "XZ_Plane", "YZ": "YZ_Plane"}}
plane_obj_name = plane_map.get(pln, "XY_Plane")

if hasattr(body, "Origin") and body.Origin:
    origin_plane = getattr(body.Origin, plane_obj_name, None)
    if not origin_plane and hasattr(body.Origin, "OriginFeatures"):
        for f in body.Origin.OriginFeatures:
            if plane_obj_name in f.Name or plane_obj_name in f.Label:
                origin_plane = f
                break
    if origin_plane:
        sketch.AttachmentSupport = [(origin_plane, "")]
        sketch.MapMode = "FlatFace"

for el in payload["elements"]:
    t = el.get("type", "").lower()
    if t == "circle":
        c = el.get("center", [0, 0])
        r = float(el.get("radius", 10.0))
        i_c = sketch.addGeometry(Part.Circle(FreeCAD.Vector(float(c[0]), float(c[1]), 0), FreeCAD.Vector(0, 0, 1), r), False)
        sketch.addConstraint(Sketcher.Constraint("Radius", i_c, r))
    elif t == "rectangle":
        c1 = el.get("corner1", [-10, -10])
        c2 = el.get("corner2", [10, 10])
        x1, y1 = float(c1[0]), float(c1[1])
        x2, y2 = float(c2[0]), float(c2[1])
        p1 = FreeCAD.Vector(x1, y1, 0)
        p2 = FreeCAD.Vector(x2, y1, 0)
        p3 = FreeCAD.Vector(x2, y2, 0)
        p4 = FreeCAD.Vector(x1, y2, 0)
        i0 = sketch.addGeometry(Part.LineSegment(p1, p2), False)
        i1 = sketch.addGeometry(Part.LineSegment(p2, p3), False)
        i2 = sketch.addGeometry(Part.LineSegment(p3, p4), False)
        i3 = sketch.addGeometry(Part.LineSegment(p4, p1), False)
        sketch.addConstraint(Sketcher.Constraint("Coincident", i0, 2, i1, 1))
        sketch.addConstraint(Sketcher.Constraint("Coincident", i1, 2, i2, 1))
        sketch.addConstraint(Sketcher.Constraint("Coincident", i2, 2, i3, 1))
        sketch.addConstraint(Sketcher.Constraint("Coincident", i3, 2, i0, 1))
        sketch.addConstraint(Sketcher.Constraint("Horizontal", i0))
        sketch.addConstraint(Sketcher.Constraint("Horizontal", i2))
        sketch.addConstraint(Sketcher.Constraint("Vertical", i1))
        sketch.addConstraint(Sketcher.Constraint("Vertical", i3))
    elif t == "line":
        s = el.get("start", [0, 0])
        e = el.get("end", [10, 10])
        sketch.addGeometry(Part.LineSegment(FreeCAD.Vector(float(s[0]), float(s[1]), 0), FreeCAD.Vector(float(e[0]), float(e[1]), 0)), False)

doc.recompute()
pad = body.newObject("PartDesign::Pad", f"Pad_{{sketch.Name}}")
pad.Profile = sketch
pad.Length = float(payload["extrude_length"])
doc.recompute()
print(f"Created Sketch on {{pln}} and Extruded Pad {{pad.Name}} with length {{pad.Length}}mm.")
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
    payload = {
        "shape_name": shape_name,
        "axis_vector": axis_vector or [0.0, 0.0, 1.0],
        "axis_base": axis_base or [0.0, 0.0, 0.0],
        "angle_degrees": float(angle_degrees),
        "result_name": result_name,
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, Part, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

src = doc.getObject(payload["shape_name"])
if not src:
    raise ValueError(f"Source object '{payload['shape_name']}' not found.")

ax = payload["axis_vector"]
base = payload["axis_base"]
rev = doc.addObject("Part::Revolution", payload["result_name"] or "Revolution")
rev.Source = src
rev.Axis = FreeCAD.Vector(float(ax[0]), float(ax[1]), float(ax[2]))
rev.Base = FreeCAD.Vector(float(base[0]), float(base[1]), float(base[2]))
rev.Angle = float(payload["angle_degrees"])
doc.recompute()
if hasattr(src, 'ViewObject') and src.ViewObject:
    src.ViewObject.Visibility = False
doc.recompute()
print(f"Revolved {{src.Name}} by {{rev.Angle}} degrees around axis {{rev.Axis}}.")
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
) -> ToolResponse:
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
    payload = {
        "object_name": object_name,
        "page_name": page_name,
        "template_name": template_name,
        "scale": float(scale),
        "projection_views": projection_views or ["Front", "Top", "Right", "Isometric"],
        "export_filepath": export_filepath,
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

target = doc.getObject(payload["object_name"])
if not target:
    raise ValueError(f"Target object '{payload['object_name']}' not found.")

import TechDraw, os
page = doc.addObject("TechDraw::DrawPage", payload["page_name"] or "TechDraw_Sheet")
template = doc.addObject("TechDraw::DrawSVGTemplate", "Template")

template_file = payload["template_name"]
res_dir = FreeCAD.getResourceDir()
cand_path = os.path.join(res_dir, "Mod", "TechDraw", "Templates", template_file)
if not os.path.exists(cand_path):
    cand_path = os.path.join(res_dir, "Mod", "TechDraw", "Templates", "Default_Template_A4_Landscape.svg")
if not os.path.exists(cand_path) and hasattr(TechDraw, "findTemplate"):
    cand_path = TechDraw.findTemplate(template_file)
template.Template = cand_path
page.Template = template
doc.recompute()

views = payload["projection_views"]
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
    dv.Scale = payload["scale"]
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

exp_path = payload.get("export_filepath")
if exp_path:
    norm_exp = exp_path.replace("\\\\", "/")
    os.makedirs(os.path.dirname(norm_exp) or ".", exist_ok=True)
    if norm_exp.lower().endswith(".pdf"):
        import TechDrawGui
        TechDrawGui.exportPageAsPdf(page, norm_exp)
    elif norm_exp.lower().endswith(".dxf"):
        TechDraw.writeDXFPage(page, norm_exp)
    else:
        import TechDrawGui
        TechDrawGui.exportPageAsSvg(page, norm_exp)
    print(f"TechDraw exported to: {{norm_exp}}")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to create TechDraw: {res.get('error')}")
    return text_response(f"✅ Technical Drawing created successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_analyze_geometry(
    object_name: str,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Calculate physical and geometric mass properties: volume, surface area, center of mass, bounding box.
    
    Args:
        object_name: Name of solid object to inspect.
        document_name: Optional document name.
    """
    payload = {
        "object_name": object_name,
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

obj_name = payload["object_name"]
obj = doc.getObject(obj_name)
if not obj or not hasattr(obj, "Shape") or obj.Shape.isNull():
    raise ValueError(f"Object '{obj_name}' not found or has no solid shape.")

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
        return text_response(f"❌ Geometry analysis failed: {res.get('error')}")
    stdout = res.get("stdout", "")
    if "__JSON_ANALYSIS__:" in stdout:
        json_str = stdout.split("__JSON_ANALYSIS__:")[1].strip()
        try:
            return json_response(json.loads(json_str))
        except Exception:
            return text_response(json_str)
    return text_response(stdout)


@mcp.tool()
def freecad_export_file(
    export_format: str,
    filepath: str,
    object_names: Optional[List[str]] = None,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Export FreeCAD 3D objects to STEP, STL, OBJ, IGES, BREP, DXF, SVG, or PDF.
    
    Args:
        export_format: 'step', 'stp', 'stl', 'obj', 'iges', 'brep', 'dxf', 'svg', or 'pdf'.
        filepath: Destination file path (e.g. 'E:/exports/bracket.step').
        object_names: List of object names to export. Defaults to all objects.
        document_name: Optional document name.
    """
    payload = {
        "export_format": export_format.lower().replace(".", "").strip(),
        "filepath": filepath.replace("\\", "/"),
        "object_names": object_names or [],
        "document_name": document_name,
    }
    payload_json = json.dumps(payload)

    code = f"""
import FreeCAD, Part, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

req_objs = payload["object_names"]
if req_objs:
    export_list = [doc.getObject(name) for name in req_objs if doc.getObject(name)]
else:
    export_list = [o for o in doc.Objects if hasattr(o, 'Shape') and o.Shape and not o.Shape.isNull()]

if not export_list:
    raise RuntimeError("No valid solid objects found to export.")

fmt = payload["export_format"]
dest = payload["filepath"]

if fmt in ["step", "stp"]:
    import Import
    Import.export(export_list, dest)
elif fmt == "stl":
    import Mesh
    mesh_objs = []
    try:
        for o in export_list:
            m = doc.addObject("Mesh::Feature", f"__temp_mesh_{{o.Name}}")
            m.Mesh = Mesh.Mesh(o.Shape.tessellate(0.5))
            mesh_objs.append(m)
        Mesh.export(mesh_objs, dest)
    finally:
        for m in mesh_objs:
            try:
                doc.removeObject(m.Name)
            except Exception:
                pass
elif fmt == "obj":
    import ImportGui
    ImportGui.export(export_list, dest)
elif fmt in ["iges", "igs"]:
    import Import
    Import.export(export_list, dest)
elif fmt == "brep":
    export_list[0].Shape.exportBrep(dest)
elif fmt == "pdf":
    pages = [o for o in doc.Objects if hasattr(o, "TypeId") and "TechDraw::DrawPage" in o.TypeId]
    if pages:
        import TechDrawGui
        TechDrawGui.exportPage(pages[0], dest)
    else:
        import Import
        Import.export(export_list, dest)
else:
    import Import
    Import.export(export_list, dest)

print(f"Exported {{len(export_list)}} objects to {{dest}} in {{fmt.upper()}} format.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Export failed: {res.get('error')}")
    return text_response(f"✅ File exported successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_import_file(
    filepath: str,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Import an external 3D CAD or mesh file (.step, .stl, .obj, .iges, .dxf) into FreeCAD.
    
    Args:
        filepath: Full path to the CAD file.
        document_name: Optional target document name.
    """
    payload = {
        "filepath": filepath.replace("\\", "/"),
        "document_name": document_name
    }
    payload_json = json.dumps(payload)
    code = f"""
import FreeCAD, Import, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    doc = FreeCAD.newDocument("Imported_CAD")
    FreeCAD.setActiveDocument(doc.Name)

norm_path = payload["filepath"]
Import.insert(norm_path, doc.Name)
doc.recompute()
print(f"Imported '{{norm_path}}' into document '{{doc.Name}}'. Total objects: {{len(doc.Objects)}}")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Import failed: {res.get('error')}")
    return text_response(f"✅ File imported successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_open_document(filepath: str) -> ToolResponse:
    """
    Open an existing FreeCAD document file (.FCStd).
    
    Args:
        filepath: Full path to the .FCStd document file.
    """
    payload = {"filepath": filepath.replace("\\", "/")}
    payload_json = json.dumps(payload)
    code = f"""
import FreeCAD, json
payload = json.loads({json.dumps(payload_json)})
norm_path = payload["filepath"]
doc = FreeCAD.openDocument(norm_path)
if doc:
    FreeCAD.setActiveDocument(doc.Name)
    print(f"Opened document '{{doc.Name}}' from '{{norm_path}}'. Objects count: {{len(doc.Objects)}}")
else:
    raise RuntimeError("Failed to open document: 'FreeCAD.openDocument' returned None.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to open document: {res.get('error')}")
    return text_response(f"✅ Document opened successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_save_document(
    filepath: Optional[str] = None,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Save an open FreeCAD document (.FCStd) to disk.
    
    Args:
        filepath: Optional destination file path to save as (e.g. 'C:/models/part.FCStd'). If omitted, saves to the document's current file path.
        document_name: Optional name of the document to save. If omitted, saves the active document.
    """
    payload = {
        "filepath": filepath.replace("\\", "/") if filepath else None,
        "document_name": document_name
    }
    payload_json = json.dumps(payload)
    code = f"""
import FreeCAD, json
payload = json.loads({json.dumps(payload_json)})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found to save.")

norm_path = payload.get("filepath")
if norm_path:
    doc.saveAs(norm_path)
    print(f"Document '{{doc.Name}}' saved as '{{norm_path}}'.")
else:
    if not getattr(doc, "FileName", None):
        raise RuntimeError("Document has no existing file path on disk. Please specify a 'filepath' parameter to save.")
    doc.save()
    print(f"Document '{{doc.Name}}' saved to '{{doc.FileName}}'.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to save document: {res.get('error')}")
    return text_response(f"✅ Document saved successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_get_active_project_info() -> ToolResponse:
    """
    Intelligently inspect and auto-detect the active project/document open in FreeCAD.
    Returns the document name, label, file path on disk, modified status, object count, 
    and complete list of model objects without needing to specify any folder or file path.
    """
    code = """
import FreeCAD, json

doc = FreeCAD.ActiveDocument
if not doc:
    docs = list(FreeCAD.listDocuments().values())
    if docs:
        doc = docs[0]
        FreeCAD.setActiveDocument(doc.Name)

if not doc:
    print(json.dumps({"has_active_project": False, "message": "No open project found in FreeCAD."}))
else:
    objs = []
    bbox = None
    for o in doc.Objects:
        objs.append({
            "name": o.Name,
            "label": getattr(o, "Label", o.Name),
            "type": getattr(o, "TypeId", "Unknown"),
            "visible": getattr(o.ViewObject, "Visibility", True) if hasattr(o, "ViewObject") and o.ViewObject else True
        })
        if hasattr(o, "Shape") and o.Shape and not o.Shape.isNull():
            b = o.Shape.BoundBox
            if bbox is None:
                bbox = [b.XMin, b.YMin, b.ZMin, b.XMax, b.YMax, b.ZMax]
            else:
                bbox = [
                    min(bbox[0], b.XMin), min(bbox[1], b.YMin), min(bbox[2], b.ZMin),
                    max(bbox[3], b.XMax), max(bbox[4], b.YMax), max(bbox[5], b.ZMax)
                ]

    data = {
        "has_active_project": True,
        "name": doc.Name,
        "label": getattr(doc, "Label", doc.Name),
        "filepath": getattr(doc, "FileName", "") or "Unsaved (in memory)",
        "object_count": len(doc.Objects),
        "bounding_box": bbox,
        "objects": objs
    }
    print(json.dumps(data))
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to detect active project: {res.get('error')}")
    try:
        data = json.loads(res.get("stdout", "{}"))
        return json_response(data)
    except Exception:
        return text_response(res.get("stdout", ""))


@mcp.tool()
def freecad_apply_color_matrix(
    object_name: Optional[str] = None,
    base_color: Optional[List[float]] = None,
    face_color_map: Optional[Dict[str, List[float]]] = None,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Apply a multi-face DiffuseColor matrix to any 3D model with solid opacity (Transparency=0)
    and high-resolution mesh deflection (Deviation=0.01).
    
    Args:
        object_name: Name of model to color (defaults to active or top solid).
        base_color: [R, G, B] for exterior surfaces (default white [0.95, 0.95, 0.95]).
        face_color_map: Optional dict mapping feature names (e.g. 'star', 'circle', 'rhombus') to RGB colors.
        document_name: Optional document name.
    """
    payload = {
        "object_name": object_name,
        "base_color": base_color or [0.95, 0.95, 0.95],
        "face_color_map": face_color_map or {},
        "document_name": document_name
    }
    code = f"""
import FreeCAD, FreeCADGui, json
payload = json.loads({json.dumps(json.dumps(payload))})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

obj_name = payload["object_name"]
if obj_name:
    obj = doc.getObject(obj_name)
else:
    solids = [o for o in doc.Objects if hasattr(o, "Shape") and o.Shape and not o.Shape.isNull()]
    visible_solids = [o for o in solids if hasattr(o, "ViewObject") and o.ViewObject and getattr(o.ViewObject, "Visibility", True)]
    obj = visible_solids[-1] if visible_solids else (solids[-1] if solids else None)

if not obj:
    raise ValueError("Target object not found.")

gui_doc = FreeCADGui.getDocument(doc.Name)
gui_obj = gui_doc.getObject(obj.Name)

base_c = payload["base_color"]
base_rgba = (float(base_c[0]), float(base_c[1]), float(base_c[2]), 1.0)

diffuse = []
bbox = obj.Shape.BoundBox
bx_min, bx_max = bbox.XMin, bbox.XMax
by_min, by_max = bbox.YMin, bbox.YMax
bz_min, bz_max = bbox.ZMin, bbox.ZMax

palette = [
    (1.0, 0.85, 0.0, 1.0),
    (0.9, 0.15, 0.15, 1.0),
    (0.0, 0.75, 0.95, 1.0),
    (0.15, 0.8, 0.25, 1.0),
    (1.0, 0.55, 0.05, 1.0),
    (0.65, 0.15, 0.85, 1.0)
]

for idx, f in enumerate(obj.Shape.Faces):
    c = f.CenterOfMass
    is_outer = (
        abs(c.x - bx_min) < 0.05 or abs(c.x - bx_max) < 0.05 or
        abs(c.y - by_min) < 0.05 or abs(c.y - by_max) < 0.05 or
        abs(c.z - bz_min) < 0.05 or abs(c.z - bz_max) < 0.05
    )
    if is_outer:
        diffuse.append(base_rgba)
    else:
        color_idx = int(abs(c.x * 3 + c.y * 5 + c.z * 7)) % len(palette)
        diffuse.append(palette[color_idx])

gui_obj.Transparency = 0
gui_obj.DiffuseColor = diffuse
gui_obj.Deviation = 0.01
doc.recompute()
FreeCADGui.updateGui()
print(f"Applied multi-color DiffuseColor matrix to {{len(diffuse)}} faces on '{{obj.Name}}'.")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to apply color matrix: {res.get('error')}")
    return text_response(f"✅ Color matrix applied successfully!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_animate_mechanism(
    object_name: Optional[str] = None,
    axis: Optional[List[float]] = None,
    angle_per_frame: float = 2.5,
    fps: int = 35,
    duration_seconds: float = 5.0,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Animate real-time smooth rotation or movement of a target object in FreeCAD's live 3D viewport
    using PySide QtCore.QTimer (30-40 FPS non-blocking execution with automatic cleanup).
    
    Args:
        object_name: Name of object to animate (defaults to top active solid).
        axis: [x, y, z] direction vector of rotation axis (default [0, 0, 1]).
        angle_per_frame: Rotation angle in degrees per frame (default 2.5).
        fps: Target frame rate (default 35).
        duration_seconds: Animation runtime in seconds before automatically stopping (default 5.0).
        document_name: Optional document name.
    """
    payload = {
        "object_name": object_name,
        "axis": axis or [0.0, 0.0, 1.0],
        "angle_per_frame": float(angle_per_frame),
        "fps": int(fps),
        "duration_seconds": float(duration_seconds),
        "document_name": document_name
    }
    code = f"""
import FreeCAD, FreeCADGui, json, time
try:
    from PySide import QtCore
except ImportError:
    try:
        from PySide2 import QtCore
    except ImportError:
        from PySide6 import QtCore

payload = json.loads({json.dumps(json.dumps(payload))})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

obj_name = payload["object_name"]
if obj_name:
    target = doc.getObject(obj_name)
else:
    solids = [o for o in doc.Objects if hasattr(o, "Shape") and o.Shape and not o.Shape.isNull()]
    visible_solids = [o for o in solids if hasattr(o, "ViewObject") and o.ViewObject and getattr(o.ViewObject, "Visibility", True)]
    target = visible_solids[-1] if visible_solids else (solids[-1] if solids else None)

if not target:
    raise ValueError("Target object for animation not found.")

rot_axis = FreeCAD.Vector(payload["axis"][0], payload["axis"][1], payload["axis"][2])
step_deg = payload["angle_per_frame"]
total_steps = int(payload["duration_seconds"] * payload["fps"])
interval_ms = int(1000.0 / payload["fps"])

if hasattr(FreeCAD, "__mcp_anim_timer") and FreeCAD.__mcp_anim_timer:
    try:
        FreeCAD.__mcp_anim_timer.stop()
    except Exception:
        pass

current_step = [0]
timer = QtCore.QTimer()

def on_tick():
    if current_step[0] >= total_steps:
        timer.stop()
        FreeCAD.__mcp_anim_timer = None
        return
    current_step[0] += 1
    cur_rot = target.Placement.Rotation
    delta_rot = FreeCAD.Rotation(rot_axis, step_deg)
    target.Placement.Rotation = delta_rot.multiply(cur_rot)
    FreeCADGui.updateGui()

timer.timeout.connect(on_tick)
timer.start(interval_ms)
FreeCAD.__mcp_anim_timer = timer
print(f"Started live QTimer mechanism animation on '{{target.Name}}' ({{payload['fps']}} FPS, {{payload['duration_seconds']}}s duration).")
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to start animation: {res.get('error')}")
    return text_response(f"✅ Mechanism animation running live!\n{res.get('stdout', '')}")


@mcp.tool()
def freecad_export_production_bundle(
    object_name: Optional[str] = None,
    output_dir: Optional[str] = None,
    bundle_name: Optional[str] = None,
    document_name: Optional[str] = None
) -> ToolResponse:
    """
    Export the model into an all-in-one manufacturing bundle containing:
    - STEP (.step) for CNC/CAM
    - STL (.stl) for 3D Printing
    - Native FreeCAD project (.FCStd) backup
    Files are neatly saved into organized subfolders within the project directory.
    
    Args:
        object_name: Name of target object (defaults to top active solid).
        output_dir: Root output directory (defaults to 'E:/AI Projects/MCP/FreeCAD/Projects').
        bundle_name: Base filename (defaults to document name or object name).
        document_name: Optional document name.
    """
    payload = {
        "object_name": object_name,
        "output_dir": (output_dir or "E:/AI Projects/MCP/FreeCAD/Projects").replace("\\\\", "/"),
        "bundle_name": bundle_name,
        "document_name": document_name
    }
    code = f"""
import FreeCAD, Part, Mesh, Import, os, json

payload = json.loads({json.dumps(json.dumps(payload))})
doc_name = payload["document_name"]
doc = FreeCAD.getDocument(doc_name) if doc_name else FreeCAD.ActiveDocument
if not doc:
    raise RuntimeError("No active document found.")

obj_name = payload["object_name"]
if obj_name:
    target = doc.getObject(obj_name)
else:
    solids = [o for o in doc.Objects if hasattr(o, "Shape") and o.Shape and not o.Shape.isNull()]
    visible_solids = [o for o in solids if hasattr(o, "ViewObject") and o.ViewObject and getattr(o.ViewObject, "Visibility", True)]
    target = visible_solids[-1] if visible_solids else (solids[-1] if solids else None)

if not target:
    raise ValueError("Target object not found to export bundle.")

base_name = payload["bundle_name"] or doc.Name or target.Name
root_dir = payload["output_dir"]

step_dir = os.path.join(root_dir, "STEP").replace("\\\\", "/")
stl_dir = os.path.join(root_dir, "STL").replace("\\\\", "/")
model_dir = os.path.join(root_dir, "CAD_Models").replace("\\\\", "/")

for d in [step_dir, stl_dir, model_dir]:
    os.makedirs(d, exist_ok=True)

step_path = os.path.join(step_dir, f"{{base_name}}.step").replace("\\\\", "/")
stl_path = os.path.join(stl_dir, f"{{base_name}}.stl").replace("\\\\", "/")
fcstd_path = os.path.join(model_dir, f"{{base_name}}.FCStd").replace("\\\\", "/")

Import.export([target], step_path)

temp_mesh = doc.addObject("Mesh::Feature", f"__bundle_mesh_{{target.Name}}")
temp_mesh.Mesh = Mesh.Mesh(target.Shape.tessellate(0.2))
Mesh.export([temp_mesh], stl_path)
doc.removeObject(temp_mesh.Name)

doc.saveAs(fcstd_path)

res_data = {{
    "bundle_name": base_name,
    "target_object": target.Name,
    "step_file": step_path,
    "stl_file": stl_path,
    "cad_project": fcstd_path
}}
print("BUNDLE_JSON:" + json.dumps(res_data))
"""
    res = FreeCADEngine.execute_code(code)
    if not res.get("success"):
        return text_response(f"❌ Failed to export production bundle: {res.get('error')}")
    stdout = res.get("stdout", "")
    if "BUNDLE_JSON:" in stdout:
        data = json.loads(stdout.split("BUNDLE_JSON:")[1].strip())
        return json_response(data)
    return text_response(stdout)


# ==============================================================================
# SECTION 3: Strategy Prompts
# ==============================================================================

@mcp.prompt()
def step_by_step_construction_rule() -> str:
    """Core rule for interactive step-by-step visual CAD modeling in FreeCAD."""
    return """
FOUNDATIONAL PROTOCOL: STEP-BY-STEP VISUAL CONSTRUCTION & LIVE ANIMATION

Whenever designing or modifying any 3D feature in FreeCAD for the user:
1. Paced Execution: Do not create the complete final object in a single instantaneous step without showing intermediate progress.
2. 4-Stage Construction Sequence:
   - Stage 1 (2D Wireframe/Points): Draw reference points and wireframe outlines edge-by-edge. In Python scripts, update the viewport (`FreeCADGui.updateGui()`) with small delays (`time.sleep(0.12)`) so the user sees the drawing happen live like an animation.
   - Stage 2 (Face Generation): Close the wireframe and generate the 2D planar face with distinct highlight colors.
   - Stage 3 (3D Volumetric Extrusion/Solid): Extrude or revolve the face into a solid 3D entity.
   - Stage 4 (Finishing & Styling): Apply fillets, boolean cuts/unions, realistic colors, and transparency.
3. Multi-Step Animation: When creating moving parts (rotors, gears, spinning stars), attach a persistent Qt QTimer to animate rotation smoothly at 30-40 FPS.
4. Report: Summarize the stages clearly with screenshots and dimensions.
"""


@mcp.prompt()
def asset_creation_strategy() -> str:
    """Standard asset creation strategy for AI CAD engineers."""
    return """
Asset Creation Strategy for FreeCAD MCP:

When creating 3D models in FreeCAD:
1. Always check existing document objects using `get_objects(doc_name)` or `freecad_status()`.
2. Follow the step-by-step visual construction protocol to give live interactive feedback.
3. For standard hardware (bolts, nuts, bearings), check `get_parts_list()` and insert using `insert_part_from_library()`.
4. For custom geometry:
   - Create basic primitives (`create_object` or `freecad_create_primitive`).
   - For complex profiles, use 2D Sketches + Extrusions (`freecad_create_sketch_and_extrude`) or Revolutions (`freecad_revolve_shape`).
   - Combine solids using Boolean operations (`freecad_boolean_operation` - fuse/cut/common).
   - Add fillets and chamfers (`freecad_apply_fillet_chamfer`).
5. Keep screenshots enabled (`include_screenshot=True`) for visual verification.
6. For heavy computation or complex boolean chains, use `execute_code_async` with `commit()` or `execute_code_headless` to prevent GUI lockups.
"""


def main():
    parser = argparse.ArgumentParser(description="FreeCAD MCP Server")
    parser.add_argument("--only-text-feedback", action="store_true", help="Only return text feedback (suppress inline screenshots)")
    parser.add_argument("--host", default=os.environ.get("FREECAD_HOST", "localhost"), help="FreeCAD RPC Host (default: localhost)")
    parser.add_argument("--port", type=int, default=int(os.environ.get("FREECAD_RPC_PORT", "9875")), help="FreeCAD RPC Port (default: 9875)")
    parser.add_argument("--auth-token", default=os.environ.get("FREECAD_MCP_TOKEN", None), help="FreeCAD RPC Auth Token")
    parser.add_argument("--freecadcmd", default=None, help="Path or command for headless FreeCADCmd")
    parser.add_argument("--test", action="store_true", help="Run unit & integration test suite")
    args, _ = parser.parse_known_args()

    if args.test:
        import unittest
        loader = unittest.defaultTestLoader
        suite = loader.discover(os.path.dirname(os.path.abspath(__file__)), pattern="test_*.py")
        runner = unittest.TextTestRunner(verbosity=2)
        result = runner.run(suite)
        sys.exit(0 if result.wasSuccessful() else 1)

    state.only_text_feedback = args.only_text_feedback
    state.rpc_host = args.host
    state.rpc_port = args.port
    state.auth_token = args.auth_token
    state.freecadcmd = parse_command(args.freecadcmd)

    logger.info(f"FreeCAD MCP Server initialized. Target RPC: {state.rpc_host}:{state.rpc_port}")
    mcp.run()


if __name__ == "__main__":
    main()
