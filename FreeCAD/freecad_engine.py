"""
FreeCAD Engine Module
---------------------
Unified orchestrator for communicating with FreeCAD.
Coordinates:
1. Live XML-RPC Connection (Port 9875)
2. Live HTTP Bridge Connection (Port 9877)
3. Headless FreeCADCmd Runner (Subprocess with crash isolation)
4. Dynamic FreeCAD installation & path discovery
"""

import os
import sys
import json
import logging
from typing import Dict, Any, List, Optional, Tuple, Union

import requests
from freecad_headless import detect_freecadcmd, run_headless, parse_command
from freecad_client import FreeCADConnection

logger = logging.getLogger("FreeCADEngine")


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
        target_token = token or cls.DEFAULT_TOKEN

        if cls._rpc_conn is None:
            cls._rpc_conn = FreeCADConnection(
                host=target_host,
                port=target_port,
                token=target_token
            )
        return cls._rpc_conn

    @classmethod
    def check_rpc_connection(cls, timeout: float = 3.0) -> Dict[str, Any]:
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
    def check_http_bridge(cls, timeout: float = 2.0) -> Dict[str, Any]:
        """Check if FreeCAD HTTP Bridge (Port 9877) is active."""
        url = f"http://{cls.DEFAULT_HOST}:{cls.HTTP_PORT}/status"
        try:
            res = requests.get(url, timeout=timeout)
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
        rpc_stat = cls.check_rpc_connection(timeout=1.5)
        http_stat = cls.check_http_bridge(timeout=1.5)
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
        if cls.check_rpc_connection(timeout=1.0).get("connected"):
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
        if cls.check_http_bridge(timeout=1.0).get("connected"):
            url = f"http://{cls.DEFAULT_HOST}:{cls.HTTP_PORT}/execute"
            try:
                r = requests.post(url, json={"code": script_code}, timeout=timeout)
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
