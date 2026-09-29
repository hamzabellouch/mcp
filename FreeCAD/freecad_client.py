"""
FreeCAD Client Module
---------------------
Robust client for communicating with FreeCAD via XML-RPC (port 9875) and HTTP Bridge (port 9877).
Supports socket timeouts, token authentication, and non-blocking background status polling.
"""

import logging
import math
import xmlrpc.client
from typing import Any, Dict, List, Optional, Tuple, Union

logger = logging.getLogger("FreeCADClient")


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

    def disconnect(self) -> None:
        transport = getattr(self.server, "_ServerProxy__transport", None)
        close = getattr(transport, "close", None)
        if callable(close):
            close()

    def ping(self) -> bool:
        try:
            return bool(self.server.ping())
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

    def create_document(self, name: str = "New_Document") -> Dict[str, Any]:
        return self.server.create_document(name)

    def list_documents(self) -> List[str]:
        return self.server.list_documents()

    def reload_document(self, doc_name: str) -> Dict[str, Any]:
        return self.server.reload_document(doc_name)

    def create_object(self, doc_name: str, obj_data: Dict[str, Any]) -> Dict[str, Any]:
        return self.server.create_object(doc_name, obj_data)

    def edit_object(self, doc_name: str, obj_name: str, obj_data: Dict[str, Any]) -> Dict[str, Any]:
        return self.server.edit_object(doc_name, obj_name, obj_data)

    def delete_object(self, doc_name: str, obj_name: str) -> Dict[str, Any]:
        return self.server.delete_object(doc_name, obj_name)

    def get_objects(self, doc_name: str) -> List[Dict[str, Any]]:
        return self.server.get_objects(doc_name)

    def get_object(self, doc_name: str, obj_name: str) -> Optional[Dict[str, Any]]:
        return self.server.get_object(doc_name, obj_name)

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
        return self.server.execute_code_async(code)

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
        return self.server.get_active_screenshot(view_name, width, height, focus_object)

    def run_fem_analysis(self, doc_name: str, analysis_name: str, timeout: int = 600) -> Dict[str, Any]:
        return self.server.run_fem_analysis(doc_name, analysis_name, timeout)

    def insert_part_from_library(self, relative_path: str) -> Dict[str, Any]:
        return self.server.insert_part_from_library(relative_path)

    def get_parts_list(self) -> Any:
        return self.server.get_parts_list()
