"""
FreeCAD MCP Add-on: Core XML-RPC Server & Security
==================================================
Consolidates protocol versioning, user settings persistence,
IP allowlisting, browser request rejection (anti-CSRF/rebinding),
and the XML-RPC server lifecycle & handler catalog.
"""

from __future__ import annotations

import base64
import contextlib
import hmac
import io
import ipaddress
import json
import math
import os
import re
import tempfile
import threading
import time
import uuid
from collections.abc import Callable
from email.message import Message
from socketserver import ThreadingMixIn
from typing import Any
from xmlrpc.client import Fault
from xmlrpc.server import SimpleXMLRPCRequestHandler, SimpleXMLRPCServer

import FreeCAD
import FreeCADGui

try:
    from PySide import QtCore
except ImportError:
    try:
        from PySide2 import QtCore
    except ImportError:
        from PySide6 import QtCore

# ---------------------------------------------------------------------
# 1. Version Information
# ---------------------------------------------------------------------
"""Addon version reported to the MCP server through get_rpc_status."""

# Keep in step with the freecad-mcp package version in pyproject.toml.
__version__ = "0.1.25"

# Bump when the RPC contract changes in a way the MCP server must know about:
# a method or parameter is added or removed, or a response shape changes.
# Must match PROTOCOL_VERSION in src/freecad_mcp/version.py.
PROTOCOL_VERSION = 1

ADDON_VERSION = __version__

# ---------------------------------------------------------------------
# 2. Settings Persistence
# ---------------------------------------------------------------------
"""Persistence of MCP RPC server settings under FreeCAD's user app data dir."""

import json
import os

import FreeCAD


_SETTINGS_FILENAME = "freecad_mcp_settings.json"

_DEFAULT_SETTINGS = {
    "remote_enabled": False,
    "allowed_ips": "127.0.0.1",
    "auto_start_rpc": True,
    "auth_token": "",  # empty = authentication disabled
}


def _get_settings_path():
    return os.path.join(FreeCAD.getUserAppDataDir(), _SETTINGS_FILENAME)


def load_settings():
    path = _get_settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                settings = json.load(f)
            for key, value in _DEFAULT_SETTINGS.items():
                if key not in settings:
                    settings[key] = value
            return settings
        except Exception as e:
            FreeCAD.Console.PrintWarning(f"Failed to load MCP settings: {e}\n")
    return dict(_DEFAULT_SETTINGS)


def save_settings(settings):
    path = _get_settings_path()
    try:
        with open(path, "w") as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        FreeCAD.Console.PrintError(f"Failed to save MCP settings: {e}\n")


# ---------------------------------------------------------------------
# 3. Security, IP Filter & Anti-CSRF
# ---------------------------------------------------------------------
"""IP-filtered, optionally token-authenticated XML-RPC server and helpers.

The server also refuses requests a web page could have sent. The IP allowlist
cannot stop those: the browser runs on an allowed machine, so a malicious page
could otherwise call execute_code (CSRF, or DNS rebinding to read the reply).
"""

import base64
import hmac
import ipaddress
import re
from email.message import Message
from socketserver import ThreadingMixIn
from xmlrpc.server import SimpleXMLRPCRequestHandler, SimpleXMLRPCServer

import FreeCAD


_XML_MEDIA_TYPES = frozenset({"text/xml", "application/xml"})


def _host_name(host_header: str) -> str:
    """Return the host of a Host header value, without port or brackets."""
    host = host_header.strip().lower()
    if host.startswith("["):  # IPv6 literal, e.g. [::1]:9875
        return host[1:].split("]", 1)[0]
    return host.rsplit(":", 1)[0].rstrip(".")


def _is_loopback_name(name: str) -> bool:
    if name == "localhost":
        return True
    try:
        return ipaddress.ip_address(name).is_loopback
    except ValueError:
        return False


def browser_request_rejection(
    headers: Message, loopback_only: bool
) -> tuple[int, str] | None:
    """Return ``(status, reason)`` for a request to refuse, or None to accept it.

    Browsers attach Origin to every POST, and can send an XML body to another
    origin only after a CORS preflight, which this server never answers.
    xmlrpc.client sends text/xml without Origin, so MCP clients are unaffected.
    The Host check stops DNS rebinding where it is decidable: a server bound to
    loopback is only ever addressed as localhost.
    """
    if headers.get("Origin") is not None:
        return 403, "requests from web pages are not accepted"
    media_type = (headers.get("Content-Type") or "").split(";", 1)[0].strip().lower()
    if media_type not in _XML_MEDIA_TYPES:
        return 415, "XML-RPC requests must use Content-Type text/xml"
    host = headers.get("Host")
    if loopback_only and host is not None and not _is_loopback_name(_host_name(host)):
        return 403, "the local RPC server only accepts requests addressed to localhost"
    return None


class BrowserGuardRequestHandler(SimpleXMLRPCRequestHandler):
    """Refuse requests a web page could have sent before they are dispatched."""

    def do_POST(self) -> None:
        rejection = browser_request_rejection(self.headers, self.server.loopback_only)
        if rejection is None:
            super().do_POST()
            return
        status, reason = rejection
        FreeCAD.Console.PrintWarning(
            f"MCP RPC: Rejected request from {self.client_address[0]}: {reason}\n"
        )
        body = reason.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Connection", "close")  # the unread body must not be parsed as a request
        self.end_headers()
        self.wfile.write(body)


def authorization_ok(header_value: str, token: str) -> bool:
    """Check an ``Authorization`` header against the configured token.

    Accepts ``Bearer <token>`` and HTTP Basic (token in the password field,
    username ignored) so stdlib clients can use ``http://:token@host:port``
    URIs. Comparisons are constant-time.
    """
    if header_value.startswith("Bearer "):
        supplied = header_value[len("Bearer "):].strip()
        return hmac.compare_digest(supplied, token)
    if header_value.startswith("Basic "):
        try:
            decoded = base64.b64decode(header_value[len("Basic "):], validate=True).decode("utf-8")
        except Exception:
            return False
        _, _, password = decoded.partition(":")
        return hmac.compare_digest(password, token)
    return False


class TokenAuthRequestHandler(BrowserGuardRequestHandler):
    """Request handler that enforces the server's auth token when one is set."""

    def parse_request(self):
        if not super().parse_request():
            return False
        token = getattr(self.server, "auth_token", "")
        if not token:
            return True  # authentication disabled
        if authorization_ok(self.headers.get("Authorization", ""), token):
            return True
        FreeCAD.Console.PrintWarning(
            f"MCP RPC: Rejected unauthenticated request from {self.client_address[0]}\n"
        )
        self.send_error(401, "Unauthorized: valid auth token required")
        return False


class FilteredXMLRPCServer(ThreadingMixIn, SimpleXMLRPCServer):
    """XML-RPC server that filters connections by allowed IP addresses/subnets
    and, when a token is configured, requires an Authorization header.

    Threaded so get_rpc_status stays answerable while a wedged GUI task blocks
    another request. Document queries and synchronous modelling handlers
    serialise onto the GUI thread through dispatch_to_gui. The opt-in
    execute_code_async worker retains its existing background execution.

    daemon_threads must stay true — ThreadingMixIn.server_close() joins
    non-daemon request threads, which would make Stop wait out the stuck
    operation.
    """

    daemon_threads = True

    def __init__(self, addr, allowed_ips_str="127.0.0.1", auth_token="", **kwargs):
        self._allowed_networks = _parse_allowed_ips(allowed_ips_str)
        self.auth_token = auth_token or ""
        # Remote clients address the server by its LAN name or IP, so only a
        # loopback-bound server can require a localhost Host header.
        self.loopback_only = _is_loopback_name(str(addr[0]).lower())
        kwargs.setdefault("requestHandler", TokenAuthRequestHandler)
        super().__init__(addr, **kwargs)

    def verify_request(self, request, client_address):
        client_ip = client_address[0]
        try:
            addr = ipaddress.ip_address(client_ip)
            for network in self._allowed_networks:
                if addr in network:
                    return True
        except ValueError:
            pass
        FreeCAD.Console.PrintWarning(
            f"MCP RPC: Rejected connection from {client_ip}\n"
        )
        return False


_COMMA_SEP_RE = re.compile(r"^\s*[^,\s]+(\s*,\s*[^,\s]+)*\s*$")


def validate_allowed_ips(allowed_ips_str):
    """Validate a comma-separated string of IP addresses/subnets.

    Returns a ``(valid, errors)`` tuple.  ``valid`` is a list of normalised
    entry strings that passed validation; ``errors`` is a list of
    human-readable error messages (empty when the input is fully valid).

    Checks performed:
    1. The overall string is well-formed comma-separated (no leading/trailing
       commas, no empty entries between commas, not blank).
    2. Each individual entry is a valid IPv4/IPv6 address or CIDR subnet
       (validated via the stdlib ``ipaddress`` module).
    """
    errors = []

    if not allowed_ips_str or not allowed_ips_str.strip():
        return [], ["Input must not be empty."]

    if not _COMMA_SEP_RE.match(allowed_ips_str):
        return [], [
            "Malformed list — check for leading/trailing commas, "
            "double commas, or missing separators."
        ]

    valid = []
    for entry in allowed_ips_str.split(","):
        entry = entry.strip()
        try:
            ipaddress.ip_network(entry, strict=False)
            valid.append(entry)
        except ValueError:
            errors.append(f"Invalid IP/subnet: '{entry}'")
    return valid, errors


def _parse_allowed_ips(allowed_ips_str):
    """Parse a comma-separated string of IPs/subnets into a list of ip_network objects."""
    valid, errors = validate_allowed_ips(allowed_ips_str)
    for msg in errors:
        FreeCAD.Console.PrintWarning(f"MCP RPC: {msg}, skipping\n")
    return [ipaddress.ip_network(entry, strict=False) for entry in valid]


# ---------------------------------------------------------------------
# 4. XML-RPC Server Implementation & Handler Registration
# ---------------------------------------------------------------------
import FreeCAD
import FreeCADGui

import contextlib
import base64
import io
import math
import os
import tempfile
import threading
import time
import uuid
from collections.abc import Callable
from typing import Any
from xmlrpc.client import Fault

try:
    from PySide import QtCore
except ImportError:
    try:
        from PySide2 import QtCore
    except ImportError:
        from PySide6 import QtCore

from .gui import register_commands, schedule_toggle_sync
from .gui import run_fem_analysis as _run_fem_analysis
from .gui import (
    cleanup_waker,
    dispatch_to_gui,
    get_dispatch_status,
    init_waker,
    process_gui_tasks,
    request_shutdown,
)
# (FilteredXMLRPCServer, validate_allowed_ips defined locally in server.py)
from .geometry import create_object_gui, edit_object_gui
from .geometry import get_parts_list, insert_part_from_library
from .geometry import Object
from .geometry import serialize_object
# (load_settings, save_settings defined locally in server.py)
# (PROTOCOL_VERSION, ADDON_VERSION defined locally in server.py)
from .gui import save_active_screenshot

rpc_server_thread = None
rpc_server_instance = None
_stop_thread = None  # drains shutdown off the GUI thread; see stop_rpc_server

# Persistent namespace for execute_code / execute_code_async. A dedicated dict
# (instead of this module's globals()) keeps user code from shadowing server
# internals like dispatch_to_gui while preserving the documented pattern of
# sharing module-level variables between successive calls.
_EXEC_NAMESPACE: dict[str, Any] = {
    "FreeCAD": FreeCAD,
    "App": FreeCAD,
    "FreeCADGui": FreeCADGui,
    "Gui": FreeCADGui,
}
_async_execution = threading.local()

# Background jobs started by execute_code_async, newest last. Errors raised off
# the GUI thread used to reach only the Report View; the registry lets the
# client read them via get_async_status. Retain all running jobs and bound only
# completed history, in completion order.
_ASYNC_JOBS: dict[str, dict[str, Any]] = {}
_ASYNC_JOBS_LOCK = threading.Lock()
_ASYNC_JOBS_KEEP = 20


def _record_job(job_id: str, **fields: Any) -> None:
    with _ASYNC_JOBS_LOCK:
        job = _ASYNC_JOBS.pop(job_id, {"id": job_id})
        job.update(fields)
        _ASYNC_JOBS[job_id] = job
        finished = [
            key for key, value in _ASYNC_JOBS.items()
            if value.get("state") in {"done", "failed"}
        ]
        for key in finished[:max(0, len(finished) - _ASYNC_JOBS_KEEP)]:
            del _ASYNC_JOBS[key]


def _ok(res) -> bool:
    """True when a GUI-thread handler returned success."""
    return res is True


def _err(res) -> dict:
    """Convert any non-True result (error string or timeout dict) to a failure dict."""
    if isinstance(res, dict):
        return res
    return {"success": False, "error": str(res)}


def _commit_async(fn: Callable[[], Any], timeout: float = 120) -> Any:
    """Run an async script's document/view writes on the GUI thread."""
    if not getattr(_async_execution, "active", False):
        raise RuntimeError("commit() is only available inside execute_code_async workers")
    res = dispatch_to_gui(
        lambda: (fn(),), timeout=timeout, operation_name="async_commit"
    )
    if isinstance(res, tuple):
        return res[0]
    error = _err(res)
    raise RuntimeError(f"commit() failed: {error['error']}")


# Keep one live namespace, including the helper: saved functions retain this
# dictionary as their globals. The thread-local guard prevents a GUI callback
# or synchronous script from waiting on its own GUI thread through commit().
_EXEC_NAMESPACE["commit"] = _commit_async


def _query_on_gui(task: Callable[[], Any], operation: str) -> Any:
    """Preserve query results while reporting dispatch failures as RPC faults."""
    # A tuple distinguishes valid results (including None and empty lists)
    # from the dispatcher's error strings and failure dictionaries.
    res = dispatch_to_gui(lambda: (task(),), operation_name=operation)
    if isinstance(res, tuple):
        return res[0]
    error = _err(res)
    code = error.get("code", "GUI_DISPATCH_FAILED")
    raise Fault(1, f"{code}: {error['error']}")


class FreeCADRPC:
    """RPC server for FreeCAD"""
    TIMEOUT = 60               # generous wait for GUI thread to become free
    EXECUTE_CODE_TIMEOUT = 90  # GUI-thread execution; use execute_code_async for heavy OCCT ops
    # Ceiling for a caller-supplied execute_code timeout. A GUI task cannot be
    # cancelled once started, so an unbounded wait would hide a wedged GUI
    # thread from the caller indefinitely.
    MAX_EXECUTE_CODE_TIMEOUT = 1800

    def ping(self):
        return True

    def get_rpc_status(self) -> dict[str, Any]:
        """Report server and GUI-dispatch health without using the GUI thread."""
        with _ASYNC_JOBS_LOCK:
            running = [j["id"] for j in _ASYNC_JOBS.values() if j.get("state") == "running"]
        return {
            "success": True,
            "rpc_server": "running",
            "gui_dispatch": get_dispatch_status(),
            "async_jobs_running": running,
            "addon_version": ADDON_VERSION,
            "protocol_version": PROTOCOL_VERSION,
            "execute_code_timeout": self.EXECUTE_CODE_TIMEOUT,
            "max_execute_code_timeout": self.MAX_EXECUTE_CODE_TIMEOUT,
        }

    def get_async_status(self, job_id: str = "") -> dict[str, Any]:
        """Report background jobs without using the GUI thread.

        With ``job_id`` returns that job (state ``running``/``done``/``failed``,
        error and traceback when failed). Without it returns all running jobs
        and up to 20 recently finished jobs. History resets when FreeCAD exits.
        """
        with _ASYNC_JOBS_LOCK:
            if job_id:
                job = _ASYNC_JOBS.get(job_id)
                if job is None:
                    return {"success": False, "error": f"unknown async job: {job_id}"}
                return {"success": True, "job": dict(job)}
            return {"success": True, "jobs": [dict(j) for j in _ASYNC_JOBS.values()]}

    def create_document(self, name="New_Document"):
        # The GUI handler reports the document's ACTUAL name — FreeCAD
        # sanitises requested names ("My Doc" -> "My_Doc") and de-duplicates
        # ("Doc" -> "Doc001"); reporting the requested name breaks every
        # follow-up call that uses it.
        res = dispatch_to_gui(
            lambda: self._create_document_gui(name),
            operation_name="create_document",
        )
        if isinstance(res, dict) and res.get("success"):
            return res
        return _err(res)

    def create_object(self, doc_name, obj_data: dict[str, Any]):
        obj = Object(
            name=obj_data.get("Name", "New_Object"),
            type=obj_data["Type"],
            analysis=obj_data.get("Analysis", None),
            properties=obj_data.get("Properties", {}),
        )
        # create_object_gui reports the created object's actual Name (see
        # its docstring) — same sanitise/de-duplicate concern as documents.
        res = dispatch_to_gui(
            lambda: self._create_object_gui(doc_name, obj),
            operation_name="create_object",
        )
        if isinstance(res, dict) and res.get("success"):
            return res
        return _err(res)

    def edit_object(self, doc_name: str, obj_name: str, properties: dict[str, Any]) -> dict[str, Any]:
        obj = Object(
            name=obj_name,
            properties=properties.get("Properties", {}),
        )
        res = dispatch_to_gui(
            lambda: self._edit_object_gui(doc_name, obj),
            operation_name="edit_object",
        )
        if _ok(res):
            return {"success": True, "object_name": obj.name}
        return _err(res)

    def delete_object(self, doc_name: str, obj_name: str):
        res = dispatch_to_gui(
            lambda: self._delete_object_gui(doc_name, obj_name),
            operation_name="delete_object",
        )
        if _ok(res):
            return {"success": True, "object_name": obj_name}
        return _err(res)


    def reload_document(self, doc_name: str) -> dict[str, Any]:
        """Close and re-open a document by name to pick up external file
        changes (e.g. edits made by another process such as `freecadcmd`
        running headlessly). Returns success once the new document is
        loaded from disk.
        """
        res = dispatch_to_gui(
            lambda: self._reload_document_gui(doc_name),
            operation_name="reload_document",
        )
        if _ok(res):
            return {"success": True, "document_name": doc_name}
        return _err(res)

    def run_fem_analysis(self, doc_name: str, analysis_name: str, timeout: int = 600) -> dict[str, Any]:
        """Run the CalculiX solver on an existing Fem::FemAnalysis and return summary results."""
        try:
            timeout_s = int(timeout)
        except (TypeError, ValueError):
            return {"success": False, "error": f"invalid timeout: {timeout!r}"}
        res = dispatch_to_gui(
            lambda: self._run_fem_analysis_gui(doc_name, analysis_name),
            timeout=timeout_s,
            operation_name="run_fem_analysis",
        )
        if isinstance(res, dict):
            return res
        return {"success": False, "error": str(res)}

    def execute_code_async(self, code: str) -> dict[str, Any]:
        """Start code execution in a background thread and return immediately.

        Use for long-running OCCT *geometry* work (fuse/cut/loft on shapes) that
        would otherwise exceed the MCP timeout. The caller should poll a document
        object for completion status (e.g. check SessionState.Label via get_object).

        Thread-safety contract — read before using this method:

        FreeCAD documents and the Coin3D scenegraph are NOT thread-safe. Code run
        here executes off the GUI thread, so it must not touch them directly.
        Assigning ``obj.Shape``, calling ``doc.recompute()``, ``doc.addObject()``,
        ``doc.save()`` or any ``ViewObject`` from this thread races the GUI thread
        and can wedge FreeCAD's event loop, after which the RPC server stops
        answering entirely.

        Safe pattern: build shapes in the background, then hand the document write
        to the GUI thread via the injected ``commit`` helper::

            box = Part.makeBox(10, 10, 10)          # background: fine
            fused = base.fuse(box).removeSplitter()  # background: fine, this is the slow part

            def apply():                             # runs on the GUI thread
                obj.Shape = fused
                doc.recompute()

            commit(apply)                            # blocks until the GUI thread ran it

        ``commit(fn, timeout=...)`` returns ``fn``'s value, or raises RuntimeError
        if the GUI dispatch failed or timed out. The helper persists so saved
        functions can reuse it in later async calls. It may only be called from
        an async worker, not from a synchronous script or GUI callback.
        """
        def _set_status(msg):
            dispatch_to_gui(
                lambda: FreeCADGui.getMainWindow().statusBar().showMessage(msg),
                operation_name="show_async_status",
            )

        def _clear_status():
            # Short timeout: this runs in the worker's finally block, so a wedged
            # or busy GUI thread must not keep the worker alive for the full
            # default dispatch timeout. Losing a status-bar reset is harmless.
            dispatch_to_gui(
                lambda: FreeCADGui.getMainWindow().statusBar().clearMessage(),
                timeout=5,
                operation_name="clear_async_status",
            )

        job_id = f"job-{uuid.uuid4().hex}"
        code_preview = code if len(code) <= 200 else code[:200] + "…"

        def worker() -> None:
            # NOTE: we do NOT redirect sys.stdout here. contextlib.redirect_stdout
            # swaps stdout process-wide, not per-thread, so it would race with the
            # GUI thread and other concurrent work. Background code should report
            # via FreeCAD.Console (which is thread-safe) instead.
            # Execute against the live dictionary. Merging a snapshot on exit
            # would restore stale values and lose deletions/concurrent writes.
            _async_execution.active = True
            outcome: dict[str, Any] = {"state": "done"}
            try:
                exec(code, _EXEC_NAMESPACE)
            except BaseException as e:
                # SystemExit/KeyboardInterrupt raised by a worker script must
                # also finish its job record rather than leave it running.
                import traceback as _tb
                outcome = {
                    "state": "failed",
                    "error": f"{type(e).__name__}: {e}",
                    "traceback": _tb.format_exc().rstrip(),
                }
            finally:
                del _async_execution.active
                # Publish the result before best-effort GUI/log cleanup. A busy
                # GUI must not prevent a client from observing script failure.
                _record_job(job_id, finished=time.time(), **outcome)
                try:
                    if outcome["state"] == "done":
                        FreeCAD.Console.PrintMessage("Async code execution completed.\n")
                    else:
                        FreeCAD.Console.PrintError(
                            f"Async code error ({job_id}): {outcome['error']}\n{outcome['traceback']}\n"
                        )
                except Exception:
                    pass
                try:
                    _clear_status()
                except Exception:
                    pass  # never let status cleanup mask or outlive the real work

        _record_job(job_id, state="running", started=time.time(), code=code_preview)
        try:
            _set_status("MCP: running background task…")
            threading.Thread(target=worker, daemon=True).start()
        except Exception as e:
            import traceback as _tb
            error = f"{type(e).__name__}: {e}"
            _record_job(
                job_id, state="failed", finished=time.time(), error=error,
                traceback=_tb.format_exc().rstrip(),
            )
            return {"success": False, "job_id": job_id, "error": error}
        return {
            "success": True,
            "job_id": job_id,
            "message": (
                "Code execution started in background. Document writes "
                "(obj.Shape = ..., recompute, addObject, save, ViewObject) must "
                "go through commit(fn) — direct writes from this thread can wedge "
                "FreeCAD."
            ),
        }

    def execute_code(self, code: str, timeout: Any = None) -> dict[str, Any]:
        """Execute Python code on the GUI thread and wait for the result.

        Runs on the GUI thread so that FreeCAD document operations
        (addObject, recompute, save) are safe and correctly ordered.
        Use execute_code_async for heavy OCCT boolean ops (fuse/cut)
        that would block the GUI thread too long.

        ``timeout`` overrides EXECUTE_CODE_TIMEOUT for this call only, capped at
        MAX_EXECUTE_CODE_TIMEOUT. Raise it for genuinely slow GUI-thread work
        that cannot move off the GUI thread, such as importing or exporting a
        large STEP assembly. Without it such a call reports a timeout while the
        task keeps running, and its result is discarded even though the work
        completes.
        """
        timeout_s = self.EXECUTE_CODE_TIMEOUT
        if timeout is not None:
            try:
                timeout_s = float(timeout)
            except (TypeError, ValueError, OverflowError):
                return {"success": False, "error": f"invalid timeout: {timeout!r}"}
            if isinstance(timeout, bool) or not math.isfinite(timeout_s) or timeout_s <= 0:
                return {"success": False, "error": f"invalid timeout: {timeout!r}"}
            timeout_s = min(timeout_s, self.MAX_EXECUTE_CODE_TIMEOUT)

        output_buffer = io.StringIO()

        def task():
            with contextlib.redirect_stdout(output_buffer):
                exec(code, _EXEC_NAMESPACE)
            return True

        res = dispatch_to_gui(
            task,
            timeout=timeout_s,
            operation_name="execute_code",
        )
        if _ok(res):
            FreeCAD.Console.PrintMessage("Python code executed successfully.\n")
            return {
                "success": True,
                "message": "Python code executed successfully.\nOutput: " + output_buffer.getvalue(),
            }
        # Log the offending code (truncated) to make errors traceable
        code_preview = code if len(code) <= 800 else code[:800] + "\n...(truncated)"
        FreeCAD.Console.PrintError(
            f"Error executing Python code: {res}\n"
            f"--- code ---\n{code_preview}\n--- end ---\n"
        )
        return _err(res)

    def get_objects(self, doc_name: str) -> list[dict[str, Any]]:
        return _query_on_gui(lambda: self._get_objects_gui(doc_name), "get_objects")

    def _get_objects_gui(self, doc_name: str) -> list[dict[str, Any]]:
        # FreeCAD.getDocument raises (not returns None) for an unknown name.
        try:
            doc = FreeCAD.getDocument(doc_name)
        except Exception:
            return []
        return [serialize_object(obj) for obj in doc.Objects]

    def get_object(self, doc_name: str, obj_name: str) -> dict[str, Any] | None:
        return _query_on_gui(
            lambda: self._get_object_gui(doc_name, obj_name), "get_object"
        )

    def _get_object_gui(self, doc_name: str, obj_name: str) -> dict[str, Any] | None:
        # FreeCAD.getDocument raises (not returns None) for an unknown name.
        try:
            doc = FreeCAD.getDocument(doc_name)
        except Exception:
            return None
        obj = doc.getObject(obj_name)
        if obj:
            return serialize_object(obj)
        return None

    def insert_part_from_library(self, relative_path):
        res = dispatch_to_gui(
            lambda: self._insert_part_from_library(relative_path),
            operation_name="insert_part_from_library",
        )
        if _ok(res):
            return {"success": True, "message": "Part inserted from library."}
        return _err(res)

    def list_documents(self) -> list[str]:
        return _query_on_gui(
            lambda: list(FreeCAD.listDocuments().keys()), "list_documents"
        )

    def get_parts_list(self):
        return get_parts_list()

    def get_active_screenshot(
        self,
        view_name: str = "Isometric",
        width: int | None = None,
        height: int | None = None,
        focus_object: str | None = None,
    ) -> str:
        """Get a screenshot of the active view as a base64-encoded PNG string.

        Returns None if the active view does not support screenshots
        (e.g., TechDraw or Spreadsheet workbench).
        """
        fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)

        def task():
            try:
                active_view = FreeCADGui.ActiveDocument.ActiveView
            except Exception:
                return False
            if active_view is None or not hasattr(active_view, "saveImage"):
                view_type = type(active_view).__name__ if active_view is not None else "None"
                FreeCAD.Console.PrintWarning(
                    f"MCP RPC: view type '{view_type}' does not support screenshots\n"
                )
                return False
            return save_active_screenshot(tmp_path, view_name, width, height, focus_object)

        try:
            res = dispatch_to_gui(task, operation_name="get_active_screenshot")
            if _ok(res):
                with open(tmp_path, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            if res is False:
                return None
            FreeCAD.Console.PrintWarning(f"MCP RPC: screenshot failed: {res}\n")
            return None
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def _create_document_gui(self, name):
        doc = FreeCAD.newDocument(name)
        doc.recompute()
        FreeCAD.Console.PrintMessage(f"Document '{doc.Name}' created via RPC.\n")
        return {"success": True, "document_name": doc.Name}

    def _create_object_gui(self, doc_name, obj: Object):
        return create_object_gui(doc_name, obj)

    def _edit_object_gui(self, doc_name: str, obj: Object):
        return edit_object_gui(doc_name, obj)

    def _run_fem_analysis_gui(self, doc_name: str, analysis_name: str):
        return _run_fem_analysis(doc_name, analysis_name)

    def _delete_object_gui(self, doc_name: str, obj_name: str):
        try:
            doc = FreeCAD.getDocument(doc_name)
        except Exception:
            FreeCAD.Console.PrintError(f"Document '{doc_name}' not found.\n")
            return f"Document '{doc_name}' not found.\n"

        try:
            doc.removeObject(obj_name)
            doc.recompute()
            FreeCAD.Console.PrintMessage(f"Object '{obj_name}' deleted via RPC.\n")
            return True
        except Exception as e:
            return str(e)


    def _reload_document_gui(self, doc_name: str):
        if doc_name not in FreeCAD.listDocuments():
            return f"Document '{doc_name}' is not loaded."
        doc = FreeCAD.getDocument(doc_name)
        file_path = doc.FileName
        if not file_path:
            return (
                f"Document '{doc_name}' has no file on disk "
                "(unsaved scratch document); nothing to reload from."
            )
        if not os.path.exists(file_path):
            return f"File for '{doc_name}' not found at {file_path!r}."
        # Close, then reopen from the same file. Reopen preserves the
        # original document name when the file was previously saved
        # under that name.
        FreeCAD.closeDocument(doc_name)
        FreeCAD.openDocument(file_path)
        FreeCAD.Console.PrintMessage(
            f"Document '{doc_name}' reloaded from '{file_path}' via RPC.\n"
        )
        return True

    def _insert_part_from_library(self, relative_path):
        try:
            insert_part_from_library(relative_path)
            return True
        except Exception as e:
            return str(e)

    def _save_active_screenshot(
        self,
        save_path: str,
        view_name: str = "Isometric",
        width: int | None = None,
        height: int | None = None,
        focus_object: str | None = None,
    ):
        return save_active_screenshot(save_path, view_name, width, height, focus_object)


def start_rpc_server(port: int = 2607) -> str:
    global rpc_server_thread, rpc_server_instance

    if rpc_server_instance:
        host, bound_port = rpc_server_instance.server_address
        return f"RPC Server already running at {host}:{bound_port} (PID {os.getpid()})."

    # A previous stop may still be draining an in-flight request off-thread;
    # binding before its server_close() would hit the old socket.
    if _stop_thread is not None and _stop_thread.is_alive():
        _stop_thread.join(timeout=5.0)
        if _stop_thread.is_alive():
            return ("RPC Server is still stopping (a request is draining); "
                    "try again in a few seconds.")

    settings = load_settings()
    remote_enabled = settings.get("remote_enabled", False)
    allowed_ips = settings.get("allowed_ips", "127.0.0.1")
    auth_token = settings.get("auth_token", "")

    if remote_enabled:
        host = "0.0.0.0"
        if not auth_token:
            FreeCAD.Console.PrintWarning(
                "MCP RPC: remote connections are enabled WITHOUT an auth token. "
                "Anyone on an allowed IP can execute code in FreeCAD. "
                "Set a token via 'Set Auth Token' in the FreeCAD MCP menu.\n"
            )
    else:
        host = "127.0.0.1"

    server = FilteredXMLRPCServer(
        (host, port),
        allowed_ips_str=allowed_ips,
        auth_token=auth_token,
        allow_none=True,
        logRequests=False,
    )
    try:
        server.register_instance(FreeCADRPC())
        init_waker()
        QtCore.QTimer.singleShot(500, process_gui_tasks)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
    except Exception:
        # A failed start must not retain a listening socket or appear running.
        # This allows the toolbar command to be retried without restarting FreeCAD.
        server.server_close()
        request_shutdown()
        cleanup_waker()
        raise

    rpc_server_instance = server
    rpc_server_thread = thread
    bound_host, bound_port = server.server_address
    msg = f"RPC Server started at {bound_host}:{bound_port} (PID {os.getpid()})."
    if remote_enabled:
        msg += f" Allowed IPs: {allowed_ips}"
    if auth_token:
        msg += " Auth token required."
    return msg


def stop_rpc_server():
    global rpc_server_instance, rpc_server_thread, _stop_thread

    if not rpc_server_instance:
        return "RPC Server was not running."

    server = rpc_server_instance
    thread = rpc_server_thread
    rpc_server_instance = None
    rpc_server_thread = None

    request_shutdown()
    cleanup_waker()

    def _shutdown_and_close():
        # shutdown() only stops the accept loop; in-flight requests run in
        # their own daemon threads and are not waited for. Kept off the GUI
        # thread so a menu command cannot block the UI. server_close() must
        # always follow, or the listening socket stays bound and Stop -> Start
        # fails with EADDRINUSE.
        try:
            server.shutdown()
            if thread is not None:
                thread.join(timeout=10.0)
                if thread.is_alive():
                    FreeCAD.Console.PrintWarning(
                        "MCP RPC: server thread still draining a request; "
                        "socket closes when it finishes.\n"
                    )
        finally:
            server.server_close()
        FreeCAD.Console.PrintMessage("RPC Server stopped.\n")

    _stop_thread = threading.Thread(target=_shutdown_and_close, daemon=True)
    _stop_thread.start()
    return "RPC Server stopping…"


register_commands()
schedule_toggle_sync()
