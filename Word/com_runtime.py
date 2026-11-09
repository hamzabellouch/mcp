"""
COM runtime helpers for the Word MCP server.

Responsibilities
----------------
* Detect whether pywin32 is available.
* Resolve the execution mode (``WORD_MCP_MODE``: auto | live | headless).
* Provide dedicated exceptions so callers can tell "Word is simply not running"
  (safe to fall back to editing the file) from "live editing failed"
  (must NOT silently fall back, because that edits the file on disk underneath
  an open Word window and the user never sees the change).
* Run every tool on ONE dedicated thread that owns the COM apartment (STA).
* Retry transient COM rejections ("call was rejected by callee").
* Detect a running WINWORD.EXE process even when COM cannot reach it
  (e.g. elevation mismatch), so we never launch/quit a second instance.

NOTE: stdout is the MCP protocol channel. Nothing in this module may print().
"""

from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import time
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from typing import Any, Callable, Optional

log = logging.getLogger("word_mcp")

try:  # pragma: no cover - exercised only on Windows
    import win32com.client  # noqa: F401
    import pythoncom
    import pywintypes

    COM_AVAILABLE = True
except ImportError:
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]
    pywintypes = None  # type: ignore[assignment]
    COM_AVAILABLE = False


# ----------------------------------------------------------------------
# Exceptions
# ----------------------------------------------------------------------
class LiveWordError(RuntimeError):
    """A live (COM) operation failed. Callers must not fall back to file editing."""


class WordNotRunning(RuntimeError):
    """Word is not running at all. Falling back to headless file editing is safe."""


class WordNotReachable(LiveWordError):
    """A WINWORD.EXE process exists but COM cannot attach to it."""


# ----------------------------------------------------------------------
# Mode
# ----------------------------------------------------------------------
VALID_MODES = ("auto", "live", "headless")


def get_mode() -> str:
    """
    auto     - use live Word when it is running, otherwise edit files directly.
               Live failures are raised, never silently turned into file edits.
    live     - COM only. Never touch files behind Word's back.
    headless - python-docx only. COM is never used.
    """
    mode = os.environ.get("WORD_MCP_MODE", "auto").strip().lower()
    return mode if mode in VALID_MODES else "auto"


def script_execution_enabled() -> bool:
    return os.environ.get("WORD_MCP_ENABLE_SCRIPT", "").strip().lower() in ("1", "true", "yes", "on")


# ----------------------------------------------------------------------
# COM apartment + retry
# ----------------------------------------------------------------------
_tls = threading.local()

# 0x80010001 RPC_E_CALL_REJECTED, 0x8001010A RPC_E_SERVERCALL_RETRYLATER
TRANSIENT_HRESULTS = {-2147418111, -2147417846}


def ensure_com() -> None:
    """Initialise COM once per thread (idempotent)."""
    if COM_AVAILABLE and not getattr(_tls, "initialised", False):
        pythoncom.CoInitialize()
        _tls.initialised = True


def is_transient_com_error(exc: BaseException) -> bool:
    if not COM_AVAILABLE or not isinstance(exc, pywintypes.com_error):
        return False
    return bool(exc.args) and exc.args[0] in TRANSIENT_HRESULTS


def com_retry(fn: Callable[..., Any], *args: Any, tries: int = 8, delay: float = 0.25, **kwargs: Any) -> Any:
    """Call ``fn`` and retry while Word answers 'busy / call rejected'."""
    for attempt in range(tries):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            if attempt == tries - 1 or not is_transient_com_error(exc):
                raise
            time.sleep(min(delay * (1.5 ** attempt), 3.0))


def word_process_exists() -> bool:
    """True if a WINWORD.EXE process exists, regardless of COM reachability."""
    if sys.platform != "win32":
        return False
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
        return "winword.exe" in out.lower()
    except Exception:  # noqa: BLE001
        return False


# ----------------------------------------------------------------------
# Single-thread COM worker
# ----------------------------------------------------------------------
class ComWorker:
    """
    Executes jobs one at a time on a single dedicated thread.

    COM objects are apartment-threaded: they must be created and used on the
    thread that initialised COM. A server that handles concurrent tool calls
    from several threads would otherwise get random failures. Serialising also
    prevents two agent calls from interleaving edits inside one document.
    """

    def __init__(self, default_timeout: Optional[float] = None) -> None:
        self._q: "queue.Queue[Optional[tuple]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._start_lock = threading.Lock()
        if default_timeout is None:
            try:
                default_timeout = float(os.environ.get("WORD_MCP_TIMEOUT", "120"))
            except ValueError:
                default_timeout = 120.0
        self.default_timeout = default_timeout

    def _loop(self) -> None:
        ensure_com()
        try:
            while True:
                item = self._q.get()
                if item is None:
                    break
                fut, fn, args, kwargs = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn(*args, **kwargs))
                except BaseException as exc:  # noqa: BLE001
                    fut.set_exception(exc)
        finally:
            if COM_AVAILABLE and getattr(_tls, "initialised", False):
                try:
                    pythoncom.CoUninitialize()
                except Exception:  # noqa: BLE001
                    pass
                _tls.initialised = False

    def _ensure_started(self) -> None:
        with self._start_lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, name="word-com-worker", daemon=True)
                self._thread.start()

    def submit(self, fn: Callable[..., Any], *args: Any, _timeout: Optional[float] = None, **kwargs: Any) -> Any:
        self._ensure_started()
        fut: Future = Future()
        self._q.put((fut, fn, args, kwargs))
        timeout = self.default_timeout if _timeout is None else _timeout
        try:
            return fut.result(timeout=timeout)
        except FutureTimeout as exc:
            raise TimeoutError(
                f"Word did not finish within {timeout:.0f}s. A modal dialog may be open in Word "
                "(save prompt, password, repair dialog). Close it and retry."
            ) from exc

    def shutdown(self) -> None:
        if self._thread and self._thread.is_alive():
            self._q.put(None)
            self._thread.join(timeout=5)


WORKER = ComWorker()
