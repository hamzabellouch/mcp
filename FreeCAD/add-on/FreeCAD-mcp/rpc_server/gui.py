"""
FreeCAD MCP Add-on: GUI Dispatch, View Manager & Commands
=========================================================
Consolidates GUI thread task dispatching, thread-safe health monitoring,
3D active view/camera screenshots, FEM execution, and Qt workbench commands.
"""

from __future__ import annotations

import os
import io
import queue
import threading
import time
import traceback
import tempfile
import itertools
from collections.abc import Callable
from typing import Any

import FreeCAD
import FreeCADGui
import ObjectsFem

try:
    from PySide import QtCore, QtWidgets
except ImportError:
    try:
        from PySide2 import QtCore, QtWidgets
    except ImportError:
        from PySide6 import QtCore, QtWidgets

from .server import load_settings, save_settings, validate_allowed_ips

# ---------------------------------------------------------------------
# 1. Dispatch Health Tracking
# ---------------------------------------------------------------------
"""Thread-safe health state for the FreeCAD GUI dispatch bridge."""

import threading
import time
from collections.abc import Callable
from typing import Any


class DispatchHealth:
    """Track the currently running GUI task and whether it timed out."""

    def __init__(self, clock: Callable[[], float] = time.monotonic):
        self._clock = clock
        self._lock = threading.Lock()
        self._active_task_id = 0
        self._operation = ""
        self._started_at = 0.0
        self._timed_out = False
        self._timeout_seconds = 0.0

    def start(self, task_id: int, operation: str) -> None:
        with self._lock:
            self._active_task_id = task_id
            self._operation = operation
            self._started_at = self._clock()
            self._timed_out = False
            self._timeout_seconds = 0.0

    def finish(self, task_id: int) -> None:
        """Clear state only when the finishing task is still the active one."""
        with self._lock:
            if self._active_task_id != task_id:
                return
            self._active_task_id = 0
            self._operation = ""
            self._started_at = 0.0
            self._timed_out = False
            self._timeout_seconds = 0.0

    def mark_timed_out(self, task_id: int, timeout: float) -> dict[str, Any] | None:
        """Mark a running task as timed out and return the resulting snapshot."""
        with self._lock:
            if self._active_task_id != task_id:
                return None
            self._timed_out = True
            self._timeout_seconds = float(timeout)
            return self._snapshot_locked()

    def snapshot(self) -> dict[str, Any]:
        with self._lock:
            return self._snapshot_locked()

    def rejection(self) -> dict[str, Any] | None:
        """Return an immediate failure while a timed-out task is still running."""
        snapshot = self.snapshot()
        if snapshot["state"] != "stuck":
            return None
        return stuck_failure(snapshot, just_timed_out=False)

    def _snapshot_locked(self) -> dict[str, Any]:
        if self._active_task_id == 0:
            return {
                "state": "healthy",
                "task_id": 0,
                "operation": "",
                "running_for_seconds": 0.0,
                "timeout_seconds": 0.0,
            }

        elapsed = max(0.0, self._clock() - self._started_at)
        return {
            "state": "stuck" if self._timed_out else "busy",
            "task_id": self._active_task_id,
            "operation": self._operation,
            "running_for_seconds": round(elapsed, 3),
            "timeout_seconds": self._timeout_seconds,
        }


def stuck_failure(
    snapshot: dict[str, Any], *, just_timed_out: bool
) -> dict[str, Any]:
    """Build the structured error returned for an uninterruptible GUI task."""
    operation = snapshot.get("operation") or "unknown GUI operation"
    running_for = float(snapshot.get("running_for_seconds", 0.0))
    timeout = float(snapshot.get("timeout_seconds", 0.0))

    if just_timed_out:
        lead = (
            f"GUI dispatch timed out after {timeout:g}s while '{operation}' was "
            "still running on FreeCAD's GUI thread."
        )
    else:
        lead = (
            f"GUI dispatch unavailable: '{operation}' timed out after "
            f"{timeout:g}s and is still running "
            f"({running_for:.1f}s elapsed)."
        )

    return {
        "success": False,
        "code": "GUI_DISPATCH_STUCK",
        "error": (
            lead
            + " The task cannot be safely cancelled. Future GUI operations "
            "will fail immediately until it finishes; restart FreeCAD if it "
            "remains stuck."
        ),
        "dispatch": snapshot,
    }


# ---------------------------------------------------------------------
# 2. GUI Thread Dispatch Loop
# ---------------------------------------------------------------------
"""GUI-thread task dispatch for the RPC server.

The XML-RPC server runs in its own thread. FreeCAD APIs that touch the GUI
or the document tree must run in the main GUI thread. This module owns the
queue that ferries wrapped callables onto the GUI thread and the helper that
RPC handlers use to invoke them.

Robustness and performance guarantees:

1. Per-call response queues: each ``dispatch_to_gui`` call owns its own
   ``queue.Queue``. A timeout in one call can never corrupt the response for
   a subsequent call.
2. Immediate wake via Qt signal: ``dispatch_to_gui`` emits a signal from the
   RPC thread; the GUI thread processes the task immediately rather than
   waiting for the next 500 ms heartbeat tick. The 500 ms heartbeat is kept
   only as a fallback.
3. Mouse-button guard: ``process_gui_tasks`` skips the current tick while
   mouse buttons are held so MCP tasks cannot interrupt 3D navigation drags,
   for at most ``MOUSE_DEFER_MAX_S`` so a lost mouse-release cannot wedge
   dispatch. A queued call that times out names the guard that held it.
4. Clean shutdown: the ``_SHUTDOWN`` sentinel sets a flag that suppresses the
   ``finally`` reschedule, so ``stop_rpc_server`` actually stops the loop.
5. Exception isolation: exceptions inside a task are caught, logged, and
   returned as error strings; they never kill the dispatch loop.
6. Stuck-task fail-fast: once a task that already started times out, later GUI
   calls fail immediately until that task returns. Status remains available
   through a GUI-independent RPC method.
7. Timeout counts from task start: the GUI thread runs queued tasks one at a
   time, so a call that arrives while another task is running waits in the
   FIFO first. That wait is budgeted separately (``queue_timeout``) and does
   not consume the task's own ``timeout``. Two concurrent ``execute_code``
   calls therefore each get their full run budget instead of the second one
   being reported stuck because the first was slow.
"""

import itertools
import queue
import threading
import time
import traceback
from typing import Any, Callable

import FreeCAD
import FreeCADGui
try:
    from PySide import QtCore, QtWidgets
except ImportError:
    try:
        from PySide2 import QtCore, QtWidgets
    except ImportError:
        from PySide6 import QtCore, QtWidgets




_rpc_request_queue: "queue.Queue[Any]" = queue.Queue()
_SHUTDOWN = object()
_processing = False  # re-entrancy guard: True while process_gui_tasks is draining
_processing_since: float = 0.0  # wall-clock time when _processing became True
_task_ids = itertools.count(1)
_dispatch_health = DispatchHealth()

# Qt can miss a mouse-release when the button is released outside the FreeCAD
# window (drag out of the viewport, release there). QApplication.mouseButtons()
# then reports the button as held indefinitely. An unbounded mouse guard turns
# that into a permanently wedged RPC server: every tick defers, no task ever
# runs, and every call fails with an unexplained timeout. Cap how long we are
# willing to believe a drag is still in progress.
MOUSE_DEFER_MAX_S = 5.0
# None = unarmed. Do not use 0.0 as the sentinel: time.monotonic() may legitimately
# return 0.0, which would re-arm the timer every tick and never reach the cap.
_mouse_defer_since: "float | None" = None
_stale_mouse_warned = False

# Why and when the last tick declined to process, recorded on the GUI thread so
# the RPC thread can report it on timeout without touching Qt off-thread. One
# tuple assignment keeps reason and time consistent without a lock.
_last_defer: "tuple[str, float] | None" = None


def _mouse_guard_should_defer(mouse_down: bool, now: float) -> bool:
    """Whether to skip this tick because mouse buttons are held.

    Defers while a genuine 3D-navigation drag is plausibly in progress, but
    gives up after ``MOUSE_DEFER_MAX_S``: buttons reported held that long
    across an MCP call are far more likely stale Qt state than a real drag,
    and deferring forever is strictly worse than interrupting a drag.
    """
    global _mouse_defer_since, _stale_mouse_warned
    if not mouse_down:
        _mouse_defer_since = None
        _stale_mouse_warned = False
        return False
    if _mouse_defer_since is None:
        _mouse_defer_since = now
        return True
    return (now - _mouse_defer_since) < MOUSE_DEFER_MAX_S


def _warn_stale_mouse_once() -> None:
    """Announce the override once per stuck-button episode."""
    global _stale_mouse_warned
    if _stale_mouse_warned:
        return
    _stale_mouse_warned = True
    FreeCAD.Console.PrintWarning(
        f"MCP RPC: mouse buttons reported held for over {MOUSE_DEFER_MAX_S:.0f}s. "
        "Treating as stale Qt input state and processing queued tasks anyway.\n"
    )


class _WakeSignal(QtCore.QObject):
    """Qt signal bridge for cross-thread GUI-task wakeup.

    Must be created on the GUI thread (``init_waker``). Emitting from the
    RPC thread is safe: Qt delivers the connection with ``QueuedConnection``,
    so the slot always fires in the GUI thread's event loop.
    """
    _sig = QtCore.Signal()

    def __init__(self):
        super().__init__()
        self._sig.connect(self._on_wake, QtCore.Qt.QueuedConnection)

    def wake(self) -> None:
        self._sig.emit()

    def _on_wake(self) -> None:
        process_gui_tasks(reschedule=False)


_waker: "_WakeSignal | None" = None


def init_waker() -> None:
    """Create the wake-signal bridge. Call once from the GUI thread."""
    global _waker
    _waker = _WakeSignal()


def cleanup_waker() -> None:
    """Release the wake-signal bridge on server stop."""
    global _waker
    _waker = None


def _flush_gui_events(delay_ms: int = 20) -> None:
    FreeCADGui.updateGui()
    app = QtWidgets.QApplication.instance()
    if app is None:
        return

    # ExcludeUserInputEvents: skip mouse/keyboard events to avoid re-entrancy
    # with ongoing navigation. ExcludeSocketNotifiers keeps network I/O out.
    flags = (
        QtCore.QEventLoop.ExcludeUserInputEvents
        | QtCore.QEventLoop.ExcludeSocketNotifiers
    )
    app.processEvents(flags, delay_ms)
    if delay_ms > 0:
        QtCore.QThread.msleep(delay_ms)
        app.processEvents(flags, delay_ms)


def process_gui_tasks(reschedule: bool = True) -> None:
    """Drain queued GUI-thread callables and optionally reschedule.

    Skips the current tick when any mouse button is held (e.g., 3D navigation
    drag) or when already executing a task (re-entrancy guard). The guard
    prevents ``doc.recompute()`` or ``processEvents()`` inside a task from
    triggering a nested ``process_gui_tasks`` call that corrupts FreeCAD state.

    ``reschedule=False`` is used by the immediate-wake path so it does not
    start a second heartbeat chain alongside the existing 500 ms one.
    """
    global _processing, _processing_since, _last_defer
    if _processing:
        return  # re-entrant call from processEvents inside a task; skip

    shutdown = False
    try:
        if _rpc_request_queue.empty():
            return  # nothing queued; skip cursor/status-bar churn on idle heartbeat ticks

        now = time.monotonic()
        mouse_down = QtWidgets.QApplication.mouseButtons() != QtCore.Qt.NoButton
        if _mouse_guard_should_defer(mouse_down, now):
            _last_defer = ("mouse buttons held (3D navigation drag)", now)
            return  # user is dragging; defer to next tick
        if mouse_down:
            _warn_stale_mouse_once()  # stale state: fall through and process anyway
        if QtWidgets.QApplication.activePopupWidget() is not None:
            _last_defer = ("a popup or context menu is open in FreeCAD", now)
            return  # context menu or popup open; defer to next tick
        if QtWidgets.QApplication.activeModalWidget() is not None:
            _last_defer = ("a modal dialog is open in FreeCAD", now)
            return  # modal dialog open; defer to next tick

        _last_defer = None
        _processing = True
        _processing_since = time.monotonic()
        app = QtWidgets.QApplication.instance()
        try:
            status_bar = FreeCADGui.getMainWindow().statusBar()
        except Exception:
            status_bar = None

        if app is not None:
            app.setOverrideCursor(QtCore.Qt.WaitCursor)
        if status_bar is not None:
            status_bar.showMessage("MCP: processing…")
        try:
            while not _rpc_request_queue.empty():
                task = _rpc_request_queue.get()
                if task is _SHUTDOWN:
                    shutdown = True
                    return
                try:
                    task()
                except Exception as e:
                    FreeCAD.Console.PrintError(
                        f"MCP RPC: unhandled exception in GUI task: {type(e).__name__}: {e}\n"
                        f"{traceback.format_exc()}"
                    )
        finally:
            if app is not None:
                app.restoreOverrideCursor()
            if status_bar is not None:
                status_bar.clearMessage()
    finally:
        _processing = False
        if not shutdown and reschedule:
            QtCore.QTimer.singleShot(500, process_gui_tasks)


def request_shutdown() -> None:
    """Post the sentinel so the next dispatch tick exits without rescheduling."""
    _rpc_request_queue.put(_SHUTDOWN)


def get_dispatch_status() -> dict[str, Any]:
    """Return GUI dispatch health without touching FreeCAD's GUI thread."""
    return _dispatch_health.snapshot()


def dispatch_to_gui(
    task: Callable[[], Any],
    timeout: float = 60,
    operation_name: str | None = None,
    queue_timeout: float | None = None,
) -> Any:
    """Run ``task`` on the GUI thread and return its result.

    Uses a per-call response queue so a timeout in one call never corrupts
    the response for a subsequent call. Wakes the GUI thread immediately via
    a Qt signal instead of waiting for the next 500 ms heartbeat.

    ``timeout`` is the run budget and starts counting when the task actually
    begins on the GUI thread. Time spent queued behind earlier tasks is
    budgeted separately by ``queue_timeout`` (defaults to ``timeout``); if the
    task has not started by then it is cancelled without marking dispatch
    stuck. A call that is already queued when an earlier task becomes stuck
    keeps waiting for its turn; only calls arriving after that point are
    rejected immediately.

    A task already running on the GUI thread cannot be interrupted; if it
    exceeds ``timeout`` it is marked as stuck and subsequent GUI calls fail
    immediately until it returns.

    Returns the task's return value on success, an error string if the task
    raises, or ``{"success": False, "error": ...}`` on timeout.
    """
    rejection = _dispatch_health.rejection()
    if rejection is not None:
        return rejection

    if queue_timeout is None:
        queue_timeout = timeout

    task_id = next(_task_ids)
    operation = operation_name or getattr(task, "__name__", "GUI operation")
    if operation == "<lambda>":
        operation = "GUI operation"

    response_queue: "queue.Queue[Any]" = queue.Queue(maxsize=1)
    state_lock = threading.Lock()
    started_event = threading.Event()
    started_at: float | None = None
    cancelled = False

    def _wrapped() -> None:
        nonlocal started_at
        with state_lock:
            if cancelled:
                return  # caller timed out and went away; don't run a stale task
            started_at = time.monotonic()
            _dispatch_health.start(task_id, operation)
            started_event.set()
        missing = object()
        res = missing
        try:
            try:
                res = task()
            except Exception as e:
                FreeCAD.Console.PrintError(
                    f"MCP RPC: GUI task raised {type(e).__name__}: {e}\n"
                    f"{traceback.format_exc()}"
                )
                res = f"{type(e).__name__}: {e}"
        finally:
            # Publish completion atomically with clearing health, so a deadline
            # racing with completion cannot report a missing successful result.
            with state_lock:
                _dispatch_health.finish(task_id)
                if res is not missing:
                    response_queue.put_nowait(res)

    queued_at = time.monotonic()
    _rpc_request_queue.put(_wrapped)
    if _waker is not None:
        _waker.wake()  # immediate wake via Qt signal (thread-safe)

    # Phase 1: wait for the task to start. Earlier queued tasks run first on
    # the GUI thread; that wait must not eat into this task's run budget.
    queue_deadline = queued_at + queue_timeout
    if not started_event.wait(max(0, queue_deadline - time.monotonic())):
        with state_lock:
            cancelled = started_at is None
    if cancelled:
        queued_for = time.monotonic() - queued_at
        last_defer = _last_defer
        if _processing:
            busy_for = time.monotonic() - _processing_since
            hint = (
                f" (GUI thread has been busy for {busy_for:.1f}s — for heavy OCCT"
                " geometry consider execute_code_async, which must apply document"
                " writes through its commit() helper)"
            )
        elif last_defer is not None and last_defer[1] >= queued_at:
            # Never silently time out on a guard: name it so the next wedge
            # diagnoses itself instead of looking like a dead server. A guard
            # that held only an earlier task says nothing about this one.
            hint = (
                f" (GUI thread is not processing tasks: {last_defer[0]}; "
                f"{_rpc_request_queue.qsize()} task(s) queued)"
            )
        else:
            hint = ""
        return {
            "success": False,
            "error": (
                f"GUI dispatch gave up after {queued_for:.1f}s waiting for "
                f"'{operation}' to start (queue_timeout={queue_timeout:g}s){hint}"
            ),
        }

    # Phase 2: count from actual GUI start, even if this RPC thread woke late.
    assert started_at is not None
    run_remaining = max(0, started_at + timeout - time.monotonic())
    try:
        return response_queue.get(timeout=run_remaining)
    except queue.Empty:
        with state_lock:
            # Completion may have won the race while we acquired the lock.
            try:
                return response_queue.get_nowait()
            except queue.Empty:
                stuck = _dispatch_health.mark_timed_out(task_id, timeout)
                if stuck is not None:
                    return stuck_failure(stuck, just_timed_out=True)
                return {"success": False, "error": f"GUI dispatch timed out after {timeout}s"}


# ---------------------------------------------------------------------
# 3. Active 3D View & Screenshot Management
# ---------------------------------------------------------------------
"""Active-view orientation, sizing, and screenshot capture."""

from typing import Any

import FreeCAD
import FreeCADGui




_VIEW_DISPATCH = {
    "Isometric": "viewIsometric",
    "Front": "viewFront",
    "Top": "viewTop",
    "Right": "viewRight",
    "Back": "viewBack",
    "Left": "viewLeft",
    "Bottom": "viewBottom",
    "Dimetric": "viewDimetric",
    "Trimetric": "viewTrimetric",
}


def _get_view_size(view: Any) -> tuple[int, int]:
    try:
        size = view.getSize()
        if isinstance(size, (list, tuple)) and len(size) >= 2:
            return max(1, int(size[0])), max(1, int(size[1]))
        return max(1, int(size.width())), max(1, int(size.height()))
    except Exception:
        return 1024, 768


# Longest edge used when the caller does not ask for a specific size. The
# screenshot's cost to an LLM client scales with its pixel count, and hosts
# commonly downscale anything larger than ~1.5k px before the model ever sees
# it, so rendering at the full window size just inflates the payload. An
# explicit width/height is always honoured as given.
MAX_AUTO_SCREENSHOT_EDGE = 1024


def _scale_to_max_edge(width: int, height: int, max_edge: int) -> tuple[int, int]:
    longest = max(width, height)
    if longest <= max_edge:
        return width, height
    scale = max_edge / longest
    return max(1, int(width * scale)), max(1, int(height * scale))


def _resolve_screenshot_size(
    view: Any,
    width: int | None,
    height: int | None,
) -> tuple[int, int]:
    view_width, view_height = _get_view_size(view)
    if width is None and height is None:
        return _scale_to_max_edge(view_width, view_height, MAX_AUTO_SCREENSHOT_EDGE)
    resolved_width = view_width if width is None else max(1, int(width))
    resolved_height = view_height if height is None else max(1, int(height))
    return resolved_width, resolved_height


_STD_COMMAND_DISPATCH = {
    "Isometric": "Std_ViewIsometric",
    "Front": "Std_ViewFront",
    "Top": "Std_ViewTop",
    "Right": "Std_ViewRight",
    "Back": "Std_ViewRear",
    "Left": "Std_ViewLeft",
    "Bottom": "Std_ViewBottom",
    "Dimetric": "Std_ViewDimetric",
    "Trimetric": "Std_ViewTrimetric",
}


def apply_view_orientation(view: Any, view_name: str) -> None:
    method_name = _VIEW_DISPATCH.get(view_name)
    if method_name is None:
        raise ValueError(f"Invalid view name: {view_name}")
    if hasattr(view, method_name):
        getattr(view, method_name)()
    else:
        # Fallback for views that lack the direct Python method
        # (e.g. some FreeCAD versions / view types)
        cmd = _STD_COMMAND_DISPATCH.get(view_name)
        if cmd:
            FreeCADGui.runCommand(cmd)
        else:
            FreeCAD.Console.PrintWarning(
                f"apply_view_orientation: no method or command for '{view_name}'\n"
            )


def save_active_screenshot(
    save_path: str,
    view_name: str = "Isometric",
    width: int | None = None,
    height: int | None = None,
    focus_object: str | None = None,
):
    """Save a PNG of the active view to ``save_path``.

    Returns ``True`` on success, or an error string on failure (preserves the
    legacy GUI-handler return contract).
    """
    try:
        view = FreeCADGui.ActiveDocument.ActiveView
        if not hasattr(view, "saveImage"):
            return "Current view does not support screenshots"

        apply_view_orientation(view, view_name)

        focused_selection = False
        # The resolved object we frame on (when focus_object is given), kept so
        # the framing can be re-applied synchronously right before saveImage().
        focus_target = None

        if focus_object:
            doc = FreeCAD.ActiveDocument
            obj = doc.getObject(focus_object) if doc else None
            if obj:
                FreeCADGui.Selection.clearSelection()
                FreeCADGui.Selection.addSelection(obj)
                FreeCADGui.SendMsgToActiveView("ViewSelection")
                focused_selection = True
                focus_target = obj
                _flush_gui_events()
                FreeCADGui.Selection.clearSelection()
            else:
                view.fitAll()
        else:
            view.fitAll()

        _flush_gui_events()
        # On macOS, when the FreeCAD window is not exposed (fully occluded or
        # minimized), saveImage() right after pumping the event loop grabs a blank
        # frame. Re-issuing the framing synchronously forces a redraw first. The
        # flush above is kept intentionally — Linux needs it for the stale-frame
        # fix (#51/#53).
        if focused_selection and focus_target is not None:
            FreeCADGui.Selection.addSelection(focus_target)
            FreeCADGui.SendMsgToActiveView("ViewSelection")
            FreeCADGui.Selection.clearSelection()
        else:
            view.fitAll()
        resolved_width, resolved_height = _resolve_screenshot_size(view, width, height)
        # On Wayland the offscreen GL contexts used by the default saveImage()
        # method render solid black; "Framebuffer" reads back the on-screen GL
        # context and captures correctly (and also works on X11/Windows/macOS).
        # FreeCAD < 1.0 lacks the method argument — fall back to the legacy call.
        try:
            view.saveImage(save_path, resolved_width, resolved_height, "Current", "Framebuffer")
        except TypeError:
            view.saveImage(save_path, resolved_width, resolved_height, "Current")

        if focused_selection:
            FreeCADGui.Selection.clearSelection()
            _flush_gui_events(delay_ms=0)
        return True
    except Exception as e:
        return str(e)


# ---------------------------------------------------------------------
# 4. CalculiX FEM Solver Execution
# ---------------------------------------------------------------------
"""CalculiX-driven FEM analysis execution."""

import tempfile
import traceback

import FreeCAD
import ObjectsFem


def run_fem_analysis(doc_name: str, analysis_name: str) -> dict:
    """Run the CalculiX solver on an existing FEM analysis container.

    Always returns a dict with at least ``success`` and ``error``/result keys
    so the caller can pass it through to the wire response unchanged.
    """
    work_dir = None
    stage = "initialization"
    try:
        stage = "document lookup"
        try:
            doc = FreeCAD.getDocument(doc_name)
        except Exception:
            return {"success": False, "error": f"Document '{doc_name}' not found."}
        analysis = doc.getObject(analysis_name)
        if analysis is None:
            return {"success": False, "error": f"Analysis '{analysis_name}' not found."}
        if analysis.TypeId not in ("Fem::FemAnalysis", "Fem::FemAnalysisPython"):
            return {"success": False, "error": f"'{analysis_name}' is not a FEM analysis (TypeId={analysis.TypeId})."}

        stage = "solver resolution"
        solver = None
        for member in analysis.Group:
            tid = getattr(member, "TypeId", "")
            if "SolverCcx" in tid or "SolverCalculix" in tid:
                solver = member
                break
        if solver is None:
            solver_factory = (
                getattr(ObjectsFem, "makeSolverCalculiXCcxTools", None)
                or getattr(ObjectsFem, "makeSolverCalculixCcxTools", None)
            )
            if solver_factory is None:
                return {"success": False, "error": "ObjectsFem has no Calculix solver factory."}
            solver = solver_factory(doc, "CalculiX")
            analysis.addObject(solver)

        stage = "femtools import"
        from femtools import ccxtools

        stage = "solver setup"
        fea = ccxtools.FemToolsCcx(analysis=analysis, solver=solver)
        fea.update_objects()

        work_dir = tempfile.mkdtemp(prefix="freecad_mcp_fem_")
        fea.setup_working_dir(work_dir)
        fea.setup_ccx()

        stage = "prerequisite check"
        prereq_msg = fea.check_prerequisites()
        if prereq_msg:
            return {"success": False, "error": f"Prerequisites failed: {prereq_msg}", "working_dir": work_dir}

        stage = "solver execution"
        fea.purge_results()
        # FemToolsCcx.run() reports CalculiX failures by returning False rather
        # than raising; None is returned on success in some FreeCAD versions,
        # so only an explicit False is treated as a failure.
        if fea.run() is False:
            return {
                "success": False,
                "error": "CalculiX solver run failed (fea.run() returned False); inspect the .dat/.frd output in working_dir.",
                "working_dir": work_dir,
            }

        stage = "result loading"
        fea.load_results()

        result_obj = None
        for member in analysis.Group:
            if "Result" in getattr(member, "TypeId", "") and hasattr(member, "vonMises"):
                result_obj = member
                break
        if result_obj is None:
            return {"success": False, "error": "Solver ran but no result object was produced.", "working_dir": work_dir}

        stage = "result extraction"
        # vonMises / DisplacementLengths can be None on a degenerate run.
        vm = list(getattr(result_obj, "vonMises", None) or [])
        disp = list(getattr(result_obj, "DisplacementLengths", None) or [])
        doc.recompute()

        return {
            "success": True,
            "result_object": result_obj.Name,
            "node_count": len(vm),
            "max_von_mises_MPa": max(vm) if vm else None,
            "min_von_mises_MPa": min(vm) if vm else None,
            "max_displacement_mm": max(disp) if disp else None,
            "working_dir": work_dir,
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"FEM analysis failed during {stage}: {type(e).__name__}: {e}",
            "traceback": traceback.format_exc(),
            "working_dir": work_dir,
        }


# ---------------------------------------------------------------------
# 5. Workbench Commands & Menus
# ---------------------------------------------------------------------
"""Qt Command classes for the MCP Addon workbench menu.

Defines the five toolbar/menu entries (Start, Stop, Toggle Auto-Start,
Toggle Remote, Configure Allowed IPs).

``register_commands()`` and ``schedule_toggle_sync()`` are invoked from
``rpc_server.py`` at import time to preserve current side-effect behavior.
"""

import FreeCAD
import FreeCADGui
try:
    from PySide import QtWidgets
except ImportError:
    try:
        from PySide2 import QtWidgets
    except ImportError:
        from PySide6 import QtWidgets

from .server import validate_allowed_ips
from .server import load_settings, save_settings


_last_server_status_msg = ""


def _report_server_command(message: str, error: bool = False) -> None:
    """Show command feedback even when the Report View is closed."""
    global _last_server_status_msg
    _last_server_status_msg = message
    printer = FreeCAD.Console.PrintError if error else FreeCAD.Console.PrintMessage
    printer(message + "\n")
    try:
        FreeCADGui.getMainWindow().statusBar().showMessage(message, 15000)
    except Exception:
        # Console feedback remains available if the main window is closing.
        pass


def _apply_action_text_update(running: bool) -> None:
    """Safely update Start_RPC_Server action text, tooltip, and statusTip."""
    try:
        try:
            from PySide6 import QtGui
            QAction = QtGui.QAction
        except ImportError:
            try:
                from PySide2 import QtWidgets
                QAction = QtWidgets.QAction
            except ImportError:
                from PySide import QtGui
                QAction = QtGui.QAction

        mw = FreeCADGui.getMainWindow()
        if not mw:
            return

        from . import server as rpc_server
        bound_info = ""
        if getattr(rpc_server, "rpc_server_instance", None):
            try:
                h, p = rpc_server.rpc_server_instance.server_address
                bound_info = f" at {h}:{p}"
            except Exception:
                bound_info = " on Port 2607"

        if running:
            text = "Stop RPC Server"
            tooltip = "Stop the running FreeCAD MCP RPC Server"
            statustip = _last_server_status_msg or f"RPC Server is active{bound_info} (Click to Stop)"
        else:
            text = "Start RPC Server"
            tooltip = "Start the FreeCAD MCP RPC Server"
            statustip = _last_server_status_msg or "Start the FreeCAD MCP RPC Server on Port 2607"

        for a in mw.findChildren(QAction):
            if a.objectName() == "Start_RPC_Server":
                a.setText(text)
                a.setToolTip(tooltip)
                a.setStatusTip(statustip)
    except Exception as e:
        FreeCAD.Console.PrintWarning(f"[MCP] Failed to update action text: {e}\n")


def _update_server_action_ui(running: bool) -> None:
    """Defer action text update so Qt finishes closing the menu popup without crash."""
    try:
        QtCore.QTimer.singleShot(150, lambda: _apply_action_text_update(running))
    except Exception:
        pass


class StartRPCServerCommand:
    def GetResources(self):
        from . import server as rpc_server
        running = getattr(rpc_server, "rpc_server_instance", None) is not None
        return {
            "MenuText": "Stop RPC Server" if running else "Start RPC Server",
            "ToolTip": "Stop the running FreeCAD MCP RPC Server" if running else "Start the FreeCAD MCP RPC Server",
            "StatusTip": "RPC Server is active on Port 2607" if running else "Start the FreeCAD MCP RPC Server on Port 2607",
        }

    def Activated(self, checked: int = 0) -> None:
        try:
            from . import server as rpc_server  # late import: avoids circular at module load
            if getattr(rpc_server, "rpc_server_instance", None) is not None:
                msg = rpc_server.stop_rpc_server()
                _report_server_command(msg)
                _update_server_action_ui(False)
            else:
                msg = rpc_server.start_rpc_server()
                _report_server_command(msg)
                _update_server_action_ui(True)
        except Exception as exc:
            _report_server_command(f"RPC Server toggle failed: {type(exc).__name__}: {exc}", error=True)

    def IsActive(self):
        return True


class StopRPCServerCommand:
    def GetResources(self):
        return {"MenuText": "Stop RPC Server", "ToolTip": "Stop RPC Server"}

    def Activated(self, checked: int = 0) -> None:
        try:
            from . import server as rpc_server
            msg = rpc_server.stop_rpc_server()
            _update_server_action_ui(False)
        except Exception as exc:
            _report_server_command(f"RPC Server failed to stop: {type(exc).__name__}: {exc}", error=True)
            return
        _report_server_command(msg)

    def IsActive(self):
        return True


class ToggleRemoteConnectionsCommand:
    def GetResources(self):
        settings = load_settings()
        return {
            "MenuText": "Remote Connections",
            "ToolTip": "Enable or disable remote connections for the RPC server.",
            "Checkable": True,
            "Checked": bool(settings.get("remote_enabled", False)),
        }

    def Activated(self, checked=0):
        from . import server as rpc_server
        settings = load_settings()
        settings["remote_enabled"] = bool(checked)
        save_settings(settings)

        if settings["remote_enabled"]:
            allowed_ips = settings.get("allowed_ips", "127.0.0.1")
            FreeCAD.Console.PrintMessage(
                f"Remote connections enabled. Allowed IPs: {allowed_ips}\n"
            )
            if not settings.get("auth_token", ""):
                FreeCAD.Console.PrintWarning(
                    "Remote connections have no auth token configured — anyone on "
                    "an allowed IP can execute code in FreeCAD. Set one via "
                    "'Set Auth Token' in the FreeCAD MCP menu.\n"
                )
        else:
            FreeCAD.Console.PrintMessage("Remote connections disabled.\n")

        if rpc_server.rpc_server_instance:
            FreeCAD.Console.PrintMessage(
                "Restart the RPC server for changes to take effect.\n"
            )

    def IsActive(self):
        return True


class ConfigureAllowedIPsCommand:
    def GetResources(self):
        return {
            "MenuText": "Configure Allowed IPs",
            "ToolTip": "Set which IP addresses or subnets are allowed to connect to the RPC server.",
        }

    def Activated(self):
        from . import server as rpc_server
        settings = load_settings()
        current_ips = settings.get("allowed_ips", "127.0.0.1")
        text, ok = QtWidgets.QInputDialog.getText(
            None,
            "Allowed IP Addresses",
            "Enter allowed IP addresses or subnets (comma-separated):\n"
            "Examples: 127.0.0.1, 192.168.1.0/24, 10.0.0.5",
            QtWidgets.QLineEdit.Normal,
            current_ips,
        )
        if ok and text.strip():
            valid, errors = validate_allowed_ips(text.strip())
            if errors:
                QtWidgets.QMessageBox.warning(
                    None,
                    "Invalid IP Configuration",
                    "The following errors were found:\n\n"
                    + "\n".join(f"• {e}" for e in errors)
                    + ("\n\nOnly valid entries will be saved."
                       if valid else "\n\nNo valid entries found. Settings not changed."),
                )
            if not valid:
                FreeCAD.Console.PrintWarning("Allowed IPs not changed — no valid entries.\n")
                return
            normalised = ", ".join(valid)
            settings["allowed_ips"] = normalised
            save_settings(settings)
            FreeCAD.Console.PrintMessage(
                f"Allowed IPs updated to: {normalised}\n"
            )
            if rpc_server.rpc_server_instance:
                FreeCAD.Console.PrintMessage(
                    "Restart the RPC server for changes to take effect.\n"
                )
        else:
            FreeCAD.Console.PrintMessage("Allowed IPs not changed.\n")

    def IsActive(self):
        return True


class SetAuthTokenCommand:
    def GetResources(self):
        return {
            "MenuText": "Set Auth Token",
            "ToolTip": "Set the shared-secret token clients must present to connect. Blank disables authentication.",
        }

    def Activated(self):
        from . import server as rpc_server
        settings = load_settings()
        current = settings.get("auth_token", "")
        text, ok = QtWidgets.QInputDialog.getText(
            None,
            "Auth Token",
            "Enter the auth token clients must present (leave blank to disable\n"
            "authentication). The MCP server passes it via --auth-token or the\n"
            "FREECAD_MCP_TOKEN environment variable.",
            QtWidgets.QLineEdit.Password,
            current,
        )
        if not ok:
            FreeCAD.Console.PrintMessage("Auth token not changed.\n")
            return
        settings["auth_token"] = text.strip()
        save_settings(settings)
        if settings["auth_token"]:
            FreeCAD.Console.PrintMessage("Auth token set — clients must authenticate.\n")
        else:
            FreeCAD.Console.PrintMessage("Auth token cleared — authentication disabled.\n")
        if rpc_server.rpc_server_instance:
            FreeCAD.Console.PrintMessage(
                "Restart the RPC server for changes to take effect.\n"
            )

    def IsActive(self):
        return True


class ToggleAutoStartCommand:
    def GetResources(self):
        settings = load_settings()
        return {
            "MenuText": "Auto-Start Server",
            "ToolTip": "Automatically start the RPC server when FreeCAD launches.",
            "Checkable": True,
            "Checked": bool(settings.get("auto_start_rpc", True)),
        }

    def Activated(self, checked=0):
        settings = load_settings()
        settings["auto_start_rpc"] = bool(checked)
        save_settings(settings)

        if settings["auto_start_rpc"]:
            FreeCAD.Console.PrintMessage(
                "MCP RPC server will start automatically on next FreeCAD launch.\n"
            )
        else:
            FreeCAD.Console.PrintMessage(
                "MCP RPC server auto-start disabled.\n"
            )

    def IsActive(self):
        return True


def register_commands() -> None:
    if not hasattr(FreeCADGui, "addCommand"):
        return
    FreeCADGui.addCommand("Start_RPC_Server", StartRPCServerCommand())
    FreeCADGui.addCommand("Stop_RPC_Server", StopRPCServerCommand())
    FreeCADGui.addCommand("Toggle_Auto_Start", ToggleAutoStartCommand())
    FreeCADGui.addCommand("Toggle_Remote_Connections", ToggleRemoteConnectionsCommand())
    FreeCADGui.addCommand("Configure_Allowed_IPs", ConfigureAllowedIPsCommand())
    FreeCADGui.addCommand("Set_Auth_Token", SetAuthTokenCommand())


def schedule_toggle_sync() -> None:
    """Compatibility no-op; toggle state is initialized by ``GetResources``.

    FreeCAD treats the presence of the ``Checkable`` resource as making an
    action checkable and uses its boolean value as the action's initial checked
    state. Loading the saved setting in ``GetResources`` therefore avoids any
    delayed QAction lookup or workbench activation at startup.
    """


def populate_mcp_menu(menu) -> None:
    """Populate or refresh the global FreeCAD MCP menu items dynamically."""
    try:
        from . import server as rpc_server

        menu.clear()

        # 1. Start / Stop Action
        running = getattr(rpc_server, "rpc_server_instance", None) is not None
        act_toggle = menu.addAction("Stop RPC Server" if running else "Start RPC Server")
        act_toggle.setObjectName("Start_RPC_Server")
        act_toggle.triggered.connect(lambda: StartRPCServerCommand().Activated())

        menu.addSeparator()

        # 2. Auto-Start
        settings = load_settings()
        act_auto = menu.addAction("Auto-Start Server")
        act_auto.setCheckable(True)
        act_auto.setChecked(bool(settings.get("auto_start_rpc", True)))
        act_auto.triggered.connect(lambda chk: ToggleAutoStartCommand().Activated(1 if chk else 0))

        # 3. Remote Connections
        act_remote = menu.addAction("Remote Connections")
        act_remote.setCheckable(True)
        act_remote.setChecked(bool(settings.get("remote_enabled", False)))
        act_remote.triggered.connect(lambda chk: ToggleRemoteConnectionsCommand().Activated(1 if chk else 0))

        menu.addSeparator()

        # 4. Configure Allowed IPs
        act_ips = menu.addAction("Configure Allowed IPs")
        act_ips.triggered.connect(lambda: ConfigureAllowedIPsCommand().Activated())

        # 5. Set Auth Token
        act_token = menu.addAction("Set Auth Token")
        act_token.triggered.connect(lambda: SetAuthTokenCommand().Activated())

    except Exception as e:
        FreeCAD.Console.PrintError(f"[MCP] Failed to populate menu: {e}\n")


def setup_global_menu() -> None:
    """Install permanent FreeCAD MCP menu into the main window menu bar."""
    try:
        try:
            from PySide6 import QtGui, QtWidgets
        except ImportError:
            try:
                from PySide2 import QtWidgets, QtGui
            except ImportError:
                from PySide import QtGui, QtWidgets

        mw = FreeCADGui.getMainWindow()
        if not mw:
            return

        menubar = mw.menuBar()
        if not menubar:
            return

        # Find existing or create new menu
        mcp_menu = None
        for a in menubar.actions():
            if a.text() == "FreeCAD MCP" or (a.menu() and a.menu().objectName() == "FreeCAD_MCP_Global_Menu"):
                mcp_menu = a.menu()
                break

        if not mcp_menu:
            mcp_menu = QtWidgets.QMenu("FreeCAD MCP", menubar)
            mcp_menu.setObjectName("FreeCAD_MCP_Global_Menu")

            # Insert before Windows or Help menu
            target_act = None
            for a in menubar.actions():
                t = a.text().replace("&", "")
                if t in ("Windows", "Window", "Help"):
                    target_act = a
                    break

            if target_act:
                menubar.insertMenu(target_act, mcp_menu)
            else:
                menubar.addMenu(mcp_menu)

        # Wire up aboutToShow and populate immediately
        try:
            mcp_menu.aboutToShow.disconnect()
        except Exception:
            pass
        mcp_menu.aboutToShow.connect(lambda: populate_mcp_menu(mcp_menu))
        populate_mcp_menu(mcp_menu)

    except Exception as e:
        FreeCAD.Console.PrintWarning(f"[MCP] Failed to install global menu: {e}\n")
