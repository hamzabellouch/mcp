"""
Excel MCP Server (Unified Edition)
---------------------------------
Model Context Protocol (MCP) server providing high-level tools to interact with Microsoft Excel.
Supports live desktop Excel automation (via Windows COM/pywin32) and headless file manipulation (via openpyxl).
Consolidated into a single portable script.
"""

from __future__ import annotations

import os
import sys
import json
import math
import datetime
import re
import subprocess
import threading
import queue
import time
import logging
from concurrent.futures import Future, TimeoutError as FutureTimeout
from typing import List, Dict, Any, Optional, Union, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import (
    BarChart,
    LineChart,
    PieChart,
    AreaChart,
    ScatterChart,
    DoughnutChart,
    Reference,
    Series,
)

try:
    import win32com.client
    import pythoncom
    import pywintypes
    COM_AVAILABLE = True
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None
    COM_AVAILABLE = False

try:
    from mcp.server.fastmcp import FastMCP
except Exception:
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except Exception:
        from mcp.server import MCPServer as FastMCP

log = logging.getLogger("excel_mcp")


# ----------------------------------------------------------------------
# Exceptions
# ----------------------------------------------------------------------
class LiveExcelError(RuntimeError):
    """A live (COM) operation failed. Callers must not fall back to file editing."""


class ExcelNotRunning(RuntimeError):
    """Excel is not running at all. Falling back to headless file editing is safe."""


class ExcelNotReachable(LiveExcelError):
    """An EXCEL.EXE process exists but COM cannot attach to it."""


# ----------------------------------------------------------------------
# Mode
# ----------------------------------------------------------------------
VALID_MODES = ("auto", "live", "headless")


def get_mode() -> str:
    """
    auto     - use live Excel when it is running, otherwise edit files directly.
    live     - COM only. Never touch files behind Excel's back.
    headless - openpyxl only. COM is never used.
    """
    mode = os.environ.get("EXCEL_MCP_MODE", "auto").strip().lower()
    return mode if mode in VALID_MODES else "auto"


def script_execution_enabled() -> bool:
    return os.environ.get("EXCEL_MCP_ENABLE_SCRIPT", "").strip().lower() in ("1", "true", "yes", "on")


# ----------------------------------------------------------------------
# COM apartment + retry
# ----------------------------------------------------------------------
_tls = threading.local()

# 0x80010001 RPC_E_CALL_REJECTED, 0x8001010A RPC_E_SERVERCALL_RETRYLATER
TRANSIENT_HRESULTS = {-2147418111, -2147417846}


def ensure_com() -> None:
    """Initialise desktop connection and COM once per thread (idempotent)."""
    if sys.platform == "win32" and not getattr(_tls, "desktop_bound", False):
        h_default = None
        try:
            import ctypes
            user32 = ctypes.windll.user32
            h_default = user32.OpenDesktopW("Default", 0, False, 0x01FF)
            if h_default:
                user32.SetThreadDesktop(h_default)
            _tls.desktop_bound = True
        except Exception:
            pass
        finally:
            if h_default:
                try:
                    import ctypes
                    ctypes.windll.user32.CloseDesktop(h_default)
                except Exception:
                    pass

    if COM_AVAILABLE and not getattr(_tls, "initialised", False):
        try:
            pythoncom.CoInitialize()
            _tls.initialised = True
        except Exception:
            pass


def is_transient_com_error(exc: BaseException) -> bool:
    if not COM_AVAILABLE or not isinstance(exc, pywintypes.com_error):
        return False
    return bool(exc.args) and exc.args[0] in TRANSIENT_HRESULTS


def com_retry(fn: Callable[..., Any], *args: Any, tries: int = 8, delay: float = 0.25, **kwargs: Any) -> Any:
    """Call fn and retry while Excel answers 'busy / call rejected'."""
    for attempt in range(tries):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            if attempt == tries - 1 or not is_transient_com_error(exc):
                raise
            time.sleep(min(delay * (1.5 ** attempt), 3.0))


def excel_process_exists() -> bool:
    """True if an EXCEL.EXE process exists, regardless of COM reachability."""
    if sys.platform != "win32":
        return False
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
        return "excel.exe" in out.lower()
    except Exception:
        return False


# ----------------------------------------------------------------------
# Single-thread COM worker
# ----------------------------------------------------------------------
class ComWorker:
    """
    Executes COM jobs one at a time on a single dedicated STA thread.
    """

    def __init__(self, default_timeout: Optional[float] = None) -> None:
        self._q: "queue.Queue[Optional[tuple]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._start_lock = threading.Lock()
        if default_timeout is None:
            try:
                default_timeout = float(os.environ.get("EXCEL_MCP_TIMEOUT", "120"))
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
                except BaseException as exc:
                    fut.set_exception(exc)
        finally:
            if COM_AVAILABLE and getattr(_tls, "initialised", False):
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                _tls.initialised = False

    def _ensure_started(self) -> None:
        with self._start_lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, name="excel-com-worker", daemon=True)
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
                f"Excel did not finish within {timeout:.0f}s. A modal dialog may be open in Excel "
                "(save prompt, formula error dialog, password). Close it and retry."
            ) from exc

    def shutdown(self) -> None:
        if self._thread and self._thread.is_alive():
            self._q.put(None)
            self._thread.join(timeout=5)


WORKER = ComWorker()


# ==========================================
# Helpers & Color Conversions
# ==========================================

def parse_hex_color(hex_str: str) -> Tuple[int, int, int]:
    """Parse '#RRGGBB' or 'RRGGBB' to (R, G, B) tuple."""
    hex_str = hex_str.strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color: {hex_str}")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return r, g, b


def hex_to_bgr_int(hex_str: str) -> int:
    """Convert hex color to COM BGR integer."""
    r, g, b = parse_hex_color(hex_str)
    return r + (g * 256) + (b * 65536)


def hex_to_argb_hex(hex_str: str) -> str:
    """Convert hex color to ARGB hex string 'FFRRGGBB' for openpyxl."""
    r, g, b = parse_hex_color(hex_str)
    return f"FF{r:02X}{g:02X}{b:02X}"


def split_sheet_and_range(range_str: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Split 'Sheet1!A1:B10' into ('Sheet1', 'A1:B10')."""
    if not range_str:
        return None, None
    if '!' in range_str:
        parts = range_str.split('!', 1)
        sheet = parts[0].strip().strip("'\"").replace("''", "'")
        rng = parts[1].strip() or None
        return (sheet or None), rng
    return None, range_str.strip()


def sanitize_value(v: Any) -> Any:
    """Convert complex objects or datetime to JSON-serializable types."""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return str(v)
    return v


# ==========================================
# Live Excel COM Operations
# ==========================================

class LiveExcelSession:
    """Manages connection to active or spawned Excel COM instance."""

    @staticmethod
    def _same_path(p1: str, p2: str) -> bool:
        if "://" in p1 or "://" in p2:
            return p1.lower() == p2.lower()
        try:
            return os.path.normcase(os.path.abspath(p1)) == os.path.normcase(os.path.abspath(p2))
        except Exception:
            return p1.lower() == p2.lower()

    @staticmethod
    def try_get_active_app():
        """Return running Excel.Application if attached, None if not running, or raise ExcelNotReachable."""
        if not COM_AVAILABLE:
            return None
        ensure_com()
        try:
            return com_retry(win32com.client.GetActiveObject, "Excel.Application")
        except Exception as exc:
            if excel_process_exists():
                raise ExcelNotReachable(
                    "EXCEL.EXE is running, but COM could not attach to it (e.g., modal dialog or elevation mismatch)."
                ) from exc
            return None

    @staticmethod
    def is_running() -> bool:
        try:
            return LiveExcelSession.try_get_active_app() is not None
        except ExcelNotReachable:
            return True

    @staticmethod
    def find_open_workbook(app, filepath: str):
        """Find an already-open workbook matching full normalized filepath."""
        for wb in app.Workbooks:
            try:
                full_name = getattr(wb, "FullName", "") or ""
                if full_name and LiveExcelSession._same_path(full_name, filepath):
                    return wb
            except Exception:
                pass
        return None

    @staticmethod
    def get_app(create_if_missing: bool = False, visible: bool = True):
        if not COM_AVAILABLE:
            raise LiveExcelError("pywin32 (win32com) is not available in this Python environment.")
        ensure_com()
        try:
            app = com_retry(win32com.client.GetActiveObject, "Excel.Application")
            if visible and not app.Visible:
                app.Visible = True
            return app
        except Exception as exc:
            if excel_process_exists():
                raise ExcelNotReachable(
                    "EXCEL.EXE is running, but COM could not attach to it (modal dialog or privilege mismatch)."
                ) from exc
            if create_if_missing:
                app = com_retry(win32com.client.Dispatch, "Excel.Application")
                app.Visible = visible
                return app
            raise ExcelNotRunning(
                "Microsoft Excel is not currently running. Please open Excel or specify a file path."
            ) from exc

    @staticmethod
    def get_workbook(app, filepath: Optional[str] = None):
        if filepath:
            wb = LiveExcelSession.find_open_workbook(app, filepath)
            if wb is not None:
                return wb
            abs_path = os.path.abspath(filepath)
            if not os.path.exists(abs_path):
                raise FileNotFoundError(f"File not found: {abs_path}")
            return com_retry(app.Workbooks.Open, abs_path)
        if app.Workbooks.Count == 0 or app.ActiveWorkbook is None:
            return app.Workbooks.Add()
        return app.ActiveWorkbook

    @staticmethod
    def get_sheet(wb, sheet_name: Optional[str] = None):
        if sheet_name:
            try:
                return wb.Sheets(sheet_name)
            except Exception:
                raise ValueError(f"Sheet '{sheet_name}' not found in workbook '{wb.Name}'.")
        if wb.ActiveSheet is None:
            raise ValueError(f"Workbook '{wb.Name}' has no active worksheet.")
        return wb.ActiveSheet


# ==========================================
# Core Engine Implementation
# ==========================================

class ExcelEngine:
    """Unified engine supporting Live COM Excel and Offline openpyxl."""

    @staticmethod
    def _resolve_live_target(
        filepath: Optional[str] = None,
        visible: bool = True,
    ) -> Tuple[Optional[Any], Optional[Any], bool]:
        """
        Resolve (app, wb, opened_in_live_mode) according to EXCEL_MCP_MODE and filepath.
        Returns (None, None, False) when openpyxl should handle the file directly.
        """
        mode = get_mode()
        if mode == "headless":
            if not filepath:
                raise ValueError("Filepath is required when EXCEL_MCP_MODE='headless'.")
            return None, None, False

        if not COM_AVAILABLE:
            if mode == "live" or not filepath:
                raise LiveExcelError("pywin32 (win32com) is not available in this Python environment.")
            return None, None, False

        if not filepath:
            app = LiveExcelSession.get_app(create_if_missing=False, visible=visible)
            if app.Workbooks.Count == 0 or app.ActiveWorkbook is None:
                raise ExcelNotRunning("Microsoft Excel is running, but no workbook is open.")
            return app, app.ActiveWorkbook, False

        app = LiveExcelSession.try_get_active_app()
        if app is not None:
            open_wb = LiveExcelSession.find_open_workbook(app, filepath)
            if open_wb is not None:
                if visible and not getattr(app, "Visible", True):
                    app.Visible = True
                return app, open_wb, False

        if mode == "live":
            app = app or LiveExcelSession.get_app(create_if_missing=True, visible=visible)
            wb = LiveExcelSession.get_workbook(app, filepath)
            return app, wb, True

        return None, None, False

    @staticmethod
    def _get_openpyxl_sheet(wb: openpyxl.Workbook, sheet_name: Optional[str] = None):
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
            return wb[sheet_name]
        if wb.active is None:
            raise ValueError("Workbook has no active worksheet.")
        return wb.active

    @staticmethod
    def _autofit_openpyxl_columns(ws) -> None:
        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    @staticmethod
    def _finalize_live_wb(app, wb, opened_in_live_mode: bool, save_changes: bool = False) -> None:
        if not opened_in_live_mode or wb is None:
            return
        prev_alerts = getattr(app, "DisplayAlerts", None)
        try:
            if prev_alerts is not None:
                app.DisplayAlerts = False
            wb.Close(SaveChanges=bool(save_changes))
        except Exception:
            pass
        finally:
            if prev_alerts is not None:
                try:
                    app.DisplayAlerts = prev_alerts
                except Exception:
                    pass

    @classmethod
    def get_status(cls) -> Dict[str, Any]:
        """Check Excel runtime status and open workbooks."""
        mode = get_mode()
        status_info = {
            "mode": mode,
            "excel_running": False,
            "excel_reachable": False,
            "com_available": COM_AVAILABLE,
            "version": None,
            "open_workbooks": [],
            "active_workbook": None,
            "active_sheet": None
        }

        if COM_AVAILABLE and mode != "headless":
            try:
                app = LiveExcelSession.try_get_active_app()
                if app is not None:
                    status_info["excel_running"] = True
                    status_info["excel_reachable"] = True
                    status_info["version"] = getattr(app, "Version", "Unknown")

                    wb_count = app.Workbooks.Count
                    workbooks = []
                    for i in range(1, wb_count + 1):
                        try:
                            wb = app.Workbooks(i)
                            sheets = [s.Name for s in wb.Sheets]
                            workbooks.append({
                                "name": wb.Name,
                                "path": wb.FullName,
                                "saved": wb.Saved,
                                "sheets": sheets
                            })
                        except Exception:
                            pass
                    status_info["open_workbooks"] = workbooks

                    if app.ActiveWorkbook:
                        status_info["active_workbook"] = app.ActiveWorkbook.Name
                    if app.ActiveSheet:
                        status_info["active_sheet"] = app.ActiveSheet.Name
            except ExcelNotReachable as exc:
                status_info["excel_running"] = True
                status_info["excel_reachable"] = False
                status_info["error"] = str(exc)
            except Exception:
                pass

        return status_info

    @classmethod
    def create_workbook(
        cls,
        filepath: Optional[str] = None,
        sheet_name: str = "Sheet1",
        headers: Optional[List[str]] = None,
        data: Optional[List[List[Any]]] = None,
        visible: bool = True,
        table_style: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new Excel workbook with optional initial data & styling."""
        mode = get_mode()
        if mode == "headless" and not filepath:
            raise ValueError("Filepath is required when EXCEL_MCP_MODE='headless'.")

        abs_target = os.path.abspath(filepath) if filepath else None
        if COM_AVAILABLE and mode != "headless":
            active_app = LiveExcelSession.try_get_active_app()
            if active_app is not None and abs_target:
                if LiveExcelSession.find_open_workbook(active_app, abs_target) is not None:
                    raise LiveExcelError(
                        f"Workbook '{abs_target}' is already open in Microsoft Excel. "
                        "Close it or modify it via write_range/clear_range instead of overwriting on disk."
                    )

        # Use COM when headless is not active and either no filepath is given, live mode is enforced, or visible is True
        if COM_AVAILABLE and mode != "headless" and (filepath is None or visible or mode == "live"):
            app = None
            wb = None
            was_running = LiveExcelSession.is_running()
            try:
                app = LiveExcelSession.get_app(create_if_missing=True, visible=visible)
                wb = app.Workbooks.Add()
                ws = wb.ActiveSheet
                if sheet_name:
                    try:
                        ws.Name = sheet_name
                    except Exception:
                        pass

                row_idx = 1
                if headers:
                    for col_idx, h in enumerate(headers, 1):
                        ws.Cells(row_idx, col_idx).Value = h
                        ws.Cells(row_idx, col_idx).Font.Bold = True
                    row_idx += 1

                if data:
                    for r in data:
                        for col_idx, val in enumerate(r, 1):
                            if isinstance(val, str) and val.startswith('='):
                                ws.Cells(row_idx, col_idx).Formula = val
                            else:
                                ws.Cells(row_idx, col_idx).Value = val
                        row_idx += 1

                if table_style and headers and data:
                    try:
                        end_col = len(headers)
                        end_row = len(data) + 1
                        end_col_letter = get_column_letter(end_col)
                        tbl_range = ws.Range(f"A1:{end_col_letter}{end_row}")
                        tbl = ws.ListObjects.Add(1, tbl_range, None, 1)  # 1 = xlSrcRange, 1 = xlYes
                        tbl.TableStyle = table_style
                    except Exception:
                        pass

                # Auto-fit columns
                ws.Columns.AutoFit()

                saved_path = None
                if abs_target:
                    os.makedirs(os.path.dirname(abs_target), exist_ok=True)
                    prev_alerts = getattr(app, "DisplayAlerts", None)
                    try:
                        if prev_alerts is not None:
                            app.DisplayAlerts = False
                        wb.SaveAs(abs_target)
                    finally:
                        if prev_alerts is not None:
                            try:
                                app.DisplayAlerts = prev_alerts
                            except Exception:
                                pass
                    saved_path = abs_target

                return {
                    "success": True,
                    "message": f"Workbook created successfully in live Excel (Sheet: '{ws.Name}')" + (f", saved to {saved_path}" if saved_path else " (unsaved in Excel)."),
                    "workbook_name": wb.Name,
                    "sheet_name": ws.Name,
                    "filepath": saved_path
                }
            except Exception as e:
                if wb is not None:
                    cls._finalize_live_wb(app, wb, True, save_changes=False)
                if not was_running and app is not None and not filepath:
                    try:
                        app.Quit()
                    except Exception:
                        pass
                if not filepath or mode == "live":
                    raise e

        # Fallback / Headless creation via openpyxl
        wb = openpyxl.Workbook()
        try:
            ws = wb.active
            ws.title = sheet_name or "Sheet1"

            row_idx = 1
            if headers:
                for col_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=h)
                    cell.font = Font(bold=True)
                row_idx += 1

            if data:
                for r in data:
                    for col_idx, val in enumerate(r, 1):
                        ws.cell(row=row_idx, column=col_idx, value=val)
                    row_idx += 1

            # Apply Table Style if specified
            if table_style and headers and data:
                try:
                    from openpyxl.worksheet.table import Table, TableStyleInfo
                    end_col = len(headers)
                    end_row = len(data) + 1
                    end_col_letter = get_column_letter(end_col)
                    tab_range = f"A1:{end_col_letter}{end_row}"
                    tab = Table(displayName="Table1", ref=tab_range)
                    style_name = table_style if table_style.startswith("TableStyle") else f"TableStyle{table_style}"
                    tab.tableStyleInfo = TableStyleInfo(
                        name=style_name,
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=True
                    )
                    ws.add_table(tab)
                except Exception:
                    pass

            # Auto-adjust column widths
            cls._autofit_openpyxl_columns(ws)

            target_path = os.path.abspath(filepath or "workbook.xlsx")
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            wb.save(target_path)

            return {
                "success": True,
                "message": f"Workbook created and saved via openpyxl to {target_path}",
                "filepath": target_path,
                "sheet_name": ws.title
            }
        finally:
            wb.close()

    @classmethod
    def open_workbook(cls, filepath: str, read_only: bool = False, visible: bool = True) -> Dict[str, Any]:
        """Open an existing workbook."""
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        mode = get_mode()
        if COM_AVAILABLE and mode != "headless":
            try:
                app = LiveExcelSession.get_app(create_if_missing=True, visible=visible)
                existing_wb = LiveExcelSession.find_open_workbook(app, abs_path)
                wb = existing_wb if existing_wb is not None else com_retry(app.Workbooks.Open, abs_path, ReadOnly=read_only)
                sheets = [s.Name for s in wb.Sheets]
                return {
                    "success": True,
                    "message": f"Opened '{wb.Name}' in Microsoft Excel.",
                    "workbook_name": wb.Name,
                    "filepath": abs_path,
                    "sheets": sheets,
                    "active_sheet": wb.ActiveSheet.Name
                }
            except Exception as e:
                if mode == "live":
                    raise e

        wb = openpyxl.load_workbook(abs_path, read_only=read_only)
        try:
            sheets = list(wb.sheetnames)
            active_title = sheets[0] if sheets else None
            try:
                if wb.active is not None:
                    active_title = wb.active.title
            except Exception:
                pass
            return {
                "success": True,
                "message": f"Loaded '{os.path.basename(abs_path)}' via openpyxl.",
                "filepath": abs_path,
                "sheets": sheets,
                "active_sheet": active_title
            }
        finally:
            wb.close()

    @classmethod
    def read_range(
        cls,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        range_address: Optional[str] = None,
        include_formulas: bool = False,
        output_format: str = "json"
    ) -> Dict[str, Any]:
        """Read data/formulas from an Excel range or sheet."""
        parsed_sheet, parsed_rng = split_sheet_and_range(range_address)
        sheet_name = parsed_sheet or sheet_name
        range_address = parsed_rng

        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                rng = ws.Range(range_address) if range_address else ws.UsedRange
                raw_values = rng.Formula if include_formulas else rng.Value

                if raw_values is None:
                    values_matrix = []
                elif not isinstance(raw_values, (tuple, list)):
                    values_matrix = [[sanitize_value(raw_values)]]
                elif len(raw_values) == 0:
                    values_matrix = []
                elif not isinstance(raw_values[0], (tuple, list)):
                    values_matrix = [[sanitize_value(v) for v in raw_values]]
                else:
                    values_matrix = [[sanitize_value(cell) for cell in row] for row in raw_values]

                result = {
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "range": rng.Address.replace('$', ''),
                    "row_count": len(values_matrix),
                    "col_count": len(values_matrix[0]) if values_matrix else 0,
                    "data": values_matrix
                }

                if output_format == "markdown_table":
                    result["markdown"] = cls._to_markdown_table(values_matrix)
                elif output_format == "records":
                    if len(values_matrix) > 1:
                        headers = [str(h) if h not in (None, "") else f"col_{i+1}" for i, h in enumerate(values_matrix[0])]
                        result["records"] = [
                            dict(zip(headers, row)) for row in values_matrix[1:]
                        ]
                    else:
                        result["records"] = []
                return result
            finally:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)

        if not filepath:
            raise ValueError("Filepath is required when Excel is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path, data_only=not include_formulas)
        try:
            ws = cls._get_openpyxl_sheet(wb, sheet_name)

            if range_address:
                cells = ws[range_address]
                if not isinstance(cells, tuple):
                    values_matrix = [[sanitize_value(cells.value)]]
                elif len(cells) == 0:
                    values_matrix = []
                elif not isinstance(cells[0], tuple):
                    if len(cells) > 1 and getattr(cells[0], "column", None) == getattr(cells[1], "column", None):
                        values_matrix = [[sanitize_value(c.value)] for c in cells]
                    else:
                        values_matrix = [[sanitize_value(c.value) for c in cells]]
                else:
                    values_matrix = [[sanitize_value(c.value) for c in row] for row in cells]
            else:
                values_matrix = [
                    [sanitize_value(v) for v in row]
                    for row in ws.iter_rows(values_only=True)
                ]

            safe_max_col = max(int(ws.max_column or 1), 1)
            safe_max_row = max(int(ws.max_row or 1), 1)
            result = {
                "filepath": abs_path,
                "sheet": ws.title,
                "range": range_address or f"A1:{get_column_letter(safe_max_col)}{safe_max_row}",
                "row_count": len(values_matrix),
                "col_count": len(values_matrix[0]) if values_matrix else 0,
                "data": values_matrix
            }

            if output_format == "markdown_table":
                result["markdown"] = cls._to_markdown_table(values_matrix)
            elif output_format == "records":
                if len(values_matrix) > 1:
                    headers = [str(h) if h not in (None, "") else f"col_{i+1}" for i, h in enumerate(values_matrix[0])]
                    result["records"] = [
                        dict(zip(headers, row)) for row in values_matrix[1:]
                    ]
                else:
                    result["records"] = []

            return result
        finally:
            wb.close()

    @classmethod
    def write_range(
        cls,
        data: List[List[Any]],
        start_cell: str = "A1",
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        auto_fit_columns: bool = True
    ) -> Dict[str, Any]:
        """Write 2D list of values or formulas into a sheet."""
        if not data:
            return {"success": False, "message": "No data provided to write."}

        parsed_sheet, parsed_start = split_sheet_and_range(start_cell)
        sheet_name = parsed_sheet or sheet_name
        start_cell = parsed_start or "A1"

        rows_written = len(data)
        cols_written = max((len(r) for r in data), default=0)
        if cols_written == 0:
            return {"success": False, "message": "No non-empty rows provided to write."}

        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                start_rng = ws.Range(start_cell)
                start_row = start_rng.Row
                start_col = start_rng.Column

                for r_idx, row_data in enumerate(data):
                    for c_idx, val in enumerate(row_data):
                        cell = ws.Cells(start_row + r_idx, start_col + c_idx)
                        if isinstance(val, str) and val.startswith('='):
                            cell.Formula = val
                        else:
                            cell.Value = val

                if auto_fit_columns and cols_written > 0:
                    end_col = start_col + cols_written - 1
                    ws.Range(ws.Cells(1, start_col), ws.Cells(1, end_col)).EntireColumn.AutoFit()

                res = {
                    "success": True,
                    "message": f"Successfully wrote {rows_written}x{cols_written} cells to {ws.Name}!{start_cell}",
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "rows_written": rows_written,
                    "cols_written": cols_written
                }
                if opened_in_live:
                    cls._finalize_live_wb(app, wb, True, save_changes=True)
                return res
            except Exception:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                raise

        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path)
        try:
            ws = cls._get_openpyxl_sheet(wb, sheet_name)
            start_cell_obj = ws[start_cell]
            start_row = start_cell_obj.row
            start_col = start_cell_obj.column

            for r_idx, row_data in enumerate(data):
                for c_idx, val in enumerate(row_data):
                    ws.cell(row=start_row + r_idx, column=start_col + c_idx, value=val)

            if auto_fit_columns:
                cls._autofit_openpyxl_columns(ws)

            wb.save(abs_path)
            return {
                "success": True,
                "message": f"Saved {rows_written}x{cols_written} cells to '{abs_path}' (Sheet: '{ws.title}') starting at {start_cell}",
                "filepath": abs_path,
                "sheet": ws.title,
                "rows_written": rows_written,
                "cols_written": cols_written
            }
        finally:
            wb.close()

    @classmethod
    def format_range(
        cls,
        range_address: str,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        font_color: Optional[str] = None,
        bg_color: Optional[str] = None,
        number_format: Optional[str] = None,
        horizontal_align: Optional[str] = None,
        border: Optional[str] = None,
        auto_fit: Optional[bool] = None
    ) -> Dict[str, Any]:
        """Apply rich styling and formatting to a range."""
        parsed_sheet, parsed_rng = split_sheet_and_range(range_address)
        sheet_name = parsed_sheet or sheet_name
        range_address = parsed_rng or "A1"

        if horizontal_align is not None and horizontal_align.lower() not in ("left", "center", "right", "justify", "general"):
            raise ValueError(f"Unsupported horizontal_align: '{horizontal_align}'")

        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                rng = ws.Range(range_address)

                if bold is not None:
                    rng.Font.Bold = bold
                if italic is not None:
                    rng.Font.Italic = italic
                if font_size is not None:
                    rng.Font.Size = font_size
                if font_name is not None:
                    rng.Font.Name = font_name
                if font_color is not None:
                    rng.Font.Color = hex_to_bgr_int(font_color)
                if bg_color is not None:
                    rng.Interior.Color = hex_to_bgr_int(bg_color)
                if number_format is not None:
                    rng.NumberFormat = number_format
                if horizontal_align is not None:
                    align_map = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130, "general": 1}
                    rng.HorizontalAlignment = align_map[horizontal_align.lower()]

                if border is not None:
                    if border in ("all_thin", "all"):
                        for border_id in [7, 8, 9, 10, 11, 12]:  # Left, Top, Bottom, Right, InsideH, InsideV
                            try:
                                rng.Borders(border_id).LineStyle = 1  # Continuous
                                rng.Borders(border_id).Weight = 2     # Thin
                            except Exception:
                                pass
                    elif border == "outline_thick":
                        rng.BorderAround(LineStyle=1, Weight=4)
                    elif border == "bottom_double":
                        rng.Borders(9).LineStyle = -4119  # xlDouble
                    elif border == "none":
                        rng.Borders.LineStyle = 0

                if auto_fit:
                    rng.EntireColumn.AutoFit()

                res = {
                    "success": True,
                    "message": f"Formatted range {range_address} in live Excel ({ws.Name})",
                    "sheet": ws.Name,
                    "range": range_address
                }
                if opened_in_live:
                    cls._finalize_live_wb(app, wb, True, save_changes=True)
                return res
            except Exception:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                raise

        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path)
        try:
            ws = cls._get_openpyxl_sheet(wb, sheet_name)

            cells = ws[range_address]
            if not isinstance(cells, tuple):
                cells_list = [cells]
            elif len(cells) > 0 and not isinstance(cells[0], tuple):
                cells_list = list(cells)
            else:
                cells_list = [c for row in cells for c in row]

            min_row = min((c.row for c in cells_list), default=1)
            max_row = max((c.row for c in cells_list), default=1)
            min_col = min((c.column for c in cells_list), default=1)
            max_col = max((c.column for c in cells_list), default=1)

            for cell in cells_list:
                current_font = cell.font or Font()
                new_font_kwargs = {
                    "name": font_name if font_name else current_font.name,
                    "size": font_size if font_size else current_font.size,
                    "bold": bold if bold is not None else current_font.bold,
                    "italic": italic if italic is not None else current_font.italic,
                    "color": hex_to_argb_hex(font_color) if font_color else current_font.color,
                    "underline": current_font.underline,
                    "strike": current_font.strike,
                }
                cell.font = Font(**new_font_kwargs)

                if bg_color:
                    argb = hex_to_argb_hex(bg_color)
                    cell.fill = PatternFill(start_color=argb, end_color=argb, fill_type="solid")

                if horizontal_align:
                    curr_align = cell.alignment or Alignment()
                    cell.alignment = Alignment(
                        horizontal=horizontal_align.lower(),
                        vertical=curr_align.vertical or "center",
                        wrap_text=curr_align.wrap_text
                    )

                if number_format:
                    cell.number_format = number_format

                if border in ("all_thin", "all"):
                    thin_side = Side(border_style="thin", color="000000")
                    cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                elif border == "bottom_double":
                    curr_b = cell.border or Border()
                    double_side = Side(border_style="double", color="000000")
                    cell.border = Border(left=curr_b.left, right=curr_b.right, top=curr_b.top, bottom=double_side)
                elif border == "outline_thick":
                    curr_b = cell.border or Border()
                    thick_side = Side(border_style="thick", color="000000")
                    cell.border = Border(
                        left=thick_side if cell.column == min_col else curr_b.left,
                        right=thick_side if cell.column == max_col else curr_b.right,
                        top=thick_side if cell.row == min_row else curr_b.top,
                        bottom=thick_side if cell.row == max_row else curr_b.bottom,
                    )
                elif border == "none":
                    cell.border = Border()

            if auto_fit:
                cls._autofit_openpyxl_columns(ws)

            wb.save(abs_path)
            return {
                "success": True,
                "message": f"Formatted range {range_address} in '{abs_path}' (Sheet: '{ws.title}')",
                "filepath": abs_path,
                "sheet": ws.title,
                "range": range_address
            }
        finally:
            wb.close()

    @classmethod
    def clear_range(
        cls,
        range_address: Optional[str] = None,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        clear_type: str = "all"
    ) -> Dict[str, Any]:
        """Clear contents, formatting, or all from a specific range or entire worksheet."""
        parsed_sheet, parsed_rng = split_sheet_and_range(range_address)
        sheet_name = parsed_sheet or sheet_name
        range_address = parsed_rng
        ct = clear_type.lower()

        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                if range_address:
                    target = ws.Range(range_address)
                    target_desc = f"range {range_address}"
                else:
                    target = ws.Cells
                    target_desc = f"entire sheet '{ws.Name}'"

                if ct == "contents":
                    target.ClearContents()
                elif ct == "formats":
                    target.ClearFormats()
                else:
                    target.Clear()
                    if not range_address:
                        ws.Columns.ColumnWidth = 10.0

                res = {
                    "success": True,
                    "message": f"Successfully cleared {target_desc} ({ct}) in live Excel",
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "range": range_address or "ALL"
                }
                if opened_in_live:
                    cls._finalize_live_wb(app, wb, True, save_changes=True)
                return res
            except Exception:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                raise

        if not filepath:
            raise ExcelNotRunning("Excel is not running and no filepath was provided.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path)
        try:
            ws = cls._get_openpyxl_sheet(wb, sheet_name)

            if range_address:
                cells = ws[range_address]
                if not isinstance(cells, tuple):
                    cells = ((cells,),)
                elif len(cells) > 0 and not isinstance(cells[0], tuple):
                    cells = (cells,)
            else:
                cells = ws.iter_rows()

            for row in cells:
                for cell in row:
                    if ct in ("all", "contents"):
                        cell.value = None
                    if ct in ("all", "formats"):
                        cell.font = Font()
                        cell.fill = PatternFill()
                        cell.border = Border()
                        cell.alignment = Alignment()
                        cell.number_format = "General"

            wb.save(abs_path)
            return {
                "success": True,
                "message": f"Successfully cleared {range_address or 'entire sheet'} ({ct}) in '{abs_path}'",
                "filepath": abs_path,
                "sheet": ws.title,
                "range": range_address or "ALL"
            }
        finally:
            wb.close()

    @classmethod
    def add_chart(
        cls,
        chart_type: str,
        data_range: str,
        title: Optional[str] = None,
        x_axis_title: Optional[str] = None,
        y_axis_title: Optional[str] = None,
        position_cell: str = "E2",
        width: int = 450,
        height: int = 280,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert a chart into a worksheet."""
        chart_type_lower = chart_type.lower()
        parsed_sheet, parsed_rng = split_sheet_and_range(data_range)
        sheet_name = parsed_sheet or sheet_name
        data_range = parsed_rng or "A1:B10"

    @classmethod
    def add_chart(
        cls,
        chart_type: str,
        data_range: str,
        title: Optional[str] = None,
        x_axis_title: Optional[str] = None,
        y_axis_title: Optional[str] = None,
        position_cell: str = "E2",
        width: int = 450,
        height: int = 280,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert a chart into a worksheet."""
        chart_type_lower = chart_type.lower()
        parsed_sheet, parsed_rng = split_sheet_and_range(data_range)
        sheet_name = parsed_sheet or sheet_name
        data_range = parsed_rng or "A1:B10"

        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                pos_rng = ws.Range(position_cell)
                left = pos_rng.Left
                top = pos_rng.Top

                chart_shape = ws.Shapes.AddChart2(201, 51, left, top, width, height)
                chart = chart_shape.Chart
                chart.SetSourceData(ws.Range(data_range))

                com_chart_types = {
                    "column_clustered": 51,
                    "column": 51,
                    "bar_clustered": 57,
                    "bar": 57,
                    "line": 4,
                    "pie": 5,
                    "area": 1,
                    "scatter": -4169,
                    "doughnut": -4120
                }
                if chart_type_lower in com_chart_types:
                    chart.ChartType = com_chart_types[chart_type_lower]

                if title:
                    chart.HasTitle = True
                    chart.ChartTitle.Text = title

                if x_axis_title:
                    try:
                        chart.Axes(1).HasTitle = True
                        chart.Axes(1).AxisTitle.Text = x_axis_title
                    except Exception:
                        pass

                if y_axis_title:
                    try:
                        chart.Axes(2).HasTitle = True
                        chart.Axes(2).AxisTitle.Text = y_axis_title
                    except Exception:
                        pass

                res = {
                    "success": True,
                    "message": f"Added {chart_type} chart at {position_cell} in live Excel ({ws.Name})",
                    "sheet": ws.Name,
                    "position": position_cell,
                    "title": title
                }
                if opened_in_live:
                    cls._finalize_live_wb(app, wb, True, save_changes=True)
                return res
            except Exception:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                raise

        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path)
        try:
            ws = cls._get_openpyxl_sheet(wb, sheet_name)

            min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(data_range)
            has_header = (max_row > min_row)

            if "scatter" in chart_type_lower:
                if max_col <= min_col:
                    raise ValueError("Scatter chart requires at least 2 columns (X and Y values).")
                chart = ScatterChart()
                x_ref = Reference(ws, min_col=min_col, min_row=min_row + 1 if has_header else min_row, max_row=max_row)
                for col_idx in range(min_col + 1, max_col + 1):
                    y_ref = Reference(ws, min_col=col_idx, min_row=min_row, max_row=max_row)
                    chart.series.append(Series(y_ref, x_ref, title_from_data=has_header))
            else:
                if "doughnut" in chart_type_lower:
                    chart = DoughnutChart()
                elif "bar" in chart_type_lower:
                    chart = BarChart()
                    chart.type = "bar"
                elif "line" in chart_type_lower:
                    chart = LineChart()
                elif "pie" in chart_type_lower:
                    chart = PieChart()
                elif "area" in chart_type_lower:
                    chart = AreaChart()
                else:
                    chart = BarChart()
                    chart.type = "col"

                if max_col > min_col:
                    data_ref = Reference(ws, min_col=min_col + 1, min_row=min_row, max_col=max_col, max_row=max_row)
                    cats_ref = Reference(ws, min_col=min_col, min_row=min_row + 1 if has_header else min_row, max_row=max_row)
                    chart.add_data(data_ref, titles_from_data=has_header)
                    chart.set_categories(cats_ref)
                else:
                    data_ref = Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
                    chart.add_data(data_ref, titles_from_data=has_header)

            if title:
                chart.title = title
            if x_axis_title and hasattr(chart, 'x_axis') and chart.x_axis is not None:
                chart.x_axis.title = x_axis_title
            if y_axis_title and hasattr(chart, 'y_axis') and chart.y_axis is not None:
                chart.y_axis.title = y_axis_title

            chart.width = width / 28.34  # convert approx pt to cm
            chart.height = height / 28.34

            ws.add_chart(chart, position_cell)
            wb.save(abs_path)

            return {
                "success": True,
                "message": f"Added {chart_type} chart at {position_cell} in '{abs_path}' (Sheet: '{ws.title}')",
                "filepath": abs_path,
                "sheet": ws.title,
                "position": position_cell
            }
        finally:
            wb.close()

    @classmethod
    def add_sheet(cls, sheet_name: str, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Add a new worksheet to the workbook."""
        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                existing_names = {s.Name.lower() for s in wb.Sheets}
                if sheet_name.lower() in existing_names:
                    raise ValueError(f"Sheet '{sheet_name}' already exists in workbook.")
                ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                ws.Name = sheet_name
                res = {
                    "success": True,
                    "message": f"Added worksheet '{sheet_name}' in live Excel workbook '{wb.Name}'.",
                    "sheet_name": sheet_name,
                    "workbook": wb.Name
                }
                if opened_in_live:
                    cls._finalize_live_wb(app, wb, True, save_changes=True)
                return res
            except Exception:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                raise

        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path)
        try:
            if sheet_name in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' already exists in workbook.")
            wb.create_sheet(title=sheet_name)
            wb.save(abs_path)
            return {
                "success": True,
                "message": f"Added worksheet '{sheet_name}' to '{abs_path}'.",
                "sheet_name": sheet_name,
                "filepath": abs_path
            }
        finally:
            wb.close()

    @classmethod
    def list_sheets(cls, filepath: Optional[str] = None) -> Dict[str, Any]:
        """List all sheets and dimensions."""
        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                sheets_info = []
                active_name = wb.ActiveSheet.Name if wb.ActiveSheet else None
                for s in wb.Sheets:
                    used = s.UsedRange
                    sheets_info.append({
                        "name": s.Name,
                        "used_range": used.Address.replace('$', ''),
                        "rows": used.Rows.Count,
                        "cols": used.Columns.Count,
                        "is_active": (s.Name == active_name)
                    })
                return {
                    "workbook": wb.Name,
                    "sheets_count": len(sheets_info),
                    "sheets": sheets_info
                }
            finally:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)

        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path, read_only=True)
        try:
            sheets_info = []
            for name in wb.sheetnames:
                ws = wb[name]
                sheets_info.append({
                    "name": name,
                    "used_range": f"A1:{get_column_letter(ws.max_column or 1)}{ws.max_row or 1}",
                    "rows": ws.max_row or 0,
                    "cols": ws.max_column or 0
                })
            return {
                "filepath": abs_path,
                "sheets_count": len(sheets_info),
                "sheets": sheets_info
            }
        finally:
            wb.close()

    @classmethod
    def export_pdf(
        cls,
        output_pdf_path: str,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        orientation: str = "landscape"
    ) -> Dict[str, Any]:
        """Export worksheet or entire workbook to PDF."""
        if get_mode() == "headless":
            raise LiveExcelError("PDF Export via Excel COM is disabled when EXCEL_MCP_MODE='headless'.")
        if not COM_AVAILABLE:
            raise RuntimeError("PDF Export via Excel requires Microsoft Excel and pywin32.")

        abs_pdf = os.path.abspath(output_pdf_path)
        os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)

        app = None
        wb = None
        was_running = False
        opened_temp_wb = False

        active_app = LiveExcelSession.try_get_active_app()
        if active_app is not None:
            was_running = True
            app = active_app
        else:
            if filepath and not os.path.exists(os.path.abspath(filepath)):
                raise FileNotFoundError(f"File not found: {os.path.abspath(filepath)}")
            app = LiveExcelSession.get_app(create_if_missing=True, visible=False)

        try:
            if filepath:
                existing_wb = LiveExcelSession.find_open_workbook(app, filepath)
                if existing_wb is not None:
                    wb = existing_wb
                else:
                    abs_path = os.path.abspath(filepath)
                    if not os.path.exists(abs_path):
                        raise FileNotFoundError(f"File not found: {abs_path}")
                    wb = com_retry(app.Workbooks.Open, abs_path, ReadOnly=True)
                    opened_temp_wb = True
            else:
                wb = LiveExcelSession.get_workbook(app, None)

            orient_val = 2 if orientation.lower() == "landscape" else 1
            if sheet_name:
                target = LiveExcelSession.get_sheet(wb, sheet_name)
                try:
                    target.PageSetup.Orientation = orient_val
                except Exception:
                    pass
            else:
                target = wb
                for s in wb.Sheets:
                    try:
                        s.PageSetup.Orientation = orient_val
                    except Exception:
                        pass

            # xlTypePDF = 0
            target.ExportAsFixedFormat(0, abs_pdf)

            return {
                "success": True,
                "message": f"Successfully exported to PDF at '{abs_pdf}'",
                "pdf_path": abs_pdf
            }
        finally:
            if opened_temp_wb and wb is not None:
                cls._finalize_live_wb(app, wb, True, save_changes=False)
            if not was_running and app is not None:
                try:
                    app.DisplayAlerts = False
                except Exception:
                    pass
                try:
                    app.Quit()
                except Exception:
                    pass

    @classmethod
    def save_workbook(cls, save_as_path: Optional[str] = None, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Save active or specified workbook."""
        app, wb, opened_in_live = cls._resolve_live_target(filepath=filepath, visible=True)
        if wb is not None:
            try:
                if save_as_path:
                    abs_save = os.path.abspath(save_as_path)
                    os.makedirs(os.path.dirname(abs_save), exist_ok=True)
                    prev_alerts = getattr(app, "DisplayAlerts", None)
                    try:
                        if prev_alerts is not None:
                            app.DisplayAlerts = False
                        wb.SaveAs(abs_save)
                    finally:
                        if prev_alerts is not None:
                            try:
                                app.DisplayAlerts = prev_alerts
                            except Exception:
                                pass
                    return {"success": True, "message": f"Workbook saved as '{abs_save}'", "filepath": abs_save}
                else:
                    wb.Save()
                    return {"success": True, "message": f"Workbook '{wb.Name}' saved successfully.", "filepath": wb.FullName}
            finally:
                cls._finalize_live_wb(app, wb, opened_in_live, save_changes=False)
                    
        if not filepath:
            raise ValueError("Filepath is required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        target = os.path.abspath(save_as_path) if save_as_path else abs_path
        os.makedirs(os.path.dirname(target), exist_ok=True)
        wb = openpyxl.load_workbook(abs_path)
        try:
            wb.save(target)
            return {"success": True, "message": f"Saved file to '{target}'", "filepath": target}
        finally:
            wb.close()

    @classmethod
    def analyze_sheet(
        cls,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        has_headers: bool = True
    ) -> Dict[str, Any]:
        """Generate statistical summary and schema profile of tabular data."""
        data_res = cls.read_range(filepath=filepath, sheet_name=sheet_name, output_format="json")
        matrix = data_res.get("data", [])
        
        if not matrix or len(matrix) == 0:
            return {"message": "Sheet is empty."}
            
        if has_headers and len(matrix) > 1:
            headers = [str(c) if c is not None else f"Column_{i+1}" for i, c in enumerate(matrix[0])]
            rows = matrix[1:]
        else:
            headers = [f"Column_{i+1}" for i in range(len(matrix[0]))]
            rows = matrix
            
        num_rows = len(rows)
        num_cols = len(headers)
        
        col_summaries = {}
        for col_idx, h in enumerate(headers):
            col_vals = [r[col_idx] for r in rows if col_idx < len(r) and r[col_idx] not in (None, "", "null")]
            null_count = num_rows - len(col_vals)
            
            # Numeric analysis
            numeric_vals = []
            for v in col_vals:
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    numeric_vals.append(float(v))
                elif isinstance(v, str):
                    try:
                        parsed_num = float(v.strip().replace(',', '').replace('$', '').replace('%', ''))
                        if math.isfinite(parsed_num):
                            numeric_vals.append(parsed_num)
                    except Exception:
                        pass
                        
            is_numeric = len(numeric_vals) >= (len(col_vals) * 0.8) and len(numeric_vals) > 0
            
            summary = {
                "inferred_type": "Number" if is_numeric else "Text/Mixed",
                "non_null_count": len(col_vals),
                "null_count": null_count,
                "unique_values_count": len(set(str(v) for v in col_vals))
            }
            
            if is_numeric and numeric_vals:
                summary.update({
                    "min": min(numeric_vals),
                    "max": max(numeric_vals),
                    "mean": round(sum(numeric_vals) / len(numeric_vals), 4),
                    "sum": round(sum(numeric_vals), 4)
                })
            else:
                # Sample unique values
                unique_sample = list(set(str(v) for v in col_vals))[:5]
                summary["sample_values"] = unique_sample
                
            col_summaries[h] = summary
            
        return {
            "sheet_name": data_res.get("sheet"),
            "total_rows": num_rows,
            "total_columns": num_cols,
            "headers": headers,
            "columns_analysis": col_summaries,
            "first_3_rows": rows[:3]
        }

    @classmethod
    def execute_script(cls, code: str, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Execute arbitrary Python code with injected Excel and openpyxl contexts."""
        if os.environ.get("EXCEL_MCP_ENABLE_SCRIPT", "").strip().lower() in ("0", "false", "no", "off"):
            raise PermissionError("excel_execute_script is disabled by EXCEL_MCP_ENABLE_SCRIPT.")

        app = None
        wb = None
        ws = None

        if COM_AVAILABLE and get_mode() != "headless":
            try:
                app = LiveExcelSession.try_get_active_app()
                if app is not None:
                    if filepath:
                        wb = LiveExcelSession.find_open_workbook(app, filepath)
                    elif app.Workbooks.Count > 0:
                        wb = app.ActiveWorkbook
                    ws = wb.ActiveSheet if wb else None
            except Exception:
                pass

        local_vars = {
            "app": app,
            "wb": wb,
            "ws": ws,
            "openpyxl": openpyxl,
            "filepath": filepath,
            "engine": cls
        }
        if COM_AVAILABLE:
            local_vars["win32com"] = win32com

        import io
        import contextlib

        stdout_buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(stdout_buf):
                exec(code, globals(), local_vars)
            stdout_content = stdout_buf.getvalue()
            return {
                "success": True,
                "output": stdout_content.strip() or "Script executed successfully with no stdout output."
            }
        except (Exception, SystemExit) as e:
            return {
                "success": False,
                "error": str(e),
                "output": stdout_buf.getvalue().strip()
            }

    @staticmethod
    def _to_markdown_table(matrix: List[List[Any]]) -> str:
        """Helper to format 2D matrix as Markdown table."""
        if not matrix:
            return ""
        def _fmt_cell(val: Any) -> str:
            if val is None:
                return ""
            return str(val).replace("|", "\\|").replace("\r\n", "<br>").replace("\n", "<br>")

        headers = matrix[0]
        rows = matrix[1:]

        header_line = "| " + " | ".join(_fmt_cell(h) for h in headers) + " |"
        sep_line = "| " + " | ".join("---" for _ in headers) + " |"
        row_lines = ["| " + " | ".join(_fmt_cell(c) for c in row) + " |" for row in rows]

        return "\n".join([header_line, sep_line] + row_lines)


def _run(fn, *args, **kwargs):
    """Execute on the dedicated COM worker thread connected to the interactive desktop."""
    return WORKER.submit(fn, *args, **kwargs)


# Initialize MCP Server
mcp = FastMCP("excel-mcp-server")


@mcp.tool()
def excel_status() -> str:
    """
    Check the current status of Microsoft Excel and active workbooks.
    Returns whether Excel is running, COM availability, version, open workbooks, and active sheet.
    """
    status = _run(ExcelEngine.get_status)
    if status.get("excel_running"):
        wb_names = [wb["name"] for wb in status.get("open_workbooks", [])]
        return (
            f"✅ Microsoft Excel (v{status.get('version')}) is RUNNING.\n"
            f"- Open Workbooks ({len(wb_names)}): {', '.join(wb_names) if wb_names else 'None'}\n"
            f"- Active Workbook: {status.get('active_workbook', 'None')}\n"
            f"- Active Sheet: {status.get('active_sheet', 'None')}"
        )
    else:
        com_status = "Available (Excel will be launched if needed)." if status.get("com_available") else "Unavailable (pywin32 not installed or non-Windows OS)."
        return (
            "ℹ️ Microsoft Excel is currently NOT running.\n"
            f"- COM automation: {com_status}\n"
            "- Headless file mode: Active (openpyxl can read/write files directly)."
        )


@mcp.tool()
def excel_create_workbook(
    filepath: Optional[str] = None,
    sheet_name: str = "Sheet1",
    headers: Optional[List[str]] = None,
    data: Optional[List[List[Any]]] = None,
    visible: bool = True,
    table_style: Optional[str] = None
) -> str:
    """
    Create a new Excel workbook with optional initial data, headers, sheet name, and table style.
    
    Args:
        filepath: Optional path where the workbook will be saved (e.g., 'reports/sales.xlsx'). If omitted, creates an unsaved workbook in live Excel.
        sheet_name: Name of the initial sheet (default: 'Sheet1').
        headers: List of column header names (e.g. ['Product', 'Price', 'Quantity', 'Total']).
        data: 2D list of row values/formulas (e.g. [['Apple', 1.5, 10, '=B2*C2'], ['Orange', 2.0, 5, '=B3*C3']]).
        visible: Whether to make Excel visible on screen when opened via COM (default: True).
        table_style: Optional Excel table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1', 'TableStyleDark1').
    """
    try:
        res = _run(
            ExcelEngine.create_workbook,
            filepath=filepath,
            sheet_name=sheet_name,
            headers=headers,
            data=data,
            visible=visible,
            table_style=table_style
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error creating workbook: {str(e)}"


@mcp.tool()
def excel_open_workbook(
    filepath: str,
    read_only: bool = False,
    visible: bool = True
) -> str:
    """
    Open an existing Excel workbook file (.xlsx, .xlsm, .xls, .csv).
    
    Args:
        filepath: Full path to the workbook file.
        read_only: Open the file in read-only mode (default: False).
        visible: Make Excel visible on screen (default: True).
    """
    try:
        res = _run(ExcelEngine.open_workbook, filepath=filepath, read_only=read_only, visible=visible)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error opening workbook: {str(e)}"


@mcp.tool()
def excel_read_range(
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    range_address: Optional[str] = None,
    include_formulas: bool = False,
    output_format: str = "markdown_table"
) -> str:
    """
    Read data and formulas from a worksheet or cell range.
    
    Args:
        filepath: Optional file path. If omitted, reads from active workbook in live Excel.
        sheet_name: Worksheet name (e.g. 'Sheet1', 'Sales'). Defaults to active sheet.
        range_address: Range to read (e.g. 'A1:D10', 'B2', or omit to read entire used range).
        include_formulas: If True, returns formulas (e.g. '=SUM(A1:A5)') instead of evaluated values (default: False).
        output_format: 'markdown_table' (default, easiest for LLM to view), 'json', or 'records'.
    """
    try:
        res = _run(
            ExcelEngine.read_range,
            filepath=filepath,
            sheet_name=sheet_name,
            range_address=range_address,
            include_formulas=include_formulas,
            output_format=output_format
        )
        if output_format == "markdown_table" and "markdown" in res:
            header = f"**Sheet:** `{res['sheet']}` | **Range:** `{res['range']}` ({res['row_count']} rows x {res['col_count']} cols)\n\n"
            return header + res["markdown"]
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error reading range: {str(e)}"


@mcp.tool()
def excel_write_range(
    data: List[List[Any]],
    start_cell: str = "A1",
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    auto_fit_columns: bool = True
) -> str:
    """
    Write a 2D matrix of values and formulas into a worksheet.
    
    Args:
        data: 2D array of rows and columns (e.g. [['Revenue', 1000], ['Cost', 600], ['Profit', '=B1-B2']]).
        start_cell: Top-left cell address to start writing (default: 'A1').
        filepath: Optional file path. If omitted, writes to active workbook in live Excel.
        sheet_name: Optional sheet name (defaults to active sheet).
        auto_fit_columns: Auto-fit column widths after writing (default: True).
    """
    try:
        res = _run(
            ExcelEngine.write_range,
            data=data,
            start_cell=start_cell,
            filepath=filepath,
            sheet_name=sheet_name,
            auto_fit_columns=auto_fit_columns
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error writing range: {str(e)}"


@mcp.tool()
def excel_format_range(
    range_address: str,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    number_format: Optional[str] = None,
    horizontal_align: Optional[str] = None,
    border: Optional[str] = None,
    auto_fit: Optional[bool] = None
) -> str:
    """
    Format a cell range with colors, font styling, borders, alignment, and number formats.
    
    Args:
        range_address: Cell range (e.g. 'A1:D1', 'B2:B20').
        filepath: Optional file path. If omitted, formats active workbook in live Excel.
        sheet_name: Optional worksheet name.
        bold: Set font bold (True / False).
        italic: Set font italic (True / False).
        font_size: Font size in points (e.g. 11, 14, 16).
        font_name: Font family name (e.g. 'Segoe UI', 'Calibri', 'Arial').
        font_color: Hex color string for text (e.g. '#FFFFFF', '#002060', '#C00000').
        bg_color: Hex color string for cell fill (e.g. '#1F497D', '#D9E1F2', '#E2EFDA').
        number_format: Excel format string (e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd', '#,##0').
        horizontal_align: Text alignment ('left', 'center', 'right').
        border: Border style ('all_thin', 'outline_thick', 'bottom_double', 'none').
        auto_fit: Auto-fit column widths (True / False).
    """
    try:
        res = _run(
            ExcelEngine.format_range,
            range_address=range_address,
            filepath=filepath,
            sheet_name=sheet_name,
            bold=bold,
            italic=italic,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            bg_color=bg_color,
            number_format=number_format,
            horizontal_align=horizontal_align,
            border=border,
            auto_fit=auto_fit
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error formatting range: {str(e)}"


@mcp.tool()
def excel_clear_range(
    range_address: Optional[str] = None,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    clear_type: str = "all"
) -> str:
    """
    Clear contents, formatting, or everything from a cell range or entire worksheet.
    
    Args:
        range_address: Optional range to clear (e.g. 'A1:D10', 'B2'). If omitted, clears the entire worksheet.
        filepath: Optional file path. If omitted, clears active workbook in live Excel.
        sheet_name: Optional worksheet name. Defaults to active worksheet.
        clear_type: 'all' (default: clears data, formulas, and formats), 'contents' (clears data only), 'formats' (clears styles only).
    """
    try:
        res = _run(
            ExcelEngine.clear_range,
            range_address=range_address,
            filepath=filepath,
            sheet_name=sheet_name,
            clear_type=clear_type
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error clearing range: {str(e)}"


@mcp.tool()
def excel_add_chart(
    chart_type: str,
    data_range: str,
    title: Optional[str] = None,
    x_axis_title: Optional[str] = None,
    y_axis_title: Optional[str] = None,
    position_cell: str = "E2",
    width: int = 450,
    height: int = 280,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> str:
    """
    Insert a chart into the worksheet based on a data range.
    
    Args:
        chart_type: Type of chart: 'column_clustered', 'bar_clustered', 'line', 'pie', 'area', 'scatter', 'doughnut'.
        data_range: Range containing headers and data for the chart (e.g. 'A1:B10', 'A1:D6').
        title: Title of the chart.
        x_axis_title: Optional X-axis category label.
        y_axis_title: Optional Y-axis value label.
        position_cell: Top-left cell where the chart will be placed (e.g. 'E2', 'G1').
        width: Chart width in points (default: 450).
        height: Chart height in points (default: 280).
        filepath: Optional file path. If omitted, adds chart to active workbook in live Excel.
        sheet_name: Optional worksheet name.
    """
    try:
        res = _run(
            ExcelEngine.add_chart,
            chart_type=chart_type,
            data_range=data_range,
            title=title,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            position_cell=position_cell,
            width=width,
            height=height,
            filepath=filepath,
            sheet_name=sheet_name
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding chart: {str(e)}"


@mcp.tool()
def excel_add_sheet(
    sheet_name: str,
    filepath: Optional[str] = None
) -> str:
    """
    Add a new worksheet to the workbook.
    
    Args:
        sheet_name: Name for the new worksheet.
        filepath: Optional file path. If omitted, adds sheet to active workbook in live Excel.
    """
    try:
        res = _run(ExcelEngine.add_sheet, sheet_name=sheet_name, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding sheet: {str(e)}"


@mcp.tool()
def excel_list_sheets(
    filepath: Optional[str] = None
) -> str:
    """
    List all worksheets in the workbook with their dimensions and row/column counts.
    
    Args:
        filepath: Optional file path. If omitted, lists sheets from active workbook in live Excel.
    """
    try:
        res = _run(ExcelEngine.list_sheets, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error listing sheets: {str(e)}"


@mcp.tool()
def excel_export_pdf(
    output_pdf_path: str,
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    orientation: str = "landscape"
) -> str:
    """
    Export the worksheet or entire workbook to a PDF file.
    
    Args:
        output_pdf_path: Destination file path for the generated PDF (e.g. 'E:/reports/report.pdf').
        filepath: Optional source Excel file path. If omitted, exports active workbook from live Excel.
        sheet_name: Optional specific sheet to export. If omitted, exports the active sheet or workbook.
        orientation: Page orientation ('landscape' or 'portrait'). Default: 'landscape'.
    """
    try:
        res = _run(
            ExcelEngine.export_pdf,
            output_pdf_path=output_pdf_path,
            filepath=filepath,
            sheet_name=sheet_name,
            orientation=orientation
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error exporting PDF: {str(e)}"


@mcp.tool()
def excel_save_workbook(
    save_as_path: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Save the currently active workbook or save it to a new file path.
    
    Args:
        save_as_path: Optional new destination path to save as.
        filepath: Optional specific workbook file to save.
    """
    try:
        res = _run(ExcelEngine.save_workbook, save_as_path=save_as_path, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error saving workbook: {str(e)}"


@mcp.tool()
def excel_analyze_sheet(
    filepath: Optional[str] = None,
    sheet_name: Optional[str] = None,
    has_headers: bool = True
) -> str:
    """
    Generate an in-depth statistical analysis, column schema profile, and data distribution of a worksheet.
    
    Args:
        filepath: Optional file path. If omitted, analyzes active worksheet in live Excel.
        sheet_name: Optional sheet name.
        has_headers: Whether the first row contains column headers (default: True).
    """
    try:
        res = _run(
            ExcelEngine.analyze_sheet,
            filepath=filepath,
            sheet_name=sheet_name,
            has_headers=has_headers
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error analyzing sheet: {str(e)}"


@mcp.tool()
def excel_execute_script(
    code: str,
    filepath: Optional[str] = None
) -> str:
    """
    Execute arbitrary Python automation code with direct access to Excel and OpenPyXL objects.
    
    Injected variables available in code:
      - `app`: Excel.Application COM instance (win32com.client)
      - `wb`: Active Workbook COM instance
      - `ws`: Active Worksheet COM instance
      - `win32com`: win32com.client module
      - `openpyxl`: openpyxl module
      - `engine`: ExcelEngine class
      - `filepath`: target filepath if provided
      
    Args:
        code: Python script string to execute.
        filepath: Optional target workbook file path.
    """
    try:
        res = _run(ExcelEngine.execute_script, code=code, filepath=filepath)
        if res.get("success"):
            return f"✅ Execution Successful:\n{res.get('output')}"
        else:
            return f"❌ Execution Failed:\nError: {res.get('error')}\nOutput:\n{res.get('output')}"
    except Exception as e:
        return f"❌ Error executing script: {str(e)}"


def main():
    mcp.run()


if __name__ == "__main__":
    main()
