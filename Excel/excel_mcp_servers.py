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
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference

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
    if sys.platform == "win32":
        try:
            import ctypes
            user32 = ctypes.windll.user32
            h_default = user32.OpenDesktopW("Default", 0, False, 0x01FF)
            if h_default:
                user32.SetThreadDesktop(h_default)
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
    hex_str = hex_str.strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    return f"FF{hex_str.upper()}"


def split_sheet_and_range(range_str: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Split 'Sheet1!A1:B10' into ('Sheet1', 'A1:B10')."""
    if not range_str:
        return None, None
    if '!' in range_str:
        parts = range_str.split('!', 1)
        sheet = parts[0].strip().strip("'\"")
        rng = parts[1].strip()
        return sheet, rng
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
    def get_app(create_if_missing: bool = False, visible: bool = True):
        if not COM_AVAILABLE:
            raise RuntimeError("pywin32 (win32com) is not available in this Python environment.")
        ensure_com()
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            if visible and not app.Visible:
                app.Visible = True
            return app
        except Exception:
            if create_if_missing:
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = visible
                return app
            raise RuntimeError("Microsoft Excel is not currently running. Please open Excel or specify a file path.")

    @staticmethod
    def get_workbook(app, filepath: Optional[str] = None):
        if filepath:
            norm_path = os.path.abspath(filepath).lower()
            for wb in app.Workbooks:
                try:
                    if wb.FullName.lower() == norm_path or wb.Name.lower() == os.path.basename(norm_path):
                        return wb
                except Exception:
                    pass
            # Not currently open, open it
            return app.Workbooks.Open(os.path.abspath(filepath))
        else:
            if app.Workbooks.Count == 0:
                return app.Workbooks.Add()
            return app.ActiveWorkbook

    @staticmethod
    def get_sheet(wb, sheet_name: Optional[str] = None):
        if sheet_name:
            try:
                return wb.Sheets(sheet_name)
            except Exception:
                raise ValueError(f"Sheet '{sheet_name}' not found in workbook '{wb.Name}'.")
        return wb.ActiveSheet


# ==========================================
# Core Engine Implementation
# ==========================================

class ExcelEngine:
    """Unified engine supporting Live COM Excel and Offline openpyxl."""

    @classmethod
    def get_status(cls) -> Dict[str, Any]:
        """Check Excel runtime status and open workbooks."""
        status_info = {
            "excel_running": False,
            "com_available": COM_AVAILABLE,
            "version": None,
            "open_workbooks": [],
            "active_workbook": None,
            "active_sheet": None
        }

        if COM_AVAILABLE:
            try:
                ensure_com()
                app = win32com.client.GetActiveObject("Excel.Application")
                status_info["excel_running"] = True
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
        
        # If COM is available and either no filepath is given or user wants it visible in Excel:
        if COM_AVAILABLE and (filepath is None or visible):
            try:
                ensure_com()
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
                if filepath:
                    abs_path = os.path.abspath(filepath)
                    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
                    wb.SaveAs(abs_path)
                    saved_path = abs_path
                
                return {
                    "success": True,
                    "message": f"Workbook created successfully in live Excel (Sheet: '{ws.Name}')" + (f", saved to {saved_path}" if saved_path else " (unsaved in Excel)."),
                    "workbook_name": wb.Name,
                    "sheet_name": ws.Name,
                    "filepath": saved_path
                }
            except Exception as e:
                # If COM fails, fall back to openpyxl if filepath provided
                if not filepath:
                    raise e

        # Fallback / Headless creation via openpyxl
        wb = openpyxl.Workbook()
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
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
            
        target_path = os.path.abspath(filepath or "workbook.xlsx")
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        wb.save(target_path)
        
        return {
            "success": True,
            "message": f"Workbook created and saved via openpyxl to {target_path}",
            "filepath": target_path,
            "sheet_name": ws.title
        }

    @classmethod
    def open_workbook(cls, filepath: str, read_only: bool = False, visible: bool = True) -> Dict[str, Any]:
        """Open an existing workbook."""
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        if COM_AVAILABLE:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=True, visible=visible)
                wb = app.Workbooks.Open(abs_path, ReadOnly=read_only)
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
                pass
        
        wb = openpyxl.load_workbook(abs_path, read_only=read_only)
        return {
            "success": True,
            "message": f"Loaded '{os.path.basename(abs_path)}' via openpyxl.",
            "filepath": abs_path,
            "sheets": wb.sheetnames,
            "active_sheet": wb.active.title if wb.active else None
        }

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
        
        # Check if sheet is embedded in range_address
        parsed_sheet, parsed_rng = split_sheet_and_range(range_address)
        sheet_name = parsed_sheet or sheet_name
        range_address = parsed_rng

        # Try Live COM if no filepath is given or Excel has the file open
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                
                if range_address:
                    rng = ws.Range(range_address)
                else:
                    rng = ws.UsedRange
                
                raw_values = rng.Formula if include_formulas else rng.Value
                
                # Normalize values to 2D list
                if not isinstance(raw_values, (tuple, list)):
                    values_matrix = [[raw_values]]
                elif len(raw_values) > 0 and not isinstance(raw_values[0], (tuple, list)):
                    values_matrix = [[v for v in raw_values]]
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
                elif output_format == "records" and len(values_matrix) > 1:
                    headers = [str(h or f"col_{i+1}") for i, h in enumerate(values_matrix[0])]
                    result["records"] = [
                        dict(zip(headers, row)) for row in values_matrix[1:]
                    ]
                return result
            except Exception as e:
                if not filepath:
                    raise e

        # Offline / file-based reading using openpyxl
        if not filepath:
            raise ValueError("Filepath is required when Excel is not running.")
        
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path, data_only=not include_formulas)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        
        if range_address:
            cells = ws[range_address]
            if not isinstance(cells, tuple):
                values_matrix = [[cells.value]]
            elif len(cells) > 0 and not isinstance(cells[0], tuple):
                values_matrix = [[c.value for c in cells]]
            else:
                values_matrix = [[sanitize_value(c.value) for c in row] for row in cells]
        else:
            values_matrix = []
            for row in ws.iter_rows(values_only=True):
                values_matrix.append([sanitize_value(v) for v in row])
        
        result = {
            "filepath": abs_path,
            "sheet": ws.title,
            "range": range_address or f"A1:{get_column_letter(ws.max_column)}{ws.max_row}",
            "row_count": len(values_matrix),
            "col_count": len(values_matrix[0]) if values_matrix else 0,
            "data": values_matrix
        }
        
        if output_format == "markdown_table":
            result["markdown"] = cls._to_markdown_table(values_matrix)
        elif output_format == "records" and len(values_matrix) > 1:
            headers = [str(h or f"col_{i+1}") for i, h in enumerate(values_matrix[0])]
            result["records"] = [
                dict(zip(headers, row)) for row in values_matrix[1:]
            ]
            
        return result

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

        # Try Live COM
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                
                start_rng = ws.Range(start_cell)
                start_row = start_rng.Row
                start_col = start_rng.Column
                
                rows_written = len(data)
                cols_written = max(len(r) for r in data) if data else 0
                
                for r_idx, row_data in enumerate(data):
                    for c_idx, val in enumerate(row_data):
                        cell = ws.Cells(start_row + r_idx, start_col + c_idx)
                        if isinstance(val, str) and val.startswith('='):
                            cell.Formula = val
                        else:
                            cell.Value = val
                
                if auto_fit_columns:
                    end_col = start_col + cols_written - 1
                    ws.Range(ws.Cells(1, start_col), ws.Cells(1, end_col)).EntireColumn.AutoFit()
                
                return {
                    "success": True,
                    "message": f"Successfully wrote {rows_written}x{cols_written} cells to {ws.Name}!{start_cell}",
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "rows_written": rows_written,
                    "cols_written": cols_written
                }
            except Exception as e:
                if not filepath:
                    raise e

        # File-based via openpyxl
        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        
        # Calculate start row and col
        start_cell_obj = ws[start_cell]
        start_row = start_cell_obj.row
        start_col = start_cell_obj.column
        
        rows_written = len(data)
        cols_written = max(len(r) for r in data) if data else 0
        
        for r_idx, row_data in enumerate(data):
            for c_idx, val in enumerate(row_data):
                ws.cell(row=start_row + r_idx, column=start_col + c_idx, value=val)
        
        if auto_fit_columns:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
        wb.save(abs_path)
        return {
            "success": True,
            "message": f"Saved {rows_written}x{cols_written} cells to '{abs_path}' (Sheet: '{ws.title}') starting at {start_cell}",
            "filepath": abs_path,
            "sheet": ws.title,
            "rows_written": rows_written,
            "cols_written": cols_written
        }

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

        # Try Live COM
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
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
                    align_map = {"left": -4131, "center": -4108, "right": -4152}
                    if horizontal_align.lower() in align_map:
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
                    
                return {
                    "success": True,
                    "message": f"Formatted range {range_address} in live Excel ({ws.Name})",
                    "sheet": ws.Name,
                    "range": range_address
                }
            except Exception as e:
                if not filepath:
                    raise e

        # File-based via openpyxl
        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        
        cells = ws[range_address]
        if not isinstance(cells, tuple):
            cells_list = [cells]
        elif len(cells) > 0 and not isinstance(cells[0], tuple):
            cells_list = list(cells)
        else:
            cells_list = [c for row in cells for c in row]
            
        for cell in cells_list:
            # Font
            current_font = cell.font or Font()
            new_font_kwargs = {
                "name": font_name if font_name else current_font.name,
                "size": font_size if font_size else current_font.size,
                "bold": bold if bold is not None else current_font.bold,
                "italic": italic if italic is not None else current_font.italic,
            }
            if font_color:
                new_font_kwargs["color"] = hex_to_argb_hex(font_color)
            cell.font = Font(**new_font_kwargs)
            
            # Fill
            if bg_color:
                cell.fill = PatternFill(start_color=hex_to_argb_hex(bg_color), end_color=hex_to_argb_hex(bg_color), fill_type="solid")
                
            # Alignment
            if horizontal_align:
                curr_align = cell.alignment or Alignment()
                cell.alignment = Alignment(
                    horizontal=horizontal_align,
                    vertical=curr_align.vertical or "center",
                    wrap_text=curr_align.wrap_text
                )
                
            # Number format
            if number_format:
                cell.number_format = number_format
                
            # Border
            if border in ("all_thin", "all"):
                thin_side = Side(border_style="thin", color="000000")
                cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            elif border == "bottom_double":
                double_side = Side(border_style="double", color="000000")
                cell.border = Border(bottom=double_side)
                
        if auto_fit:
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
        wb.save(abs_path)
        return {
            "success": True,
            "message": f"Formatted range {range_address} in '{abs_path}' (Sheet: '{ws.title}')",
            "filepath": abs_path,
            "sheet": ws.title,
            "range": range_address
        }

    @classmethod
    def clear_range(
        cls,
        range_address: Optional[str] = None,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        clear_type: str = "all"
    ) -> Dict[str, Any]:
        """
        Clear contents, formatting, or all from a specific range or entire worksheet.
        
        Args:
            range_address: Optional range to clear (e.g. 'A1:C10'). If omitted, clears entire worksheet.
            filepath: Optional file path. If omitted, targets active workbook in live Excel.
            sheet_name: Optional sheet name. Defaults to active sheet.
            clear_type: 'all' (contents + styling), 'contents' (data only), 'formats' (styling only).
        """
        if COM_AVAILABLE:
            try:
                ensure_com()
                app = win32com.client.GetActiveObject("Excel.Application")
                wb = LiveExcelSession.get_workbook(app, filepath)
                ws = LiveExcelSession.get_sheet(wb, sheet_name)
                
                if range_address:
                    target = ws.Range(range_address)
                    target_desc = f"range {range_address}"
                else:
                    target = ws.Cells
                    target_desc = f"entire sheet '{ws.Name}'"
                    
                ct = clear_type.lower()
                if ct == "contents":
                    target.ClearContents()
                elif ct == "formats":
                    target.ClearFormats()
                else:
                    target.Clear()
                    if not range_address:
                        ws.Columns.ColumnWidth = 10.0
                        
                return {
                    "success": True,
                    "message": f"Successfully cleared {target_desc} ({ct}) in live Excel",
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "range": range_address or "ALL"
                }
            except Exception as com_err:
                if not filepath:
                    raise com_err

        if not filepath:
            raise RuntimeError("Excel is not running and no filepath was provided.")
            
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")
            
        wb = openpyxl.load_workbook(abs_path)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found.")
            ws = wb[sheet_name]
        else:
            ws = wb.active
            
        ct = clear_type.lower()
        if range_address:
            cells = ws[range_address]
            if not isinstance(cells, tuple):
                cells = ((cells,),)
            elif not isinstance(cells[0], tuple):
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

        # Try Live COM
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
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
                        
                return {
                    "success": True,
                    "message": f"Added {chart_type} chart at {position_cell} in live Excel ({ws.Name})",
                    "sheet": ws.Name,
                    "position": position_cell,
                    "title": title
                }
            except Exception as e:
                if not filepath:
                    raise e

        # File-based via openpyxl
        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        
        if "bar" in chart_type_lower:
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
            
        if title:
            chart.title = title
        if x_axis_title and hasattr(chart, 'x_axis'):
            chart.x_axis.title = x_axis_title
        if y_axis_title and hasattr(chart, 'y_axis'):
            chart.y_axis.title = y_axis_title
            
        # Parse range coords
        min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(data_range)
        has_header = (max_row > min_row)

        if max_col > min_col:
            # First column is categories, subsequent columns are data series
            data_ref = Reference(ws, min_col=min_col + 1, min_row=min_row, max_col=max_col, max_row=max_row)
            cats_ref = Reference(ws, min_col=min_col, min_row=min_row + 1 if has_header else min_row, max_row=max_row)
            chart.add_data(data_ref, titles_from_data=has_header)
            chart.set_categories(cats_ref)
        else:
            # Single-column data range: use entire column as data series
            data_ref = Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
            chart.add_data(data_ref, titles_from_data=has_header)
        
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

    @classmethod
    def add_sheet(cls, sheet_name: str, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Add a new worksheet to the workbook."""
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
                ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                ws.Name = sheet_name
                return {
                    "success": True,
                    "message": f"Added worksheet '{sheet_name}' in live Excel workbook '{wb.Name}'.",
                    "sheet_name": sheet_name,
                    "workbook": wb.Name
                }
            except Exception as e:
                if not filepath:
                    raise e
                    
        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path)
        if sheet_name in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' already exists in workbook.")
        ws = wb.create_sheet(title=sheet_name)
        wb.save(abs_path)
        return {
            "success": True,
            "message": f"Added worksheet '{sheet_name}' to '{abs_path}'.",
            "sheet_name": sheet_name,
            "filepath": abs_path
        }

    @classmethod
    def list_sheets(cls, filepath: Optional[str] = None) -> Dict[str, Any]:
        """List all sheets and dimensions."""
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
                sheets_info = []
                for s in wb.Sheets:
                    used = s.UsedRange
                    sheets_info.append({
                        "name": s.Name,
                        "used_range": used.Address.replace('$', ''),
                        "rows": used.Rows.Count,
                        "cols": used.Columns.Count,
                        "is_active": (s.Name == wb.ActiveSheet.Name)
                    })
                return {
                    "workbook": wb.Name,
                    "sheets_count": len(sheets_info),
                    "sheets": sheets_info
                }
            except Exception as e:
                if not filepath:
                    raise e
                    
        if not filepath:
            raise ValueError("Filepath required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        wb = openpyxl.load_workbook(abs_path, read_only=True)
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

    @classmethod
    def export_pdf(
        cls,
        output_pdf_path: str,
        filepath: Optional[str] = None,
        sheet_name: Optional[str] = None,
        orientation: str = "landscape"
    ) -> Dict[str, Any]:
        """Export worksheet or entire workbook to PDF."""
        if not COM_AVAILABLE:
            raise RuntimeError("PDF Export via Excel requires Microsoft Excel and pywin32.")
            
        abs_pdf = os.path.abspath(output_pdf_path)
        os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
        
        ensure_com()
        app = LiveExcelSession.get_app(create_if_missing=True, visible=False)
        wb = LiveExcelSession.get_workbook(app, filepath)
        
        target = wb
        if sheet_name:
            target = LiveExcelSession.get_sheet(wb, sheet_name)
            
        # Page orientation: xlLandscape = 2, xlPortrait = 1
        try:
            target.PageSetup.Orientation = 2 if orientation.lower() == "landscape" else 1
        except Exception:
            pass
            
        # xlTypePDF = 0
        target.ExportAsFixedFormat(0, abs_pdf)
        
        return {
            "success": True,
            "message": f"Successfully exported to PDF at '{abs_pdf}'",
            "pdf_path": abs_pdf
        }

    @classmethod
    def save_workbook(cls, save_as_path: Optional[str] = None, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Save active or specified workbook."""
        if COM_AVAILABLE and not filepath:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                wb = LiveExcelSession.get_workbook(app, filepath)
                if save_as_path:
                    abs_save = os.path.abspath(save_as_path)
                    os.makedirs(os.path.dirname(abs_save), exist_ok=True)
                    wb.SaveAs(abs_save)
                    return {"success": True, "message": f"Workbook saved as '{abs_save}'", "filepath": abs_save}
                else:
                    wb.Save()
                    return {"success": True, "message": f"Workbook '{wb.Name}' saved successfully.", "filepath": wb.FullName}
            except Exception as e:
                if not filepath:
                    raise e
                    
        if not filepath:
            raise ValueError("Filepath is required when Excel application is not running.")
            
        abs_path = os.path.abspath(filepath)
        target = os.path.abspath(save_as_path) if save_as_path else abs_path
        wb = openpyxl.load_workbook(abs_path)
        wb.save(target)
        return {"success": True, "message": f"Saved file to '{target}'", "filepath": target}

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
                        numeric_vals.append(float(v.replace(',', '').replace('$', '').replace('%', '')))
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
        app = None
        wb = None
        ws = None
        
        if COM_AVAILABLE:
            try:
                ensure_com()
                app = LiveExcelSession.get_app(create_if_missing=False)
                if app:
                    wb = LiveExcelSession.get_workbook(app, filepath)
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
            
        # Capture stdout safely using contextlib
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
        except Exception as e:
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
        headers = matrix[0]
        rows = matrix[1:]
        
        header_line = "| " + " | ".join(str(h if h is not None else "") for h in headers) + " |"
        sep_line = "| " + " | ".join("---" for _ in headers) + " |"
        row_lines = ["| " + " | ".join(str(c if c is not None else "") for c in row) + " |" for row in rows]
        
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
