"""
Word MCP Server (Unified Edition)
---------------------------------
Model Context Protocol (MCP) server providing high-level tools to interact with Microsoft Word.
Supports live desktop Word automation (via Windows COM/pywin32) and headless document manipulation (via python-docx).
Consolidated into a single unified script.
"""

from __future__ import annotations

try:
    from mcp.server.fastmcp import FastMCP
except Exception:
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except Exception:
        from mcp.server import MCPServer as FastMCP


# ==========================================
# 1. COM Runtime & Worker Thread
# ==========================================

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


import contextlib
import io
import json
import logging
import math
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import zipfile
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from typing import Any, Callable, Dict, List, Optional, Tuple, Union
from xml.sax.saxutils import escape as xml_escape

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


# ==========================================
# 2. Word Automation Engine
# ==========================================

"""
Word Automation Engine (Enterprise Edition)
-------------------------------------------
Provides a high-level abstraction for interacting with Microsoft Word.
Supports:
1. Live Windows COM Automation (via pywin32) for interactive real-time control of Microsoft Word desktop application.
2. Headless python-docx engine for direct, high-performance file manipulation without requiring Word to be running.
"""

# Headless docx engine
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


# ==========================================
# Color, XML & Style Helpers
# ==========================================

# Word WdBuiltinStyle constants so built-in styles work regardless of Office UI language
COM_BUILTIN_STYLES: Dict[str, int] = {
    "normal": -1,
    "heading 1": -2,
    "heading 2": -3,
    "heading 3": -4,
    "heading 4": -5,
    "heading 5": -6,
    "heading 6": -7,
    "title": -63,
    "subtitle": -75,
    "list bullet": -49,
    "list number": -50,
    "quote": -181,
    "intense quote": -182,
}


def apply_com_style(target, style_name: str) -> None:
    """Apply a built-in Word style by locale-independent constant, falling back to name."""
    key = style_name.strip().lower()
    builtin_id = COM_BUILTIN_STYLES.get(key)
    if builtin_id is not None:
        try:
            target.Style = builtin_id
            return
        except Exception:
            pass
    try:
        target.Style = style_name
    except Exception:
        pass


def resolve_docx_style(doc, style_name: Optional[str]) -> Optional[str]:
    """Return style_name if it exists in doc.styles, else None."""
    if not style_name:
        return None
    for s in doc.styles:
        if s.name == style_name:
            return style_name
    return None


def parse_hex_color(hex_str: str) -> Tuple[int, int, int]:
    """Parse '#RRGGBB' or 'RRGGBB' or 3-digit shorthand into (R, G, B) tuple."""
    if hex_str is None:
        raise ValueError("Hex color string cannot be None.")
    cleaned = str(hex_str).strip().lstrip('#')
    if len(cleaned) == 3:
        cleaned = ''.join([c * 2 for c in cleaned])
    if len(cleaned) != 6:
        raise ValueError(f"Invalid hex color string: '{hex_str}'. Expected 6-digit hex like '#1F497D'.")
    try:
        r = int(cleaned[0:2], 16)
        g = int(cleaned[2:4], 16)
        b = int(cleaned[4:6], 16)
    except ValueError as exc:
        raise ValueError(f"Invalid hex color string: '{hex_str}'. Expected 6-digit hex like '#1F497D'.") from exc
    return r, g, b


def hex_to_bgr_int(hex_str: str) -> int:
    """Convert hex color string to Word COM BGR integer (wdColor)."""
    r, g, b = parse_hex_color(hex_str)
    return r + (g * 256) + (b * 65536)


def bgr_int_to_hex(bgr_val: Any) -> Optional[str]:
    """Convert a Word COM BGR integer color to '#RRGGBB', or None for automatic/undefined."""
    try:
        val = int(bgr_val)
    except (TypeError, ValueError):
        return None
    if val < 0 or val > 0xFFFFFF:
        return None
    r = val & 0xFF
    g = (val >> 8) & 0xFF
    b = (val >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


COM_HIGHLIGHT_MAP: Dict[str, int] = {
    "YELLOW": 7, "BRIGHT_GREEN": 4, "GREEN": 11, "TURQUOISE": 3, "CYAN": 3,
    "PINK": 5, "MAGENTA": 5, "RED": 6, "BLUE": 2, "DARK_BLUE": 9,
    "DARK_RED": 13, "DARK_YELLOW": 14, "GRAY_25": 16, "LIGHT_GRAY": 16,
    "GRAY_50": 15, "GRAY": 15, "TEAL": 10, "VIOLET": 12, "PURPLE": 12, "BLACK": 1,
}


_RICH_TOKEN_RE = re.compile(
    r'(\*\*\*(.+?)\*\*\*'
    r'|\*\*(.+?)\*\*'
    r'|\*(.+?)\*'
    r'|`([^`]+)`'
    r'|\[([^\]]+)\]\((https?://[^\s)]+)\))'
)


def _parse_rich_inline(text: str) -> List[Dict[str, Any]]:
    """
    Parse lightweight inline markdown (***both***, **bold**, *italic*, `code`, [label](url))
    into a list of run descriptors.
    """
    segments: List[Dict[str, Any]] = []
    pos = 0
    for m in _RICH_TOKEN_RE.finditer(text):
        if m.start() > pos:
            segments.append({
                "kind": "text", "text": text[pos:m.start()],
                "bold": False, "italic": False, "code": False, "url": None,
            })
        if m.group(2) is not None:
            segments.append({"kind": "text", "text": m.group(2), "bold": True, "italic": True, "code": False, "url": None})
        elif m.group(3) is not None:
            segments.append({"kind": "text", "text": m.group(3), "bold": True, "italic": False, "code": False, "url": None})
        elif m.group(4) is not None:
            segments.append({"kind": "text", "text": m.group(4), "bold": False, "italic": True, "code": False, "url": None})
        elif m.group(5) is not None:
            segments.append({"kind": "text", "text": m.group(5), "bold": False, "italic": False, "code": True, "url": None})
        elif m.group(6) is not None:
            segments.append({"kind": "link", "text": m.group(6), "bold": False, "italic": False, "code": False, "url": m.group(7)})
        pos = m.end()
    if pos < len(text) or not segments:
        segments.append({
            "kind": "text", "text": text[pos:],
            "bold": False, "italic": False, "code": False, "url": None,
        })
    return [s for s in segments if s["text"] or not text]


def _add_docx_hyperlink(
    paragraph,
    text: str,
    url: str,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: bool = False,
    italic: bool = False,
) -> None:
    """Append a clickable OpenXML hyperlink to a python-docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run_el = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    if font_name:
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rFonts.set(qn("w:cs"), font_name)
        rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if font_size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(round(font_size * 2))))
        rPr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(round(font_size * 2))))
        rPr.append(szCs)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    run_el.append(rPr)
    t_el = OxmlElement("w:t")
    t_el.set(qn("xml:space"), "preserve")
    t_el.text = text
    run_el.append(t_el)
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)


def _docx_insert_paragraph(
    doc,
    text: str = "",
    style: Optional[str] = None,
    position: Optional[str] = None,
    paragraph_index: Optional[int] = None,
):
    """Create a paragraph at start, before a 1-based paragraph_index, or at the end (default)."""
    resolved = resolve_docx_style(doc, style)
    pos = (position or "end").strip().lower()
    if paragraph_index is not None and 1 <= paragraph_index <= len(doc.paragraphs):
        return doc.paragraphs[paragraph_index - 1].insert_paragraph_before(text, style=resolved)
    if pos == "start" and len(doc.paragraphs) > 0:
        return doc.paragraphs[0].insert_paragraph_before(text, style=resolved)
    return doc.add_paragraph(text, style=resolved)


def hex_to_rgb_color(hex_str: str) -> RGBColor:
    """Convert hex color string to python-docx RGBColor."""
    r, g, b = parse_hex_color(hex_str)
    return RGBColor(r, g, b)


def clean_hex(hex_str: str) -> str:
    """Return validated 6-character uppercase hex string without leading '#'."""
    r, g, b = parse_hex_color(hex_str)
    return f"{r:02X}{g:02X}{b:02X}"


def is_arabic_text(text: str) -> bool:
    """Detect if string contains Arabic/Persian/Hebrew characters for RTL formatting."""
    return bool(text and re.search(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\u0590-\u05FF]', text))


def set_paragraph_rtl(paragraph, is_rtl: bool = True) -> None:
    """Set or clear OpenXML w:bidi paragraph direction for proper RTL rendering."""
    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.find(qn("w:bidi"))
    if existing is not None:
        pPr.remove(existing)
    if is_rtl:
        pPr.append(OxmlElement("w:bidi"))


def set_run_font_name(run, font_name: str) -> None:
    """Apply font family to ASCII, high-ANSI, and Complex Script (Arabic/RTL) slots."""
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:cs"), font_name)


def get_docx_alignment(align_str: Optional[str]):
    """Convert string ('left', 'center', 'right', 'justify') to docx WD_ALIGN_PARAGRAPH."""
    if not align_str:
        return None
    a = align_str.strip().lower()
    if a == "center":
        return WD_ALIGN_PARAGRAPH.CENTER
    elif a in ("right", "end"):
        return WD_ALIGN_PARAGRAPH.RIGHT
    elif a in ("justify", "justified"):
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    return WD_ALIGN_PARAGRAPH.LEFT


def get_com_alignment(align_str: Optional[str]) -> int:
    """Convert string to Word COM alignment constant (wdAlignParagraph*)."""
    if not align_str:
        return 0  # wdAlignParagraphLeft
    a = align_str.strip().lower()
    if a == "center":
        return 1  # wdAlignParagraphCenter
    elif a in ("right", "end"):
        return 2  # wdAlignParagraphRight
    elif a in ("justify", "justified"):
        return 3  # wdAlignParagraphJustify
    return 0


def get_highlight_enum(color_str: Optional[str]):
    """Convert color name string to docx WD_COLOR_INDEX."""
    if not color_str:
        return None
    name = color_str.strip().upper().replace(" ", "_").replace("-", "_")
    color_map = {
        "YELLOW": WD_COLOR_INDEX.YELLOW,
        "BRIGHT_GREEN": WD_COLOR_INDEX.BRIGHT_GREEN,
        "GREEN": WD_COLOR_INDEX.GREEN,
        "TURQUOISE": WD_COLOR_INDEX.TURQUOISE,
        "CYAN": WD_COLOR_INDEX.TURQUOISE,
        "PINK": WD_COLOR_INDEX.PINK,
        "MAGENTA": WD_COLOR_INDEX.PINK,
        "RED": WD_COLOR_INDEX.RED,
        "BLUE": WD_COLOR_INDEX.BLUE,
        "DARK_BLUE": WD_COLOR_INDEX.DARK_BLUE,
        "DARK_RED": WD_COLOR_INDEX.DARK_RED,
        "DARK_YELLOW": WD_COLOR_INDEX.DARK_YELLOW,
        "GRAY_25": WD_COLOR_INDEX.GRAY_25,
        "GRAY_50": WD_COLOR_INDEX.GRAY_50,
        "LIGHT_GRAY": WD_COLOR_INDEX.GRAY_25,
        "GRAY": WD_COLOR_INDEX.GRAY_50,
        "TEAL": WD_COLOR_INDEX.TEAL,
        "VIOLET": WD_COLOR_INDEX.VIOLET,
        "PURPLE": WD_COLOR_INDEX.VIOLET,
        "BLACK": WD_COLOR_INDEX.BLACK
    }
    return color_map.get(name, WD_COLOR_INDEX.YELLOW)


def _replace_tcpr_child(tcPr, tag_name: str, elem) -> None:
    """Replace an existing w:tcPr child element with the same tag if present."""
    existing = tcPr.find(qn(tag_name))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(elem)


def set_cell_background(cell, hex_color: str):
    """Set background shading color of a docx table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{clean_hex(hex_color)}"/>')
    _replace_tcpr_child(tcPr, "w:shd", shd)


def set_cell_margins(cell, top_dpt: int = 120, bottom_dpt: int = 120, left_dpt: int = 160, right_dpt: int = 160):
    """Set padding/margins of a docx table cell in twentieths of a point (dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top_dpt}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom_dpt}" w:type="dxa"/>'
        f'<w:left w:w="{left_dpt}" w:type="dxa"/>'
        f'<w:right w:w="{right_dpt}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    _replace_tcpr_child(tcPr, "w:tcMar", tcMar)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """Set custom borders on a docx cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = [f'<w:tcBorders {nsdecls("w")}>']
    for side_name, border_spec in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if border_spec is not None:
            val = border_spec.get('val', 'single')
            sz = border_spec.get('sz', '4')
            color = clean_hex(border_spec.get('color', 'auto')) if border_spec.get('color') != 'auto' else 'auto'
            borders_xml.append(f'<w:{side_name} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>')
        else:
            borders_xml.append(f'<w:{side_name} w:val="nil"/>')
    borders_xml.append('</w:tcBorders>')
    tcBorders = parse_xml(''.join(borders_xml))
    _replace_tcpr_child(tcPr, "w:tcBorders", tcBorders)


def _open_docx_or_dotx(filepath: str):
    """Open a .docx, .dotx, or .dotm file with python-docx (converting template content type in memory)."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in (".dotx", ".dotm"):
        return docx.Document(filepath)
    buf = io.BytesIO()
    with zipfile.ZipFile(filepath, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                ).replace(
                    b"application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                )
            zout.writestr(item, data)
    buf.seek(0)
    return docx.Document(buf)


def _iter_section_headers_footers(doc):
    """Yield unique header and footer objects across all document sections."""
    seen = set()
    for sec in doc.sections:
        for hf in (
            sec.header,
            sec.first_page_header,
            sec.even_page_header,
            sec.footer,
            sec.first_page_footer,
            sec.even_page_footer,
        ):
            elem_id = id(hf._element)
            if elem_id not in seen:
                seen.add(elem_id)
                yield hf


def _iter_table_paragraphs(table):
    """Recursively yield all paragraphs inside a table and any nested tables."""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for nested in cell.tables:
                yield from _iter_table_paragraphs(nested)


def _replace_in_paragraph_runs(
    paragraph,
    pattern: re.Pattern,
    repl_fn: Callable[[re.Match], str],
    max_replacements: Optional[int] = None,
) -> int:
    """
    Replace matches inside a paragraph while preserving run formatting across split runs.

    ``max_replacements`` limits total replacements performed in this paragraph (None = unlimited).
    """
    runs = list(paragraph.runs)
    if not runs:
        return 0
    full_run_text = "".join(r.text or "" for r in runs)
    if not full_run_text:
        return 0

    matches = list(pattern.finditer(full_run_text))
    if not matches:
        return 0

    if max_replacements is not None:
        if max_replacements <= 0:
            return 0
        matches = matches[:max_replacements]

    # Map each run to its character span [start, end) in full_run_text.
    spans = []
    pos = 0
    for idx, r in enumerate(runs):
        txt = r.text or ""
        spans.append((pos, pos + len(txt), idx))
        pos += len(txt)

    # Patch right-to-left so earlier run indices and character offsets remain valid.
    for m in reversed(matches):
        m_start, m_end = m.start(), m.end()
        replacement = repl_fn(m)
        first_hit = True
        for r_start, r_end, r_idx in spans:
            if r_end <= m_start or r_start >= m_end:
                continue
            cur = runs[r_idx].text or ""
            local_start = max(0, m_start - r_start)
            local_end = min(len(cur), m_end - r_start)
            if first_hit:
                prefix = cur[:local_start]
                suffix = cur[m_end - r_start:] if m_end <= r_end else ""
                runs[r_idx].text = prefix + replacement + suffix
                first_hit = False
            elif m_end < r_end:
                runs[r_idx].text = cur[local_end:]
            else:
                runs[r_idx].text = ""

    return len(matches)


# ==========================================
# Live Word COM Management
# ==========================================

def require_filepath(filepath: Optional[str]) -> str:
    """Headless edits need an explicit file. Never invent a hidden default file."""
    if not filepath:
        raise ValueError(
            "No filepath was given and there is no active document in live Word. "
            "Pass `filepath`, or open/create a document in Word first."
        )
    return os.path.abspath(filepath)


class LiveWordSession:
    """
    Manages the COM connection to the Word instance the USER is looking at.

    Design rules (these exist because of real failure modes):
      1. ATTACH to the running instance (GetActiveObject). Never start a second one
         while a WINWORD.EXE process exists - Dispatch() would join the user's
         instance and a later Quit() would close the user's Word.
      2. Only quit an instance this server launched itself, only when it is hidden
         and holds no documents (PDF export).
      3. Never launch an invisible Word for general tools: a hidden instance
         silently swallows documents the user opens later.
      4. Live failures are raised (see handle_live_failure) instead of silently
         falling back to editing the file on disk behind Word's back.
      5. DisplayAlerts is suppressed only for the duration of one tool call.
    """

    _saved_alerts: Optional[int] = None
    _last_launched_by_us: bool = False

    # ---- mode / failure policy -------------------------------------------------
    @staticmethod
    def mode() -> str:
        return get_mode()

    @staticmethod
    def should_use_live() -> bool:
        mode = get_mode()
        if mode == "headless":
            return False
        if not COM_AVAILABLE:
            if mode == "live":
                raise LiveWordError(
                    "WORD_MCP_MODE=live requires Windows with pywin32 installed (pip install pywin32)."
                )
            return False
        return True

    @staticmethod
    def handle_live_failure(exc: BaseException) -> None:
        """
        Called from `except` blocks around live COM code.
        Returns normally only when falling back to headless editing is SAFE
        (Word not running / file not open yet, and mode is not 'live').
        Otherwise raises LiveWordError so the agent and the user see the problem.
        """
        benign = isinstance(exc, (WordNotRunning, FileNotFoundError))
        if benign and get_mode() != "live":
            return
        if isinstance(exc, LiveWordError):
            raise exc
        if isinstance(exc, WordNotRunning):
            raise LiveWordError(
                "Microsoft Word is not running (WORD_MCP_MODE=live). Start Word or use WORD_MCP_MODE=auto."
            ) from exc
        if isinstance(exc, FileNotFoundError):
            raise LiveWordError(str(exc)) from exc
        raise LiveWordError(
            f"Live Word operation failed ({type(exc).__name__}: {exc}). "
            "The document file was NOT modified on disk to avoid conflicting with the open Word window. "
            "Check that Word is idle (no open dialog) and retry."
        ) from exc

    # ---- attach / launch -------------------------------------------------------
    @staticmethod
    def _attach():
        """Return the running Word.Application COM object, or None."""
        if not COM_AVAILABLE:
            return None
        ensure_com()
        try:
            return win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return None

    @staticmethod
    def is_word_running() -> bool:
        """True if Word is running (reachable via COM or at least as a process)."""
        return LiveWordSession._attach() is not None or word_process_exists()

    @classmethod
    def get_app(cls, create_if_missing: bool = True, visible: bool = True, hidden_ok: bool = False):
        """
        Attach to the running Word. Launch one only if no Word process exists at all.

        visible=False is honoured only with hidden_ok=True (PDF export, which quits
        the instance afterwards). Everywhere else a launched Word is always visible.
        """
        if not COM_AVAILABLE:
            raise LiveWordError("pywin32 (win32com) is not installed or unavailable.")
        cls._last_launched_by_us = False

        app = cls._attach()
        if app is not None:
            if visible:
                try:
                    if not app.Visible:
                        app.Visible = True
                except Exception:
                    pass
                try:
                    app.Activate()
                except Exception:
                    pass
            cls._suppress_alerts(app)
            return app

        if word_process_exists():
            raise WordNotReachable(
                "Word is running but COM cannot attach to it. Common causes: Word is elevated "
                "('Run as administrator') while this server is not (or the reverse); Word is still "
                "starting; or a start-up/recovery dialog is open. No second instance was launched, "
                "to avoid interfering with your session. Match privilege levels or close the dialog."
            )

        if not create_if_missing:
            raise WordNotRunning("Microsoft Word is not currently running.")

        app = com_retry(win32com.client.Dispatch, "Word.Application")
        cls._last_launched_by_us = True
        try:
            app.Visible = True if not hidden_ok else bool(visible)
        except Exception:
            pass
        cls._suppress_alerts(app)
        return app

    @classmethod
    def launched_by_us(cls) -> bool:
        """Whether the most recent get_app() call started a brand-new Word process."""
        return cls._last_launched_by_us

    # ---- alerts ----------------------------------------------------------------
    @classmethod
    def _suppress_alerts(cls, app) -> None:
        try:
            if cls._saved_alerts is None:
                cls._saved_alerts = int(app.DisplayAlerts)
            app.DisplayAlerts = 0  # wdAlertsNone
        except Exception:
            pass

    @classmethod
    def restore_ui(cls) -> None:
        """Give DisplayAlerts back to the user. Call after every tool invocation."""
        if cls._saved_alerts is None:
            return
        saved, cls._saved_alerts = cls._saved_alerts, None
        app = cls._attach()
        if app is None:
            return
        try:
            app.DisplayAlerts = saved
        except Exception:
            pass

    # ---- documents -------------------------------------------------------------
    @staticmethod
    def _norm_local_path(path: str) -> str:
        abs_p = os.path.abspath(path)
        try:
            return os.path.normcase(os.path.realpath(abs_p))
        except Exception:
            return os.path.normcase(abs_p)

    @staticmethod
    def find_open_document(app, abs_path: str):
        """Return the already-open document whose path equals abs_path, else None."""
        target = LiveWordSession._norm_local_path(abs_path)
        try:
            for doc in app.Documents:
                try:
                    doc_path = getattr(doc, "FullName", "")
                    if doc_path and not doc_path.lower().startswith(("http://", "https://")):
                        if LiveWordSession._norm_local_path(doc_path) == target:
                            return doc
                except Exception:
                    continue
        except Exception:
            pass
        return None

    @staticmethod
    def com_prepare_range(
        doc,
        position: Optional[str] = None,
        paragraph_index: Optional[int] = None,
    ) -> Any:
        """
        Return a collapsed COM Range for insertion without moving the user's Selection.

        Supports ``paragraph_index`` (1-based, inserts before that paragraph),
        ``position='start'``, or default ``position='end'``.
        """
        try:
            p_count = int(doc.Paragraphs.Count)
        except Exception:
            p_count = 0
        try:
            content_text = doc.Content.Text
        except Exception:
            content_text = ""
        try:
            has_tables = int(doc.Tables.Count) > 0
        except Exception:
            has_tables = False

        if p_count <= 1 and content_text in ("", "\r") and not has_tables:
            return doc.Range(0, 0)

        if paragraph_index is not None and 1 <= paragraph_index <= p_count:
            pos = doc.Paragraphs(paragraph_index).Range.Start
            rng = doc.Range(pos, pos)
            rng.InsertBefore("\r")
            return doc.Range(pos, pos)

        if (position or "end").strip().lower() == "start":
            rng = doc.Range(0, 0)
            rng.InsertBefore("\r")
            return doc.Range(0, 0)

        if p_count > 0:
            last_rng = doc.Paragraphs.Last.Range
            if last_rng.Text in ("", "\r"):
                # Avoid inserting inside a trailing table cell if the final paragraph is in a table
                try:
                    in_tbl = bool(last_rng.Information(12))
                except Exception:
                    in_tbl = False
                if not in_tbl:
                    return doc.Range(last_rng.Start, last_rng.Start)

        end_pos = max(0, int(doc.Content.End) - 1) if content_text.endswith("\r") else int(doc.Content.End)
        rng = doc.Range(end_pos, end_pos)
        rng.InsertAfter("\r")
        rng.Collapse(0)  # wdCollapseEnd
        return rng

    @staticmethod
    def com_prepare_append(app, doc) -> Any:
        """Backward-compatible wrapper returning a Selection at the end of doc."""
        sel = app.Selection
        sel.EndKey(Unit=6)  # wdStory
        try:
            last_text = doc.Paragraphs.Last.Range.Text if doc.Paragraphs.Count > 0 else ""
            has_tables = doc.Tables.Count > 0
            in_table = bool(sel.Information(12))  # wdWithInTable
            if last_text not in ("", "\r") or (has_tables and in_table):
                sel.TypeParagraph()
        except Exception:
            sel.TypeParagraph()
        return sel

    @staticmethod
    def get_document(app, filepath: Optional[str] = None, create_if_empty: bool = True):
        """Get document from Word COM application by filepath or active document."""
        if not app:
            return None
        if filepath:
            abs_path = os.path.abspath(filepath)
            doc = LiveWordSession.find_open_document(app, abs_path)
            if doc is not None:
                com_retry(doc.Activate)
                return doc

            if os.path.exists(abs_path):
                doc = com_retry(
                    app.Documents.Open, abs_path, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False
                )
                com_retry(doc.Activate)
                return doc
            elif create_if_empty:
                doc = com_retry(app.Documents.Add)
                com_retry(doc.Activate)
                return doc
            else:
                raise FileNotFoundError(f"Document file not found: {abs_path}")
        else:
            if com_retry(lambda: app.Documents.Count) > 0:
                doc = com_retry(lambda: app.ActiveDocument)
                com_retry(doc.Activate)
                return doc
            elif create_if_empty:
                doc = com_retry(app.Documents.Add)
                com_retry(doc.Activate)
                return doc
            return None


# ==========================================
# Core Word Automation Engine
# ==========================================

class WordEngine:
    """
    Main engine unifying live Word COM desktop automation and headless python-docx file processing.
    """

    @classmethod
    def get_status(cls) -> Dict[str, Any]:
        """Check live status of Microsoft Word (never launches Word)."""
        status = {
            "mode": get_mode(),
            "word_running": False,
            "word_process_found": False,
            "com_available": COM_AVAILABLE,
            "version": None,
            "open_documents": [],
            "active_document": None,
            "active_document_path": None,
            "hint": None,
        }
        if not COM_AVAILABLE:
            status["hint"] = "pywin32 is not installed: only headless (file) editing is available."
            return status

        app = LiveWordSession._attach()
        if app is None:
            status["word_process_found"] = word_process_exists()
            if status["word_process_found"]:
                status["hint"] = (
                    "A Word process exists but COM cannot attach to it (elevation mismatch, "
                    "start-up dialog, or Word still loading)."
                )
            return status

        status["word_running"] = True
        status["word_process_found"] = True
        try:
            status["version"] = str(getattr(app, "Version", ""))
        except Exception:
            pass

        docs = []
        try:
            for doc in app.Documents:
                try:
                    docs.append({
                        "name": getattr(doc, "Name", ""),
                        "path": getattr(doc, "FullName", ""),
                        "saved": getattr(doc, "Saved", True),
                    })
                except Exception:
                    continue
        except Exception:
            pass
        status["open_documents"] = docs

        try:
            if app.Documents.Count > 0 and app.ActiveDocument:
                status["active_document"] = app.ActiveDocument.Name
                try:
                    status["active_document_path"] = app.ActiveDocument.FullName
                except Exception:
                    status["active_document_path"] = ""
        except Exception:
            pass
        return status

    @classmethod
    def create_document(
        cls,
        filepath: Optional[str] = None,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        author: Optional[str] = None,
        visible: bool = True
    ) -> Dict[str, Any]:
        """
        Create a new Microsoft Word document with optional title/subtitle header block.
        If visible=True and COM is available, opens and displays the new document live in Microsoft Word.
        """
        if filepath:
            abs_path = os.path.abspath(filepath)
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        else:
            abs_path = None

        if (visible or get_mode() == "live") and LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=True, visible=True)
                if abs_path and LiveWordSession.find_open_document(app, abs_path) is not None:
                    raise LiveWordError(
                        f"Document '{abs_path}' is already open in Microsoft Word. Close it first before recreating."
                    )
                doc = app.Documents.Add()
                doc.Activate()
                sel = app.Selection
                sel.HomeKey(Unit=6)  # wdStory

                if title:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 24
                    sel.Font.Bold = True
                    sel.Font.Color = hex_to_bgr_int("#1F497D")
                    is_rtl = is_arabic_text(title)
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_rtl else "left")
                    try:
                        sel.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    sel.ParagraphFormat.SpaceAfter = 6
                    sel.TypeText(title + "\n")

                if subtitle:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 13
                    sel.Font.Italic = True
                    sel.Font.Bold = False
                    sel.Font.Color = hex_to_bgr_int("#595959")
                    is_rtl = is_arabic_text(subtitle)
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_rtl else "left")
                    try:
                        sel.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    sel.ParagraphFormat.SpaceAfter = 14
                    sel.TypeText(subtitle + "\n")

                if author:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 10
                    sel.Font.Italic = False
                    sel.Font.Color = hex_to_bgr_int("#7F7F7F")
                    is_rtl = is_arabic_text(author)
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_rtl else "left")
                    try:
                        sel.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    sel.ParagraphFormat.SpaceAfter = 18
                    author_label = "المؤلف: " if is_rtl else "Author: "
                    sel.TypeText(f"{author_label}{author}\n")

                sel.Font.Reset()
                try:
                    sel.ParagraphFormat.Reset()
                except Exception:
                    pass

                if abs_path:
                    doc.SaveAs2(abs_path)
                    return {
                        "mode": "live_com",
                        "status": "created_and_displayed_live",
                        "doc_name": doc.Name,
                        "filepath": abs_path,
                        "title": title
                    }
                else:
                    return {
                        "mode": "live_com",
                        "status": "created_and_displayed_live",
                        "doc_name": doc.Name,
                        "title": title
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        if abs_path and LiveWordSession.should_use_live():
            attached_app = LiveWordSession._attach()
            if attached_app is not None and LiveWordSession.find_open_document(attached_app, abs_path) is not None:
                raise LiveWordError(
                    f"Document '{abs_path}' is currently open in Microsoft Word; refusing to overwrite on disk."
                )

        # Headless docx creation
        doc = docx.Document()

        if title:
            title_style = resolve_docx_style(doc, "Title")
            p_title = doc.add_paragraph(style=title_style)
            is_rtl = is_arabic_text(title)
            p_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
            set_paragraph_rtl(p_title, is_rtl)
            p_title.paragraph_format.space_after = Pt(6)
            p_title.paragraph_format.space_before = Pt(0)
            run_title = p_title.add_run(title)
            set_run_font_name(run_title, "Segoe UI")
            run_title.font.size = Pt(24)
            run_title.font.bold = True
            run_title.font.color.rgb = hex_to_rgb_color("#1F497D")

        if subtitle:
            sub_style = resolve_docx_style(doc, "Subtitle")
            p_sub = doc.add_paragraph(style=sub_style)
            is_rtl = is_arabic_text(subtitle)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
            set_paragraph_rtl(p_sub, is_rtl)
            p_sub.paragraph_format.space_after = Pt(12)
            p_sub.paragraph_format.space_before = Pt(0)
            run_sub = p_sub.add_run(subtitle)
            set_run_font_name(run_sub, "Segoe UI")
            run_sub.font.size = Pt(13)
            run_sub.font.italic = True
            run_sub.font.color.rgb = hex_to_rgb_color("#595959")

        if author:
            p_auth = doc.add_paragraph()
            is_rtl = is_arabic_text(author)
            p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
            set_paragraph_rtl(p_auth, is_rtl)
            p_auth.paragraph_format.space_after = Pt(18)
            p_auth.paragraph_format.space_before = Pt(0)
            author_label = "المؤلف: " if is_rtl else "Author: "
            run_auth = p_auth.add_run(f"{author_label}{author}")
            set_run_font_name(run_auth, "Segoe UI")
            run_auth.font.size = Pt(10)
            run_auth.font.color.rgb = hex_to_rgb_color("#7F7F7F")

        target_path = abs_path or os.path.abspath("New_Document.docx")
        doc.save(target_path)
        if visible and sys.platform == "win32":
            try:
                os.startfile(target_path)
            except Exception:
                pass
        return {
            "mode": "headless_docx",
            "status": "created_and_saved",
            "filepath": target_path,
            "title": title
        }

    @classmethod
    def open_document(
        cls,
        filepath: str,
        read_only: bool = False,
        visible: bool = True
    ) -> Dict[str, Any]:
        """Open an existing Word document."""
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")


        if visible and sys.platform == "win32" and not COM_AVAILABLE:
            try:
                os.startfile(abs_path)
            except Exception:
                pass

        if LiveWordSession.should_use_live():
            try:
                # NOTE: `visible` is intentionally not used to hide Word: a hidden Word
                # instance would swallow documents the user opens later. When visible=False
                # in auto mode and Word is not running, fall back to headless verification.
                app = LiveWordSession.get_app(
                    create_if_missing=(visible or get_mode() == "live"),
                    visible=True,
                )
                doc = LiveWordSession.find_open_document(app, abs_path)
                already_open = doc is not None
                if not already_open:
                    doc = com_retry(
                        app.Documents.Open,
                        abs_path,
                        ConfirmConversions=False,
                        ReadOnly=read_only,
                        AddToRecentFiles=False,
                    )
                com_retry(doc.Activate)
                return {
                    "mode": "live_com",
                    "status": "already_open_activated" if already_open else "opened_live_on_screen",
                    "doc_name": doc.Name,
                    "filepath": abs_path,
                    "read_only": read_only
                }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        doc = docx.Document(abs_path)
        return {
            "mode": "headless_docx",
            "status": "verified_and_loaded",
            "filepath": abs_path,
            "paragraphs_count": len(doc.paragraphs),
            "tables_count": len(doc.tables)
        }

    @classmethod
    def _extract_com_runs(cls, para_range, txt: str) -> List[Dict[str, Any]]:
        """Read run-level formatting from a COM paragraph Range, coalescing Words only if mixed."""
        def _flag(val: Any) -> Optional[bool]:
            if val in (9999999, None):
                return None
            return bool(val)

        try:
            f = para_range.Font
            b = _flag(getattr(f, "Bold", None))
            it = _flag(getattr(f, "Italic", None))
            ul_raw = getattr(f, "Underline", None)
            ul = None if ul_raw in (9999999, None) else bool(ul_raw)
            sz_raw = getattr(f, "Size", None)
            sz = None if sz_raw in (9999999, None) else float(sz_raw)
            col = bgr_int_to_hex(getattr(f, "Color", None))
        except Exception:
            return [{"text": txt, "bold": None, "italic": None, "underline": None, "font_size": None, "color": None}]

        if None not in (b, it):
            return [{"text": txt, "bold": b, "italic": it, "underline": ul, "font_size": sz, "color": col}]

        # Mixed formatting within the paragraph: coalesce contiguous Word tokens
        runs: List[Dict[str, Any]] = []
        try:
            for w in para_range.Words:
                wt = (w.Text or "").replace("\r", "").replace("\x07", "")
                if not wt:
                    continue
                wf = w.Font
                wb = bool(getattr(wf, "Bold", 0)) if getattr(wf, "Bold", 0) != 9999999 else False
                wi = bool(getattr(wf, "Italic", 0)) if getattr(wf, "Italic", 0) != 9999999 else False
                wu = bool(getattr(wf, "Underline", 0)) if getattr(wf, "Underline", 0) not in (0, 9999999, None) else False
                wsz_raw = getattr(wf, "Size", None)
                wsz = None if wsz_raw in (9999999, None) else float(wsz_raw)
                wcol = bgr_int_to_hex(getattr(wf, "Color", None))
                if (
                    runs
                    and runs[-1]["bold"] == wb
                    and runs[-1]["italic"] == wi
                    and runs[-1]["underline"] == wu
                    and runs[-1]["font_size"] == wsz
                    and runs[-1]["color"] == wcol
                ):
                    runs[-1]["text"] += wt
                else:
                    runs.append({"text": wt, "bold": wb, "italic": wi, "underline": wu, "font_size": wsz, "color": wcol})
        except Exception:
            pass
        return runs or [{"text": txt, "bold": b, "italic": it, "underline": ul, "font_size": sz, "color": col}]

    @classmethod
    def _read_live_com_fallback(
        cls,
        doc,
        output_format: str,
        include_tables: bool,
        max_paragraphs: Optional[int],
    ) -> Dict[str, Any]:
        """Extract rich document structure and formatting directly via COM."""
        paragraphs_data = []
        md_lines = []
        plain_lines = []
        para_count = 0

        for para in doc.Paragraphs:
            if max_paragraphs and para_count >= max_paragraphs:
                break
            p_rng = para.Range
            txt = p_rng.Text.replace("\x07", "").strip("\r\n ").strip()
            if not txt:
                continue
            para_count += 1
            try:
                style_name = str(getattr(para.Style, "NameLocal", "") or getattr(para, "Style", "Normal"))
            except Exception:
                style_name = "Normal"
            try:
                level = int(para.OutlineLevel)
            except Exception:
                level = 10
            if 1 <= level <= 6 and not style_name.startswith("Heading"):
                style_name = f"Heading {level}"

            try:
                list_type = int(p_rng.ListFormat.ListType)
            except Exception:
                list_type = 0

            runs_data = cls._extract_com_runs(p_rng, txt)
            paragraphs_data.append({
                "index": para_count,
                "style": style_name,
                "text": txt,
                "runs": runs_data,
            })
            plain_lines.append(txt)

            if 1 <= level <= 6:
                md_lines.append(f"{'#' * level} {txt}\n")
            elif list_type == 2:
                md_lines.append(f"- {txt}")
            elif list_type in (3, 4):
                md_lines.append(f"1. {txt}")
            else:
                parts = []
                for mr in runs_data:
                    rt = mr["text"]
                    if mr.get("bold") and mr.get("italic"):
                        parts.append(f"***{rt}***")
                    elif mr.get("bold"):
                        parts.append(f"**{rt}**")
                    elif mr.get("italic"):
                        parts.append(f"*{rt}*")
                    else:
                        parts.append(rt)
                md_lines.append("".join(parts) + "\n")

        tables_data = []
        if include_tables:
            try:
                for t_idx in range(1, doc.Tables.Count + 1):
                    tbl = doc.Tables(t_idx)
                    tbl_rows = []
                    for r_idx in range(1, tbl.Rows.Count + 1):
                        row_cells = []
                        for c_idx in range(1, tbl.Columns.Count + 1):
                            try:
                                cell_txt = tbl.Cell(r_idx, c_idx).Range.Text.replace("\r\x07", "").replace("\x07", "").strip()
                            except Exception:
                                cell_txt = ""
                            row_cells.append(cell_txt.replace("\n", " "))
                        tbl_rows.append(row_cells)
                    tables_data.append({
                        "table_index": t_idx,
                        "rows_count": len(tbl_rows),
                        "cols_count": len(tbl_rows[0]) if tbl_rows else 0,
                        "data": tbl_rows,
                    })
                    if tbl_rows:
                        plain_lines.append("")
                        for rc in tbl_rows:
                            plain_lines.append("\t".join(rc))
                        plain_lines.append("")
                        if output_format == "markdown":
                            md_lines.append(f"\n**Table {t_idx}** ({len(tbl_rows)} rows x {len(tbl_rows[0])} cols):\n")
                            headers = [c.replace("|", "\\|") for c in tbl_rows[0]]
                            md_lines.append("| " + " | ".join(headers) + " |")
                            md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                            for rc in tbl_rows[1:]:
                                escaped = [c.replace("|", "\\|") for c in rc]
                                md_lines.append("| " + " | ".join(escaped) + " |")
                            md_lines.append("\n")
            except Exception:
                pass

        comments_data = []
        try:
            for c_idx in range(1, doc.Comments.Count + 1):
                cm = doc.Comments(c_idx)
                comments_data.append({
                    "author": getattr(cm, "Author", "") or "",
                    "text": getattr(cm.Range, "Text", "").strip() if hasattr(cm, "Range") else "",
                })
        except Exception:
            pass

        res: Dict[str, Any] = {
            "document_name": getattr(doc, "Name", ""),
            "filepath": getattr(doc, "FullName", "") or getattr(doc, "Name", ""),
            "mode": "live_com",
            "total_paragraphs": len(paragraphs_data),
            "total_tables": len(tables_data),
            "total_words": sum(len(p["text"].split()) for p in paragraphs_data),
        }
        if output_format == "markdown":
            if comments_data:
                md_lines.append("\n---\n**Comments:**\n")
                for cm in comments_data:
                    md_lines.append(f"- **{cm['author']}:** {cm['text']}")
            res["markdown"] = "\n".join(md_lines)
        elif output_format == "text":
            if comments_data:
                plain_lines.append("\nComments:")
                for cm in comments_data:
                    plain_lines.append(f"[{cm['author']}]: {cm['text']}")
            res["text"] = "\n".join(plain_lines)
        else:
            res["paragraphs"] = paragraphs_data
            res["tables"] = tables_data
            if comments_data:
                res["comments"] = comments_data
        return res

    @classmethod
    def read_document(
        cls,
        filepath: Optional[str] = None,
        output_format: str = "markdown",
        include_tables: bool = True,
        max_paragraphs: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Read text and structure from a Word document or the active document.
        """
        temp_copy = None
        source_display_path = None
        if LiveWordSession.should_use_live():
            try:
                app = None
                doc = None
                if not filepath:
                    app = LiveWordSession.get_app(create_if_missing=False)
                    if app.Documents.Count > 0:
                        doc = app.ActiveDocument
                else:
                    abs_target = os.path.abspath(filepath)
                    app = LiveWordSession._attach()
                    if app is not None:
                        doc = LiveWordSession.find_open_document(app, abs_target)
                    elif get_mode() == "live":
                        app = LiveWordSession.get_app(create_if_missing=False)
                        doc = LiveWordSession.find_open_document(app, abs_target)

                if doc is not None:
                    source_display_path = getattr(doc, "FullName", "") or getattr(doc, "Name", "")
                    if output_format == "text":
                        return {
                            "document_name": doc.Name,
                            "filepath": source_display_path,
                            "mode": "live_com",
                            "text": doc.Content.Text.replace("\r", "\n"),
                        }
                    fd, temp_path = tempfile.mkstemp(prefix="_temp_mcp_read_", suffix=".docx")
                    os.close(fd)
                    try:
                        doc.SaveCopyAs(temp_path)
                        filepath = temp_path
                        temp_copy = temp_path
                    except Exception:
                        try:
                            os.remove(temp_path)
                        except OSError:
                            pass
                        return cls._read_live_com_fallback(doc, output_format, include_tables, max_paragraphs)
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        if not filepath:
            raise ValueError("No filepath provided and no active document found in Microsoft Word.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        try:
            doc = docx.Document(abs_path)
        finally:
            if temp_copy:
                try:
                    os.remove(temp_copy)
                except OSError:
                    pass

        paragraphs_data = []
        md_lines = []
        plain_lines = []

        para_count = 0
        for p in doc.paragraphs:
            if max_paragraphs and para_count >= max_paragraphs:
                break

            p_text = p.text.strip()
            if not p_text and not p.runs:
                continue

            para_count += 1
            style_name = p.style.name if p.style else "Normal"

            runs_data = []
            for r in p.runs:
                runs_data.append({
                    "text": r.text,
                    "bold": r.bold,
                    "italic": r.italic,
                    "underline": r.underline,
                    "font_size": r.font.size.pt if r.font.size else None,
                    "color": str(r.font.color.rgb) if r.font.color and r.font.color.rgb else None
                })

            paragraphs_data.append({
                "index": para_count,
                "style": style_name,
                "text": p_text,
                "runs": runs_data
            })
            plain_lines.append(p_text)

            # Markdown conversion with smart run merging
            is_title_heading = (
                style_name.startswith("Title") or
                (p.runs and len(p.runs) > 0 and p.runs[0].font.size and p.runs[0].font.size.pt >= 20 and p.runs[0].bold)
            )
            heading_match = re.match(r'^Heading\s+([1-6])\b', style_name)

            if heading_match:
                h_lvl = int(heading_match.group(1))
                md_lines.append(f"{'#' * h_lvl} {p_text}\n")
            elif is_title_heading:
                md_lines.append(f"# {p_text}\n")
            elif style_name.startswith("Subtitle"):
                md_lines.append(f"*{p_text}*\n")
            elif style_name.startswith("List Bullet"):
                md_lines.append(f"- {p_text}")
            elif style_name.startswith("List Number"):
                md_lines.append(f"1. {p_text}")
            elif style_name == "Quote":
                md_lines.append(f"> {p_text}\n")
            else:
                # Merge consecutive runs with identical formatting
                merged_runs = []
                for r in p.runs:
                    txt = r.text
                    if not txt:
                        continue
                    b = bool(r.bold)
                    it = bool(r.italic)
                    if merged_runs and merged_runs[-1]["bold"] == b and merged_runs[-1]["italic"] == it:
                        merged_runs[-1]["text"] += txt
                    else:
                        merged_runs.append({"text": txt, "bold": b, "italic": it})

                formatted_para = []
                for mr in merged_runs:
                    txt = mr["text"]
                    if mr["bold"] and mr["italic"]:
                        formatted_para.append(f"***{txt}***")
                    elif mr["bold"]:
                        formatted_para.append(f"**{txt}**")
                    elif mr["italic"]:
                        formatted_para.append(f"*{txt}*")
                    else:
                        formatted_para.append(txt)
                md_lines.append("".join(formatted_para) + "\n")

        tables_data = []
        if include_tables and doc.tables:
            for t_idx, tbl in enumerate(doc.tables, 1):
                tbl_rows = []
                for row in tbl.rows:
                    tbl_rows.append([cell.text.strip().replace("\n", " ") for cell in row.cells])

                tables_data.append({
                    "table_index": t_idx,
                    "rows_count": len(tbl.rows),
                    "cols_count": len(tbl.columns),
                    "data": tbl_rows
                })

                if tbl_rows:
                    plain_lines.append("")
                    for row_cells in tbl_rows:
                        plain_lines.append("\t".join(row_cells))
                    plain_lines.append("")

                if output_format == "markdown" and tbl_rows:
                    md_lines.append(f"\n**Table {t_idx}** ({len(tbl.rows)} rows x {len(tbl.columns)} cols):\n")
                    headers = [h.replace("|", "\\|") for h in tbl_rows[0]]
                    md_lines.append("| " + " | ".join(headers) + " |")
                    md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                    for row_cells in tbl_rows[1:]:
                        escaped = [c.replace("|", "\\|") for c in row_cells]
                        md_lines.append("| " + " | ".join(escaped) + " |")
                    md_lines.append("\n")

        comments_data = []
        if hasattr(doc, "comments"):
            try:
                for cm in doc.comments:
                    comments_data.append({
                        "author": getattr(cm, "author", "") or "",
                        "text": getattr(cm, "text", "") or "",
                    })
            except Exception:
                pass

        total_words = sum(len(p["text"].split()) for p in paragraphs_data)

        res = {
            "filepath": source_display_path or abs_path,
            "total_paragraphs": len(paragraphs_data),
            "total_tables": len(tables_data),
            "total_words": total_words
        }

        if output_format == "markdown":
            if comments_data:
                md_lines.append("\n---\n**Comments:**\n")
                for cm in comments_data:
                    md_lines.append(f"- **{cm['author']}:** {cm['text']}")
            res["markdown"] = "\n".join(md_lines)
        elif output_format == "text":
            if comments_data:
                plain_lines.append("\nComments:")
                for cm in comments_data:
                    plain_lines.append(f"[{cm['author']}]: {cm['text']}")
            res["text"] = "\n".join(plain_lines)
        else:
            res["paragraphs"] = paragraphs_data
            res["tables"] = tables_data
            if comments_data:
                res["comments"] = comments_data

        return res

    @classmethod
    def get_outline(cls, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Extract hierarchical outline (Headings 1-6, tables, word count)."""
        read_res = cls.read_document(filepath=filepath, output_format="json", include_tables=True)
        outline_items = []
        for p in read_res.get("paragraphs", []):
            style = p.get("style", "")
            if "Heading" in style or style in ("Title", "Subtitle"):
                level = 1
                if style.startswith("Heading"):
                    try:
                        level = int(style.replace("Heading", "").strip().split()[0])
                    except Exception:
                        level = 1
                elif style == "Subtitle":
                    level = 2

                outline_items.append({
                    "title": p.get("text"),
                    "level": level,
                    "style": style,
                    "paragraph_index": p.get("index")
                })

        return {
            "filepath": read_res.get("filepath"),
            "total_words": read_res.get("total_words"),
            "total_paragraphs": read_res.get("total_paragraphs"),
            "total_tables": read_res.get("total_tables"),
            "outline_count": len(outline_items),
            "outline": outline_items
        }

    @classmethod
    def add_heading(
        cls,
        text: str,
        level: int = 1,
        filepath: Optional[str] = None,
        align: Optional[str] = None,
        color: Optional[str] = None,
        position: Optional[str] = None,
        paragraph_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Add a heading (Level 1 to 6) to the document."""
        level = max(1, min(6, int(level)))
        is_rtl = is_arabic_text(text)

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc, position=position, paragraph_index=paragraph_index)
                    rng.Text = text
                    apply_com_style(rng, f"Heading {level}")
                    rng.Font.Reset()
                    try:
                        rng.ParagraphFormat.Reset()
                    except Exception:
                        pass
                    if align:
                        rng.ParagraphFormat.Alignment = get_com_alignment(align)
                    elif is_rtl:
                        rng.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight
                    else:
                        rng.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
                    try:
                        rng.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    if color:
                        rng.Font.Color = hex_to_bgr_int(color)
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "heading_added_live",
                        "doc_name": doc.Name,
                        "level": level,
                        "text": text
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        heading = _docx_insert_paragraph(
            doc,
            text,
            style=f"Heading {level}",
            position=position,
            paragraph_index=paragraph_index,
        )
        if align:
            heading.alignment = get_docx_alignment(align)
            if align.strip().lower() in ("right", "end") and is_rtl:
                set_paragraph_rtl(heading, True)
        elif is_rtl:
            heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_paragraph_rtl(heading, True)
        if color and heading.runs:
            heading.runs[0].font.color.rgb = hex_to_rgb_color(color)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "heading_added",
            "filepath": abs_path,
            "level": level,
            "text": text
        }

    @classmethod
    def add_paragraph(
        cls,
        text: str,
        filepath: Optional[str] = None,
        style: Optional[str] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: Optional[float] = None,
        font_name: Optional[str] = None,
        font_color: Optional[str] = None,
        highlight_color: Optional[str] = None,
        align: Optional[str] = None,
        line_spacing: Optional[float] = None,
        space_before: Optional[float] = None,
        space_after: Optional[float] = None,
        position: Optional[str] = None,
        paragraph_index: Optional[int] = None,
        rich_formatting: bool = False,
    ) -> Dict[str, Any]:
        """Add a rich paragraph with optional inline markdown formatting, links, color, highlight, and positioning."""
        is_rtl = is_arabic_text(text)

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc, position=position, paragraph_index=paragraph_index)
                    apply_com_style(rng, style or "Normal")
                    rng.Font.Reset()
                    try:
                        rng.ParagraphFormat.Reset()
                    except Exception:
                        pass
                    if align:
                        rng.ParagraphFormat.Alignment = get_com_alignment(align)
                    elif is_rtl:
                        rng.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight
                    try:
                        rng.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    if space_after is not None:
                        rng.ParagraphFormat.SpaceAfter = space_after
                    if space_before is not None:
                        rng.ParagraphFormat.SpaceBefore = space_before
                    if line_spacing is not None:
                        rng.ParagraphFormat.LineSpacingRule = 5  # wdLineSpaceMultiple
                        rng.ParagraphFormat.LineSpacing = line_spacing * 12

                    hl_idx = None
                    if highlight_color:
                        hl_key = highlight_color.strip().upper().replace(" ", "_").replace("-", "_")
                        hl_idx = COM_HIGHLIGHT_MAP.get(hl_key, 7)

                    def _style_com_run(target_rng, seg_bold: bool, seg_italic: bool, seg_code: bool) -> None:
                        target_rng.Font.Bold = bool(bold or seg_bold)
                        target_rng.Font.Italic = bool(italic or seg_italic)
                        target_rng.Font.Underline = 1 if underline else 0
                        if seg_code:
                            target_rng.Font.Name = "Consolas"
                        elif font_name:
                            target_rng.Font.Name = font_name
                        if font_size:
                            target_rng.Font.Size = font_size
                        if font_color:
                            target_rng.Font.Color = hex_to_bgr_int(font_color)
                        if hl_idx is not None:
                            try:
                                target_rng.HighlightColorIndex = hl_idx
                            except Exception:
                                pass

                    if rich_formatting:
                        cur_pos = int(rng.Start)
                        for seg in _parse_rich_inline(text):
                            seg_rng = doc.Range(cur_pos, cur_pos)
                            if seg["kind"] == "link":
                                try:
                                    doc.Hyperlinks.Add(Anchor=seg_rng, Address=seg["url"], TextToDisplay=seg["text"])
                                    cur_pos = int(seg_rng.End) if int(seg_rng.End) > cur_pos else cur_pos + len(seg["text"])
                                except Exception:
                                    seg_rng.Text = seg["text"]
                                    _style_com_run(seg_rng, False, False, False)
                                    cur_pos = int(seg_rng.End)
                            else:
                                seg_rng.Text = seg["text"]
                                _style_com_run(seg_rng, seg["bold"], seg["italic"], seg["code"])
                                cur_pos = int(seg_rng.End)
                    else:
                        rng.Text = text
                        _style_com_run(rng, False, False, False)

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "paragraph_added_live",
                        "doc_name": doc.Name,
                        "text_length": len(text)
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        p = _docx_insert_paragraph(doc, "", style=style, position=position, paragraph_index=paragraph_index)
        if align:
            p.alignment = get_docx_alignment(align)
        elif is_rtl:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if is_rtl:
            set_paragraph_rtl(p, True)
        if space_before is not None:
            p.paragraph_format.space_before = Pt(space_before)
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
        if line_spacing is not None:
            p.paragraph_format.line_spacing = line_spacing

        segments = _parse_rich_inline(text) if rich_formatting else [
            {"kind": "text", "text": text, "bold": False, "italic": False, "code": False, "url": None}
        ]
        for seg in segments:
            if seg["kind"] == "link":
                _add_docx_hyperlink(
                    p,
                    text=seg["text"],
                    url=seg["url"],
                    font_name=font_name,
                    font_size=font_size,
                    bold=bold,
                    italic=italic,
                )
                continue
            run = p.add_run(seg["text"])
            if bold or seg["bold"]:
                run.bold = True
            if italic or seg["italic"]:
                run.italic = True
            if underline:
                run.underline = True
            if font_size:
                run.font.size = Pt(font_size)
            if seg["code"]:
                set_run_font_name(run, "Consolas")
                if not font_color:
                    run.font.color.rgb = hex_to_rgb_color("#B91C1C")
            elif font_name:
                set_run_font_name(run, font_name)
            if font_color:
                run.font.color.rgb = hex_to_rgb_color(font_color)
            if highlight_color:
                hl = get_highlight_enum(highlight_color)
                if hl:
                    run.font.highlight_color = hl

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "paragraph_added",
            "filepath": abs_path,
            "text_length": len(text)
        }

    @classmethod
    def add_bullet_list(
        cls,
        items: List[str],
        filepath: Optional[str] = None,
        list_type: str = "bullet"
    ) -> Dict[str, Any]:
        """Add a bulleted or numbered list."""
        style_name = "List Number" if list_type.lower() in ("number", "numbered", "1", "numeric") else "List Bullet"

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc)
                    for idx, item in enumerate(items):
                        if idx > 0:
                            rng.Collapse(0)
                            rng.InsertAfter("\r")
                            rng.Collapse(0)
                        rng.Text = item
                        apply_com_style(rng, style_name)
                        is_rtl = is_arabic_text(item)
                        if is_rtl:
                            rng.ParagraphFormat.Alignment = 2
                            try:
                                rng.ParagraphFormat.ReadingOrder = 0
                            except Exception:
                                pass
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "list_added_live",
                        "doc_name": doc.Name,
                        "items_count": len(items),
                        "list_type": list_type
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()
        resolved_style = resolve_docx_style(doc, style_name)

        for item in items:
            p = doc.add_paragraph(item, style=resolved_style)
            if is_arabic_text(item):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_paragraph_rtl(p, True)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "list_added",
            "filepath": abs_path,
            "items_count": len(items),
            "list_type": list_type
        }

    @classmethod
    def add_callout_box(
        cls,
        text: str,
        filepath: Optional[str] = None,
        title: Optional[str] = None,
        box_type: str = "info",
        bg_color: Optional[str] = None,
        border_color: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a styled callout / alert box with custom background and left/right border accent."""
        is_rtl = is_arabic_text(text) or (bool(title) and is_arabic_text(title or ""))
        box_themes = {
            "info": {"bg": "#EDF5FC", "border": "#1F497D", "icon": "ℹ️ ", "title": "معلومات" if is_rtl else "Information"},
            "tip": {"bg": "#EDF7ED", "border": "#2E7D32", "icon": "💡 ", "title": "نصيحة" if is_rtl else "Tip"},
            "warning": {"bg": "#FFF4E5", "border": "#ED6C02", "icon": "⚠️ ", "title": "تنبيه" if is_rtl else "Warning"},
            "note": {"bg": "#F5F5F5", "border": "#616161", "icon": "📝 ", "title": "ملاحظة" if is_rtl else "Note"},
            "quote": {"bg": "#FAFAFA", "border": "#757575", "icon": "❝ ", "title": "اقتباس" if is_rtl else "Quote"},
            "important": {"bg": "#FDEDED", "border": "#D32F2F", "icon": "🚨 ", "title": "هام جداً" if is_rtl else "Important"}
        }
        theme = box_themes.get(box_type.lower(), box_themes["info"])
        final_bg = bg_color or theme["bg"]
        final_border = border_color or theme["border"]
        final_title = title if title is not None else f"{theme['icon']}{theme['title']}"
        is_rtl = is_rtl or (bool(final_title) and is_arabic_text(final_title))

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc)
                    tbl = doc.Tables.Add(Range=rng, NumRows=1, NumColumns=1)
                    try:
                        tbl.Borders.Enable = False
                    except Exception:
                        pass
                    cell = tbl.Cell(1, 1)
                    cell.Shading.BackgroundPatternColor = hex_to_bgr_int(final_bg)

                    accent_border_idx = -4 if is_rtl else -2  # wdBorderRight=-4, wdBorderLeft=-2
                    for b_idx in (-1, -2, -3, -4):
                        cell.Borders(b_idx).LineStyle = 0
                    cell.Borders(accent_border_idx).LineStyle = 1
                    cell.Borders(accent_border_idx).LineWidth = 24  # 3.0 pt
                    cell.Borders(accent_border_idx).Color = hex_to_bgr_int(final_border)

                    try:
                        cell.TopPadding = 7
                        cell.BottomPadding = 7
                        cell.LeftPadding = 8 if is_rtl else 10
                        cell.RightPadding = 10 if is_rtl else 8
                    except Exception:
                        pass

                    cell_range = cell.Range
                    cell_range.Text = (final_title + "\n" if final_title else "") + text
                    cell_range.Font.Name = "Segoe UI"
                    cell_range.Font.Size = 10
                    cell_range.Font.Italic = (box_type.lower() == "quote")
                    cell_range.Font.Color = hex_to_bgr_int("#262626")
                    cell_range.ParagraphFormat.Alignment = 2 if is_rtl else 0
                    try:
                        cell_range.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                    except Exception:
                        pass
                    if final_title:
                        first_line = cell.Range.Paragraphs(1).Range
                        first_line.Font.Size = 11
                        first_line.Font.Bold = True
                        first_line.Font.Italic = False
                        first_line.Font.Color = hex_to_bgr_int(final_border)

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "callout_box_added_live",
                        "doc_name": doc.Name,
                        "box_type": box_type,
                        "title": final_title
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Clear outer table borders
        tblPr = table._tbl.tblPr
        existing_borders = tblPr.find(qn("w:tblBorders"))
        if existing_borders is not None:
            tblPr.remove(existing_borders)
        tblBorders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="none"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:insideH w:val="none"/>'
            f'<w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(tblBorders)

        cell = table.cell(0, 0)
        cell.width = Inches(6.5)

        border_spec = {"val": "single", "sz": "24", "color": final_border}
        if is_rtl:
            set_cell_borders(cell, top=None, bottom=None, left=None, right=border_spec)
        else:
            set_cell_borders(cell, top=None, bottom=None, left=border_spec, right=None)
        set_cell_background(cell, final_bg)
        set_cell_margins(
            cell,
            top_dpt=140,
            bottom_dpt=140,
            left_dpt=160 if is_rtl else 200,
            right_dpt=200 if is_rtl else 160,
        )

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4) if text else Pt(0)
        if is_rtl:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_paragraph_rtl(p, True)

        if final_title:
            run_title = p.add_run(final_title + "\n")
            set_run_font_name(run_title, "Segoe UI")
            run_title.font.size = Pt(11)
            run_title.font.bold = True
            run_title.font.color.rgb = hex_to_rgb_color(final_border)

        if text:
            run_text = p.add_run(text)
            set_run_font_name(run_text, "Segoe UI")
            run_text.font.size = Pt(10)
            run_text.font.italic = (box_type.lower() == "quote")
            run_text.font.color.rgb = hex_to_rgb_color("#262626")

        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(6)
        spacer.paragraph_format.space_after = Pt(6)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "callout_box_added",
            "filepath": abs_path,
            "box_type": box_type,
            "title": final_title
        }

    @classmethod
    def add_table(
        cls,
        headers: Optional[List[str]] = None,
        data: Optional[List[List[Any]]] = None,
        filepath: Optional[str] = None,
        table_style: str = "Table Grid",
        header_bg_color: Optional[str] = "#1F497D",
        header_font_color: Optional[str] = "#FFFFFF",
        alternate_row_color: Optional[str] = "#F2F4F8",
        col_widths: Optional[List[float]] = None,
        alignment: str = "center",
        auto_fit: bool = True,
        merges: Optional[List[List[int]]] = None,
    ) -> Dict[str, Any]:
        """Add a styled table with custom headers, zebra rows, column widths, and optional cell merges."""
        data = data or []
        if not headers and not data:
            raise ValueError("Cannot add an empty table: provide `headers`, `data`, or both.")
        num_cols = len(headers) if headers else max((len(r) for r in data), default=1)
        if num_cols <= 0:
            raise ValueError("Table must have at least one column.")
        num_rows = (1 if headers else 0) + len(data)

        validated_merges: List[Tuple[int, int, int, int]] = []
        if merges:
            for spec in merges:
                if len(spec) != 4:
                    raise ValueError(f"Each merge spec must have 4 integers [r1, c1, r2, c2], got {spec}.")
                r1, c1, r2, c2 = (int(x) for x in spec)
                if not (0 <= r1 <= r2 < num_rows and 0 <= c1 <= c2 < num_cols):
                    raise ValueError(
                        f"Merge coordinates {spec} out of bounds for {num_rows}x{num_cols} table (0-based)."
                    )
                validated_merges.append((r1, c1, r2, c2))

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc)
                    tbl = doc.Tables.Add(Range=rng, NumRows=num_rows, NumColumns=num_cols)
                    if table_style:
                        try:
                            tbl.Style = table_style
                        except Exception:
                            pass
                    try:
                        tbl.Rows.Alignment = get_com_alignment(alignment)
                        tbl.Rows.AllowBreakAcrossPages = 0
                    except Exception:
                        pass

                    row_offset = 1
                    if headers:
                        try:
                            tbl.Rows(1).HeadingFormat = -1
                        except Exception:
                            pass
                        for c_i, h_text in enumerate(headers, 1):
                            cell = tbl.Cell(1, c_i)
                            cell.Range.Text = str(h_text)
                            cell.Range.Font.Bold = True
                            cell.Range.Font.Name = "Segoe UI"
                            cell.Range.Font.Size = 10.5
                            cell.Range.ParagraphFormat.Alignment = 1  # Center
                            try:
                                cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                            except Exception:
                                pass
                            if header_font_color:
                                cell.Range.Font.Color = hex_to_bgr_int(header_font_color)
                            if header_bg_color:
                                cell.Shading.BackgroundPatternColor = hex_to_bgr_int(header_bg_color)
                        row_offset = 2

                    for r_i, row_values in enumerate(data):
                        curr_row = row_offset + r_i
                        is_even = (r_i % 2 == 1)
                        for c_i in range(1, num_cols + 1):
                            cell = tbl.Cell(curr_row, c_i)
                            val = row_values[c_i - 1] if (c_i - 1) < len(row_values) else ""
                            val_str = str(val) if val is not None else ""
                            cell.Range.Text = val_str
                            cell.Range.Font.Name = "Segoe UI"
                            cell.Range.Font.Size = 9.5
                            cell.Range.Font.Color = hex_to_bgr_int("#262626")
                            try:
                                cell.VerticalAlignment = 1
                            except Exception:
                                pass
                            try:
                                num = float(val_str.replace("$", "").replace(",", "").replace("%", ""))
                                is_num = math.isfinite(num) and bool(val_str.strip())
                            except ValueError:
                                is_num = False
                            cell.Range.ParagraphFormat.Alignment = 2 if (is_num or is_arabic_text(val_str)) else 0
                            if is_even and alternate_row_color:
                                cell.Shading.BackgroundPatternColor = hex_to_bgr_int(alternate_row_color)

                    if col_widths:
                        try:
                            tbl.AllowAutoFit = False
                        except Exception:
                            pass
                        for c_i, w in enumerate(col_widths, 1):
                            if c_i <= num_cols:
                                tbl.Columns(c_i).Width = w * 72

                    for r1, c1, r2, c2 in validated_merges:
                        tbl.Cell(r1 + 1, c1 + 1).Merge(tbl.Cell(r2 + 1, c2 + 1))

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "table_added_live",
                        "doc_name": doc.Name,
                        "rows": num_rows,
                        "cols": num_cols
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        table = doc.add_table(rows=num_rows, cols=num_cols)
        resolved_tbl_style = resolve_docx_style(doc, table_style)
        if resolved_tbl_style:
            table.style = resolved_tbl_style
        align_low = (alignment or "center").strip().lower()
        if align_low == "center":
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
        elif align_low == "right":
            table.alignment = WD_TABLE_ALIGNMENT.RIGHT
        else:
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = auto_fit if not col_widths else False

        current_row_idx = 0
        if headers:
            hdr_row = table.rows[0]
            trPr = hdr_row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            for c_idx, text in enumerate(headers):
                cell = hdr_row.cells[c_idx]
                if header_bg_color:
                    set_cell_background(cell, header_bg_color)
                set_cell_margins(cell, top_dpt=120, bottom_dpt=120, left_dpt=140, right_dpt=140)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if is_arabic_text(str(text)):
                    set_paragraph_rtl(p, True)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                run = p.add_run(str(text))
                set_run_font_name(run, "Segoe UI")
                run.font.size = Pt(10.5)
                run.font.bold = True
                if header_font_color:
                    run.font.color.rgb = hex_to_rgb_color(header_font_color)

            current_row_idx = 1

        for r_i, row_values in enumerate(data):
            row = table.rows[current_row_idx + r_i]
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            is_even = (r_i % 2 == 1)
            row_bg = alternate_row_color if (is_even and alternate_row_color) else None

            for c_i in range(num_cols):
                cell = row.cells[c_i]
                val = row_values[c_i] if c_i < len(row_values) else ""
                val_str = "" if val is None else str(val)

                if row_bg:
                    set_cell_background(cell, row_bg)
                set_cell_margins(cell, top_dpt=100, bottom_dpt=100, left_dpt=120, right_dpt=120)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)

                try:
                    num = float(val_str.replace("$", "").replace(",", "").replace("%", ""))
                    is_num = math.isfinite(num) and bool(val_str.strip())
                except ValueError:
                    is_num = False
                if is_num:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif is_arabic_text(val_str):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_paragraph_rtl(p, True)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                run = p.add_run(val_str)
                set_run_font_name(run, "Segoe UI")
                run.font.size = Pt(9.5)
                run.font.color.rgb = hex_to_rgb_color("#262626")

        if col_widths:
            tblPr = table._tbl.tblPr
            existing_layout = tblPr.find(qn("w:tblLayout"))
            if existing_layout is not None:
                tblPr.remove(existing_layout)
            tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
            for c_i, w in enumerate(col_widths):
                if c_i < len(table.columns):
                    table.columns[c_i].width = Inches(w)
            for row in table.rows:
                for c_i, w in enumerate(col_widths):
                    if c_i < len(row.cells):
                        row.cells[c_i].width = Inches(w)

        for r1, c1, r2, c2 in validated_merges:
            table.cell(r1, c1).merge(table.cell(r2, c2))

        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(6)
        spacer.paragraph_format.space_after = Pt(6)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "table_added",
            "filepath": abs_path,
            "rows": num_rows,
            "cols": num_cols
        }

    @classmethod
    def add_image(
        cls,
        image_path: str,
        filepath: Optional[str] = None,
        width_inches: Optional[float] = None,
        height_inches: Optional[float] = None,
        caption: Optional[str] = None,
        align: str = "center"
    ) -> Dict[str, Any]:
        """Insert an image into the document."""
        abs_img_path = os.path.abspath(image_path)
        if not os.path.exists(abs_img_path):
            raise FileNotFoundError(f"Image file not found: {abs_img_path}")

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc)
                    shape = doc.InlineShapes.AddPicture(
                        FileName=abs_img_path,
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Range=rng,
                    )
                    if width_inches and height_inches:
                        try:
                            shape.LockAspectRatio = 0  # msoFalse
                        except Exception:
                            pass
                        shape.Width = width_inches * 72
                        shape.Height = height_inches * 72
                    elif width_inches:
                        try:
                            shape.LockAspectRatio = -1  # msoTrue
                        except Exception:
                            pass
                        shape.Width = width_inches * 72
                    elif height_inches:
                        try:
                            shape.LockAspectRatio = -1
                        except Exception:
                            pass
                        shape.Height = height_inches * 72
                    else:
                        try:
                            shape.LockAspectRatio = -1
                            if shape.Width > 5.5 * 72:
                                shape.Width = 5.5 * 72
                        except Exception:
                            pass
                    shape.Range.ParagraphFormat.Alignment = get_com_alignment(align)
                    if caption:
                        cap_rng = LiveWordSession.com_prepare_range(doc)
                        cap_rng.ParagraphFormat.Alignment = get_com_alignment(align)
                        cap_rng.Font.Name = "Segoe UI"
                        cap_rng.Font.Italic = True
                        cap_rng.Font.Size = 9
                        cap_rng.Font.Color = hex_to_bgr_int("#595959")
                        is_rtl = is_arabic_text(caption)
                        try:
                            cap_rng.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                        except Exception:
                            pass
                        prefix = "الشكل: " if is_rtl else "Figure: "
                        cap_rng.Text = f"{prefix}{caption}"
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "image_added_live",
                        "doc_name": doc.Name,
                        "image": abs_img_path
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        p = doc.add_paragraph()
        p.alignment = get_docx_alignment(align) or WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)

        w = Inches(width_inches) if width_inches else None
        h = Inches(height_inches) if height_inches else None

        run_img = p.add_run()
        if w and h:
            run_img.add_picture(abs_img_path, width=w, height=h)
        elif w:
            run_img.add_picture(abs_img_path, width=w)
        elif h:
            run_img.add_picture(abs_img_path, height=h)
        else:
            run_img.add_picture(abs_img_path, width=Inches(5.5))

        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = get_docx_alignment(align) or WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(10)
            is_rtl = is_arabic_text(caption)
            if is_rtl:
                set_paragraph_rtl(p_cap, True)
            prefix = "الشكل: " if is_rtl else "Figure: "
            run_cap = p_cap.add_run(f"{prefix}{caption}")
            set_run_font_name(run_cap, "Segoe UI")
            run_cap.font.size = Pt(9)
            run_cap.font.italic = True
            run_cap.font.color.rgb = hex_to_rgb_color("#595959")

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "image_added",
            "filepath": abs_path,
            "image": abs_img_path,
            "caption": caption
        }

    @classmethod
    def add_page_break(cls, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Insert a page break."""
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc)
                    rng.InsertBreak(Type=7)
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {"mode": "live_com", "status": "page_break_added_live", "doc_name": doc.Name}
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()
        doc.add_page_break()
        doc.save(abs_path)
        return {"mode": "headless_docx", "status": "page_break_added", "filepath": abs_path}

    @classmethod
    def _com_replace_long_in_story(
        cls,
        doc,
        cur,
        find_text: str,
        replace_text: str,
        match_case: bool,
        match_whole_word: bool,
        replace_all: bool,
    ) -> int:
        """Search and replace text longer than Word Find's 255-character limit in a COM story range."""
        norm_find = find_text.replace("^p", "\r").replace("\r\n", "\r").replace("\n", "\r")
        flags = 0 if match_case else re.IGNORECASE
        pat_str = rf"\b{re.escape(norm_find)}\b" if match_whole_word else re.escape(norm_find)
        pattern = re.compile(pat_str, flags)

        anchor = norm_find.split("\r", 1)[0][:200]
        if not anchor:
            # Pure multi-newline or paragraph-break pattern: scan story text directly
            story_txt = cur.Text or ""
            matches = list(pattern.finditer(story_txt))
            if not matches:
                return 0
            if not replace_all:
                matches = matches[:1]
            base = int(cur.Start)
            for m in reversed(matches):
                sub = doc.Range(base + m.start(), base + m.end())
                if replace_text == "":
                    sub.Delete()
                else:
                    sub.Text = replace_text
            return len(matches)

        count = 0
        rng = cur.Duplicate
        find_obj = rng.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Text = anchor
        find_obj.Forward = True
        find_obj.Wrap = 0  # wdFindStop
        find_obj.MatchCase = match_case
        find_obj.MatchWholeWord = False

        while find_obj.Execute():
            m_start = int(rng.Start)
            probe_end = min(int(cur.End), m_start + len(norm_find) + 64)
            probe = doc.Range(m_start, probe_end)
            probe_txt = probe.Text or ""
            m = pattern.match(probe_txt)
            if m:
                exact_end = min(int(cur.End), m_start + m.end())
                target = doc.Range(m_start, exact_end)
                # Adjust end slightly if embedded field/control chars shifted character count
                if (target.Text or "") != probe_txt[:m.end()]:
                    while exact_end < int(cur.End) and len(target.Text or "") < m.end():
                        exact_end += 1
                        target = doc.Range(m_start, exact_end)
                if replace_text == "":
                    target.Delete()
                    next_pos = m_start
                else:
                    target.Text = replace_text
                    next_pos = int(target.End)
                count += 1
                if not replace_all:
                    break
            else:
                next_pos = int(rng.End)
            story_end = int(cur.End)
            if next_pos >= story_end:
                break
            rng.SetRange(next_pos, story_end)
        return count

    @classmethod
    def search_and_replace(
        cls,
        find_text: str,
        replace_text: str,
        filepath: Optional[str] = None,
        match_case: bool = False,
        match_whole_word: bool = False,
        replace_all: bool = True
    ) -> Dict[str, Any]:
        """Search and replace text across document (body, tables, headers, and footers)."""
        if not find_text:
            raise ValueError("`find_text` must not be empty.")

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    replace_count_live = 0
                    for story in doc.StoryRanges:
                        cur = story
                        while cur is not None:
                            if len(find_text) > 255:
                                replace_count_live += cls._com_replace_long_in_story(
                                    doc, cur, find_text, replace_text, match_case, match_whole_word, replace_all
                                )
                            else:
                                rng = cur.Duplicate
                                find_obj = rng.Find
                                find_obj.ClearFormatting()
                                find_obj.Replacement.ClearFormatting()
                                find_obj.Text = find_text
                                find_obj.Forward = True
                                find_obj.Wrap = 0  # wdFindStop
                                find_obj.MatchCase = match_case
                                find_obj.MatchWholeWord = match_whole_word
                                while find_obj.Execute():
                                    if replace_text == "":
                                        rng.Delete()
                                    else:
                                        rng.Text = replace_text
                                    replace_count_live += 1
                                    if not replace_all:
                                        break
                                    next_pos = rng.Start if replace_text == "" else rng.End
                                    story_end = cur.End
                                    if next_pos >= story_end:
                                        break
                                    rng.SetRange(next_pos, story_end)
                            if replace_count_live > 0 and not replace_all:
                                break
                            try:
                                cur = cur.NextStoryRange
                            except Exception:
                                cur = None
                        if replace_count_live > 0 and not replace_all:
                            break

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "replaced_live",
                        "doc_name": doc.Name,
                        "find_text": find_text,
                        "replace_text": replace_text,
                        "replacements_count": replace_count_live,
                        "success": replace_count_live > 0
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        doc = docx.Document(abs_path)
        replace_count = 0

        escaped_find = re.escape(find_text)
        pattern_str = rf'\b{escaped_find}\b' if match_whole_word else escaped_find
        flags = 0 if match_case else re.IGNORECASE
        regex = re.compile(pattern_str, flags)

        def replace_in_paragraphs(paras):
            nonlocal replace_count
            for p in paras:
                if not replace_all and replace_count >= 1:
                    return
                remaining = None if replace_all else (1 - replace_count)
                n = _replace_in_paragraph_runs(p, regex, lambda _m: replace_text, max_replacements=remaining)
                replace_count += n

        replace_in_paragraphs(doc.paragraphs)
        for tbl in doc.tables:
            if not replace_all and replace_count >= 1:
                break
            replace_in_paragraphs(_iter_table_paragraphs(tbl))
        for hf in _iter_section_headers_footers(doc):
            if not replace_all and replace_count >= 1:
                break
            replace_in_paragraphs(hf.paragraphs)
            for tbl in hf.tables:
                if not replace_all and replace_count >= 1:
                    break
                replace_in_paragraphs(_iter_table_paragraphs(tbl))

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "replaced",
            "filepath": abs_path,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements_count": replace_count
        }

    # ==========================================
    # Advanced Enterprise Tools (1-6)
    # ==========================================

    @classmethod
    def set_header_footer(
        cls,
        header_text: Optional[str] = None,
        footer_text: Optional[str] = None,
        include_page_number: bool = True,
        different_first_page: bool = False,
        header_align: str = "right",
        footer_align: str = "center",
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Set custom running header, footer, and dynamic page numbering.
        """
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    for s_idx in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(s_idx)
                        sec.PageSetup.DifferentFirstPageHeaderFooter = different_first_page

                        if header_text is not None:
                            hdr = sec.Headers(1)
                            p_rng = hdr.Range.Paragraphs(1).Range
                            if p_rng.End > p_rng.Start and p_rng.Text.endswith("\r"):
                                p_rng.End = p_rng.End - 1
                            p_rng.Text = header_text
                            p_rng.ParagraphFormat.Alignment = get_com_alignment(header_align)
                            is_rtl = is_arabic_text(header_text)
                            try:
                                p_rng.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                            except Exception:
                                pass
                            p_rng.Font.Name = "Segoe UI"
                            p_rng.Font.Size = 9
                            p_rng.Font.Color = hex_to_bgr_int("#7F7F7F")

                        if footer_text is not None or include_page_number:
                            ftr = sec.Footers(1)
                            ftr.Range.Text = (footer_text + (" — " if include_page_number else "")) if footer_text else ""
                            ftr.Range.ParagraphFormat.Alignment = get_com_alignment(footer_align)
                            if footer_text and is_arabic_text(footer_text):
                                try:
                                    ftr.Range.ParagraphFormat.ReadingOrder = 0
                                except Exception:
                                    pass
                            ftr.Range.Font.Name = "Segoe UI"
                            ftr.Range.Font.Size = 9
                            ftr.Range.Font.Color = hex_to_bgr_int("#7F7F7F")
                            if include_page_number:
                                end_rng = ftr.Range.Paragraphs(1).Range
                                if end_rng.End > end_rng.Start and end_rng.Text.endswith("\r"):
                                    end_rng.End = end_rng.End - 1
                                end_rng.Collapse(0)  # wdCollapseEnd
                                ftr.Range.Fields.Add(Range=end_rng, Type=33, PreserveFormatting=True)  # wdFieldPage

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "header_footer_set_live",
                        "doc_name": doc.Name,
                        "header": header_text,
                        "footer": footer_text,
                        "page_number": include_page_number
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        for s_idx, section in enumerate(doc.sections):
            section.different_first_page_header_footer = different_first_page

            if header_text is not None:
                hdr = section.header
                if s_idx == 0 or not hdr.is_linked_to_previous:
                    p_h = hdr.paragraphs[0]
                    p_h.alignment = get_docx_alignment(header_align) or WD_ALIGN_PARAGRAPH.RIGHT
                    set_paragraph_rtl(p_h, is_arabic_text(header_text))
                    p_h.text = ""
                    r_h = p_h.add_run(header_text)
                    set_run_font_name(r_h, "Segoe UI")
                    r_h.font.size = Pt(9)
                    r_h.font.color.rgb = hex_to_rgb_color("#7F7F7F")

            if footer_text is not None or include_page_number:
                ftr = section.footer
                if s_idx == 0 or not ftr.is_linked_to_previous:
                    p_f = ftr.paragraphs[0]
                    p_f.alignment = get_docx_alignment(footer_align) or WD_ALIGN_PARAGRAPH.CENTER
                    set_paragraph_rtl(p_f, bool(footer_text and is_arabic_text(footer_text)))
                    for child in list(p_f._p):
                        if child.tag != qn("w:pPr"):
                            p_f._p.remove(child)
                    if footer_text:
                        r_f = p_f.add_run(footer_text + (" — " if include_page_number else ""))
                        set_run_font_name(r_f, "Segoe UI")
                        r_f.font.size = Pt(9)
                        r_f.font.color.rgb = hex_to_rgb_color("#7F7F7F")

                    if include_page_number:
                        fld = parse_xml(
                            f'<w:fldSimple {nsdecls("w")} w:instr="PAGE">'
                            f'<w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr><w:t>1</w:t></w:r>'
                            f'</w:fldSimple>'
                        )
                        p_f._p.append(fld)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "header_footer_set",
            "filepath": abs_path,
            "header": header_text,
            "footer": footer_text,
            "page_number": include_page_number
        }

    @classmethod
    def insert_toc(
        cls,
        filepath: Optional[str] = None,
        title: str = "جدول المحتويات",
        upper_heading_level: int = 1,
        lower_heading_level: int = 3,
        position: str = "start"
    ) -> Dict[str, Any]:
        """
        Insert an automated Table of Contents (TOC) with page references and heading levels.
        """
        upper_heading_level = max(1, min(9, int(upper_heading_level)))
        lower_heading_level = max(upper_heading_level, min(9, int(lower_heading_level)))

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    rng = LiveWordSession.com_prepare_range(doc, position=position)

                    if title:
                        rng.Text = title
                        apply_com_style(rng, "Heading 1")
                        rng.Font.Name = "Segoe UI"
                        rng.Font.Color = hex_to_bgr_int("#1F497D")
                        is_rtl = is_arabic_text(title)
                        rng.ParagraphFormat.Alignment = 2 if is_rtl else 0
                        try:
                            rng.ParagraphFormat.ReadingOrder = 0 if is_rtl else 1
                        except Exception:
                            pass
                        rng.Collapse(0)  # wdCollapseEnd
                        rng.InsertAfter("\r")
                        rng.Collapse(0)
                        apply_com_style(rng, "Normal")

                    toc = doc.TablesOfContents.Add(
                        Range=rng,
                        UseHeadingStyles=True,
                        UpperHeadingLevel=upper_heading_level,
                        LowerHeadingLevel=lower_heading_level,
                        RightAlignPageNumbers=True
                    )
                    toc.Update()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "toc_inserted_live",
                        "doc_name": doc.Name,
                        "title": title
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()
        h1_style = resolve_docx_style(doc, "Heading 1")

        # Build TOC paragraph
        if position == "start" and len(doc.paragraphs) > 0:
            first_p = doc.paragraphs[0]
            if title:
                p_h = first_p.insert_paragraph_before(title, style=h1_style)
                if is_arabic_text(title):
                    p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_paragraph_rtl(p_h, True)
                if p_h.runs:
                    set_run_font_name(p_h.runs[0], "Segoe UI")
                    p_h.runs[0].font.color.rgb = hex_to_rgb_color("#1F497D")
            p_toc = first_p.insert_paragraph_before()
            fldSimple = parse_xml(
                f'<w:fldSimple {nsdecls("w")} w:dirty="true" '
                f'w:instr="TOC \\o &quot;{upper_heading_level}-{lower_heading_level}&quot; \\h \\z \\u"/>'
            )
            p_toc._p.append(fldSimple)
            first_p.insert_paragraph_before()  # Spacing
        else:
            if title:
                h = doc.add_heading(title, level=1)
                if is_arabic_text(title):
                    h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_paragraph_rtl(h, True)
                if h.runs:
                    set_run_font_name(h.runs[0], "Segoe UI")
                    h.runs[0].font.color.rgb = hex_to_rgb_color("#1F497D")
            p = doc.add_paragraph()
            fldSimple = parse_xml(
                f'<w:fldSimple {nsdecls("w")} w:dirty="true" '
                f'w:instr="TOC \\o &quot;{upper_heading_level}-{lower_heading_level}&quot; \\h \\z \\u"/>'
            )
            p._p.append(fldSimple)
            doc.add_paragraph()

        settings_el = doc.settings._element
        if settings_el.find(qn("w:updateFields")) is None:
            settings_el.append(parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>'))

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "toc_inserted",
            "filepath": abs_path,
            "title": title,
            "note": "Table of Contents field inserted; Word will auto-render it when opened."
        }

    @classmethod
    def fill_template(
        cls,
        template_path: str,
        output_path: str,
        replacements: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Fill a Word template (.docx / .dotx) by replacing {{placeholders}} with dynamic values.
        """
        abs_tpl = os.path.abspath(template_path)
        abs_out = os.path.abspath(output_path)
        if not os.path.exists(abs_tpl):
            raise FileNotFoundError(f"Template file not found: {abs_tpl}")

        if LiveWordSession.should_use_live():
            attached = LiveWordSession._attach()
            if attached is not None and LiveWordSession.find_open_document(attached, abs_out) is not None:
                raise LiveWordError(
                    f"Output document '{abs_out}' is currently open in Microsoft Word; refusing to overwrite on disk."
                )

        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        doc = _open_docx_or_dotx(abs_tpl)

        total_replacements = 0
        if replacements:
            tag_map: Dict[str, str] = {}
            tag_variants: List[str] = []
            for k, v in replacements.items():
                val_str = str(v)
                for tag in (f"{{{{{k}}}}}", f"{{{k}}}"):
                    tag_map[tag] = val_str
                    tag_variants.append(tag)
            # Match longer tags and {{key}} before {key} in a single pass to prevent chained substitutions
            tag_variants.sort(key=len, reverse=True)
            pattern = re.compile("|".join(re.escape(t) for t in tag_variants))

            def replace_in_paras(paras) -> int:
                count = 0
                for p in paras:
                    count += _replace_in_paragraph_runs(p, pattern, lambda m: tag_map[m.group(0)])
                return count

            total_replacements += replace_in_paras(doc.paragraphs)
            for tbl in doc.tables:
                total_replacements += replace_in_paras(_iter_table_paragraphs(tbl))
            for hf in _iter_section_headers_footers(doc):
                total_replacements += replace_in_paras(hf.paragraphs)
                for tbl in hf.tables:
                    total_replacements += replace_in_paras(_iter_table_paragraphs(tbl))

        doc.save(abs_out)
        return {
            "status": "template_rendered",
            "template": abs_tpl,
            "output": abs_out,
            "replacements_applied": total_replacements,
            "keys": list(replacements.keys())
        }

    @classmethod
    def add_comment(
        cls,
        comment_text: str,
        target_text: Optional[str] = None,
        paragraph_index: Optional[int] = None,
        author: str = "AI Reviewer",
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add a review comment to a specific phrase or paragraph.
        """
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    target_range = None
                    if target_text:
                        find_obj = doc.Content.Find
                        find_obj.ClearFormatting()
                        find_obj.Replacement.ClearFormatting()
                        find_obj.Forward = True
                        find_obj.Wrap = 1  # wdFindContinue
                        find_obj.Text = target_text
                        if find_obj.Execute():
                            target_range = find_obj.Parent

                    if target_range is None and paragraph_index and 1 <= paragraph_index <= doc.Paragraphs.Count:
                        target_range = doc.Paragraphs(paragraph_index).Range

                    if target_range is None:
                        target_range = doc.Content.Paragraphs(1).Range if doc.Content.Paragraphs.Count > 0 else doc.Content

                    old_username = getattr(app, "UserName", "")
                    try:
                        app.UserName = author
                        doc.Comments.Add(Range=target_range, Text=comment_text)
                    finally:
                        if old_username:
                            try:
                                app.UserName = old_username
                            except Exception:
                                pass
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "comment_added_live",
                        "doc_name": doc.Name,
                        "author": author,
                        "comment": comment_text,
                        "target_text": target_text
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        # Locate insertion target (fall back to paragraph_index if target_text not found)
        target_para = None
        if target_text:
            for p in doc.paragraphs:
                if target_text in p.text:
                    target_para = p
                    break
        if target_para is None and paragraph_index and 1 <= paragraph_index <= len(doc.paragraphs):
            target_para = doc.paragraphs[paragraph_index - 1]

        if hasattr(doc, "add_comment"):
            if target_para is None:
                target_para = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph("")
            runs = list(target_para.runs)
            if target_text:
                matching = [r for r in runs if target_text in (r.text or "")]
                if matching:
                    runs = matching
            if not runs:
                runs = [target_para.add_run("")]
            initials = "".join(part[0] for part in author.split() if part)[:3].upper() or "AI"
            doc.add_comment(runs=runs, text=comment_text, author=author, initials=initials)
            doc.save(abs_path)
            return {
                "mode": "headless_docx",
                "status": "comment_added",
                "filepath": abs_path,
                "author": author,
                "comment": comment_text,
                "target_text": target_text
            }

        # Fallback annotation for python-docx < 1.2.0
        if target_para:
            p_note = target_para.insert_paragraph_before()
        else:
            p_note = doc.add_paragraph()

        p_note.paragraph_format.space_before = Pt(4)
        p_note.paragraph_format.space_after = Pt(4)
        r_author = p_note.add_run(f"[{author}]: ")
        r_author.font.bold = True
        r_author.font.size = Pt(9.5)
        r_author.font.color.rgb = hex_to_rgb_color("#C00000")

        r_txt = p_note.add_run(comment_text)
        r_txt.font.italic = True
        r_txt.font.size = Pt(9.5)
        r_txt.font.color.rgb = hex_to_rgb_color("#595959")

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "comment_annotated",
            "filepath": abs_path,
            "author": author,
            "comment": comment_text
        }

    @classmethod
    def set_page_setup(
        cls,
        filepath: Optional[str] = None,
        orientation: Optional[str] = None,
        paper_size: Optional[str] = None,
        margin_preset: Optional[str] = None,
        top_margin_inches: Optional[float] = None,
        bottom_margin_inches: Optional[float] = None,
        left_margin_inches: Optional[float] = None,
        right_margin_inches: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Configure page layout, orientation (Portrait/Landscape), paper size (A4/Letter/Legal), and margins.
        """
        margin_presets = {
            "normal": (1.0, 1.0, 1.0, 1.0),
            "narrow": (0.5, 0.5, 0.5, 0.5),
            "moderate": (1.0, 1.0, 0.75, 0.75),
            "wide": (1.0, 1.0, 2.0, 2.0)
        }
        if margin_preset and margin_preset.lower() in margin_presets:
            p_top, p_bot, p_left, p_right = margin_presets[margin_preset.lower()]
            if top_margin_inches is None:
                top_margin_inches = p_top
            if bottom_margin_inches is None:
                bottom_margin_inches = p_bot
            if left_margin_inches is None:
                left_margin_inches = p_left
            if right_margin_inches is None:
                right_margin_inches = p_right

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    for s_idx in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(s_idx)
                        if paper_size:
                            ps = paper_size.lower()
                            if ps == "a4":
                                sec.PageSetup.PaperSize = 7  # wdPaperA4
                            elif ps in ("letter", "us_letter"):
                                sec.PageSetup.PaperSize = 2  # wdPaperLetter
                            elif ps == "legal":
                                sec.PageSetup.PaperSize = 4  # wdPaperLegal
                        if orientation:
                            sec.PageSetup.Orientation = 1 if orientation.lower() in ("landscape", "horizontal") else 0
                        if top_margin_inches is not None:
                            sec.PageSetup.TopMargin = top_margin_inches * 72
                        if bottom_margin_inches is not None:
                            sec.PageSetup.BottomMargin = bottom_margin_inches * 72
                        if left_margin_inches is not None:
                            sec.PageSetup.LeftMargin = left_margin_inches * 72
                        if right_margin_inches is not None:
                            sec.PageSetup.RightMargin = right_margin_inches * 72

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "page_setup_updated_live",
                        "doc_name": doc.Name,
                        "orientation": orientation,
                        "paper_size": paper_size
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        for sec in doc.sections:
            if paper_size:
                ps = paper_size.lower()
                if ps == "a4":
                    sec.page_width = Inches(8.27)
                    sec.page_height = Inches(11.69)
                elif ps in ("letter", "us_letter"):
                    sec.page_width = Inches(8.5)
                    sec.page_height = Inches(11.0)
                elif ps == "legal":
                    sec.page_width = Inches(8.5)
                    sec.page_height = Inches(14.0)

            if orientation:
                is_landscape = orientation.lower() in ("landscape", "horizontal")
                sec.orientation = WD_ORIENT.LANDSCAPE if is_landscape else WD_ORIENT.PORTRAIT
                if is_landscape and sec.page_width < sec.page_height:
                    sec.page_width, sec.page_height = sec.page_height, sec.page_width
                elif not is_landscape and sec.page_width > sec.page_height:
                    sec.page_width, sec.page_height = sec.page_height, sec.page_width

            if top_margin_inches is not None:
                sec.top_margin = Inches(top_margin_inches)
            if bottom_margin_inches is not None:
                sec.bottom_margin = Inches(bottom_margin_inches)
            if left_margin_inches is not None:
                sec.left_margin = Inches(left_margin_inches)
            if right_margin_inches is not None:
                sec.right_margin = Inches(right_margin_inches)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "page_setup_updated",
            "filepath": abs_path,
            "orientation": orientation,
            "paper_size": paper_size
        }

    @classmethod
    def add_watermark(
        cls,
        text: str,
        filepath: Optional[str] = None,
        color: str = "#D3D3D3",
        font_name: str = "Segoe UI",
        font_size: int = 54,
        diagonal: bool = True
    ) -> Dict[str, Any]:
        """
        Add a background text watermark (e.g. 'CONFIDENTIAL', 'DRAFT', 'مسودة').
        """
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sec = doc.Sections(1)
                    hdr = sec.Headers(1)
                    for i in range(hdr.Shapes.Count, 0, -1):
                        try:
                            if hdr.Shapes(i).Name == "WordWatermarkObject":
                                hdr.Shapes(i).Delete()
                        except Exception:
                            pass
                    if hdr.Range.Paragraphs.Count < 2:
                        hdr.Range.InsertAfter("\r")
                    anchor_rng = hdr.Range.Paragraphs.Last.Range
                    shape = hdr.Shapes.AddTextEffect(0, text, font_name, font_size, False, False, 0, 0, anchor_rng)
                    shape.Name = "WordWatermarkObject"
                    shape.Fill.Solid()
                    shape.Fill.ForeColor.RGB = hex_to_bgr_int(color)
                    try:
                        shape.Fill.Transparency = 0.7
                    except Exception:
                        pass
                    shape.Line.Visible = False
                    if diagonal:
                        shape.Rotation = -45
                    try:
                        shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionMargin
                        shape.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionMargin
                        shape.Left = -999995  # wdShapeCenter
                        shape.Top = -999995   # wdShapeCenter
                        shape.WrapFormat.Type = 3  # wdWrapNone
                        shape.ZOrder(5)       # msoSendBehindText
                    except Exception:
                        pass

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "watermark_added_live",
                        "doc_name": doc.Name,
                        "watermark_text": text
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        sec = doc.sections[0]
        header = sec.header

        # Remove any existing WordWatermarkObject paragraph before inserting
        for el in list(header._element):
            if el.xpath('.//*[@id="WordWatermarkObject"]'):
                header._element.remove(el)

        # Add high-fidelity VML watermark shape to header in OpenXML
        rot = "-45" if diagonal else "0"
        hex_c = clean_hex(color)
        safe_text = xml_escape(text, {'"': '&quot;', "'": '&apos;'})
        safe_font = xml_escape(font_name, {'"': '&quot;', "'": '&apos;'})
        watermark_xml = (
            f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'
            f'<w:pPr><w:pStyle w:val="Header"/></w:pPr>'
            f'<w:r>'
            f'<w:rPr><w:noProof/></w:rPr>'
            f'<w:pict>'
            f'<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m0,0l0,21600,21600,21600,21600,0xe">'
            f'<v:path textpathok="t" o:connecttype="rect"/>'
            f'<v:textpath on="t" fitshape="t"/>'
            f'<v:handles><v:h position="#BOTTOMRIGHT" xrange="0,21600"/></v:handles>'
            f'<o:lock v:ext="edit" text="t" shapetype="t"/>'
            f'</v:shapetype>'
            f'<v:shape id="WordWatermarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" '
            f'style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251656192;'
            f'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
            f'mso-position-vertical:center;mso-position-vertical-relative:margin;rotation:{rot}" '
            f'fillcolor="#{hex_c}" stroked="f">'
            f'<v:fill opacity="0.2"/>'
            f'<v:textpath string="{safe_text}" style="font-family:\'{safe_font}\';font-size:{int(font_size)}pt;font-weight:bold"/>'
            f'</v:shape>'
            f'</w:pict>'
            f'</w:r>'
            f'</w:p>'
        )
        try:
            wm_elem = parse_xml(watermark_xml)
            header._element.append(wm_elem)
        except Exception:
            p_fb = header.add_paragraph()
            p_fb.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_fb.add_run(f"[{text.upper()}]")
            set_run_font_name(r, font_name)
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = hex_to_rgb_color(color)

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "watermark_added",
            "filepath": abs_path,
            "watermark_text": text
        }

    # ==========================================
    # PDF Export, Save & Script Execution
    # ==========================================

    @classmethod
    def export_pdf(
        cls,
        output_pdf_path: str,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Export Word document to high-fidelity PDF."""
        abs_pdf_path = os.path.abspath(output_pdf_path)
        os.makedirs(os.path.dirname(abs_pdf_path), exist_ok=True)

        if not COM_AVAILABLE:
            raise LiveWordError("Live Microsoft Word COM (pywin32) is required to export high-fidelity PDF documents.")

        # Attach to the user's Word if there is one. A hidden instance is launched only
        # when NO Word process exists, and only that instance is ever quit afterwards.
        app = LiveWordSession.get_app(create_if_missing=True, visible=False, hidden_ok=True)
        launched_here = LiveWordSession.launched_by_us()
        doc = None
        doc_opened_here = False

        try:
            if filepath:
                abs_path = os.path.abspath(filepath)
                doc = LiveWordSession.find_open_document(app, abs_path)
                if not doc:
                    if not os.path.exists(abs_path):
                        raise FileNotFoundError(f"Document file not found: {abs_path}")
                    doc = com_retry(
                        app.Documents.Open, abs_path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
                    )
                    doc_opened_here = True
            else:
                if app.Documents.Count > 0:
                    doc = app.ActiveDocument
                else:
                    raise ValueError("No filepath specified and no active document in Microsoft Word.")

            doc.ExportAsFixedFormat(
                OutputFileName=abs_pdf_path,
                ExportFormat=17,  # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,     # wdExportOptimizeForPrint
                Range=0,          # wdExportAllDocument
                Item=7            # wdExportDocumentWithMarkup
            )

            return {
                "status": "exported_pdf",
                "pdf_path": abs_pdf_path,
                "source_doc": getattr(doc, "Name", "document")
            }
        finally:
            if doc_opened_here and doc:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            if launched_here and app:
                try:
                    # Never close a Word the user could be using: only our own, now-empty instance.
                    if app.Documents.Count == 0:
                        app.Quit(SaveChanges=0)
                except Exception:
                    pass

    @classmethod
    def save_document(
        cls,
        save_as_path: Optional[str] = None,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Save the document or save it to a new location."""
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    if save_as_path:
                        abs_save = os.path.abspath(save_as_path)
                        os.makedirs(os.path.dirname(abs_save), exist_ok=True)
                        doc.SaveAs2(abs_save)
                        return {"mode": "live_com", "status": "saved_as", "filepath": abs_save}
                    else:
                        if not getattr(doc, "Path", ""):
                            raise ValueError(
                                "Active document has never been saved. Provide `save_as_path` to save it to disk."
                            )
                        doc.Save()
                        return {"mode": "live_com", "status": "saved", "filepath": getattr(doc, "FullName", doc.Name)}
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        if filepath:
            abs_path = os.path.abspath(filepath)
            if not os.path.exists(abs_path):
                raise FileNotFoundError(f"Document file not found: {abs_path}")
            doc = docx.Document(abs_path)
            target = os.path.abspath(save_as_path) if save_as_path else abs_path
            os.makedirs(os.path.dirname(target), exist_ok=True)
            doc.save(target)
            return {"mode": "headless_docx", "status": "saved", "filepath": target}

        raise ValueError("Cannot save document: No active COM document and no filepath specified.")

    @classmethod
    def execute_script(
        cls,
        code: str,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Execute arbitrary Python automation code (disabled unless WORD_MCP_ENABLE_SCRIPT=1)."""
        if not script_execution_enabled():
            raise PermissionError(
                "word_execute_script is disabled. It runs arbitrary Python on this machine. "
                "Set the environment variable WORD_MCP_ENABLE_SCRIPT=1 for the MCP server to enable it."
            )
        app = None
        doc = None
        docx_doc = None

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        if filepath and os.path.exists(os.path.abspath(filepath)):
            try:
                docx_doc = docx.Document(os.path.abspath(filepath))
            except Exception:
                pass

        redirected_output = io.StringIO()

        execution_scope = {
            "app": app,
            "doc": doc,
            "docx_doc": docx_doc,
            "docx": docx,
            "win32com": win32com.client if COM_AVAILABLE else None,
            "engine": cls,
            "filepath": filepath,
            "os": os,
            "sys": sys,
            "json": json,
            "re": re,
            "Inches": Inches,
            "Pt": Pt,
            "RGBColor": RGBColor,
            "WD_ALIGN_PARAGRAPH": WD_ALIGN_PARAGRAPH,
            "WD_COLOR_INDEX": WD_COLOR_INDEX,
            "WD_TABLE_ALIGNMENT": WD_TABLE_ALIGNMENT,
            "WD_ALIGN_VERTICAL": WD_ALIGN_VERTICAL,
            "WD_SECTION": WD_SECTION,
            "WD_ORIENT": WD_ORIENT,
            "parse_xml": parse_xml,
            "OxmlElement": OxmlElement
        }

        try:
            # redirect_stdout always restores the real stdout (the MCP protocol channel),
            # even if the script raises or calls sys.exit().
            with contextlib.redirect_stdout(redirected_output):
                exec(code, execution_scope)
            return {
                "success": True,
                "output": redirected_output.getvalue() or "Script executed successfully without output."
            }
        except (Exception, SystemExit) as e:
            return {
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc(),
                "output": redirected_output.getvalue()
            }

    @classmethod
    def modify_paragraph(
        cls,
        paragraph_index: Optional[int] = None,
        target_text: Optional[str] = None,
        new_text: Optional[str] = None,
        delete: bool = False,
        filepath: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        font_size: Optional[float] = None,
        font_name: Optional[str] = None,
        font_color: Optional[str] = None,
        highlight_color: Optional[str] = None,
        align: Optional[str] = None,
        rtl: Optional[bool] = None,
        line_spacing: Optional[float] = None,
        space_before: Optional[float] = None,
        space_after: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Update text/formatting of an existing paragraph (by 1-based index or target_text) or delete it."""
        if paragraph_index is None and not target_text:
            raise ValueError("Provide either `paragraph_index` (1-based) or `target_text` to locate the paragraph.")

        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    target_para = None
                    matched_index = None
                    p_count = int(doc.Paragraphs.Count)
                    if paragraph_index is not None:
                        if not (1 <= paragraph_index <= p_count):
                            raise IndexError(f"paragraph_index {paragraph_index} out of range (1..{p_count}).")
                        target_para = doc.Paragraphs(paragraph_index)
                        matched_index = paragraph_index
                    else:
                        for idx in range(1, p_count + 1):
                            p_cand = doc.Paragraphs(idx)
                            if target_text in (p_cand.Range.Text or ""):
                                target_para = p_cand
                                matched_index = idx
                                break
                        if target_para is None:
                            raise ValueError(f"No paragraph found containing `{target_text}`.")

                    p_rng = target_para.Range
                    if delete:
                        p_rng.Delete()
                        if filepath and hasattr(doc, "Save"):
                            doc.Save()
                        return {
                            "mode": "live_com",
                            "status": "paragraph_deleted_live",
                            "doc_name": doc.Name,
                            "paragraph_index": matched_index,
                        }

                    # Keep trailing paragraph mark intact when modifying text/formatting
                    edit_rng = p_rng.Duplicate
                    if edit_rng.End > edit_rng.Start and (edit_rng.Text or "").endswith("\r"):
                        edit_rng.End = int(edit_rng.End) - 1

                    if new_text is not None:
                        edit_rng.Text = new_text

                    effective_rtl = rtl if rtl is not None else (is_arabic_text(edit_rng.Text or "") if new_text is not None else None)
                    if effective_rtl is not None:
                        try:
                            p_rng.ParagraphFormat.ReadingOrder = 0 if effective_rtl else 1
                        except Exception:
                            pass
                    if align is not None:
                        p_rng.ParagraphFormat.Alignment = get_com_alignment(align)
                    elif effective_rtl is True:
                        p_rng.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight
                    elif effective_rtl is False and align is None and rtl is not None:
                        p_rng.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft

                    if space_before is not None:
                        p_rng.ParagraphFormat.SpaceBefore = space_before
                    if space_after is not None:
                        p_rng.ParagraphFormat.SpaceAfter = space_after
                    if line_spacing is not None:
                        p_rng.ParagraphFormat.LineSpacingRule = 5
                        p_rng.ParagraphFormat.LineSpacing = line_spacing * 12

                    if bold is not None:
                        edit_rng.Font.Bold = bool(bold)
                    if italic is not None:
                        edit_rng.Font.Italic = bool(italic)
                    if underline is not None:
                        edit_rng.Font.Underline = 1 if underline else 0
                    if font_name is not None:
                        edit_rng.Font.Name = font_name
                    if font_size is not None:
                        edit_rng.Font.Size = font_size
                    if font_color is not None:
                        edit_rng.Font.Color = hex_to_bgr_int(font_color)
                    if highlight_color is not None:
                        hl_key = highlight_color.strip().upper().replace(" ", "_").replace("-", "_")
                        edit_rng.HighlightColorIndex = COM_HIGHLIGHT_MAP.get(hl_key, 7)

                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "paragraph_updated_live",
                        "doc_name": doc.Name,
                        "paragraph_index": matched_index,
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        doc = docx.Document(abs_path)
        target_para = None
        matched_index = None
        if paragraph_index is not None:
            if not (1 <= paragraph_index <= len(doc.paragraphs)):
                raise IndexError(f"paragraph_index {paragraph_index} out of range (1..{len(doc.paragraphs)}).")
            target_para = doc.paragraphs[paragraph_index - 1]
            matched_index = paragraph_index
        else:
            for idx, p_cand in enumerate(doc.paragraphs, 1):
                if target_text in p_cand.text:
                    target_para = p_cand
                    matched_index = idx
                    break
            if target_para is None:
                raise ValueError(f"No paragraph found containing `{target_text}`.")

        if delete:
            parent = target_para._element.getparent()
            if len(doc.paragraphs) <= 1:
                target_para.text = ""
            else:
                parent.remove(target_para._element)
            doc.save(abs_path)
            return {
                "mode": "headless_docx",
                "status": "paragraph_deleted",
                "filepath": abs_path,
                "paragraph_index": matched_index,
            }

        if new_text is not None:
            # Preserve first run formatting if present, then replace text cleanly
            first_run_fmt = None
            if target_para.runs:
                r0 = target_para.runs[0]
                first_run_fmt = (
                    r0.bold,
                    r0.italic,
                    r0.underline,
                    r0.font.size,
                    r0.font.name,
                    r0.font.color.rgb if r0.font.color else None,
                    r0.font.highlight_color,
                )
            target_para.text = ""
            r_new = target_para.add_run(new_text)
            if first_run_fmt:
                b0, i0, u0, sz0, fn0, fc0, hl0 = first_run_fmt
                if b0 is not None:
                    r_new.bold = b0
                if i0 is not None:
                    r_new.italic = i0
                if u0 is not None:
                    r_new.underline = u0
                if sz0 is not None:
                    r_new.font.size = sz0
                if fn0:
                    set_run_font_name(r_new, fn0)
                if fc0 is not None:
                    r_new.font.color.rgb = fc0
                if hl0 is not None:
                    r_new.font.highlight_color = hl0

        effective_rtl = rtl if rtl is not None else (is_arabic_text(target_para.text) if new_text is not None else None)
        if effective_rtl is not None:
            set_paragraph_rtl(target_para, effective_rtl)
        if align is not None:
            target_para.alignment = get_docx_alignment(align)
        elif effective_rtl is True:
            target_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif effective_rtl is False and rtl is not None:
            target_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if space_before is not None:
            target_para.paragraph_format.space_before = Pt(space_before)
        if space_after is not None:
            target_para.paragraph_format.space_after = Pt(space_after)
        if line_spacing is not None:
            target_para.paragraph_format.line_spacing = line_spacing

        if not target_para.runs and any(x is not None for x in (bold, italic, underline, font_size, font_name, font_color, highlight_color)):
            target_para.add_run("")
        for r in target_para.runs:
            if bold is not None:
                r.bold = bold
            if italic is not None:
                r.italic = italic
            if underline is not None:
                r.underline = underline
            if font_size is not None:
                r.font.size = Pt(font_size)
            if font_name is not None:
                set_run_font_name(r, font_name)
            if font_color is not None:
                r.font.color.rgb = hex_to_rgb_color(font_color)
            if highlight_color is not None:
                hl = get_highlight_enum(highlight_color)
                if hl:
                    r.font.highlight_color = hl

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "paragraph_updated",
            "filepath": abs_path,
            "paragraph_index": matched_index,
        }

    @classmethod
    def clear_document(
        cls,
        filepath: Optional[str] = None,
        include_headers_footers: bool = False,
    ) -> Dict[str, Any]:
        """Clear all body content (and optionally headers/footers), leaving a blank page."""
        if LiveWordSession.should_use_live():
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    doc.Content.Delete()
                    try:
                        doc.Content.Font.Reset()
                        doc.Content.ParagraphFormat.Reset()
                        doc.Content.ParagraphFormat.ReadingOrder = 1
                    except Exception:
                        pass
                    if include_headers_footers:
                        for s_idx in range(1, doc.Sections.Count + 1):
                            sec = doc.Sections(s_idx)
                            for hf_idx in (1, 2, 3):
                                for hf in (sec.Headers(hf_idx), sec.Footers(hf_idx)):
                                    try:
                                        hf.Range.Delete()
                                    except Exception:
                                        pass
                                    try:
                                        for sh_idx in range(hf.Shapes.Count, 0, -1):
                                            hf.Shapes(sh_idx).Delete()
                                    except Exception:
                                        pass
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "cleared_live",
                        "doc_name": doc.Name,
                        "cleared_headers_footers": include_headers_footers,
                    }
            except Exception as _live_exc:
                LiveWordSession.handle_live_failure(_live_exc)

        abs_path = require_filepath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        doc = docx.Document(abs_path)
        body = doc.element.body
        for child in list(body):
            if child.tag != qn("w:sectPr"):
                body.remove(child)
        doc.add_paragraph()

        if include_headers_footers:
            for hf in _iter_section_headers_footers(doc):
                el = hf._element
                for child in list(el):
                    el.remove(child)
                el.append(OxmlElement("w:p"))

        doc.save(abs_path)
        return {
            "mode": "headless_docx",
            "status": "cleared",
            "filepath": abs_path,
            "cleared_headers_footers": include_headers_footers,
        }



# ==========================================
# 3. FastMCP Server & Tool Registrations
# ==========================================


def _job(fn, *args, **kwargs):
    """Runs on the single COM thread. Always hands DisplayAlerts back to the user afterwards."""
    try:
        return fn(*args, **kwargs)
    finally:
        try:
            LiveWordSession.restore_ui()
        except Exception:  # noqa: BLE001
            log.debug("restore_ui failed", exc_info=True)


def _run(fn, *args, **kwargs):
    """
    Execute an engine call on the dedicated COM thread (COM objects are apartment-threaded,
    and serialising calls keeps two agent requests from interleaving edits in one document).
    """
    return WORKER.submit(_job, fn, *args, **kwargs)

# Initialize FastMCP Server
mcp = FastMCP("word-mcp-server")


@mcp.tool()
def word_status() -> str:
    """
    Check the current status of Microsoft Word and active documents.
    Returns the execution mode, whether Word is running, version, open documents and the active document.
    """
    status = _run(WordEngine.get_status)
    mode = status.get("mode")
    if status["word_running"]:
        doc_names = [d["name"] for d in status["open_documents"]]
        return (
            f"✅ Microsoft Word (v{status.get('version')}) is RUNNING and reachable. Mode: {mode}\n"
            f"- Open Documents ({len(doc_names)}): {', '.join(doc_names) if doc_names else 'None'}\n"
            f"- Active Document: {status.get('active_document') or 'None'}\n"
            f"- Active Document Path: {status.get('active_document_path') or 'None'}\n"
            "- Edits are applied live inside Word (you will see them on screen)."
        )
    if status.get("word_process_found"):
        return (
            f"⚠️ A Word process exists but COM cannot attach to it. Mode: {mode}\n"
            f"- {status.get('hint')}\n"
            "- Live edits will fail with a clear error; files are NOT edited behind Word's back."
        )
    com = "Available" if status.get("com_available") else "NOT available (install pywin32 on Windows)"
    return (
        f"ℹ️ Microsoft Word is currently NOT running. Mode: {mode}\n"
        f"- COM automation: {com}.\n"
        "- Files that are not open in Word can be edited directly (python-docx)."
    )


@mcp.tool()
def word_create_document(
    filepath: Optional[str] = None,
    title: Optional[str] = None,
    subtitle: Optional[str] = None,
    author: Optional[str] = None,
    visible: bool = True
) -> str:
    """
    Create a new Microsoft Word document with optional initial title, subtitle, and author metadata block.

    Args:
        filepath: Optional path where document will be saved (e.g., 'E:/reports/Q1_Report.docx'). If omitted, creates an unsaved document in live Word.
        title: Main document title banner (e.g. 'Project Proposal & Technical Plan').
        subtitle: Subtitle or summary subtitle (e.g. 'Version 2.0 | Confidential').
        author: Author name or department.
        visible: Whether to make Word visible on screen when launched via COM (default: True).
    """
    try:
        res = _run(WordEngine.create_document,
            filepath=filepath,
            title=title,
            subtitle=subtitle,
            author=author,
            visible=visible
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error creating document: {str(e)}"


@mcp.tool()
def word_open_document(
    filepath: str,
    read_only: bool = False,
    visible: bool = True
) -> str:
    """
    Open an existing Word document (.docx, .doc, .rtf).

    Args:
        filepath: Full file path to the document file.
        read_only: Open the document in read-only mode (default: False).
        visible: Make Word window visible on screen (default: True).
    """
    try:
        res = _run(WordEngine.open_document, filepath=filepath, read_only=read_only, visible=visible)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error opening document: {str(e)}"


@mcp.tool()
def word_read_document(
    filepath: Optional[str] = None,
    output_format: str = "markdown",
    include_tables: bool = True,
    max_paragraphs: Optional[int] = None
) -> str:
    """
    Read text and structure from a Word document or the active document.

    Args:
        filepath: Optional file path. If omitted, reads from active document in live Word.
        output_format: Format of returned content: 'markdown' (default, easiest for LLM to digest), 'text', or 'json'.
        include_tables: Whether to parse and include tables (default: True).
        max_paragraphs: Maximum number of paragraphs to read (useful for long docs).
    """
    try:
        res = _run(WordEngine.read_document,
            filepath=filepath,
            output_format=output_format,
            include_tables=include_tables,
            max_paragraphs=max_paragraphs
        )
        if output_format == "markdown" and "markdown" in res:
            file_label = res.get("filepath") or res.get("document_name") or "Active Document"
            header = (
                f"**File:** `{file_label}` | **Paragraphs:** {res['total_paragraphs']} | "
                f"**Tables:** {res['total_tables']} | **Words:** {res['total_words']}\n\n---\n\n"
            )
            return header + res["markdown"]
        elif output_format == "text" and "text" in res:
            return res["text"]
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error reading document: {str(e)}"


@mcp.tool()
def word_get_outline(
    filepath: Optional[str] = None
) -> str:
    """
    Extract the document hierarchy / table of contents outline (Headings 1-6, paragraph indices, and statistics).

    Args:
        filepath: Optional file path. If omitted, inspects active document in live Word.
    """
    try:
        res = _run(WordEngine.get_outline, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error getting outline: {str(e)}"


@mcp.tool()
def word_add_heading(
    text: str,
    level: int = 1,
    filepath: Optional[str] = None,
    align: Optional[str] = None,
    color: Optional[str] = None,
    position: Optional[str] = None,
    paragraph_index: Optional[int] = None,
) -> str:
    """
    Add a heading (Level 1 to 6) to the document.

    Args:
        text: Heading text.
        level: Heading level (1 to 6). Default: 1.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        align: Optional text alignment ('left', 'center', 'right'). Auto-detects RTL/right for Arabic if omitted.
        color: Optional hex color for heading text (e.g. '#1F497D', '#2B2B2B').
        position: Optional insertion position: 'end' (default) or 'start'.
        paragraph_index: Optional 1-based paragraph index to insert before.
    """
    try:
        res = _run(WordEngine.add_heading,
            text=text,
            level=level,
            filepath=filepath,
            align=align,
            color=color,
            position=position,
            paragraph_index=paragraph_index,
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding heading: {str(e)}"


@mcp.tool()
def word_add_paragraph(
    text: str,
    filepath: Optional[str] = None,
    style: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
    font_color: Optional[str] = None,
    highlight_color: Optional[str] = None,
    align: Optional[str] = None,
    line_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    position: Optional[str] = None,
    paragraph_index: Optional[int] = None,
    rich_formatting: bool = False,
) -> str:
    """
    Add a rich paragraph with font formatting, optional inline markdown, links, color, highlight, and positioning.

    Args:
        text: Paragraph text content. When rich_formatting=True, supports **bold**, *italic*, `code`, and [label](https://url).
        filepath: Optional file path. If omitted, adds to active document in live Word.
        style: Optional style name (e.g. 'Normal', 'Quote', 'Intense Quote').
        bold: Set bold font (True / False).
        italic: Set italic font (True / False).
        underline: Set underline font (True / False).
        font_size: Font size in points (e.g. 11, 12, 14).
        font_name: Font family name (e.g. 'Segoe UI', 'Calibri', 'Arial', 'Times New Roman').
        font_color: Hex color string (e.g. '#333333', '#C00000', '#1F497D').
        highlight_color: Highlight color name ('YELLOW', 'GREEN', 'CYAN', 'PINK', 'BLUE', 'RED', 'GRAY_25').
        align: Paragraph alignment ('left', 'center', 'right', 'justify').
        line_spacing: Line spacing multiplier (e.g. 1.0, 1.15, 1.5).
        space_before: Space before paragraph in points.
        space_after: Space after paragraph in points.
        position: Optional insertion position: 'end' (default) or 'start'.
        paragraph_index: Optional 1-based paragraph index to insert before.
        rich_formatting: Parse inline **bold**, *italic*, `code`, and [label](url) inside text (default: False).
    """
    try:
        res = _run(WordEngine.add_paragraph,
            text=text,
            filepath=filepath,
            style=style,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            highlight_color=highlight_color,
            align=align,
            line_spacing=line_spacing,
            space_before=space_before,
            space_after=space_after,
            position=position,
            paragraph_index=paragraph_index,
            rich_formatting=rich_formatting,
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding paragraph: {str(e)}"


@mcp.tool()
def word_add_bullet_list(
    items: List[str],
    filepath: Optional[str] = None,
    list_type: str = "bullet"
) -> str:
    """
    Add a list of items (bulleted or numbered).

    Args:
        items: List of string items to insert.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        list_type: 'bullet' (bulleted list) or 'number' (numbered 1, 2, 3 list).
    """
    try:
        res = _run(WordEngine.add_bullet_list,
            items=items,
            filepath=filepath,
            list_type=list_type
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding list: {str(e)}"


@mcp.tool()
def word_add_callout_box(
    text: str,
    filepath: Optional[str] = None,
    title: Optional[str] = None,
    box_type: str = "info",
    bg_color: Optional[str] = None,
    border_color: Optional[str] = None
) -> str:
    """
    Add a styled callout / alert block with light background and colored accent border.

    Args:
        text: Callout body text.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        title: Optional callout title header (e.g. 'Important Note', 'Key Takeaway').
        box_type: Theme style: 'info' (Blue), 'tip' (Green), 'warning' (Orange), 'important' (Red), 'note' (Gray), 'quote' (Quote).
        bg_color: Optional custom background hex color (e.g. '#EDF5FC').
        border_color: Optional custom border hex color (e.g. '#1F497D').
    """
    try:
        res = _run(WordEngine.add_callout_box,
            text=text,
            filepath=filepath,
            title=title,
            box_type=box_type,
            bg_color=bg_color,
            border_color=border_color
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding callout box: {str(e)}"


@mcp.tool()
def word_add_table(
    headers: Optional[List[str]] = None,
    data: Optional[List[List[Any]]] = None,
    filepath: Optional[str] = None,
    table_style: str = "Table Grid",
    header_bg_color: str = "#1F497D",
    header_font_color: str = "#FFFFFF",
    alternate_row_color: str = "#F2F4F8",
    col_widths: Optional[List[float]] = None,
    alignment: str = "center",
    auto_fit: bool = True,
    merges: Optional[List[List[int]]] = None,
) -> str:
    """
    Add a professionally styled table with custom header fills, zebra row striping, column widths, and optional cell merges.

    Args:
        headers: List of column header titles (e.g. ['Task', 'Owner', 'Due Date', 'Status']).
        data: 2D list of rows and cell values.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        table_style: Base table style name (default: 'Table Grid').
        header_bg_color: Hex color for header row background (default: Navy '#1F497D').
        header_font_color: Hex color for header text (default: White '#FFFFFF').
        alternate_row_color: Hex color for alternating row shading (default: Soft Blue/Gray '#F2F4F8').
        col_widths: List of column widths in inches (e.g. [2.0, 1.5, 1.5, 1.5]).
        alignment: Table alignment ('center', 'left', or 'right').
        auto_fit: Automatically fit columns to contents (default: True).
        merges: Optional list of 0-based cell merge ranges [start_row, start_col, end_row, end_col].
    """
    try:
        res = _run(WordEngine.add_table,
            headers=headers,
            data=data,
            filepath=filepath,
            table_style=table_style,
            header_bg_color=header_bg_color,
            header_font_color=header_font_color,
            alternate_row_color=alternate_row_color,
            col_widths=col_widths,
            alignment=alignment,
            auto_fit=auto_fit,
            merges=merges,
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding table: {str(e)}"


@mcp.tool()
def word_add_image(
    image_path: str,
    filepath: Optional[str] = None,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    caption: Optional[str] = None,
    align: str = "center"
) -> str:
    """
    Insert an image into the Word document with scaling, alignment, and optional figure caption.

    Args:
        image_path: Full file path to the image (PNG, JPG, BMP, etc.).
        filepath: Optional document file path. If omitted, adds to active document in live Word.
        width_inches: Optional width in inches (e.g. 5.5).
        height_inches: Optional height in inches.
        caption: Optional figure caption text displayed below the image.
        align: Image alignment ('center', 'left', 'right').
    """
    try:
        res = _run(WordEngine.add_image,
            image_path=image_path,
            filepath=filepath,
            width_inches=width_inches,
            height_inches=height_inches,
            caption=caption,
            align=align
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding image: {str(e)}"


@mcp.tool()
def word_add_page_break(
    filepath: Optional[str] = None
) -> str:
    """
    Insert a page break into the document.

    Args:
        filepath: Optional file path. If omitted, inserts into active document in live Word.
    """
    try:
        res = _run(WordEngine.add_page_break, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding page break: {str(e)}"


@mcp.tool()
def word_search_and_replace(
    find_text: str,
    replace_text: str,
    filepath: Optional[str] = None,
    match_case: bool = False,
    match_whole_word: bool = False,
    replace_all: bool = True
) -> str:
    """
    Find and replace text across the document (paragraphs, headings, tables).

    Args:
        find_text: The string to search for.
        replace_text: The replacement string.
        filepath: Optional file path. If omitted, replaces in active document in live Word.
        match_case: Case-sensitive search (default: False).
        match_whole_word: Match whole words only (default: False).
        replace_all: Replace all instances (default: True).
    """
    try:
        res = _run(WordEngine.search_and_replace,
            find_text=find_text,
            replace_text=replace_text,
            filepath=filepath,
            match_case=match_case,
            match_whole_word=match_whole_word,
            replace_all=replace_all
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error replacing text: {str(e)}"


# ==========================================
# Advanced Enterprise Tools (1-6)
# ==========================================

@mcp.tool()
def word_set_header_footer(
    header_text: Optional[str] = None,
    footer_text: Optional[str] = None,
    include_page_number: bool = True,
    different_first_page: bool = False,
    header_align: str = "right",
    footer_align: str = "center",
    filepath: Optional[str] = None
) -> str:
    """
    Set custom running header, footer, and dynamic page numbering (Page X).

    Args:
        header_text: Text to display in top header (e.g. 'Confidential Report | Q1 2026').
        footer_text: Text to display in bottom footer (e.g. 'Antigravity AI').
        include_page_number: Whether to append dynamic page numbering (default: True).
        different_first_page: Suppress header/footer on title/first page (default: False).
        header_align: Header alignment ('left', 'center', 'right'). Default: 'right'.
        footer_align: Footer alignment ('left', 'center', 'right'). Default: 'center'.
        filepath: Optional file path. If omitted, applies to active document in live Word.
    """
    try:
        res = _run(WordEngine.set_header_footer,
            header_text=header_text,
            footer_text=footer_text,
            include_page_number=include_page_number,
            different_first_page=different_first_page,
            header_align=header_align,
            footer_align=footer_align,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error setting header/footer: {str(e)}"


@mcp.tool()
def word_insert_toc(
    filepath: Optional[str] = None,
    title: str = "جدول المحتويات",
    upper_heading_level: int = 1,
    lower_heading_level: int = 3,
    position: str = "start"
) -> str:
    """
    Insert an automated Table of Contents (TOC) with heading levels and page numbers.

    Args:
        filepath: Optional file path. If omitted, inserts into active document in live Word.
        title: Title above TOC (default: 'جدول المحتويات' or 'Table of Contents').
        upper_heading_level: Top heading level to include (default: 1).
        lower_heading_level: Deepest heading level to include (default: 3).
        position: Insertion position: 'start' (at beginning of document) or 'end'.
    """
    try:
        res = _run(WordEngine.insert_toc,
            filepath=filepath,
            title=title,
            upper_heading_level=upper_heading_level,
            lower_heading_level=lower_heading_level,
            position=position
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error inserting TOC: {str(e)}"


@mcp.tool()
def word_fill_template(
    template_path: str,
    output_path: str,
    replacements: Dict[str, Any]
) -> str:
    """
    Render a pre-designed Word template (.docx / .dotx) by filling {{tags}} with data.

    Args:
        template_path: Path to template docx file.
        output_path: Destination path for generated document.
        replacements: Key-value map of tags to replace (e.g. {'client_name': 'Acme Corp', 'amount': '$5,000'}).
    """
    try:
        res = _run(WordEngine.fill_template,
            template_path=template_path,
            output_path=output_path,
            replacements=replacements
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error filling template: {str(e)}"


@mcp.tool()
def word_add_comment(
    comment_text: str,
    target_text: Optional[str] = None,
    paragraph_index: Optional[int] = None,
    author: str = "AI Reviewer",
    filepath: Optional[str] = None
) -> str:
    """
    Add a review comment / feedback annotation to a specific text string or paragraph.

    Args:
        comment_text: The review comment note.
        target_text: Specific text phrase in the document to attach the comment to.
        paragraph_index: Optional 1-based paragraph index to attach the comment to.
        author: Reviewer name (default: 'AI Reviewer').
        filepath: Optional file path. If omitted, comments in active document in live Word.
    """
    try:
        res = _run(WordEngine.add_comment,
            comment_text=comment_text,
            target_text=target_text,
            paragraph_index=paragraph_index,
            author=author,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding comment: {str(e)}"


@mcp.tool()
def word_set_page_setup(
    filepath: Optional[str] = None,
    orientation: Optional[str] = None,
    paper_size: Optional[str] = None,
    margin_preset: Optional[str] = None,
    top_margin_inches: Optional[float] = None,
    bottom_margin_inches: Optional[float] = None,
    left_margin_inches: Optional[float] = None,
    right_margin_inches: Optional[float] = None
) -> str:
    """
    Configure page layout: orientation (Portrait/Landscape), paper size (A4/Letter), and margins.

    Args:
        filepath: Optional file path. If omitted, configures active document in live Word.
        orientation: 'portrait' (vertical) or 'landscape' (horizontal).
        paper_size: 'A4', 'Letter', or 'Legal'.
        margin_preset: Margin preset: 'normal' (1 in), 'narrow' (0.5 in), 'moderate', or 'wide'.
        top_margin_inches: Custom top margin in inches.
        bottom_margin_inches: Custom bottom margin in inches.
        left_margin_inches: Custom left margin in inches.
        right_margin_inches: Custom right margin in inches.
    """
    try:
        res = _run(WordEngine.set_page_setup,
            filepath=filepath,
            orientation=orientation,
            paper_size=paper_size,
            margin_preset=margin_preset,
            top_margin_inches=top_margin_inches,
            bottom_margin_inches=bottom_margin_inches,
            left_margin_inches=left_margin_inches,
            right_margin_inches=right_margin_inches
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error configuring page setup: {str(e)}"


@mcp.tool()
def word_add_watermark(
    text: str,
    filepath: Optional[str] = None,
    color: str = "#D3D3D3",
    font_name: str = "Segoe UI",
    font_size: int = 54,
    diagonal: bool = True
) -> str:
    """
    Add a subtle background text watermark (e.g. 'CONFIDENTIAL', 'DRAFT', 'مسودة').

    Args:
        text: Watermark text string.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        color: Hex color string (default: '#D3D3D3' light gray).
        font_name: Font family name (default: 'Segoe UI').
        font_size: Font size in points (default: 54).
        diagonal: Angle diagonally across page (default: True).
    """
    try:
        res = _run(WordEngine.add_watermark,
            text=text,
            filepath=filepath,
            color=color,
            font_name=font_name,
            font_size=font_size,
            diagonal=diagonal
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding watermark: {str(e)}"


# ==========================================
# Export, Save & Script Execution Tools
# ==========================================

@mcp.tool()
def word_export_pdf(
    output_pdf_path: str,
    filepath: Optional[str] = None
) -> str:
    """
    Export the Word document to a high-fidelity PDF file using Microsoft Word's native engine.

    Args:
        output_pdf_path: Target path for the generated PDF (e.g. 'E:/reports/report.pdf').
        filepath: Optional source document path. If omitted, exports active document from live Word.
    """
    try:
        res = _run(WordEngine.export_pdf, output_pdf_path=output_pdf_path, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error exporting PDF: {str(e)}"


@mcp.tool()
def word_modify_paragraph(
    paragraph_index: Optional[int] = None,
    target_text: Optional[str] = None,
    new_text: Optional[str] = None,
    delete: bool = False,
    filepath: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
    font_color: Optional[str] = None,
    highlight_color: Optional[str] = None,
    align: Optional[str] = None,
    rtl: Optional[bool] = None,
    line_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
) -> str:
    """
    Modify formatting or text of an existing paragraph (located by 1-based paragraph_index or target_text), or delete it.

    Args:
        paragraph_index: 1-based paragraph index to modify.
        target_text: Substring to locate the first matching paragraph when paragraph_index is omitted.
        new_text: Optional replacement text for the paragraph.
        delete: If True, deletes the matched paragraph from the document.
        filepath: Optional file path. If omitted, modifies the active document in live Word.
        bold: Optional bold toggle.
        italic: Optional italic toggle.
        underline: Optional underline toggle.
        font_size: Optional font size in points.
        font_name: Optional font family name.
        font_color: Optional hex font color (e.g. '#1F497D').
        highlight_color: Optional highlight color name.
        align: Optional alignment ('left', 'center', 'right', 'justify').
        rtl: Optional RTL paragraph direction toggle (True for RTL, False for LTR).
        line_spacing: Optional line spacing multiplier.
        space_before: Optional space before in points.
        space_after: Optional space after in points.
    """
    try:
        res = _run(
            WordEngine.modify_paragraph,
            paragraph_index=paragraph_index,
            target_text=target_text,
            new_text=new_text,
            delete=delete,
            filepath=filepath,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            highlight_color=highlight_color,
            align=align,
            rtl=rtl,
            line_spacing=line_spacing,
            space_before=space_before,
            space_after=space_after,
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error modifying paragraph: {str(e)}"


@mcp.tool()
def word_clear_document(
    filepath: Optional[str] = None,
    include_headers_footers: bool = False,
) -> str:
    """
    Clear all content from the active document or specified file, leaving a clean blank page.

    Args:
        filepath: Optional file path. If omitted, clears the active document in live Word.
        include_headers_footers: Also clear headers, footers, and watermarks (default: False).
    """
    try:
        res = _run(
            WordEngine.clear_document,
            filepath=filepath,
            include_headers_footers=include_headers_footers,
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error clearing document: {str(e)}"


@mcp.tool()
def word_save_document(
    save_as_path: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Save the currently active document or save it to a new file path.

    Args:
        save_as_path: Optional new destination path to save as.
        filepath: Optional specific document file to save.
    """
    try:
        res = _run(WordEngine.save_document, save_as_path=save_as_path, filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error saving document: {str(e)}"


def word_execute_script(
    code: str,
    filepath: Optional[str] = None
) -> str:
    """
    Execute arbitrary Python automation code with direct access to Word COM and python-docx objects.

    Injected variables available in code:
      - `app`: Word.Application COM instance (win32com.client)
      - `doc`: Active Document COM instance
      - `docx_doc`: python-docx Document instance (if filepath exists)
      - `docx`: docx module
      - `win32com`: win32com.client module
      - `engine`: WordEngine class
      - `Inches`, `Pt`, `RGBColor`, `WD_ALIGN_PARAGRAPH`, `WD_COLOR_INDEX`

    Args:
        code: Python script string to execute.
        filepath: Optional target document file path.
    """
    try:
        res = _run(WordEngine.execute_script, code=code, filepath=filepath)
        if res.get("success"):
            return f"✅ Execution Successful:\n{res.get('output')}"
        else:
            return f"❌ Execution Failed:\nError: {res.get('error')}\nOutput:\n{res.get('output')}"
    except Exception as e:
        return f"❌ Error executing script: {str(e)}"


# Registered only when WORD_MCP_ENABLE_SCRIPT=1: it executes arbitrary Python on this machine.
if script_execution_enabled():
    word_execute_script = mcp.tool()(word_execute_script)


def main():
    """Main entrypoint for CLI and script execution."""
    # stdout is the MCP protocol channel: log to stderr only.
    logging.basicConfig(
        stream=sys.stderr,
        level=os.environ.get("WORD_MCP_LOG", "WARNING").upper(),
        format="%(asctime)s %(name)s %(levelname)s %(message)s",
    )
    log.info("word-mcp starting (mode=%s)", get_mode())
    try:
        mcp.run()
    finally:
        WORKER.shutdown()


if __name__ == "__main__":
    main()
