"""
Word Automation Engine (Enterprise Edition)
-------------------------------------------
Provides a high-level abstraction for interacting with Microsoft Word.
Supports:
1. Live Windows COM Automation (via pywin32) for interactive real-time control of Microsoft Word desktop application.
2. Headless python-docx engine for direct, high-performance file manipulation without requiring Word to be running.
"""

import os
import io
import sys
import json
import re
import traceback
from typing import List, Dict, Any, Optional, Union, Tuple

# Headless docx engine
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# Windows COM Automation
try:
    import win32com.client
    import pythoncom
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


# ==========================================
# Color & Style Helpers
# ==========================================

def parse_hex_color(hex_str: str) -> Tuple[int, int, int]:
    """Parse '#RRGGBB' or 'RRGGBB' or 3-digit shorthand into (R, G, B) tuple."""
    hex_str = hex_str.strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color string: '{hex_str}'. Expected 6-digit hex like '#1F497D'.")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return r, g, b


def cleanup_stale_lock_files(target_path: Optional[str] = None):
    """Remove orphaned Microsoft Word ~$*.docx hidden lock files to prevent 'File In Use' dialogs."""
    try:
        if target_path:
            dir_path = os.path.dirname(os.path.abspath(target_path)) if (os.path.isfile(target_path) or '.' in os.path.basename(target_path)) else os.path.abspath(target_path)
        else:
            dir_path = os.getcwd()
        if os.path.exists(dir_path):
            for fname in os.listdir(dir_path):
                if fname.startswith("~$") and fname.endswith(".docx"):
                    lock_file = os.path.join(dir_path, fname)
                    try:
                        if sys.platform == "win32":
                            try:
                                import ctypes
                                ctypes.windll.kernel32.SetFileAttributesW(lock_file, 128)  # FILE_ATTRIBUTE_NORMAL
                            except Exception:
                                pass
                        os.remove(lock_file)
                    except Exception:
                        pass
    except Exception:
        pass


def hex_to_bgr_int(hex_str: str) -> int:
    """Convert hex color string to Word COM BGR integer (wdColor)."""
    r, g, b = parse_hex_color(hex_str)
    return r + (g * 256) + (b * 65536)


def hex_to_rgb_color(hex_str: str) -> RGBColor:
    """Convert hex color string to python-docx RGBColor."""
    r, g, b = parse_hex_color(hex_str)
    return RGBColor(r, g, b)


def clean_hex(hex_str: str) -> str:
    """Return 6-character hex string without leading '#' for XML fill values."""
    hex_str = hex_str.strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    return hex_str.upper()


def is_arabic_text(text: str) -> bool:
    """Detect if string contains Arabic/Persian/Hebrew characters for RTL formatting."""
    return bool(re.search(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\u0590-\u05FF]', text))


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


def set_cell_background(cell, hex_color: str):
    """Set background shading color of a docx table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{clean_hex(hex_color)}"/>')
    tcPr.append(shd)


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
    tcPr.append(tcMar)


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
    tcPr.append(tcBorders)


# ==========================================
# Live Word COM Management
# ==========================================

class LiveWordSession:
    """Manages COM connection to Microsoft Word desktop application safely."""

    @staticmethod
    def is_word_running() -> bool:
        """Check if Microsoft Word is running without launching a new instance."""
        if not COM_AVAILABLE:
            return False
        try:
            pythoncom.CoInitialize()
            app = win32com.client.GetActiveObject("Word.Application")
            return app is not None
        except Exception:
            return False

    @staticmethod
    def get_app(create_if_missing: bool = True, visible: bool = True):
        """Get running Word COM instance or launch one if create_if_missing=True."""
        if not COM_AVAILABLE:
            raise RuntimeError("pywin32 (win32com) is not installed or unavailable.")
        pythoncom.CoInitialize()
        try:
            # First try to get already running instance
            app = win32com.client.GetActiveObject("Word.Application")
            if visible and hasattr(app, "Visible") and not app.Visible:
                app.Visible = True
            if hasattr(app, "DisplayAlerts"):
                app.DisplayAlerts = 0
            if visible:
                try:
                    app.Activate()
                except Exception:
                    pass
            return app
        except Exception:
            if create_if_missing:
                app = win32com.client.Dispatch("Word.Application")
                if hasattr(app, "Visible"):
                    app.Visible = visible
                if hasattr(app, "DisplayAlerts"):
                    app.DisplayAlerts = 0
                return app
            raise RuntimeError("Microsoft Word is not currently running.")

    @staticmethod
    def get_document(app, filepath: Optional[str] = None, create_if_empty: bool = True):
        """Get document from Word COM application by filepath or active document."""
        if not app:
            return None
        if filepath:
            abs_path = os.path.abspath(filepath)
            try:
                for doc in app.Documents:
                    try:
                        doc_path = getattr(doc, "FullName", "")
                        if doc_path and os.path.abspath(doc_path).lower() == abs_path.lower():
                            doc.Activate()
                            return doc
                    except Exception:
                        pass
            except Exception:
                pass

            if os.path.exists(abs_path):
                doc = app.Documents.Open(abs_path, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)
                doc.Activate()
                return doc
            elif create_if_empty:
                doc = app.Documents.Add()
                doc.Activate()
                return doc
            else:
                raise FileNotFoundError(f"Document file not found: {abs_path}")
        else:
            try:
                if app.Documents.Count > 0:
                    doc = app.ActiveDocument
                    doc.Activate()
                    return doc
                elif create_if_empty:
                    doc = app.Documents.Add()
                    doc.Activate()
                    return doc
            except Exception:
                pass
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
        """Check live status of Microsoft Word."""
        status = {
            "word_running": False,
            "com_available": COM_AVAILABLE,
            "version": None,
            "open_documents": [],
            "active_document": None,
            "active_document_path": None
        }
        if not COM_AVAILABLE:
            return status

        try:
            pythoncom.CoInitialize()
            app = win32com.client.GetActiveObject("Word.Application")
            status["word_running"] = True
            status["version"] = str(getattr(app, "Version", ""))
            
            docs = []
            try:
                for doc in app.Documents:
                    try:
                        docs.append({
                            "name": getattr(doc, "Name", ""),
                            "path": getattr(doc, "FullName", ""),
                            "saved": getattr(doc, "Saved", True)
                        })
                    except Exception:
                        pass
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
            cleanup_stale_lock_files(abs_path)
        else:
            abs_path = None
            cleanup_stale_lock_files()

        if COM_AVAILABLE and visible:
            try:
                app = LiveWordSession.get_app(create_if_missing=True, visible=True)
                doc = app.Documents.Add()
                doc.Activate()
                sel = app.Selection
                sel.HomeKey(Unit=6)  # wdStory

                if title:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 24
                    sel.Font.Bold = True
                    sel.Font.Color = hex_to_bgr_int("#1F497D")
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_arabic_text(title) else "left")
                    sel.ParagraphFormat.SpaceAfter = 6
                    sel.TypeText(title + "\n")

                if subtitle:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 13
                    sel.Font.Italic = True
                    sel.Font.Bold = False
                    sel.Font.Color = hex_to_bgr_int("#595959")
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_arabic_text(subtitle) else "left")
                    sel.ParagraphFormat.SpaceAfter = 14
                    sel.TypeText(subtitle + "\n")

                if author:
                    sel.Font.Name = "Segoe UI"
                    sel.Font.Size = 10
                    sel.Font.Italic = False
                    sel.Font.Color = hex_to_bgr_int("#7F7F7F")
                    sel.ParagraphFormat.Alignment = get_com_alignment("right" if is_arabic_text(author) else "left")
                    sel.ParagraphFormat.SpaceAfter = 18
                    author_label = "المؤلف: " if is_arabic_text(author) else "Author: "
                    sel.TypeText(f"{author_label}{author}\n")

                sel.Font.Reset()

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
            except Exception:
                pass

        # Headless docx creation
        doc = docx.Document()
        styles = [s.name for s in doc.styles]
        
        if title:
            title_style = "Title" if "Title" in styles else None
            p_title = doc.add_paragraph(style=title_style)
            p_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic_text(title) else WD_ALIGN_PARAGRAPH.LEFT
            p_title.paragraph_format.space_after = Pt(6)
            p_title.paragraph_format.space_before = Pt(0)
            run_title = p_title.add_run(title)
            run_title.font.name = "Segoe UI"
            run_title.font.size = Pt(24)
            run_title.font.bold = True
            run_title.font.color.rgb = hex_to_rgb_color("#1F497D")

        if subtitle:
            sub_style = "Subtitle" if "Subtitle" in styles else None
            p_sub = doc.add_paragraph(style=sub_style)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic_text(subtitle) else WD_ALIGN_PARAGRAPH.LEFT
            p_sub.paragraph_format.space_after = Pt(12)
            p_sub.paragraph_format.space_before = Pt(0)
            run_sub = p_sub.add_run(subtitle)
            run_sub.font.name = "Segoe UI"
            run_sub.font.size = Pt(13)
            run_sub.font.italic = True
            run_sub.font.color.rgb = hex_to_rgb_color("#595959")

        if author:
            p_auth = doc.add_paragraph()
            p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic_text(author) else WD_ALIGN_PARAGRAPH.LEFT
            p_auth.paragraph_format.space_after = Pt(18)
            p_auth.paragraph_format.space_before = Pt(0)
            author_label = "المؤلف: " if is_arabic_text(author) else "Author: "
            run_auth = p_auth.add_run(f"{author_label}{author}")
            run_auth.font.name = "Segoe UI"
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

        cleanup_stale_lock_files(abs_path)

        if visible and sys.platform == "win32" and not COM_AVAILABLE:
            try:
                os.startfile(abs_path)
            except Exception:
                pass

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=True, visible=visible)
                doc = app.Documents.Open(abs_path, ReadOnly=read_only)
                if visible:
                    doc.Activate()
                return {
                    "mode": "live_com",
                    "status": "opened_live_on_screen",
                    "doc_name": doc.Name,
                    "filepath": abs_path,
                    "read_only": read_only
                }
            except Exception:
                pass

        doc = docx.Document(abs_path)
        return {
            "mode": "headless_docx",
            "status": "verified_and_loaded",
            "filepath": abs_path,
            "paragraphs_count": len(doc.paragraphs),
            "tables_count": len(doc.tables)
        }

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
        if not filepath and COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                if app.Documents.Count > 0:
                    doc = app.ActiveDocument
                    text_content = doc.Content.Text
                    if output_format == "text":
                        return {
                            "document_name": doc.Name,
                            "filepath": getattr(doc, "FullName", ""),
                            "mode": "live_com",
                            "text": text_content.replace("\r", "\n")
                        }
                    else:
                        temp_path = os.path.join(os.environ.get("TEMP", "."), f"_temp_mcp_read_{os.getpid()}.docx")
                        doc.SaveCopyAs(temp_path)
                        filepath = temp_path
            except Exception:
                pass

        if not filepath:
            raise ValueError("No filepath provided and no active document found in Microsoft Word.")

        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        doc = docx.Document(abs_path)
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

            if style_name.startswith("Heading 1") or is_title_heading:
                md_lines.append(f"# {p_text}\n")
            elif style_name.startswith("Heading 2"):
                md_lines.append(f"## {p_text}\n")
            elif style_name.startswith("Heading 3"):
                md_lines.append(f"### {p_text}\n")
            elif style_name.startswith("Heading 4"):
                md_lines.append(f"#### {p_text}\n")
            elif style_name.startswith("Subtitle"):
                md_lines.append(f"*{p_text}*\n")
            elif style_name.startswith("List Bullet") or style_name == "List Bullet":
                md_lines.append(f"- {p_text}")
            elif style_name.startswith("List Number") or style_name == "List Number":
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
                    headers = tbl_rows[0]
                    md_lines.append("| " + " | ".join(headers) + " |")
                    md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                    for row_cells in tbl_rows[1:]:
                        md_lines.append("| " + " | ".join(row_cells) + " |")
                    md_lines.append("\n")

        total_words = sum(len(p["text"].split()) for p in paragraphs_data)

        res = {
            "filepath": abs_path,
            "total_paragraphs": len(paragraphs_data),
            "total_tables": len(tables_data),
            "total_words": total_words
        }

        if output_format == "markdown":
            res["markdown"] = "\n".join(md_lines)
        elif output_format == "text":
            res["text"] = "\n".join(plain_lines)
        else:
            res["paragraphs"] = paragraphs_data
            res["tables"] = tables_data

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
        align: Optional[str] = "left",
        color: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a heading (Level 1 to 6) to the document."""
        level = max(1, min(6, int(level)))
        
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    try:
                        sel.Style = f"Heading {level}"
                    except Exception:
                        pass
                    sel.ParagraphFormat.Alignment = get_com_alignment(align)
                    if color:
                        sel.Font.Color = hex_to_bgr_int(color)
                    sel.TypeText(text)
                    sel.Font.Reset()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "heading_added_live",
                        "doc_name": doc.Name,
                        "level": level,
                        "text": text
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        heading = doc.add_heading(text, level=level)
        if align:
            heading.alignment = get_docx_alignment(align)
        elif is_arabic_text(text):
            heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT
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
        space_after: Optional[float] = None
    ) -> Dict[str, Any]:
        """Add a rich paragraph with font formatting, color, highlight, and spacing."""
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    if style:
                        try:
                            sel.Style = style
                        except Exception:
                            pass
                    sel.Font.Bold = bold
                    sel.Font.Italic = italic
                    sel.Font.Underline = 1 if underline else 0
                    if font_name:
                        sel.Font.Name = font_name
                    if font_size:
                        sel.Font.Size = font_size
                    if font_color:
                        sel.Font.Color = hex_to_bgr_int(font_color)
                    if align:
                        sel.ParagraphFormat.Alignment = get_com_alignment(align)
                    elif is_arabic_text(text):
                        sel.ParagraphFormat.Alignment = 2  # Right
                    if space_after is not None:
                        sel.ParagraphFormat.SpaceAfter = space_after
                    if space_before is not None:
                        sel.ParagraphFormat.SpaceBefore = space_before
                    if line_spacing is not None:
                        sel.ParagraphFormat.LineSpacing = line_spacing * 12
                    
                    sel.TypeText(text)
                    sel.Font.Reset()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "paragraph_added_live",
                        "doc_name": doc.Name,
                        "text_length": len(text)
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        p = doc.add_paragraph(style=style if style and style in [s.name for s in doc.styles] else None)
        if align:
            p.alignment = get_docx_alignment(align)
        elif is_arabic_text(text):
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if space_before is not None:
            p.paragraph_format.space_before = Pt(space_before)
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
        if line_spacing is not None:
            p.paragraph_format.line_spacing = line_spacing

        run = p.add_run(text)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if underline:
            run.underline = True
        if font_size:
            run.font.size = Pt(font_size)
        if font_name:
            run.font.name = font_name
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
        
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    for item in items:
                        sel.TypeParagraph()
                        try:
                            sel.Style = style_name
                        except Exception:
                            pass
                        sel.TypeText(item)
                    sel.Font.Reset()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "list_added_live",
                        "doc_name": doc.Name,
                        "items_count": len(items),
                        "list_type": list_type
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        for item in items:
            p = doc.add_paragraph(item, style=style_name if style_name in [s.name for s in doc.styles] else None)
            if is_arabic_text(item):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

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
        """Add a styled callout / alert box with custom background and left border accent."""
        box_themes = {
            "info": {"bg": "#EDF5FC", "border": "#1F497D", "icon": "ℹ️ ", "title": "معلومات" if is_arabic_text(text) else "Information"},
            "tip": {"bg": "#EDF7ED", "border": "#2E7D32", "icon": "💡 ", "title": "نصيحة" if is_arabic_text(text) else "Tip"},
            "warning": {"bg": "#FFF4E5", "border": "#ED6C02", "icon": "⚠️ ", "title": "تنبيه" if is_arabic_text(text) else "Warning"},
            "note": {"bg": "#F5F5F5", "border": "#616161", "icon": "📝 ", "title": "ملاحظة" if is_arabic_text(text) else "Note"},
            "quote": {"bg": "#FAFAFA", "border": "#757575", "icon": "❝ ", "title": "اقتباس" if is_arabic_text(text) else "Quote"},
            "important": {"bg": "#FDEDED", "border": "#D32F2F", "icon": "🚨 ", "title": "هام جداً" if is_arabic_text(text) else "Important"}
        }
        theme = box_themes.get(box_type.lower(), box_themes["info"])
        final_bg = bg_color or theme["bg"]
        final_border = border_color or theme["border"]
        final_title = title if title is not None else f"{theme['icon']}{theme['title']}"

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    tbl = doc.Tables.Add(Range=sel.Range, NumRows=1, NumColumns=1)
                    cell = tbl.Cell(1, 1)
                    cell.Shading.BackgroundPatternColor = hex_to_bgr_int(final_bg)
                    
                    cell.Borders(-1).LineStyle = 0  # wdBorderTop = -1
                    cell.Borders(-3).LineStyle = 0  # wdBorderBottom = -3
                    cell.Borders(-4).LineStyle = 0  # wdBorderRight = -4
                    cell.Borders(-2).LineStyle = 1  # wdBorderLeft = -2
                    cell.Borders(-2).LineWidth = 24 # 3.0 pt
                    cell.Borders(-2).Color = hex_to_bgr_int(final_border)
                    
                    cell_range = cell.Range
                    cell_range.Text = (final_title + "\n" if final_title else "") + text
                    if final_title:
                        first_line = cell.Range.Paragraphs(1).Range
                        first_line.Font.Bold = True
                        first_line.Font.Color = hex_to_bgr_int(final_border)
                    
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "callout_box_added_live",
                        "doc_name": doc.Name,
                        "box_type": box_type,
                        "title": final_title
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Clear outer table borders
        tblPr = table._tbl.tblPr
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

        set_cell_background(cell, final_bg)
        set_cell_margins(cell, top_dpt=140, bottom_dpt=140, left_dpt=200, right_dpt=160)
        
        # Left border for LTR, Right border if primary language is Arabic
        border_side = "right" if is_arabic_text(text) else "left"
        border_spec = {"val": "single", "sz": "24", "color": final_border}
        if border_side == "left":
            set_cell_borders(cell, top=None, bottom=None, left=border_spec, right=None)
        else:
            set_cell_borders(cell, top=None, bottom=None, left=None, right=border_spec)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4) if text else Pt(0)
        if is_arabic_text(text) or (final_title and is_arabic_text(final_title)):
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        if final_title:
            run_title = p.add_run(final_title + "\n")
            run_title.font.name = "Segoe UI"
            run_title.font.size = Pt(11)
            run_title.font.bold = True
            run_title.font.color.rgb = hex_to_rgb_color(final_border)

        if text:
            run_text = p.add_run(text)
            run_text.font.name = "Segoe UI"
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
        auto_fit: bool = True
    ) -> Dict[str, Any]:
        """Add a styled table with custom headers, zebra rows, and widths."""
        data = data or []
        num_cols = len(headers) if headers else (len(data[0]) if data else 1)
        num_rows = (1 if headers else 0) + len(data)

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    tbl = doc.Tables.Add(Range=sel.Range, NumRows=num_rows, NumColumns=num_cols)
                    
                    row_offset = 1
                    if headers:
                        for c_i, h_text in enumerate(headers, 1):
                            cell = tbl.Cell(1, c_i)
                            cell.Range.Text = str(h_text)
                            cell.Range.Font.Bold = True
                            cell.Range.Font.Name = "Segoe UI"
                            cell.Range.Font.Size = 10.5
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
                            cell.Range.Text = str(val) if val is not None else ""
                            cell.Range.Font.Name = "Segoe UI"
                            cell.Range.Font.Size = 9.5
                            if is_even and alternate_row_color:
                                cell.Shading.BackgroundPatternColor = hex_to_bgr_int(alternate_row_color)

                    if col_widths:
                        for c_i, w in enumerate(col_widths, 1):
                            if c_i <= num_cols:
                                tbl.Columns(c_i).Width = w * 72

                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "table_added_live",
                        "doc_name": doc.Name,
                        "rows": num_rows,
                        "cols": num_cols
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        table = doc.add_table(rows=num_rows, cols=num_cols)
        if table_style and table_style in [s.name for s in doc.styles]:
            table.style = table_style
        table.alignment = WD_TABLE_ALIGNMENT.CENTER if alignment.lower() == "center" else WD_TABLE_ALIGNMENT.LEFT
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
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                run = p.add_run(str(text))
                run.font.name = "Segoe UI"
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
                    float(val_str.replace("$", "").replace(",", "").replace("%", ""))
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                except ValueError:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic_text(val_str) else WD_ALIGN_PARAGRAPH.LEFT

                run = p.add_run(val_str)
                run.font.name = "Segoe UI"
                run.font.size = Pt(9.5)
                run.font.color.rgb = hex_to_rgb_color("#262626")

        if col_widths:
            for c_i, w in enumerate(col_widths):
                if c_i < len(table.columns):
                    table.columns[c_i].width = Inches(w)
            for row in table.rows:
                for c_i, w in enumerate(col_widths):
                    if c_i < len(row.cells):
                        row.cells[c_i].width = Inches(w)

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

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    shape = doc.InlineShapes.AddPicture(FileName=abs_img_path, Range=sel.Range)
                    if width_inches:
                        shape.Width = width_inches * 72
                    if height_inches:
                        shape.Height = height_inches * 72
                    sel.ParagraphFormat.Alignment = get_com_alignment(align)
                    if caption:
                        sel.EndKey(Unit=6)
                        sel.TypeParagraph()
                        sel.Font.Italic = True
                        sel.Font.Size = 9
                        sel.Font.Color = hex_to_bgr_int("#595959")
                        sel.TypeText(f"Figure: {caption}")
                        sel.Font.Reset()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "image_added_live",
                        "doc_name": doc.Name,
                        "image": abs_img_path
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        p = doc.add_paragraph()
        p.alignment = get_docx_alignment(align)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)

        w = Inches(width_inches) if width_inches else None
        h = Inches(height_inches) if height_inches else None

        if w and h:
            doc.add_picture(abs_img_path, width=w, height=h)
        elif w:
            doc.add_picture(abs_img_path, width=w)
        elif h:
            doc.add_picture(abs_img_path, height=h)
        else:
            doc.add_picture(abs_img_path, width=Inches(5.5))

        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = get_docx_alignment(align)
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(10)
            prefix = "الشكل: " if is_arabic_text(caption) else "Figure: "
            run_cap = p_cap.add_run(f"{prefix}{caption}")
            run_cap.font.name = "Segoe UI"
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
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    sel.EndKey(Unit=6)
                    sel.InsertBreak(Type=7)
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {"mode": "live_com", "status": "page_break_added_live", "doc_name": doc.Name}
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()
        doc.add_page_break()
        doc.save(abs_path)
        return {"mode": "headless_docx", "status": "page_break_added", "filepath": abs_path}

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
        """Search and replace text across document."""
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    find_obj = doc.Content.Find
                    find_obj.ClearFormatting()
                    find_obj.Replacement.ClearFormatting()
                    find_obj.Text = find_text
                    find_obj.Replacement.Text = replace_text
                    find_obj.Forward = True
                    find_obj.Wrap = 1  # wdFindContinue
                    find_obj.MatchCase = match_case
                    find_obj.MatchWholeWord = match_whole_word
                    
                    replace_const = 2 if replace_all else 1
                    success = find_obj.Execute(Replace=replace_const)
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "replaced_live",
                        "doc_name": doc.Name,
                        "find_text": find_text,
                        "replace_text": replace_text,
                        "success": bool(success)
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
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
                if not p.text or not regex.search(p.text):
                    continue
                # Try single run replacements first
                replaced_run = False
                for r in p.runs:
                    if r.text and regex.search(r.text):
                        count = 1 if (not replace_all and replace_count >= 1) else 0
                        new_text, n = regex.subn(replace_text, r.text, count=count)
                        if n > 0:
                            r.text = new_text
                            replace_count += n
                            replaced_run = True
                            if not replace_all:
                                return
                # If spanned across runs and still present
                if regex.search(p.text) and not replaced_run:
                    count = 1 if (not replace_all and replace_count >= 1) else 0
                    new_full, n = regex.subn(replace_text, p.text, count=count)
                    if n > 0:
                        p.text = new_full
                        replace_count += n
                        if not replace_all:
                            return

        replace_in_paragraphs(doc.paragraphs)
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)
        for sec in doc.sections:
            replace_in_paragraphs(sec.header.paragraphs)
            replace_in_paragraphs(sec.footer.paragraphs)

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
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    for s_idx in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(s_idx)
                        sec.PageSetup.DifferentFirstPageHeaderFooter = different_first_page
                        
                        if header_text is not None:
                            hdr = sec.Headers(1)
                            hdr.Range.Text = header_text
                            hdr.Range.ParagraphFormat.Alignment = get_com_alignment(header_align)
                            hdr.Range.Font.Name = "Segoe UI"
                            hdr.Range.Font.Size = 9
                            hdr.Range.Font.Color = hex_to_bgr_int("#7F7F7F")

                        if footer_text is not None or include_page_number:
                            ftr = sec.Footers(1)
                            ftr.Range.Text = footer_text or ""
                            ftr.Range.ParagraphFormat.Alignment = get_com_alignment(footer_align)
                            ftr.Range.Font.Name = "Segoe UI"
                            ftr.Range.Font.Size = 9
                            ftr.Range.Font.Color = hex_to_bgr_int("#7F7F7F")
                            if include_page_number:
                                ftr.PageNumbers.Add(PageNumberAlignment=get_com_alignment(footer_align), FirstPage=not different_first_page)

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
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        for section in doc.sections:
            section.different_first_page_header_footer = different_first_page
            
            if header_text is not None:
                hdr = section.header
                p_h = hdr.paragraphs[0]
                p_h.alignment = get_docx_alignment(header_align)
                p_h.text = ""
                r_h = p_h.add_run(header_text)
                r_h.font.name = "Segoe UI"
                r_h.font.size = Pt(9)
                r_h.font.color.rgb = hex_to_rgb_color("#7F7F7F")

            if footer_text is not None or include_page_number:
                ftr = section.footer
                p_f = ftr.paragraphs[0]
                p_f.alignment = get_docx_alignment(footer_align)
                p_f.text = ""
                if footer_text:
                    r_f = p_f.add_run(footer_text + (" — " if include_page_number else ""))
                    r_f.font.name = "Segoe UI"
                    r_f.font.size = Pt(9)
                    r_f.font.color.rgb = hex_to_rgb_color("#7F7F7F")
                
                if include_page_number:
                    fld = parse_xml(
                        f'<w:fldSimple {nsdecls("w")} w:instr="PAGE">'
                        f'<w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/><w:sz w:val="18"/><w:color w:val="7F7F7F"/></w:rPr><w:t>1</w:t></w:r>'
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
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sel = app.Selection
                    if position == "start":
                        sel.HomeKey(Unit=6)
                    else:
                        sel.EndKey(Unit=6)
                    
                    if title:
                        sel.Style = "Heading 1"
                        sel.TypeText(title + "\n")
                        sel.Style = "Normal"

                    rng = sel.Range
                    toc = doc.TablesOfContents.Add(
                        Range=rng,
                        UseHeadingStyles=True,
                        UpperHeadingLevel=upper_heading_level,
                        LowerHeadingLevel=lower_heading_level,
                        RightAlignPageNumbers=True
                    )
                    toc.Update()
                    sel.EndKey(Unit=6)
                    sel.TypeParagraph()
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "toc_inserted_live",
                        "doc_name": doc.Name,
                        "title": title
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        # Build TOC paragraph
        if position == "start" and len(doc.paragraphs) > 0:
            first_p = doc.paragraphs[0]
            if title:
                p_h = first_p.insert_paragraph_before(title, style="Heading 1" if "Heading 1" in [s.name for s in doc.styles] else None)
                if p_h.runs:
                    p_h.runs[0].font.name = "Segoe UI"
                    p_h.runs[0].font.color.rgb = hex_to_rgb_color("#1F497D")
            p_toc = first_p.insert_paragraph_before()
            fldSimple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\o &quot;{upper_heading_level}-{lower_heading_level}&quot; \\h \\z \\u"/>')
            p_toc._p.append(fldSimple)
            first_p.insert_paragraph_before()  # Spacing
        else:
            if title:
                h = doc.add_heading(title, level=1)
                if h.runs:
                    h.runs[0].font.name = "Segoe UI"
                    h.runs[0].font.color.rgb = hex_to_rgb_color("#1F497D")
            p = doc.add_paragraph()
            fldSimple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\o &quot;{upper_heading_level}-{lower_heading_level}&quot; \\h \\z \\u"/>')
            p._p.append(fldSimple)
            doc.add_paragraph()

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

        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        doc = docx.Document(abs_tpl)

        def replace_in_paras(paras):
            count = 0
            for p in paras:
                if not p.text:
                    continue
                for k, v in replacements.items():
                    val_str = str(v)
                    tags = [f"{{{{{k}}}}}", f"{{{k}}}"]
                    for tag in tags:
                        if tag not in p.text:
                            continue
                        replaced_run = False
                        for r in p.runs:
                            if tag in r.text:
                                r.text = r.text.replace(tag, val_str)
                                count += 1
                                replaced_run = True
                        if tag in p.text and not replaced_run:
                            p.text = p.text.replace(tag, val_str)
                            count += 1
            return count

        total_replacements = replace_in_paras(doc.paragraphs)
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    total_replacements += replace_in_paras(cell.paragraphs)
        for sec in doc.sections:
            total_replacements += replace_in_paras(sec.header.paragraphs)
            total_replacements += replace_in_paras(sec.footer.paragraphs)

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
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    target_range = None
                    if target_text:
                        find_obj = doc.Content.Find
                        find_obj.ClearFormatting()
                        find_obj.Text = target_text
                        if find_obj.Execute():
                            target_range = find_obj.Parent
                    
                    if target_range is None and paragraph_index and paragraph_index <= doc.Paragraphs.Count:
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
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()
        
        # Locate insertion target
        target_para = None
        if target_text:
            for p in doc.paragraphs:
                if target_text in p.text:
                    target_para = p
                    break
        elif paragraph_index and 1 <= paragraph_index <= len(doc.paragraphs):
            target_para = doc.paragraphs[paragraph_index - 1]

        # Add review note block
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
        Configure page layout, orientation (Portrait/Landscape), paper size (A4/Letter), and margins.
        """
        margin_presets = {
            "normal": (1.0, 1.0, 1.0, 1.0),
            "narrow": (0.5, 0.5, 0.5, 0.5),
            "moderate": (1.0, 1.0, 0.75, 0.75),
            "wide": (1.0, 1.0, 2.0, 2.0)
        }
        if margin_preset and margin_preset.lower() in margin_presets:
            p_top, p_bot, p_left, p_right = margin_presets[margin_preset.lower()]
            top_margin_inches = top_margin_inches or p_top
            bottom_margin_inches = bottom_margin_inches or p_bot
            left_margin_inches = left_margin_inches or p_left
            right_margin_inches = right_margin_inches or p_right

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    for s_idx in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections(s_idx)
                        if orientation:
                            sec.PageSetup.Orientation = 1 if orientation.lower() in ("landscape", "horizontal") else 0
                        if paper_size:
                            ps = paper_size.lower()
                            if ps == "a4":
                                sec.PageSetup.PaperSize = 7  # wdPaperA4
                            elif ps in ("letter", "us_letter"):
                                sec.PageSetup.PaperSize = 0  # wdPaperLetter
                            elif ps == "legal":
                                sec.PageSetup.PaperSize = 4  # wdPaperLegal
                        if top_margin_inches:
                            sec.PageSetup.TopMargin = top_margin_inches * 72
                        if bottom_margin_inches:
                            sec.PageSetup.BottomMargin = bottom_margin_inches * 72
                        if left_margin_inches:
                            sec.PageSetup.LeftMargin = left_margin_inches * 72
                        if right_margin_inches:
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
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        for sec in doc.sections:
            if orientation:
                is_landscape = orientation.lower() in ("landscape", "horizontal")
                sec.orientation = WD_ORIENT.LANDSCAPE if is_landscape else WD_ORIENT.PORTRAIT
                if is_landscape and sec.page_width < sec.page_height:
                    sec.page_width, sec.page_height = sec.page_height, sec.page_width
                elif not is_landscape and sec.page_width > sec.page_height:
                    sec.page_width, sec.page_height = sec.page_height, sec.page_width

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

            if top_margin_inches:
                sec.top_margin = Inches(top_margin_inches)
            if bottom_margin_inches:
                sec.bottom_margin = Inches(bottom_margin_inches)
            if left_margin_inches:
                sec.left_margin = Inches(left_margin_inches)
            if right_margin_inches:
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
        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
                if doc:
                    sec = doc.Sections(1)
                    hdr = sec.Headers(1)
                    shape = hdr.Shapes.AddTextEffect(0, text, font_name, font_size, False, False, 100, 100)
                    shape.Fill.Solid()
                    shape.Fill.ForeColor.RGB = hex_to_bgr_int(color)
                    shape.Line.Visible = False
                    if diagonal:
                        shape.Rotation = -45
                    
                    if filepath and hasattr(doc, "Save"):
                        doc.Save()
                    return {
                        "mode": "live_com",
                        "status": "watermark_added_live",
                        "doc_name": doc.Name,
                        "watermark_text": text
                    }
            except Exception:
                pass

        abs_path = os.path.abspath(filepath) if filepath else os.path.abspath("document.docx")
        doc = docx.Document(abs_path) if os.path.exists(abs_path) else docx.Document()

        sec = doc.sections[0]
        header = sec.header
        p = header.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add high-fidelity VML watermark shape to header in OpenXML
        rot = "-45" if diagonal else "0"
        hex_c = clean_hex(color)
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
            f'<v:textpath string="{text}" style="font-family:\'{font_name}\';font-size:{font_size}pt;font-weight:bold"/>'
            f'</v:shape>'
            f'</w:pict>'
            f'</w:r>'
            f'</w:p>'
        )
        try:
            wm_elem = parse_xml(watermark_xml)
            header._element.append(wm_elem)
        except Exception:
            r = p.add_run(f"[{text.upper()}]")
            r.font.name = font_name
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
            raise RuntimeError("Live Microsoft Word COM is required to export high-fidelity PDF documents.")

        was_running = LiveWordSession.is_word_running()
        app = LiveWordSession.get_app(create_if_missing=True, visible=False)
        doc = None
        doc_opened_here = False

        try:
            if filepath:
                abs_path = os.path.abspath(filepath)
                # Check if already open
                for d in app.Documents:
                    try:
                        if os.path.abspath(d.FullName).lower() == abs_path.lower():
                            doc = d
                            break
                    except Exception:
                        pass
                if not doc:
                    doc = app.Documents.Open(abs_path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
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
            if not was_running and app:
                try:
                    app.Quit()
                except Exception:
                    pass

    @classmethod
    def save_document(
        cls,
        save_as_path: Optional[str] = None,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Save the document or save it to a new location."""
        if COM_AVAILABLE:
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
                        doc.Save()
                        return {"mode": "live_com", "status": "saved", "filepath": getattr(doc, "FullName", doc.Name)}
            except Exception:
                pass

        if filepath:
            abs_path = os.path.abspath(filepath)
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
        """Execute arbitrary Python automation code."""
        app = None
        doc = None
        docx_doc = None

        if COM_AVAILABLE:
            try:
                app = LiveWordSession.get_app(create_if_missing=False)
                doc = LiveWordSession.get_document(app, filepath, create_if_empty=False)
            except Exception:
                pass

        if filepath and os.path.exists(os.path.abspath(filepath)):
            try:
                docx_doc = docx.Document(os.path.abspath(filepath))
            except Exception:
                pass

        old_stdout = sys.stdout
        redirected_output = io.StringIO()
        sys.stdout = redirected_output

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
            exec(code, execution_scope)
            sys.stdout = old_stdout
            return {
                "success": True,
                "output": redirected_output.getvalue() or "Script executed successfully without output."
            }
        except Exception as e:
            sys.stdout = old_stdout
            return {
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc(),
                "output": redirected_output.getvalue()
            }
