"""
Word MCP Server (Enterprise Edition)
------------------------------------
Model Context Protocol (MCP) server providing high-level tools to interact with Microsoft Word.
Supports both live desktop Word automation (via COM/pywin32) and headless document manipulation (via python-docx).
"""

import os
import sys
import json
import logging
from typing import List, Optional, Dict, Any

# Ensure server directory is in sys.path
SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
if SERVER_DIR not in sys.path:
    sys.path.insert(0, SERVER_DIR)

try:
    from mcp.server.fastmcp import FastMCP
except (ImportError, ModuleNotFoundError):
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except ImportError:
        raise ImportError("Please install mcp: pip install mcp")


from com_runtime import WORKER, get_mode, script_execution_enabled
from word_engine import WordEngine, LiveWordSession

log = logging.getLogger("word_mcp")


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
            header = (
                f"**File:** `{res['filepath']}` | **Paragraphs:** {res['total_paragraphs']} | "
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
    align: str = "left",
    color: Optional[str] = None
) -> str:
    """
    Add a heading (Level 1 to 6) to the document.

    Args:
        text: Heading text.
        level: Heading level (1 to 6). Default: 1.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        align: Text alignment ('left', 'center', 'right').
        color: Optional hex color for heading text (e.g. '#1F497D', '#2B2B2B').
    """
    try:
        res = _run(WordEngine.add_heading,
            text=text,
            level=level,
            filepath=filepath,
            align=align,
            color=color
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
    space_after: Optional[float] = None
) -> str:
    """
    Add a rich paragraph with font formatting, color, highlight, and spacing.

    Args:
        text: Paragraph text content.
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
            space_after=space_after
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
    auto_fit: bool = True
) -> str:
    """
    Add a professionally styled table with custom header fills, zebra row striping, and column widths.

    Args:
        headers: List of column header titles (e.g. ['Task', 'Owner', 'Due Date', 'Status']).
        data: 2D list of rows and cell values.
        filepath: Optional file path. If omitted, adds to active document in live Word.
        table_style: Base table style name (default: 'Table Grid').
        header_bg_color: Hex color for header row background (default: Navy '#1F497D').
        header_font_color: Hex color for header text (default: White '#FFFFFF').
        alternate_row_color: Hex color for alternating row shading (default: Soft Blue/Gray '#F2F4F8').
        col_widths: List of column widths in inches (e.g. [2.0, 1.5, 1.5, 1.5]).
        alignment: Table alignment ('center' or 'left').
        auto_fit: Automatically fit columns to contents (default: True).
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
            auto_fit=auto_fit
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
