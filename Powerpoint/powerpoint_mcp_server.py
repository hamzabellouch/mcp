"""
PowerPoint MCP Server (Enterprise Edition)
------------------------------------------
Model Context Protocol (MCP) server providing high-level tools to interact with Microsoft PowerPoint.
Supports both live desktop PowerPoint automation (via COM/pywin32) and headless presentation manipulation (via python-pptx).
"""

import os
import sys
import json
from typing import List, Optional, Dict, Any, Union

# Ensure server directory is in sys.path
SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
if SERVER_DIR not in sys.path:
    sys.path.insert(0, SERVER_DIR)

try:
    from mcp.server.fastmcp import FastMCP
except Exception:
    try:
        from mcp.server.mcpserver import MCPServer as FastMCP
    except Exception:
        from mcp.server import MCPServer as FastMCP

from powerpoint_engine import PowerPointEngine

# Initialize FastMCP Server
mcp = FastMCP("powerpoint-mcp-server")


# ==========================================
# 1. Status & Presentation Management Tools
# ==========================================

@mcp.tool()
def powerpoint_status() -> str:
    """
    Check the current status of Microsoft PowerPoint and active presentations.
    Returns whether PowerPoint is running, COM availability, version, open presentations, and active presentation.
    """
    status = PowerPointEngine.get_status()
    if status["powerpoint_running"]:
        prs_names = [p["name"] for p in status["open_presentations"]]
        return (
            f"✅ Microsoft PowerPoint (v{status.get('version')}) is RUNNING.\n"
            f"- Open Presentations ({len(prs_names)}): {', '.join(prs_names) if prs_names else 'None'}\n"
            f"- Active Presentation: {status.get('active_presentation', 'None')}\n"
            f"- Active Presentation Path: {status.get('active_presentation_path', 'None')}\n"
            f"- Active Slide Index: {status.get('active_slide_index', 'None')}"
        )
    else:
        return (
            "ℹ️ Microsoft PowerPoint is currently NOT running.\n"
            "- COM automation: Available (PowerPoint will be launched if needed).\n"
            "- Headless file mode: Active (python-pptx can read/write files directly)."
        )


@mcp.tool()
def powerpoint_create_presentation(
    filepath: Optional[str] = None,
    title: Optional[str] = None,
    subtitle: Optional[str] = None,
    author: Optional[str] = None,
    widescreen: bool = True,
    theme: str = "modern_blue",
    visible: bool = True
) -> str:
    """
    Create a new Microsoft PowerPoint presentation with optional title slide, widescreen (16:9) layout, and theme palette.

    Args:
        filepath: Optional path where presentation will be saved (e.g. 'E:/presentations/AI_Pitch.pptx'). Defaults to 'Presentation.pptx'.
        title: Main presentation title (e.g. 'AI Transformation Strategy 2026').
        subtitle: Subtitle or tagline (e.g. 'Enterprise Roadmap & Business Case').
        author: Author name, presenter, or organization.
        widescreen: Use 16:9 widescreen format (default: True). If False, uses 4:3 standard format.
        theme: Theme color palette preset ('modern_blue', 'dark_tech', 'corporate_teal', 'warm_crimson', 'minimal_emerald'). Default: 'modern_blue'.
        visible: Whether to open and show presentation in live PowerPoint desktop app (default: True).
    """
    try:
        res = PowerPointEngine.create_presentation(
            filepath=filepath,
            title=title,
            subtitle=subtitle,
            author=author,
            widescreen=widescreen,
            theme=theme,
            visible=visible
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error creating presentation: {str(e)}"


@mcp.tool()
def powerpoint_open_presentation(
    filepath: str,
    read_only: bool = False,
    visible: bool = True
) -> str:
    """
    Open an existing PowerPoint presentation (.pptx, .ppt, .potx).

    Args:
        filepath: Full file path to the presentation file.
        read_only: Open the presentation in read-only mode (default: False).
        visible: Make PowerPoint application visible on desktop (default: True).
    """
    try:
        res = PowerPointEngine.open_presentation(
            filepath=filepath,
            read_only=read_only,
            visible=visible
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error opening presentation: {str(e)}"


@mcp.tool()
def powerpoint_read_presentation(
    filepath: Optional[str] = None,
    output_format: str = "markdown",
    include_notes: bool = True
) -> str:
    """
    Read full content and structure of a PowerPoint presentation (slides, titles, text boxes, tables, speaker notes).

    Args:
        filepath: Optional file path. If omitted, reads from active presentation in live PowerPoint.
        output_format: Output format ('markdown' or 'json'). Default: 'markdown'.
        include_notes: Include speaker notes in output (default: True).
    """
    try:
        return PowerPointEngine.read_presentation(
            filepath=filepath,
            output_format=output_format,
            include_notes=include_notes
        )
    except Exception as e:
        return f"❌ Error reading presentation: {str(e)}"


@mcp.tool()
def powerpoint_get_outline(filepath: Optional[str] = None) -> str:
    """
    Extract high-level outline of presentation including slide titles, shape counts, tables, charts, and word statistics.

    Args:
        filepath: Optional file path. If omitted, uses active presentation.
    """
    try:
        res = PowerPointEngine.get_outline(filepath=filepath)
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error extracting outline: {str(e)}"


@mcp.tool()
def powerpoint_save_presentation(
    filepath: Optional[str] = None,
    save_as_path: Optional[str] = None
) -> str:
    """
    Save the active or specified PowerPoint presentation, or save as a new file.

    Args:
        filepath: Optional file path of target presentation. If omitted, saves active presentation.
        save_as_path: Optional new destination path to save a copy.
    """
    try:
        res = PowerPointEngine.save_presentation(
            filepath=filepath,
            save_as_path=save_as_path
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error saving presentation: {str(e)}"


@mcp.tool()
def powerpoint_export_pdf(
    output_pdf_path: str,
    filepath: Optional[str] = None
) -> str:
    """
    Export PowerPoint presentation to high-quality PDF using native PowerPoint COM engine.

    Args:
        output_pdf_path: Destination path for the exported PDF file (e.g. 'E:/reports/PitchDeck.pdf').
        filepath: Optional source presentation file path. If omitted, exports active presentation.
    """
    try:
        res = PowerPointEngine.export_pdf(
            output_pdf_path=output_pdf_path,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error exporting PDF: {str(e)}"


@mcp.tool()
def powerpoint_export_slide_images(
    output_dir: str,
    filepath: Optional[str] = None,
    slide_index: Optional[int] = None,
    image_format: str = "PNG",
    width: int = 1920,
    height: int = 1080
) -> str:
    """
    Export slide(s) to high-resolution PNG or JPG image files.

    Args:
        output_dir: Directory where exported images will be saved.
        filepath: Optional source presentation file path. If omitted, exports from active presentation.
        slide_index: Optional 1-based slide index to export single slide. If omitted, exports all slides.
        image_format: Image format ('PNG' or 'JPG'). Default: 'PNG'.
        width: Image width in pixels (default: 1920 for Full HD).
        height: Image height in pixels (default: 1080 for Full HD).
    """
    try:
        res = PowerPointEngine.export_slide_images(
            output_dir=output_dir,
            filepath=filepath,
            slide_index=slide_index,
            image_format=image_format,
            width=width,
            height=height
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error exporting slide images: {str(e)}"


# ==========================================
# 2. Slide Operations & Layout Tools
# ==========================================

@mcp.tool()
def powerpoint_add_slide(
    filepath: Optional[str] = None,
    layout: str = "title_and_content",
    title: Optional[str] = None,
    subtitle: Optional[str] = None,
    background_color: Optional[str] = None,
    slide_index: Optional[int] = None
) -> str:
    """
    Add a new slide to the presentation with specified layout, title, and background color.

    Args:
        filepath: Optional file path. If omitted, adds to active presentation in live PowerPoint.
        layout: Layout preset name ('title', 'title_and_content', 'section_header', 'two_content', 'comparison', 'title_only', 'blank'). Default: 'title_and_content'.
        title: Optional title text for the slide.
        subtitle: Optional subtitle text for the slide.
        background_color: Optional hex background color (e.g. '#F8FAFC' or '#0F172A').
        slide_index: Optional 1-based position where the slide should be inserted. Defaults to end.
    """
    try:
        res = PowerPointEngine.add_slide(
            filepath=filepath,
            layout=layout,
            title=title,
            subtitle=subtitle,
            background_color=background_color,
            slide_index=slide_index
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding slide: {str(e)}"


@mcp.tool()
def powerpoint_duplicate_slide(
    slide_index: int,
    filepath: Optional[str] = None
) -> str:
    """
    Duplicate a slide by 1-based index.

    Args:
        slide_index: 1-based index of the slide to duplicate.
        filepath: Optional file path. If omitted, duplicates in active presentation.
    """
    try:
        res = PowerPointEngine.duplicate_slide(
            slide_index=slide_index,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error duplicating slide: {str(e)}"


@mcp.tool()
def powerpoint_delete_slide(
    slide_index: int,
    filepath: Optional[str] = None
) -> str:
    """
    Delete a slide by 1-based index.

    Args:
        slide_index: 1-based index of the slide to delete.
        filepath: Optional file path. If omitted, deletes from active presentation.
    """
    try:
        res = PowerPointEngine.delete_slide(
            slide_index=slide_index,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error deleting slide: {str(e)}"


@mcp.tool()
def powerpoint_reorder_slide(
    from_index: int,
    to_index: int,
    filepath: Optional[str] = None
) -> str:
    """
    Move a slide from one position to another (1-based indices).

    Args:
        from_index: 1-based index of the slide to move.
        to_index: 1-based target index.
        filepath: Optional file path. If omitted, reorders in active presentation.
    """
    try:
        res = PowerPointEngine.reorder_slide(
            from_index=from_index,
            to_index=to_index,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error reordering slide: {str(e)}"


@mcp.tool()
def powerpoint_set_slide_background(
    color: str,
    slide_index: Optional[int] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Set solid background color for a specific slide or all slides in the presentation.

    Args:
        color: Hex color string (e.g. '#0F172A', '#F8FAFC', '#1F497D').
        slide_index: Optional 1-based slide index. If omitted, sets background for all slides.
        filepath: Optional file path. If omitted, applies to active presentation.
    """
    try:
        res = PowerPointEngine.set_slide_background(
            color=color,
            slide_index=slide_index,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error setting slide background: {str(e)}"


@mcp.tool()
def powerpoint_set_slide_notes(
    notes: str,
    slide_index: int,
    filepath: Optional[str] = None,
    append: bool = False
) -> str:
    """
    Add or update speaker notes for a specific slide.

    Args:
        notes: Speaker notes text.
        slide_index: 1-based index of the slide.
        filepath: Optional file path. If omitted, applies to active presentation.
        append: Append to existing notes instead of overwriting (default: False).
    """
    try:
        res = PowerPointEngine.set_slide_notes(
            notes=notes,
            slide_index=slide_index,
            filepath=filepath,
            append=append
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error setting slide notes: {str(e)}"


# ==========================================
# 3. Content, Typography & Visual Tools
# ==========================================

@mcp.tool()
def powerpoint_add_text_box(
    text: str,
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 1.5,
    width_inches: float = 6.0,
    height_inches: float = 2.0,
    font_size: int = 18,
    font_name: str = "Segoe UI",
    font_color: str = "#1E293B",
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    align: str = "left",
    bg_color: Optional[str] = None,
    border_color: Optional[str] = None,
    border_width_pt: float = 1.0,
    word_wrap: bool = True,
    filepath: Optional[str] = None
) -> str:
    """
    Add a rich formatted text box to a slide with position, typography, background fill, and border styling.

    Args:
        text: Text content.
        slide_index: 1-based index of the target slide (default: 1).
        left_inches: Horizontal distance from left of slide in inches (default: 1.0).
        top_inches: Vertical distance from top of slide in inches (default: 1.5).
        width_inches: Width of text box in inches (default: 6.0).
        height_inches: Height of text box in inches (default: 2.0).
        font_size: Font size in points (default: 18).
        font_name: Font family name (default: 'Segoe UI').
        font_color: Hex font color (default: '#1E293B').
        bold: Bold text formatting (default: False).
        italic: Italic text formatting (default: False).
        underline: Underline text formatting (default: False).
        align: Text alignment ('left', 'center', 'right', 'justify'). Default: 'left'.
        bg_color: Optional background fill hex color for the text box.
        border_color: Optional border outline hex color.
        border_width_pt: Border outline thickness in points (default: 1.0).
        word_wrap: Wrap text automatically (default: True).
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_text_box(
            text=text,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            bold=bold,
            italic=italic,
            underline=underline,
            align=align,
            bg_color=bg_color,
            border_color=border_color,
            border_width_pt=border_width_pt,
            word_wrap=word_wrap,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding text box: {str(e)}"


@mcp.tool()
def powerpoint_add_bullet_points(
    items: List[Union[str, Dict[str, Any]]],
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 8.0,
    height_inches: float = 4.0,
    font_size: int = 16,
    font_name: str = "Segoe UI",
    font_color: str = "#1E293B",
    line_spacing: float = 1.3,
    filepath: Optional[str] = None
) -> str:
    """
    Add structured bullet points with multi-level nesting and line spacing.

    Args:
        items: List of strings or dicts with keys: {'text': '...', 'level': 0, 'bold': False, 'color': '#...'}.
        slide_index: 1-based index of target slide (default: 1).
        left_inches: Distance from left in inches (default: 1.0).
        top_inches: Distance from top in inches (default: 2.0).
        width_inches: Box width in inches (default: 8.0).
        height_inches: Box height in inches (default: 4.0).
        font_size: Base font size in points (default: 16).
        font_name: Font name (default: 'Segoe UI').
        font_color: Hex color string (default: '#1E293B').
        line_spacing: Line spacing multiple (default: 1.3).
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_bullet_points(
            items=items,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            font_size=font_size,
            font_name=font_name,
            font_color=font_color,
            line_spacing=line_spacing,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding bullet points: {str(e)}"


@mcp.tool()
def powerpoint_add_shape(
    shape_type: str = "rounded_rectangle",
    slide_index: int = 1,
    left_inches: float = 2.0,
    top_inches: float = 2.0,
    width_inches: float = 3.0,
    height_inches: float = 1.5,
    text: Optional[str] = None,
    fill_color: str = "#1F497D",
    border_color: Optional[str] = None,
    border_width_pt: float = 1.0,
    text_color: str = "#FFFFFF",
    font_size: int = 16,
    font_name: str = "Segoe UI",
    bold: bool = True,
    filepath: Optional[str] = None
) -> str:
    """
    Add geometric or callout shapes with custom fill, outline border, and centered text.

    Args:
        shape_type: Shape type name ('rectangle', 'rounded_rectangle', 'oval', 'circle', 'diamond', 'triangle', 'arrow_right', 'arrow_left', 'arrow_up', 'arrow_down', 'chevron', 'hexagon', 'star', 'callout', 'cloud'). Default: 'rounded_rectangle'.
        slide_index: 1-based index of the target slide (default: 1).
        left_inches: Distance from left in inches (default: 2.0).
        top_inches: Distance from top in inches (default: 2.0).
        width_inches: Shape width in inches (default: 3.0).
        height_inches: Shape height in inches (default: 1.5).
        text: Optional centered text inside the shape.
        fill_color: Hex fill color (default: '#1F497D').
        border_color: Optional outline border hex color.
        border_width_pt: Outline border width in points (default: 1.0).
        text_color: Hex text color (default: '#FFFFFF').
        font_size: Font size for text (default: 16).
        font_name: Font name (default: 'Segoe UI').
        bold: Bold text inside shape (default: True).
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_shape(
            shape_type=shape_type,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            text=text,
            fill_color=fill_color,
            border_color=border_color,
            border_width_pt=border_width_pt,
            text_color=text_color,
            font_size=font_size,
            font_name=font_name,
            bold=bold,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding shape: {str(e)}"


@mcp.tool()
def powerpoint_add_card(
    title: str,
    value: str,
    description: Optional[str] = None,
    badge: Optional[str] = None,
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 3.5,
    height_inches: float = 2.2,
    card_bg_color: str = "#FFFFFF",
    primary_color: str = "#1F497D",
    text_color: str = "#1E293B",
    border_color: str = "#E2E8F0",
    font_name: str = "Segoe UI",
    filepath: Optional[str] = None
) -> str:
    """
    Add a modern executive KPI / Metric Card (Title, big Highlight Value, Description, and Badge) with clean styling.

    Args:
        title: Card header title (e.g. 'Annual Revenue' or 'إجمالي الإيرادات').
        value: Large stat / KPI value (e.g. '$4.2M' or '+38% YoY').
        description: Descriptive body text (e.g. 'Target achieved across 12 countries').
        badge: Optional status badge (e.g. 'TOP PERFORMER' or 'ON TRACK').
        slide_index: 1-based index of target slide (default: 1).
        left_inches: Distance from left in inches (default: 1.0).
        top_inches: Distance from top in inches (default: 2.0).
        width_inches: Card width in inches (default: 3.5).
        height_inches: Card height in inches (default: 2.2).
        card_bg_color: Background color for the card (default: White '#FFFFFF').
        primary_color: Highlight color for big metric value (default: Navy '#1F497D').
        text_color: Body text color (default: '#1E293B').
        border_color: Subtle border stroke color (default: '#E2E8F0').
        font_name: Font name (default: 'Segoe UI').
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_card(
            title=title,
            value=value,
            description=description,
            badge=badge,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            card_bg_color=card_bg_color,
            primary_color=primary_color,
            text_color=text_color,
            border_color=border_color,
            font_name=font_name,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding KPI card: {str(e)}"


@mcp.tool()
def powerpoint_add_image(
    image_path_or_url: str,
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 1.5,
    width_inches: Optional[float] = 5.0,
    height_inches: Optional[float] = None,
    caption: Optional[str] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Insert an image from local file path or public URL with custom dimensions and optional caption.

    Args:
        image_path_or_url: Absolute local file path (e.g. 'E:/assets/chart.png') or HTTP(S) URL.
        slide_index: 1-based slide index (default: 1).
        left_inches: Distance from left in inches (default: 1.0).
        top_inches: Distance from top in inches (default: 1.5).
        width_inches: Width in inches (default: 5.0).
        height_inches: Optional height in inches (maintains aspect ratio if omitted).
        caption: Optional caption text placed below image.
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_image(
            image_path_or_url=image_path_or_url,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            caption=caption,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding image: {str(e)}"


@mcp.tool()
def powerpoint_add_table(
    headers: List[str],
    data: List[List[Any]],
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 11.0,
    height_inches: float = 3.5,
    header_bg_color: str = "#1F497D",
    header_font_color: str = "#FFFFFF",
    alternate_row_color: Optional[str] = "#F8FAFC",
    font_size: int = 13,
    font_name: str = "Segoe UI",
    col_widths: Optional[List[float]] = None,
    filepath: Optional[str] = None
) -> str:
    """
    Add a professionally styled table with custom header fills, zebra row striping, and column widths.

    Args:
        headers: List of column header names (e.g. ['Quarter', 'Revenue', 'Growth', 'Status']).
        data: 2D list of rows and cell values.
        slide_index: 1-based slide index (default: 1).
        left_inches: Distance from left in inches (default: 1.0).
        top_inches: Distance from top in inches (default: 2.0).
        width_inches: Total table width in inches (default: 11.0).
        height_inches: Total table height in inches (default: 3.5).
        header_bg_color: Hex fill color for header row (default: Navy '#1F497D').
        header_font_color: Hex text color for header row (default: White '#FFFFFF').
        alternate_row_color: Hex fill color for zebra alternating rows (default: '#F8FAFC').
        font_size: Font size in points (default: 13).
        font_name: Font name (default: 'Segoe UI').
        col_widths: Optional list of individual column widths in inches (e.g. [3.0, 2.5, 2.5, 3.0]).
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_table(
            headers=headers,
            data=data,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            header_bg_color=header_bg_color,
            header_font_color=header_font_color,
            alternate_row_color=alternate_row_color,
            font_size=font_size,
            font_name=font_name,
            col_widths=col_widths,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding table: {str(e)}"


@mcp.tool()
def powerpoint_add_chart(
    chart_type: str,
    categories: List[str],
    series_data: List[Dict[str, Any]],
    title: Optional[str] = None,
    slide_index: int = 1,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 8.0,
    height_inches: float = 4.5,
    filepath: Optional[str] = None
) -> str:
    """
    Add a native PowerPoint chart (Column, Bar, Line, Pie, Doughnut, Area) with series data and legend.

    Args:
        chart_type: Chart type ('column', 'column_stacked', 'bar', 'bar_stacked', 'line', 'line_markers', 'pie', 'doughnut', 'area').
        categories: List of category labels (e.g. ['Q1', 'Q2', 'Q3', 'Q4']).
        series_data: List of series dicts: [{'name': '2025', 'values': [100, 120, 140, 160]}, {'name': '2026', 'values': [130, 150, 180, 210]}].
        title: Optional chart title.
        slide_index: 1-based slide index (default: 1).
        left_inches: Distance from left in inches (default: 1.0).
        top_inches: Distance from top in inches (default: 2.0).
        width_inches: Chart width in inches (default: 8.0).
        height_inches: Chart height in inches (default: 4.5).
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_chart(
            chart_type=chart_type,
            categories=categories,
            series_data=series_data,
            title=title,
            slide_index=slide_index,
            left_inches=left_inches,
            top_inches=top_inches,
            width_inches=width_inches,
            height_inches=height_inches,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding chart: {str(e)}"


@mcp.tool()
def powerpoint_add_timeline(
    milestones: List[Dict[str, str]],
    slide_index: int = 1,
    title: Optional[str] = "Project Roadmap & Timeline",
    top_inches: float = 2.5,
    primary_color: str = "#1F497D",
    accent_color: str = "#418AB3",
    filepath: Optional[str] = None
) -> str:
    """
    Create a visual roadmap / timeline slide with sequential milestone nodes, dates, titles, descriptions, and connectors.

    Args:
        milestones: List of milestone dicts: [{'date': 'Q1 2026', 'title': 'Discovery', 'description': 'Scope & Architecture'}, ...].
        slide_index: 1-based slide index (default: 1).
        title: Slide title (default: 'Project Roadmap & Timeline').
        top_inches: Vertical baseline position in inches (default: 2.5).
        primary_color: Primary node color (default: '#1F497D').
        accent_color: Accent border and date color (default: '#418AB3').
        filepath: Optional file path. If omitted, adds to active presentation.
    """
    try:
        res = PowerPointEngine.add_timeline(
            milestones=milestones,
            slide_index=slide_index,
            title=title,
            top_inches=top_inches,
            primary_color=primary_color,
            accent_color=accent_color,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error adding timeline: {str(e)}"


# ==========================================
# 4. Enterprise & Advanced Tools
# ==========================================

@mcp.tool()
def powerpoint_search_and_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    filepath: Optional[str] = None
) -> str:
    """
    Search and replace text across all slides, text frames, tables, and speaker notes.

    Args:
        find_text: Text string to find.
        replace_text: Replacement text.
        match_case: Case-sensitive search (default: False).
        filepath: Optional file path. If omitted, applies to active presentation.
    """
    try:
        res = PowerPointEngine.search_and_replace(
            find_text=find_text,
            replace_text=replace_text,
            match_case=match_case,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error during search and replace: {str(e)}"


@mcp.tool()
def powerpoint_fill_template(
    template_path: str,
    output_path: str,
    replacements: Dict[str, Any]
) -> str:
    """
    Fill a PowerPoint template presentation by replacing {{placeholders}} with dynamic values.

    Args:
        template_path: Path to template .pptx file containing {{tags}}.
        output_path: Destination path for filled presentation.
        replacements: Dictionary of tag names to replacement values (e.g. {'client': 'Acme Corp', 'date': '2026-09-29'}).
    """
    try:
        res = PowerPointEngine.fill_template(
            template_path=template_path,
            output_path=output_path,
            replacements=replacements
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error filling template: {str(e)}"


@mcp.tool()
def powerpoint_apply_theme(
    theme: str,
    filepath: Optional[str] = None
) -> str:
    """
    Apply a cohesive color palette and font theme across all slides in the presentation.

    Args:
        theme: Theme preset ('modern_blue', 'dark_tech', 'corporate_teal', 'warm_crimson', 'minimal_emerald').
        filepath: Optional file path. If omitted, applies to active presentation.
    """
    try:
        res = PowerPointEngine.apply_theme(
            theme=theme,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error applying theme: {str(e)}"


@mcp.tool()
def powerpoint_execute_script(
    code: str,
    filepath: Optional[str] = None
) -> str:
    """
    Execute custom Python script with access to presentation objects (prs, pptx, app, Inches, Pt, RGBColor).

    Args:
        code: Python script to execute.
        filepath: Optional target presentation file path.
    """
    try:
        res = PowerPointEngine.execute_script(
            code=code,
            filepath=filepath
        )
        return json.dumps(res, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"❌ Error executing script: {str(e)}"


# ==========================================
# Main Entry Point
# ==========================================

def main():
    """Main entrypoint for CLI and script execution."""
    mcp.run()


if __name__ == "__main__":
    main()
