"""
PowerPoint Automation Engine (Enterprise Edition)
------------------------------------------------
Provides a comprehensive, production-grade abstraction for interacting with Microsoft PowerPoint.
Supports:
1. Live Windows COM Automation (via pywin32) for interactive real-time control of Microsoft PowerPoint desktop application.
2. Headless python-pptx engine for direct, high-performance file manipulation without requiring PowerPoint to be installed or open.
"""

import os
import io
import sys
import json
import re
import traceback
import tempfile
import urllib.request
import urllib.parse
from typing import List, Dict, Any, Optional, Union, Tuple

# Headless python-pptx engine
import pptx
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

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
    hex_str = str(hex_str).strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color string: '{hex_str}'. Expected 6-digit hex like '#1F497D'.")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return r, g, b


def hex_to_rgb_color(hex_str: str) -> RGBColor:
    """Convert hex color string to python-pptx RGBColor."""
    r, g, b = parse_hex_color(hex_str)
    return RGBColor(r, g, b)


def hex_to_bgr_int(hex_str: str) -> int:
    """Convert hex color string to PowerPoint COM BGR integer (RGB value in COM is (R) + (G*256) + (B*65536))."""
    r, g, b = parse_hex_color(hex_str)
    return r + (g * 256) + (b * 65536)


def clean_hex(hex_str: str) -> str:
    """Return 6-character hex string without leading '#'."""
    hex_str = str(hex_str).strip().lstrip('#')
    if len(hex_str) == 3:
        hex_str = ''.join([c * 2 for c in hex_str])
    return hex_str.upper()


def is_arabic_text(text: str) -> bool:
    """Detect if string contains Arabic/Persian/Hebrew characters for RTL formatting."""
    return bool(re.search(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\u0590-\u05FF]', str(text)))


def cleanup_stale_lock_files(target_path: Optional[str] = None):
    """Remove orphaned Microsoft PowerPoint ~$*.pptx hidden lock files when PowerPoint is not actively running."""
    try:
        # Never touch lock files if PowerPoint is currently running to prevent corruption
        if COM_AVAILABLE and LivePowerPointSession.is_running():
            return
        if target_path:
            dir_path = os.path.dirname(os.path.abspath(target_path)) if (os.path.isfile(target_path) or '.' in os.path.basename(target_path)) else os.path.abspath(target_path)
        else:
            dir_path = os.getcwd()
        if os.path.exists(dir_path):
            for fname in os.listdir(dir_path):
                if fname.startswith("~$") and fname.endswith(".pptx"):
                    lock_file = os.path.join(dir_path, fname)
                    try:
                        if sys.platform == "win32":
                            try:
                                import ctypes
                                ctypes.windll.kernel32.SetFileAttributesW(lock_file, 128)
                            except Exception:
                                pass
                        os.remove(lock_file)
                    except Exception:
                        pass
    except Exception:
        pass


def get_pptx_alignment(align_str: Optional[str]) -> PP_ALIGN:
    """Convert string ('left', 'center', 'right', 'justify') to pptx PP_ALIGN."""
    if not align_str:
        return PP_ALIGN.LEFT
    a = align_str.strip().lower()
    if a in ("center", "centre"):
        return PP_ALIGN.CENTER
    elif a in ("right", "end"):
        return PP_ALIGN.RIGHT
    elif a in ("justify", "justified"):
        return PP_ALIGN.JUSTIFY
    return PP_ALIGN.LEFT


def get_com_alignment(align_str: Optional[str]) -> int:
    """Convert string to PowerPoint COM alignment constant (ppAlign*)."""
    if not align_str:
        return 1  # ppAlignLeft
    a = align_str.strip().lower()
    if a in ("center", "centre"):
        return 2  # ppAlignCenter
    elif a in ("right", "end"):
        return 3  # ppAlignRight
    elif a in ("justify", "justified"):
        return 4  # ppAlignJustify
    return 1  # ppAlignLeft


# ==========================================
# Preset Mappings & Definitions
# ==========================================

# Shape Name to python-pptx MSO_SHAPE mapping
SHAPE_TYPE_MAP = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "oval": MSO_SHAPE.OVAL,
    "circle": MSO_SHAPE.OVAL,
    "diamond": MSO_SHAPE.DIAMOND,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "right_triangle": MSO_SHAPE.RIGHT_TRIANGLE,
    "arrow_right": MSO_SHAPE.RIGHT_ARROW,
    "arrow_left": MSO_SHAPE.LEFT_ARROW,
    "arrow_up": MSO_SHAPE.UP_ARROW,
    "arrow_down": MSO_SHAPE.DOWN_ARROW,
    "chevron": MSO_SHAPE.CHEVRON,
    "pentagon": MSO_SHAPE.PENTAGON,
    "hexagon": MSO_SHAPE.HEXAGON,
    "octagon": MSO_SHAPE.OCTAGON,
    "star_5": getattr(MSO_SHAPE, "5_POINT_STAR", getattr(MSO_SHAPE, "STAR_5_POINT", MSO_SHAPE.RECTANGLE)),
    "star": getattr(MSO_SHAPE, "STAR_5_POINT", MSO_SHAPE.RECTANGLE),
    "heart": MSO_SHAPE.HEART,
    "cloud": MSO_SHAPE.CLOUD,
    "callout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "rounded_callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    "oval_callout": MSO_SHAPE.OVAL_CALLOUT,
    "cube": MSO_SHAPE.CUBE,
    "donut": MSO_SHAPE.DONUT,
    "can": MSO_SHAPE.CAN,
}

# COM Shape Type mapping (MsoAutoShapeType)
COM_SHAPE_MAP = {
    "rectangle": 1,          # msoShapeRectangle
    "rect": 1,
    "rounded_rectangle": 5,  # msoShapeRoundedRectangle
    "round_rect": 5,
    "oval": 9,               # msoShapeOval
    "circle": 9,
    "diamond": 4,            # msoShapeDiamond
    "triangle": 7,           # msoShapeIsoscelesTriangle
    "arrow_right": 13,       # msoShapeRightArrow
    "arrow_left": 14,        # msoShapeLeftArrow
    "arrow_up": 15,          # msoShapeUpArrow
    "arrow_down": 16,        # msoShapeDownArrow
    "chevron": 52,           # msoShapeChevron
    "pentagon": 51,          # msoShapePentagon
    "hexagon": 10,           # msoShapeHexagon
    "octagon": 6,            # msoShapeOctagon
    "star": 92,              # msoShape5pointStar
    "heart": 21,             # msoShapeHeart
    "cloud": 29,             # msoShapeCloud
    "callout": 61,           # msoShapeRectangularCallout
}

# Layout Name to Slide Layout Index
LAYOUT_MAP = {
    "title": 0,
    "title_and_content": 1,
    "title_content": 1,
    "content": 1,
    "section_header": 2,
    "two_content": 3,
    "comparison": 4,
    "title_only": 5,
    "blank": 6,
    "content_with_caption": 7,
    "picture_with_caption": 8
}

# Chart Name to XL_CHART_TYPE mapping
CHART_TYPE_MAP = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "column_stacked_100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE,
    "line_markers": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
    "area_stacked": XL_CHART_TYPE.AREA_STACKED,
}

# Theme Palettes
THEME_PALETTES = {
    "modern_blue": {
        "primary": "#1F497D",
        "secondary": "#418AB3",
        "accent": "#F2AF00",
        "bg_light": "#F8FAFC",
        "bg_dark": "#0F172A",
        "text_dark": "#1E293B",
        "text_light": "#FFFFFF",
        "card_bg": "#FFFFFF",
        "border": "#E2E8F0",
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    },
    "dark_tech": {
        "primary": "#38BDF8",
        "secondary": "#818CF8",
        "accent": "#F43F5E",
        "bg_light": "#1E293B",
        "bg_dark": "#0B0F19",
        "text_dark": "#F8FAFC",
        "text_light": "#FFFFFF",
        "card_bg": "#1E293B",
        "border": "#334155",
        "font_heading": "Segoe UI Semibold",
        "font_body": "Segoe UI"
    },
    "corporate_teal": {
        "primary": "#0D9488",
        "secondary": "#14B8A6",
        "accent": "#F59E0B",
        "bg_light": "#F0FDFA",
        "bg_dark": "#134E4A",
        "text_dark": "#0F172A",
        "text_light": "#FFFFFF",
        "card_bg": "#FFFFFF",
        "border": "#CCFBF1",
        "font_heading": "Calibri",
        "font_body": "Calibri"
    },
    "warm_crimson": {
        "primary": "#991B1B",
        "secondary": "#DC2626",
        "accent": "#D97706",
        "bg_light": "#FEF2F2",
        "bg_dark": "#450A0A",
        "text_dark": "#1C1917",
        "text_light": "#FFFFFF",
        "card_bg": "#FFFFFF",
        "border": "#FEE2E2",
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    },
    "minimal_emerald": {
        "primary": "#047857",
        "secondary": "#10B981",
        "accent": "#6366F1",
        "bg_light": "#F0FDF4",
        "bg_dark": "#064E3B",
        "text_dark": "#111827",
        "text_light": "#FFFFFF",
        "card_bg": "#FFFFFF",
        "border": "#D1FAE5",
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }
}


# ==========================================
# Live PowerPoint COM Session Manager
# ==========================================

class LivePowerPointSession:
    """Manages interactive connection to Microsoft PowerPoint desktop application via pywin32 COM."""

    @staticmethod
    def get_app(create_if_missing: bool = False, visible: bool = True):
        if not COM_AVAILABLE:
            raise RuntimeError("pywin32 (win32com) is not installed in this Python environment.")
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            if visible and not app.Visible:
                app.Visible = True
            return app
        except Exception:
            if create_if_missing:
                app = win32com.client.Dispatch("PowerPoint.Application")
                if visible:
                    app.Visible = True
                return app
            return None

    @staticmethod
    def is_running() -> bool:
        if not COM_AVAILABLE:
            return False
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            return app is not None
        except Exception:
            return False

    @staticmethod
    def get_active_presentation(app=None):
        if app is None:
            app = LivePowerPointSession.get_app(create_if_missing=False)
        if not app:
            return None
        try:
            return app.ActivePresentation
        except Exception:
            if app.Presentations.Count > 0:
                return app.Presentations(1)
            return None

    @staticmethod
    def find_presentation_by_path(filepath: str, app=None):
        if app is None:
            app = LivePowerPointSession.get_app(create_if_missing=False)
        if not app:
            return None
        target = os.path.abspath(filepath).lower()
        try:
            for i in range(1, app.Presentations.Count + 1):
                p = app.Presentations(i)
                try:
                    if p.FullName and os.path.abspath(p.FullName).lower() == target:
                        return p
                    if os.path.basename(target) == p.Name.lower():
                        return p
                except Exception:
                    pass
        except Exception:
            pass
        return None


def save_pptx_safely(prs: pptx.Presentation, target_path: str):
    """Save python-pptx Presentation safely, closing and reopening in COM if open in live PowerPoint."""
    import time
    abs_path = os.path.abspath(target_path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    com_prs_was_open = False
    app = None
    if COM_AVAILABLE:
        try:
            pythoncom.CoInitialize()
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(abs_path, app)
                if prs_com:
                    com_prs_was_open = True
                    try:
                        prs_com.Saved = True
                        prs_com.Close()
                    except Exception:
                        pass
        except Exception:
            pass

    time.sleep(0.05)
    prs.save(abs_path)

    if com_prs_was_open and app:
        try:
            app.Presentations.Open(abs_path)
        except Exception:
            pass


# ==========================================
# Unified PowerPoint Automation Engine
# ==========================================

class PowerPointEngine:
    """Enterprise-grade dual-mode engine for Microsoft PowerPoint automation."""

    @classmethod
    def get_status(cls) -> Dict[str, Any]:
        """Get live status of PowerPoint and active presentations."""
        cleanup_stale_lock_files()
        is_running = LivePowerPointSession.is_running()
        result = {
            "powerpoint_running": is_running,
            "com_available": COM_AVAILABLE,
            "open_presentations": [],
            "active_presentation": None,
            "active_presentation_path": None,
            "active_slide_index": None,
            "version": None
        }

        if is_running:
            try:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    result["version"] = getattr(app, "Version", "Unknown")
                    for i in range(1, app.Presentations.Count + 1):
                        prs = app.Presentations(i)
                        try:
                            result["open_presentations"].append({
                                "name": prs.Name,
                                "path": prs.FullName if hasattr(prs, "FullName") else None,
                                "slides_count": prs.Slides.Count,
                                "saved": prs.Saved
                            })
                        except Exception:
                            pass
                    
                    active_prs = LivePowerPointSession.get_active_presentation(app)
                    if active_prs:
                        result["active_presentation"] = active_prs.Name
                        try:
                            result["active_presentation_path"] = active_prs.FullName
                        except Exception:
                            pass
                        try:
                            result["active_slide_index"] = app.ActiveWindow.View.Slide.SlideIndex
                        except Exception:
                            pass
            except Exception as e:
                result["com_error"] = str(e)

        return result

    @classmethod
    def create_presentation(
        cls,
        filepath: Optional[str] = None,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        author: Optional[str] = None,
        widescreen: bool = True,
        theme: str = "modern_blue",
        visible: bool = True
    ) -> Dict[str, Any]:
        """Create a new presentation with 16:9 or 4:3 aspect ratio and optional title slide."""
        cleanup_stale_lock_files(filepath)
        palette = THEME_PALETTES.get(theme, THEME_PALETTES["modern_blue"])

        # Create using python-pptx
        prs = pptx.Presentation()
        if widescreen:
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        else:
            prs.slide_width = Inches(10.0)
            prs.slide_height = Inches(7.5)

        # Add Title Slide if title provided
        if title:
            slide_layout = prs.slide_layouts[0]  # Title slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Apply background color
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb_color(palette["bg_light"])

            # Set title text
            if slide.shapes.title:
                slide.shapes.title.text = title
                tf = slide.shapes.title.text_frame
                for p in tf.paragraphs:
                    p.font.name = palette["font_heading"]
                    p.font.size = Pt(40)
                    p.font.bold = True
                    p.font.color.rgb = hex_to_rgb_color(palette["primary"])
                    if is_arabic_text(title):
                        p.alignment = PP_ALIGN.RIGHT

            # Set subtitle / author
            sub_text_parts = []
            if subtitle:
                sub_text_parts.append(subtitle)
            if author:
                sub_text_parts.append(f"Author: {author}")
            
            if len(slide.placeholders) > 1 and sub_text_parts:
                subtitle_shape = slide.placeholders[1]
                subtitle_shape.text = "\n".join(sub_text_parts)
                for p in subtitle_shape.text_frame.paragraphs:
                    p.font.name = palette["font_body"]
                    p.font.size = Pt(20)
                    p.font.color.rgb = hex_to_rgb_color(palette["secondary"])
                    if is_arabic_text(subtitle_shape.text):
                        p.alignment = PP_ALIGN.RIGHT

        out_path = os.path.abspath(filepath) if filepath else os.path.abspath("Presentation.pptx")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        prs.save(out_path)

        # If COM is running or visible requested, open in live PowerPoint
        live_opened = False
        if COM_AVAILABLE and visible:
            try:
                app = LivePowerPointSession.get_app(create_if_missing=True, visible=True)
                if app:
                    app.Presentations.Open(out_path)
                    live_opened = True
            except Exception:
                pass

        return {
            "status": "success",
            "message": f"Presentation created successfully at '{out_path}'.",
            "filepath": out_path,
            "slides_count": len(prs.slides),
            "widescreen": widescreen,
            "theme": theme,
            "live_opened": live_opened
        }

    @classmethod
    def open_presentation(
        cls,
        filepath: str,
        read_only: bool = False,
        visible: bool = True
    ) -> Dict[str, Any]:
        """Open an existing presentation in live PowerPoint or prepare for editing."""
        cleanup_stale_lock_files(filepath)
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Presentation file not found at: {abs_path}")

        # Open in live COM if available and visible requested
        live_connected = False
        slides_count = 0
        if COM_AVAILABLE and visible:
            try:
                app = LivePowerPointSession.get_app(create_if_missing=True, visible=True)
                if app:
                    prs = LivePowerPointSession.find_presentation_by_path(abs_path, app)
                    if not prs:
                        prs = app.Presentations.Open(abs_path, ReadOnly=read_only)
                    slides_count = prs.Slides.Count
                    live_connected = True
            except Exception:
                pass

        if not live_connected:
            prs = pptx.Presentation(abs_path)
            slides_count = len(prs.slides)

        return {
            "status": "success",
            "message": f"Presentation '{os.path.basename(abs_path)}' opened successfully.",
            "filepath": abs_path,
            "slides_count": slides_count,
            "live_connected": live_connected,
            "read_only": read_only
        }

    @classmethod
    def save_presentation(
        cls,
        filepath: Optional[str] = None,
        save_as_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """Save active or specified presentation."""
        target_path = os.path.abspath(filepath) if filepath else None
        save_as = os.path.abspath(save_as_path) if save_as_path else None

        if COM_AVAILABLE:
            try:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs = None
                    if target_path:
                        prs = LivePowerPointSession.find_presentation_by_path(target_path, app)
                    if not prs:
                        prs = LivePowerPointSession.get_active_presentation(app)
                    
                    if prs:
                        if save_as:
                            os.makedirs(os.path.dirname(save_as), exist_ok=True)
                            try:
                                prs.SaveCopyAs(save_as)
                            except Exception:
                                try:
                                    prs.SaveAs(save_as)
                                except Exception:
                                    src = getattr(prs, "FullName", target_path)
                                    if src and os.path.exists(src):
                                        import shutil
                                        shutil.copy2(src, save_as)
                            return {"status": "success", "message": f"Presentation saved as '{save_as}'.", "filepath": save_as}
                        else:
                            prs.Save()
                            return {"status": "success", "message": f"Active presentation '{prs.Name}' saved successfully.", "filepath": getattr(prs, "FullName", target_path)}
            except Exception:
                pass

        if target_path and os.path.exists(target_path):
            if save_as:
                os.makedirs(os.path.dirname(save_as), exist_ok=True)
                import shutil
                shutil.copy2(target_path, save_as)
                return {"status": "success", "message": f"Presentation saved as '{save_as}'.", "filepath": save_as}
            return {"status": "success", "message": f"File '{target_path}' is saved and up to date.", "filepath": target_path}

        raise RuntimeError("No active presentation or valid filepath provided to save.")

    @classmethod
    def export_pdf(
        cls,
        output_pdf_path: str,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Export presentation to high-quality PDF using native PowerPoint COM engine."""
        cleanup_stale_lock_files(filepath)
        out_pdf = os.path.abspath(output_pdf_path)
        os.makedirs(os.path.dirname(out_pdf), exist_ok=True)

        if not COM_AVAILABLE:
            raise RuntimeError("Live COM is required for native PDF export. Please ensure pywin32 is installed and PowerPoint is available.")

        pythoncom.CoInitialize()
        app = LivePowerPointSession.get_app(create_if_missing=True, visible=False)
        if not app:
            raise RuntimeError("Unable to launch Microsoft PowerPoint for PDF export.")

        close_after = False
        prs = None
        if filepath:
            abs_fp = os.path.abspath(filepath)
            prs = LivePowerPointSession.find_presentation_by_path(abs_fp, app)
            if not prs:
                prs = app.Presentations.Open(abs_fp, WithWindow=False)
                close_after = True
        else:
            prs = LivePowerPointSession.get_active_presentation(app)

        if not prs:
            raise RuntimeError("No open presentation found to export to PDF.")

        try:
            # ppSaveAsPDF = 32
            prs.SaveAs(out_pdf, 32)
        finally:
            if close_after:
                try:
                    prs.Close()
                except Exception:
                    pass

        if not os.path.exists(out_pdf):
            raise RuntimeError(f"Failed to generate PDF at '{out_pdf}'.")

        return {
            "status": "success",
            "message": f"Presentation successfully exported to PDF at '{out_pdf}'.",
            "pdf_path": out_pdf,
            "size_bytes": os.path.getsize(out_pdf)
        }

    @classmethod
    def export_slide_images(
        cls,
        output_dir: str,
        filepath: Optional[str] = None,
        slide_index: Optional[int] = None,
        image_format: str = "PNG",
        width: int = 1920,
        height: int = 1080
    ) -> Dict[str, Any]:
        """Export individual slide or all slides as high-resolution images (.png or .jpg)."""
        cleanup_stale_lock_files(filepath)
        out_dir = os.path.abspath(output_dir)
        os.makedirs(out_dir, exist_ok=True)
        img_format = image_format.upper()
        if img_format not in ("PNG", "JPG", "JPEG"):
            img_format = "PNG"

        if not COM_AVAILABLE:
            raise RuntimeError("PowerPoint COM is required for slide image export.")

        pythoncom.CoInitialize()
        app = LivePowerPointSession.get_app(create_if_missing=True, visible=False)
        if not app:
            raise RuntimeError("Unable to launch PowerPoint for image export.")

        close_after = False
        prs = None
        if filepath:
            abs_fp = os.path.abspath(filepath)
            prs = LivePowerPointSession.find_presentation_by_path(abs_fp, app)
            if not prs:
                prs = app.Presentations.Open(abs_fp, WithWindow=False)
                close_after = True
        else:
            prs = LivePowerPointSession.get_active_presentation(app)

        if not prs:
            raise RuntimeError("No presentation found to export slides from.")

        generated_files = []
        try:
            if slide_index is not None:
                # 1-indexed
                slide = prs.Slides(slide_index)
                img_path = os.path.join(out_dir, f"slide_{slide_index}.{img_format.lower()}")
                slide.Export(img_path, img_format, width, height)
                generated_files.append(img_path)
            else:
                for idx in range(1, prs.Slides.Count + 1):
                    slide = prs.Slides(idx)
                    img_path = os.path.join(out_dir, f"slide_{idx}.{img_format.lower()}")
                    slide.Export(img_path, img_format, width, height)
                    generated_files.append(img_path)
        finally:
            if close_after:
                try:
                    prs.Close()
                except Exception:
                    pass

        return {
            "status": "success",
            "message": f"Exported {len(generated_files)} slide image(s) to '{out_dir}'.",
            "output_directory": out_dir,
            "images": generated_files
        }

    @classmethod
    def read_presentation(
        cls,
        filepath: Optional[str] = None,
        output_format: str = "markdown",
        include_notes: bool = True
    ) -> Union[str, Dict[str, Any]]:
        """Read full presentation structure and content into structured Markdown or JSON."""
        # Find file or active presentation
        target_path = None
        if filepath:
            target_path = os.path.abspath(filepath)
        elif COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    try:
                        target_path = prs_com.FullName
                    except Exception:
                        pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath or 'No active presentation'}")

        prs = pptx.Presentation(target_path)
        slides_data = []

        for idx, slide in enumerate(prs.slides, start=1):
            slide_info = {
                "slide_index": idx,
                "title": "",
                "texts": [],
                "tables": [],
                "shapes": [],
                "speaker_notes": ""
            }

            # Title
            if slide.shapes.title and slide.shapes.title.text:
                slide_info["title"] = slide.shapes.title.text.strip()

            # Shapes and text
            for shape in slide.shapes:
                if shape.has_text_frame:
                    t = shape.text_frame.text.strip()
                    if t and t != slide_info["title"]:
                        slide_info["texts"].append(t)
                
                if shape.has_table:
                    tbl = shape.table
                    table_matrix = []
                    for row in tbl.rows:
                        row_vals = [cell.text.strip() for cell in row.cells]
                        table_matrix.append(row_vals)
                    slide_info["tables"].append(table_matrix)

                if shape.shape_type:
                    slide_info["shapes"].append(str(shape.name))

            # Notes
            if include_notes and slide.has_notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                if notes_tf and notes_tf.text:
                    slide_info["speaker_notes"] = notes_tf.text.strip()

            slides_data.append(slide_info)

        if output_format.lower() == "json":
            return json.dumps({
                "filepath": target_path,
                "total_slides": len(slides_data),
                "slides": slides_data
            }, indent=2, ensure_ascii=False)

        # Markdown format
        md_lines = [
            f"# 📊 Presentation Overview: {os.path.basename(target_path)}",
            f"**Total Slides:** {len(slides_data)}",
            "---"
        ]

        for s in slides_data:
            title_display = s["title"] if s["title"] else f"Slide {s['slide_index']}"
            md_lines.append(f"\n## Slide {s['slide_index']}: {title_display}")
            
            if s["texts"]:
                for t in s["texts"]:
                    md_lines.append(f"- {t}")

            if s["tables"]:
                for tbl_idx, tbl in enumerate(s["tables"], start=1):
                    md_lines.append(f"\n**Table {tbl_idx}:**")
                    if tbl:
                        headers = tbl[0]
                        md_lines.append("| " + " | ".join(headers) + " |")
                        md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                        for r in tbl[1:]:
                            md_lines.append("| " + " | ".join(r) + " |")

            if include_notes and s["speaker_notes"]:
                md_lines.append(f"\n> 📝 **Speaker Notes:** {s['speaker_notes']}")

            md_lines.append("\n---")

        return "\n".join(md_lines)

    @classmethod
    def get_outline(cls, filepath: Optional[str] = None) -> Dict[str, Any]:
        """Extract high-level outline, slide titles, counts, and word statistics."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        outline = []
        total_words = 0

        for idx, slide in enumerate(prs.slides, start=1):
            title = slide.shapes.title.text.strip() if (slide.shapes.title and slide.shapes.title.text) else f"[Untitled Slide {idx}]"
            shape_count = len(slide.shapes)
            has_table = any(s.has_table for s in slide.shapes)
            has_chart = any(s.has_chart for s in slide.shapes)
            
            slide_text = " ".join([s.text_frame.text for s in slide.shapes if s.has_text_frame])
            words = len(slide_text.split())
            total_words += words

            outline.append({
                "slide_index": idx,
                "title": title,
                "shape_count": shape_count,
                "has_table": has_table,
                "has_chart": has_chart,
                "word_count": words
            })

        return {
            "filepath": target_path,
            "total_slides": len(prs.slides),
            "total_words": total_words,
            "outline": outline
        }

    @classmethod
    def add_slide(
        cls,
        filepath: Optional[str] = None,
        layout: str = "title_and_content",
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        background_color: Optional[str] = None,
        slide_index: Optional[int] = None
    ) -> Dict[str, Any]:
        """Add a new slide with specified layout, title, and background."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        # COM Live Modification if active
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    layout_val = 2  # ppLayoutText
                    if layout in ("title", "title_slide"):
                        layout_val = 1
                    elif layout in ("blank", "empty"):
                        layout_val = 12
                    elif layout in ("two_content", "comparison"):
                        layout_val = 3
                    
                    pos = slide_index if slide_index is not None else (prs_com.Slides.Count + 1)
                    slide = prs_com.Slides.Add(pos, layout_val)
                    if title and slide.Shapes.HasTitle:
                        slide.Shapes.Title.TextFrame.TextRange.Text = title
                    if background_color:
                        bgr = hex_to_bgr_int(background_color)
                        slide.Background.Fill.Solid()
                        slide.Background.Fill.ForeColor.RGB = bgr
                    
                    if target_path:
                        prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Slide added at index {slide.SlideIndex} in live PowerPoint.",
                        "slide_index": slide.SlideIndex,
                        "total_slides": prs_com.Slides.Count
                    }

        # Headless python-pptx modification
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        layout_idx = LAYOUT_MAP.get(layout.lower(), 1)
        if layout_idx >= len(prs.slide_layouts):
            layout_idx = 1
        slide_layout = prs.slide_layouts[layout_idx]
        
        slide = prs.slides.add_slide(slide_layout)
        
        # If slide_index specified and not at end, python-pptx appends to end, we can record position
        actual_index = len(prs.slides)

        if title and slide.shapes.title:
            slide.shapes.title.text = title
            if is_arabic_text(title):
                for p in slide.shapes.title.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.RIGHT

        if subtitle and len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle
            if is_arabic_text(subtitle):
                for p in slide.placeholders[1].text_frame.paragraphs:
                    p.alignment = PP_ALIGN.RIGHT

        if background_color:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb_color(background_color)

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Slide added successfully at index {actual_index}.",
            "slide_index": actual_index,
            "total_slides": len(prs.slides)
        }

    @classmethod
    def duplicate_slide(
        cls,
        slide_index: int,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Duplicate a slide at the specified 1-based index."""
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(filepath, app) if filepath else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    if slide_index < 1 or slide_index > prs_com.Slides.Count:
                        raise ValueError(f"Invalid slide index {slide_index}. Presentation has {prs_com.Slides.Count} slides.")
                    slide = prs_com.Slides(slide_index)
                    dup = slide.Duplicate()
                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Slide {slide_index} duplicated successfully in live PowerPoint.",
                        "new_slide_index": dup.SlideIndex,
                        "total_slides": prs_com.Slides.Count
                    }

        # Headless mode
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        if not COM_AVAILABLE:
            raise RuntimeError("Slide duplication requires Microsoft PowerPoint COM automation (pywin32) on Windows.")

        # Headless python-pptx duplication via COM fallback
        pythoncom.CoInitialize()
        app = LivePowerPointSession.get_app(create_if_missing=True, visible=False)
        prs_com = app.Presentations.Open(target_path, WithWindow=False)
        try:
            slide = prs_com.Slides(slide_index)
            dup = slide.Duplicate()
            prs_com.Save()
            total = prs_com.Slides.Count
            new_idx = dup.SlideIndex
        finally:
            prs_com.Close()

        return {
            "status": "success",
            "message": f"Slide {slide_index} duplicated successfully.",
            "new_slide_index": new_idx,
            "total_slides": total
        }

    @classmethod
    def delete_slide(
        cls,
        slide_index: int,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Delete slide by 1-based index."""
        target_path = os.path.abspath(filepath) if filepath else None

        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    if slide_index < 1 or slide_index > prs_com.Slides.Count:
                        raise ValueError(f"Invalid slide index {slide_index}.")
                    prs_com.Slides(slide_index).Delete()
                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Slide {slide_index} deleted successfully in live PowerPoint.",
                        "total_slides": prs_com.Slides.Count
                    }

        # Pure Python headless fallback via python-pptx
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}. Total slides: {len(prs.slides)}.")

        sldId_lst = prs.slides._sldIdLst
        sldId = sldId_lst[slide_index - 1]
        try:
            rId = sldId.rId
            prs.part.drop_rel(rId)
        except Exception:
            pass
        sldId_lst.remove(sldId)
        save_pptx_safely(prs, target_path)

        return {
            "status": "success",
            "message": f"Slide {slide_index} deleted successfully.",
            "total_slides": len(prs.slides)
        }

    @classmethod
    def reorder_slide(
        cls,
        from_index: int,
        to_index: int,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Move a slide from from_index to to_index (1-based)."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    prs_com.Slides(from_index).MoveTo(to_index)
                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Slide moved from {from_index} to {to_index} successfully.",
                        "total_slides": prs_com.Slides.Count
                    }

        # Pure Python headless fallback via python-pptx
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if from_index < 1 or from_index > len(prs.slides):
            raise ValueError(f"Invalid from_index {from_index}. Total slides: {len(prs.slides)}.")
        if to_index < 1 or to_index > len(prs.slides):
            raise ValueError(f"Invalid to_index {to_index}. Total slides: {len(prs.slides)}.")

        sldId_lst = prs.slides._sldIdLst
        sld_elem = sldId_lst[from_index - 1]
        sldId_lst.remove(sld_elem)
        sldId_lst.insert(to_index - 1, sld_elem)
        save_pptx_safely(prs, target_path)

        return {
            "status": "success",
            "message": f"Slide moved from {from_index} to {to_index} successfully.",
            "total_slides": len(prs.slides)
        }

    @classmethod
    def set_slide_background(
        cls,
        color: str,
        slide_index: Optional[int] = None,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Set background color for a specific slide or all slides."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        # python-pptx
        if target_path and os.path.exists(target_path):
            prs = pptx.Presentation(target_path)
            target_slides = [prs.slides[slide_index - 1]] if (slide_index and slide_index <= len(prs.slides)) else prs.slides
            for s in target_slides:
                bg = s.background
                fill = bg.fill
                fill.solid()
                fill.fore_color.rgb = hex_to_rgb_color(color)
            save_pptx_safely(prs, target_path)

        # COM
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    bgr = hex_to_bgr_int(color)
                    if slide_index:
                        prs_com.Slides(slide_index).Background.Fill.Solid()
                        prs_com.Slides(slide_index).Background.Fill.ForeColor.RGB = bgr
                    else:
                        for s in prs_com.Slides:
                            s.Background.Fill.Solid()
                            s.Background.Fill.ForeColor.RGB = bgr
                    prs_com.Save()

        return {
            "status": "success",
            "message": f"Background color '{color}' applied successfully.",
            "slide_index": slide_index or "all"
        }

    @classmethod
    def set_slide_notes(
        cls,
        notes: str,
        slide_index: int,
        filepath: Optional[str] = None,
        append: bool = False
    ) -> Dict[str, Any]:
        """Add or update speaker notes for a specific slide."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        # COM
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    slide = prs_com.Slides(slide_index)
                    existing = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text if append else ""
                    new_text = f"{existing}\n{notes}".strip() if existing else notes
                    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = new_text
                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Speaker notes set for slide {slide_index} in live PowerPoint.",
                        "slide_index": slide_index
                    }

        # python-pptx
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        if append and tf.text:
            tf.text = f"{tf.text}\n{notes}".strip()
        else:
            tf.text = notes
        save_pptx_safely(prs, target_path)

        return {
            "status": "success",
            "message": f"Speaker notes set for slide {slide_index}.",
            "slide_index": slide_index
        }

    @classmethod
    def add_text_box(
        cls,
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
    ) -> Dict[str, Any]:
        """Add a rich formatted text box with typography and box styling."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        # COM
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    slide = prs_com.Slides(slide_index)
                    # Points in COM: 1 inch = 72 pt
                    shape = slide.Shapes.AddTextbox(
                        1,  # msoTextOrientationHorizontal
                        left_inches * 72,
                        top_inches * 72,
                        width_inches * 72,
                        height_inches * 72
                    )
                    tf = shape.TextFrame
                    tf.WordWrap = -1 if word_wrap else 0
                    tr = tf.TextRange
                    tr.Text = text
                    tr.Font.Name = font_name
                    tr.Font.Size = font_size
                    tr.Font.Bold = -1 if bold else 0
                    tr.Font.Italic = -1 if italic else 0
                    tr.Font.Underline = -1 if underline else 0
                    tr.Font.Color.RGB = hex_to_bgr_int(font_color)
                    tr.ParagraphFormat.Alignment = get_com_alignment(align)

                    if bg_color:
                        shape.Fill.Solid()
                        shape.Fill.ForeColor.RGB = hex_to_bgr_int(bg_color)
                    if border_color:
                        shape.Line.Visible = -1
                        shape.Line.ForeColor.RGB = hex_to_bgr_int(border_color)
                        shape.Line.Weight = border_width_pt

                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Text box added to slide {slide_index} in live PowerPoint.",
                        "slide_index": slide_index,
                        "shape_name": shape.Name
                    }

        # python-pptx
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        txBox = slide.shapes.add_textbox(Inches(left_inches), Inches(top_inches), Inches(width_inches), Inches(height_inches))
        tf = txBox.text_frame
        tf.word_wrap = word_wrap
        
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.italic = italic
        p.font.underline = underline
        p.font.color.rgb = hex_to_rgb_color(font_color)
        
        # Alignment & RTL
        if is_arabic_text(text) and align == "left":
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = get_pptx_alignment(align)

        # Fill & Border
        if bg_color:
            txBox.fill.solid()
            txBox.fill.fore_color.rgb = hex_to_rgb_color(bg_color)
        if border_color:
            txBox.line.color.rgb = hex_to_rgb_color(border_color)
            txBox.line.width = Pt(border_width_pt)

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Text box added to slide {slide_index}.",
            "slide_index": slide_index,
            "shape_name": txBox.name
        }

    @classmethod
    def add_bullet_points(
        cls,
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
    ) -> Dict[str, Any]:
        """Add structured bullet points with multi-level nesting and styling."""
        target_path = os.path.abspath(filepath) if filepath else None
        
        # python-pptx
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        txBox = slide.shapes.add_textbox(Inches(left_inches), Inches(top_inches), Inches(width_inches), Inches(height_inches))
        tf = txBox.text_frame
        tf.word_wrap = True

        for idx, item in enumerate(items):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            
            level = 0
            text_val = ""
            item_bold = False
            item_color = font_color
            
            if isinstance(item, dict):
                text_val = item.get("text", "")
                level = int(item.get("level", 0))
                item_bold = bool(item.get("bold", False))
                item_color = item.get("color", font_color)
            else:
                text_val = str(item)

            p.text = text_val
            p.level = level
            p.font.name = font_name
            p.font.size = Pt(font_size - (level * 2))
            p.font.bold = item_bold
            p.font.color.rgb = hex_to_rgb_color(item_color)
            p.line_spacing = line_spacing

            if is_arabic_text(text_val):
                p.alignment = PP_ALIGN.RIGHT

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Added {len(items)} bullet points to slide {slide_index}.",
            "slide_index": slide_index,
            "items_count": len(items)
        }

    @classmethod
    def add_shape(
        cls,
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
    ) -> Dict[str, Any]:
        """Add geometric or callout shapes with fill, border, and centered text."""
        target_path = os.path.abspath(filepath) if filepath else None
        st_clean = shape_type.strip().lower()

        # COM
        if COM_AVAILABLE:
            app = LivePowerPointSession.get_app(create_if_missing=False)
            if app:
                prs_com = LivePowerPointSession.find_presentation_by_path(target_path, app) if target_path else LivePowerPointSession.get_active_presentation(app)
                if prs_com:
                    slide = prs_com.Slides(slide_index)
                    com_type_id = COM_SHAPE_MAP.get(st_clean, 5) # Default rounded rect
                    shape = slide.Shapes.AddShape(
                        com_type_id,
                        left_inches * 72,
                        top_inches * 72,
                        width_inches * 72,
                        height_inches * 72
                    )
                    shape.Fill.Solid()
                    shape.Fill.ForeColor.RGB = hex_to_bgr_int(fill_color)
                    if border_color:
                        shape.Line.Visible = -1
                        shape.Line.ForeColor.RGB = hex_to_bgr_int(border_color)
                        shape.Line.Weight = border_width_pt
                    else:
                        shape.Line.Visible = 0

                    if text:
                        tr = shape.TextFrame.TextRange
                        tr.Text = text
                        tr.Font.Name = font_name
                        tr.Font.Size = font_size
                        tr.Font.Bold = -1 if bold else 0
                        tr.Font.Color.RGB = hex_to_bgr_int(text_color)
                        tr.ParagraphFormat.Alignment = 2 # Center

                    prs_com.Save()
                    return {
                        "status": "success",
                        "message": f"Shape '{shape_type}' added to slide {slide_index} in live PowerPoint.",
                        "slide_index": slide_index,
                        "shape_name": shape.Name
                    }

        # python-pptx
        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        mso_type = SHAPE_TYPE_MAP.get(st_clean, MSO_SHAPE.ROUNDED_RECTANGLE)
        shape = slide.shapes.add_shape(
            mso_type,
            Inches(left_inches),
            Inches(top_inches),
            Inches(width_inches),
            Inches(height_inches)
        )

        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_to_rgb_color(fill_color)

        if border_color:
            shape.line.color.rgb = hex_to_rgb_color(border_color)
            shape.line.width = Pt(border_width_pt)
        else:
            shape.line.fill.background()

        if text:
            tf = shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = text
            p.font.name = font_name
            p.font.size = Pt(font_size)
            p.font.bold = bold
            p.font.color.rgb = hex_to_rgb_color(text_color)
            p.alignment = PP_ALIGN.CENTER

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Shape '{shape_type}' added to slide {slide_index}.",
            "slide_index": slide_index,
            "shape_name": shape.name
        }

    @classmethod
    def add_card(
        cls,
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
    ) -> Dict[str, Any]:
        """Add modern executive KPI Metric Card with title, highlight value, description, and badge."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        # 1. Background Card Shape
        card_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left_inches),
            Inches(top_inches),
            Inches(width_inches),
            Inches(height_inches)
        )
        card_shape.fill.solid()
        card_shape.fill.fore_color.rgb = hex_to_rgb_color(card_bg_color)
        card_shape.line.color.rgb = hex_to_rgb_color(border_color)
        card_shape.line.width = Pt(1.5)

        # 2. Text Frame inside card
        tf = card_shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)
        tf.margin_bottom = Inches(0.2)

        # Title line
        p_title = tf.paragraphs[0]
        p_title.text = title.upper()
        p_title.font.name = font_name
        p_title.font.size = Pt(11)
        p_title.font.bold = True
        p_title.font.color.rgb = hex_to_rgb_color("#64748B")
        if is_arabic_text(title):
            p_title.alignment = PP_ALIGN.RIGHT

        # Value line (Big highlight)
        p_val = tf.add_paragraph()
        p_val.text = value
        p_val.font.name = font_name
        p_val.font.size = Pt(28)
        p_val.font.bold = True
        p_val.font.color.rgb = hex_to_rgb_color(primary_color)
        p_val.space_before = Pt(4)
        if is_arabic_text(value):
            p_val.alignment = PP_ALIGN.RIGHT

        # Description / Badge line
        if description or badge:
            p_desc = tf.add_paragraph()
            full_desc = f"{badge} • {description}" if (badge and description) else (badge or description)
            p_desc.text = full_desc
            p_desc.font.name = font_name
            p_desc.font.size = Pt(10)
            p_desc.font.color.rgb = hex_to_rgb_color(text_color)
            p_desc.space_before = Pt(4)
            if is_arabic_text(full_desc):
                p_desc.alignment = PP_ALIGN.RIGHT

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"KPI Card '{title}' added to slide {slide_index}.",
            "slide_index": slide_index,
            "title": title,
            "value": value
        }

    @classmethod
    def add_image(
        cls,
        image_path_or_url: str,
        slide_index: int = 1,
        left_inches: float = 1.0,
        top_inches: float = 1.5,
        width_inches: Optional[float] = 5.0,
        height_inches: Optional[float] = None,
        caption: Optional[str] = None,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert an image from local file path or URL with custom dimensions and optional caption."""
        target_path = os.path.abspath(filepath) if filepath else None
        temp_file_path = None
        
        try:
            local_img = image_path_or_url
            # Download if URL
            if image_path_or_url.startswith("http://") or image_path_or_url.startswith("https://"):
                parsed_url = urllib.parse.urlparse(image_path_or_url)
                ext = os.path.splitext(parsed_url.path)[1]
                if not ext or len(ext) > 5:
                    ext = ".png"

                req = urllib.request.Request(
                    image_path_or_url,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                    }
                )
                with urllib.request.urlopen(req, timeout=10) as response:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        tmp.write(response.read())
                        temp_file_path = tmp.name
                        local_img = temp_file_path

            if not os.path.exists(local_img):
                raise FileNotFoundError(f"Image file not found at: {local_img}")

            if not target_path or not os.path.exists(target_path):
                if COM_AVAILABLE:
                    app = LivePowerPointSession.get_app(create_if_missing=False)
                    if app:
                        prs_com = LivePowerPointSession.get_active_presentation(app)
                        if prs_com:
                            try:
                                target_path = prs_com.FullName
                            except Exception:
                                pass

            if not target_path or not os.path.exists(target_path):
                raise FileNotFoundError(f"Presentation file not found: {filepath}")

            prs = pptx.Presentation(target_path)
            if slide_index < 1 or slide_index > len(prs.slides):
                raise ValueError(f"Invalid slide index {slide_index}.")
            slide = prs.slides[slide_index - 1]

            left = Inches(left_inches)
            top = Inches(top_inches)
            w = Inches(width_inches) if width_inches else None
            h = Inches(height_inches) if height_inches else None

            pic = slide.shapes.add_picture(local_img, left, top, width=w, height=h)

            if caption:
                # Place caption right below the image
                cap_top = top + (pic.height or Inches(3.0)) + Inches(0.1)
                cap_box = slide.shapes.add_textbox(left, cap_top, pic.width or Inches(5.0), Inches(0.5))
                p = cap_box.text_frame.paragraphs[0]
                p.text = caption
                p.font.size = Pt(11)
                p.font.italic = True
                p.font.color.rgb = hex_to_rgb_color("#64748B")
                p.alignment = PP_ALIGN.CENTER

            save_pptx_safely(prs, target_path)
            return {
                "status": "success",
                "message": f"Image added to slide {slide_index} successfully.",
                "slide_index": slide_index,
                "shape_name": pic.name
            }
        finally:
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                except Exception:
                    pass

    @classmethod
    def add_table(
        cls,
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
    ) -> Dict[str, Any]:
        """Add professionally styled table with header fills, zebra striping, and cell formatting."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        rows = len(data) + 1
        cols = len(headers)

        table_shape = slide.shapes.add_table(
            rows,
            cols,
            Inches(left_inches),
            Inches(top_inches),
            Inches(width_inches),
            Inches(height_inches)
        )
        table = table_shape.table

        # Column widths
        if col_widths and len(col_widths) == cols:
            for c_idx, w in enumerate(col_widths):
                table.columns[c_idx].width = Inches(w)

        # Header formatting
        for col_idx, h_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = str(h_text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb_color(header_bg_color)
            for p in cell.text_frame.paragraphs:
                p.font.name = font_name
                p.font.size = Pt(font_size)
                p.font.bold = True
                p.font.color.rgb = hex_to_rgb_color(header_font_color)
                if is_arabic_text(str(h_text)):
                    p.alignment = PP_ALIGN.RIGHT
                else:
                    p.alignment = PP_ALIGN.CENTER

        # Data rows formatting with alternating colors
        for row_idx, row_vals in enumerate(data, start=1):
            is_alt = (row_idx % 2 == 0) and (alternate_row_color is not None)
            for col_idx in range(cols):
                val = row_vals[col_idx] if col_idx < len(row_vals) else ""
                cell = table.cell(row_idx, col_idx)
                cell.text = str(val)
                if is_alt:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = hex_to_rgb_color(alternate_row_color)
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = hex_to_rgb_color("#FFFFFF")

                for p in cell.text_frame.paragraphs:
                    p.font.name = font_name
                    p.font.size = Pt(font_size - 1)
                    p.font.color.rgb = hex_to_rgb_color("#1E293B")
                    if is_arabic_text(str(val)):
                        p.alignment = PP_ALIGN.RIGHT

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Table ({rows}x{cols}) added to slide {slide_index}.",
            "slide_index": slide_index,
            "rows": rows,
            "columns": cols
        }

    @classmethod
    def add_chart(
        cls,
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
    ) -> Dict[str, Any]:
        """Add native PowerPoint chart (Column, Bar, Line, Pie, Doughnut, Area) with series data."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        # Prepare chart data
        chart_data = CategoryChartData()
        chart_data.categories = categories

        for s in series_data:
            s_name = s.get("name", "Series")
            s_values = s.get("values", [])
            chart_data.add_series(s_name, tuple(s_values))

        ctype = CHART_TYPE_MAP.get(chart_type.strip().lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

        chart_shape = slide.shapes.add_chart(
            ctype,
            Inches(left_inches),
            Inches(top_inches),
            Inches(width_inches),
            Inches(height_inches),
            chart_data
        )

        chart = chart_shape.chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False

        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Chart '{chart_type}' added to slide {slide_index}.",
            "slide_index": slide_index,
            "chart_type": chart_type,
            "series_count": len(series_data)
        }

    @classmethod
    def add_timeline(
        cls,
        milestones: List[Dict[str, str]],
        slide_index: int = 1,
        title: Optional[str] = "Project Roadmap & Timeline",
        top_inches: float = 2.5,
        primary_color: str = "#1F497D",
        accent_color: str = "#418AB3",
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create visual roadmap / timeline slide with sequential milestone nodes and connectors."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        if slide_index < 1 or slide_index > len(prs.slides):
            raise ValueError(f"Invalid slide index {slide_index}.")
        slide = prs.slides[slide_index - 1]

        # Slide Title if specified
        if title and slide.shapes.title:
            slide.shapes.title.text = title

        count = len(milestones)
        if count == 0:
            return {"status": "error", "message": "No milestones provided."}

        # Calculate horizontal positions
        slide_w = prs.slide_width / Inches(1) # width in inches
        margin = 1.0
        available_w = slide_w - (margin * 2)
        step = available_w / max(1, count)
        card_w = min(2.4, step * 0.85)

        # Draw connecting line across
        line_top = top_inches + 1.2
        line_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(margin + 0.2),
            Inches(line_top),
            Inches(available_w - 0.4),
            Inches(0.06)
        )
        line_shape.fill.solid()
        line_shape.fill.fore_color.rgb = hex_to_rgb_color("#CBD5E1")
        line_shape.line.fill.background()

        for idx, m in enumerate(milestones):
            c_x = margin + (idx * step) + ((step - card_w) / 2)
            m_date = m.get("date", f"Phase {idx+1}")
            m_title = m.get("title", f"Milestone {idx+1}")
            m_desc = m.get("description", "")

            # Milestone Node Circle
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(c_x + (card_w / 2) - 0.2),
                Inches(line_top - 0.17),
                Inches(0.4),
                Inches(0.4)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = hex_to_rgb_color(primary_color)
            circle.line.color.rgb = hex_to_rgb_color("#FFFFFF")
            circle.line.width = Pt(2)
            circle.text_frame.text = str(idx + 1)
            circle.text_frame.paragraphs[0].font.size = Pt(10)
            circle.text_frame.paragraphs[0].font.bold = True
            circle.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

            # Milestone Card Box (Alternating top/bottom)
            is_top = (idx % 2 == 0)
            box_top = (top_inches - 1.2) if is_top else (line_top + 0.4)
            box = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(c_x),
                Inches(box_top),
                Inches(card_w),
                Inches(1.2)
            )
            box.fill.solid()
            box.fill.fore_color.rgb = hex_to_rgb_color("#FFFFFF")
            box.line.color.rgb = hex_to_rgb_color(accent_color)
            box.line.width = Pt(1.5)

            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.1)
            tf.margin_right = Inches(0.1)
            tf.margin_top = Inches(0.1)

            p0 = tf.paragraphs[0]
            p0.text = m_date.upper()
            p0.font.size = Pt(9)
            p0.font.bold = True
            p0.font.color.rgb = hex_to_rgb_color(accent_color)

            p1 = tf.add_paragraph()
            p1.text = m_title
            p1.font.size = Pt(11)
            p1.font.bold = True
            p1.font.color.rgb = hex_to_rgb_color("#1E293B")

            if m_desc:
                p2 = tf.add_paragraph()
                p2.text = m_desc
                p2.font.size = Pt(9)
                p2.font.color.rgb = hex_to_rgb_color("#64748B")

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Timeline with {count} milestones added to slide {slide_index}.",
            "slide_index": slide_index,
            "milestones_count": count
        }

    @classmethod
    def search_and_replace(
        cls,
        find_text: str,
        replace_text: str,
        match_case: bool = False,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Search and replace text across all slides, text boxes, tables, and notes."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        prs = pptx.Presentation(target_path)
        replacements_count = 0
        pattern = re.compile(re.escape(find_text), 0 if match_case else re.IGNORECASE)

        for slide in prs.slides:
            # Check shapes and textframes
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if pattern.search(run.text):
                                run.text = pattern.sub(replace_text, run.text)
                                replacements_count += 1
                
                # Check tables
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if pattern.search(run.text):
                                        run.text = pattern.sub(replace_text, run.text)
                                        replacements_count += 1

            # Check notes
            if slide.has_notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                for paragraph in notes_tf.paragraphs:
                    for run in paragraph.runs:
                        if pattern.search(run.text):
                            run.text = pattern.sub(replace_text, run.text)
                            replacements_count += 1

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Replaced {replacements_count} occurrence(s) of '{find_text}' with '{replace_text}'.",
            "replacements_count": replacements_count
        }

    @classmethod
    def fill_template(
        cls,
        template_path: str,
        output_path: str,
        replacements: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Fill a PowerPoint template by replacing {{tags}} with dynamic data."""
        in_path = os.path.abspath(template_path)
        out_path = os.path.abspath(output_path)
        if not os.path.exists(in_path):
            raise FileNotFoundError(f"Template presentation not found at: {in_path}")

        prs = pptx.Presentation(in_path)
        total_replacements = 0

        for key, val in replacements.items():
            tag = f"{{{{{key}}}}}"
            str_val = str(val)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            if tag in paragraph.text:
                                paragraph.text = paragraph.text.replace(tag, str_val)
                                total_replacements += 1
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    if tag in paragraph.text:
                                        paragraph.text = paragraph.text.replace(tag, str_val)
                                        total_replacements += 1

        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        save_pptx_safely(prs, out_path)
        return {
            "status": "success",
            "message": f"Template filled successfully with {total_replacements} replacements saved to '{out_path}'.",
            "output_path": out_path,
            "replacements_made": total_replacements
        }

    @classmethod
    def apply_theme(
        cls,
        theme: str,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Apply cohesive color palette and styling to entire presentation."""
        target_path = os.path.abspath(filepath) if filepath else None
        if not target_path or not os.path.exists(target_path):
            if COM_AVAILABLE:
                app = LivePowerPointSession.get_app(create_if_missing=False)
                if app:
                    prs_com = LivePowerPointSession.get_active_presentation(app)
                    if prs_com:
                        try:
                            target_path = prs_com.FullName
                        except Exception:
                            pass

        if not target_path or not os.path.exists(target_path):
            raise FileNotFoundError(f"Presentation file not found: {filepath}")

        palette = THEME_PALETTES.get(theme.lower(), THEME_PALETTES["modern_blue"])
        prs = pptx.Presentation(target_path)

        for slide in prs.slides:
            # Set background
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = hex_to_rgb_color(palette["bg_light"])

            # Style title
            if slide.shapes.title:
                for p in slide.shapes.title.text_frame.paragraphs:
                    p.font.name = palette["font_heading"]
                    p.font.color.rgb = hex_to_rgb_color(palette["primary"])

        save_pptx_safely(prs, target_path)
        return {
            "status": "success",
            "message": f"Theme '{theme}' applied to all {len(prs.slides)} slides.",
            "theme": theme,
            "palette": palette
        }

    @classmethod
    def execute_script(
        cls,
        code: str,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """Execute arbitrary Python script with access to presentation objects (prs, slide, pptx, app)."""
        target_path = os.path.abspath(filepath) if filepath else None
        prs = pptx.Presentation(target_path) if (target_path and os.path.exists(target_path)) else None
        app = LivePowerPointSession.get_app(create_if_missing=False) if COM_AVAILABLE else None

        stdout_buf = io.StringIO()
        stderr_buf = io.StringIO()
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        local_env = {
            "pptx": pptx,
            "prs": prs,
            "pptx_prs": prs,
            "app": app,
            "filepath": target_path,
            "Inches": Inches,
            "Pt": Pt,
            "RGBColor": RGBColor,
            "MSO_SHAPE": MSO_SHAPE,
            "PP_ALIGN": PP_ALIGN,
            "XL_CHART_TYPE": XL_CHART_TYPE,
            "json": json,
            "os": os,
            "sys": sys
        }

        import contextlib
        has_error = False
        try:
            with contextlib.redirect_stdout(stdout_buf), contextlib.redirect_stderr(stderr_buf):
                exec(code, globals(), local_env)
            if prs and target_path:
                save_pptx_safely(prs, target_path)
        except Exception:
            has_error = True
            stderr_buf.write(traceback.format_exc())

        stderr_val = stderr_buf.getvalue()
        is_error = has_error or (bool(stderr_val.strip()) and ("Traceback" in stderr_val or "Error" in stderr_val or "Exception" in stderr_val))

        return {
            "status": "error" if is_error else "success",
            "stdout": stdout_buf.getvalue(),
            "stderr": stderr_val
        }
