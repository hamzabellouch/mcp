"""
Comprehensive Verification Test Suite for Word MCP Server (Enterprise Edition)
-------------------------------------------------------------------------------
Verifies all 22 tools across headless generation, template rendering (.docx & .dotx),
RTL/Arabic formatting, edge-case bug regressions, COM runtime safety rules,
mocked live COM paths, and live PDF export (when Word COM is reachable).
"""

import asyncio
import json
import os
import sys
import tempfile
import threading
import unittest
import zipfile
from unittest.mock import patch

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image

# Ensure UTF-8 output encoding on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# Route bytecode cache to Tests/.cache
sys.pycache_prefix = os.path.abspath(os.path.join(os.path.dirname(__file__), ".cache"))

# Ensure Word project directory is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Word"))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# The script tool is only registered when explicitly enabled; enable it for this run.
os.environ.setdefault("WORD_MCP_ENABLE_SCRIPT", "1")

import word_mcp_servers as srv  # noqa: E402
from word_mcp_servers import (  # noqa: E402
    ComWorker,
    LiveWordError,
    LiveWordSession,
    WordEngine,
    WordNotReachable,
    WordNotRunning,
    com_retry,
    get_mode,
    require_filepath,
)


# ----------------------------------------------------------------------
# Helpers & Fake Word COM Objects for Unit Tests
# ----------------------------------------------------------------------
class FakeFont:
    def __init__(self):
        self.Name = None
        self.Size = None
        self.Bold = False
        self.Italic = False
        self.Underline = 0
        self.Color = None

    def Reset(self):
        self.Bold = False
        self.Italic = False
        self.Underline = 0
        self.Color = None


class FakeParagraphFormat:
    def __init__(self):
        self.Alignment = 0
        self.ReadingOrder = 1
        self.SpaceBefore = 0
        self.SpaceAfter = 0
        self.LineSpacing = 12
        self.LineSpacingRule = 0

    def Reset(self):
        self.Alignment = 0
        self.SpaceBefore = 0
        self.SpaceAfter = 0
        self.LineSpacingRule = 0


class FakeFind:
    def __init__(self, range_obj):
        self.Parent = range_obj
        self.Text = ""
        self.Forward = True
        self.Wrap = 1
        self.MatchCase = False
        self.MatchWholeWord = False
        self.Replacement = type("Rep", (), {"Text": "", "ClearFormatting": lambda s: None})()

    def ClearFormatting(self):
        pass

    def Execute(self, Replace=0):
        parent = self.Parent
        haystack = parent._backing[parent.Start:parent.End] if parent._backing is not None else parent.Text
        if self.Text and self.Text in haystack:
            idx = haystack.index(self.Text)
            m_start = parent.Start + idx
            m_end = m_start + len(self.Text)
            if Replace in (1, 2):
                old = parent.Text
                if Replace == 2:
                    parent.Text = old.replace(self.Text, self.Replacement.Text)
                else:
                    parent.Text = old.replace(self.Text, self.Replacement.Text, 1)
                return True
            parent.Start = m_start
            parent.End = m_end
            return True
        return False


class FakeFields:
    def __init__(self):
        self.added = []

    def Add(self, Range=None, Type=None, PreserveFormatting=True):
        self.added.append({"type": Type, "preserve": PreserveFormatting})
        return type("Field", (), {})()


class FakeParagraph:
    def __init__(self, range_obj, outline_level=10):
        self.Range = range_obj
        self.Style = "Normal"
        self.OutlineLevel = outline_level


class FakeParagraphs(list):
    @property
    def Count(self):
        return len(self)

    @property
    def Last(self):
        return self[-1]

    def __call__(self, idx):
        return self[idx - 1]


class FakeRange:
    def __init__(self, text="", _is_para_child=False, _owner=None):
        self._text = text
        self._owner = _owner
        self.Font = FakeFont()
        self.ParagraphFormat = FakeParagraphFormat()
        self.Fields = FakeFields()
        self.Find = FakeFind(self)
        self.ListFormat = type("ListFormat", (), {"ListType": 0})()
        self.HighlightColorIndex = 0
        self.Start = 0
        self.End = len(text)
        self.Style = -1
        self.NextStoryRange = None
        self.Words = []
        if not _is_para_child:
            self.Paragraphs = FakeParagraphs([FakeParagraph(self)])

    @property
    def _backing(self):
        return self._owner._text if self._owner is not None else None

    @property
    def Text(self):
        if self._owner is not None:
            return self._owner._text[self.Start:self.End]
        return self._text

    @Text.setter
    def Text(self, val):
        if self._owner is not None:
            cur = self._owner._text
            self._owner._text = cur[:self.Start] + val + cur[self.End:]
            self._owner.End = len(self._owner._text)
            self.End = self.Start + len(val)
        else:
            self._text = val
            self.End = len(val)

    @property
    def Duplicate(self):
        dup = FakeRange(self.Text, _is_para_child=True, _owner=self._owner or self)
        dup.Start = self.Start
        dup.End = self.End
        return dup

    def Delete(self):
        self.Text = ""

    def SetRange(self, s, e):
        self.Start = s
        self.End = e

    def Collapse(self, direction=0):
        if direction == 0:
            self.Start = self.End
        else:
            self.End = self.Start

    def InsertBefore(self, s):
        self.Text = s + self.Text

    def InsertAfter(self, s):
        self.Text = self.Text + s

    def Information(self, kind):
        return False


class FakePageSetup:
    def __init__(self):
        self.PaperSize = 2
        self.Orientation = 0
        self.TopMargin = 72
        self.BottomMargin = 72
        self.LeftMargin = 72
        self.RightMargin = 72
        self.DifferentFirstPageHeaderFooter = False


class FakeHeaderFooter:
    def __init__(self, text=""):
        self.Range = FakeRange(text)
        self.Shapes = type("Shapes", (), {"Count": 0})()


class FakeSection:
    def __init__(self):
        self.PageSetup = FakePageSetup()
        self._hdr = FakeHeaderFooter("Existing Header\r")
        self._ftr = FakeHeaderFooter("\r")

    def Headers(self, idx):
        return self._hdr

    def Footers(self, idx):
        return self._ftr


class FakeSections(list):
    @property
    def Count(self):
        return len(self)

    def __call__(self, idx):
        return self[idx - 1]


class FakeSelection:
    def __init__(self, doc):
        self._doc = doc
        self.Font = FakeFont()
        self.ParagraphFormat = FakeParagraphFormat()
        self.Range = FakeRange()
        self.Style = -1
        self.Start = 0
        self.End = 0

    def HomeKey(self, Unit=6):
        self.Start = 0
        self.End = 0

    def EndKey(self, Unit=6):
        self.Start = len(self._doc.Content.Text)
        self.End = self.Start

    def TypeParagraph(self):
        self._doc.Content.Text += "\r"

    def TypeText(self, s):
        self._doc.Content.Text += s

    def Information(self, kind):
        return False


class FakeDoc:
    def __init__(self, path, has_path=True):
        self.FullName = path if has_path else os.path.basename(path)
        self.Name = os.path.basename(path)
        self.Path = os.path.dirname(path) if has_path else ""
        self.activated = 0
        self.saved = 0
        self.closed = False
        self.Content = FakeRange("Body hello\r")
        self.Paragraphs = FakeParagraphs([FakeParagraph(FakeRange("Body hello\r", _is_para_child=True, _owner=self.Content))])
        self.Sections = FakeSections([FakeSection()])
        self.Tables = FakeParagraphs([])
        self.StoryRanges = [self.Content, FakeRange("Header hello\r")]
        self.Comments = type("Comments", (), {"Count": 0, "added": [], "Add": lambda s, **kw: s.added.append(kw)})()
        self.Hyperlinks = type("Hyperlinks", (), {"added": [], "Add": lambda s, **kw: s.added.append(kw)})()
        self.exported = []
        self.created_ranges = []
        self.TablesOfContents = type(
            "TOCs",
            (),
            {"Add": lambda s, **kw: type("TOC", (), {"Update": lambda t: None})()},
        )()

    def Activate(self):
        self.activated += 1

    def Save(self):
        self.saved += 1

    def SaveAs2(self, p):
        self.FullName = p
        self.Path = os.path.dirname(p)
        self.saved += 1

    def Close(self, SaveChanges=False):
        self.closed = True

    def Range(self, s, e):
        r = FakeRange("", _is_para_child=True, _owner=self.Content)
        r.Start = s
        r.End = e
        self.created_ranges.append(r)
        return r

    def ExportAsFixedFormat(self, **kwargs):
        self.exported.append(kwargs)


class FakeDocs(list):
    @property
    def Count(self):
        return len(self)

    def Open(self, path, **kwargs):
        doc = FakeDoc(path)
        self.append(doc)
        return doc

    def Add(self):
        doc = FakeDoc("Document1", has_path=False)
        self.append(doc)
        return doc


class FakeApp:
    def __init__(self, docs=()):
        self.Documents = FakeDocs(docs)
        self.Visible = True
        self.DisplayAlerts = -1
        self.UserName = "Default User"
        self.quit_called = False
        self.quit_args = None

    @property
    def Selection(self):
        return FakeSelection(self.ActiveDocument)

    def Quit(self, SaveChanges=0):
        self.quit_called = True
        self.quit_args = SaveChanges

    def Activate(self):
        pass

    @property
    def ActiveDocument(self):
        return self.Documents[-1]


# ----------------------------------------------------------------------
# 1. Unit & Mocked-COM Safety Tests
# ----------------------------------------------------------------------
class TestWordRuntimeAndSafety(unittest.TestCase):
    def setUp(self):
        self._orig_mode = os.environ.get("WORD_MCP_MODE")
        self._orig_script = os.environ.get("WORD_MCP_ENABLE_SCRIPT")
        os.environ.pop("WORD_MCP_MODE", None)
        LiveWordSession._saved_alerts = None
        LiveWordSession._last_launched_by_us = False

    def tearDown(self):
        if self._orig_mode is None:
            os.environ.pop("WORD_MCP_MODE", None)
        else:
            os.environ["WORD_MCP_MODE"] = self._orig_mode
        if self._orig_script is None:
            os.environ.pop("WORD_MCP_ENABLE_SCRIPT", None)
        else:
            os.environ["WORD_MCP_ENABLE_SCRIPT"] = self._orig_script
        LiveWordSession._saved_alerts = None
        LiveWordSession._last_launched_by_us = False

    def test_fastmcp_tools_registered(self):
        async def _list():
            return await srv.mcp.list_tools()

        tools = asyncio.run(_list())
        names = {t.name for t in tools}
        expected = {
            "word_status",
            "word_create_document",
            "word_open_document",
            "word_read_document",
            "word_get_outline",
            "word_add_heading",
            "word_add_paragraph",
            "word_add_bullet_list",
            "word_add_callout_box",
            "word_add_table",
            "word_add_image",
            "word_add_page_break",
            "word_search_and_replace",
            "word_set_header_footer",
            "word_insert_toc",
            "word_fill_template",
            "word_add_comment",
            "word_set_page_setup",
            "word_add_watermark",
            "word_export_pdf",
            "word_save_document",
            "word_modify_paragraph",
            "word_clear_document",
            "word_execute_script",
        }
        self.assertEqual(expected - names, set())

    def test_mode_defaults_and_validation(self):
        self.assertEqual(get_mode(), "auto")
        os.environ["WORD_MCP_MODE"] = "LIVE"
        self.assertEqual(get_mode(), "live")
        os.environ["WORD_MCP_MODE"] = "invalid_mode"
        self.assertEqual(get_mode(), "auto")

    def test_headless_mode_never_uses_live(self):
        os.environ["WORD_MCP_MODE"] = "headless"
        self.assertFalse(LiveWordSession.should_use_live())

    def test_live_mode_without_pywin32_raises(self):
        os.environ["WORD_MCP_MODE"] = "live"
        with patch.object(srv, "COM_AVAILABLE", False):
            with self.assertRaises(LiveWordError):
                LiveWordSession.should_use_live()

    def test_handle_live_failure_policy(self):
        LiveWordSession.handle_live_failure(WordNotRunning("not running"))
        os.environ["WORD_MCP_MODE"] = "live"
        with self.assertRaises(LiveWordError):
            LiveWordSession.handle_live_failure(WordNotRunning("not running"))
        os.environ["WORD_MCP_MODE"] = "auto"
        with self.assertRaises(WordNotReachable):
            LiveWordSession.handle_live_failure(WordNotReachable("unreachable"))
        with self.assertRaises(LiveWordError) as ctx:
            LiveWordSession.handle_live_failure(RuntimeError("COM error"))
        self.assertIn("NOT modified on disk", str(ctx.exception))

    def test_live_failure_never_touches_file_on_disk(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "report.docx")
            docx.Document().save(target)
            with open(target, "rb") as f:
                before = f.read()

            class BrokenDoc(FakeDoc):
                def Range(self, s, e):
                    raise RuntimeError("COM exploded")

            app = FakeApp([BrokenDoc(target)])
            with patch.object(srv, "COM_AVAILABLE", True), \
                 patch.object(LiveWordSession, "_attach", staticmethod(lambda: app)), \
                 patch.object(srv, "word_process_exists", lambda: True):
                with self.assertRaises(LiveWordError):
                    WordEngine.add_paragraph(text="hello", filepath=target)

            with open(target, "rb") as f:
                self.assertEqual(f.read(), before)

    def test_attach_never_launches_duplicate_word(self):
        dispatched = []
        fake_win32 = type("W", (), {"client": type("C", (), {"Dispatch": staticmethod(lambda n: dispatched.append(n))})})
        with patch.object(srv, "COM_AVAILABLE", True), \
             patch.object(LiveWordSession, "_attach", staticmethod(lambda: None)), \
             patch.object(srv, "word_process_exists", lambda: True), \
             patch.object(srv, "win32com", fake_win32):
            with self.assertRaises(WordNotReachable):
                LiveWordSession.get_app(create_if_missing=True)
        self.assertEqual(dispatched, [])

    def test_not_running_no_create_raises_word_not_running(self):
        with patch.object(srv, "COM_AVAILABLE", True), \
             patch.object(LiveWordSession, "_attach", staticmethod(lambda: None)), \
             patch.object(srv, "word_process_exists", lambda: False):
            with self.assertRaises(WordNotRunning):
                LiveWordSession.get_app(create_if_missing=False)

    def test_launch_visibility_and_display_alerts_restore(self):
        launched = FakeApp()
        launched.Visible = False
        fake_win32 = type("W", (), {"client": type("C", (), {"Dispatch": staticmethod(lambda n: launched)})})
        with patch.object(srv, "COM_AVAILABLE", True), \
             patch.object(LiveWordSession, "_attach", staticmethod(lambda: None)), \
             patch.object(srv, "word_process_exists", lambda: False), \
             patch.object(srv, "win32com", fake_win32):
            app = LiveWordSession.get_app(create_if_missing=True, visible=False)
            self.assertTrue(app.Visible)
            self.assertTrue(LiveWordSession.launched_by_us())

            launched.Visible = True
            app = LiveWordSession.get_app(create_if_missing=True, visible=False, hidden_ok=True)
            self.assertFalse(app.Visible)

        attached = FakeApp()
        attached.DisplayAlerts = -1
        with patch.object(srv, "COM_AVAILABLE", True), \
             patch.object(LiveWordSession, "_attach", staticmethod(lambda: attached)):
            LiveWordSession.get_app()
            self.assertFalse(LiveWordSession.launched_by_us())
            self.assertEqual(attached.DisplayAlerts, 0)
            LiveWordSession.restore_ui()
            self.assertEqual(attached.DisplayAlerts, -1)

    def test_find_open_document_case_and_cloud_urls(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = os.path.join(tmp, "Report.docx")
            cloud = FakeDoc("https://contoso.sharepoint.com/sites/doc.docx", has_path=False)
            cloud.FullName = "https://contoso.sharepoint.com/sites/doc.docx"
            local = FakeDoc(p)
            app = FakeApp([cloud, local])
            self.assertIs(LiveWordSession.find_open_document(app, p.lower()), local)
            self.assertIsNone(LiveWordSession.find_open_document(app, os.path.join(tmp, "other.docx")))

    def test_require_filepath_and_execute_script_safety(self):
        with self.assertRaises(ValueError):
            require_filepath(None)
        os.environ.pop("WORD_MCP_ENABLE_SCRIPT", None)
        with self.assertRaises(PermissionError):
            WordEngine.execute_script(code="print(1)")

        os.environ["WORD_MCP_ENABLE_SCRIPT"] = "1"
        os.environ["WORD_MCP_MODE"] = "headless"
        real_stdout = sys.stdout
        res_ok = WordEngine.execute_script(code="print('ok')")
        self.assertTrue(res_ok["success"])
        self.assertIs(sys.stdout, real_stdout)
        res_exit = WordEngine.execute_script(code="import sys; print('before'); sys.exit(2)")
        self.assertFalse(res_exit["success"])
        self.assertIn("before", res_exit["output"])
        self.assertIs(sys.stdout, real_stdout)

    def test_worker_and_com_retry(self):
        w = ComWorker(default_timeout=5)
        try:
            ids = {w.submit(threading.get_ident) for _ in range(4)}
            self.assertEqual(len(ids), 1)
            self.assertNotIn(threading.get_ident(), ids)
            with self.assertRaises(KeyError):
                w.submit(lambda: {}["missing"])
            self.assertEqual(w.submit(lambda: 42), 42)
        finally:
            w.shutdown()

        w2 = ComWorker(default_timeout=0.15)
        try:
            import time
            with self.assertRaises(TimeoutError) as ctx:
                w2.submit(time.sleep, 0.8)
            self.assertIn("modal dialog", str(ctx.exception))
        finally:
            w2.shutdown()

        class Transient(Exception):
            pass

        calls = {"n": 0}
        with patch.object(srv, "is_transient_com_error", lambda e: isinstance(e, Transient)), \
             patch.object(srv.time, "sleep", lambda s: None):
            def flaky():
                calls["n"] += 1
                if calls["n"] < 3:
                    raise Transient()
                return "ok"

            self.assertEqual(com_retry(flaky), "ok")
            self.assertEqual(calls["n"], 3)

    def test_mocked_live_com_fixes(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc_path = os.path.join(tmp, "live.docx")
            docx.Document().save(doc_path)
            fdoc = FakeDoc(doc_path)
            app = FakeApp([fdoc])

            with patch.object(srv, "COM_AVAILABLE", True), \
                 patch.object(LiveWordSession, "_attach", staticmethod(lambda: app)), \
                 patch.object(srv, "word_process_exists", lambda: True):
                # 1. Letter paper size must be wdPaperLetter = 2, and 0.0 margins must apply
                WordEngine.set_page_setup(
                    filepath=doc_path, paper_size="letter", orientation="landscape", top_margin_inches=0.0
                )
                ps = fdoc.Sections(1).PageSetup
                self.assertEqual(ps.PaperSize, 2)
                self.assertEqual(ps.Orientation, 1)
                self.assertEqual(ps.TopMargin, 0.0)

                # 2. Search and replace across StoryRanges (body + header)
                res_sr = WordEngine.search_and_replace(
                    find_text="hello", replace_text="world", filepath=doc_path, replace_all=True
                )
                self.assertTrue(res_sr["success"])
                self.assertIn("world", fdoc.StoryRanges[1].Text)

                # 3. Inline PAGE field in footer
                WordEngine.set_header_footer(
                    header_text="Hdr", footer_text="Ftr", include_page_number=True, filepath=doc_path
                )
                self.assertEqual(fdoc.Sections(1).Footers(1).Range.Fields.added[0]["type"], 33)

                # 4. Untitled save guard prevents Save As modal hang
                untitled = FakeDoc("Document2", has_path=False)
                app.Documents.append(untitled)
                with self.assertRaises(LiveWordError):
                    WordEngine.save_document()

                # 5. export_pdf quits only when launched_by_us()
                pdf_out = os.path.join(tmp, "out.pdf")
                res_pdf = WordEngine.export_pdf(output_pdf_path=pdf_out, filepath=doc_path)
                self.assertEqual(res_pdf["status"], "exported_pdf")
                self.assertFalse(app.quit_called)

                # 6. Empty table raises ValueError cleanly
                with self.assertRaises(ValueError):
                    WordEngine.add_table(headers=None, data=[], filepath=doc_path)

                # 7. Empty find_text raises ValueError cleanly
                with self.assertRaises(ValueError):
                    WordEngine.search_and_replace(find_text="", replace_text="x", filepath=doc_path)

                # 8. add_comment with out-of-range paragraph_index=0 falls back cleanly without crashing
                cmt = WordEngine.add_comment(comment_text="Note", paragraph_index=0, filepath=doc_path)
                self.assertEqual(cmt["status"], "comment_added_live")
                self.assertEqual(len(fdoc.Comments.added), 1)

                # 9. insert_toc at start works cleanly in COM
                toc = WordEngine.insert_toc(filepath=doc_path, title="TOC", position="start")
                self.assertEqual(toc["status"], "toc_inserted_live")

                # 10. add_paragraph uses Range (does not touch Selection) and keeps RTL ReadingOrder=0 and Alignment=2
                sel_sentinel = object()
                with patch.object(FakeApp, "Selection", new_callable=lambda: property(lambda s: sel_sentinel)):
                    fdoc.created_ranges.clear()
                    WordEngine.add_paragraph(text="مرحبا بالعالم", filepath=doc_path)
                    rtl_rng = [r for r in fdoc.created_ranges if r.ParagraphFormat.ReadingOrder == 0]
                    self.assertTrue(rtl_rng, "Expected Range-based insertion with ReadingOrder=0")
                    self.assertEqual(rtl_rng[-1].ParagraphFormat.Alignment, 2)
                    self.assertIn("مرحبا بالعالم", fdoc.Content.Text)

                # 11. search_and_replace handles strings longer than Word Find's 255-char limit in COM
                long_prefix = "LONG_TOKEN_" + ("X" * 280)
                fdoc.Content.Text = f"Start {long_prefix} End\r"
                res_long = WordEngine.search_and_replace(
                    find_text=long_prefix, replace_text="SHORT", filepath=doc_path
                )
                self.assertTrue(res_long["success"])
                self.assertEqual(fdoc.Content.Text, "Start SHORT End\r")

                # 12. modify_paragraph updates RTL/alignment and delete works in COM
                fdoc.Paragraphs[0].Range = FakeRange("مرحبا\r", _is_para_child=True, _owner=fdoc.Content)
                fdoc.Content.Text = "مرحبا\r"
                mod_res = WordEngine.modify_paragraph(paragraph_index=1, rtl=True, bold=True, filepath=doc_path)
                self.assertEqual(mod_res["status"], "paragraph_updated_live")
                self.assertEqual(fdoc.Paragraphs[0].Range.ParagraphFormat.ReadingOrder, 0)
                self.assertEqual(fdoc.Paragraphs[0].Range.ParagraphFormat.Alignment, 2)

                # 13. clear_document wipes body content in COM
                clr_res = WordEngine.clear_document(filepath=doc_path)
                self.assertEqual(clr_res["status"], "cleared_live")
                self.assertEqual(fdoc.Content.Text, "")

                # 14. _read_live_com_fallback extracts uniform run formatting
                styled_rng = FakeRange("Styled Bold\r", _is_para_child=True)
                styled_rng.Font.Bold = -1
                styled_rng.Font.Italic = 0
                styled_rng.Font.Size = 14
                styled_rng.Font.Color = 0x7D491F  # BGR for #1F497D
                fdoc.Paragraphs[0] = FakeParagraph(styled_rng)
                live_read = WordEngine._read_live_com_fallback(fdoc, "json", False, None)
                self.assertTrue(live_read["paragraphs"][0]["runs"][0]["bold"])
                self.assertEqual(live_read["paragraphs"][0]["runs"][0]["color"], "#1F497D")

                # 15. Out-of-bounds table merge raises ValueError
                with self.assertRaises(ValueError):
                    WordEngine.add_table(headers=["A", "B"], data=[["1", "2"]], merges=[[0, 0, 5, 1]], filepath=doc_path)


# ----------------------------------------------------------------------
# 2. Full Headless + Live PDF End-to-End Verification Suite (All 22 Tools)
# ----------------------------------------------------------------------
TEST_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEST_FILE = os.path.join(TEST_DIR, "temp_enterprise_output.docx")
TEST_SAVED_AS = os.path.join(TEST_DIR, "temp_saved_as.docx")
TEST_TEMPLATE = os.path.join(TEST_DIR, "temp_template.docx")
TEST_DOTX = os.path.join(TEST_DIR, "temp_template.dotx")
TEST_FILLED = os.path.join(TEST_DIR, "temp_filled.docx")
TEST_FILLED_DOTX = os.path.join(TEST_DIR, "temp_filled_dotx.docx")
TEST_IMAGE = os.path.join(TEST_DIR, "temp_word_img.png")
TEST_PDF = os.path.join(TEST_DIR, "temp_enterprise_output.pdf")


def _make_dotx(src_docx: str, dst_dotx: str) -> None:
    with zipfile.ZipFile(src_docx, "r") as zin, zipfile.ZipFile(dst_dotx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"wordprocessingml.document.main+xml",
                    b"wordprocessingml.template.main+xml",
                )
            zout.writestr(item, data)


def _cleanup_temp_files():
    for f in [TEST_FILE, TEST_SAVED_AS, TEST_TEMPLATE, TEST_DOTX, TEST_FILLED, TEST_FILLED_DOTX, TEST_IMAGE, TEST_PDF]:
        if os.path.exists(f):
            try:
                os.remove(f)
            except Exception:
                pass


def run_e2e_tests() -> bool:
    print("=======================================================")
    print("   WORD MCP SERVER ENTERPRISE VERIFICATION SUITE       ")
    print("=======================================================")

    failures = []

    def check_tool(name, fn, *args, **kwargs):
        out = fn(*args, **kwargs)
        if isinstance(out, str) and out.startswith("❌"):
            failures.append((name, out))
        return out

    Image.new("RGB", (320, 180), color=(31, 73, 125)).save(TEST_IMAGE)

    orig_mode = os.environ.get("WORD_MCP_MODE")
    try:
        # 1. Status Check (real mode)
        print("\n[1/22] Testing word_status()...")
        status_out = srv.word_status()
        print("Status Result:\n", status_out)
        status_dict = WordEngine.get_status()

        # Run document generation tools in deterministic headless mode so an open
        # desktop Word window on the user's machine is never disturbed.
        os.environ["WORD_MCP_MODE"] = "headless"

        # 2. Create Document
        print("\n[2/22] Testing word_create_document()...")
        create_out = check_tool(
            "word_create_document",
            srv.word_create_document,
            filepath=TEST_FILE,
            title="تقرير استراتيجية الذكاء الاصطناعي 2026",
            subtitle="وثيقة هندسية وفنية معتمدة | سري للغاية",
            author="فريق الذكاء الاصطناعي المتقدم",
            visible=False,
        )
        print("Create Output:\n", create_out)

        # 3. Open Document
        print("\n[3/22] Testing word_open_document()...")
        open_out = check_tool("word_open_document", srv.word_open_document, filepath=TEST_FILE, visible=False)
        print("Open Output:\n", open_out)

        # 4. Page Setup (test landscape+A4 ordering & zero margin, then switch back to portrait A4)
        print("\n[4/22] Testing word_set_page_setup()...")
        check_tool(
            "word_set_page_setup(landscape)",
            srv.word_set_page_setup,
            filepath=TEST_FILE,
            paper_size="A4",
            orientation="landscape",
            top_margin_inches=0.0,
        )
        d_ps = docx.Document(TEST_FILE)
        assert d_ps.sections[0].page_width > d_ps.sections[0].page_height, "A4 landscape width must exceed height"
        assert d_ps.sections[0].top_margin == 0, "Zero top margin must be preserved"

        ps_out = check_tool(
            "word_set_page_setup",
            srv.word_set_page_setup,
            filepath=TEST_FILE,
            paper_size="A4",
            margin_preset="normal",
            orientation="portrait",
        )
        print("Page Setup Output:\n", ps_out)

        # 5. Header & Footer (verify repeated calls don't duplicate PAGE fields)
        print("\n[5/22] Testing word_set_header_footer()...")
        check_tool(
            "word_set_header_footer(1)",
            srv.word_set_header_footer,
            filepath=TEST_FILE,
            header_text="Initial Header",
            footer_text="Initial Footer",
            include_page_number=True,
        )
        hf_out = check_tool(
            "word_set_header_footer",
            srv.word_set_header_footer,
            filepath=TEST_FILE,
            header_text="شركة حمزة للتقنية | وثيقة استراتيجية",
            footer_text="Generated by Antigravity Word MCP",
            include_page_number=True,
            header_align="right",
            footer_align="center",
        )
        print("Header/Footer Output:\n", hf_out)
        d_hf = docx.Document(TEST_FILE)
        fld_count = len(d_hf.sections[0].footer.paragraphs[0]._p.findall(qn("w:fldSimple")))
        assert fld_count == 1, f"Expected 1 PAGE field after repeated call, got {fld_count}"

        # 6. Watermark (test XML special characters & preserve header alignment)
        print("\n[6/22] Testing word_add_watermark()...")
        wm_out = check_tool(
            "word_add_watermark",
            srv.word_add_watermark,
            filepath=TEST_FILE,
            text='CONFIDENTIAL & "DRAFT"',
            color="#E0E0E0",
        )
        print("Watermark Output:\n", wm_out)
        d_wm = docx.Document(TEST_FILE)
        assert d_wm.sections[0].header.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
        assert d_wm.sections[0].header._element.xpath('.//*[@id="WordWatermarkObject"]')

        # 7. Headings & Rich Paragraphs
        print("\n[7/22] Testing word_add_heading() & word_add_paragraph()...")
        check_tool(
            "word_add_heading",
            srv.word_add_heading,
            text="1. الملخص التنفيذي والرؤية",
            level=1,
            filepath=TEST_FILE,
            color="#1F497D",
        )
        d_hd = docx.Document(TEST_FILE)
        last_heading = d_hd.paragraphs[-1]
        assert last_heading.alignment == WD_ALIGN_PARAGRAPH.RIGHT, "Arabic heading must auto-align RIGHT"
        assert last_heading._p.pPr.find(qn("w:bidi")) is not None, "Arabic heading must set w:bidi"

        check_tool(
            "word_add_paragraph",
            srv.word_add_paragraph,
            text="تستعرض هذه الوثيقة البنية التحتية المتطورة لبروتوكول سياق النماذج (MCP) وتكامله مع تطبيقات مايكروسوفت أوفيس.",
            filepath=TEST_FILE,
            font_size=11,
            font_name="Segoe UI",
            line_spacing=1.15,
            highlight_color="YELLOW",
        )

        # 8. Comment
        print("\n[8/22] Testing word_add_comment()...")
        cmt_out = check_tool(
            "word_add_comment",
            srv.word_add_comment,
            filepath=TEST_FILE,
            comment_text="تم تدقيق هذه الفقرة واعتمادها رسمياً للربع الأول.",
            target_text="الملخص التنفيذي",
            author="مدقق الذكاء الاصطناعي",
        )
        print("Add Comment Output:\n", cmt_out)

        # 9. Callout Box
        print("\n[9/22] Testing word_add_callout_box()...")
        check_tool(
            "word_add_callout_box",
            srv.word_add_callout_box,
            text="يدعم النظام المعالجة الخلفية الفائقة السرعة عبر python-docx والتحكم اللحظي التفاعلي عبر COM Automation.",
            filepath=TEST_FILE,
            title="ميزة معمارية رئيسية",
            box_type="tip",
        )

        # 10. Bullet List
        print("\n[10/22] Testing word_add_bullet_list()...")
        check_tool(
            "word_add_heading(2)",
            srv.word_add_heading,
            text="2. القدرات التشغيلية والميزات",
            level=1,
            filepath=TEST_FILE,
            color="#1F497D",
        )
        check_tool(
            "word_add_bullet_list",
            srv.word_add_bullet_list,
            items=[
                "توليد مستندات معقدة بتنسيق احترافي وهندسي متقن",
                "إدراج الجداول مع التظليل المتناوب والترويسات المميزة",
                "دعم القوالب الجاهزة وتعبئة النماذج الديناميكية",
                "التصدير إلى PDF بأعلى درجات الدقة والوضوح",
            ],
            filepath=TEST_FILE,
            list_type="bullet",
        )

        # 11. Table
        print("\n[11/22] Testing word_add_table()...")
        check_tool(
            "word_add_heading(3)",
            srv.word_add_heading,
            text="3. مصفوفة مقارنة الخوادم",
            level=2,
            filepath=TEST_FILE,
            color="#2B2B2B",
        )
        check_tool(
            "word_add_table",
            srv.word_add_table,
            headers=["الخادم (MCP)", "المحرك الأساسي", "حالة الاتصال", "زمن الاستجابة"],
            data=[
                ["Word MCP Server", "Live COM & python-docx", "فعال ومحدث (22 أداة)", "< 40ms"],
                ["Excel MCP Server", "Live COM & openpyxl", "فعال ومحدث (13 أداة)", "< 40ms"],
                ["Blender MCP Server", "Bridge Socket API", "فعال ومحدث (7 أدوات)", "< 100ms"],
            ],
            filepath=TEST_FILE,
            header_bg_color="#1F497D",
            header_font_color="#FFFFFF",
            alternate_row_color="#F2F4F8",
            col_widths=[1.8, 1.8, 1.8, 1.1],
        )

        # 12. Image (verify aligned inside same paragraph)
        print("\n[12/22] Testing word_add_image()...")
        img_out = check_tool(
            "word_add_image",
            srv.word_add_image,
            image_path=TEST_IMAGE,
            filepath=TEST_FILE,
            width_inches=3.0,
            caption="مخطط معماري تجريبي",
            align="center",
        )
        print("Add Image Output:\n", img_out)
        d_img = docx.Document(TEST_FILE)
        pic_paras = [p for p in d_img.paragraphs if p._p.xpath(".//w:drawing")]
        assert len(pic_paras) == 1 and pic_paras[0].alignment == WD_ALIGN_PARAGRAPH.CENTER

        # 13. Page Break
        print("\n[13/22] Testing word_add_page_break()...")
        pb_out = check_tool("word_add_page_break", srv.word_add_page_break, filepath=TEST_FILE)
        print("Page Break Output:\n", pb_out)

        # 14. Search & Replace (header text + replace_all=False + cross-run formatting)
        print("\n[14/22] Testing word_search_and_replace()...")
        sr_out = json.loads(
            check_tool(
                "word_search_and_replace",
                srv.word_search_and_replace,
                filepath=TEST_FILE,
                find_text="وثيقة استراتيجية",
                replace_text="وثيقة استراتيجية معتمدة",
            )
        )
        assert sr_out["replacements_count"] == 1

        # Cross-run formatting preservation & replace_all=False check
        d_sr = docx.Document()
        p_sr = d_sr.add_paragraph()
        r1 = p_sr.add_run("Bold ")
        r1.bold = True
        r2 = p_sr.add_run("TOK")
        r2.bold = True
        r3 = p_sr.add_run("EN end TOKEN")
        d_sr.add_paragraph("TOKEN in para 2")
        sr_path = TEST_SAVED_AS
        d_sr.save(sr_path)
        sr_one = WordEngine.search_and_replace(
            find_text="TOKEN", replace_text="C:\\100", filepath=sr_path, replace_all=False
        )
        assert sr_one["replacements_count"] == 1
        d_sr2 = docx.Document(sr_path)
        assert d_sr2.paragraphs[0].text == "Bold C:\\100 end TOKEN"
        assert d_sr2.paragraphs[0].runs[0].bold is True
        assert d_sr2.paragraphs[1].text == "TOKEN in para 2"

        # 15. Insert TOC
        print("\n[15/22] Testing word_insert_toc()...")
        toc_out = check_tool(
            "word_insert_toc",
            srv.word_insert_toc,
            filepath=TEST_FILE,
            title="فهرس المحتويات (TOC)",
            upper_heading_level=1,
            lower_heading_level=5,
        )
        print("TOC Output:\n", toc_out)
        d_toc = docx.Document(TEST_FILE)
        assert d_toc.settings._element.find(qn("w:updateFields")) is not None

        # 16. Fill Template (.docx AND .dotx, including cross-run & non-chaining)
        print("\n[16/22] Testing word_fill_template()...")
        check_tool(
            "word_create_document(template)",
            srv.word_create_document,
            filepath=TEST_TEMPLATE,
            title="فاتورة رقم {{invoice_no}}",
            subtitle="العميل: {{client_name}} | التاريخ: {{date}}",
            visible=False,
        )
        # Add split-run placeholder to verify formatting-preserving cross-run fill
        d_tpl = docx.Document(TEST_TEMPLATE)
        p_split = d_tpl.add_paragraph()
        r_lbl = p_split.add_run("المبلغ المستحق: ")
        r_lbl.bold = True
        p_split.add_run("{{total_")
        p_split.add_run("amount}} مع خصم {{discount}}.")
        d_tpl.save(TEST_TEMPLATE)

        fill_out = json.loads(
            check_tool(
                "word_fill_template",
                srv.word_fill_template,
                template_path=TEST_TEMPLATE,
                output_path=TEST_FILLED,
                replacements={
                    "invoice_no": "INV-2026-089",
                    "client_name": "شركة حمزة القابضة",
                    "date": "2026-09-29",
                    "total_amount": "$12,500",
                    "discount": "10%",
                },
            )
        )
        print("Fill Template Output:\n", json.dumps(fill_out, indent=2, ensure_ascii=False))
        assert fill_out["replacements_applied"] == 5
        d_filled = docx.Document(TEST_FILLED)
        assert "$12,500" in d_filled.paragraphs[-1].text
        assert d_filled.paragraphs[-1].runs[0].bold is True

        _make_dotx(TEST_TEMPLATE, TEST_DOTX)
        fill_dotx = json.loads(
            check_tool(
                "word_fill_template(.dotx)",
                srv.word_fill_template,
                template_path=TEST_DOTX,
                output_path=TEST_FILLED_DOTX,
                replacements={
                    "invoice_no": "DOTX-1",
                    "client_name": "Client",
                    "date": "2026-11-25",
                    "total_amount": "$99",
                    "discount": "5%",
                },
            )
        )
        assert fill_dotx["replacements_applied"] == 5
        assert "DOTX-1" in docx.Document(TEST_FILLED_DOTX).paragraphs[0].text

        # 17-18. Read Document & Outline
        print("\n[17-18/22] Testing word_read_document() & word_get_outline()...")
        read_res = check_tool("word_read_document", srv.word_read_document, filepath=TEST_FILE, output_format="markdown")
        print("Markdown Preview:\n", read_res[:400] + "...\n")
        outline = check_tool("word_get_outline", srv.word_get_outline, filepath=TEST_FILE)
        outline_data = json.loads(outline)
        print("Outline Items Count:", outline_data.get("outline_count"))
        assert outline_data.get("outline_count", 0) >= 4

        # 20. Save Document
        print("\n[20/24] Testing word_save_document()...")
        save_out = check_tool("word_save_document", srv.word_save_document, filepath=TEST_FILE, save_as_path=TEST_SAVED_AS)
        print("Save Document Output:\n", save_out)
        assert os.path.exists(TEST_SAVED_AS)

        # 21. Rich Inline Formatting, Hyperlinks, Positional Insertion, Table Merges & Modify/Clear
        print("\n[21-22/24] Testing rich_formatting, merges, word_modify_paragraph() & word_clear_document()...")
        check_tool(
            "word_add_paragraph(rich)",
            srv.word_add_paragraph,
            text="Text with **bold**, *italic*, `code`, and [Link](https://example.com).",
            filepath=TEST_SAVED_AS,
            rich_formatting=True,
            position="start",
        )
        d_rich = docx.Document(TEST_SAVED_AS)
        assert d_rich.paragraphs[0]._p.xpath(".//w:hyperlink"), "Expected clickable OOXML hyperlink"
        assert any(r.bold and r.text == "bold" for r in d_rich.paragraphs[0].runs)

        check_tool(
            "word_add_table(merges)",
            srv.word_add_table,
            headers=["Group", "Group", "Value"],
            data=[["A", "B", "10"]],
            merges=[[0, 0, 0, 1]],
            filepath=TEST_SAVED_AS,
        )
        d_mrg = docx.Document(TEST_SAVED_AS)
        last_tbl = d_mrg.tables[-1]
        assert last_tbl.cell(0, 0)._tc is last_tbl.cell(0, 1)._tc

        mod_out = json.loads(
            check_tool(
                "word_modify_paragraph",
                srv.word_modify_paragraph,
                paragraph_index=1,
                new_text="نص عربي محدث",
                rtl=True,
                bold=True,
                filepath=TEST_SAVED_AS,
            )
        )
        assert mod_out["status"] == "paragraph_updated"
        d_mod = docx.Document(TEST_SAVED_AS)
        assert d_mod.paragraphs[0].text == "نص عربي محدث"
        assert d_mod.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
        assert d_mod.paragraphs[0]._p.pPr.find(qn("w:bidi")) is not None

        clr_out = json.loads(
            check_tool(
                "word_clear_document",
                srv.word_clear_document,
                filepath=TEST_SAVED_AS,
                include_headers_footers=True,
            )
        )
        assert clr_out["status"] == "cleared"
        d_clr = docx.Document(TEST_SAVED_AS)
        assert sum(len(p.text.strip()) for p in d_clr.paragraphs) == 0
        assert len(d_clr.tables) == 0

        # 23. Execute Script
        print("\n[23/24] Testing word_execute_script()...")
        script_code = """
doc = docx_doc or docx.Document(filepath)
print(f"Document inspection complete: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables.")
"""
        exec_out = check_tool("word_execute_script", srv.word_execute_script, code=script_code, filepath=TEST_FILE)
        print("Script Output:\n", exec_out)

        # 24. Export PDF via Word COM (live if reachable or Word not running)
        print("\n[24/24] Testing word_export_pdf()...")
        if sys.platform != "win32":
            print("PDF Note: skipped on non-Windows OS.")
        elif status_dict.get("word_process_found") and not status_dict.get("word_running"):
            print(
                "PDF Note: skipped live export because an existing WINWORD.EXE process is unreachable "
                "from this session (elevation/desktop isolation). Mocked COM export verified above."
            )
        else:
            pdf_out = check_tool("word_export_pdf", srv.word_export_pdf, output_pdf_path=TEST_PDF, filepath=TEST_FILE)
            print("PDF Export Output:\n", pdf_out)
            if os.path.exists(TEST_PDF):
                print(f"✅ Generated PDF verified ({os.path.getsize(TEST_PDF)} bytes).")
    finally:
        if orig_mode is None:
            os.environ.pop("WORD_MCP_MODE", None)
        else:
            os.environ["WORD_MCP_MODE"] = orig_mode
        _cleanup_temp_files()

    print("\n=======================================================")
    if failures:
        print(f"   {len(failures)} TOOL CALL(S) FAILED:")
        for name, out in failures:
            print(f"   - {name}: {out[:160]}")
        print("=======================================================")
        return False
    print("   ALL 24 WORD MCP E2E CHECKS SUCCEEDED!")
    print("=======================================================")
    return True


def run_tests() -> int:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(TestWordRuntimeAndSafety)
    unit_res = unittest.TextTestRunner(verbosity=2).run(suite)
    e2e_ok = run_e2e_tests()
    return 0 if (unit_res.wasSuccessful() and e2e_ok) else 1


if __name__ == "__main__":
    sys.exit(run_tests())
