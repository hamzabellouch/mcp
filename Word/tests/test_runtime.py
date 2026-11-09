"""
Platform-independent tests for the live/headless decision logic.

They use a fake Word application so they run on Linux/macOS CI as well as Windows.
What they protect against (the original bugs):
  * a live COM failure silently turning into a file edit behind an open Word window
  * Quit() on a Word instance the user is working in
  * hidden Word instances
  * DisplayAlerts left disabled
  * stdout (the MCP channel) left redirected by word_execute_script
"""

import os
import sys
import threading

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import com_runtime  # noqa: E402
import word_engine  # noqa: E402
from word_engine import LiveWordSession, WordEngine, require_filepath  # noqa: E402
from com_runtime import ComWorker, LiveWordError, WordNotRunning, WordNotReachable  # noqa: E402


# ---------------------------------------------------------------- fakes
class FakeDoc:
    def __init__(self, path):
        self.FullName = path
        self.Name = os.path.basename(path)
        self.activated = 0

    def Activate(self):
        self.activated += 1

    def Save(self):
        pass


class FakeDocs(list):
    @property
    def Count(self):
        return len(self)


class FakeApp:
    def __init__(self, docs=()):
        self.Documents = FakeDocs(docs)
        self.Visible = True
        self.DisplayAlerts = -1
        self.quit_called = False

    def Quit(self):
        self.quit_called = True

    def Activate(self):
        pass

    @property
    def ActiveDocument(self):
        return self.Documents[-1]


@pytest.fixture(autouse=True)
def clean_env(monkeypatch):
    monkeypatch.delenv("WORD_MCP_MODE", raising=False)
    monkeypatch.delenv("WORD_MCP_ENABLE_SCRIPT", raising=False)
    LiveWordSession._saved_alerts = None
    LiveWordSession._last_launched_by_us = False
    yield


@pytest.fixture
def fake_com(monkeypatch):
    """Pretend pywin32 exists, with a controllable Word."""
    state = {"app": None, "process": False}
    monkeypatch.setattr(word_engine, "COM_AVAILABLE", True)
    monkeypatch.setattr(com_runtime, "COM_AVAILABLE", True)
    monkeypatch.setattr(LiveWordSession, "_attach", staticmethod(lambda: state["app"]))
    monkeypatch.setattr(word_engine, "word_process_exists", lambda: state["process"])
    return state


# ---------------------------------------------------------------- mode
def test_mode_defaults_and_validation(monkeypatch):
    assert com_runtime.get_mode() == "auto"
    monkeypatch.setenv("WORD_MCP_MODE", "LIVE")
    assert com_runtime.get_mode() == "live"
    monkeypatch.setenv("WORD_MCP_MODE", "nonsense")
    assert com_runtime.get_mode() == "auto"


def test_headless_mode_never_uses_live(monkeypatch, fake_com):
    monkeypatch.setenv("WORD_MCP_MODE", "headless")
    assert LiveWordSession.should_use_live() is False


def test_live_mode_without_pywin32_is_an_error(monkeypatch):
    monkeypatch.setattr(word_engine, "COM_AVAILABLE", False)
    monkeypatch.setenv("WORD_MCP_MODE", "live")
    with pytest.raises(LiveWordError):
        LiveWordSession.should_use_live()


# ---------------------------------------------------------------- failure policy (the core bug)
def test_word_not_running_falls_back_in_auto_mode():
    LiveWordSession.handle_live_failure(WordNotRunning("not running"))  # must not raise


def test_word_not_running_is_error_in_live_mode(monkeypatch):
    monkeypatch.setenv("WORD_MCP_MODE", "live")
    with pytest.raises(LiveWordError):
        LiveWordSession.handle_live_failure(WordNotRunning("x"))


def test_unexpected_com_failure_never_falls_back_to_file_editing():
    with pytest.raises(LiveWordError) as ei:
        LiveWordSession.handle_live_failure(RuntimeError("call was rejected by callee"))
    assert "NOT modified on disk" in str(ei.value)


def test_unreachable_word_is_never_benign():
    with pytest.raises(WordNotReachable):
        LiveWordSession.handle_live_failure(WordNotReachable("elevated"))


def test_tool_raises_instead_of_editing_file_when_live_fails(tmp_path, monkeypatch, fake_com):
    """Word is up, the doc is open, COM errors out mid-operation: the file must stay untouched."""
    target = tmp_path / "report.docx"
    import docx

    docx.Document().save(str(target))
    before = target.read_bytes()

    class BrokenApp(FakeApp):
        @property
        def Selection(self):
            raise RuntimeError("COM exploded")

    app = BrokenApp([FakeDoc(str(target))])
    fake_com["app"] = app
    fake_com["process"] = True
    monkeypatch.setattr(LiveWordSession, "get_app", classmethod(lambda cls, **k: app))

    with pytest.raises(LiveWordError):
        WordEngine.add_paragraph(text="hello", filepath=str(target))
    assert target.read_bytes() == before


# ---------------------------------------------------------------- attach / launch rules
def test_attach_never_launches_a_second_word(fake_com, monkeypatch):
    fake_com["app"] = None
    fake_com["process"] = True  # WINWORD.EXE exists but COM cannot reach it
    dispatched = []
    monkeypatch.setattr(word_engine, "win32com", type("W", (), {"client": type("C", (), {"Dispatch": staticmethod(lambda n: dispatched.append(n))})}))
    with pytest.raises(WordNotReachable):
        LiveWordSession.get_app(create_if_missing=True)
    assert dispatched == []


def test_not_running_and_no_create_raises_word_not_running(fake_com):
    with pytest.raises(WordNotRunning):
        LiveWordSession.get_app(create_if_missing=False)


def test_launch_is_always_visible_unless_hidden_ok(fake_com, monkeypatch):
    launched = FakeApp()
    launched.Visible = False
    client = type("C", (), {"Dispatch": staticmethod(lambda n: launched)})
    monkeypatch.setattr(word_engine, "win32com", type("W", (), {"client": client}))

    app = LiveWordSession.get_app(create_if_missing=True, visible=False)  # hidden_ok False
    assert app.Visible is True
    assert LiveWordSession.launched_by_us() is True

    launched.Visible = True
    app = LiveWordSession.get_app(create_if_missing=True, visible=False, hidden_ok=True)
    assert app.Visible is False


def test_attach_resets_launched_flag(fake_com):
    LiveWordSession._last_launched_by_us = True  # stale from an earlier call
    fake_com["app"] = FakeApp()
    LiveWordSession.get_app()
    assert LiveWordSession.launched_by_us() is False


def test_display_alerts_restored(fake_com):
    app = FakeApp()
    app.DisplayAlerts = -1
    fake_com["app"] = app
    LiveWordSession.get_app()
    assert app.DisplayAlerts == 0
    LiveWordSession.restore_ui()
    assert app.DisplayAlerts == -1


def test_find_open_document_is_case_insensitive(tmp_path):
    p = str(tmp_path / "A.docx")
    app = FakeApp([FakeDoc(p)])
    assert LiveWordSession.find_open_document(app, p.lower()) is app.Documents[0]
    assert LiveWordSession.find_open_document(app, str(tmp_path / "other.docx")) is None


# ---------------------------------------------------------------- headless path safety
def test_require_filepath_rejects_missing_path():
    with pytest.raises(ValueError):
        require_filepath(None)


def test_no_document_docx_is_created_in_cwd(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("WORD_MCP_MODE", "headless")
    with pytest.raises(ValueError):
        WordEngine.add_paragraph(text="x")
    assert not (tmp_path / "document.docx").exists()


def test_headless_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("WORD_MCP_MODE", "headless")
    f = str(tmp_path / "t.docx")
    WordEngine.create_document(filepath=f, title="Title", visible=False)
    res = WordEngine.add_paragraph(text="Hello", filepath=f)
    assert res["mode"] == "headless_docx"
    out = WordEngine.read_document(filepath=f, output_format="text")
    assert "Hello" in out["text"]


# ---------------------------------------------------------------- execute_script
def test_execute_script_disabled_by_default():
    with pytest.raises(PermissionError):
        WordEngine.execute_script(code="print(1)")


def test_execute_script_restores_stdout(monkeypatch, tmp_path):
    monkeypatch.setenv("WORD_MCP_ENABLE_SCRIPT", "1")
    monkeypatch.setenv("WORD_MCP_MODE", "headless")
    real = sys.stdout
    res = WordEngine.execute_script(code="print('hi')")
    assert res["success"] and "hi" in res["output"]
    assert sys.stdout is real
    res = WordEngine.execute_script(code="raise ValueError('boom')")
    assert not res["success"]
    assert sys.stdout is real


# ---------------------------------------------------------------- worker thread
def test_worker_runs_jobs_on_one_thread_serially():
    w = ComWorker(default_timeout=5)
    ids = {w.submit(threading.get_ident) for _ in range(5)}
    assert len(ids) == 1
    assert threading.get_ident() not in ids
    w.shutdown()


def test_worker_propagates_exceptions():
    w = ComWorker(default_timeout=5)
    with pytest.raises(KeyError):
        w.submit(lambda: {}["x"])
    assert w.submit(lambda: 7) == 7  # still alive afterwards
    w.shutdown()


def test_worker_timeout_message():
    import time

    w = ComWorker(default_timeout=0.2)
    with pytest.raises(TimeoutError) as ei:
        w.submit(time.sleep, 1.0)
    assert "modal dialog" in str(ei.value)
    w.shutdown()


def test_com_retry_only_retries_transient(monkeypatch):
    class Transient(Exception):
        pass

    monkeypatch.setattr(com_runtime, "is_transient_com_error", lambda e: isinstance(e, Transient))
    monkeypatch.setattr(com_runtime.time, "sleep", lambda s: None)
    calls = {"n": 0}

    def flaky():
        calls["n"] += 1
        if calls["n"] < 3:
            raise Transient()
        return "ok"

    assert com_runtime.com_retry(flaky) == "ok" and calls["n"] == 3

    def hard():
        raise ValueError("nope")

    with pytest.raises(ValueError):
        com_runtime.com_retry(hard)
