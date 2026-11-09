# Changelog

## 1.1.0

### Fixed
- **Changes not visible in the open Word window.** Every tool used to wrap its COM code in
  `except Exception: pass` and then silently edit the `.docx` on disk with python-docx. Live failures
  are now raised as `LiveWordError`; the file is never rewritten behind an open Word window.
- **Word closing unexpectedly.** `export_pdf` called `Quit()` whenever `GetActiveObject` failed to see Word
  (e.g. elevation mismatch), which could close the user's own Word. The server now attaches only, detects an
  existing `WINWORD.EXE` process, refuses to launch a second instance, and quits only an instance it started
  itself that holds no documents.
- Hidden Word instances: launched Word is always visible (a hidden instance swallows documents opened later).
- `DisplayAlerts` was disabled permanently; it is now restored after each tool call.
- `word_execute_script` could leave `sys.stdout` redirected (the MCP channel); now uses `contextlib.redirect_stdout`.
- Removed deletion of `~$*.docx` lock files, which belong to documents currently open in Word.
- No more silent creation of `document.docx` in the working directory when no `filepath` is given.
- `read_document` of the active document no longer leaves temp copies and has a live-text fallback.
- `test_mcp.py` no longer prints "all tools verified" when tools failed.

### Added
- `com_runtime.py`: single dedicated COM thread, transient-error retry, process detection, execution modes.
- `WORD_MCP_MODE` (`auto` | `live` | `headless`), `WORD_MCP_TIMEOUT`, `WORD_MCP_LOG`.
- `word_execute_script` is opt-in via `WORD_MCP_ENABLE_SCRIPT=1`.
- `word_status` reports mode, process/COM reachability and hints.
- `tests/test_runtime.py`: platform-independent tests with a fake Word.

### Known limitations
- Several live tools still type at the user's `Selection` (moves the cursor). Migrating them to `Range` is the next step.
- Transient COM rejections are retried when attaching/opening documents, not on every single COM call.
- Live COM paths could only be verified with a simulated Word here; run `python test_mcp.py` on Windows with Word.
