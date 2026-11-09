# Word MCP

<h3 align="center">Enterprise-Grade Document Generation, Live COM Automation & Headless Processing.</h3>

<p align="center">
Empower LLMs and AI agents to create, format, and manipulate Microsoft Word documents with full RTL/Arabic support and PDF export.
</p>

<img width="1200" height="628" alt="word" src="https://github.com/user-attachments/assets/6793f026-88a9-4c6c-9ce1-960cbe183a37" />


## Overview

**Microsoft Word MCP Server (Enterprise Edition)** is a robust, enterprise-grade Model Context Protocol (MCP) server providing 21 specialized tools (plus an opt-in 22nd, `word_execute_script`) for complete document generation, styling, and management.

It combines interactive live desktop control via Windows COM with cross-platform headless document processing via `python-docx` and `lxml`. AI assistants can construct formatted reports, edit tables, manage dynamic tables of contents (TOC), insert watermarks, and export print-ready PDFs through simple natural language prompts.

## 🏛️ Dual-Engine Architecture

This server is powered by a robust dual-engine design that automatically selects the optimal execution engine based on your operating system and environment:

```
                  ┌─────────────────────────────────────┐
                  │          AI Assistant / LLM         │
                  └──────────────────┬──────────────────┘
                                     │ (MCP Protocol)
                                     ▼
                  ┌─────────────────────────────────────┐
                  │           Word MCP Server           │
                  │        (21 + 1 opt-in Tools)        │
                  └──────────────────┬──────────────────┘
                                     │
                 ┌───────────────────┴───────────────────┐
                 ▼                                       ▼
    ┌───────────────────────────┐           ┌───────────────────────────┐
    │  Live Desktop COM Engine  │           │   Headless Fast Engine    │
    │  (pywin32 / Windows COM)  │           │  (python-docx + lxml)     │
    ├───────────────────────────┤           ├───────────────────────────┤
    │ • Real-time interactive   │           │ • Cross-platform          │
    │   document manipulation   │           │   (Windows, Linux, macOS) │
    │ • Native dynamic TOC &    │           │ • Ultra-fast document gen │
    │   field updates           │           │ • No Microsoft Office     │
    │ • True-fidelity PDF export│           │   installation required   │
    │ • Live review comments    │           │ • Templating & XML batch  │
    └───────────────────────────┘           └───────────────────────────┘
```

1. **Live Desktop COM Automation (Windows)**:
   - Connects directly to running Microsoft Word instances or launches Word on-demand via `pywin32`.
   - Supports active window editing, cursor movement, native Table of Contents (TOC) generation, background watermarks, review comments, and high-fidelity PDF export via Word's native print/export engine.

2. **Headless High-Performance Engine (Cross-Platform)**:
   - Built on top of `python-docx` and `lxml` for lightning-fast server-side processing without requiring Microsoft Office.
   - Ideal for CI/CD pipelines, Docker containers, cloud environments, batch document generation, and template rendering across Windows, macOS, and Linux.

## 🛠️ Complete Tool Catalog (21 Tools + 1 opt-in)

### 1. Core Document Management

| Tool Name | Description | Key Parameters |
| :--- | :--- | :--- |
| `word_status` | Check running state of Word, active documents, COM availability, and document paths. | *None* |
| `word_create_document` | Create a new document with an optional styled title banner, subtitle, and author metadata. | `filepath`, `title`, `subtitle`, `author`, `visible` |
| `word_open_document` | Open an existing document from disk for reading or editing. | `filepath`, `read_only`, `visible` |
| `word_read_document` | Read and convert complete document content into formatted Markdown, JSON, or plain text. | `filepath`, `output_format`, `include_tables`, `max_paragraphs` |
| `word_get_outline` | Extract the hierarchical heading outline, word count, paragraph count, and section stats. | `filepath` |
| `word_save_document` | Save the current active document or save to a new target path. | `save_as_path`, `filepath` |
| `word_export_pdf` | Export documents to high-fidelity PDF using Word's native rendering engine. | `output_pdf_path`, `filepath` |

### 2. Rich Content & Formatting

| Tool Name | Description | Key Parameters |
| :--- | :--- | :--- |
| `word_add_heading` | Insert styled headings (Levels 1–6) with customizable colors, alignment, and spacing. | `text`, `level`, `filepath`, `align`, `color` |
| `word_add_paragraph` | Add rich text paragraphs with precise control over font name, size, bold, italic, color, highlights, and line spacing. | `text`, `bold`, `italic`, `font_size`, `font_name`, `font_color`, `highlight_color`, `align` |
| `word_add_bullet_list` | Insert formatted bulleted or numbered lists. | `items`, `filepath`, `list_type` |
| `word_add_callout_box` | Insert modern highlighted callout boxes (`info`, `tip`, `warning`, `important`, `quote`). | `text`, `title`, `box_type`, `bg_color`, `border_color` |
| `word_add_table` | Insert professional tables with customizable header styling, column widths, and alternating row shading (zebra striping). | `headers`, `data`, `header_bg_color`, `header_font_color`, `alternate_row_color`, `col_widths` |
| `word_add_image` | Insert images with custom width/height dimensions, alignment, and figure captions. | `image_path`, `width_inches`, `height_inches`, `caption`, `align` |
| `word_add_page_break` | Insert a clean page break into the document. | `filepath` |
| `word_search_and_replace`| Smart find-and-replace across all paragraphs and tables with case matching. | `find_text`, `replace_text`, `match_case`, `replace_all` |

### 3. Enterprise & Layout Configuration

| Tool Name | Description | Key Parameters |
| :--- | :--- | :--- |
| `word_set_page_setup` | Configure paper sizes (`A4`, `Letter`, `Legal`), orientation (`portrait`, `landscape`), and margin presets (`normal`, `narrow`, `wide`). | `paper_size`, `orientation`, `margin_preset`, `top_margin_inches` |
| `word_set_header_footer`| Configure headers, footers, dynamic page numbering (`Page X`), and different first page options. | `header_text`, `footer_text`, `include_page_number`, `different_first_page` |
| `word_insert_toc` | Insert and update an automated native Table of Contents (TOC) with page numbers. | `title`, `upper_heading_level`, `lower_heading_level`, `position` |
| `word_add_watermark` | Add semi-transparent diagonal or horizontal text watermarks (e.g., `CONFIDENTIAL`, `DRAFT`). | `text`, `color`, `font_size`, `diagonal` |
| `word_add_comment` | Insert review and audit comments attached to specific text passages or paragraphs. | `comment_text`, `target_text`, `paragraph_index`, `author` |
| `word_fill_template` | Render template documents by replacing `{{placeholder}}` tags with dynamic values. | `template_path`, `output_path`, `replacements` |

### 4. Advanced Automation

| Tool Name | Description | Key Parameters |
| :--- | :--- | :--- |
| `word_execute_script` | **Disabled by default.** Executes arbitrary Python with direct access to COM and python-docx objects. Only registered when `WORD_MCP_ENABLE_SCRIPT=1`. | `code`, `filepath` |

## 🎛️ Live vs. Headless: how the server decides (read this first)

The server never edits a `.docx` behind the back of a Word window that has it open.

| `WORD_MCP_MODE` | Behaviour |
| :--- | :--- |
| `auto` *(default)* | If Word is running, **attach to that exact Word window** and edit live (you see every change). If Word is **not** running, edit the file directly with `python-docx`. |
| `live` | COM only. If Word is not running or cannot be reached, the tool fails with a clear message. |
| `headless` | `python-docx` only. COM is never used (except PDF export, which always needs Word). |

Safety rules built into the COM layer:

- **Attach, never duplicate.** The server attaches to the running Word. If a `WINWORD.EXE` process exists but COM cannot reach it (typically *Word elevated vs. server not elevated*, or a start-up dialog), the tool fails with an explanation instead of launching a second instance.
- **A live failure is an error, not a silent file edit.** Earlier versions swallowed COM errors and fell back to rewriting the file on disk, which is why changes did not appear (and could be overwritten by Word later).
- **Word is only quit if this server started it** (hidden instance used for PDF export, and only when it holds no documents). Your own Word window is never closed.
- **No invisible Word.** Word launched by a tool is always visible.
- **One COM thread.** All tool calls run serially on a single dedicated thread, which is what COM requires and prevents two agent calls from interleaving edits.
- **`DisplayAlerts` is restored** after every call; `~$` lock files are no longer deleted.
- Transient "call was rejected by callee" errors are retried while attaching/opening documents.

Other environment variables: `WORD_MCP_TIMEOUT` (seconds per tool call, default 120), `WORD_MCP_LOG` (`DEBUG`/`INFO`/`WARNING`, logs go to **stderr** because stdout is the MCP channel).

## 🔒 Security Notice (`word_execute_script`)

> [!WARNING]
> `word_execute_script` is **not exposed unless you set `WORD_MCP_ENABLE_SCRIPT=1`** in the server's environment. When enabled it allows arbitrary Python execution within the host environment.
>
> - **Trust Boundaries**: Only enable it for trusted AI clients and applications.
> - **Execution Context**: The script runs with the same permissions as the host Python process.
> - **Sandboxing**: In multi-tenant or untrusted environments, run the server inside an isolated container or virtual environment.

## 📦 Installation & Setup

### Prerequisites
- Python 3.10 or higher
- Microsoft Word (optional, required only for live desktop COM features and native PDF export on Windows)

### 1. Install from Source

```bash
git clone https://github.com/your-org/word-mcp-server.git
cd word-mcp-server

# Install dependencies
pip install -r requirements.txt

# Or install in editable mode
pip install -e .
```

### 2. Run the Server

```bash
# Run using console entrypoint
word-mcp

# Or run directly with Python
python word_mcp_server.py
```

## 🔌 Client Configurations

Add the Word MCP Server to your preferred AI agent environment using the snippets below:

### 1. Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": [
        "-m",
        "word_mcp_server"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

*Or with an absolute path:*

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": [
        "/path/to/word-mcp-server/word_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8",
        "PYTHONPATH": "/path/to/word-mcp-server"
      }
    }
  }
}
```

### 2. Cursor (`mcp.json`)

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "python",
      "args": [
        "/path/to/word-mcp-server/word_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

### 3. Google Antigravity / Gemini CLI (`mcp_config.json`)

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": [
        "/path/to/word-mcp-server/word_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8",
        "PYTHONPATH": "/path/to/word-mcp-server"
      }
    }
  }
}
```

### 4. VS Code (Cline / Roo Code / Continue)

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": [
        "/path/to/word-mcp-server/word_mcp_server.py"
      ],
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

## 🧪 Verification & Testing

Unit tests (any OS, use a fake Word, verify the live/headless safety rules):

```bash
pip install pytest
python -m pytest -q tests
```

End-to-end smoke test of all tools (on Windows with Word installed this also exercises live COM and PDF export; on other systems PDF export is skipped):

```bash
python test_mcp.py
```

The script exits with a non-zero code and lists every failing tool if anything fails.

> [!WARNING]
> Document generation and template variable substitution can modify files on disk. There is always a possibility of error, so we assume no responsibility for any inaccuracies.

## 🩺 Troubleshooting

| Symptom | Likely cause / fix |
| :--- | :--- |
| `Word is running but COM cannot attach to it` | Word and the server run at different privilege levels (one "as administrator"), or Word shows a start-up/recovery dialog. Run both normally and close the dialog. |
| `Word did not finish within 120s` | A modal dialog is open in Word (save prompt, password, repair). Close it, retry. Raise `WORD_MCP_TIMEOUT` for very large documents. |
| Tool says live edit failed and the file was not modified | By design. Word was reachable but rejected the operation. Make sure Word is idle and retry; use `word_status` to inspect state. |
| Changes not visible | Check `word_status`: in `headless` mode or when Word is closed, edits go to the file; open it afterwards. Use `WORD_MCP_MODE=live` to forbid file edits. |

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

