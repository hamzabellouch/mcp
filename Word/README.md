# 📝 Word MCP

<h3 align="center">Enterprise-Grade Document Generation, Live COM Automation & Headless Processing.</h3>

<p align="center">
Empower LLMs and AI agents to create, format, and manipulate Microsoft Word documents with full RTL/Arabic support and PDF export.
</p>

## Overview

**Microsoft Word MCP Server (Enterprise Edition)** is a robust, enterprise-grade Model Context Protocol (MCP) server providing 22 specialized tools for complete document generation, styling, and management.

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
                  │        (22 Specialized Tools)       │
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

## 🛠️ Complete Tool Catalog (22 Tools)

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
| `word_execute_script` | Execute arbitrary Python automation scripts with direct access to COM and python-docx objects. | `code`, `filepath` |

## 🔒 Security Notice (`word_execute_script`)

> [!WARNING]
> The `word_execute_script` tool allows arbitrary Python code execution within the host environment to enable complex automations and COM object manipulations.
> 
> - **Trust Boundaries**: Only connect this MCP server to trusted AI clients and applications.
> - **Execution Context**: The script runs with the same permissions as the host Python process.
> - **Sandboxing**: When deploying in multi-tenant or untrusted environments, execute the server inside an isolated container or virtual environment.

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

Run the included comprehensive 16-step test suite to verify that all 22 tools execute flawlessly:

```bash
python test_mcp.py
```

Expected output:
```text
=======================================================
   WORD MCP SERVER ENTERPRISE VERIFICATION SUITE       
=======================================================
[1/16] Testing word_status()...
[2/16] Testing word_create_document()...
...
=======================================================
   ALL 22 WORD MCP TOOLS VERIFIED & WORKING FLAWLESSLY! 
=======================================================
```

> [!WARNING]
> Document generation and template variable substitution can modify files on disk. There is always a possibility of error, so we assume no responsibility for any inaccuracies.

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

