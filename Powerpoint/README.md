# 📊 PowerPoint MCP

<h3 align="center">High-Performance Presentation Automation, Dynamic Slide Generation & KPI Visualization.</h3>

<p align="center">
Create executive slide decks, native charts, KPI metric cards, timelines, and high-resolution PDF/PNG exports directly via natural language.
</p>

## Overview

**Microsoft PowerPoint MCP Server (Enterprise Edition)** is a comprehensive Model Context Protocol (MCP) server providing 26 specialized tools for presentation automation, slide generation, and multi-format document management.

Featuring an intelligent dual-engine architecture, it bridges AI assistants (Claude Desktop, Cursor, Gemini, VS Code) with live PowerPoint desktop instances on Windows, while offering 100% cross-platform headless presentation creation via `python-pptx` on macOS, Linux, and server environments.

## ⚡ Architecture: Dual-Engine Design

This MCP server features an intelligent dual-engine architecture that seamlessly selects the optimal execution engine based on the environment and request:

```
                          ┌──────────────────────────┐
                          │   AI Model / Assistant   │
                          │ (Claude, Cursor, Gemini) │
                          └─────────────┬────────────┘
                                        │ JSON-RPC (stdio)
                          ┌─────────────▼────────────┐
                          │  PowerPoint MCP Server   │
                          │   (26 Specialized Tools) │
                          └─────────────┬────────────┘
                                        │
                 ┌──────────────────────┴──────────────────────┐
                 │                                             │
      ┌──────────▼──────────┐                       ┌──────────▼──────────┐
      │  Windows COM Engine │                       │ Headless python-pptx│
      │   (pywin32 / Live)  │                       │   (Cross-Platform)  │
      ├─────────────────────┤                       ├─────────────────────┤
      │ • Interactive UI    │                       │ • Linux/macOS/Win   │
      │ • Live slide sync   │                       │ • Zero Office req   │
      │ • Native PDF export │                       │ • Fast batch jobs   │
      │ • High-res images   │                       │ • Template filling  │
      └─────────────────────┘                       └─────────────────────┘
```

1. **Live Desktop COM Engine (`pywin32`)**:
   - Directly manipulates the running Microsoft PowerPoint desktop application in real time.
   - Provides native vector PDF rendering and slide-to-image exports at up to 4K resolution.
   - Safely coordinates file locks, auto-saving and reloading when transitioning between headless and COM modes.

2. **Headless Python Engine (`python-pptx`)**:
   - Operates cross-platform (Windows, macOS, Linux) without requiring Microsoft PowerPoint or Office installation.
   - Ideal for CI/CD pipelines, containerized environments, cloud backends, and headless batch document generation.
   - Direct OpenXML manipulation with safe slide relationships and packaging lifecycle management.

## 🛠️ Complete Tools Reference (26 Specialized Tools)

### 1. Presentation Lifecycle & Document Management
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `powerpoint_status` | Check live status of PowerPoint application, version, and active presentations. | _None_ |
| `powerpoint_create_presentation` | Create a new presentation with 16:9 widescreen or 4:3 format, title slide, and theme palette. | `filepath`, `title`, `subtitle`, `author`, `widescreen`, `theme`, `visible` |
| `powerpoint_open_presentation` | Open an existing PowerPoint presentation file (`.pptx`, `.ppt`, `.potx`). | `filepath`, `read_only`, `visible` |
| `powerpoint_read_presentation` | Read full presentation structure, slide titles, body text, tables, and notes into Markdown or JSON. | `filepath`, `output_format`, `include_notes` |
| `powerpoint_get_outline` | Extract high-level hierarchical outline, slide titles, shape counts, and word statistics. | `filepath` |
| `powerpoint_save_presentation` | Save active or target presentation in place or as a new file copy. | `filepath`, `save_as_path` |
| `powerpoint_export_pdf` | Export presentation to high-quality PDF using native PowerPoint COM engine. | `output_pdf_path`, `filepath` |
| `powerpoint_export_slide_images` | Export specific slide or all slides as high-resolution PNG or JPG image files. | `output_dir`, `filepath`, `slide_index`, `image_format`, `width`, `height` |

### 2. Slide Structure & Navigation
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `powerpoint_add_slide` | Add a new slide with layout preset, title, subtitle, and custom background color. | `filepath`, `layout`, `title`, `subtitle`, `background_color`, `slide_index` |
| `powerpoint_duplicate_slide` | Duplicate an existing slide at a 1-based index with all shapes and styles intact. | `slide_index`, `filepath` |
| `powerpoint_delete_slide` | Delete a slide by 1-based index with proper OpenXML relationship cleanup. | `slide_index`, `filepath` |
| `powerpoint_reorder_slide` | Move a slide from one position index to another (1-based). | `from_index`, `to_index`, `filepath` |
| `powerpoint_set_slide_background` | Set solid hex background color for a specific slide or all slides. | `color`, `slide_index`, `filepath` |
| `powerpoint_set_slide_notes` | Set or append speaker notes for a specific slide. | `notes`, `slide_index`, `filepath`, `append` |

### 3. Visual Content, Data & Typography
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `powerpoint_add_text_box` | Add a rich formatted text box with custom font, size, colors, borders, and alignment. | `text`, `slide_index`, `left_inches`, `top_inches`, `width_inches`, `height_inches`, `font_size`, `font_name`, `font_color`, `bg_color`, `border_color` |
| `powerpoint_add_bullet_points` | Add multi-level nested bullet points with customized colors, boldness, and line spacing. | `items`, `slide_index`, `left_inches`, `top_inches`, `width_inches`, `font_size`, `font_color`, `line_spacing` |
| `powerpoint_add_shape` | Add geometric shapes (rectangles, diamonds, arrows, chevrons, stars, callouts) with fill and text. | `shape_type`, `slide_index`, `left_inches`, `top_inches`, `width_inches`, `height_inches`, `text`, `fill_color`, `border_color` |
| `powerpoint_add_card` | Add modern executive KPI / Metric Cards with title, large stat highlight, description, and badge. | `title`, `value`, `description`, `badge`, `slide_index`, `left_inches`, `top_inches`, `primary_color` |
| `powerpoint_add_image` | Insert image from local path or URL with custom dimensions, caption, and safe temporary caching. | `image_path_or_url`, `slide_index`, `left_inches`, `top_inches`, `width_inches`, `height_inches`, `caption` |
| `powerpoint_add_table` | Add professionally formatted tables with header fills, zebra row striping, and column widths. | `headers`, `data`, `slide_index`, `left_inches`, `top_inches`, `header_bg_color`, `alternate_row_color`, `col_widths` |
| `powerpoint_add_chart` | Insert native charts (Column, Bar, Line, Pie, Doughnut, Area) with series data and legends. | `chart_type`, `categories`, `series_data`, `title`, `slide_index`, `left_inches`, `top_inches` |
| `powerpoint_add_timeline` | Create visual roadmap / timeline slides with sequential milestone nodes and connectors. | `milestones`, `slide_index`, `title`, `top_inches`, `primary_color`, `accent_color` |

### 4. Enterprise Automation & Dynamic Templating
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `powerpoint_search_and_replace` | Search and replace text across all slides, text frames, tables, and speaker notes. | `find_text`, `replace_text`, `match_case`, `filepath` |
| `powerpoint_fill_template` | Fill a PowerPoint template presentation by replacing dynamic `{{tags}}` with data. | `template_path`, `output_path`, `replacements` |
| `powerpoint_apply_theme` | Apply cohesive color palette presets (`modern_blue`, `dark_tech`, `corporate_teal`, `warm_crimson`, `minimal_emerald`). | `theme`, `filepath` |
| `powerpoint_execute_script` | Execute arbitrary Python scripts with direct access to presentation objects (`prs`, `slide`, `pptx`, `app`). | `code`, `filepath` |

## 🔒 Security Notice

> [!WARNING]
> **Arbitrary Script Execution Notice**:
> The tool `powerpoint_execute_script` allows execution of arbitrary Python code with direct access to the local runtime environment and presentation objects.
> - Only enable or invoke this tool in trusted environments.
> - Never execute untrusted or unverified code snippets from external sources without prior review.
> - Exceptions and runtime errors during execution are captured and returned with `"status": "error"`.

## 🚀 Client Configuration & Setup

### 1. Installation

```bash
# Clone the repository
git clone https://github.com/your-org/powerpoint-mcp-server.git
cd powerpoint-mcp-server

# Create and activate virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
# Or install in editable mode
pip install -e .
```

### 2. Ready-to-Use Client Configurations

#### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "-m",
        "powerpoint_mcp_server"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8",
        "PYTHONPATH": "/path/to/powerpoint-mcp-server"
      }
    }
  }
}
```

#### Cursor (`mcp.json` or Settings > MCP)

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "/path/to/powerpoint-mcp-server/powerpoint_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

#### Google Antigravity / Gemini CLI (`mcp_config.json`)

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "/path/to/powerpoint-mcp-server/powerpoint_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

#### VS Code (Cline / Roo Code / Windsurf)

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": [
        "/path/to/powerpoint-mcp-server/powerpoint_mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

## 🧪 Testing & Verification

Run the comprehensive test suite verifying all 26 tools:

```bash
python test_mcp.py
```

Expected output:
```text
=======================================================
   POWERPOINT MCP SERVER ENTERPRISE VERIFICATION SUITE  
=======================================================

[1/26] Testing powerpoint_status()...
[2/26] Testing powerpoint_create_presentation()...
...
[25/26] Testing powerpoint_export_pdf()...
[26/26] Testing powerpoint_export_slide_images()...

=======================================================
   ALL 26 POWERPOINT MCP TOOLS VERIFIED & WORKING!     
=======================================================
```

> [!WARNING]
> Presentation layout generation and script execution can alter slide contents and files. There is always a possibility of error, so we assume no responsibility for any inaccuracies.

### <a name="Copyright©2026"></a> Copyright © 2026

Thank you for checking out Apex. If you have any feedback or suggestions, feel free to contact us:
hamzabellouchcontact@gmail.com

Stay connected and follow us on:  
[Facebook](https://facebook.com/hamzabellouch0) | [Instagram](https://instagram.com/hamzabellouch0) | [Twitter](https://twitter.com/hamzabellouch0) | [Telegram](https://t.me/hammzabellouch) | [LinkedIn](https://www.linkedin.com/in/hamzabellouch)

