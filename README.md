# 🎨 PPTX MCP Server

A **Model Context Protocol (MCP) server** that lets any AI model create
standardised, high-quality PowerPoint presentations via a clean JSON API.

Plug it in once — every AI that connects gets the same consistent, on-brand
slide quality every time, with zero design knowledge required.

---

## Features

| Capability | Detail |
|---|---|
| **13 slide layouts** | title, section, two-column, icon trio, stat callout, bullet, quote, timeline, comparison, chart, full image, text body, thank you |
| **6 brand themes** | Midnight Executive, Coral Energy, Forest Calm, Teal Trust, Charcoal Minimal, Berry & Cream |
| **Chart types** | Bar, Column, Line, Area, Pie, Doughnut |
| **Design rules baked in** | No AI-looking stripe accents, no cream defaults, safe font stack, shadow cards, proper contrast |
| **Transports** | stdio (Claude Desktop / local) and SSE (hosted/remote) |

---

## Quick Start

### 1. Install dependencies

```bash
# Python
pip install mcp

# Node (pptxgenjs is the rendering engine)
npm install -g pptxgenjs
```

### 2. Run the server (stdio)

```bash
python pptx_mcp_server.py
```

### 3. Run as HTTP/SSE server

```bash
python pptx_mcp_server.py --transport sse --port 8000
```

---

## Claude Desktop Config

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "pptx-maker": {
      "command": "python",
      "args": ["/absolute/path/to/pptx_mcp_server.py"]
    }
  }
}
```

---

## Tools

### `list_themes()`
Returns all brand themes with colour swatches and font choices.

### `list_slide_layouts()`
Returns all 13 supported layout types with descriptions.

### `preview_theme(theme_id)`
Returns the full colour/font spec for a theme (useful for AI to reference).

### `create_presentation(spec)`
The main tool. Accepts a JSON spec, generates a `.pptx` in `./output/`.

---

## Spec Schema

```json
{
  "title":    "My Deck",
  "author":   "Arbaz",
  "subject":  "Product Pitch",
  "theme_id": "midnight_executive",
  "filename": "my_deck.pptx",
  "slides": [ ...slide objects... ]
}
```

### Slide objects by layout

#### `title_slide`
```json
{ "layout": "title_slide", "title": "...", "subtitle": "...", "tagline": "..." }
```

#### `section_divider`
```json
{ "layout": "section_divider", "number": "01", "title": "...", "caption": "..." }
```

#### `two_column`
```json
{
  "layout": "two_column",
  "title": "...",
  "left_heading": "...", "left_body": "...",
  "right_heading": "...", "right_body": "..."
}
```

#### `icon_trio`
```json
{
  "layout": "icon_trio",
  "title": "...",
  "items": [
    {"heading": "...", "body": "..."},
    {"heading": "...", "body": "..."},
    {"heading": "...", "body": "..."}
  ]
}
```

#### `stat_callout`
```json
{
  "layout": "stat_callout",
  "title": "...",
  "stats": [
    {"value": "94%", "label": "Customer Satisfaction", "delta": "+12% YoY"},
    {"value": "£2.4M", "label": "Revenue", "delta": ""}
  ]
}
```

#### `bullet_list`
```json
{ "layout": "bullet_list", "title": "...", "bullets": ["Point 1", "Point 2"] }
```

#### `quote_highlight`
```json
{ "layout": "quote_highlight", "quote": "...", "attribution": "— Name, Role" }
```

#### `timeline`
```json
{
  "layout": "timeline",
  "title": "...",
  "steps": [
    {"label": "Phase 1", "text": "Discovery"},
    {"label": "Phase 2", "text": "Build"}
  ]
}
```

#### `comparison`
```json
{
  "layout": "comparison",
  "title": "...",
  "left_heading": "Option A", "left_points": ["Pro 1", "Pro 2"],
  "right_heading": "Option B", "right_points": ["Pro 1", "Pro 2"]
}
```

#### `chart_slide`
```json
{
  "layout": "chart_slide",
  "title": "...",
  "caption": "Source: ...",
  "chart_type": "bar",
  "chart_data": {
    "series_name": "Revenue",
    "labels": ["Q1", "Q2", "Q3", "Q4"],
    "values": [120, 145, 162, 198]
  }
}
```
Supported `chart_type` values: `bar`, `line`, `pie`, `doughnut`, `area`

#### `image_full`
```json
{ "layout": "image_full", "title": "...", "caption": "...", "image_url": "https://..." }
```

#### `text_body`
```json
{ "layout": "text_body", "title": "...", "body": "...", "note": "Optional footnote" }
```

#### `thank_you`
```json
{ "layout": "thank_you", "message": "Thank You", "contact": "...", "website": "..." }
```

---

## Available Themes

| ID | Name | Vibe |
|---|---|---|
| `midnight_executive` | Midnight Executive | Navy + ice blue — professional, authoritative |
| `coral_energy` | Coral Energy | Coral + gold — bold, startup-friendly |
| `forest_calm` | Forest Calm | Green + moss — sustainability, wellness |
| `teal_trust` | Teal Trust | Teal + seafoam — fintech, SaaS, healthcare |
| `charcoal_minimal` | Charcoal Minimal | Charcoal + off-white — luxury, editorial |
| `berry_cream` | Berry & Cream | Berry + cream — fashion, lifestyle |

---

## Design Principles (enforced automatically)

- No decorative stripe accents (AI-generated tell)
- No cream/beige defaults
- No accent underlines under titles
- Cards use shadow + tint, never edge borders
- Safe font stack (Calibri, Cambria, Bookman Old Style)
- Minimum 0.5" margins throughout
- Every layout has a visual element — no text-only slides

---

## Output

Files are saved to `./output/<filename>.pptx` relative to the server directory.
