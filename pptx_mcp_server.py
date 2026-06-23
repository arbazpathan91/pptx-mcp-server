"""
PPTX MCP Server
===============
A Model Context Protocol server that lets any AI create standardised,
high-quality PowerPoint presentations via a clean tool API.

Tools exposed:
  - create_presentation   : build a full deck from a structured spec
  - list_themes           : inspect available brand themes
  - list_slide_layouts    : inspect available layout types
  - preview_theme         : get full colour/font spec for a theme

Usage (stdio transport):
  python pptx_mcp_server.py

Usage (SSE transport, for remote/hosted):
  python pptx_mcp_server.py --transport sse --port 8000
"""

import json
import os
import subprocess
import sys
import tempfile
import textwrap
from pathlib import Path

from mcp.server.fastmcp import FastMCP

# ─────────────────────────────────────────────
# Server bootstrap
# ─────────────────────────────────────────────
mcp = FastMCP("pptx-maker")

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

SCRIPTS_DIR = Path(__file__).parent / "scripts"
SCRIPTS_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# Brand Themes
# ─────────────────────────────────────────────
THEMES: dict[str, dict] = {
    "midnight_executive": {
        "label": "Midnight Executive",
        "description": "Deep navy with ice-blue accents. Professional, authoritative.",
        "primary": "1E2761",
        "secondary": "CADCFC",
        "accent": "4FC3F7",
        "background_dark": "141B4D",
        "background_light": "F8FAFF",
        "text_dark": "0D1333",
        "text_light": "FFFFFF",
        "font_heading": "Cambria",
        "font_body": "Calibri",
    },
    "coral_energy": {
        "label": "Coral Energy",
        "description": "Vibrant coral and gold. Energetic, bold, startup-friendly.",
        "primary": "E8445A",
        "secondary": "F9C74F",
        "accent": "2F3C7E",
        "background_dark": "1A1A2E",
        "background_light": "FFFDF7",
        "text_dark": "1A1A2E",
        "text_light": "FFFFFF",
        "font_heading": "Calibri",
        "font_body": "Calibri",
    },
    "forest_calm": {
        "label": "Forest Calm",
        "description": "Forest green with moss and cream. Sustainability, wellness, nature.",
        "primary": "2C5F2D",
        "secondary": "97BC62",
        "accent": "F5F5F5",
        "background_dark": "1B3A1C",
        "background_light": "F5F8F0",
        "text_dark": "1B3A1C",
        "text_light": "FFFFFF",
        "font_heading": "Cambria",
        "font_body": "Calibri",
    },
    "teal_trust": {
        "label": "Teal Trust",
        "description": "Teal and seafoam. Fintech, SaaS, healthcare.",
        "primary": "028090",
        "secondary": "02C39A",
        "accent": "F0F4F8",
        "background_dark": "014F5A",
        "background_light": "F0FAFB",
        "text_dark": "012E35",
        "text_light": "FFFFFF",
        "font_heading": "Calibri",
        "font_body": "Calibri",
    },
    "charcoal_minimal": {
        "label": "Charcoal Minimal",
        "description": "Charcoal, off-white, black. Ultra-clean, editorial, luxury.",
        "primary": "36454F",
        "secondary": "8FA3AD",
        "accent": "212121",
        "background_dark": "1C2526",
        "background_light": "F7F7F7",
        "text_dark": "1C2526",
        "text_light": "F2F2F2",
        "font_heading": "Bookman Old Style",
        "font_body": "Calibri",
    },
    "berry_cream": {
        "label": "Berry & Cream",
        "description": "Berry, dusty rose, cream. Fashion, beauty, lifestyle.",
        "primary": "6D2E46",
        "secondary": "A26769",
        "accent": "ECE2D0",
        "background_dark": "3D1226",
        "background_light": "FBF8F4",
        "text_dark": "3D1226",
        "text_light": "FFFFFF",
        "font_heading": "Century Schoolbook",
        "font_body": "Calibri",
    },
}

# ─────────────────────────────────────────────
# Slide Layout Registry
# ─────────────────────────────────────────────
LAYOUTS = {
    "title_slide": "Full-bleed dark cover. Large title, subtitle, optional tagline.",
    "section_divider": "Dark background section break with large section number + title.",
    "two_column": "Left: heading + body text. Right: image placeholder or icon grid.",
    "icon_trio": "Three equal columns, each with icon in circle + heading + 2-line body.",
    "stat_callout": "Up to 4 large KPI numbers with labels. Impact data slide.",
    "text_body": "Light background, left-aligned title + multi-paragraph prose.",
    "image_full": "Full-bleed image with text overlay at bottom.",
    "bullet_list": "Classic title + bullet list (use sparingly; max 5 bullets).",
    "quote_highlight": "Large pull-quote centred, attribution below, dark bg.",
    "timeline": "Horizontal process / timeline with up to 5 steps.",
    "comparison": "Two-column Before/After or Option A / Option B layout.",
    "chart_slide": "Title + one chart (bar, line, or pie) + optional insight caption.",
    "thank_you": "Closing dark slide with large thank-you message and contact info.",
}

# ─────────────────────────────────────────────
# JS template generator
# ─────────────────────────────────────────────

def _build_js(spec: dict, theme: dict, output_path: str) -> str:
    """Return a complete PptxGenJS Node script for the given spec."""

    slides_js = []
    for i, slide_spec in enumerate(spec["slides"]):
        slides_js.append(_render_slide(slide_spec, theme, i))

    slides_block = "\n\n  ".join(slides_js)

    return textwrap.dedent(f"""
        const pptxgen = require('pptxgenjs');
        const pres = new pptxgen();
        pres.layout = 'LAYOUT_16x9';
        pres.title   = {json.dumps(spec.get('title', 'Presentation'))};
        pres.author  = {json.dumps(spec.get('author', 'AI Generated'))};
        pres.subject = {json.dumps(spec.get('subject', ''))};

        // ── Brand palette ──
        const T = {json.dumps(theme)};

        // ── Helper: fresh shadow ──
        const shadow = () => ({{ type:'outer', color:'000000', blur:8, offset:3, angle:45, opacity:0.12 }});

        {slides_block}

        pres.writeFile({{ fileName: {json.dumps(output_path)} }})
            .then(() => console.log('OK:' + {json.dumps(output_path)}))
            .catch(e => {{ console.error('ERR:' + e.message); process.exit(1); }});
    """).strip()


def _render_slide(s: dict, t: dict, idx: int) -> str:
    """Dispatch to per-layout renderer."""
    layout = s.get("layout", "text_body")
    dispatch = {
        "title_slide":     _slide_title,
        "section_divider": _slide_section,
        "two_column":      _slide_two_col,
        "icon_trio":       _slide_icon_trio,
        "stat_callout":    _slide_stat,
        "text_body":       _slide_text_body,
        "bullet_list":     _slide_bullet,
        "quote_highlight": _slide_quote,
        "timeline":        _slide_timeline,
        "comparison":      _slide_comparison,
        "chart_slide":     _slide_chart,
        "thank_you":       _slide_thank_you,
        "image_full":      _slide_image_full,
    }
    fn = dispatch.get(layout, _slide_text_body)
    return fn(s, t, idx)


# ── Individual layout renderers ─────────────────────────────────────────────

def _slide_title(s, t, idx):
    title    = json.dumps(s.get("title", "Presentation Title"))
    subtitle = json.dumps(s.get("subtitle", ""))
    tagline  = json.dumps(s.get("tagline", ""))
    return f"""
  // Slide {idx+1}: Title
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_dark }};
    // accent bar top
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.06, fill:{{ color:T.accent }}, line:{{ color:T.accent }} }});
    // large title
    sl.addText({title}, {{
      x:0.7, y:1.4, w:8.6, h:2.0,
      fontSize:52, fontFace:T.font_heading, bold:true, color:T.text_light,
      align:'left', valign:'middle', margin:0
    }});
    // subtitle
    sl.addText({subtitle}, {{
      x:0.7, y:3.5, w:8.6, h:0.9,
      fontSize:22, fontFace:T.font_body, color:T.secondary, align:'left', margin:0
    }});
    // tagline
    if ({tagline}) sl.addText({tagline}, {{
      x:0.7, y:4.5, w:8.6, h:0.7,
      fontSize:14, fontFace:T.font_body, italic:true, color:T.secondary,
      align:'left', margin:0, transparency:20
    }});
    // bottom accent
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:5.525, w:3.5, h:0.1, fill:{{ color:T.primary }}, line:{{ color:T.primary }} }});
  }}"""


def _slide_section(s, t, idx):
    number  = json.dumps(str(s.get("number", idx)))
    title   = json.dumps(s.get("title", "Section"))
    caption = json.dumps(s.get("caption", ""))
    return f"""
  // Slide {idx+1}: Section Divider
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.primary }};
    // big section number
    sl.addText({number}, {{
      x:0.5, y:0.3, w:2.5, h:2.5,
      fontSize:120, fontFace:T.font_heading, bold:true, color:T.text_light,
      transparency:70, align:'left', margin:0
    }});
    // title
    sl.addText({title}, {{
      x:0.7, y:2.1, w:8.6, h:1.6,
      fontSize:40, fontFace:T.font_heading, bold:true, color:T.text_light,
      align:'left', margin:0
    }});
    // caption
    if ({caption}) sl.addText({caption}, {{
      x:0.7, y:3.85, w:8.0, h:0.8,
      fontSize:16, fontFace:T.font_body, color:T.secondary, align:'left', margin:0
    }});
  }}"""


def _slide_two_col(s, t, idx):
    title      = json.dumps(s.get("title", ""))
    left_head  = json.dumps(s.get("left_heading", ""))
    left_body  = json.dumps(s.get("left_body", ""))
    right_head = json.dumps(s.get("right_heading", ""))
    right_body = json.dumps(s.get("right_body", ""))
    return f"""
  // Slide {idx+1}: Two Column
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    // title
    sl.addText({title}, {{
      x:0.5, y:0.3, w:9.0, h:0.7,
      fontSize:28, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'left', margin:0
    }});
    // divider line
    sl.addShape(pres.shapes.LINE, {{ x:0.5, y:1.1, w:9.0, h:0, line:{{ color:T.primary, width:2 }} }});
    // left column card
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:0.4, y:1.25, w:4.3, h:3.9,
      fill:{{ color:'FFFFFF' }}, rectRadius:0.1, line:{{ color:'E5E7EB', width:1 }}, shadow:shadow()
    }});
    sl.addText({left_head}, {{
      x:0.65, y:1.45, w:3.8, h:0.55,
      fontSize:16, fontFace:T.font_heading, bold:true, color:T.primary, align:'left', margin:0
    }});
    sl.addText({left_body}, {{
      x:0.65, y:2.1, w:3.8, h:2.8,
      fontSize:14, fontFace:T.font_body, color:T.text_dark, align:'left', valign:'top', margin:0
    }});
    // right column card
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:5.3, y:1.25, w:4.3, h:3.9,
      fill:{{ color:'FFFFFF' }}, rectRadius:0.1, line:{{ color:'E5E7EB', width:1 }}, shadow:shadow()
    }});
    sl.addText({right_head}, {{
      x:5.55, y:1.45, w:3.8, h:0.55,
      fontSize:16, fontFace:T.font_heading, bold:true, color:T.primary, align:'left', margin:0
    }});
    sl.addText({right_body}, {{
      x:5.55, y:2.1, w:3.8, h:2.8,
      fontSize:14, fontFace:T.font_body, color:T.text_dark, align:'left', valign:'top', margin:0
    }});
  }}"""


def _slide_icon_trio(s, t, idx):
    title = json.dumps(s.get("title", ""))
    items = s.get("items", [{"heading": "Point 1", "body": "Description"}, {"heading": "Point 2", "body": "Description"}, {"heading": "Point 3", "body": "Description"}])
    items = items[:3]

    cards_js = []
    xs = [0.4, 3.7, 7.0]
    emojis = ["✦", "◆", "★"]
    for ci, (item, x) in enumerate(zip(items, xs)):
        h = json.dumps(item.get("heading", ""))
        b = json.dumps(item.get("body", ""))
        em = json.dumps(emojis[ci])
        cards_js.append(f"""
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:{x}, y:1.3, w:2.7, h:3.8,
      fill:{{ color:'FFFFFF' }}, rectRadius:0.12, line:{{ color:'E5E7EB', width:1 }}, shadow:shadow()
    }});
    sl.addShape(pres.shapes.OVAL, {{
      x:{x+0.95}, y:1.5, w:0.8, h:0.8,
      fill:{{ color:T.primary }}, line:{{ color:T.primary }}
    }});
    sl.addText({em}, {{
      x:{x+0.95}, y:1.5, w:0.8, h:0.8,
      fontSize:18, color:T.text_light, align:'center', valign:'middle', margin:0
    }});
    sl.addText({h}, {{
      x:{x+0.1}, y:2.5, w:2.5, h:0.65,
      fontSize:15, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'center', margin:0
    }});
    sl.addText({b}, {{
      x:{x+0.1}, y:3.25, w:2.5, h:1.65,
      fontSize:13, fontFace:T.font_body, color:'4B5563', align:'center', valign:'top', margin:0
    }});""")

    cards_block = "\n".join(cards_js)
    return f"""
  // Slide {idx+1}: Icon Trio
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addText({title}, {{
      x:0.5, y:0.25, w:9.0, h:0.75,
      fontSize:28, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'center', margin:0
    }});
    sl.addShape(pres.shapes.LINE, {{ x:3.5, y:1.1, w:3.0, h:0, line:{{ color:T.accent, width:3 }} }});
    {cards_block}
  }}"""


def _slide_stat(s, t, idx):
    title = json.dumps(s.get("title", "Key Metrics"))
    stats = s.get("stats", [])[:4]
    n     = len(stats)
    col_w = 9.0 / max(n, 1)
    cards = []
    for si, stat in enumerate(stats):
        x = 0.5 + si * col_w
        v = json.dumps(str(stat.get("value", "—")))
        l = json.dumps(stat.get("label", ""))
        d = json.dumps(stat.get("delta", ""))
        cards.append(f"""
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:{x:.2f}, y:1.4, w:{col_w-0.2:.2f}, h:3.5,
      fill:{{ color:'FFFFFF' }}, rectRadius:0.12, line:{{ color:'E5E7EB', width:1 }}, shadow:shadow()
    }});
    sl.addText({v}, {{
      x:{x:.2f}, y:1.9, w:{col_w-0.2:.2f}, h:1.8,
      fontSize:64, fontFace:T.font_heading, bold:true, color:T.primary, align:'center', margin:0
    }});
    sl.addText({l}, {{
      x:{x:.2f}, y:3.75, w:{col_w-0.2:.2f}, h:0.6,
      fontSize:13, fontFace:T.font_body, color:'4B5563', align:'center', margin:0
    }});
    if ({d}) sl.addText({d}, {{
      x:{x:.2f}, y:4.4, w:{col_w-0.2:.2f}, h:0.35,
      fontSize:12, fontFace:T.font_body, italic:true, color:T.secondary, align:'center', margin:0
    }});""")

    cards_block = "\n".join(cards)
    return f"""
  // Slide {idx+1}: Stat Callout
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addText({title}, {{
      x:0.5, y:0.2, w:9.0, h:0.85,
      fontSize:28, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'center', margin:0
    }});
    {cards_block}
  }}"""


def _slide_text_body(s, t, idx):
    title = json.dumps(s.get("title", ""))
    body  = json.dumps(s.get("body", ""))
    note  = json.dumps(s.get("note", ""))
    return f"""
  // Slide {idx+1}: Text Body
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:1.0, fill:{{ color:T.primary }}, line:{{ color:T.primary }} }});
    sl.addText({title}, {{
      x:0.6, y:0.1, w:8.8, h:0.8,
      fontSize:26, fontFace:T.font_heading, bold:true, color:T.text_light, align:'left', margin:0
    }});
    sl.addText({body}, {{
      x:0.6, y:1.2, w:8.8, h:4.0,
      fontSize:16, fontFace:T.font_body, color:T.text_dark, align:'left', valign:'top', margin:0
    }});
    if ({note}) sl.addText({note}, {{
      x:0.6, y:5.2, w:8.8, h:0.3,
      fontSize:11, fontFace:T.font_body, italic:true, color:'9CA3AF', align:'left', margin:0
    }});
  }}"""


def _slide_bullet(s, t, idx):
    title   = json.dumps(s.get("title", ""))
    bullets = s.get("bullets", [])[:5]
    items_js = ",\n      ".join(
        f'{{ text:{json.dumps(b)}, options:{{ bullet:true, breakLine:true, fontSize:16, fontFace:T.font_body, color:T.text_dark }} }}'
        for b in bullets
    )
    return f"""
  // Slide {idx+1}: Bullet List
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:1.0, fill:{{ color:T.primary }}, line:{{ color:T.primary }} }});
    sl.addText({title}, {{
      x:0.6, y:0.1, w:8.8, h:0.8,
      fontSize:26, fontFace:T.font_heading, bold:true, color:T.text_light, align:'left', margin:0
    }});
    sl.addText([
      {items_js}
    ], {{ x:0.8, y:1.2, w:8.4, h:4.2, valign:'top', paraSpaceAfter:6 }});
  }}"""


def _slide_quote(s, t, idx):
    quote       = json.dumps(s.get("quote", ""))
    attribution = json.dumps(s.get("attribution", ""))
    return f"""
  // Slide {idx+1}: Quote
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_dark }};
    sl.addText('\u201c', {{
      x:0.3, y:0.1, w:2, h:2,
      fontSize:140, fontFace:T.font_heading, color:T.primary, transparency:40, margin:0
    }});
    sl.addText({quote}, {{
      x:1.0, y:1.1, w:8.0, h:3.2,
      fontSize:26, fontFace:T.font_heading, italic:true, color:T.text_light,
      align:'center', valign:'middle', margin:0
    }});
    sl.addShape(pres.shapes.LINE, {{ x:3.5, y:4.5, w:3.0, h:0, line:{{ color:T.accent, width:2 }} }});
    sl.addText({attribution}, {{
      x:1.0, y:4.7, w:8.0, h:0.6,
      fontSize:14, fontFace:T.font_body, color:T.secondary, align:'center', margin:0
    }});
  }}"""


def _slide_timeline(s, t, idx):
    title = json.dumps(s.get("title", ""))
    steps = s.get("steps", [])[:5]
    n     = len(steps)
    step_w = 9.0 / max(n, 1)
    steps_js = []
    for si, step in enumerate(steps):
        x   = 0.5 + si * step_w
        cx  = x + step_w / 2 - 0.1
        lbl = json.dumps(step.get("label", f"Step {si+1}"))
        txt = json.dumps(step.get("text", ""))
        steps_js.append(f"""
    sl.addShape(pres.shapes.OVAL, {{
      x:{cx:.2f}, y:2.25, w:0.5, h:0.5,
      fill:{{ color:T.primary }}, line:{{ color:T.primary }}
    }});
    sl.addText(String({si+1}), {{
      x:{cx:.2f}, y:2.25, w:0.5, h:0.5,
      fontSize:13, bold:true, color:T.text_light, align:'center', valign:'middle', margin:0
    }});
    sl.addText({lbl}, {{
      x:{x:.2f}, y:1.5, w:{step_w-0.1:.2f}, h:0.55,
      fontSize:13, fontFace:T.font_heading, bold:true, color:T.primary, align:'center', margin:0
    }});
    sl.addText({txt}, {{
      x:{x:.2f}, y:3.0, w:{step_w-0.1:.2f}, h:2.2,
      fontSize:12, fontFace:T.font_body, color:T.text_dark, align:'center', valign:'top', margin:0
    }});""")

    steps_block = "\n".join(steps_js)
    connector_w = 9.0 - step_w * 0.6
    return f"""
  // Slide {idx+1}: Timeline
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addText({title}, {{
      x:0.5, y:0.2, w:9.0, h:0.8,
      fontSize:28, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'center', margin:0
    }});
    sl.addShape(pres.shapes.LINE, {{
      x:{0.5 + step_w*0.5:.2f}, y:2.5, w:{connector_w:.2f}, h:0,
      line:{{ color:T.secondary, width:2 }}
    }});
    {steps_block}
  }}"""


def _slide_comparison(s, t, idx):
    title   = json.dumps(s.get("title", "Comparison"))
    l_head  = json.dumps(s.get("left_heading", "Option A"))
    r_head  = json.dumps(s.get("right_heading", "Option B"))
    l_pts   = s.get("left_points", [])
    r_pts   = s.get("right_points", [])

    def pts_js(pts, color):
        return ",\n        ".join(
            f'{{ text:{json.dumps(p)}, options:{{ bullet:true, breakLine:true, color:{json.dumps(color)}, fontSize:14, fontFace:T.font_body }} }}'
            for p in pts
        )

    return f"""
  // Slide {idx+1}: Comparison
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addText({title}, {{
      x:0.5, y:0.2, w:9.0, h:0.75,
      fontSize:28, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'center', margin:0
    }});
    // Left panel
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:0.4, y:1.1, w:4.3, h:4.2,
      fill:{{ color:T.primary }}, rectRadius:0.12, line:{{ color:T.primary }}, shadow:shadow()
    }});
    sl.addText({l_head}, {{
      x:0.5, y:1.2, w:4.1, h:0.6,
      fontSize:18, fontFace:T.font_heading, bold:true, color:T.text_light, align:'center', margin:0
    }});
    sl.addText([{pts_js(l_pts, 'FFFFFF')}], {{
      x:0.6, y:1.9, w:3.9, h:3.2, valign:'top', paraSpaceAfter:5
    }});
    // Right panel
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {{
      x:5.3, y:1.1, w:4.3, h:4.2,
      fill:{{ color:'FFFFFF' }}, rectRadius:0.12, line:{{ color:T.primary, width:2 }}, shadow:shadow()
    }});
    sl.addText({r_head}, {{
      x:5.4, y:1.2, w:4.1, h:0.6,
      fontSize:18, fontFace:T.font_heading, bold:true, color:T.primary, align:'center', margin:0
    }});
    sl.addText([{pts_js(r_pts, t['text_dark'])}], {{
      x:5.5, y:1.9, w:3.9, h:3.2, valign:'top', paraSpaceAfter:5
    }});
    // VS badge
    sl.addShape(pres.shapes.OVAL, {{ x:4.62, y:2.7, w:0.75, h:0.75, fill:{{ color:T.accent }}, line:{{ color:T.accent }} }});
    sl.addText('VS', {{ x:4.62, y:2.7, w:0.75, h:0.75, fontSize:13, bold:true, color:T.text_light, align:'center', valign:'middle', margin:0 }});
  }}"""


def _slide_chart(s, t, idx):
    title   = json.dumps(s.get("title", ""))
    caption = json.dumps(s.get("caption", ""))
    ctype   = s.get("chart_type", "bar").upper()
    data    = s.get("chart_data", {"labels": ["A", "B", "C"], "values": [30, 50, 20], "series_name": "Data"})
    labels  = json.dumps(data.get("labels", []))
    values  = json.dumps(data.get("values", []))
    sname   = json.dumps(data.get("series_name", "Data"))

    pptx_chart = {
        "BAR":  "pres.charts.BAR",
        "LINE": "pres.charts.LINE",
        "PIE":  "pres.charts.PIE",
        "AREA": "pres.charts.AREA",
        "DOUGHNUT": "pres.charts.DOUGHNUT",
    }.get(ctype, "pres.charts.BAR")

    bar_dir = "" if ctype in ("PIE", "DOUGHNUT", "LINE", "AREA") else "barDir:'col',"
    show_pct = "showPercent:true," if ctype in ("PIE", "DOUGHNUT") else ""

    return f"""
  // Slide {idx+1}: Chart
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_light }};
    sl.addText({title}, {{
      x:0.5, y:0.2, w:9.0, h:0.7,
      fontSize:26, fontFace:T.font_heading, bold:true, color:T.text_dark, align:'left', margin:0
    }});
    sl.addChart({pptx_chart}, [{{ name:{sname}, labels:{labels}, values:{values} }}], {{
      x:0.5, y:1.0, w:9.0, h:4.1,
      {bar_dir} {show_pct}
      chartColors:[T.primary, T.secondary, T.accent],
      chartArea:{{ fill:{{ color:'FFFFFF' }}, roundedCorners:true }},
      catAxisLabelColor:'64748B', valAxisLabelColor:'64748B',
      valGridLine:{{ color:'E2E8F0', size:0.5 }}, catGridLine:{{ style:'none' }},
      showLegend:false, dataLabelColor:'1E293B'
    }});
    if ({caption}) sl.addText({caption}, {{
      x:0.5, y:5.2, w:9.0, h:0.3,
      fontSize:11, fontFace:T.font_body, italic:true, color:'9CA3AF', align:'left', margin:0
    }});
  }}"""


def _slide_thank_you(s, t, idx):
    message = json.dumps(s.get("message", "Thank You"))
    contact = json.dumps(s.get("contact", ""))
    website = json.dumps(s.get("website", ""))
    return f"""
  // Slide {idx+1}: Thank You
  {{
    const sl = pres.addSlide();
    sl.background = {{ color: T.background_dark }};
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:0, w:10, h:0.08, fill:{{ color:T.accent }}, line:{{ color:T.accent }} }});
    sl.addText({message}, {{
      x:0.7, y:1.2, w:8.6, h:2.2,
      fontSize:56, fontFace:T.font_heading, bold:true, color:T.text_light,
      align:'center', valign:'middle', margin:0
    }});
    sl.addShape(pres.shapes.LINE, {{ x:2.5, y:3.6, w:5.0, h:0, line:{{ color:T.primary, width:2 }} }});
    if ({contact}) sl.addText({contact}, {{
      x:0.7, y:3.9, w:8.6, h:0.6,
      fontSize:18, fontFace:T.font_body, color:T.secondary, align:'center', margin:0
    }});
    if ({website}) sl.addText({website}, {{
      x:0.7, y:4.6, w:8.6, h:0.5,
      fontSize:15, fontFace:T.font_body, italic:true, color:T.accent, align:'center', margin:0
    }});
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:5.525, w:10, h:0.1, fill:{{ color:T.primary }}, line:{{ color:T.primary }} }});
  }}"""


def _slide_image_full(s, t, idx):
    title   = json.dumps(s.get("title", ""))
    caption = json.dumps(s.get("caption", ""))
    img_url = s.get("image_url", "")
    img_block = ""
    if img_url:
        img_block = f"sl.addImage({{ path:{json.dumps(img_url)}, x:0, y:0, w:10, h:5.625, sizing:{{type:'cover',w:10,h:5.625}} }});"
    return f"""
  // Slide {idx+1}: Full Image
  {{
    const sl = pres.addSlide();
    sl.background = {{ color:'000000' }};
    {img_block}
    sl.addShape(pres.shapes.RECTANGLE, {{ x:0, y:3.8, w:10, h:1.825, fill:{{ color:'000000', transparency:40 }}, line:{{ color:'000000' }} }});
    sl.addText({title}, {{
      x:0.6, y:3.9, w:8.8, h:1.0,
      fontSize:30, fontFace:T.font_heading, bold:true, color:'FFFFFF', align:'left', margin:0
    }});
    if ({caption}) sl.addText({caption}, {{
      x:0.6, y:5.0, w:8.8, h:0.45,
      fontSize:13, fontFace:T.font_body, color:'E5E7EB', italic:true, align:'left', margin:0
    }});
  }}"""


# ─────────────────────────────────────────────
# Core execution helper
# ─────────────────────────────────────────────

def _run_js(js_code: str, output_path: str) -> str:
    """Write JS to a temp file, run with node, return the output file path."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(js_code)
        tmp_js = f.name
    try:
        result = subprocess.run(
            ["node", tmp_js],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(f"Node error: {result.stderr.strip()}")
        return output_path
    finally:
        os.unlink(tmp_js)


# ─────────────────────────────────────────────
# MCP Tools
# ─────────────────────────────────────────────

@mcp.tool()
def list_themes() -> str:
    """
    List all available brand themes with their colours and fonts.
    Call this before create_presentation to pick the right theme_id.
    """
    out = []
    for tid, info in THEMES.items():
        out.append(
            f"**{tid}** — {info['label']}\n"
            f"  {info['description']}\n"
            f"  Primary: #{info['primary']}  Secondary: #{info['secondary']}  Accent: #{info['accent']}\n"
            f"  Fonts: {info['font_heading']} (heading) / {info['font_body']} (body)"
        )
    return "\n\n".join(out)


@mcp.tool()
def list_slide_layouts() -> str:
    """
    List all supported slide layout types with descriptions.
    Use these layout names in the 'layout' field of each slide spec.
    """
    return "\n".join(f"**{k}**: {v}" for k, v in LAYOUTS.items())


@mcp.tool()
def preview_theme(theme_id: str) -> str:
    """
    Return the full colour and font specification for a given theme_id.

    Args:
        theme_id: One of the theme IDs returned by list_themes().
    """
    t = THEMES.get(theme_id)
    if not t:
        return f"Unknown theme '{theme_id}'. Call list_themes() for valid options."
    return json.dumps(t, indent=2)


@mcp.tool()
def create_presentation(spec: str) -> str:
    """
    Generate a standardised, high-quality PowerPoint (.pptx) file.

    Args:
        spec: JSON string with the full presentation specification (see schema below).

    ── Spec Schema ──────────────────────────────────────────────────────────
    {
      "title":   "Presentation Title",          // required
      "author":  "Your Name",                   // optional
      "subject": "Topic",                        // optional
      "theme_id": "midnight_executive",          // optional, default: midnight_executive
      "filename": "my_deck.pptx",               // optional, default: output.pptx
      "slides": [                               // required, array of slide objects
        {
          "layout": "title_slide",              // required — see list_slide_layouts()
          "title": "...",
          "subtitle": "...",
          "tagline": "..."
        },
        {
          "layout": "two_column",
          "title": "...",
          "left_heading": "...", "left_body": "...",
          "right_heading": "...", "right_body": "..."
        },
        {
          "layout": "icon_trio",
          "title": "...",
          "items": [
            {"heading": "...", "body": "..."},
            {"heading": "...", "body": "..."},
            {"heading": "...", "body": "..."}
          ]
        },
        {
          "layout": "stat_callout",
          "title": "...",
          "stats": [
            {"value": "94%", "label": "Customer Satisfaction", "delta": "+12% YoY"},
            {"value": "£2.4M", "label": "Revenue", "delta": ""}
          ]
        },
        {
          "layout": "bullet_list",
          "title": "...",
          "bullets": ["Point one", "Point two", "Point three"]
        },
        {
          "layout": "quote_highlight",
          "quote": "The best way to predict the future is to create it.",
          "attribution": "— Peter Drucker"
        },
        {
          "layout": "timeline",
          "title": "...",
          "steps": [
            {"label": "Phase 1", "text": "Discovery & scoping"},
            {"label": "Phase 2", "text": "Build"}
          ]
        },
        {
          "layout": "comparison",
          "title": "...",
          "left_heading": "Option A", "left_points": ["Pro 1", "Pro 2"],
          "right_heading": "Option B", "right_points": ["Pro 1", "Pro 2"]
        },
        {
          "layout": "chart_slide",
          "title": "...",
          "caption": "Source: ...",
          "chart_type": "bar",         // bar | line | pie | doughnut | area
          "chart_data": {
            "series_name": "Revenue",
            "labels": ["Q1", "Q2", "Q3", "Q4"],
            "values": [120, 145, 162, 198]
          }
        },
        {
          "layout": "thank_you",
          "message": "Thank You",
          "contact": "hello@example.com",
          "website": "www.example.com"
        }
      ]
    }
    ─────────────────────────────────────────────────────────────────────────

    Returns a message with the path to the generated .pptx file.
    """
    try:
        parsed = json.loads(spec)
    except json.JSONDecodeError as e:
        return f"❌ Invalid JSON in spec: {e}"

    if "slides" not in parsed or not parsed["slides"]:
        return "❌ spec must contain a non-empty 'slides' array."

    theme_id = parsed.get("theme_id", "midnight_executive")
    theme    = THEMES.get(theme_id, THEMES["midnight_executive"])
    filename = parsed.get("filename", "output.pptx")
    if not filename.endswith(".pptx"):
        filename += ".pptx"

    output_path = str(OUTPUT_DIR / filename)

    try:
        js_code = _build_js(parsed, theme, output_path)
        _run_js(js_code, output_path)
    except Exception as e:
        return f"❌ Generation failed: {e}"

    size_kb = Path(output_path).stat().st_size // 1024
    n_slides = len(parsed["slides"])
    return (
        f"✅ Presentation created successfully!\n"
        f"   File:   {output_path}\n"
        f"   Slides: {n_slides}\n"
        f"   Size:   {size_kb} KB\n"
        f"   Theme:  {theme['label']}\n\n"
        f"The file is ready for download or further editing in PowerPoint."
    )


# ─────────────────────────────────────────────
# Entry point

if __name__ == "__main__":
    import uvicorn
    from mcp.server.sse import SseServerTransport
    from starlette.applications import Starlette
    from starlette.responses import JSONResponse, Response
    from starlette.routing import Route, Mount
    from starlette.requests import Request

    port = int(os.environ.get("PORT", 8000))

    sse_transport = SseServerTransport("/messages/")

    async def handle_sse(request: Request):
        """SSE transport - for Claude Desktop, older clients"""
        async with sse_transport.connect_sse(
            request.scope, request.receive, request._send
        ) as streams:
            await mcp._mcp_server.run(
                streams[0], streams[1],
                mcp._mcp_server.create_initialization_options()
            )

    async def handle_streamable(request: Request):
        """Streamable HTTP transport - for Gemini, ChatGPT, newer clients"""
        body = await request.body()
        headers = dict(request.headers)

        import json as _json
        from mcp.types import JSONRPCMessage

        try:
            payload = _json.loads(body)
        except Exception:
            return Response("Bad Request", status_code=400)

        # Handle initialize / ping / tools/list / tools/call
        method = payload.get("method", "")
        msg_id = payload.get("id", 1)

        if method == "initialize":
            result = {
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "pptx-maker", "version": "1.0.0"}
                }
            }
            return JSONResponse(result)

        elif method == "notifications/initialized":
            return Response(status_code=204)

        elif method == "ping":
            return JSONResponse({"jsonrpc": "2.0", "id": msg_id, "result": {}})

        elif method == "tools/list":
            tools = []
            for tool_name, tool_fn in mcp._tool_manager._tools.items():
                tools.append({
                    "name": tool_name,
                    "description": tool_fn.description or "",
                    "inputSchema": tool_fn.parameters or {"type": "object", "properties": {}}
                })
            return JSONResponse({
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {"tools": tools}
            })

        elif method == "tools/call":
            params = payload.get("params", {})
            tool_name = params.get("name", "")
            arguments = params.get("arguments", {})

            try:
                result = await mcp._tool_manager.call_tool(tool_name, arguments)
                content = [{"type": "text", "text": str(result)}]
                return JSONResponse({
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "result": {"content": content, "isError": False}
                })
            except Exception as e:
                return JSONResponse({
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "result": {
                        "content": [{"type": "text", "text": f"Error: {str(e)}"}],
                        "isError": True
                    }
                })

        return JSONResponse({
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": {"code": -32601, "message": f"Method not found: {method}"}
        })

    async def health(request: Request):
        return JSONResponse({
            "status": "ok",
            "server": "pptx-mcp-server",
            "transport": ["sse (/sse)", "streamable-http (POST /)"]
        })

    starlette_app = Starlette(
        routes=[
            Route("/", endpoint=health, methods=["GET"]),
            Route("/", endpoint=handle_streamable, methods=["POST"]),
            Route("/health", endpoint=health, methods=["GET"]),
            Route("/mcp", endpoint=handle_streamable, methods=["POST"]),
            Route("/sse", endpoint=handle_sse, methods=["GET"]),
            Mount("/messages/", app=sse_transport.handle_post_message),
        ]
    )

    uvicorn.run(starlette_app, host="0.0.0.0", port=port)