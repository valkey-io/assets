#!/usr/bin/env python3
"""Build slides/Valkey placeholder slides.pptx from the Valkey template.

Re-runnable: strips the template's slides, builds the placeholder slides fresh,
rasterizes PNG fallbacks for each SVG, and embeds the SVG icons.

    python3 slides/build_placeholder_deck.py

Run from the repo root. Requires python-pptx, Pillow, and Google Chrome.
"""
import math
import os
import re
import shutil
import subprocess
import tempfile
import sys
import xml.etree.ElementTree as ET

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_svg import add_svg
from pptx_layout import (add_placeholder, audit, drawing_canvas, new_layout,
                         solid_fill, strip_stock_layouts)

HERE = os.path.dirname(os.path.abspath(__file__))
# Donor for theme/master/layouts. Deliberately the community deck: it is the only
# git-tracked pptx here, so it is recoverable, and it is not a deck being actively
# edited in PowerPoint. Never point this at a deck someone has open.
TEMPLATE = os.path.join(HERE, "Valkey community deck.pptx")
OUT = os.path.join(HERE, "Valkey placeholder slides.pptx")
LOGOS = os.path.join(os.path.dirname(HERE), "Valkey Branding", "logo svgs")
NEEDED_LAYOUTS = ("Title Only", "Title + Cont", "Holding")

# Brand logo aspect ratios, from each file's viewBox. Never distort these.
LOGO_ASPECT = {"horizontal": 187.9 / 63.5, "stacked": 778.6 / 669.9,
               "mark": 282.6 / 326.0}

# The brand pack ships only lockups (glyph + "Valkey" wordmark). At corner-mark size
# the wordmark is illegible, so make_marks() derives a glyph-only SVG from the stacked
# logo by dropping the wordmark group and tightening the viewBox to the glyph. The
# viewBox was measured by rasterising the glyph and taking its bounding box.
MARK_VIEWBOX = "248.0 0 282.6 326.0"
MARK_FILES = ("valkey-mark-color", "valkey-mark-white")
ICONS = os.path.join(HERE, "icons")
PNGS = os.path.join(ICONS, "png")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# --- template geometry (see ../slides-skill.md) ---
TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H = 609600, 400019, 11062800, 443198
CONTENT_LEFT, CONTENT_TOP, CONTENT_W, CONTENT_H = 609600, 1445931, 11062800, 4565111
SLIDE_W, SLIDE_H = 12192000, 6858000

# Text runs pin Arial explicitly. The theme fonts stay Play / Century Gothic, and
# inherited placeholder styling still uses them, but every run this builder creates
# names Arial so the deck is consistent. Community deck slides 7-8 do the same.
FONT = "Arial"
TITLE_FONT = "Arial"

GAP = 200000
PAD = 182880
BAR_H = 73152
CARD_ICON = 1097280     # 1.2in, sized to hold its own against an 18pt heading
GALLERY_ICON = 914400   # 1.0in, fits the 4x3 grid cell with a label beneath
CARD_H = 3900000        # bottom lands at 5345931, clear of the logo at 5943597

# name -> (heading, body). Placeholder copy is deliberately generic.
ICON_LABELS = {
    "horizontal-scaling": "Horizontal scaling",
    "vertical-scaling": "Vertical scaling",
    "replication": "Replication",
    "performance": "Performance",
    "memory-efficiency": "Memory efficiency",
    "cost-efficiency": "Cost efficiency",
    "security": "Security",
    "hardening": "Hardening",
    "reliability": "Reliability",
    "observability": "Observability",
    "migration": "Migration",
    "community": "Community",
    "durability": "Durability",
    "general-scaling": "Scaling",
    "data-list": "Lists",
    "data-hash": "Hashes",
}


# The Valkey palette has no red. ACCENT_3 (#632637) is the nearest thing and reads
# as brown, not as failure. Failure states therefore use an explicit RGB red: the
# one documented exception to the theme-colours-only rule.
# The template's accent6 was peach (#FBD8BF). patch_theme() rewrites it to the brand
# rose so it is a real theme colour, pickable in PowerPoint, not a hardcoded RGB.
BRAND_ROSE_HEX = "E0A2AF"                    # Pantone 494 C, CMYK 4 45 15 0
CARD_ROSE = MSO_THEME_COLOR.ACCENT_6         # the patched accent6
CARD_TEXT = MSO_THEME_COLOR.ACCENT_2         # #1A2026, already in the theme
# The deep purple pairing. Nothing uses it: it is not offered as a layout because it
# needs LIGHT_1 text and white icon variants that do not exist, so it is not a fill
# swap. Kept named so opting in stays one call:
#   slide_two_box(..., (ACCENT_5, CARD_PURPLE), texts=(CARD_TEXT, CARD_TEXT_ON_DARK))
CARD_PURPLE = MSO_THEME_COLOR.ACCENT_4       # #30176E, Pantone 2685 C: already in the theme
CARD_TEXT_ON_DARK = MSO_THEME_COLOR.LIGHT_1  # #1A2026 is unreadable on the deep purple

FAIL_RED = RGBColor(0xD3, 0x2F, 0x2F)        # default
FAIL_RED_BRICK = RGBColor(0xC0, 0x39, 0x2B)  # darker alternative


def _set_colour(target, colour):
    """Accept either a theme colour or an explicit RGBColor."""
    if isinstance(colour, RGBColor):
        target.rgb = colour
    else:
        target.theme_color = colour


def svg_size(svg_path):
    """Intrinsic render size of an SVG, as Chrome will treat it.

    An SVG with explicit width/height renders at that size regardless of the
    window. One with only a viewBox scales to fill the viewport, so the viewBox
    dimensions are the right window size to preserve its aspect ratio.
    """
    root = ET.parse(svg_path).getroot()

    def dim(v):
        return math.ceil(float(re.sub(r"[a-z%]+$", "", v))) if v else None

    w, h = dim(root.get("width")), dim(root.get("height"))
    if w and h:
        return w, h
    vb = root.get("viewBox").split()
    return math.ceil(float(vb[2])), math.ceil(float(vb[3]))


def rasterize(svg_path, png_path, scale=4):
    """Render an SVG to a transparent PNG fallback via headless Chrome.

    Always goes through an HTML wrapper with an explicitly sized <img>. Pointing
    Chrome straight at a .svg is unreliable: the Illustrator-exported brand logos
    render as a fully transparent image that way, and an SVG that declares
    width/height renders at its natural size inside the window rather than
    scaling to fit, which silently corner-anchors the artwork. Sizing the <img>
    in CSS sidesteps both.
    """
    w, h = svg_size(svg_path)
    tw, th = w * scale, h * scale
    with tempfile.TemporaryDirectory() as tmp:
        wrapper = os.path.join(tmp, "w.html")
        with open(wrapper, "w") as fh:
            fh.write(f'<html><body style="margin:0">'
                     f'<img src="file://{svg_path}" '
                     f'style="width:{tw}px;height:{th}px;display:block"></body></html>')
        subprocess.run(
            [CHROME, "--headless", "--disable-gpu", "--no-sandbox",
             "--default-background-color=00000000", "--virtual-time-budget=3000",
             f"--screenshot={png_path}", f"--window-size={tw},{th}",
             f"file://{wrapper}"],
            check=True, capture_output=True,
        )
    if not os.path.exists(png_path) or os.path.getsize(png_path) < 200:
        raise RuntimeError(f"rasterize produced nothing for {svg_path}")

    # The artwork must fill the canvas, or it renders small and corner-anchored
    # in any viewer that falls back to the PNG instead of the SVG.
    from PIL import Image
    im = Image.open(png_path).convert("RGBA")
    bbox = im.getbbox()
    if bbox is None:
        raise RuntimeError(f"{png_path}: fully transparent, nothing rendered")
    fill = max((bbox[2] - bbox[0]) / im.size[0], (bbox[3] - bbox[1]) / im.size[1])
    if fill < 0.6:
        raise RuntimeError(f"{png_path}: content fills only {fill:.0%} of canvas")


def logo_src(name):
    """Derived marks live in icons/; the shipped lockups live in the brand folder."""
    return ICONS if name.startswith("valkey-mark-") else LOGOS


def make_marks():
    """Derive glyph-only SVGs from the stacked lockups."""
    from lxml import etree as LET
    os.makedirs(ICONS, exist_ok=True)
    for colour in ("color", "white"):
        tree = LET.parse(os.path.join(LOGOS, f"valkey-stacked-{colour}.svg"))
        root = tree.getroot()
        for g in root.iter("{http://www.w3.org/2000/svg}g"):
            if (g.get("id") or "").startswith("text"):
                g.getparent().remove(g)
                break
        else:
            raise RuntimeError(f"wordmark group not found in valkey-stacked-{colour}.svg")
        root.set("viewBox", MARK_VIEWBOX)
        w, h = MARK_VIEWBOX.split()[2:]
        root.set("width", w)
        root.set("height", h)
        tree.write(os.path.join(ICONS, f"valkey-mark-{colour}.svg"))


def logo(slide, variant, colour, left, top, height):
    """Place a brand logo, deriving width from the real aspect ratio."""
    name = f"valkey-{variant}-{colour}"
    width = int(height * LOGO_ASPECT[variant])
    add_svg(slide, os.path.join(logo_src(name), f"{name}.svg"),
            os.path.join(PNGS, f"{name}.png"),
            left, top, width, height, name=name)
    return width


def text_box(slide, left, top, width, height, runs, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, spacing=None, italic=False):
    """Add a plain text box. `runs` is a list of (text, size_pt, bold, theme_color)."""
    box = slide.shapes.add_textbox(Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    para = tf.paragraphs[0]
    para.alignment = align
    if spacing is not None:
        para.line_spacing = spacing
    for text, size, bold, color in runs:
        r = para.add_run()
        r.text = text
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic or None
        _set_colour(r.font.color, color)
    return box


def rect(slide, left, top, width, height, fill=None, line=None):
    """Rectangle. `fill`/`line` are theme colors; None means no fill / no border."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(left)), Emu(int(top)),
                                Emu(int(width)), Emu(int(height)))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        _set_colour(sh.fill.fore_color, fill)
    if line is None:
        sh.line.fill.background()
    else:
        _set_colour(sh.line.color, line)
        sh.line.width = Pt(0.75)
    sh.text_frame.text = ""
    return sh


def patch_theme(prs):
    """Rewrite accent6 in the theme colour scheme to the brand rose.

    The theme part is not parsed by python-pptx, so edit its blob directly. Target
    only the clrScheme entry: accent6 may also appear as an explicit srgbClr
    elsewhere, and those are not ours to change.
    """
    for part in prs.part.package.iter_parts():
        if str(part.partname) != "/ppt/theme/theme1.xml":
            continue
        xml = part.blob.decode("utf-8")
        patched, n = re.subn(
            r'(<a:accent6>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
            r"\g<1>" + BRAND_ROSE_HEX + r"\g<2>", xml, count=1)
        if n != 1:
            raise RuntimeError("could not find accent6 in the theme colour scheme")
        part._blob = patched.encode("utf-8")
        print(f"theme accent6 -> #{BRAND_ROSE_HEX}")
        return
    raise RuntimeError("theme1.xml not found")


def strip_slides(prs):
    for i in range(len(prs.slides) - 1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]


def new_slide(prs, layout_name, title=None):
    layout = prs.slide_layouts.get_by_name(layout_name)
    if layout is None:
        raise SystemExit(f"layout {layout_name!r} not in this template")
    slide = prs.slides.add_slide(layout)
    if title is not None and slide.shapes.title is not None:
        tf = slide.shapes.title.text_frame
        tf.text = title
        for run in tf.paragraphs[0].runs:
            run.font.name = TITLE_FONT
    elif slide.shapes.title is not None:
        # drop an unused title placeholder so PowerPoint shows no empty field
        slide.shapes._spTree.remove(slide.shapes.title._element)
    return slide


CARD_HEAD_H, CARD_HEAD_GAP, CARD_BODY_GAP = 400110, 137160, 45720


def box_columns(n):
    """Column width and left edges for an n-up card row."""
    w = (CONTENT_W - (n - 1) * GAP) // n
    return w, [CONTENT_LEFT + i * (w + GAP) for i in range(n)]


def card_geometry(x, y, w, h):
    """The icon, heading and body rectangles inside a card.

    Shared by card() and the Two Box / Three Box layouts, so a slide built by this
    script and a slide built from the layout in PowerPoint cannot drift apart.
    """
    inner_w = w - 2 * PAD
    iy = y + PAD
    hy = iy + CARD_ICON + CARD_HEAD_GAP
    by = hy + CARD_HEAD_H + CARD_BODY_GAP
    return ((x + PAD, iy, CARD_ICON, CARD_ICON),
            (x + PAD, hy, inner_w, CARD_HEAD_H),
            (x + PAD, by, inner_w, h - (by - y) - PAD))


def card(slide, x, y, w, h, fill, icon, heading, body, border=None, text=CARD_TEXT):
    """One card: background, icon, heading, body.

    fill=None + border gives an uncoloured outline card. No accent stripe: the fill
    already defines the card, and a thin bar on top of it is decoration. `text`
    colours both the heading and the body, so a dark fill can pass LIGHT_1.
    """
    rect(slide, x, y, w, h, fill=fill, line=border)
    icon_r, head_r, body_r = card_geometry(x, y, w, h)
    add_svg(slide, os.path.join(ICONS, f"{icon}.svg"),
            os.path.join(PNGS, f"{icon}.png"), *icon_r, name=icon)
    text_box(slide, *head_r, [(heading, 18, True, text)])
    text_box(slide, *body_r, [(body, 13, False, text)], spacing=1.2)


def slide_two_box(prs, title, left_spec, right_spec, fills, border=None, texts=None):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, title)
    col_w, xs = box_columns(2)
    texts = texts or (CARD_TEXT, CARD_TEXT)
    for i, (spec, fill) in enumerate(zip((left_spec, right_spec), fills)):
        card(slide, xs[i], CONTENT_TOP, col_w, CARD_H, fill, *spec,
             border=border, text=texts[i])
    return slide


def slide_two_box_calm(prs, title, left_spec, right_spec):
    """Alternative to the purple/peach pairing: purple against neutral grey.

    Keeps the two cards distinguishable using only purple against neutral grey.
    """
    return slide_two_box(prs, title, left_spec, right_spec,
                         (MSO_THEME_COLOR.ACCENT_5, MSO_THEME_COLOR.LIGHT_2))


def slide_three_box(prs, title, specs, fill):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, title)
    col_w, xs = box_columns(3)
    for i, spec in enumerate(specs):
        card(slide, xs[i], CONTENT_TOP, col_w, CARD_H, fill, *spec)
    return slide


def source_note(slide, text):
    """Source citation, pinned to the logo baseline. Never put this in the body."""
    return text_box(slide, CONTENT_LEFT, 5943597, CONTENT_W, 300000,
                    [(text, 10, False, MSO_THEME_COLOR.DARK_1)], italic=True)


V8_TITLE = "Valkey 8.0 (September 2024)"
V8_SOURCE = "Source: valkey.io/blog/valkey-8-0-ga/"
V8_ITEMS = [("3x", " throughput with multi-threaded I/O"),
            ("25%+", " memory reduction with embedded hashtable entries"),
            ("1M+", " RPS on a single node"),
            "Dual-channel replication for faster failovers",
            "Improved cluster bus reliability",
            "Enhanced observability with per-slot metrics"]
V8_ICONS = ["performance", "memory-efficiency", "vertical-scaling",
            "replication", "reliability", "observability"]


def item_runs(item, size, text_colour=MSO_THEME_COLOR.DARK_1):
    """(highlight, rest) tuple -> two runs; plain string -> one."""
    if isinstance(item, tuple):
        return [(item[0], size, True, MSO_THEME_COLOR.ACCENT_1),
                (item[1], size, False, text_colour)]
    return [(item, size, False, text_colour)]


def slide_bullets_cards(prs, title, items, source=None):
    """Alternative A: one soft row card per item, blue left edge instead of a caret."""
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, title)
    row_h, gap = 520000, 130000
    for i, item in enumerate(items):
        y = CONTENT_TOP + i * (row_h + gap)
        rect(slide, CONTENT_LEFT, y, CONTENT_W, row_h, fill=MSO_THEME_COLOR.LIGHT_2)
        rect(slide, CONTENT_LEFT, y, BAR_H, row_h, fill=MSO_THEME_COLOR.ACCENT_1)
        text_box(slide, CONTENT_LEFT + BAR_H + PAD, y, CONTENT_W - BAR_H - 2 * PAD, row_h,
                 item_runs(item, 18), anchor=MSO_ANCHOR.MIDDLE)
    if source:
        source_note(slide, source)
    return slide


def slide_bullets_soft(prs, title, items, source=None):
    """Alternative B: still a list, just quieter. Small square bullet, softer text,
    narrower measure, more air between lines."""
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, title)
    row = 620000
    for i, item in enumerate(items):
        y = CONTENT_TOP + i * row
        rect(slide, CONTENT_LEFT, y + 105000, 110000, 110000, fill=MSO_THEME_COLOR.ACCENT_1)
        text_box(slide, CONTENT_LEFT + 360000, y, 8500000, 520000,
                 item_runs(item, 20, MSO_THEME_COLOR.DARK_2), spacing=1.3)
    if source:
        source_note(slide, source)
    return slide


def slide_bullets_icons(prs, title, items, icons, source=None):
    """Alternative C: an icon per row with hairline dividers. Most designed of the
    three, and it puts the icon library to work."""
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, title)
    row, icon = 640000, 520000
    for i, (item, name) in enumerate(zip(items, icons)):
        y = CONTENT_TOP + i * row
        add_svg(slide, os.path.join(ICONS, f"{name}.svg"),
                os.path.join(PNGS, f"{name}.png"),
                CONTENT_LEFT, y + (row - icon) // 2 - 40000, icon, icon, name=name)
        text_box(slide, CONTENT_LEFT + 760000, y, CONTENT_W - 760000, row - 80000,
                 item_runs(item, 18), anchor=MSO_ANCHOR.MIDDLE)
        if i < len(items) - 1:
            rect(slide, CONTENT_LEFT, y + row - 40000, CONTENT_W, 12700,
                 fill=MSO_THEME_COLOR.LIGHT_2)
    if source:
        source_note(slide, source)
    return slide


def slide_gallery(prs, names):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "Icon library")
    cols = 4
    rows = math.ceil(len(names) / cols)
    # The template's content box is NOT symmetric: left inset 609600 but width
    # 11062800, which overhangs the right margin by 90000 EMU. That is invisible
    # in prose but obvious in a grid, so centre this on the slide instead.
    grid_w = SLIDE_W - 2 * CONTENT_LEFT
    grid_left = (SLIDE_W - grid_w) // 2
    cell_w = grid_w // cols
    row_pitch = CONTENT_H // rows
    label_gap, label_h = 91440, 300000

    # Shrink the icon if the row count demands it, so extra rows never overlap.
    icon = min(GALLERY_ICON, row_pitch - label_gap - label_h - 120000)

    # Centre the grid on the PAGE, not in the content box. Anchoring it to
    # CONTENT_TOP left the block sitting 0.33in below the page's vertical centre,
    # which is visible on a slide that is nothing but a grid.
    block_h = (rows - 1) * row_pitch + icon + label_gap + label_h
    grid_top = (SLIDE_H - block_h) // 2

    for i, name in enumerate(names):
        cx = grid_left + (i % cols) * cell_w
        iy = grid_top + (i // cols) * row_pitch
        add_svg(slide, os.path.join(ICONS, f"{name}.svg"),
                os.path.join(PNGS, f"{name}.png"),
                cx + (cell_w - icon) // 2, iy, icon, icon, name=name)
        text_box(slide, cx, iy + icon + label_gap, cell_w, label_h,
                 [(ICON_LABELS[name], 12, False, MSO_THEME_COLOR.DARK_2)],
                 align=PP_ALIGN.CENTER)
    return slide


# --- layouts in the slide master -----------------------------------------------
# These designs used to be drawn shape-by-shape onto a slide, which meant every reuse
# was a copy-paste and nothing was editable as a placeholder. They are slide layouts
# now: the decoration lives in the master and PowerPoint's New Slide gallery offers
# them by name.
#
# The split in every builder below is the same one: decoration is drawn onto the
# layout with the ordinary slide helpers via drawing_canvas(), and every piece of text
# a presenter replaces is an add_placeholder(). Static layout shapes are inherited by
# the slide but cannot be selected there, so anything that needs editing must be a
# placeholder, and anything that must be deletable per-slide has to be one shape.
#
# All of these are white-background designs, so they build on nothing: a fresh layout
# inherits the master's white (lt1). The stock title layouts look dark because the
# *layout* carries a full-bleed background picture, which is not something you can
# turn off without covering it up.

TITLE_TXT = "Presentation title goes here"
SUB_TXT = "A single line that sets up the talk"
WHO, ROLE = "Madelyn Olson", "Valkey Maintainer and Engineer"

# Where the master's own logo sits. A full-bleed fill covers it (slide and layout
# shapes draw over master shapes), so those layouts re-place a white one here.
MASTER_LOGO_TOP, MASTER_LOGO_H = 5943597, 443198
MASTER_LOGO_RIGHT = 11824163

LAYOUT_TITLE_ONLY = "Valkey Title Only"
LAYOUT_TITLE_BAND = "Valkey Title Band"
LAYOUT_TITLE_TOP = "Valkey Title Top Banner"
LAYOUT_TITLE_BOTTOM = "Valkey Title Bottom Banner"
LAYOUT_INSIGHT = "Valkey Insight"
LAYOUT_INSIGHT_BOLD = "Valkey Insight Bold"
LAYOUT_BIG_NUMBER = "Valkey Big Number"
LAYOUT_LIST_ROWS = "Valkey List Rows"
LAYOUT_TWO_BOX = "Valkey Two Box"
LAYOUT_THREE_BOX = "Valkey Three Box"
LAYOUT_THANKS = "Valkey Thank You"


def layout_title(layout):
    """The standard title placeholder, at the master's own title geometry."""
    add_placeholder(layout, "title", 0, TITLE_LEFT, TITLE_TOP, TITLE_W, TITLE_H,
                    "Title is an assertion, not a label", 32, bold=True,
                    font=TITLE_FONT, anchor="ctr", name="Title")


def layout_title_only(prs):
    """A title and an empty canvas: the layout every custom slide is built on.

    Not a design, a foundation. The stock `Title Only` did this job, and the diagram
    slides all sat on it, which is why it has to exist before the stock layouts can go.
    """
    lo = new_layout(prs, LAYOUT_TITLE_ONLY)
    layout_title(lo)
    return lo


def layout_title_band(prs):
    """Title option A: full-height Valkey blue band on the right, stacked white logo."""
    lo = new_layout(prs, LAYOUT_TITLE_BAND)
    c = drawing_canvas(lo)
    band_w = 4114800                                  # right third
    band_x = SLIDE_W - band_w
    rect(c, band_x, 0, band_w, SLIDE_H, fill=MSO_THEME_COLOR.ACCENT_1)

    logo_h = 1600200
    logo_w = int(logo_h * LOGO_ASPECT["stacked"])
    logo(c, "stacked", "white", band_x + (band_w - logo_w) // 2,
         (SLIDE_H - logo_h) // 2, logo_h)

    col_w = band_x - CONTENT_LEFT - 685800
    add_placeholder(lo, "ctrTitle", 0, CONTENT_LEFT, 2360930, col_w, 1500000,
                    TITLE_TXT, 40, bold=True, spacing=1.1, name="Title")
    add_placeholder(lo, "subTitle", 1, CONTENT_LEFT, 3160000, col_w, 400000,
                    SUB_TXT, 18, colour=MSO_THEME_COLOR.DARK_2, name="Subtitle")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 4000000, 4000000, 300000,
                    WHO, 16, bold=True, name="Speaker")
    add_placeholder(lo, "body", 3, CONTENT_LEFT, 4330000, 4000000, 300000,
                    ROLE, 13, colour=MSO_THEME_COLOR.DARK_2, name="Speaker role")
    return lo


def layout_title_top_banner(prs):
    """Title option B: dark banner across the top with the horizontal white logo."""
    lo = new_layout(prs, LAYOUT_TITLE_TOP)
    c = drawing_canvas(lo)
    band_h = 1828800
    rect(c, 0, 0, SLIDE_W, band_h, fill=MSO_THEME_COLOR.ACCENT_2)
    rect(c, 0, band_h, SLIDE_W, BAR_H, fill=MSO_THEME_COLOR.ACCENT_1)

    logo_h = 731520
    logo_w = int(logo_h * LOGO_ASPECT["horizontal"])
    logo(c, "horizontal", "white", SLIDE_W - CONTENT_LEFT - logo_w,
         (band_h - logo_h) // 2, logo_h)

    add_placeholder(lo, "ctrTitle", 0, CONTENT_LEFT, 2700000, 9800000, 1600000,
                    TITLE_TXT, 44, bold=True, spacing=1.1, name="Title")
    add_placeholder(lo, "subTitle", 1, CONTENT_LEFT, 3700000, 9800000, 400000,
                    SUB_TXT, 18, colour=MSO_THEME_COLOR.DARK_2, name="Subtitle")
    # The slide version set the name and role as two runs in one box. A placeholder
    # is styled per shape, so they split into two.
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 4600000, 5000000, 300000,
                    WHO, 16, bold=True, name="Speaker")
    add_placeholder(lo, "body", 3, CONTENT_LEFT, 4930000, 5000000, 300000,
                    ROLE, 13, colour=MSO_THEME_COLOR.DARK_2, name="Speaker role")
    return lo


def layout_title_bottom_banner(prs):
    """Title option C: centred colour logo on white, speaker in a bottom banner."""
    lo = new_layout(prs, LAYOUT_TITLE_BOTTOM)
    c = drawing_canvas(lo)
    band_h = 1188720
    band_y = SLIDE_H - band_h
    # Taller than 914403, so it covers the master's corner logo rather than
    # colliding with it. The centred lockup above is the only mark on the slide.
    rect(c, 0, band_y, SLIDE_W, band_h, fill=MSO_THEME_COLOR.ACCENT_1)

    logo_h = 1005840
    logo_w = int(logo_h * LOGO_ASPECT["horizontal"])
    logo(c, "horizontal", "color", (SLIDE_W - logo_w) // 2, 1100000, logo_h)

    rule_w = 914400
    rect(c, (SLIDE_W - rule_w) // 2, 3560000, rule_w, 45720,
         fill=MSO_THEME_COLOR.ACCENT_1)

    add_placeholder(lo, "ctrTitle", 0, CONTENT_LEFT, 2600000, CONTENT_W, 1300000,
                    TITLE_TXT, 40, bold=True, align="ctr", spacing=1.1, name="Title")
    add_placeholder(lo, "subTitle", 1, CONTENT_LEFT, 3900000, CONTENT_W, 400000,
                    SUB_TXT, 18, colour=MSO_THEME_COLOR.DARK_2, align="ctr",
                    name="Subtitle")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, band_y + 400000, CONTENT_W, 400000,
                    WHO, 18, bold=True, colour=MSO_THEME_COLOR.LIGHT_1, align="ctr",
                    name="Speaker")
    add_placeholder(lo, "body", 3, CONTENT_LEFT, band_y + 780000, CONTENT_W, 300000,
                    ROLE, 14, colour=MSO_THEME_COLOR.LIGHT_1, align="ctr",
                    name="Speaker role")
    return lo


def layout_insight(prs):
    """One idea, nothing else on the slide.

    No title placeholder, deliberately: the statement is the title. The point is the
    pause, so resist adding a second bullet, a chart, or a logo lockup.
    """
    lo = new_layout(prs, LAYOUT_INSIGHT)
    c = drawing_canvas(lo)
    rect(c, CONTENT_LEFT, 2200000, 1371600, 54864, fill=MSO_THEME_COLOR.ACCENT_1)
    # 300000 EMU below the rule, matching the same lockup on the closing layout
    add_placeholder(lo, "body", 1, CONTENT_LEFT, 2500000, 9144000, 1800000,
                    "Steady state is a lie. The peak is the only number that matters.",
                    40, bold=True, spacing=1.15, name="Statement")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 4400000, 9144000, 400000,
                    "Optional caption. Delete it if the line stands alone.", 16,
                    colour=MSO_THEME_COLOR.DARK_2, name="Caption")
    return lo


def layout_insight_bold(prs):
    """The loud version: full-bleed Valkey blue. Once per talk, at the turn."""
    lo = new_layout(prs, LAYOUT_INSIGHT_BOLD)
    c = drawing_canvas(lo)
    rect(c, 0, 0, SLIDE_W, SLIDE_H, fill=MSO_THEME_COLOR.ACCENT_1)
    # the fill hid the master logo, so put a white one back in the same spot
    w = int(MASTER_LOGO_H * LOGO_ASPECT["horizontal"])
    logo(c, "horizontal", "white", MASTER_LOGO_RIGHT - w, MASTER_LOGO_TOP,
         MASTER_LOGO_H)

    add_placeholder(lo, "body", 1, CONTENT_LEFT, 2600000, CONTENT_W, 1800000,
                    "You cannot provision your way out of a workload that moves.",
                    40, bold=True, colour=MSO_THEME_COLOR.LIGHT_1, align="ctr",
                    spacing=1.15, name="Statement")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 4500000, CONTENT_W, 400000,
                    "Optional caption.", 16, colour=MSO_THEME_COLOR.LIGHT_1,
                    align="ctr", name="Caption")
    return lo


def layout_big_number(prs):
    """One measurement, big enough to be the whole slide."""
    lo = new_layout(prs, LAYOUT_BIG_NUMBER)
    add_placeholder(lo, "body", 1, CONTENT_LEFT, 1800000, CONTENT_W, 2100000,
                    "3x", 130, bold=True, colour=MSO_THEME_COLOR.ACCENT_1,
                    align="ctr", spacing=1.0, name="Number")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 3950000, CONTENT_W, 500000,
                    "throughput on the same hardware", 28, bold=True, align="ctr",
                    name="Label")
    add_placeholder(lo, "body", 3, CONTENT_LEFT, 4560000, CONTENT_W, 400000,
                    "Source: valkey.io/blog/...", 16, colour=MSO_THEME_COLOR.DARK_2,
                    align="ctr", name="Source")
    return lo


LIST_ROWS = 5                   # four to six is the ceiling; five fits with air
LIST_ROW_H, LIST_ROW_GAP = 520000, 130000


def layout_list_rows(prs):
    """The default list treatment, built exactly like `slide_bullets_cards()`.

    Two flat shapes per row, no gradient: a static `ACCENT_1` bar flush left, and a
    `LIGHT_2` placeholder holding the text. A gradient fill can fake the accent edge
    inside one shape, but PowerPoint then shows the row as gradient-filled in its own
    UI, which is not what a flat card is.

    **The placeholder starts after the bar rather than spanning it.** Layout shapes
    draw beneath slide shapes, so a full-width row placeholder would cover the bar it
    is supposed to sit next to. Keeping them side by side also means the grey card
    travels with the text, so deleting a row on a slide leaves only the hairline bar
    rather than a whole empty card.
    """
    lo = new_layout(prs, LAYOUT_LIST_ROWS)
    canvas = drawing_canvas(lo)
    layout_title(lo)
    for i in range(LIST_ROWS):
        y = CONTENT_TOP + i * (LIST_ROW_H + LIST_ROW_GAP)
        rect(canvas, CONTENT_LEFT, y, BAR_H, LIST_ROW_H, fill=MSO_THEME_COLOR.ACCENT_1)
        add_placeholder(lo, "body", i + 1, CONTENT_LEFT + BAR_H, y,
                        CONTENT_W - BAR_H, LIST_ROW_H,
                        "One idea, one line", 18, anchor="ctr",
                        insets=(PAD, 0, PAD, 0),
                        fill=solid_fill(MSO_THEME_COLOR.LIGHT_2),
                        name=f"Row {i + 1}")
    return lo


def _layout_boxes(prs, name, n, fills, specs, texts=None):
    """Shared body of the Two Box and Three Box layouts.

    Card background and icon are static; heading and body are placeholders. The icon
    is baked in rather than made a picture placeholder so a new slide looks right
    immediately: swapping it means replacing the picture on the slide.
    """
    lo = new_layout(prs, name)
    c = drawing_canvas(lo)
    layout_title(lo)
    col_w, xs = box_columns(n)
    texts = texts or (CARD_TEXT,) * n
    for i, (icon, heading, body) in enumerate(specs):
        rect(c, xs[i], CONTENT_TOP, col_w, CARD_H, fill=fills[i])
        icon_r, head_r, body_r = card_geometry(xs[i], CONTENT_TOP, col_w, CARD_H)
        add_svg(c, os.path.join(ICONS, f"{icon}.svg"),
                os.path.join(PNGS, f"{icon}.png"), *icon_r, name=icon)
        add_placeholder(lo, "body", 2 * i + 1, *head_r, heading, 18, bold=True,
                        colour=texts[i], name=f"Heading {i + 1}")
        add_placeholder(lo, "body", 2 * i + 2, *body_r, body, 13, colour=texts[i],
                        spacing=1.2, name=f"Body {i + 1}")
    return lo


LOREM = ("Replace this with one supporting sentence. Keep it to a single idea so the "
         "card stays scannable from the back of the room.")


def layout_two_box(prs):
    """The generic two-up. ACCENT_5 against the brand rose.

    One fill pairing, not three: recolouring a card on a slide is two clicks, and a
    layout per palette variant is three near-identical entries in the New Slide
    gallery. The deep purple needs LIGHT_1 text and white icons, so it is not a
    drop-in swap and is deliberately not offered here.
    """
    return _layout_boxes(
        prs, LAYOUT_TWO_BOX, 2,
        (MSO_THEME_COLOR.ACCENT_5, CARD_ROSE),
        [("vertical-scaling", "Heading one", LOREM),
         ("horizontal-scaling", "Heading two", LOREM)])


def layout_three_box(prs):
    """The generic three-up, neutral grey cards."""
    return _layout_boxes(
        prs, LAYOUT_THREE_BOX, 3,
        (MSO_THEME_COLOR.LIGHT_2,) * 3,
        [("performance", "Heading one", LOREM),
         ("security", "Heading two", LOREM),
         ("reliability", "Heading three", LOREM)])


def layout_thanks(prs):
    """Closing slide on white, logo in a bottom banner. Bookends title option C."""
    lo = new_layout(prs, LAYOUT_THANKS)
    c = drawing_canvas(lo)
    band_h = 1188720
    band_y = SLIDE_H - band_h
    rect(c, 0, band_y, SLIDE_W, band_h, fill=MSO_THEME_COLOR.ACCENT_1)
    rect(c, CONTENT_LEFT, 3100000, 1371600, 54864, fill=MSO_THEME_COLOR.ACCENT_1)

    logo_h = 609600
    logo_w = int(logo_h * LOGO_ASPECT["horizontal"])
    logo(c, "horizontal", "white", (SLIDE_W - logo_w) // 2,
         band_y + (band_h - logo_h) // 2, logo_h)

    add_placeholder(lo, "body", 1, CONTENT_LEFT, 1900000, CONTENT_W, 1000000,
                    "Thank you", 54, bold=True, name="Closing")
    add_placeholder(lo, "body", 2, CONTENT_LEFT, 3400000, 6000000, 400000,
                    WHO, 18, bold=True, name="Speaker")
    for i, handle in enumerate(("@reconditerose", "github.com/madolson")):
        add_placeholder(lo, "body", 3 + i, CONTENT_LEFT, 3820000 + i * 340000,
                        6000000, 340000, handle, 16,
                        colour=MSO_THEME_COLOR.DARK_2, name=f"Contact {i + 1}")
    return lo


NEW_LAYOUTS = (layout_title_only,
               layout_title_band, layout_title_top_banner, layout_title_bottom_banner,
               layout_insight, layout_insight_bold, layout_big_number,
               layout_list_rows, layout_two_box, layout_three_box, layout_thanks)

LAYOUT_PREFIX = "Valkey "


def build_layouts(prs):
    """Add every Valkey layout to the master, then assert they are well-formed."""
    names = [build(prs).name for build in NEW_LAYOUTS]
    assert all(n.startswith(LAYOUT_PREFIX) for n in names), names
    audit(prs, names)
    print(f"added {len(names)} layouts to the slide master: {', '.join(names)}")
    return names


def slide_from_layout(prs, layout_name, texts):
    """Add a slide from one of the Valkey layouts and fill its placeholders.

    A new slide's placeholders come up empty: PowerPoint does not copy the layout's
    prompt text onto it. Filling them with the placeholder copy is what makes this a
    starter slide instead of a blank one.

    The runs deliberately name no font and no size, so they inherit the layout's
    lstStyle. That is the entire point of the layout existing, and it is the one place
    in this script where not pinning Arial is correct.
    """
    layout = prs.slide_layouts.get_by_name(layout_name)
    if layout is None:
        raise SystemExit(f"layout {layout_name!r} not in this deck")
    slide = prs.slides.add_slide(layout)
    phs = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)
    if len(phs) != len(texts):
        raise SystemExit(f"{layout_name}: {len(phs)} placeholders, {len(texts)} strings")
    for ph, text in zip(phs, texts):
        ph.text_frame.text = text
    return slide


def slide_intro(prs):
    """Opening slide. Bottom banner of the three title options, because Valkey Thank
    You was designed to bookend that one."""
    return slide_from_layout(prs, LAYOUT_TITLE_BOTTOM, [TITLE_TXT, SUB_TXT, WHO, ROLE])


def slide_thanks(prs):
    """Closing slide."""
    return slide_from_layout(
        prs, LAYOUT_THANKS,
        ["Thank you", WHO, "@reconditerose", "github.com/madolson"])


# --- primitives for feature and topology slides ---------------------------------


def chip(slide, x, y, w, h, text, fill=MSO_THEME_COLOR.ACCENT_2,
         fg=MSO_THEME_COLOR.LIGHT_1, size=12):
    """A small filled label. Stands in for a code/command token.

    The template defines no monospace font and the two-font rule says do not add
    one, so commands are set in the brand font inside a dark chip instead of a
    terminal block.
    """
    rect(slide, x, y, w, h, fill=fill)
    text_box(slide, x, y, w, h, [(text, size, True, fg)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def arrow(slide, x1, y1, x2, y2, colour=MSO_THEME_COLOR.ACCENT_1, pt=1.5,
          dashed=False, head=True):
    """Thin connector with a real arrowhead. Reads far cleaner than block arrows."""
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)),
                                   Emu(int(x2)), Emu(int(y2)))
    _set_colour(c.line.color, colour)
    c.line.width = Pt(pt)
    ln = c.line._get_or_add_ln()
    # child order matters in CT_LineProperties: prstDash before tailEnd
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    if head:
        ln.append(ln.makeelement(qn("a:tailEnd"),
                                 {"type": "triangle", "w": "med", "len": "med"}))
    return c


NODE_MARK_H = 220000     # small corner glyph, only on boxes that are a real instance


def node(slide, x, y, w, h, label, sub=None, mark=True):
    """A topology node: grey card, name, optional role line.

    `mark` puts the Valkey glyph in the upper-right corner, where these boxes used to
    carry a blue top bar. On by default because every caller of node() is a real
    Valkey instance or shard; content cards elsewhere deliberately do not get one.
    """
    rect(slide, x, y, w, h, fill=MSO_THEME_COLOR.LIGHT_2)
    text_box(slide, x, y + h // 2 - (300000 if sub else 150000), w, 300000,
             [(label, 16, True, MSO_THEME_COLOR.DARK_1)], align=PP_ALIGN.CENTER)
    if sub:
        text_box(slide, x, y + h // 2 + 20000, w, 260000,
                 [(sub, 11, False, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)
    if mark:
        mw = int(NODE_MARK_H * LOGO_ASPECT["mark"])
        logo(slide, "mark", "color", x + w - 110000 - mw, y + 110000, NODE_MARK_H)


def bar_rows(slide, x, y, w, rows, unit="", label_w=1100000, value_w=900000,
             row_h=260000, row_gap=360000, highlight=True, scale_max=None):
    """Horizontal bars for a small distribution. `rows` is [(label, value), ...].

    The largest bar gets ACCENT_1, the rest ACCENT_5, so skew is the thing the eye
    lands on.
    """
    peak = max(v for _, v in rows) or 1
    full = scale_max or peak
    track_x = x + label_w
    track_w = w - label_w - value_w - 160000
    for i, (label, value) in enumerate(rows):
        ry = y + i * (row_h + row_gap)
        text_box(slide, x, ry - 40000, label_w - 100000, 340000,
                 [(label, 11, False, MSO_THEME_COLOR.DARK_2)], anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, track_x, ry, track_w, row_h, fill=MSO_THEME_COLOR.LIGHT_2)
        hot = highlight and value == peak
        fill = MSO_THEME_COLOR.ACCENT_1 if hot else MSO_THEME_COLOR.ACCENT_5
        rect(slide, track_x, ry, max(int(track_w * value / full), 20000), row_h, fill=fill)
        text_box(slide, track_x + track_w + 160000, ry - 40000, value_w, 340000,
                 [(f"{value:,}{unit}", 11, hot,
                   MSO_THEME_COLOR.DARK_1 if hot else MSO_THEME_COLOR.DARK_2)],
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# Cores are pinned, so the process gets 400% and uses 200%. Spread across the host
# that reads as half idle; the main thread is at 99% of the single core it can use.
HOST_TOTAL, HOST_CAPACITY = 200, 400
THREAD_CPU = [("main thread", 99), ("io thread 1", 36),
              ("io thread 2", 34), ("io thread 3", 31)]

METRIC = "used_active_main_thread"


def slide_metric_compare(prs):
    """Feature emphasis by contradiction: the aggregate vs the thread that matters."""
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "The host looks idle. The main thread is full.")

    gap, pad = 400000, 300000
    col_w = (SLIDE_W - 2 * CONTENT_LEFT - gap) // 2
    col_h, top = 3400000, 1650000

    # left: the aggregate reading
    rect(slide, CONTENT_LEFT, top, col_w, col_h, fill=MSO_THEME_COLOR.LIGHT_2)
    text_box(slide, CONTENT_LEFT + pad, top + 260000, col_w - 2 * pad, 460000,
             [("host CPU", 24, True, MSO_THEME_COLOR.DARK_2)])
    text_box(slide, CONTENT_LEFT + pad, top + 900000, col_w - 2 * pad, 700000,
             [(f"{HOST_TOTAL}%", 44, True, MSO_THEME_COLOR.DARK_2),
              (f" of {HOST_CAPACITY}%", 18, False, MSO_THEME_COLOR.DARK_2)])
    bar_rows(slide, CONTENT_LEFT + pad, top + 1900000, col_w - 2 * pad,
             [("4 cores pinned", HOST_TOTAL)], unit="%", scale_max=HOST_CAPACITY,
             highlight=False, label_w=1500000, value_w=800000)
    text_box(slide, CONTENT_LEFT + pad, top + col_h - 480000, col_w - 2 * pad, 340000,
             [("Half the machine is doing nothing.", 12, False, MSO_THEME_COLOR.DARK_2)])

    # right: the same instant, per thread
    rx = CONTENT_LEFT + col_w + gap
    rect(slide, rx, top, col_w, col_h, line=MSO_THEME_COLOR.ACCENT_1)
    text_box(slide, rx + pad, top + 260000, col_w - 2 * pad, 460000,
             [(METRIC, 20, True, MSO_THEME_COLOR.ACCENT_1)])
    bar_rows(slide, rx + pad, top + 900000, col_w - 2 * pad, THREAD_CPU, unit="%",
             scale_max=100, label_w=950000, value_w=800000, row_gap=250000)
    text_box(slide, rx + pad, top + col_h - 480000, col_w - 2 * pad, 340000,
             [("One core, and it is full. That is the ceiling.", 12, False,
               MSO_THEME_COLOR.DARK_2)])
    return slide


# --- topology slides ------------------------------------------------------------
# Modernised from The Vertical Scaling wall slides 11-13. Square corners, grey node
# cards with a blue top bar, thin arrows with real arrowheads, correct slot ranges.

NODE_W, NODE_H = 3017520, 1100000


# --- how to shard: which slots actually cost you --------------------------------
# CLUSTER SLOT-STATS, Valkey 8.0. ORDERBY takes key-count, cpu-usec,
# network-bytes-in, network-bytes-out. Only key-count is reported by default;
# the rest need `cluster-slot-stats-enabled yes`.
# https://valkey.io/commands/cluster-slot-stats/
#
# The same five slots measured two ways. Key count says they are interchangeable.
SLOT_KEYS = [("slot 12703", 20450), ("slot 4488", 19980), ("slot 9021", 20310),
             ("slot 15330", 19740), ("slot 702", 20120)]
SLOT_CPU = [("slot 12703", 41200), ("slot 4488", 9800), ("slot 9021", 8100),
            ("slot 15330", 7400), ("slot 702", 6900)]


def slide_slot_stats(prs):
    """Which slot to move. The default metric hides the answer."""
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER SLOT-STATS")

    gap, pad = 400000, 300000
    col_w = (SLIDE_W - 2 * CONTENT_LEFT - gap) // 2
    col_h, top = 3800000, 1650000

    for i, (metric, rows, unit, note) in enumerate((
            ("key-count", SLOT_KEYS, "", "On by default. Every slot looks the same."),
            ("cpu-usec", SLOT_CPU, " us", "Needs cluster-slot-stats-enabled yes."))):
        x = CONTENT_LEFT + i * (col_w + gap)
        rect(slide, x, top, col_w, col_h,
             fill=MSO_THEME_COLOR.LIGHT_2 if i == 0 else None,
             line=None if i == 0 else MSO_THEME_COLOR.ACCENT_1)
        text_box(slide, x + pad, top + 260000, col_w - 2 * pad, 460000,
                 [(metric, 24, True,
                   MSO_THEME_COLOR.DARK_2 if i == 0 else MSO_THEME_COLOR.ACCENT_1)])
        bar_rows(slide, x + pad, top + 900000, col_w - 2 * pad, rows, unit=unit,
                 label_w=950000, value_w=900000, row_gap=250000, highlight=i == 1)
        text_box(slide, x + pad, top + col_h - 480000, col_w - 2 * pad, 340000,
                 [(note, 12, False, MSO_THEME_COLOR.DARK_2)])
    return slide


def slide_standalone(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "Standalone")
    ny = 2850000
    chip_w, hop = 1400000, 1200000
    gx = (SLIDE_W - (chip_w + hop + NODE_W)) // 2      # centre the whole group
    nx = gx + chip_w + hop
    chip(slide, gx, ny + (NODE_H - 400000) // 2, chip_w, 400000,
         "Clients", fill=MSO_THEME_COLOR.DARK_2)
    arrow(slide, gx + chip_w + 60000, ny + NODE_H // 2, nx - 160000, ny + NODE_H // 2)
    node(slide, nx, ny, NODE_W, NODE_H, "Primary", "all 16384 slots")
    return slide


def slide_replication(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "Replication")
    ny = 2850000
    px = 3550000
    rx = px + NODE_W + 1500000
    chip(slide, CONTENT_LEFT, ny + (NODE_H - 400000) // 2, 1400000, 400000,
         "Clients", fill=MSO_THEME_COLOR.DARK_2)
    arrow(slide, CONTENT_LEFT + 1500000, ny + NODE_H // 2, px - 160000, ny + NODE_H // 2)
    text_box(slide, CONTENT_LEFT + 1500000, ny + NODE_H // 2 - 400000, 1900000, 300000,
             [("writes", 12, False, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)

    node(slide, px, ny, NODE_W, NODE_H, "Primary", "accepts writes")
    node(slide, rx, ny, NODE_W, NODE_H, "Replica", "serves reads")
    arrow(slide, px + NODE_W + 120000, ny + NODE_H // 2, rx - 120000, ny + NODE_H // 2)
    text_box(slide, px + NODE_W, ny + NODE_H // 2 - 400000, 1500000, 300000,
             [("replicates", 12, False, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)
    return slide


def slide_clustering(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "Clustering")
    gap = 400000
    grid_w = SLIDE_W - 2 * CONTENT_LEFT
    col_w = (grid_w - 2 * gap) // 3
    slot_y, prim_y, rep_y = 1450000, 1920000, 3420000
    ranges = ["slots 0 – 5460", "slots 5461 – 10922", "slots 10923 – 16383"]

    for i, rng in enumerate(ranges):
        x = CONTENT_LEFT + i * (col_w + gap)
        text_box(slide, x, slot_y, col_w, 320000,
                 [(rng, 12, True, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)
        node(slide, x, prim_y, col_w, NODE_H, "Primary", f"shard {i + 1}")
        node(slide, x, rep_y, col_w, NODE_H, "Replica", "read + failover")
        arrow(slide, x + col_w // 2, prim_y + NODE_H + 60000,
              x + col_w // 2, rep_y - 60000, pt=1.25)
        # cluster bus: short dashed hops in the gaps only, so nothing overlaps a card
        if i < 2:
            y = prim_y + NODE_H // 2
            arrow(slide, x + col_w + 60000, y, x + col_w + gap - 60000, y,
                  colour=MSO_THEME_COLOR.DARK_2, pt=1.25, dashed=True, head=False)
    return slide


# --- slot migration ------------------------------------------------------------
# Legacy key-by-key first, then atomic slot migration (Valkey 9.0).
# https://valkey.io/blog/atomic-slot-migration/

MIG_Y, MIG_W, MIG_H = 2150000, 3200000, 1150000
SRC_X = 1300000
TGT_X = SLIDE_W - SRC_X - MIG_W
STATE_Y = MIG_Y + MIG_H + 160000

# Job states, shown as a status bar along the bottom of every migration slide.
STATES = ["Initializing", "Snapshotting", "Streaming changes", "Finalized"]
FAILED_STATES = STATES[:-1] + ["Failed"]
STRIP_Y = 5000000


def step_strip(slide, y, steps, active, active_fill=MSO_THEME_COLOR.ACCENT_1):
    """Job states with the current one(s) filled. Reads as a status bar."""
    gap = 300000
    grid_w = SLIDE_W - 2 * CONTENT_LEFT
    w = (grid_w - (len(steps) - 1) * gap) // len(steps)
    active = active if isinstance(active, (set, frozenset, list, tuple)) else {active}
    for i, label in enumerate(steps):
        x = CONTENT_LEFT + i * (w + gap)
        on = i in active
        rect(slide, x, y, w, 400000,
             fill=active_fill if on else MSO_THEME_COLOR.LIGHT_2)
        text_box(slide, x, y, w, 400000,
                 [(label, 11, on, MSO_THEME_COLOR.LIGHT_1 if on else MSO_THEME_COLOR.DARK_2)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < len(steps) - 1:
            arrow(slide, x + w + 60000, y + 200000, x + w + gap - 60000, y + 200000,
                  colour=MSO_THEME_COLOR.DARK_2, pt=1.0)


def mig_slot_label(slide, text):
    text_box(slide, CONTENT_LEFT, 1450000, SLIDE_W - 2 * CONTENT_LEFT, 340000,
             [(text, 14, True, MSO_THEME_COLOR.ACCENT_1)], align=PP_ALIGN.CENTER)


def mig_arrow(slide, label, reverse=False, colour=MSO_THEME_COLOR.ACCENT_1, dashed=False):
    y = MIG_Y + MIG_H // 2
    a, b = SRC_X + MIG_W + 140000, TGT_X - 140000
    if reverse:
        arrow(slide, b, y, a, y, colour=colour, dashed=dashed)
    else:
        arrow(slide, a, y, b, y, colour=colour, dashed=dashed)
    text_box(slide, a, y - 440000, b - a, 400000,
             [(label, 12, False, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)


def mig_pair(slide, src_state, src_fill, tgt_sub, stage_label, stage_fill):
    """Source and target shards. One geometry, reused by all six slides.

    The staging area on the target is why atomic rollback is cheap: incoming data
    lands there, separate from the live dataset, and is never announced.
    """
    node(slide, SRC_X, MIG_Y, MIG_W, MIG_H, "Shard 1", "source")
    node(slide, TGT_X, MIG_Y, MIG_W, MIG_H, "Shard 2", tgt_sub)
    chip(slide, SRC_X, STATE_Y, MIG_W, 400000, src_state, fill=src_fill)
    chip(slide, TGT_X, STATE_Y, MIG_W, 400000, stage_label, fill=stage_fill,
         fg=MSO_THEME_COLOR.DARK_1 if stage_fill == MSO_THEME_COLOR.ACCENT_5
         else MSO_THEME_COLOR.LIGHT_1)


# The legacy key-by-key protocol, kept as the "before" half of the comparison.
# Its bottom bar shows the operator's command sequence, since the old mechanism had
# no job states to report.
LEGACY_STEPS = ["SETSLOT IMPORTING", "SETSLOT MIGRATING", "GETKEYSINSLOT + MIGRATE",
                "SETSLOT NODE"]


def slide_legacy_migration(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER SETSLOT")
    mig_slot_label(slide, "slot 5461")
    mig_pair(slide, "MIGRATING", MSO_THEME_COLOR.ACCENT_2,
             "target", "IMPORTING", MSO_THEME_COLOR.ACCENT_2)
    mig_arrow(slide, "one key per MIGRATE, clients follow -ASK")
    step_strip(slide, STRIP_Y, LEGACY_STEPS, active=2)
    return slide


def slide_legacy_rollback(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER SETSLOT STABLE")
    mig_slot_label(slide, "slot 5461")
    mig_pair(slide, "still owns the slot", MSO_THEME_COLOR.ACCENT_1,
             "target", "holds keys it does not own", FAIL_RED)
    mig_arrow(slide, "keys must be moved back by hand", reverse=True, colour=FAIL_RED)
    step_strip(slide, STRIP_Y, LEGACY_STEPS, active=2, active_fill=FAIL_RED)
    return slide


def slide_slot_init(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER MIGRATESLOTS")
    mig_slot_label(slide, "slots 5461 – 10922")
    mig_pair(slide, "still owns the slots", MSO_THEME_COLOR.ACCENT_2,
             "target", "staging area created", MSO_THEME_COLOR.ACCENT_5)
    mig_arrow(slide, "slot migration link established", dashed=True)
    step_strip(slide, STRIP_Y, STATES, active=0)
    return slide


def slide_slot_migration(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER SYNCSLOTS")
    mig_slot_label(slide, "slots 5461 – 10922")
    mig_pair(slide, "still owns the slots", MSO_THEME_COLOR.ACCENT_2,
             "target", "staging area", MSO_THEME_COLOR.ACCENT_5)
    mig_arrow(slide, "snapshot, then live stream of mutations")
    step_strip(slide, STRIP_Y, STATES, active={1, 2})
    return slide


def slide_roll_forward(prs):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER SYNCSLOTS")
    mig_slot_label(slide, "slots 5461 – 10922")
    mig_pair(slide, "paused, then drops its copy", MSO_THEME_COLOR.DARK_2,
             "target", "owns the slots", MSO_THEME_COLOR.ACCENT_1)
    mig_arrow(slide, "caught up, ownership handed over")
    step_strip(slide, STRIP_Y, STATES, active=3)
    return slide


def slide_rollback(prs, red=FAIL_RED):
    slide = new_slide(prs, LAYOUT_TITLE_ONLY, "CLUSTER CANCELSLOTMIGRATIONS")
    mig_slot_label(slide, "slots 5461 – 10922")
    mig_pair(slide, "never stopped owning them", MSO_THEME_COLOR.ACCENT_1,
             "target", "staging area discarded", red)
    mig_arrow(slide, "transfer abandoned", colour=red)
    step_strip(slide, STRIP_Y, FAILED_STATES, active=3, active_fill=red)
    return slide


LOGO_FILES = ("valkey-stacked-white", "valkey-horizontal-white",
              "valkey-horizontal-color", "valkey-stacked-color")


def main():
    names = sorted(ICON_LABELS)
    missing = [n for n in names if not os.path.exists(os.path.join(ICONS, f"{n}.svg"))]
    missing += [n for n in LOGO_FILES if not os.path.exists(os.path.join(LOGOS, f"{n}.svg"))]
    if missing:
        raise SystemExit(f"missing SVGs: {missing}")

    # Preflight the donor before touching anything on disk. A template missing a
    # layout used to fail *after* the copy had already overwritten a good OUT.
    donor = Presentation(TEMPLATE)
    absent = [n for n in NEEDED_LAYOUTS if donor.slide_layouts.get_by_name(n) is None]
    if absent:
        raise SystemExit(f"{TEMPLATE} is missing required layouts: {absent}")

    make_marks()
    os.makedirs(PNGS, exist_ok=True)
    for n in names:
        rasterize(os.path.join(ICONS, f"{n}.svg"), os.path.join(PNGS, f"{n}.png"))
    for n in MARK_FILES:
        rasterize(os.path.join(ICONS, f"{n}.svg"), os.path.join(PNGS, f"{n}.png"), scale=4)
    for n in LOGO_FILES:
        # scale 2: the stacked logo is 779x670, so 4x would be needlessly large
        rasterize(os.path.join(LOGOS, f"{n}.svg"), os.path.join(PNGS, f"{n}.png"), scale=2)
    print(f"rasterized {len(names)} icon + {len(LOGO_FILES)} logo PNG fallbacks")

    # Build to a temp path and move into place only on success, so a mid-build
    # failure cannot leave OUT truncated or half-written.
    tmp_out = OUT + ".building"
    shutil.copyfile(TEMPLATE, tmp_out)
    prs = Presentation(tmp_out)
    patch_theme(prs)
    strip_slides(prs)

    # The reusable frames live in the slide master now, so they are pickable from
    # PowerPoint's New Slide gallery instead of being copied off a demo slide. The
    # demo slides that existed only to be copied are gone with them.
    build_layouts(prs)

    # --- opening slide ---
    slide_intro(prs)

    # --- feature emphasis: one metric for scaling operations ---
    slide_metric_compare(prs)

    # --- topology: standalone / replication / clustering ---
    slide_slot_stats(prs)
    slide_standalone(prs)
    slide_replication(prs)
    slide_clustering(prs)
    slide_legacy_migration(prs)
    slide_legacy_rollback(prs)
    slide_slot_init(prs)
    slide_slot_migration(prs)
    slide_roll_forward(prs)
    slide_rollback(prs)
    slide_rollback(prs, red=FAIL_RED_BRICK)   # alt: darker brick red

    # --- list slides ---
    # The generic "List placeholder" slide is gone: that shape is the Valkey List Rows
    # layout now. These three are the same real content in the three list treatments,
    # kept so the soft and icon-row alternatives can still be compared and picked.
    slide_bullets_cards(prs, V8_TITLE, V8_ITEMS, source=V8_SOURCE)    # A: row cards
    slide_bullets_soft(prs, V8_TITLE, V8_ITEMS, source=V8_SOURCE)     # B: soft list
    slide_bullets_icons(prs, V8_TITLE, V8_ITEMS, V8_ICONS, source=V8_SOURCE)  # C: icon rows

    # --- worked card examples ---
    # The lorem two-box and three-box demos are gone: those are the Valkey Two Box and
    # Valkey Three Box layouts. What is left is content, which a layout cannot carry.
    slide_two_box_calm(
        prs, "Two ways to grow (calmer palette)",
        ("vertical-scaling", "Vertical scaling",
         "Add CPU and memory to one node. Simple to operate, bounded by the largest "
         "instance you can buy, and a bigger blast radius when it fails."),
        ("horizontal-scaling", "Horizontal scaling",
         "Shard the keyspace across nodes. Scales past any single machine, at the cost "
         "of more moving parts and less headroom for bursts."),
    )
    slide_two_box(
        prs, "Two ways to grow",
        ("vertical-scaling", "Vertical scaling",
         "Add CPU and memory to one node. Simple to operate, bounded by the largest "
         "instance you can buy, and a bigger blast radius when it fails."),
        ("horizontal-scaling", "Horizontal scaling",
         "Shard the keyspace across nodes. Scales past any single machine, at the cost "
         "of more moving parts and less headroom for bursts."),
        (MSO_THEME_COLOR.ACCENT_5, CARD_ROSE),
    )

    # 5. worked three-box example
    slide_three_box(
        prs, "What we optimize for",
        [("performance", "Performance",
          "Throughput per core, and predictable tail latency under load."),
         ("hardening", "Hardening",
          "Safe defaults, fuzzing, and a shrinking surface area for untrusted input."),
         ("cost-efficiency", "Cost efficiency",
          "Same workload on smaller hardware means a smaller bill.")],
        MSO_THEME_COLOR.LIGHT_2,
    )

    # icon library: a reference sheet, not a frame, so it stays a slide
    slide_gallery(prs, names)

    # --- closing slide ---
    slide_thanks(prs)

    # Every slide is on a Valkey layout by now, so the stock layouts are dead weight:
    # 15 entries cluttering the New Slide gallery, each with its own full-bleed
    # background picture. remove() refuses any that a slide still depends on.
    dropped = strip_stock_layouts(prs, LAYOUT_PREFIX)
    print(f"removed {len(dropped)} stock layouts: {', '.join(dropped)}")

    prs.save(tmp_out)
    Presentation(tmp_out)          # reopen: if it raises, PowerPoint would complain too
    os.replace(tmp_out, OUT)
    print(f"wrote {OUT}")

    check = Presentation(OUT)
    print(f"verify: {len(check.slides)} slides, {len(check.slide_layouts)} layouts")
    for i, s in enumerate(check.slides, 1):
        pics = len(s.shapes._spTree.findall(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}pic"))
        t = s.shapes.title.text if s.shapes.title is not None else "(no title)"
        print(f"  {i}: {t!r} shapes={len(s.shapes)} pictures={pics}")

    # Every added layout must still resolve after a save/reopen cycle, and must be
    # usable: adding a slide from it is the thing a presenter actually does.
    added = [lo for lo in check.slide_layouts if lo.name.startswith("Valkey ")]
    audit(check, [lo.name for lo in added])
    for lo in added:
        s = check.slides.add_slide(lo)
        print(f"  layout {lo.name!r}: {len(lo.shapes)} shapes, "
              f"{len(lo.placeholders)} placeholders -> "
              f"slide gets {len(s.placeholders)} editable")
        check.part.drop_rel(check.slides._sldIdLst[-1].rId)
        del check.slides._sldIdLst[-1]


if __name__ == "__main__":
    main()
