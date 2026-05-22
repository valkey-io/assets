# Slides Skill: Updating the Valkey Community Deck

## File Location

`slides/Valkey community deck.pptx`

## Deck Structure (42 slides)

### Section 1: General (slides 1-18)

High-level project overview, release summaries, integrations, and migrations.

| Slide | Title | Layout | Update Frequency |
|-------|-------|--------|-----------------|
| 1 | (Title slide) | Title + Speaker | Stable |
| 2 | [Divider] Background | Divider | Stable |
| 3 | Introducing the Valkey project | Title + Cont | Stable |
| 4 | Backed by industry leaders | Title + Cont | Leave as-is |
| 5 | The Valkey project is thriving | Title + Cont | Every release (stats) |
| 6 | Latest developments in Valkey OSS | Title + Cont | Every major release (timeline) |
| 7 | Valkey 7.2 (April 2024) | Title + Cont | Stable |
| 8 | Valkey 8.0 (September 2024) | Title + Cont | Stable |
| 9 | Valkey 8.1 (April 2025) | Title + Cont | Stable |
| 10 | Valkey 9.0 (October 2025) | Title + Cont | Stable |
| 11 | Valkey 9.1 (May 2026) | Title + Cont | Stable |
| 12 | Get involved in the project | Title + Cont | Stable |
| 13 | [Divider] Integrations | Divider | Stable |
| 14 | Valkey module distribution | Title + 3x Cont with Headings | When modules change |
| 15 | Valkey Admin | Title + Cont + Right Img Small | When Valkey Admin updates |
| 16 | [Divider] Migrations | Divider | Stable |
| 17 | Migrating from Redis OSS to Valkey | Title + Cont | Stable |
| 18 | Strategies for migration | Title + Cont | Stable |

### Section 2: Features (slides 19-33)

Feature-specific deep dives with technical detail.

| Slide | Title | Layout | Update Frequency |
|-------|-------|--------|-----------------|
| 19 | [Divider] Memory Efficiency | Divider | Stable |
| 20-26 | A new hash table (7 slides) | Title + Cont | Stable |
| 27 | [Divider] Performance | Divider | Stable |
| 28-32 | Thread scalable command processing (5 slides) | Title + Cont | Stable |
| 33 | Multithreaded performance | Title + Cont | Stable |

Slides 34-42 are template/thank-you slides kept for layout reference.

## What to Update

### Slide 5: Project Stats

Update the bullet point numbers. Sources:
- Contributors: https://github.com/valkey-io/valkey (check contributor count)
- Container pulls: https://hub.docker.com/r/valkey/valkey
- PRs: GitHub pull request count across valkey-io org
- AWS blog posts often cite consolidated stats (search "Valkey turns" or "Valkey anniversary")

Speaker notes contain the managed provider list. Per instruction, keep this as-is unless explicitly told to update.

### Slide 5: Formatting

Each bullet has two runs: the number (ACCENT_1, bold) and the description (DARK_1, not bold). When updating stats, preserve this two-run structure:

```python
from pptx.enum.dml import MSO_THEME_COLOR

# Clear existing runs from paragraph XML first
p_elem = para._p
for r in p_elem.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}r'):
    p_elem.remove(r)

# Add number run (blue, bold)
run1 = para.add_run()
run1.text = "100M+ "
run1.font.bold = True
run1.font.size = Pt(28)  # 355600 EMU
run1.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

# Add description run (black, not bold)
run2 = para.add_run()
run2.text = "container pulls"
run2.font.bold = False
run2.font.size = Pt(28)
run2.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

### Slide 6: Timeline

The timeline is a group shape containing date labels and release labels. The date text boxes are children of a group shape (Shape 1), not top-level shapes. The release label text boxes ("Valkey 8.0 Release") are separate top-level text boxes.

```python
# Access group shape children for dates:
for shape in slide.shapes:
    if shape.shape_type == 6:  # GROUP
        for child in shape.shapes:
            if child.has_text_frame:
                print(child.text_frame.text)
```

### Adding a New Release Overview Slide

When a new major/minor version ships, add a summary slide after the existing release slides (currently after slide 11). Use layout index 1 ("Title + Cont"). Include a source citation at the bottom in 10pt italic.

## How to Edit the PPTX Programmatically

The deck is a standard OOXML pptx file. Use `python-pptx` for edits.

### Setup

```bash
pip3 install python-pptx
```

### Reading Content

```python
from pptx import Presentation

prs = Presentation("slides/Valkey community deck.pptx")
slide = prs.slides[4]  # 0-indexed, so slide 5 = index 4

for shape in slide.shapes:
    if shape.has_text_frame:
        print(shape.text_frame.text)

# Speaker notes:
notes = slide.notes_slide
for shape in notes.shapes:
    if shape.placeholder_format and shape.placeholder_format.idx == 1:
        print(shape.text_frame.text)
```

### Updating Text

```python
# Update runs within existing paragraphs to preserve formatting:
for para in tf.paragraphs:
    for run in para.runs:
        if "old text" in run.text:
            run.text = run.text.replace("old text", "new text")

prs.save("slides/Valkey community deck.pptx")
```

### Adding Slides

```python
layout = prs.slide_layouts[1]  # "Title + Cont"
slide = prs.slides.add_slide(layout)  # appends to end

# Move to desired position (0-indexed):
slides = prs.slides._sldIdLst
new_elem = slides[-1]
slides.remove(new_elem)
slides.insert(target_index, new_elem)
```

### Removing Slides

**WARNING:** `drop_rel` + delete from `_sldIdLst` leaves orphan XML in the zip, causing "needs repair" errors in PowerPoint. Prefer overwriting slide content in-place or adding new slides without deleting old ones. If deletion is unavoidable, validate the output opens cleanly.

```python
# ALWAYS remove from highest index first to avoid shifting issues
rId = prs.slides._sldIdLst[index].rId
prs.part.drop_rel(rId)
del prs.slides._sldIdLst[index]
```

## Release Overview Slide Style

Release overview slides (7-11) follow a consistent format. Key rules:

### Bullet Style

Bullets use a ► caret in Valkey blue (accent1). This requires setting paragraph properties XML directly:

```python
from pptx.oxml import parse_xml

# First paragraph (no space before):
BULLET_PPR_FIRST = '''<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" marL="457200" marR="0" lvl="0" indent="-350520" algn="l" rtl="0">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="0"/></a:spcBef>
  <a:spcAft><a:spcPts val="0"/></a:spcAft>
  <a:buClr><a:schemeClr val="accent1"/></a:buClr>
  <a:buSzPts val="1920"/>
  <a:buFont typeface="Play"/>
  <a:buChar char="&#9658;"/>
</a:pPr>'''

# Subsequent paragraphs (10pt space before for breathing room):
BULLET_PPR = '''<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" marL="457200" marR="0" lvl="0" indent="-350520" algn="l" rtl="0">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="1000"/></a:spcBef>
  <a:spcAft><a:spcPts val="0"/></a:spcAft>
  <a:buClr><a:schemeClr val="accent1"/></a:buClr>
  <a:buSzPts val="1920"/>
  <a:buFont typeface="Play"/>
  <a:buChar char="&#9658;"/>
</a:pPr>'''
```

### Content Rules

- **Numbers get highlighted** in ACCENT_1 (Valkey blue) + bold. Only highlight measurable improvements (percentages, throughput numbers). Do not highlight ordinals like "1st" or "first".
- **Add `+` to numbers** where the improvement is "at least" that amount (e.g., "20%+", "1B+", "100M+").
- **Font size**: 24pt for release slides, 28pt for slide 5 stats.
- **Keep bullets short and punchy.** Lead with the number when possible.
- **Feature descriptions** (no number) use plain DARK_1 text, no bold.

### Source Citation

Source is a separate text box pinned to the bottom of the slide, aligned vertically with the Valkey logo:

```python
from pptx.util import Emu

SOURCE_LEFT = Emu(609600)
SOURCE_TOP = Emu(5943597)  # Aligned with Valkey logo Y position
SOURCE_WIDTH = Emu(9000000)
SOURCE_HEIGHT = Emu(300000)

txBox = slide.shapes.add_textbox(SOURCE_LEFT, SOURCE_TOP, SOURCE_WIDTH, SOURCE_HEIGHT)
tf = txBox.text_frame
run = tf.paragraphs[0].add_run()
run.text = "Source: valkey.io/blog/..."
run.font.size = Pt(10)
run.font.italic = True
run.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

Do NOT put the source inside the body placeholder. It floats up with the content and looks wrong.

## Custom Slide Style (Cards and Diagrams)

For slides that go beyond the standard placeholder layouts (comparison cards, architecture diagrams, etc.), build with native shapes directly. Use layout index 7 ("Title Only") or 14 ("Holding", zero placeholders) and remove any unwanted placeholders from the slide.

### Standard Coordinates

All custom slides must align with the deck's standard positioning:

```python
TITLE_LEFT = 609600
TITLE_TOP = 400019
TITLE_W = 11062800
TITLE_H = 443198
CONTENT_LEFT = 609600
CONTENT_TOP = 1445931
CONTENT_W = 11062800
CONTENT_H = 4565111
```

### Shape Rules

- Use `MSO_SHAPE.RECTANGLE` for all shapes. No rounded corners unless explicitly needed.
- Text shapes: `fill.background()`, `line.fill.background()`, margins all zero, `vertical_anchor = MSO_ANCHOR.MIDDLE`.
- Card backgrounds: `fill.solid()` with theme color, `line.fill.background()` (no border) or thin border with `line.width = Pt(0.75)`.
- Accent bars: `RECTANGLE`, height=73152 EMU, solid fill with accent color, no line.
- Always use theme colors (`MSO_THEME_COLOR.*`) instead of hardcoded RGB values.

### Theme Colors (Valkey scheme)

| Token | Hex | Use |
|-------|-----|-----|
| DARK_1 | #000000 | Primary text |
| DARK_2 | #232F3E | Secondary text, subtitles |
| LIGHT_1 | #FFFFFF | White |
| LIGHT_2 | #F1F3F3 | Card backgrounds (neutral) |
| ACCENT_1 | #6983FF | Primary accent, highlights, bars |
| ACCENT_2 | #1A2026 | Dark accent |
| ACCENT_3 | #632637 | Warm/red accent |
| ACCENT_4 | #30176E | Purple accent |
| ACCENT_5 | #BCB5E7 | Light purple (card fill) |
| ACCENT_6 | #FBD8BF | Light peach (card fill) |

### Card Pattern (Two-Column Comparison)

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR

col_gap = Emu(200000)
col_w = (CONTENT_W - col_gap) // 2
col_h = Emu(2800000)  # adjust to content
col_y = CONTENT_TOP
pad = Emu(182880)
inner_w = col_w - 2 * pad

lx = CONTENT_LEFT
rx = CONTENT_LEFT + col_w + col_gap

# Card background
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, lx, col_y, col_w, col_h)
bg.fill.solid()
bg.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_5
bg.line.fill.background()

# Top accent bar
bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, lx, col_y, col_w, Emu(73152))
bar.fill.solid()
bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
bar.line.fill.background()

# Card title (18pt bold)
# Card body (13pt regular + 12pt gray descriptions)
```

### Card Pattern (Three-Column)

```python
card_gap = Emu(200000)
card_w = (CONTENT_W - 2 * card_gap) // 3
card_h = Emu(1500000)
card_y = CONTENT_TOP

for i in range(3):
    cx = CONTENT_LEFT + i * (card_w + card_gap)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, card_y, card_w, card_h)
    bg.fill.solid()
    bg.fill.fore_color.theme_color = MSO_THEME_COLOR.LIGHT_2
    bg.line.fill.background()
```

### Removing Layout Placeholders

When using "Title Only" layout for fully custom slides, remove the title placeholder so it doesn't show an empty field:

```python
spTree = slide.shapes._spTree
for shape in list(slide.shapes):
    try:
        if shape.placeholder_format.idx == 0:
            spTree.remove(shape._element)
    except ValueError:
        pass
```

### Slide Master Limitations

python-pptx cannot add new slide layouts to the slide master. The `SlideLayouts` collection only supports `get_by_name`, `index`, and `remove`. Adding layouts requires manual OOXML manipulation.

## Images and Diagrams

**Prefer native PowerPoint shapes for diagrams.** Rectangles, connectors, and text boxes are editable, copyable, and scale cleanly. Use `add_shape()` and `add_connector()` instead of generating PNGs. When a diagram has a grouping wrapper (e.g. dashed box around related items), put the label in a separate text box positioned above the child shapes so it isn't hidden behind them.

**Use raster images (PNG/JPG) for integration/product screenshots.** Layout 9 ("Title + Cont + Right Img Small") has a picture placeholder for this. Source screenshots from the project's GitHub repo `/screenshots` directory when available.

### Inserting SVGs

python-pptx doesn't support SVG insertion natively (PIL can't read SVGs). However, SVGs can be inserted by bypassing PIL and directly creating the OOXML package parts. PowerPoint renders SVGs via the `asvg:svgBlip` extension element with a PNG fallback for older viewers.

**Critical:** The `Part` constructor signature is `Part(partname, content_type, package, blob)`. Getting the argument order wrong results in 0-byte files that show "image can't be displayed."

```python
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image
import io

def add_svg_to_slide(slide, svg_path, left, top, width, height, png_fallback_path=None):
    """
    Insert SVG with PNG fallback using the asvg:svgBlip extension.
    If no png_fallback_path provided, generates a minimal placeholder PNG.
    """
    with open(svg_path, 'rb') as f:
        svg_bytes = f.read()

    # PNG fallback (PowerPoint shows this if SVG unsupported)
    if png_fallback_path:
        with open(png_fallback_path, 'rb') as f:
            png_bytes = f.read()
    else:
        img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        png_bytes = buf.getvalue()

    slide_part = slide.part
    package = slide_part.package

    # Find next available image number
    import re
    existing = [str(p._partname) for p in package.iter_parts()
                if hasattr(p, '_partname') and '/media/' in str(p._partname)]
    max_num = 0
    for name in existing:
        m = re.search(r'image(\d+)', name)
        if m:
            max_num = max(max_num, int(m.group(1)))
    next_num = max_num + 1

    # Add parts - NOTE: arg order is (partname, content_type, package, blob)
    png_part = Part(PackURI(f'/ppt/media/image{next_num}.png'), 'image/png', package, png_bytes)
    png_rId = slide_part.relate_to(png_part, RT.IMAGE)

    svg_part = Part(PackURI(f'/ppt/media/image{next_num + 1}.svg'), 'image/svg+xml', package, svg_bytes)
    svg_rId = slide_part.relate_to(svg_part, RT.IMAGE)

    # Build pic element with SVG extension
    spTree = slide.shapes._spTree
    pic = etree.SubElement(spTree, qn('p:pic'))

    nvPicPr = etree.SubElement(pic, qn('p:nvPicPr'))
    cNvPr = etree.SubElement(nvPicPr, qn('p:cNvPr'))
    cNvPr.set('id', str(next_num + 100))
    cNvPr.set('name', 'SVG Image')
    cNvPicPr = etree.SubElement(nvPicPr, qn('p:cNvPicPr'))
    picLocks = etree.SubElement(cNvPicPr, qn('a:picLocks'))
    picLocks.set('noChangeAspect', '1')
    etree.SubElement(nvPicPr, qn('p:nvPr'))

    blipFill = etree.SubElement(pic, qn('p:blipFill'))
    blip = etree.SubElement(blipFill, qn('a:blip'))
    blip.set(qn('r:embed'), png_rId)

    extLst = etree.SubElement(blip, qn('a:extLst'))
    ext = etree.SubElement(extLst, qn('a:ext'))
    ext.set('uri', '{96DAC541-7B7A-43D3-8B79-37D633B846F1}')
    SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
    svgBlip = etree.SubElement(ext, '{%s}svgBlip' % SVG_NS)
    svgBlip.set(qn('r:embed'), svg_rId)

    stretch = etree.SubElement(blipFill, qn('a:stretch'))
    etree.SubElement(stretch, qn('a:fillRect'))

    spPr = etree.SubElement(pic, qn('p:spPr'))
    xfrm = etree.SubElement(spPr, qn('a:xfrm'))
    off = etree.SubElement(xfrm, qn('a:off'))
    off.set('x', str(left))
    off.set('y', str(top))
    ext_elem = etree.SubElement(xfrm, qn('a:ext'))
    ext_elem.set('cx', str(width))
    ext_elem.set('cy', str(height))
    prstGeom = etree.SubElement(spPr, qn('a:prstGeom'))
    prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, qn('a:avLst'))

    return svg_rId
```

## Available Slide Layouts

| Index | Name | Placeholders |
|-------|------|-------------|
| 0 | Title + Speaker | title, 3x body |
| 1 | Title + Cont | title, body |
| 2 | Divider | title only |
| 3 | Title + 3x Cont with Headings | title, 6x body |
| 4 | Title + Sub + Cont | title, 2x body |
| 5 | Thank You x1 | title, body |
| 6 | Title + 2x Speaker | title, 5x body |
| 7 | Title Only | title |
| 8 | Title + Subtitle | title, body |
| 9 | Title + Cont + Right Img Small | title, body, picture |
| 10 | Title + 2x Cont with Headings | title, 4x body |
| 11 | Quote | title, 2x body |
| 12 | Code | body |
| 13 | Thank You 2x | title, 3x body |
| 14 | Holding | (none) |
| 15 | Full Image | picture |
| 16 | Video | body |

## Fonts

The deck uses:
- **Play** for headers/titles (theme major font)
- **Century Gothic** for body text (theme minor font)

Do not introduce other fonts. Theme references `+mj-lt` and `+mn-lt` resolve to Play and Century Gothic respectively.

## Pitfalls and Rules

1. **Formatting is per-run, not per-paragraph.** Never overwrite `runs[0].text` with a full line if the paragraph originally had multiple styled runs. You will lose color/bold formatting. Always inspect the original run structure first.

2. **Delete slides from highest index first.** Removing a slide shifts all subsequent indices. If you delete index 5 then try to delete index 10, you're actually deleting what was index 11.

3. **Deleted slides leave orphan files in the zip.** After structural changes (add/remove/reorder), re-open and re-save the file to clean up.

4. **New slides have no pre-populated runs.** Placeholders from layouts may have empty text frames. Always use `para.add_run()` rather than assuming `para.runs[0]` exists.

5. **Verify after structural changes.** Always print the final slide list after any add/remove/reorder operation. Off-by-one errors in index math are common.

6. **Batch operations order: delete, then add, then reorder.** Do all deletions first (reverse index order), then all additions, then reorder once at the end.

7. **Use theme colors, not hardcoded RGB.** Ensures slides adapt if the theme changes. Use `MSO_THEME_COLOR.*` constants.

8. **No rounded corners on custom shapes.** Use `MSO_SHAPE.RECTANGLE` not `ROUNDED_RECTANGLE` for the clean card style.

9. **Remove unused placeholders.** When using "Title Only" or other layouts for custom slides, remove placeholder elements from `_spTree` so empty fields don't show.

10. **Align to standard coordinates.** Custom slides must use the same title position (left=609600, top=400019) and content start (top=1445931) as placeholder-based slides.

## Research Sources for Updates

1. **Release info**: https://valkey.io/download/releases/
2. **Blog announcements**: https://valkey.io/blog/
3. **GitHub releases**: https://github.com/valkey-io/valkey/releases
4. **Year-end reviews**: https://valkey.io/blog/2025-year-end/ (pattern: year-end)
5. **Stats**: AWS database blog, LFX analytics, Docker Hub pull counts
6. **Release cadence**: https://valkey.io/topics/releases/ (major yearly, minor as needed)
7. **Managed providers**: Keep existing list from speaker notes unless told otherwise
