# Slides Skill: Updating the Valkey Community Deck

## File Location

`slides/Valkey community deck.pptx`

## Deck Structure (48 slides)

### Section 1: Background (slides 2-6)
| Slide | Title | Update Frequency |
|-------|-------|-----------------|
| 2 | [Divider] Background | Stable |
| 3 | Introducing the Valkey project | Stable |
| 4 | Backed by industry leaders | Leave as-is (corporate supporters) |
| 5 | The Valkey project is thriving | Every release (stats) |
| 6 | Latest developments in Valkey OSS | Every major release (timeline) |

### Section 2: Feature Deep Dive (slides 7-30)
| Slide | Title | Update Frequency |
|-------|-------|-----------------|
| 7 | [Divider] Feature Deep Dive | Stable |
| 8 | Valkey 7.2 (April 2024) | Stable |
| 9 | Valkey 8.0 (September 2024) | Stable |
| 10 | Valkey 8.1 (April 2025) | Stable |
| 11 | Valkey 9.0 (October 2025) | Stable |
| 12 | Valkey 9.1 (May 2026) | Stable |
| 13 | [Divider] Valkey 8.0 | Stable |
| 14-22 | Valkey 8.0 deep dives | Stable |
| 23 | [Divider] Valkey 8.1 | Stable |
| 24-31 | Valkey 8.1 deep dives | Stable |
| 32 | What's next | Every release |
| 33 | Features under active development | Every release (roadmap) |

### Section 3: Integrations (slides 34-37)
| Slide | Title | Update Frequency |
|-------|-------|-----------------|
| 34 | [Divider] Integrations | Stable |
| 35 | Valkey module distribution | When modules change |
| 36 | Valkey Glide | Stable |
| 37 | Get involved in the project | Stable |

### Section 4: Migrations (slides 38-40)
| Slide | Title | Update Frequency |
|-------|-------|-----------------|
| 38 | [Divider] Migrations | Stable |
| 39 | Migrating from Redis OSS to Valkey | Stable |
| 40 | Strategies for migration | Stable |

Slides 41-49 are template/thank-you slides kept for layout reference.

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

### Slide 33: Features Under Active Development

Remove items that have shipped. Add items from the current roadmap. Uses paragraph levels:
- Level 0: category headers ("Reliability", "New functionality")
- Level 1: bullet items under each category

Speaker notes should explain what shipped and what's next.

### Adding a New Release Overview Slide

When a new major/minor version ships, add a summary slide after the existing release slides (currently after slide 12). Use layout index 1 ("Title + Cont"). Include a source citation at the bottom in 10pt italic.

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

```python
# ALWAYS remove from highest index first to avoid shifting issues
rId = prs.slides._sldIdLst[index].rId
prs.part.drop_rel(rId)
del prs.slides._sldIdLst[index]
```

## Release Overview Slide Style

Release overview slides (8-12) follow a consistent format. Key rules:

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

## Pitfalls and Rules

1. **Formatting is per-run, not per-paragraph.** Never overwrite `runs[0].text` with a full line if the paragraph originally had multiple styled runs. You will lose color/bold formatting. Always inspect the original run structure first.

2. **Delete slides from highest index first.** Removing a slide shifts all subsequent indices. If you delete index 5 then try to delete index 10, you're actually deleting what was index 11.

3. **Deleted slides leave orphan files in the zip.** After structural changes (add/remove/reorder), re-open and re-save the file to clean up.

4. **New slides have no pre-populated runs.** Placeholders from layouts may have empty text frames. Always use `para.add_run()` rather than assuming `para.runs[0]` exists.

5. **Verify after structural changes.** Always print the final slide list after any add/remove/reorder operation. Off-by-one errors in index math are common.

6. **Batch operations order: delete, then add, then reorder.** Do all deletions first (reverse index order), then all additions, then reorder once at the end.

## Available Slide Layouts

| Index | Name | Placeholders |
|-------|------|-------------|
| 0 | Title + Speaker | title, 3x body |
| 1 | Title + Cont | title, body |
| 2 | Divider | title only |
| 3 | Title + 3x Cont with Headings | title, 6x body |
| 4 | Title + Sub + Cont | title, 2x body |
| 7 | Title Only | title |
| 8 | Title + Subtitle | title, body |
| 9 | Title + Cont + Right Img Small | title, body, picture |
| 10 | Title + 2x Cont with Headings | title, 4x body |

## Fonts

The deck uses:
- **Play** for headers/titles (theme major font)
- **Century Gothic** for body text (theme minor font)

Do not introduce other fonts. Theme references `+mj-lt` and `+mn-lt` resolve to Play and Century Gothic respectively.

## Images

All images are raster (PNG/JPG). Zero SVGs. If replacing images, note that pptx SVG support requires newer PowerPoint versions.

## Research Sources for Updates

1. **Release info**: https://valkey.io/download/releases/
2. **Blog announcements**: https://valkey.io/blog/
3. **GitHub releases**: https://github.com/valkey-io/valkey/releases
4. **Year-end reviews**: https://valkey.io/blog/2025-year-end/ (pattern: year-end)
5. **Stats**: AWS database blog, LFX analytics, Docker Hub pull counts
6. **Release cadence**: https://valkey.io/topics/releases/ (major yearly, minor as needed)
7. **Managed providers**: Keep existing list from speaker notes unless told otherwise
