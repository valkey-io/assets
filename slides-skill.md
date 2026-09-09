# Slides Skill: Authoring Valkey-Themed Slides

How to build and edit slides in the Valkey template. This file is about the template and the
technique, not about any one deck. Deck-specific notes live next to the deck in `slides/*.md`.

All facts below were verified against the four decks in `slides/` (Valkey community deck, The
Vertical Scaling wall, Percona Caching Deck, Percona Live 2026 Keynote). The theme is identical in
all four, so anything in the Template section is safe to rely on.

## Before You Touch a Deck

1. **Check for a lock file.** `ls slides/` and look for `~$<name>.pptx`. That means PowerPoint has
   the file open. Your `prs.save()` will be silently clobbered the next time the user hits save,
   and PowerPoint may write its own stale in-memory copy over the file at any moment. Ask them to
   close it first, and do not use an open deck as a template donor.
2. **These pptx files are almost all untracked in git.** Only `Valkey community deck.pptx` is
   committed. There is no undo for the others. Before any destructive step, copy the file aside, and
   consider suggesting the user commit first.
3. **Dump the deck before editing.** Walk every shape, position, run, chart series and speaker
   note, and read it back. Never edit a slide you have not dumped: run structure and shape names
   are not guessable.
4. **Copy to `/tmp` and test there first** for anything structural (adding, deleting, reordering
   slides, replacing chart data). Save, reopen, re-dump, confirm, then apply to the real file.

Requires `python-pptx` (`pip3 install python-pptx`), tested at 1.0.2, plus Pillow and Google
Chrome for the icon and rendering tooling.

## Tooling in This Repo

| Script | What it does |
|--------|-------------|
| `slides/pptx_svg.py` | `add_svg(slide, svg, png_fallback, left, top, w, h)` to embed an SVG. |
| `slides/pptx_layout.py` | Add slide layouts to the master, which python-pptx cannot do. See Adding a Layout to the Master. |
| `slides/build_placeholder_deck.py` | Re-runnable builder for the placeholder deck. Good worked example of cards, icon embedding, and building a deck from scratch. |
| `slides/icons/` | 16 reusable Valkey-blue outline icons, plus 4x PNG fallbacks in `icons/png/`. |

## Template Facts

### Canvas

16:9, `12192000 x 6858000` EMU (13.333in x 7.5in). 914400 EMU per inch, 12700 per point.

### Theme Colors

Always use `MSO_THEME_COLOR.*`, never hardcoded RGB. Identical across all decks.

| Token | Hex | Use |
|-------|-----|-----|
| `DARK_1` | `#000000` | Primary text |
| `DARK_2` | `#232F3E` | Secondary text, subtitles, column headings |
| `LIGHT_1` | `#FFFFFF` | White; text on dark backgrounds (dividers, title, thank-you) |
| `LIGHT_2` | `#F1F3F3` | Neutral card background |
| `ACCENT_1` | `#6983FF` | **Valkey blue.** Bullets, highlighted numbers, accent bars, links |
| `ACCENT_2` | `#1A2026` | Dark accent |
| `ACCENT_3` | `#632637` | Dark maroon. **Reads brown, not red** — see below |
| `ACCENT_4` | `#30176E` | Purple accent |
| `ACCENT_5` | `#BCB5E7` | Light purple card fill |
| `ACCENT_6` | `#E0A2AF` | Brand rose, Pantone 494 C. **Patched in**, replacing the stock peach `#FBD8BF` |

Hyperlink and followed-hyperlink are both `#6983FF`.

**Official card colours, and where they live in the theme:**

| Colour | Hex | RGB | CMYK | Pantone | Theme slot |
|---|---|---|---|---|---|
| Rose | `#E0A2AF` | 224 162 175 | 4 45 15 0 | 494 C | `ACCENT_6`, patched in |
| Deep purple | `#30176E` | 48 23 110 | 95 100 20 14 | 2685 C | `ACCENT_4`, already correct |
| Card text | `#1A2026` | 26 32 38 | | | `ACCENT_2`, already correct |

**Prefer patching the theme over hardcoding RGB.** The stock `accent6` was peach; rewriting that one
scheme entry makes the rose a real theme colour, so it shows up in PowerPoint's colour picker and
every reference stays a theme reference. python-pptx does not parse the theme part, so edit the blob:

```python
for part in prs.part.package.iter_parts():
    if str(part.partname) == "/ppt/theme/theme1.xml":
        xml = part.blob.decode("utf-8")
        xml, n = re.subn(r'(<a:accent6>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
                         r"\g<1>E0A2AF\g<2>", xml, count=1)
        part._blob = xml.encode("utf-8")
```

Anchor the regex on the `clrScheme` entry and assert exactly one substitution: the same hex may
appear as an explicit `srgbClr` elsewhere, and those are not yours to change.

On the deep purple, `#1A2026` text is unreadable — pass `LIGHT_1`, and note the blue concept icons
also lose contrast on it. A dark fill needs white icon variants, which do not exist yet.

**The palette has no red.** `ACCENT_3` is the closest and it reads as brown, which is wrong for a
failure or abort state. This is the one documented exception to theme-colours-only: use an explicit
`RGBColor` for failure states. `#D32F2F` reads clearly red next to the Valkey blue; `#C0392B` is a
darker brick alternative. Keep the exception to failure semantics, and do not start hand-picking
RGB for anything the palette already covers.

Supporting both means a small shim, since python-pptx has separate setters:

```python
def _set_colour(target, colour):
    if isinstance(colour, RGBColor):
        target.rgb = colour
    else:
        target.theme_color = colour
```

Route every colour assignment through it: run fonts, shape fills, shape lines, **and connector
lines**. Missing one raises `ValueError: (211, 47, 47) is not a valid MSO_THEME_COLOR_INDEX`, which
is the giveaway that a code path still assumes a theme colour.

For two-way comparisons, `ACCENT_5` vs `ACCENT_6` reads clearly and stays on-brand. Reserve
`ACCENT_1` for the thing you want the eye to land on. If everything is blue, nothing is.

### Fonts

The theme fonts are:

- **Play** — theme major font, `+mj-lt`. The master applies it to titles *and* to the body
  placeholder, which is a quirk worth knowing: body text is not Century Gothic by default.
- **Century Gothic** — theme minor font, `+mn-lt`.

**Runs pin Arial.** `Arial` is the standard for text, set explicitly with `run.font.name` on every
run rather than left to inherit. Community deck slides 7-8 already do this, so the convention
matches the real decks. In `build_placeholder_deck.py` it is two constants, `FONT` and
`TITLE_FONT`, so headings can be moved back to Play in one line if you want that split.

Three fonts total, and no more: Play and Century Gothic in the theme, Arial on runs. Anything else
risks not being installed on the presenting machine — that rules out Open Sans and the rest of the
Google fonts unless you are certain about the room's laptop.

**Set the font on titles too, not just body.** The classic symptom is a deck where most titles look
right and a couple look different: those slides had body runs pinning Arial while their titles
inherited Play. Audit it rather than trusting it — collect `run.font.name` across every title and
assert the set has exactly one member.

### Type Scale

Not from the master, despite where it looks like it lives: `slideMaster1.xml`'s `titleStyle` and
`bodyStyle` are 14pt across all four levels. Every size below comes from an `a:lstStyle` on a
*layout* placeholder, which the slide's placeholder then inherits. That matters when you add a
layout: putting the sizes in its `lstStyle` is not a workaround, it is the same mechanism the stock
layouts use.


| Role | Size | Weight | Color |
|------|------|--------|-------|
| Slide title | 32pt | bold | `DARK_1` |
| Body level 1 | 24pt | regular | `DARK_1` |
| Body level 2 | 20pt | regular | `DARK_1` |
| Body level 3 | 18pt | regular | `DARK_1` |
| Divider / Quote / title-slide title | 44pt | regular | `LIGHT_1` |
| Column heading (`… with Headings` layouts) | 20pt | bold | `DARK_2` |
| Column body (`… with Headings` layouts) | 18pt | regular | inherit |
| Source citation | 10pt | italic | `DARK_1` |

Two things the table does not capture:

- **Leave placeholder titles unstyled** so they inherit 32pt. Most title placeholders across the
  decks set no explicit run size. The overrides that exist are 48pt (opening title slides) and
  28pt.
- **Hand-built title shapes use 28pt bold**, consistently, on every custom slide in every deck.
  When you fake a title with a rectangle on a `Title Only` or `Holding` slide, use 28pt bold at
  the standard title coordinates, not 32pt.

Card text does not have one house size; it scales to the card. Observed: a small 14pt bold eyebrow
label over 24pt bold rows for a full-width two-card comparison, and 18pt bold titles with 13pt
body for denser three-column cards. Pick sizes for the density you have and keep the two cards
identical.

### Standard Geometry

Every slide must line up on this grid. Content is inset 609600 EMU (0.667in) from the left.

```python
TITLE_LEFT   = 609600
TITLE_TOP    = 400019
TITLE_W      = 11062800
TITLE_H      = 443198

CONTENT_LEFT = 609600
CONTENT_TOP  = 1445931
CONTENT_W    = 11062800
CONTENT_H    = 4565111
```

The Valkey logo sits in the master at `(10389732, 5943597)`, size `1434431 x 443198`. Anything
pinned to the bottom of a slide (source lines, footnotes) should align to `y = 5943597` so it sits
on the same baseline as the logo.

**The content box is not symmetric.** Left inset is `609600` but width is `11062800`, so it
overhangs the right margin by `90000` EMU (0.1in). Text does not reveal this, but a grid or a row
of evenly divided cells does: the last column sits closer to the right edge than the first does to
the left. For grids, centre on the slide instead:

```python
grid_w = SLIDE_W - 2 * CONTENT_LEFT      # 10972800, symmetric
cell_w = grid_w // cols
grid_left = CONTENT_LEFT
```

The same applies vertically. `CONTENT_TOP + CONTENT_H` bottoms out at `6011042`, so a block anchored
to `CONTENT_TOP` sits low: its centre lands at `3728486` against a page centre of `3429000`, a third
of an inch off. For a slide that is mostly one block, measure the block and centre it on the page:

```python
block_h = (rows - 1) * row_pitch + row_content_h
grid_top = (SLIDE_H - block_h) // 2
```

Then assert it: the top gap and bottom gap should come out equal. Eyeballing will not catch 0.33in.

## Choosing a Layout

**Never hardcode a layout index.** Indices differ in every deck: `Title Only` is index 6 in the
community deck, 8 in The Vertical Scaling wall, 4 in the Percona Caching deck, 2 in the Percona
keynote. **The placeholder deck has none of the stock layouts at all** — its master is the eleven
Valkey layouts and nothing else. Resolve by name:

```python
layout = prs.slide_layouts.get_by_name("Title + Cont")
if layout is None:
    raise SystemExit("layout missing from this deck")
slide = prs.slides.add_slide(layout)
```

Layout names are consistent where they exist, but not every deck has every layout. Check first.

### Stock Layouts

Present in the community deck, The Vertical Scaling wall, and both Percona decks. **Not in the
placeholder deck**, which replaced them; see The Valkey Layouts.

| Layout | Placeholders | Use it for |
|--------|-------------|------------|
| `Title + Cont` | title(0), body(1) | The default. Title plus bulleted content. |
| `Title Only` | title(0) | Custom slides: diagrams, cards, anything hand-built. |
| `Divider` | title(0), 44pt on dark | Section breaks. |
| `Title + Sub + Cont` | title(0), body(1), sub(2, 11pt bold) | Title needing a kicker line above the content. |
| `Title + Subtitle` | title(0), body(1) | Title with a one-line subtitle, no body. |
| `Title + 2x Cont with Headings` | title(0), heading(1,3), body(2,4) | Two-column comparison with real headings. |
| `Title + 3x Cont with Headings` | title(0), heading(1,3,5), body(2,4,6) | Three-column comparison. |
| `Title + Cont + Right Img Small` | title(0), body(1), picture(2) | Screenshot or product shot on the right. Picture placeholder is full-bleed right: `(7884543, 0)` size `4307456 x 6858000`. |
| `Quote` | title(0, 44pt), body(1,2) | Pull quote with attribution. |
| `Code` | title(0) only | Code slides. Add the code as your own text box; there is no body placeholder. |
| `Full Image` | picture(2), full bleed | Edge-to-edge image. |
| `Video` | media(2), full bleed | Embedded video. |
| `Title + 2x Speaker` | title(0), body(1–5) | Opening slide, up to two presenters. |
| `Thank You 2x` | body(1–4), no title | Closing slide, up to two presenters. |
| `Holding` / `DEFAULT` | none | Blank canvas for fully custom work. |

Prefer the `… with Headings` layouts over hand-built cards when the content is genuinely a
two- or three-column comparison. They inherit the type scale for free.

### The Valkey Layouts

`build_placeholder_deck.py` builds eleven layouts and then **deletes all fifteen stock ones**, so the
placeholder deck's master is only these. They are the designs this repo used to hand-draw onto a
slide every time. All are white-background, all put decoration in the layout and text in
placeholders.

| Layout | Placeholders | Use it for |
|--------|-------------|------------|
| `Valkey Title Only` | title | The foundation, not a design. Every custom slide (diagrams, anything hand-built) sits on this. Replaces the stock `Title Only`. |
| `Valkey Title Band` | title, subtitle, speaker, role | Title slide, blue band down the right third. |
| `Valkey Title Top Banner` | title, subtitle, speaker, role | Title slide, navy banner across the top. |
| `Valkey Title Bottom Banner` | title, subtitle, speaker, role | Title slide, centred lockup, speaker in a blue bottom banner. Bookends `Valkey Thank You`. |
| `Valkey Insight` | statement, caption | One assertion on white under an accent rule. No title: the statement is the title. |
| `Valkey Insight Bold` | statement, caption | Same, full-bleed Valkey blue. Once per talk. |
| `Valkey Big Number` | number, label, source | One measurement at 130pt. |
| `Valkey List Rows` | title, 5 rows | The default list treatment. Each row is one placeholder, so unused rows delete cleanly. |
| `Valkey Two Box` | title, 2x (heading, body) | Two-up cards, `ACCENT_5` against the rose, icons baked in. |
| `Valkey Three Box` | title, 3x (heading, body) | Three-up cards, `LIGHT_2`. |
| `Valkey Thank You` | closing, speaker, 2 contacts | Closing slide. |

One layout per design, not one per palette variant. Recolouring a card on a slide is two clicks;
three near-identical gallery entries is clutter. The deep purple is the exception worth knowing
about and it is deliberately absent: it needs `LIGHT_1` text and white icons, so it is not a fill
swap.

### Removing the Stock Layouts

```python
dropped = strip_stock_layouts(prs, "Valkey ")   # after every slide is on one of yours
```

Order is not optional: **build the replacements and every slide first, delete last.**
`SlideLayouts.remove()` raises `ValueError` on a layout that is the basis for a slide, and that
refusal is the safety net. If it fires, a slide is still on a stock layout and the fix is to rebuild
it on one of yours, not to skip the delete. This is the whole reason `Valkey Title Only` exists: 19
diagram slides were on the stock `Title Only`.

Removal drops the master→layout relationship, so each layout part and any media only it referenced
falls out of the package on save. The stock layouts each carry a full-bleed background picture, so
this is most of why the placeholder deck went from 894 KB to 541 KB. It also removes the four shapes
in the deck that overhung the canvas, all of them Google-Slides-exported backgrounds in stock
layouts.

Two things to know afterwards:

- **Pasting a slide in from another Valkey deck drags its layout back in.** PowerPoint brings the
  source layout along, so the master grows a `Title + Cont` again. Re-running the builder is what
  cleans that up; nothing detects it automatically.
- **Assert it, do not assume it.** A removed layout can leave an orphaned part in the zip if the
  relationship graph is wrong. Compare the count of `ppt/slideLayouts/slideLayoutN.xml` entries in
  the package against `len(prs.slide_layouts)`, and check that every slide's
  `slide_layout.name` is one of yours.

## Adding a Layout to the Master

python-pptx cannot: `SlideLayouts` exposes only `get_by_name`, `index` and `remove`. That is the
whole reason this repo used to hand-draw reusable frames onto slides. `slides/pptx_layout.py`
closes the gap by building the part directly.

```python
from pptx_layout import new_layout, drawing_canvas, add_placeholder, audit

layout = new_layout(prs, "Valkey Two Box")     # empty, registered on the master
canvas = drawing_canvas(layout)                # slide-shaped, so the slide helpers work
rect(canvas, x, CONTENT_TOP, w, h, fill=MSO_THEME_COLOR.ACCENT_5)
add_placeholder(layout, "body", 1, x, y, w, h, "Heading one", 18, bold=True)
audit(prs, ["Valkey Two Box"])
```

`drawing_canvas` is the piece that saves the work: `layout.shapes` is a read-only `LayoutShapes`,
but `SlideShapes` only needs an spTree and a parent with a `.part`, so wrapping it lets `rect()`,
`text_box()`, `logo()` and `add_svg()` target a layout unchanged.

Three things have to line up or PowerPoint reports a repair, and it is always one of these:

1. The layout part exists with content type `PML_SLIDE_LAYOUT`.
2. A relationship in **each** direction between layout and master. The layout→master rel is what
   makes `SlideLayout.slide_master` resolve; the master→layout rel is what the id list points at.
3. A `p:sldLayoutId` entry in the master's `sldLayoutIdLst`, with an `id` in
   `2147483648 .. 4294967295`.

`SlideLayoutPart.load`'s signature is `(partname, content_type, package, blob)` — the same argument
order trap as `Part()` in `pptx_svg.py`, and it fails deep inside lxml rather than where you made
the mistake.

### What Belongs in a Layout, and What Cannot

- **Decoration goes in the layout, text goes in placeholders.** Static layout shapes are inherited
  and drawn on the slide but cannot be selected there. Anything a presenter edits has to be a
  placeholder or it is unreachable.
- **The layout's placeholder text is prompt text, not content.** A new slide comes up empty and
  correctly styled; PowerPoint does not copy the layout's words. Set `hasCustomPrompt="1"` and the
  prompt replaces "Click to add text", which is worth doing: "Heading one" beats the generic.
- **Per-run formatting cannot live in a layout.** A blue highlighted number is two runs with
  different colours; `lstStyle` styles a whole level. Keep that in the generator.
- **Static layout shapes draw *beneath* slide shapes.** A placeholder therefore cannot sit on top of
  layout decoration it is meant to sit beside: a full-width row placeholder covers the accent bar at
  its left edge. Inset the placeholder past the static shape instead. In the list layout the bar is
  static at `CONTENT_LEFT` width `BAR_H`, and the row placeholder starts at `CONTENT_LEFT + BAR_H`.
- **Do not fake a one-sided accent with a gradient fill.** A DrawingML outline has no per-side
  control, so a two-stop `a:gradFill` looks like a clever way to get an accent edge inside a single
  shape, and it is wrong twice: PowerPoint then shows the shape as gradient-filled in its own UI,
  which is not what a flat card is, and the hard stop it relies on is fragile (pitfall 12). Use two
  shapes. `solid_fill()` gives the placeholder its flat card, and the accent is its own static shape.
- **What a placeholder's own fill buys you is deletion.** The card travels with the text, so removing
  a row on a slide takes its background with it and leaves only the hairline bar, not an empty card.
- **Sizes go in `a:lstStyle`,** not on the layout's own runs. See Type Scale.
- `pptx_layout.audit()` catches the failures that are silent otherwise: an unregistered layout, a
  layout that does not resolve back to the master, duplicate `(type, idx)` pairs, duplicate shape
  ids, and a layout with no placeholders at all. Run it before saving.

## Writing the Content

Design rules that matter more than any of the code below:

- **One idea per slide.** If a slide needs two sentences to summarize, it is two slides.
- **Titles are assertions, not labels.** "Throughput is a cliff" beats "Throughput". State the
  claim so the slide has a point even if you say nothing.
- **Build complex diagrams across slides.** Duplicate the slide and add one element per step. This
  is the house style for architecture and threading diagrams. Keep builds in strictly increasing
  order: every slide adds, none removes.
- **Bullets are short and punchy.** Lead with the number. Aim for one line each; four to six per
  slide.
- **Keep the text minimal. This is the rule that gets broken most.** A slide is a visual aid, not a
  document. Explanatory paragraphs under a diagram do not get read: the audience is either reading
  or listening, never both. If a diagram needs a paragraph to make sense, fix the diagram or put the
  prose in the speaker notes. Prefer a label on the shape over a caption below it, and a plain title
  over a title plus a subtitle plus a caption. When in doubt, cut.
- **Do not decorate labels.** A short piece of text does not need a coloured box behind it. Use the
  real title placeholder as the header rather than inventing a filled chip above the content.
- **No accent stripe on top of a filled card.** A card's fill already defines it; a thin bar across
  its top edge is pure decoration, and it is the single most recognisable tell of a machine-made
  slide. Same for node cards in diagrams. Leave the card plain.

  Accents that are *not* a line on top of a box are fine and worth keeping: a coloured left edge on
  a list row, a full-width rule marking the boundary of a banner, a short rule separating a title
  from a subtitle, hairline dividers between rows. The distinction the reviewer drew, and it is a
  good one, is the horizontal stripe sitting on a box's top edge.

  To audit: for each thin shape, check whether a larger shape shares its exact left, top and width.
  If so it is a line on top of a box. Delete those; leave the rest.
- **Cite your numbers.** Any chart or stat gets a source line.

### Bullet Style

Valkey bullets are a ► caret in Valkey blue. python-pptx cannot set bullet characters through its
API, so set the paragraph properties XML directly.

```python
from pptx.oxml import parse_xml

BULLET_PPR = '''<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    marL="457200" marR="0" lvl="0" indent="-350520" algn="l" rtl="0">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="{spc}"/></a:spcBef>
  <a:spcAft><a:spcPts val="0"/></a:spcAft>
  <a:buClr><a:schemeClr val="accent1"/></a:buClr>
  <a:buSzPts val="1920"/>
  <a:buFont typeface="Play"/>
  <a:buChar char="&#9658;"/>
</a:pPr>'''

# spc="0" on the first paragraph, spc="1000" (10pt) on the rest for breathing room.
para._p.insert(0, parse_xml(BULLET_PPR.format(spc=0)))
```

The template's own `marL=457200 indent=-228600` is the inherited default; the caret style uses
`indent=-350520` to fit the wider glyph. Verified against community deck slides 7-8, which are the
reference for this style.

**Know when not to use it.** The caret style is raw text on white. In a deck built from cards and
filled shapes it is the only slide type with no fill and no structure, and it reads as harsh and
out of place. In that context use a row treatment instead: one `LIGHT_2` row per item with a
`73152` EMU `ACCENT_1` left edge, text at 18pt, vertically centred. Same information, same
highlight rule, and it borrows the deck's existing vocabulary. See `slide_bullets_cards()` in
`build_placeholder_deck.py`. Reach for carets when the surrounding slides are also text-led, as in
the community deck.

Those community deck runs pin `latin typeface="Arial"`, overriding the theme font. That is now the
house convention rather than a quirk to avoid: see Fonts.

### Highlighting Numbers

Measurable improvements get `ACCENT_1` + bold; the rest of the line stays `DARK_1` regular. This
is the single most recognizable thing about the deck style.

```python
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt

r1 = para.add_run()
r1.text = "20%+ "
r1.font.bold = True
r1.font.size = Pt(24)
r1.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

r2 = para.add_run()
r2.text = "faster on the same hardware"
r2.font.bold = False
r2.font.size = Pt(24)
r2.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

Rules:
- Highlight percentages, throughput, latency, counts. **Do not** highlight ordinals ("1st",
  "first") or version numbers.
- Append `+` when the improvement is "at least" that much: `20%+`, `100M+`, `1B+`.
- One highlighted number per bullet. Two competes with itself.

### Source Citations

Put the source in its own text box pinned to the logo baseline. **Never** put it in the body
placeholder; it floats up under the last bullet and looks broken.

```python
from pptx.util import Emu, Pt

box = slide.shapes.add_textbox(Emu(609600), Emu(5943597), Emu(9000000), Emu(300000))
run = box.text_frame.paragraphs[0].add_run()
run.text = "Source: valkey.io/blog/..."
run.font.size = Pt(10)
run.font.italic = True
run.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

## Custom Slides

For anything no layout covers, start from `Title Only` (or `Holding` / `DEFAULT` for no title at
all) and build with native shapes. Native shapes stay editable and scale cleanly; never render a
diagram to PNG.

**Before you build custom, ask whether it should be a layout instead.** If the design will be used
more than once, hand-drawing it onto each slide is the wrong answer: add it to the master. Custom
slides are for content that only exists once, which in practice means diagrams.

### Shape Rules

- `MSO_SHAPE.RECTANGLE` for cards and callouts. No rounded corners.
  (Older diagram slides use `ROUNDED_RECTANGLE`. Match the slide you are editing, but square is
  the house style for new work.)
- Text-only shapes: `fill.background()`, `line.fill.background()`, all four margins zero,
  `vertical_anchor = MSO_ANCHOR.MIDDLE`.
- Card backgrounds: `fill.solid()` with a theme color, `line.fill.background()` for no border, or
  a hairline with `line.width = Pt(0.75)`.
- Accent bars: `RECTANGLE`, height `73152` EMU, solid accent fill, no line.
- Card padding: `182880` EMU (0.2in) inside the card edge.
- Hand-built titles: 28pt bold at `(609600, 400019)`, size `11062800 x 443198`.
- Remove placeholders you are not using, or PowerPoint shows empty "Click to add title" fields:

```python
spTree = slide.shapes._spTree
for shape in list(slide.shapes):
    try:
        if shape.placeholder_format.idx == 0:
            spTree.remove(shape._element)
    except (ValueError, AttributeError):
        pass
```

### Two-Column Cards

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Emu

gap  = 200000
col_w = (CONTENT_W - gap) // 2
col_h = 2800000                      # adjust to content
lx, rx = CONTENT_LEFT, CONTENT_LEFT + col_w + gap

for x, fill in ((lx, MSO_THEME_COLOR.ACCENT_5), (rx, MSO_THEME_COLOR.ACCENT_6)):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, CONTENT_TOP, col_w, col_h)
    bg.fill.solid()
    bg.fill.fore_color.theme_color = fill
    bg.line.fill.background()

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, CONTENT_TOP, col_w, 73152)
    bar.fill.solid()
    bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    bar.line.fill.background()
```

Three columns: `card_w = (CONTENT_W - 2 * gap) // 3`, same pattern.

Label cards with what they *are*. "State A" / "State B" tells the audience nothing; "Steady state"
/ "Rebalancing" does.

### Diagrams

- Use `add_shape()` + `add_connector()`. Connectors with `size=(0, h)` are vertical lifelines;
  `size=(w, 0)` are horizontal handoffs.
- For a dashed grouping box around related shapes, put the label in a **separate text box above**
  the children, not in the box itself, or it renders behind them.
- Overlay-based animation (stacking `req` and `retry` boxes at identical coordinates on one slide)
  is fragile and confusing to edit. Prefer one slide per state.

## Charts

Charts carry their own embedded workbook. Edit them through the chart API, not by touching text
runs. `replace_data` is verified working for both category and XY charts.

```python
from pptx.chart.data import CategoryChartData, XyChartData

chart = next(s.chart for s in slide.shapes if getattr(s, "has_chart", False) and s.has_chart)

# Category charts (LINE, LINE_MARKERS, BAR, COLUMN, ...)
cd = CategoryChartData()
cd.categories = ["10K", "20K", "30K"]
cd.add_series("p50", (167.0, 203.0, 228.0))
chart.replace_data(cd)

# XY scatter charts
xd = XyChartData()
s = xd.add_series("P50 Latency")
for x, y in [(0, 25.0), (1, 28.0), (2, 25.0)]:
    s.add_data_point(x, y)
chart.replace_data(xd)
```

**`replace_data` replaces the whole series set.** If you pass one series to a three-series chart,
you now have a one-series chart, and the deleted series' formatting is gone. Always re-add every
series you want to keep. Dump the chart first to get the existing series names and values.

Category labels can be empty strings when the shape of the curve is the point and the absolute
values are not. That is a legitimate choice, but label the axes.

## Images and SVGs

Raster (PNG/JPG) for screenshots and product shots; `Title + Cont + Right Img Small` has a picture
placeholder for exactly this. Native shapes for everything diagram-shaped.

### Brand Logos

Eight official logo SVGs live in `Valkey Branding/logo svgs/`:
`valkey-{horizontal,stacked}-{black,white,color,color-light}.svg`.

- **horizontal** aspect ratio `187.9 / 63.5` = **2.959**
- **stacked** aspect ratio `778.6 / 669.9` = **1.162**

Always derive width from height using those ratios; never type two independent numbers. Use the
`white` variants on `ACCENT_1` / `ACCENT_2` fills and `color` on white. The `logo()` helper in
`build_placeholder_deck.py` takes a height and does the arithmetic.

### White-Background Title Slides

The slide master's background is already white (`lt1`). The stock title-ish layouts
(`Title + 2x Speaker`, `Divider`, `Thank You 2x`) look dark because the *layout* carries a
full-bleed background **picture**, which is why their placeholder text is `LIGHT_1`. You cannot
make them white without covering that image up.

For a white title slide, build custom on **`Holding`**: zero placeholders, no background picture,
so the master's white shows through and you have a clean canvas. `Title Only` also works if you
want a real title placeholder.

**Every layout inherits the master's shapes** (`showMasterSp` is unset everywhere), so the Valkey
logo at `(10389732, 5943597)` renders bottom-right on every slide. Two consequences:

- A full-bleed fill **covers** it, because slide shapes draw over master shapes. Put a white
  horizontal logo back at the same spot: height `443198`, right edge `11824163`.
- A bottom banner taller than `914403` also covers it, which is usually what you want on a title or
  closing slide. If you add your own logo elsewhere on a slide that still shows the corner mark,
  you get two logos. Check before shipping.

### Icon Library

`slides/icons/` holds 16 reusable outline icons, all `96x96` viewBox, `#6983FF` stroke at width 4,
no fills: `community`, `cost-efficiency`, `data-hash`, `data-list`, `durability`, `general-scaling`,
`hardening`, `horizontal-scaling`, `memory-efficiency`, `migration`, `observability`, `performance`,
`reliability`, `replication`, `security`, `vertical-scaling`.

Keep a validator for these and run it on every addition: single root `<svg>` at `96x96`, every element
`fill="none"` with a `#6983FF` or `#232F3E` stroke, 3-8 elements, geometry inside the viewBox, and none
of `<style>`, `class=`, `style=`, gradients, filters, masks, `clipPath`, `<text>` or `<use>`. **Exclude
the derived `valkey-mark-*.svg`** — those come from the brand pack and legitimately carry a gradient
and a `<style>` block.

A grid of icons should adapt rather than assume: `rows = ceil(n / cols)`, and shrink the icon to
`row_pitch - label_gap - label_h - padding` so an extra row never overlaps the one above.

Sizing that works: **1097280 EMU (1.2in)** for an icon above an 18pt card heading, **914400 EMU
(1.0in)** in a 4-column gallery grid. 0.7in looks like a bullet point next to real text; do not go
smaller than 1in on a card.

To add more, keep the constraints: single root `<svg>` with `width="96" height="96"
viewBox="0 0 96 96"`, presentation attributes only, and no `<style>`, CSS classes, `style=`,
gradients, filters, masks, `clipPath`, `<text>`, or `<use>`. PowerPoint is strict, and those
features render inconsistently or not at all.

### Embedding SVGs

Use `slides/pptx_svg.py`:

```python
from pptx_svg import add_svg
add_svg(slide, "slides/icons/security.svg", "slides/icons/png/security.png",
        left, top, 1097280, 1097280, name="security")
```

python-pptx cannot insert SVGs (PIL cannot decode them), so that helper builds the package parts by
hand. PowerPoint renders the `asvg:svgBlip` extension and falls back to the embedded PNG in older
viewers. Three traps, all of which produce silently wrong output:

1. **`Part`'s signature is `Part(partname, content_type, package, blob)`.** Wrong argument order
   gives a 0-byte part and an "image cannot be displayed" box.
2. **A PNG fallback is required, and it must be correct.** Rasterize with headless Chrome, but
   **always go through an HTML wrapper with an explicitly sized `<img>`.** Pointing Chrome straight
   at a `.svg` fails in two different ways:

   - Illustrator/Inkscape exports (all the brand logos) render as a **fully transparent image**.
     Not a timing issue; `--virtual-time-budget` does not help.
   - An SVG that declares `width`/`height` renders at that natural size inside the window rather
     than scaling to fit, so an oversized window silently corner-anchors the artwork with
     transparent padding around it.

   The wrapper sidesteps both, because CSS sizing always wins:

   ```html
   <html><body style="margin:0">
     <img src="file:///abs/path/icon.svg" style="width:384px;height:384px;display:block">
   </body></html>
   ```

   ```bash
   chrome --headless --disable-gpu --default-background-color=00000000 \
     --window-size=384,384 --screenshot=out.png file:///tmp/wrapper.html
   ```

   Then assert it: the rendered content's bounding box should fill most of the canvas. Under ~60%
   means something went wrong. `rasterize()` in `build_placeholder_deck.py` does all of this.
3. **`shape.image` returns the PNG fallback, not the SVG.** It resolves the `a:blip` `r:embed`. To
   read the actual SVG, follow the `svgBlip` `r:embed` instead.

## Running Header

A running header belongs on **every** slide, including the ones that already have a title. Two
mistakes to avoid, both of which look reasonable in code:

- **Too small.** A 10pt line reads as a footnote. 16pt is visible without competing with a title.
- **Only on untitled slides.** Skipping titled slides means the header lands *only* on the title and
  closing slides, where it duplicates the big title. That is the opposite of a running header.

What works:

- Right-aligned in a band **above** the title: `(609600, 130000)`, size `11062800 x 260000`,
  vertically centred, 16pt regular.
- Right-aligned and above, so it coexists with any slide's own title however long that title runs.
  Do not right-align it *within* the title band: a 44-character title at 32pt is already 11.4in
  wide, and the header needs 1.9in of the 13.3in slide.
- `DARK_2` normally, `LIGHT_1` where the top of the slide is a dark fill.
- Its own text box per slide, so it can be deleted freely.

Add it in a pass over the finished deck, not inside each generator, so a newly added slide cannot
forget it, and run it last so it draws over any background fill. Detect the dark background rather
than hardcoding slide numbers: walk the autoshapes for a solid dark-theme fill covering the header
band.

## The Glyph, and Where to Use It

**The brand pack ships only lockups** (glyph plus the "Valkey" wordmark), and at corner-mark size the
wordmark is illegible. Derive a glyph-only asset instead of shipping an illegible lockup: parse
`valkey-stacked-<colour>.svg`, remove the `<g id="text...">` wordmark group, and tighten the viewBox
to the glyph. The glyph's box is `248.0 0 282.6 326.0`, aspect **0.867** — measured by rasterising
the wordmark-free SVG and taking its bounding box, which is the reliable way to find it.

Where the glyph belongs: the **upper-right corner of boxes that are a real instance or shard**,
`220000` EMU tall, inset `110000` — exactly where those boxes used to carry a blue top bar. It fills
the space the stripe vacated, stops a grid of grey rectangles looking inert, and unlike the stripe it
carries information: a box with the glyph is a server, a box without one is a client, a proxy, or a
content card.

Where it does not belong: **loose in the corner of the slide.** The master already puts a full lockup
bottom-right, so a second mark floating at the top is just another piece of furniture. Brand marks
should be attached to something they identify.

**Cache the image parts.** The same logo placed on 21 boxes would otherwise add 42 media parts to
the package. Key a cache on `(id(package), svg_path)` and reuse the `Part` objects;
`slide_part.relate_to()` returns the existing rId when a slide already references one. In the
placeholder deck this keeps 62 pictures down to 35 media parts.

## Speaker Notes

```python
slide.notes_slide.notes_text_frame.text = "..."
```

Notes are a common source of embarrassment: duplicating a slide duplicates its notes, so notes
routinely end up describing a completely different slide. When you copy a slide, either rewrite
the notes or clear them. Dump notes before presenting.

## Structural Edits

Order of operations, always: **delete (highest index first) → add → reorder once at the end.**

```python
# Add, then move into position (0-indexed).
slide = prs.slides.add_slide(layout)          # appends
lst = prs.slides._sldIdLst
lst.insert(target_index, lst[-1])             # insert moves the element
```

Deleting a slide is safe **as long as you drop the relationship and remove the `sldIdLst` entry
together.** python-pptx only serializes parts reachable through the relationship graph, so the
orphaned slide part is dropped on save. Verified: stripping all 31 slides from a deck left zero
`ppt/slides/*` parts and preserved all 17 layouts, the master, and the theme.

```python
rId = prs.slides._sldIdLst[index].rId
prs.part.drop_rel(rId)          # both lines, or you leave a dangling r:id
del prs.slides._sldIdLst[index]
```

Removing the `sldIdLst` entry *without* `drop_rel` is what produces the "needs repair" prompt.
Always reopen the saved file afterwards to confirm.

### Starting a New Deck From the Template

There is no blank Valkey template file, so clone an existing deck and empty it. Use the deck with
the richest layout set as the donor:

```python
import shutil
shutil.copyfile("slides/The Vertical Scaling wall.pptx", "slides/New deck.pptx")
prs = Presentation("slides/New deck.pptx")
for i in range(len(prs.slides) - 1, -1, -1):     # highest index first
    prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
    del prs.slides._sldIdLst[i]
```

You keep the theme, master, all layouts, and the logo. A handful of media parts referenced by the
master and layouts stay behind; that is correct, not leakage. See
`slides/build_placeholder_deck.py` for the full pattern.

If a deck lacks a layout you need, add it with `slides/pptx_layout.py` rather than hand-drawing the
design onto every slide. See Adding a Layout to the Master.

## Verifying Your Work

There is no LibreOffice or PowerPoint CLI here, but you can still **see your slides**: read the real
geometry, theme fills, text runs and embedded images out of the file, lay them out as HTML, and
screenshot that with headless Chrome. An approximate renderer built that way reliably catches the
things that actually go wrong: text overflowing a card, shapes colliding, icons at the wrong scale,
dead space, bad alignment.

The rest of this section is what such a renderer can and cannot tell you, learned by building one.
It is worth the hour, and the gaps below are worth knowing before you trust any output.

**Draw master and layout content beneath each slide's own shapes** (the corner Valkey logo, the dark
full-bleed backgrounds on the title layouts), or what you see omits inherited artwork and
misrepresents the slide badly.

It is an approximation, not a renderer. Expect these gaps: Play and Century Gothic fall back to a generic
sans so line breaks differ slightly; autofit, **bullet glyphs**, **paragraph `spcBef` spacing**,
rotation, gradients, and shadows are ignored; charts draw as labelled grey boxes; placeholder text
inherited from a layout does not appear. **Never claim a slide is visually correct on the strength
of an approximate render alone.** It proves layout, not typography. Say which you verified, and check the
things it cannot show (carets, spacing) in the XML instead.

**Text inheritance from a layout is the hardest gap to close,** and it limits what any quick renderer
tells you about a slide built from one. Position and size of the box come through, because those are on the slide's own
placeholder, but nothing from the layout's `lstStyle` does: a 130pt big number and a 13pt card body
render at the same default, and `algn="ctr"` renders left. Renderer gaps rather than deck bugs, and they
cannot be settled that way. To check type scale, add a slide from the layout, fill the placeholders,
and **open it in PowerPoint once.**

**Do resolve a placeholder's *fill* from its layout placeholder, and do draw linear `a:gradFill`.**
Skipping either hid a real bug once: a row placeholder whose card background lives on the layout
rendered as bare text, so a broken fill shipped unseen. If you use a fill mechanism your renderer
does not know about, teach it rather than working around it.

What you *can* establish without opening it: that the mechanism is the one the template already
relies on everywhere. Both size and alignment reach slides through layout `lstStyle` in the stock
layouts — that is where the whole type scale comes from, and `Video` sets `algn="ctr"` there. If
inheritance were broken, every stock layout would be broken too.

A caution learned the hard way, four times: a renderer's own bugs make *correct* decks look broken.
An alignment map keyed on XML values instead of python-pptx enum names; the same for theme colours;
PowerPoint's internal text margins applied as CSS padding without `box-sizing:border-box`, which drew
every filled shape up to 0.2in wider than it really was; and a hardcoded colour scheme that rendered
the patched `accent6` as the stock peach. Read the theme from `theme1.xml`, and draw unmapped theme
colours in magenta with a warning rather than quietly substituting black.

**When the render looks wrong, confirm against the XML before "fixing" the deck.** Measuring the
rendered PNG (bounding boxes of a known colour) against the EMU geometry in the file is the fastest
way to tell a renderer bug from a real one.

Alongside it, always:

- Re-dump every slide you touched and read it back.
- Check bounds arithmetic: `left + width <= 12192000`, `top + height <= 6858000`.
- Count runs before and after. A paragraph that had two runs and now has one lost its formatting.
- Reopen the saved file with python-pptx. If it raises, PowerPoint will complain too.

For a pixel-accurate check, LibreOffice (`brew install --cask libreoffice`) would allow
`soffice --headless --convert-to png`. Worth suggesting when typography matters.

## Pitfalls

1. **Formatting is per-run, not per-paragraph.** Never assign a whole line to `runs[0].text`. A
   bullet with a highlighted number is two runs; collapsing them kills the color and bold. Dump
   the run structure first, then edit runs in place or rebuild all of them.
2. **New placeholders have zero runs.** Use `para.add_run()`; do not assume `para.runs[0]` exists.
3. **Layout indices are per-deck.** Resolve by name, every time.
4. **`replace_data` drops series you omit.** Re-supply the full set.
5. **Delete from the highest index down,** and always pair `drop_rel` with the `sldIdLst` deletion.
6. **Group shapes hide their children.** `shape.shape_type == 6` means recurse into
   `shape.shapes`; top-level iteration will not find the text inside.
7. **Theme colors, not RGB.** `MSO_THEME_COLOR.ACCENT_1`, not `RGBColor(0x69, 0x83, 0xFF)`.
8. **Duplicated slides carry stale notes and duplicate shape names.** Both bite later.
9. **Check for `~$` lock files before saving.**
10. **Off-template layouts (`DEFAULT`, `Blank`) lose the footer and logo.** If a slide looks
    naked compared to its neighbors, check what layout it is on.
11. **Never derive a shape id from anything but the shape tree.** `add_svg` used to set
    `id = media_number + 1000`, which collides with the ids python-pptx hands out as soon as
    pictures and shapes interleave: it left duplicate ids on 12 of the 31 slides in the placeholder
    deck. PowerPoint tolerated it silently, which is why it survived. Allocate from
    `max(existing) + 1` over the tree you are appending to, and assert uniqueness after building.
12. **Gradient stop positions must be strictly increasing.** `a:gsLst` positions are required to be
    distinct; duplicates are undefined and PowerPoint resolves them by dropping a stop, turning a
    hard edge into a full-width wash. Nothing in this repo generates a gradient any more, for the
    reason above, but `pptx_layout.audit()` still checks it: a deck edited in PowerPoint can bring
    one in.
13. **The mock renderer is not a source of truth about the theme.** It used to hardcode the colour
    scheme, so a deck whose `accent6` had been patched to the brand rose rendered in the old stock
    peach. It reads `theme1.xml` now. When a colour looks wrong in the mock, check the theme blob
    before changing the deck.

## Deck-Specific Notes

Per-deck structure, update cadence, and content sources live alongside each deck:

- `slides/valkey-placeholder-slides.md` — the placeholder/starter slide library

When you learn something new about a deck, write it there. When you learn something about the
template or the tooling, write it here.
