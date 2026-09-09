# Valkey placeholder slides

File: `slides/Valkey placeholder slides.pptx`
21 slides, 11 layouts, and **no stock layouts at all**. The reusable frames are in the slide master;
the slides are the things a layout cannot carry, bookended by a placeholder intro and thank you.

**Generated, not hand-edited.** Rebuild with:

```bash
python3 slides/build_placeholder_deck.py
```

That script strips the template, adds the eleven layouts, rebuilds all 21 slides, deletes the fifteen
stock layouts, re-rasterizes the PNG fallbacks, and re-embeds the SVGs and brand logos. It builds to
a temp file and only moves it into place on success, so a failed run cannot destroy a good deck. Edit
the script, not the pptx, or your changes vanish on the next rebuild. Once you copy a slide into a
real deck, it is yours and the script no longer touches it.

## No stock layouts

The master is the eleven Valkey layouts and nothing else. `Title + Cont`, `Divider`, `Quote`, `Code`,
`Full Image`, `Video`, `Holding` and the rest are gone, along with the full-bleed background picture
each of them carried: 894 KB down to 528 KB, and the four canvas-overhanging shapes in the package
were all in those layouts.

`strip_stock_layouts()` runs **last**, after every slide exists. `SlideLayouts.remove()` refuses a
layout a slide still depends on, which is why `Valkey Title Only` had to be added first: all 19
diagram and card slides sat on the stock `Title Only`. Mechanics and the pasting caveat are in
`../slides-skill.md`.

The trade this makes: no bulleted-content layout, no blank, no `Divider`, no `Quote`. If you want any
of those here, add them as Valkey layouts rather than reinstating the stock ones. The other three
decks in `slides/` still have the full stock set.

## Layouts, not demo slides

Ten designs used to be hand-drawn onto a slide so you could copy the slide. They are slide layouts
now, plus `Valkey Title Only` as the foundation, so PowerPoint's New Slide gallery offers them by
name and the text is placeholders. The demo slides that existed only to be copied are gone:
`build_layouts()` in `build_placeholder_deck.py` is where they live.

| Layout | Placeholders | Was |
|---|---|---|
| `Valkey Title Only` | title | the stock `Title Only`, which every custom slide needed |
| `Valkey Title Band` | title, subtitle, speaker, role | title option A |
| `Valkey Title Top Banner` | title, subtitle, speaker, role | title option B |
| `Valkey Title Bottom Banner` | title, subtitle, speaker, role | title option C |
| `Valkey Insight` | statement, caption | insight, quiet |
| `Valkey Insight Bold` | statement, caption | insight, loud |
| `Valkey Big Number` | number, label, source | insight, one number |
| `Valkey List Rows` | title, 5 rows | list placeholder |
| `Valkey Two Box` | title, 2x (heading, body) | 3 two-box placeholder slides |
| `Valkey Three Box` | title, 3x (heading, body) | three-box placeholder |
| `Valkey Thank You` | closing, speaker, 2 contacts | thank you |

All three title options survive as three layouts, because they are three answers to the same
question and picking one is the presenter's call. The two-box **fill** variants did not: purple,
rose and neutral were one design in three colours, and recolouring a card on a slide is two clicks.
The rose pairing (`ACCENT_5` + the patched `ACCENT_6`) is the one in the master. The deep purple is
deliberately absent: it needs `LIGHT_1` text and white icon variants that do not exist, so it is not
a fill swap.

Mechanics of adding a layout, and what a layout can and cannot express, are in
`../slides-skill.md`. Two decisions specific to these:

- **Card icons are static, not picture placeholders.** A picture placeholder would be swappable but
  a new slide would come up with an empty grey box. Baking the icon in means the slide looks right
  immediately; swapping it means replacing the picture on the slide.
- **A list row is two flat shapes, exactly like `slide_bullets_cards()` builds slide 14.** A static
  `ACCENT_1` bar at `CONTENT_LEFT` width `BAR_H`, then a `LIGHT_2` placeholder starting at
  `CONTENT_LEFT + BAR_H` holding the text. Five rows, which is inside the four-to-six ceiling.

  It was one placeholder with a two-stop `a:gradFill` faking the accent edge, and that was wrong on
  both counts: it shipped malformed once (duplicate stop positions, so every row rendered as a
  blue-to-grey wash), and even correct it made PowerPoint show a flat card as gradient-filled. The
  placeholder has to start *after* the bar rather than spanning it, because layout shapes draw
  beneath slide shapes and a full-width row would cover the bar. Side by side also keeps the card
  travelling with the text, so deleting a row leaves only the hairline bar.

## Slides

| # | Title | Shape |
|---|-------|-------|
| 1 | Presentation title goes here | **Placeholder intro**, on `Valkey Title Bottom Banner`. Replace four lines and go. |
| 2 | Metric comparison | Same instant, two readings: the 200%-of-400% aggregate looks half idle, the main thread is at 99% of the one core it can use. |
| 3 | `CLUSTER SLOT-STATS` | Which slot to move: `key-count` flat vs `cpu-usec` skewed, same five slots. |
| 4 | Standalone | One node, clients arrow in. |
| 5 | Replication | Primary plus replica, replication arrow, writes/reads labelled. |
| 6 | Clustering | Three shards with correct slot ranges, primary over replica, dashed cluster-bus hops. |
| 7 | `CLUSTER SETSLOT` | **Legacy** key-by-key migration. Source `MIGRATING`, target `IMPORTING`, clients following `-ASK`. |
| 8 | `CLUSTER SETSLOT STABLE` | **Legacy** abort. Keys already moved are stranded and must be hand-migrated back. |
| 9 | `CLUSTER MIGRATESLOTS` | Atomic: Initializing. Link established, staging area created, no data moving yet. |
| 10 | `CLUSTER SYNCSLOTS` | Atomic: Snapshotting and Streaming changes into the staging area. |
| 11 | `CLUSTER SYNCSLOTS` | Atomic: Finalized. Source pauses and drops its copy, target owns the slots. |
| 12 | `CLUSTER CANCELSLOTMIGRATIONS` | Atomic: Failed. Staging area discarded; the source never stopped owning the slots. |
| 13 | `CLUSTER CANCELSLOTMIGRATIONS` | Same, in the darker brick red. Delete whichever loses. |
| 14 | Valkey 8.0 — row cards | **Default** list treatment, with real content. Same shape as `Valkey List Rows`. |
| 15 | Valkey 8.0 — soft list | **Alt B.** Small square bullet, 20pt `DARK_2`, narrower measure, more air. |
| 16 | Valkey 8.0 — icon rows | **Alt C.** An icon per row with hairline dividers. Most designed. |
| 17 | Two ways to grow (calmer palette) | Worked example: purple against neutral grey. |
| 18 | Two ways to grow | Worked example: purple against the rose. |
| 19 | What we optimize for | Worked example: performance / hardening / cost efficiency. |
| 20 | Icon library | All 16 icons in a 4x4 grid with labels. Reference sheet, not a talk slide. |
| 21 | Thank you | **Placeholder closing**, on `Valkey Thank You`. Bookends slide 1. |

Slides 2-20 are content, not frames. Slides 15-16 and 17-18 are still alternatives to pick between,
and slide 13 is still the brick-red comparison.

Slides 1 and 21 are the exception: they are the only slides here built by
`slide_from_layout()`, which adds a slide from a layout and fills the placeholders. Their runs
deliberately name no font and no size, so everything comes from the layout's `lstStyle`. That is the
one place in the builder where not pinning Arial is correct, and it is also what makes them the live
demonstration that the inheritance chain works. Slide 1 uses the bottom-banner title option because
`Valkey Thank You` was designed to bookend that one; swap `slide_intro()` to `LAYOUT_TITLE_BAND` or
`LAYOUT_TITLE_TOP` if you prefer another.

## List slides

`slide_bullets_cards()` is the default and slide 14 uses it: one `LIGHT_2` row per item with a blue
left edge, text at 18pt, numbers highlighted in `ACCENT_1` bold. `Valkey List Rows` is the same
shape as a layout.

**The caret-bullet style is deliberately not in this deck.** Community deck slides 7-8 use ► carets
in Valkey blue on raw white, and that is genuinely the community deck's house style, but in a
card-led deck it reads harsh: it is the only slide type with no fill, no card, and no structure, so
it looks like it came from a different deck. The caret XML recipe is still documented in
`../slides-skill.md` for when you are editing the community deck itself.

Two softer alternatives remain for picking: `slide_bullets_soft()` (small square bullet, 20pt
`DARK_2`, narrower measure) and `slide_bullets_icons()` (an icon per row with hairline dividers).
Slides 14-16 are the same list in all three, so you can compare and delete the two you do not want.

All three share `V8_ITEMS`, `V8_ICONS`, `V8_TITLE` and `V8_SOURCE`, so changing the content changes
all of them. `item_runs()` turns a `(highlight, rest)` tuple into two runs and a plain string into
one, which is what lets the generators share content. Highlight measurable numbers only, never
ordinals or version numbers. **This is why the highlight cannot move into a layout:** it is two runs
with different colours, and a layout's `lstStyle` styles a whole level.

## Which metric, for which question

Two different scaling questions, two different metrics, and they are not interchangeable:

| Slides | Metric | Question it answers |
|---|---|---|
| 2 | `used_active_main_thread` | Are we at the vertical ceiling? One thread, one core, no more room. |
| 3 | `CLUSTER SLOT-STATS` `cpu-usec` | If we shard, which slot moves? |

Slide 3 lands on the trap that makes this a real argument: **`key-count` is the only metric reported
by default.** `cpu-usec`, `network-bytes-in` and `network-bytes-out` need
`cluster-slot-stats-enabled yes`, so the metric most people rebalance on is the one that says every
slot is identical. Verified against https://valkey.io/commands/cluster-slot-stats/
(`CLUSTER SLOT-STATS`, Valkey 8.0).

## Feature emphasis and topology

`slide_metric_compare()` is the feature-emphasis shape. It leans on
`bar_rows(slide, x, y, w, [(label, value), ...], highlight=True)`, which draws a light track per row
and fills it proportionally, colouring the peak `ACCENT_1` and the rest `ACCENT_5`. Pass
`highlight=False` when the point is that the distribution is *flat*.

The metric it features is **`used_active_main_thread`** (the `METRIC` constant). Valkey serialises
command execution on one thread, so that thread is the ceiling and host CPU averages hide it. This
is the vertical-scaling-wall argument, not a rebalancing argument: an earlier draft used per-slot
`cpu-usec` from `CLUSTER SLOT-STATS`, which answers "which slot should I move", a different question.
Swap `METRIC`, `HOST_CPU`, and `THREAD_CPU` to retarget it.

The template has no monospace font and the two-font rule says do not add one, so command names go in
a dark `chip()` set in the brand font rather than a terminal block. If you want true monospace,
`Courier New` is the only safe cross-platform choice, and it will look dated.

`slide_standalone()`, `slide_replication()`, and `slide_clustering()` modernise The Vertical Scaling
wall slides 11-13: square corners instead of rounded, grey `node()` cards, and `arrow()` connectors
with real arrowheads instead of block arrows. Slot ranges are the correct ones
(0–5460 / 5461–10922 / 10923–16383); the originals had 8196/8197 and 16343, which are wrong.

`arrow()` sets `a:tailEnd` (and `a:prstDash`) directly, because python-pptx has no arrowhead API.
Child order inside `a:ln` matters: `prstDash` before `tailEnd`. An approximate renderer will not draw
arrowheads or dashes, so verify those in the XML, not the render.

## Slot migration

Slides 7-13 are one run: two slides on the **legacy** mechanism, then four on **atomic slot
migration**, the Valkey 9.0 design, plus the brick-red variant. Source:
https://valkey.io/blog/atomic-slot-migration/

The legacy slides exist **only as the "before" half of the comparison** and are labelled with their
own commands. Never present that protocol as the current design: it produced `-ASK` redirects, broke
multi-key commands mid-migration, and had no real rollback.

All six share one geometry so old and new read against each other:

- `mig_pair()` draws source and target shards at `SRC_X` / `TGT_X`, width `MIG_W`, with a state badge
  under each. The badges and their colours are the only per-slide difference.
- `mig_arrow()` is the single labelled arrow. `reverse=True` points it back, `dashed=True` means a
  link that exists but is not carrying data yet.
- `step_strip()` fills the bottom bar. It takes an int or a **set**, because snapshotting and
  streaming overlap.

Three strips, all four cells wide so the bar never changes size between slides:

| Constant | Cells |
|---|---|
| `LEGACY_STEPS` | `SETSLOT IMPORTING` / `SETSLOT MIGRATING` / `GETKEYSINSLOT + MIGRATE` / `SETSLOT NODE` |
| `STATES` | `Initializing` / `Snapshotting` / `Streaming changes` / `Finalized` |
| `FAILED_STATES` | same, with `Failed` in the failure red |

The legacy bar lists commands because the old mechanism had no job states to report. That is itself
part of the argument: `CLUSTER GETSLOTMIGRATIONS` has nothing to be equivalent to.

The point slide 12 exists to make: incoming data lands in a **staging area** on the target, separate
from the live dataset, and ownership is never broadcast until finalization. Cancelling is just
discarding it, and out-of-memory, failover, and network partitions are handled the same way.

## Fonts

Every run this builder creates names **Arial** (`FONT` / `TITLE_FONT`), titles included, and so does
every added layout's `lstStyle`. Checked by collecting `run.font.name` across the deck and every
layout `a:latin` typeface: one value, `Arial`. The theme fonts are still Play and Century Gothic,
which is what inherited placeholder styling uses; the constants exist so headings can go back to
Play in one line.

There is no running header. `add_headers()` existed to put a deck title on the seven custom slides
that had no real title of their own, and every one of those is a layout now.

## Review fixes and colour alternatives

Four points from a review pass, and what happened to each:

- **Accent stripes on top of boxes deleted.** Every card and node card had a thin `ACCENT_1` bar
  across its top edge; those are gone, on the two/three-box cards, the topology and migration node
  cards, and the comparison cards. The fill defines the card on its own.

  Deliberately **kept**, because they are not lines on top of a box: the blue left edge on list rows,
  the full-width rule under the navy banner on the top-banner title layout, the centred rule on the
  bottom-banner one, the rule above the insight statement, the rule on the closing layout, and the
  hairline row dividers on the icon-row list. To re-audit: for each thin shape, check whether a
  larger shape shares its exact left, top and width. There is no build-time check for this.
- **Failure red.** `ACCENT_3` (`#632637`) reads brown. Failure states use an explicit
  `FAIL_RED = #D32F2F`; `FAIL_RED_BRICK = #C0392B` is the darker alternative, shown on the extra
  `CLUSTER CANCELSLOTMIGRATIONS` slide so the two can be compared side by side. Delete whichever
  loses. See `../slides-skill.md` for why RGB is allowed here.
- **Title fonts.** Consistent: every title is Arial. See Fonts.
- **Peach replaced in the theme.** `patch_theme()` rewrites the template's `accent6` from the stock
  peach `#FBD8BF` to the brand rose `#E0A2AF` (Pantone 494 C), so the rose is a real theme colour and
  the cards reference `ACCENT_6` rather than a hardcoded RGB. Card text is `ACCENT_2` (`#1A2026`) and
  the deep purple is `ACCENT_4` (`#30176E`, Pantone 2685 C) — both already correct in the theme.
  **Caveat:** the concept icons only exist in `#6983FF`, so they lose contrast on the deep purple;
  picking it means generating white icon variants.

  Read the scheme from `theme1.xml` when checking this. A renderer with a hardcoded palette shows the
  old peach and makes a correct deck look wrong.

## Instance marks

`node()` puts the Valkey **glyph** (not the full lockup — the wordmark is illegible this small) in the
**upper-right** corner of every box, `220000` EMU tall, inset `110000`: exactly where these boxes used
to carry a blue top bar. `make_marks()` derives `valkey-mark-color.svg` and `valkey-mark-white.svg`
from the stacked lockups at build time by dropping the wordmark group and tightening the viewBox to
`248.0 0 282.6 326.0`. It is only on `node()`, which is exactly the boxes that are a real Valkey
instance or shard — so a box with the mark is a server and a box without one is a client, a proxy, or
a content card. Sub-labels are `230000` tall so they stop where the mark begins.

`pptx_svg.py` caches image parts per package, because the same logo lands on many boxes and would
otherwise add two media parts per placement. 62 pictures currently need 41 media parts.

Shape ids come from `max(existing) + 1` over the target shape tree. They used to be derived from the
media number, which collided with python-pptx's own ids and left duplicates on 12 of the 31 slides
in the previous build. See `../slides-skill.md` pitfall 11.

## Card anatomy

Every card is four shapes, built in this order (z-order matters, the icon must land last):

1. Background rectangle, solid theme fill, no border (or no fill + hairline border, uncoloured)
2. Icon, `1097280` EMU (1.2in) square, inset by `PAD`
3. Heading text box, 18pt bold
4. Body text box, 13pt, line spacing 1.2, filling the remaining height

`card_geometry()` returns those three inner rectangles and is shared by `card()` and the Two Box /
Three Box layouts, so a slide this script builds and a slide built from the layout in PowerPoint
cannot drift apart. `box_columns(n)` does the same for the column edges.

Geometry: `GAP=200000`, `PAD=182880`, `CARD_H=3900000`. Two-up columns are `5431400` wide,
three-up are `3554266`. Cards start at `CONTENT_TOP=1445931`, so they bottom out at `5345931`,
clear of the logo at `5943597`.

`card()` takes `fill`, `border`, and `text`. `fill=None, border=DARK_2` gives an outlined white card.
Neither is used by the current slides, which are all filled cards.

The icon-library grid is centred on the **page**, both axes, not in the content box. Two separate
traps, both invisible in prose and obvious on a slide that is nothing but a grid:

- Horizontally, the template's content box is asymmetric (left inset `609600`, width `11062800`, so
  it overhangs the right margin by `90000` EMU). Use `SLIDE_W - 2 * CONTENT_LEFT` instead.
- Vertically, anchoring the first row to `CONTENT_TOP` left the block sitting `299485` EMU (0.33in)
  below the page's vertical centre. Measure the real block height and centre it:
  `grid_top = (SLIDE_H - block_h) // 2`.

The donor template is the **community deck**, not the talk deck. It is the only git-tracked pptx
here, so it is recoverable. It is read-only in the build, but it does get opened in PowerPoint, so
the donor is whatever was last saved to disk.

## Icons

**16** outline icons in `slides/icons/`, with 4x PNG fallbacks in `slides/icons/png/`:
community, cost-efficiency, data-hash, data-list, durability, general-scaling, hardening,
horizontal-scaling, memory-efficiency, migration, observability, performance, reliability,
replication, security, vertical-scaling. Names and sizing guidance are in `../slides-skill.md`. The
PNGs are generated; do not hand-edit them.

The gallery adapts: `rows = ceil(len(names) / 4)` and the icon shrinks to fit the row pitch, so
adding a row does not make them overlap. At 16 icons that is a 4x4 grid with a `629837` EMU icon,
still centred on the page to the EMU.

`valkey-mark-*.svg` are derived brand assets, not concept icons — exclude them from the concept-icon
validator, which forbids the gradients and `<style>` blocks the brand files legitimately use.

## Verification status

Structurally verified: 21 slides, 11 layouts, zip integrity, no zero-byte parts, 62 pictures each
carrying a resolving `svgBlip` plus a non-empty PNG fallback, no duplicate shape ids on any slide or
layout, **all geometry inside the canvas with nothing left over** (the four overhanging shapes went
with the stock layouts), brand logo aspect ratios undistorted, reopens cleanly. Every layout name
starts with `Valkey `, every slide sits on one, and the count of `slideLayoutN.xml` parts in the zip
matches `len(prs.slide_layouts)`, so no orphaned layout part survived the removal. Every layout
yields the expected placeholder count when a slide is added from it. `theme1.xml` confirmed carrying
`accent6 = E0A2AF`. Of 271 slide runs, all pinned fonts are Arial and exactly 8 inherit from a
layout: the four lines on slide 1 and the four on slide 21.

Layout verified with an approximate HTML render, which substitutes a generic sans for Play and
Century Gothic. **Not verified in PowerPoint**, and two things specifically cannot be:

- **Type scale and alignment on slides built from a layout.** The mock resolves no *text*
  inheritance, so the `lstStyle` sizes (130pt big number, 54pt closing, 40pt insight, 18/13pt cards)
  render at its default and the centred title on slide 1 renders left. The mechanism is the one the
  stock layouts used for the entire template type scale, and one of them (`Video`) set `algn="ctr"`
  the same way, which is the evidence it is right. It has not been seen.

The list rows render identically to generator-built slide 14, checked by resolving the placeholder
fill from the layout rather than reading the slide's own empty `spPr`. A renderer that skips that
shows the rows as bare text, which is how the gradient bug above stayed hidden.

The check that matters: copy the deck, strip its slides, add one slide from every layout, and fill
every placeholder, asserting the placeholder count is what you expected. That is what proved all
eleven still work after the stock layouts were removed. Open the result once and confirm the type
scale, which no headless render will show you.

```python
prs = Presentation(copy_of_deck)
for i in range(len(prs.slides) - 1, -1, -1):
    prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
    del prs.slides._sldIdLst[i]
for layout in prs.slide_layouts:
    slide = prs.slides.add_slide(layout)
    phs = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)
    assert len(phs) == len(sample_text[layout.name]), layout.name
```
