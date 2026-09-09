# Valkey community deck

File: `slides/Valkey community deck.pptx`
43 slides. General-purpose reusable deck for community talks. Read `../slides-skill.md` first for
template mechanics.

Verified against the file on 2026-08-26. The previous version of this table (in `slides-skill.md`)
claimed 42 slides, named slide 15 "Valkey Admin", and listed layout indices that did not match the
file. Re-dump before trusting any of it: `python3 dump_deck.py "Valkey community deck.pptx"`.

## Structure

### Section 1: General (1-18)

| Slide | Title | Layout | Update cadence |
|-------|-------|--------|----------------|
| 1 | (Title slide) | Holding | Stable |
| 2 | Background | Divider | Stable |
| 3 | Introducing the Valkey project | Title + Cont | Stable |
| 4 | Backed by industry leaders | Title + Cont | Leave as-is |
| 5 | The Valkey project is thriving | Title + Cont | **Every release** (stats) |
| 6 | Latest developments in Valkey OSS | Title + Cont | **Every major release** (timeline) |
| 7 | Valkey 7.2 (April 2024) | Title + Cont | Stable |
| 8 | Valkey 8.0 (September 2024) | Title + Cont | Stable |
| 9 | Valkey 8.1 (April 2025) | Title + Cont | Stable |
| 10 | Valkey 9.0 (October 2025) | Title + Cont | Stable |
| 11 | Valkey 9.1 (May 2026) | Title + Cont | Stable |
| 12 | Get involved in the project | Title + Cont | Stable |
| 13 | Integrations | Divider | Stable |
| 14 | Valkey module distribution | Title + 3x Cont with Headings | When modules change |
| 15 | Valkey GLIDE | Title + Cont + Right Img Small | When GLIDE updates |
| 16 | Migrations | Divider | Stable |
| 17 | Migrating from Redis OSS to Valkey | Title + Cont | Stable |
| 18 | Strategies for migration | Title + Cont | Stable |

### Section 2: Features (19-33)

| Slide | Title | Layout |
|-------|-------|--------|
| 19 | Memory Efficiency | Divider |
| 20-26 | A new hash table (7-slide build) | Title + Cont |
| 27 | Performance | Divider |
| 28-32 | Thread scalable command processing (5-slide build) | Title + Cont |
| 33 | Multithreaded performance | Title + Cont |

### Slides 34-43

Template/reference slides kept for layout examples, mostly empty: Thank You 2x (34),
Title + 2x Speaker (35), Divider (36), Title + Sub + Cont (37), Title Only (38),
Title + Subtitle (39), Title + Cont + Right Img Small (40), plus three built-out
Title Only slides: "Three queries, one answer" (41), "Standalone vs. Cluster mode" (42),
"Standalone" (43).

## What to update

### Slide 5: project stats

Update the bullet numbers. Each bullet is two runs at 28pt: the number (`ACCENT_1`, bold) and the
description (`DARK_1`, regular). Preserve that structure. See the number-highlighting section of
`../slides-skill.md`.

Sources:
- Contributors: https://github.com/valkey-io/valkey
- Container pulls: https://hub.docker.com/r/valkey/valkey
- PRs: GitHub PR count across the valkey-io org
- AWS database blog often has consolidated stats (search "Valkey turns" / "Valkey anniversary")

Speaker notes hold the managed provider list. **Leave it as-is unless explicitly asked to change
it.**

### Slide 6: release timeline

The timeline is a group shape. Date labels are children of the group (`Shape 1`); the release
labels ("Valkey 8.0 Release") are separate top-level text boxes. Recurse into the group:

```python
for shape in slide.shapes:
    if shape.shape_type == 6:  # GROUP
        for child in shape.shapes:
            if child.has_text_frame:
                print(child.text_frame.text)
```

### Adding a release overview slide

When a new major/minor ships, add a summary slide after the existing release slides (currently
after 11) using the `Title + Cont` layout. Follow the release-slide house style: 24pt bullets,
numbers highlighted in `ACCENT_1` bold, feature descriptions in plain `DARK_1`, and a 10pt italic
source citation pinned to the logo baseline.

## Content sources

1. Releases: https://valkey.io/download/releases/
2. Blog: https://valkey.io/blog/
3. GitHub releases: https://github.com/valkey-io/valkey/releases
4. Year-end reviews: https://valkey.io/blog/2025-year-end/ (pattern: `<year>-year-end`)
5. Stats: AWS database blog, LFX analytics, Docker Hub
6. Release cadence: https://valkey.io/topics/releases/ (major yearly, minor as needed)
