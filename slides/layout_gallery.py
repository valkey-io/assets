#!/usr/bin/env python3
"""Generate one filled placeholder slide per layout, to check the layouts work.

    python3 slides/layout_gallery.py [deck.pptx] [out.pptx]

Defaults to the placeholder deck and `/tmp/layout-gallery.pptx`. Strips the deck's
slides, then adds one slide from every layout and fills its placeholders in idx order.

This is the check `mock_render.py` cannot do on its own: it resolves no placeholder
inheritance, so open the result in PowerPoint to confirm the type scale, the centred
title alignment, and the list row's gradient left edge. What it does prove headlessly
is that every layout is still usable — that adding a slide from it works and yields the
placeholders you expect. Run it after any change to the master.

Copy is per-layout sample text, not the real prompt strings, so a layout whose
placeholder count changed fails loudly instead of rendering half-filled.
"""
import os
import sys

from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DECK = os.path.join(HERE, "Valkey placeholder slides.pptx")
DEFAULT_OUT = "/tmp/layout-gallery.pptx"

TITLE_SET = ["Scaling Valkey past one core",
             "What the ceiling looks like and how to move it",
             "Madelyn Olson", "Valkey Maintainer and Engineer"]

SAMPLE = {
    "Valkey Title Only": ["Every custom slide is built on this"],
    "Valkey Title Band": TITLE_SET,
    "Valkey Title Top Banner": TITLE_SET,
    "Valkey Title Bottom Banner": TITLE_SET,
    "Valkey Insight": [
        "Steady state is a lie. The peak is the only number that matters.",
        "One idea per slide. A second bullet means a second slide."],
    "Valkey Insight Bold": [
        "You cannot provision your way out of a workload that moves.",
        "The loud version. Use it once, at the turn."],
    "Valkey Big Number": [
        "3x", "throughput on the same hardware",
        "Valkey 8.0, multi-threaded I/O. Source: valkey.io/blog/valkey-8-0-ga/"],
    "Valkey List Rows": [
        "Valkey 8.0 (September 2024)",
        "3x throughput with multi-threaded I/O",
        "25%+ memory reduction with embedded hashtable entries",
        "1M+ RPS on a single node",
        "Dual-channel replication for faster failovers",
        "Improved cluster bus reliability"],
    "Valkey Two Box": [
        "Two ways to grow", "Vertical scaling",
        "Add CPU and memory to one node. Simple to operate, bounded by the largest "
        "instance you can buy.",
        "Horizontal scaling",
        "Shard the keyspace across nodes. Scales past any single machine, at the cost "
        "of more moving parts."],
    "Valkey Three Box": [
        "What we optimize for", "Performance",
        "Throughput per core, and predictable tail latency under load.", "Hardening",
        "Safe defaults, fuzzing, and a shrinking surface area for untrusted input.",
        "Cost efficiency", "Same workload on smaller hardware means a smaller bill."],
    "Valkey Thank You": [
        "Thank you", "Madelyn Olson", "@reconditerose", "github.com/madolson"],
}


def build(deck=DEFAULT_DECK, out=DEFAULT_OUT):
    import shutil
    shutil.copyfile(deck, out)
    prs = Presentation(out)

    names = [lo.name for lo in prs.slide_layouts]
    unknown = [n for n in names if n not in SAMPLE]
    if unknown:
        raise SystemExit(f"no sample copy for: {unknown}. Add it to SAMPLE.")

    for i in range(len(prs.slides) - 1, -1, -1):        # highest index first
        prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
        del prs.slides._sldIdLst[i]

    for name in names:                                  # gallery order = master order
        slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(name))
        phs = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)
        texts = SAMPLE[name]
        if len(phs) != len(texts):
            raise SystemExit(
                f"{name}: layout offers {len(phs)} placeholders, SAMPLE has {len(texts)}")
        for ph, text in zip(phs, texts):
            ph.text_frame.text = text

    prs.save(out)
    Presentation(out)        # reopen: if it raises, PowerPoint would complain too
    print(f"wrote {out}: {len(names)} layouts, one slide each")
    for name in names:
        print(f"  {name}")
    return out


if __name__ == "__main__":
    build(*(sys.argv[1:3] or []))
