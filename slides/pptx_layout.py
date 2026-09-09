#!/usr/bin/env python3
"""Add slide layouts to a presentation's slide master.

python-pptx cannot do this: `SlideLayouts` only exposes `get_by_name`, `index` and
`remove`, so a design that should live in the master has to be hand-drawn onto every
slide instead. This module builds the `slideLayoutN.xml` part directly and wires it
into the master, which is the only missing piece.

Two entry points:

    layout = new_layout(prs, "Valkey Two Box")   # empty layout, registered on the master
    canvas = drawing_canvas(layout)              # gives the layout a slide-like shapes API
    add_placeholder(layout, "body", 1, ...)      # the editable bits

`drawing_canvas` is what lets the existing slide-drawing helpers target a layout
unchanged: pass the canvas anywhere a slide is expected. Static shapes drawn this way
are inherited by every slide on the layout and are not editable there, so draw only
decoration and put every piece of replaceable text in a placeholder.
"""
from lxml import etree

from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.parts.slide import SlideLayoutPart
from pptx.shapes.shapetree import SlideShapes

# sldLayoutId/@id is ST_SlideLayoutId: 2147483648 .. 4294967295.
_MIN_LAYOUT_ID = 2147483648

_EMPTY_LAYOUT = """<p:sldLayout %s type="cust" preserve="1">
  <p:cSld name="%s">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


def _next_layout_num(package):
    nums = []
    for part in package.iter_parts():
        name = str(getattr(part, "partname", ""))
        if name.startswith("/ppt/slideLayouts/slideLayout"):
            nums.append(int(name.rsplit("slideLayout", 1)[1].split(".")[0]))
    return max(nums, default=0) + 1


def new_layout(prs, name, preserve=True):
    """Create an empty slide layout named `name` and register it on the master.

    Three things have to line up or PowerPoint reports a repair: the layout part
    itself, a relationship in each direction between it and the master, and a
    `p:sldLayoutId` entry in the master's `sldLayoutIdLst`.
    """
    if prs.slide_layouts.get_by_name(name) is not None:
        raise ValueError(f"layout {name!r} already exists")

    master_part = prs.slide_master.part
    package = master_part.package
    partname = PackURI(f"/ppt/slideLayouts/slideLayout{_next_layout_num(package)}.xml")
    blob = (_EMPTY_LAYOUT % (nsdecls("p", "a", "r"), name)).encode("utf-8")

    # load()'s signature is (partname, content_type, package, blob). Same argument
    # order trap as Part() in pptx_svg.py; getting it wrong fails inside lxml.
    layout_part = SlideLayoutPart.load(partname, CT.PML_SLIDE_LAYOUT, package, blob)
    if not preserve:
        layout_part._element.set("preserve", "0")

    # Both directions. The layout->master rel is what makes SlideLayout.slide_master
    # resolve; the master->layout rel is what the sldLayoutId below points at.
    layout_part.relate_to(master_part, RT.SLIDE_MASTER)
    rId = master_part.relate_to(layout_part, RT.SLIDE_LAYOUT)

    lst = prs.slide_master._element.get_or_add_sldLayoutIdLst()
    used = [int(e.get("id")) for e in lst if e.get("id")]
    entry = parse_xml(f'<p:sldLayoutId {nsdecls("p", "r")} '
                      f'id="{max(used, default=_MIN_LAYOUT_ID - 1) + 1}" r:id="{rId}"/>')
    lst.append(entry)

    return layout_part.slide_layout


def drawing_canvas(layout):
    """A slide-shaped object that draws into `layout`.

    `layout.shapes` is a read-only `LayoutShapes`, but `SlideShapes` only needs an
    spTree and a parent with a `.part`, and `SlideLayoutPart` supports everything it
    reaches for. Wrapping it here means the slide helpers work on a layout as-is.
    """
    return _Canvas(layout)


class _Canvas:
    def __init__(self, layout):
        self.layout = layout
        self.shapes = SlideShapes(layout._element.cSld.spTree, layout)

    @property
    def part(self):
        return self.layout.part


# --- placeholders ---------------------------------------------------------------

_ALIGN = {"l": "l", "ctr": "ctr", "r": "r"}
_ANCHOR = {"t": "t", "ctr": "ctr", "b": "b"}


def _scheme(colour):
    return MSO_THEME_COLOR.to_xml(colour)


def _next_shape_id(spTree):
    ids = [int(e.get("id")) for e in spTree.iter() if e.tag == qn("p:cNvPr")]
    return max(ids, default=1) + 1


def add_placeholder(layout, ph_type, idx, left, top, width, height, prompt,
                    size, bold=False, colour=MSO_THEME_COLOR.DARK_1, font="Arial",
                    align="l", anchor="t", spacing=None, italic=False,
                    insets=(0, 0, 0, 0), fill=None, name=None):
    """Add one placeholder to `layout`.

    `prompt` is the custom prompt PowerPoint shows in place of "Click to add text";
    it is not copied onto new slides, so the slide starts empty and correctly styled.
    `insets` is (left, top, right, bottom) in EMU. `fill` is XML for a fill element,
    for the cases where the shape behind the text has to travel with it.
    """
    spTree = layout._element.cSld.spTree
    line_spc = (f'<a:lnSpc><a:spcPct val="{int(spacing * 100000)}"/></a:lnSpc>'
                if spacing is not None else "")
    rpr = (f'<a:defRPr sz="{int(size * 100)}" b="{1 if bold else 0}" '
           f'i="{1 if italic else 0}">'
           f'<a:solidFill><a:schemeClr val="{_scheme(colour)}"/></a:solidFill>'
           f'<a:latin typeface="{font}"/></a:defRPr>')
    li, ti, ri, bi = insets
    ph_idx = "" if ph_type in ("title", "ctrTitle") else f' idx="{idx}"'

    xml = f"""<p:sp {nsdecls("p", "a")}>
      <p:nvSpPr>
        <p:cNvPr id="{_next_shape_id(spTree)}" name="{name or prompt}"/>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="{ph_type}"{ph_idx} hasCustomPrompt="1"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{int(left)}" y="{int(top)}"/>
          <a:ext cx="{int(width)}" cy="{int(height)}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        {fill or ""}
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" lIns="{int(li)}" tIns="{int(ti)}" rIns="{int(ri)}"
                  bIns="{int(bi)}" anchor="{_ANCHOR[anchor]}"><a:normAutofit/></a:bodyPr>
        <a:lstStyle>
          <a:lvl1pPr algn="{_ALIGN[align]}" marL="0" indent="0">{line_spc}<a:buNone/>{rpr}</a:lvl1pPr>
        </a:lstStyle>
        <a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{prompt}</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>"""
    sp = parse_xml(xml)
    spTree.append(sp)
    return sp


def solid_fill(colour):
    """A flat theme-colour fill, for a placeholder whose background travels with it.

    Flat, not a gradient. A gradient can fake a one-sided accent edge in a single
    shape, which is tempting when a card has to travel with its text, but it is the
    wrong answer twice: PowerPoint shows the shape as gradient-filled in its own UI,
    which is not what a flat card is, and the hard stop it depends on is fragile. Put
    the accent in its own static shape beside the placeholder instead, not inside its
    fill, and keep the placeholder clear of it so the layout shape is not covered.
    """
    return f'<a:solidFill><a:schemeClr val="{_scheme(colour)}"/></a:solidFill>'


def strip_stock_layouts(prs, keep_prefix):
    """Remove every layout whose name does not start with `keep_prefix`.

    `SlideLayouts.remove()` refuses a layout that is the basis for a slide, and that
    refusal is the safety net rather than a nuisance: if this raises, a slide is still
    sitting on a stock layout, and the fix is to build it on one of yours, not to skip
    the delete. Removal drops the master->layout relationship, so the layout part and
    any media only it referenced (the stock layouts' full-bleed background pictures)
    fall out of the package on save.
    """
    doomed = [lo for lo in prs.slide_layouts if not lo.name.startswith(keep_prefix)]
    for layout in doomed:
        prs.slide_layouts.remove(layout)
    return [lo.name for lo in doomed]


def audit(prs, names):
    """Assert the added layouts are well-formed enough for PowerPoint to open."""
    problems = []
    for name in names:
        layout = prs.slide_layouts.get_by_name(name)
        if layout is None:
            problems.append(f"{name}: not registered on the master")
            continue
        if layout.slide_master is not prs.slide_master:
            problems.append(f"{name}: does not resolve back to the master")
        seen, ids = set(), set()
        for ph in layout.placeholders:
            key = (ph.element.ph_type, ph.placeholder_format.idx)
            if key in seen:
                problems.append(f"{name}: duplicate placeholder {key}")
            seen.add(key)
        for e in layout._element.cSld.spTree.iter():
            if e.tag != qn("p:cNvPr"):
                continue
            if e.get("id") in ids:
                problems.append(f"{name}: duplicate shape id {e.get('id')}")
            ids.add(e.get("id"))
        if not seen:
            problems.append(f"{name}: no placeholders, so nothing is editable")
        # a:gsLst positions must be distinct and increasing. Duplicates are undefined,
        # and PowerPoint resolves them by dropping a stop, which silently turns a hard
        # edge into a full-width wash. Nothing else catches this.
        for gsLst in layout._element.cSld.spTree.iter(qn("a:gsLst")):
            pos = [int(gs.get("pos", 0)) for gs in gsLst.findall(qn("a:gs"))]
            if any(b <= a for a, b in zip(pos, pos[1:])):
                problems.append(f"{name}: gradient stops not strictly increasing: {pos}")
    if problems:
        raise RuntimeError("layout audit failed:\n  " + "\n  ".join(problems))
    return len(names)
