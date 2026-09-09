#!/usr/bin/env python3
"""Insert SVG images into a python-pptx slide.

python-pptx cannot add SVGs (PIL cannot decode them), so this builds the OOXML
package parts by hand. PowerPoint renders the SVG through the `asvg:svgBlip`
extension and falls back to an embedded PNG in older viewers.

Requires a real PNG fallback. Generate one with:
    chrome --headless --screenshot=out.png --window-size=96,96 file:///abs/path/icon.svg
"""
import re
from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"


def _next_media_num(package):
    nums = []
    for part in package.iter_parts():
        name = str(getattr(part, "_partname", ""))
        m = re.search(r"image(\d+)\.", name)
        if m:
            nums.append(int(m.group(1)))
    return max(nums, default=0) + 1


_PART_CACHE = {}


def _image_parts(package, svg_path, png_fallback_path):
    """Create the PNG+SVG parts once per package and reuse them.

    The same logo can appear on many slides; without this each placement adds two
    fresh media parts and the file grows for no reason.
    """
    key = (id(package), svg_path)
    if key not in _PART_CACHE:
        with open(svg_path, "rb") as f:
            svg_bytes = f.read()
        with open(png_fallback_path, "rb") as f:
            png_bytes = f.read()
        n = _next_media_num(package)
        _PART_CACHE[key] = (
            Part(PackURI(f"/ppt/media/image{n}.png"), "image/png", package, png_bytes),
            Part(PackURI(f"/ppt/media/image{n + 1}.svg"), "image/svg+xml", package, svg_bytes),
        )
    return _PART_CACHE[key]


def _next_shape_id(spTree):
    """One past the highest shape id in this tree.

    Deriving the id from the media number instead looks fine and is not: media parts
    are shared across placements, and python-pptx hands out ids from this same
    sequence, so the two schemes collide as soon as pictures and shapes interleave.
    That produced duplicate ids on 12 of the 31 slides in the placeholder deck.
    """
    ids = [int(e.get("id")) for e in spTree.iter()
           if e.tag == qn("p:cNvPr") and e.get("id")]
    return max(ids, default=1) + 1


def add_svg(slide, svg_path, png_fallback_path, left, top, width, height, name="SVG Image"):
    """Add an SVG picture to `slide`. Coordinates and size are EMU ints.

    NOTE: Part's signature is Part(partname, content_type, package, blob).
    Getting that order wrong yields a 0-byte part and an "image cannot be
    displayed" box in PowerPoint.
    """
    slide_part = slide.part
    png_part, svg_part = _image_parts(slide_part.package, svg_path, png_fallback_path)
    png_rId = slide_part.relate_to(png_part, RT.IMAGE)
    svg_rId = slide_part.relate_to(svg_part, RT.IMAGE)

    shape_id = _next_shape_id(slide.shapes._spTree)
    pic = etree.SubElement(slide.shapes._spTree, qn("p:pic"))

    nvPicPr = etree.SubElement(pic, qn("p:nvPicPr"))
    cNvPr = etree.SubElement(nvPicPr, qn("p:cNvPr"))
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", name)
    cNvPicPr = etree.SubElement(nvPicPr, qn("p:cNvPicPr"))
    etree.SubElement(cNvPicPr, qn("a:picLocks")).set("noChangeAspect", "1")
    etree.SubElement(nvPicPr, qn("p:nvPr"))

    blipFill = etree.SubElement(pic, qn("p:blipFill"))
    blip = etree.SubElement(blipFill, qn("a:blip"))
    blip.set(qn("r:embed"), png_rId)
    ext = etree.SubElement(etree.SubElement(blip, qn("a:extLst")), qn("a:ext"))
    ext.set("uri", SVG_EXT_URI)
    etree.SubElement(ext, f"{{{SVG_NS}}}svgBlip").set(qn("r:embed"), svg_rId)
    etree.SubElement(etree.SubElement(blipFill, qn("a:stretch")), qn("a:fillRect"))

    spPr = etree.SubElement(pic, qn("p:spPr"))
    xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    e = etree.SubElement(xfrm, qn("a:ext"))
    e.set("cx", str(int(width)))
    e.set("cy", str(int(height)))
    prstGeom = etree.SubElement(spPr, qn("a:prstGeom"))
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom, qn("a:avLst"))

    return pic
