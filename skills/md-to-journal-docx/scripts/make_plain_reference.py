"""Build the plain-mode style references from pandoc's own default reference.docx.

usage:  make_plain_reference.py [OUT_DIR]      (default: ../assets)
writes  plain_reference.docx     no first-line indent, space between paragraphs
        plain_reference_zh.docx  same, body paragraphs indented two characters (Chinese convention)

Re-run after a pandoc upgrade so the style set matches the pandoc version in use.

Style choices (neutral document look):
- body: Times New Roman / 宋体, 12 pt, 1.25 lines, 6 pt after; text colour black
- title 18 pt, headings 16 / 14 / 12 pt: bold, Times New Roman / 黑体, black, no theme colour
- table text 10.5 pt (via the Table style, so lists outside tables keep the body size)
- captions 10.5 pt centred, not italic; code Consolas 10 pt
- theme font and colour references removed, so WPS and Word render the same fonts
"""
import io
import os
import subprocess
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pypandoc  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
LATIN, BODY_EA, HEAD_EA, CODE = "Times New Roman", "宋体", "黑体", "Consolas"

# child order inside w:rPr / w:pPr (schema sequence) for the elements set here
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "color", "sz", "szCs", "lang"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "spacing", "ind", "contextualSpacing", "jc", "outlineLvl"]


def q(tag):
    return f"{{{W}}}{tag}"


def child(parent, tag, order):
    """Get or create w:tag inside parent at its schema position."""
    el = parent.find(f"w:{tag}", NS)
    if el is not None:
        return el
    el = etree.Element(q(tag))
    pos = order.index(tag)
    for i, c in enumerate(parent):
        name = etree.QName(c).localname
        if name in order and order.index(name) > pos:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def props(style, kind):
    tag, order = ("rPr", RPR_ORDER) if kind == "r" else ("pPr", PPR_ORDER)
    el = style.find(f"w:{tag}", NS)
    if el is None:
        el = etree.SubElement(style, q(tag))
        # schema order in w:style: … pPr, rPr, tblPr, trPr, tcPr, tblStylePr
        later = ("rPr", "tblPr", "trPr", "tcPr", "tblStylePr") if kind == "p" else ("tblPr", "trPr", "tcPr", "tblStylePr")
        for c in style:
            if etree.QName(c).localname in later and c is not el:
                c.addprevious(el)
                break
    return el, order


def set_fonts(rpr, latin, ea):
    f = child(rpr, "rFonts", RPR_ORDER)
    for k in list(f.attrib):
        del f.attrib[k]
    for a in ("ascii", "hAnsi", "cs"):
        f.set(q(a), latin)
    f.set(q("eastAsia"), ea)


def set_val(parent, tag, order, val=None, **attrs):
    el = child(parent, tag, order)
    if val is not None:
        el.set(q("val"), str(val))
    for k, v in attrs.items():
        el.set(q(k), str(v))
    return el


def drop(parent, *tags):
    for t in tags:
        for el in parent.findall(f"w:{t}", NS):
            parent.remove(el)


def style(root, sid):
    return root.find(f"w:style[@w:styleId='{sid}']", NS)


def build(styles_xml, zh):
    root = etree.fromstring(styles_xml)
    # remove every theme font/colour reference
    for el in root.iter(q("rFonts"), q("color")):
        for k in list(el.attrib):
            if "theme" in etree.QName(k).localname.lower():
                del el.attrib[k]
    for el in list(root.iter(q("color"))):
        if el.get(q("val")) not in (None, "auto") and el.getparent().getparent().get(q("styleId")) != "Hyperlink":
            el.getparent().remove(el)

    d = root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
    set_fonts(d, LATIN, BODY_EA)
    set_val(d, "sz", RPR_ORDER, 24)
    set_val(d, "szCs", RPR_ORDER, 24)
    dp = root.find("w:docDefaults/w:pPrDefault/w:pPr", NS)
    set_val(dp, "spacing", PPR_ORDER, after=120, line=300, lineRule="auto")

    for sid in ("BodyText", "FirstParagraph"):
        s = style(root, sid)
        if s is None:
            continue
        p, o = props(s, "p")
        sp = set_val(p, "spacing", o, before=0, after=120)
        sp.set(q("line"), "300"); sp.set(q("lineRule"), "auto")
        if zh:
            set_val(p, "ind", o, firstLineChars=200, firstLine=480)

    # table cells and tight lists (Compact) inherit Body Text: no first-line indent there
    s = style(root, "Compact")
    if s is not None and zh:
        p, o = props(s, "p")
        set_val(p, "ind", o, firstLineChars=0, firstLine=0)

    heads = {"Title": 36, "Subtitle": 28, "Heading1": 32, "Heading2": 28, "Heading3": 24}
    heads.update({f"Heading{n}": 24 for n in range(4, 10)})
    for sid, hp in heads.items():
        s = style(root, sid)
        if s is None:
            continue
        r, o = props(s, "r")
        set_fonts(r, LATIN, HEAD_EA)
        drop(r, "i", "iCs", "color")
        set_val(r, "b", o); set_val(r, "bCs", o)
        set_val(r, "sz", o, hp); set_val(r, "szCs", o, hp)
        p, po = props(s, "p")
        if sid.startswith("Heading"):
            set_val(p, "spacing", po, before=240, after=120)

    for sid in ("Caption", "TableCaption", "ImageCaption"):
        s = style(root, sid)
        if s is None:
            continue
        r, o = props(s, "r")
        drop(r, "i", "iCs")
        set_val(r, "sz", o, 21); set_val(r, "szCs", o, 21)
        p, po = props(s, "p")
        set_val(p, "jc", po, "center")
        if sid == "TableCaption":
            set_val(p, "keepNext", po)

    t = style(root, "Table")
    if t is not None:
        r, o = props(t, "r")
        set_val(r, "sz", o, 21); set_val(r, "szCs", o, 21)

    for sid in ("VerbatimChar", "SourceCode"):
        s = style(root, sid)
        if s is not None:
            r, o = props(s, "r")
            set_fonts(r, CODE, BODY_EA)
            set_val(r, "sz", o, 20); set_val(r, "szCs", o, 20)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
    os.makedirs(out_dir, exist_ok=True)
    default = subprocess.run([pypandoc.get_pandoc_path(), "--print-default-data-file", "reference.docx"],
                             capture_output=True, check=True).stdout
    for name, zh in (("plain_reference.docx", False), ("plain_reference_zh.docx", True)):
        zin = zipfile.ZipFile(io.BytesIO(default))
        path = os.path.join(out_dir, name)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/styles.xml":
                    data = build(data, zh)
                zout.writestr(item, data)
        print("wrote", os.path.abspath(path), f"(pandoc {pypandoc.get_pandoc_version()})")


if __name__ == "__main__":
    main()
