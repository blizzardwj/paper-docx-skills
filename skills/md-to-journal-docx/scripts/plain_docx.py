"""Plain mode (no journal template): Markdown -> DOCX with the bundled style reference
(assets/plain_reference[_zh].docx, made by make_plain_reference.py) and Word-native (OMML) equations. Writes a build report in the same shape as build_docx.ps1 so make_docx.ps1 can
run the same host save, counts and page renders.

usage:  plain_docx.py MANUSCRIPT.md OUT.docx REPORT_DIR [--east-asia-font 宋体] [--lang zh-CN]

Fails (exit 1) on what pandoc only warns about: LaTeX it could not convert (printed as raw TeX)
and pictures it could not find. All such lines are listed at once.
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # -I drops the script dir
import pypandoc  # noqa: E402
from lxml import etree  # noqa: E402
from tex2omml import fix_math_runs  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NSW = {"w": W}
ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
FATAL = ("Could not convert TeX math", "Could not fetch resource", "Could not find image",
         "Could not determine image size")


def sha256(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest().upper()


BORDERS = "<w:tblBorders>" + "".join(
    f'<w:{e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"


def add_table_borders(m):
    """Grid lines like a Markdown preview (pandoc's default table style has none). tblBorders must
    come before shd/tblLayout/tblCellMar/tblLook in w:tblPr."""
    tp = m.group(0)
    if "<w:tblBorders" in tp:
        return tp
    at = re.search(r"<w:(?:shd|tblLayout|tblCellMar|tblLook)\b|</w:tblPr>", tp).start()
    return tp[:at] + BORDERS + tp[at:]


def keep_table_together(m):
    """keepNext on every row but the last (references/pitfalls.md #13); a table taller than a page
    still breaks."""
    tbl = m.group(0)
    rows = re.findall(r"<w:tr[ >].*?</w:tr>", tbl, flags=re.S)
    for row in rows[:-1]:
        new = re.sub(r"<w:pPr>(<w:pStyle [^>]*/>)?", lambda p: p.group(0) + "<w:keepNext/>", row)
        new = re.sub(r"<w:p>(?!<w:pPr>)", "<w:p><w:pPr><w:keepNext/></w:pPr>", new)
        tbl = tbl.replace(row, new, 1)
    return tbl


PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]


def put(parent, tag, order, **attrs):
    """Set w:tag (with attributes) inside parent at its schema position, replacing an existing one."""
    for old in parent.findall(f"w:{tag}", NSW):
        parent.remove(old)
    el = etree.Element(f"{{{W}}}{tag}")
    for k, v in attrs.items():
        el.set(f"{{{W}}}{k}", str(v))
    pos = order.index(tag)
    for i, c in enumerate(parent):
        name = etree.QName(c).localname
        if name in order and order.index(name) > pos:
            parent.insert(i, el)
            return
    parent.append(el)


def ppr(p):
    el = p.find("w:pPr", NSW)
    if el is None:
        el = etree.Element(f"{{{W}}}pPr")
        p.insert(0, el)
    return el


def layout_fixes(xml):
    """Pictures and display equations: no first-line indent (pictures centred); tables centred;
    space between a table and the next paragraph."""
    root = etree.fromstring(xml.encode("utf-8"))
    body = root.find("w:body", NSW)
    for p in root.iter(f"{{{W}}}p"):
        has_text = any((t.text or "").strip() for t in p.iter(f"{{{W}}}t"))
        if p.find(".//w:drawing", NSW) is not None and not has_text:
            put(ppr(p), "ind", PPR_ORDER, firstLineChars=0, firstLine=0)
            put(ppr(p), "jc", PPR_ORDER, val="center")
        elif p.find(f"{{{M}}}oMathPara") is not None:
            put(ppr(p), "ind", PPR_ORDER, firstLineChars=0, firstLine=0)
    for tbl in body.iter(f"{{{W}}}tbl"):
        put(tbl.find("w:tblPr", NSW), "jc", TBLPR_ORDER, val="center")
        nxt = tbl.getnext()
        if nxt is not None and nxt.tag == f"{{{W}}}p":
            put(ppr(nxt), "spacing", PPR_ORDER, before=120, after=120)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")


def walk(node, counts, suspects):
    if isinstance(node, dict):
        t = node.get("t")
        if t == "Math":
            counts["display_formulas" if node["c"][0]["t"] == "DisplayMath" else "inline_formulas"] += 1
        elif t == "Table":
            counts["tables"] += 1
        elif t == "Image":
            counts["figures"] += 1
        elif t == "Str" and "$" in node["c"]:
            suspects.append(f"'$' left as text (not parsed as math: unbalanced braces, or a space after the "
                            f"opening / before the closing $): {node['c']}")
        elif t in ("RawInline", "RawBlock") and node["c"][0] in ("tex", "latex"):
            suspects.append(f"raw LaTeX outside math (dropped from DOCX): {node['c'][1][:80]}")
        for v in node.values():
            walk(v, counts, suspects)
    elif isinstance(node, list):
        for v in node:
            walk(v, counts, suspects)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manuscript")
    ap.add_argument("out_docx")
    ap.add_argument("report_dir")
    ap.add_argument("--east-asia-font", default="宋体")
    ap.add_argument("--lang", default="", help="document language; default: zh-CN if the text has CJK, else none")
    ap.add_argument("--allow-literal-dollar", action="store_true", help="the text really contains $ signs")
    ap.add_argument("--reference-doc", default="", help="style reference; default: assets/plain_reference[_zh].docx by language")
    a = ap.parse_args()

    src = os.path.abspath(a.manuscript)
    text = open(src, encoding="utf-8-sig").read()
    lang = a.lang or ("zh-CN" if re.search(r"[一-鿿]", text) else "")
    pandoc = pypandoc.get_pandoc_path()
    ref = a.reference_doc or os.path.join(ASSETS, "plain_reference_zh.docx" if lang.startswith("zh") else "plain_reference.docx")
    if not os.path.exists(ref):
        sys.exit(f"style reference missing: {ref} (run make_plain_reference.py)")
    # like a Markdown preview: alt text is not turned into an extra caption
    base = [pandoc, src, "-f", "markdown-implicit_figures", "--resource-path", os.path.dirname(src)]
    if lang:
        base += ["-M", f"lang={lang}"]

    ast = subprocess.run(base + ["-t", "json"], capture_output=True, check=True).stdout
    counts = {"display_formulas": 0, "inline_formulas": 0, "tables": 0, "figures": 0}
    suspects = []
    walk(json.loads(ast), counts, suspects)
    if a.allow_literal_dollar:
        suspects = [s for s in suspects if not s.startswith("'$'")]

    r = subprocess.run(base + ["-t", "docx", "--reference-doc", ref, "-o", a.out_docx], capture_output=True)
    # pandoc's missing-translation notices (zh.yaml, "The term Table …") don't affect the DOCX
    warnings = [ln.strip() for ln in r.stderr.decode("utf-8", "replace").splitlines()
                if ln.strip() and "translation" not in ln and "zh.yaml" not in ln]
    if r.returncode != 0:
        sys.exit("pandoc failed:\n" + "\n".join(warnings))
    fatal = suspects + [w for w in warnings if any(k in w for k in FATAL)]

    # equation fonts/spaces (references/pitfalls.md #20, #21); sizes inherit the paragraph
    tmp = a.out_docx + ".tmp"
    with zipfile.ZipFile(a.out_docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = fix_math_runs(data.decode("utf-8"), None, a.east_asia_font)
                xml = re.sub(r"<w:tblPr>.*?</w:tblPr>", add_table_borders, xml, flags=re.S)
                xml = re.sub(r"<w:tbl>.*?</w:tbl>", keep_table_together, xml, flags=re.S)
                xml = layout_fixes(xml)
                omath = len(re.findall(r"<m:oMath[ >]", xml))
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, a.out_docx)

    expected = counts["display_formulas"] + counts["inline_formulas"]
    if omath != expected and not fatal:
        fatal.append(f"{omath} equations in the DOCX, {expected} in the Markdown")
    os.makedirs(a.report_dir, exist_ok=True)
    report = {"mode": "plain (no template)", "math_engine": "pandoc", "pandoc_version": pypandoc.get_pandoc_version(),
              "style_reference": os.path.abspath(ref), "style_reference_sha256": sha256(ref),
              "lang": lang or None, "source_path": src, "source_sha256": sha256(src), "output_path": os.path.abspath(a.out_docx),
              "counts": counts, "omath_in_docx": omath, "warnings": warnings, "errors": fatal,
              "content_ledger_complete": None}
    json.dump(report, open(os.path.join(a.report_dir, "manuscript_build_report.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print(json.dumps(counts), f"omath={omath}", f"warnings={len(warnings)}")
    if fatal:
        print("ERRORS:\n  " + "\n  ".join(fatal))
        sys.exit(1)


if __name__ == "__main__":
    main()
