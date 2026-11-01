"""LaTeX -> Word OMML through pandoc's texmath, post-processed for Word/WPS rendering.

usage:  tex2omml.py --batch IN.json OUT.json
  IN.json : [{"key": "...", "tex": "...", "display": bool, "size_pt": 10.5}, ...]
  OUT.json: {"<key>": {"omml": "<m:oMath ...>", "error": null | "message"}, ...}

Post-processing (see references/pitfalls.md, "Formula engine"):
- every m:r gets explicit Cambria Math / East-Asian font and the requested size (pandoc emits none,
  and WPS then renders digits and brackets in the body font, italic);
- zero-width spaces texmath inserts around n-ary operators are dropped;
- a leading space inside \\text{ ...} is kept as a no-break space (math zones drop plain spaces).
"""
import argparse
import json
import os
import re
import sys
import tempfile
import zipfile

import pypandoc

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def run_props(size_pt, east_asia):
    fonts = f'<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="{east_asia}"/>'
    if size_pt is None:    # inherit the paragraph size (plain mode)
        return f"<w:rPr>{fonts}</w:rPr>"
    hp = int(round(size_pt * 2))
    return f'<w:rPr>{fonts}<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/></w:rPr>'


def fix_math_runs(om, size_pt, east_asia):
    """Font/size/space fixes on every m:r of an OMML string (one equation or a whole document part)."""
    om = om.replace("\u200b", "")
    rpr = run_props(size_pt, east_asia)

    def fix_run(m):
        r = re.sub(r"<w:rPr>.*?</w:rPr>|<w:rPr\s*/>", "", m.group(0), flags=re.S)
        if re.search(r"<m:rPr\s*/>|<m:rPr>", r):
            r = re.sub(r"(<m:rPr\s*/>|<m:rPr>.*?</m:rPr>)", lambda x: x.group(1) + rpr, r, count=1, flags=re.S)
        else:
            r = r.replace("<m:r>", "<m:r>" + rpr, 1)
        # text runs (m:nor) keep a leading space visible
        if "<m:nor" in r:
            r = re.sub(r"(<m:t[^>]*>) ", "\\1\u00a0", r)
        return r

    om = re.sub(r"<m:r>.*?</m:r>", fix_run, om, flags=re.S)
    om = re.sub(r"<m:r>(?:(?!</m:r>).)*?<m:t[^>]*></m:t></m:r>", "", om, flags=re.S)
    return om


def postprocess(om, size_pt, east_asia):
    om = fix_math_runs(om, size_pt, east_asia)
    # declare namespaces on the root so the fragment can be imported on its own
    om = re.sub(r"^<m:oMath\b[^>]*>", f'<m:oMath xmlns:m="{M_NS}" xmlns:w="{W_NS}">', om, count=1)
    return om


def convert_batch(items, east_asia):
    """One pandoc call for all formulas. Each formula follows its own marker paragraph, so a
    formula pandoc cannot parse (it then emits plain text) is reported without shifting the rest."""
    parts = []
    for k, it in enumerate(items):
        parts.append(f"QQF{k:05d}QQ")
        parts.append(f"$${it['tex']}$$" if it["display"] else f"${it['tex']}$")
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "f.docx")
        pypandoc.convert_text("\n\n".join(parts), "docx", format="markdown", outputfile=out)
        xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8")
    paras = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
    by_marker = {}
    for idx, p in enumerate(paras):
        m = re.search(r"QQF(\d{5})QQ", p)
        if m and idx + 1 < len(paras):
            by_marker[int(m.group(1))] = paras[idx + 1]
    result = {}
    for k, it in enumerate(items):
        oms = re.findall(r"<m:oMath>.*?</m:oMath>", by_marker.get(k, ""), re.S)
        if len(oms) != 1 or "<m:t" not in oms[0]:
            result[it["key"]] = {"omml": None, "error": "pandoc produced no single math object (unsupported or malformed LaTeX)"}
            continue
        result[it["key"]] = {"omml": postprocess(oms[0], it.get("size_pt", 10.5), east_asia), "error": None}
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--batch", nargs=2, metavar=("IN", "OUT"), required=True)
    ap.add_argument("--east-asia-font", default="宋体")
    a = ap.parse_args()
    items = json.load(open(a.batch[0], encoding="utf-8-sig"))
    res = convert_batch(items, a.east_asia_font) if items else {}
    json.dump(res, open(a.batch[1], "w", encoding="utf-8"), ensure_ascii=False)
    bad = [k for k, v in res.items() if v["error"]]
    print(f"formulas: {len(items)}, errors: {len(bad)}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
