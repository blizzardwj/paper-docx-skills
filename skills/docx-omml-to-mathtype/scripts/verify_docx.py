"""Check a converted .docx against the original. Exit code 0 only if everything matches.

usage: verify_docx.py ORIGINAL.docx CONVERTED.docx
"""
import re
import sys
import zipfile


def load(path):
    z = zipfile.ZipFile(path)
    return z, z.read("word/document.xml").decode("utf-8")


def bookmarks(xml):
    return set(re.findall(r'<w:bookmarkStart\b[^>]*w:name="([^"]+)"', xml))


def plain(xml):
    """Concatenated text of all w:t runs outside equations (robust to run splitting / whitespace
    between tags, which WPS changes when it re-saves)."""
    xml = re.sub(r"<m:oMathPara>.*?</m:oMathPara>|<m:oMath>.*?</m:oMath>|<w:object>.*?</w:object>", "", xml, flags=re.S)
    text = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", xml))
    return re.sub(r"\s+", " ", text).strip()   # hosts may collapse non-preserved spaces


def main():
    zo, orig = load(sys.argv[1])
    zc, conv = load(sys.argv[2])
    n_orig = len(re.findall(r"<m:oMath>", orig))
    pre = len(re.findall(r'ProgID="Equation\.DSMT', orig))
    checks = []

    left = len(re.findall(r"<m:oMath>", conv))
    checks.append(("no OMML equations left", left == 0, f"{left} left (original {n_orig})"))
    dsmt = len(re.findall(r'ProgID="Equation\.DSMT', conv))
    checks.append(("one MathType object per equation", dsmt == n_orig + pre, f"{dsmt} Equation.DSMT objects, expected {n_orig + pre}"))
    emb = [n for n in zc.namelist() if n.startswith("word/embeddings/") and not n.endswith("/")]
    checks.append(("embedded OLE parts present", len(emb) >= dsmt, f"{len(emb)} parts"))
    markers = re.findall(r"QQMTEQ\d+QQ", conv)
    checks.append(("no markers left", not markers, ", ".join(markers[:10]) or "none"))
    lost = sorted(bookmarks(orig) - bookmarks(conv))
    checks.append(("all bookmarks kept", not lost, f"lost: {lost[:10]}" if lost else f"{len(bookmarks(conv))} kept"))
    po, pc = plain(orig), plain(conv)
    detail = f"{len(po)} chars"
    if po != pc:
        i = next((k for k in range(min(len(po), len(pc))) if po[k] != pc[k]), min(len(po), len(pc)))
        a = max(0, i - 30)
        detail = f"first diff at {i}: orig {po[a:i+30]!r} | conv {pc[a:i+30]!r}"
    checks.append(("body text unchanged", po == pc, detail))

    for name, ok, d in checks:
        print(f"[{'OK' if ok else 'FAIL'}] {name}: {d}")
    sys.exit(0 if all(ok for _, ok, _ in checks) else 1)


if __name__ == "__main__":
    main()
