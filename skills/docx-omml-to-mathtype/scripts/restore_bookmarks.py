"""Re-wrap each MathType object run with the bookmarks its OMML equation had in the original.
(Pasting over a marker that a bookmark exactly encloses deletes the bookmark.)

usage: restore_bookmarks.py ORIGINAL.docx CONVERTED.docx OUT.docx
Assumes every equation was converted: the n-th <w:object> run in CONVERTED is equation n of ORIGINAL.
"""
import re
import sys
import zipfile

EQ_RE = re.compile(r"<m:oMathPara>.*?</m:oMathPara>|<m:oMath>.*?</m:oMath>", re.S)
START_RE = re.compile(r'<w:bookmarkStart\b[^>]*w:id="(\d+)"[^>]*/>')
END_RE = re.compile(r'<w:bookmarkEnd\b[^>]*w:id="(\d+)"[^>]*/>')
OBJ_RUN_RE = re.compile(r"<w:r\b(?:(?!<w:r\b).)*?<w:object\b.*?</w:object>.*?</w:r>", re.S)


def eq_bookmarks(xml):
    """For each equation: bookmark start tags fully owned by it (adjacent before/inside, ended inside/adjacent after)."""
    out = []
    for m in EQ_RE.finditer(xml):
        before = re.search(r"(?:<w:bookmarkStart\b[^>]*/>)*$", xml[:m.start()]).group(0)
        after = re.match(r"(?:<w:bookmarkEnd\b[^>]*/>)*", xml[m.end():]).group(0)
        frag = m.group(0)
        tags = re.findall(r"<w:bookmarkStart\b[^>]*/>", before) + re.findall(r"<w:bookmarkStart\b[^>]*/>", frag)
        ids = START_RE.findall(before) + START_RE.findall(frag)
        end_ids = set(END_RE.findall(frag) + END_RE.findall(after))
        out.append([t for t, i in zip(tags, ids) if i in end_ids])
    return out


def main():
    orig_path, conv_path, out_path = sys.argv[1:4]
    orig = zipfile.ZipFile(orig_path).read("word/document.xml").decode("utf-8")
    zin = zipfile.ZipFile(conv_path)
    conv = zin.read("word/document.xml").decode("utf-8")

    marks = eq_bookmarks(orig)
    runs = list(OBJ_RUN_RE.finditer(conv))
    pre_existing = len(re.findall(r'ProgID="Equation\.DSMT', orig))
    if len(runs) != len(marks) + pre_existing or pre_existing:
        sys.exit(f"{len(runs)} object runs vs {len(marks)} equations (+{pre_existing} pre-existing objects); "
                 "not restoring bookmarks automatically")

    present = set(re.findall(r'<w:bookmarkStart\b[^>]*w:name="([^"]+)"', conv))
    next_id = max([int(i) for i in START_RE.findall(conv) + START_RE.findall(orig)] + [0]) + 1
    pieces, pos, restored = [], 0, 0
    for run, starts in zip(runs, marks):
        pre, post = [], []
        for st in starts:
            name = re.search(r'w:name="([^"]+)"', st).group(1)
            if name in present:
                continue
            pre.append(f'<w:bookmarkStart w:id="{next_id}" w:name="{name}"/>')
            post.append(f'<w:bookmarkEnd w:id="{next_id}"/>')
            next_id += 1; restored += 1
        pieces += [conv[pos:run.start()], "".join(pre), run.group(0), "".join(post)]
        pos = run.end()
    pieces.append(conv[pos:])

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = "".join(pieces).encode("utf-8")
            zout.writestr(item, data)
    print(f"object runs: {len(runs)}, bookmarks restored: {restored}")


if __name__ == "__main__":
    main()
