"""Render PDF regions for visual review: around text snippets, or whole pages.

usage: render_pages.py PDF OUTDIR [--pages 3,5] [snippet ...]
Snippets are searched in page text; a crop from just above the hit is saved per snippet.
"""
import argparse

import pymupdf

ap = argparse.ArgumentParser()
ap.add_argument("pdf"); ap.add_argument("outdir")
ap.add_argument("--pages", help="comma-separated 1-based page numbers to render whole")
ap.add_argument("snippets", nargs="*")
args = ap.parse_args()

doc = pymupdf.open(args.pdf)
print("pages:", doc.page_count)
for n in [int(p) for p in args.pages.split(",")] if args.pages else []:
    out = f"{args.outdir}/page{n:02d}.png"
    doc[n - 1].get_pixmap(dpi=110).save(out)
    print("page", n, "->", out)
for k, snip in enumerate(args.snippets):
    for page in doc:
        hits = page.search_for(snip)
        if not hits:
            continue
        r = hits[0]
        clip = pymupdf.Rect(page.rect.x0 + 30, max(0, r.y0 - 20), page.rect.x1 - 30, min(page.rect.y1, r.y0 + 170))
        out = f"{args.outdir}/crop{k:02d}_p{page.number + 1}.png"
        page.get_pixmap(dpi=150, clip=clip).save(out)
        print(f"snippet {k} -> {out}")
        break
    else:
        print(f"snippet {k} not found")
