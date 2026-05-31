#!/usr/bin/env python3
"""
fix-favicon-footer.py — one-time idempotent chrome sweep across all site pages.
  1) Ensures every page declares the canonical favicon (assets/icons/favicon.svg
     + .ico fallback), with depth-correct relative paths. Replaces any existing
     rel="icon"/"shortcut icon"/"apple-touch-icon" links so there are no stale
     refs (e.g. brand/odysee-mark-light.svg) and no missing declarations.
  2) Replaces the stale footer label "Brand v1.0" -> "Brand V5".
Idempotent: run again => no changes. Run from repo root: python fix-favicon-footer.py
"""
import re, sys
from pathlib import Path

ROOT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
FILES = [ROOT/"index.html"] + sorted((ROOT/"pages").glob("*.html"))

ICON_RE = re.compile(r'[ \t]*<link[^>]*rel=["\'](?:icon|shortcut icon|apple-touch-icon)["\'][^>]*>\s*\n?',
                     re.IGNORECASE)

def block(depth):
    return (f'<link rel="icon" type="image/svg+xml" href="{depth}assets/icons/favicon.svg">\n'
            f'<link rel="icon" type="image/x-icon" href="{depth}assets/icons/favicon.ico">\n')

def process(path):
    if not path.exists():
        return f"– missing  {path}"
    txt = path.read_text(encoding="utf-8")
    orig = txt
    depth = "" if path.parent == ROOT else "../"
    # 1. favicon: strip existing icon links, insert canonical block before </head>
    txt = ICON_RE.sub("", txt)
    if "</head>" in txt:
        txt = txt.replace("</head>", block(depth) + "</head>", 1)
    # 2. footer label — SCOPED to .site-footer__right only
    #    (avoids corrupting task names like "מיגרציה Brand v1.0" on project-plan)
    txt = re.sub(
        r'(<div class="site-footer__right"[^>]*>)(.*?)(</div>)',
        lambda m: m.group(1) + m.group(2).replace("Brand v1.0", "Brand V5") + m.group(3),
        txt, flags=re.S)
    if txt == orig:
        return f"= no-op   {path}"
    path.write_text(txt, encoding="utf-8")
    fav = "favicon✓"
    foot = "footer✓" if re.search(r'<div class="site-footer__right"[^>]*>.*?Brand v1\.0.*?</div>', orig, re.S) else ""
    return f"✓ updated {path}  {fav} {foot}"

if __name__ == "__main__":
    for f in FILES:
        print(process(f))
