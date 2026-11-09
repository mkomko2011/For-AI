#!/usr/bin/env python3
"""Office ファイル内で Meiryo UI 以外のフォントを列挙する。違反があれば exit 1。
使い方: python check_fonts.py <file> [<file> ...]"""
import re, sys, zipfile
from collections import Counter
FONT = "Meiryo UI"
PATS = [r'<name val="([^"]*)"/>',                      # xlsx styles
        r'<a:(?:latin|ea|cs)\b[^>]*?typeface="([^"]*)"',  # theme / pptx / charts
        r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]*)"']       # docx
def check(path):
    bad = Counter()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not n.endswith('.xml'): continue
            s = z.read(n).decode('utf-8', 'ignore')
            for p in PATS:
                for f in re.findall(p, s):
                    if f and f != FONT: bad[(n, f)] += 1
    if bad:
        print(f"NG {path}")
        for (n, f), c in bad.most_common(): print(f"  {f!r} x{c}  in {n}")
    else:
        print(f"OK {path}: all fonts are {FONT}")
    return not bad
if __name__ == '__main__':
    if len(sys.argv) < 2: sys.exit(__doc__)
    sys.exit(0 if all([check(p) for p in sys.argv[1:]]) else 1)
