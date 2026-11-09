#!/usr/bin/env python3
"""Office ファイル（xlsx/xlsm/pptx/docx）のフォントを Meiryo UI に統一する。
ZIP 内の XML を直接書き換えるため、Excel の計算済みキャッシュ値は保持される。
使い方: python apply_meiryo.py <file> [<out>]   （out 省略時は上書き）"""
import re, sys, zipfile, shutil, os, tempfile
FONT = "Meiryo UI"

def fix_xlsx_styles(s):
    s = re.sub(r'<name val="[^"]*"/>', f'<name val="{FONT}"/>', s)
    s = re.sub(r'<charset val="\d+"/>', '<charset val="128"/>', s)
    return s

def fix_drawingml(s):
    # theme / slides / charts: latin・ea・cs の typeface を置換（+mj-lt 等のテーマ参照も含め統一）
    s = re.sub(r'(<a:(?:latin|ea|cs)\b[^>]*?typeface=")[^"]*(")', rf'\g<1>{FONT}\2', s)
    s = re.sub(r'(<a:font script="[^"]*" typeface=")[^"]*(")', rf'\g<1>{FONT}\2', s)
    return s

def fix_wordml(s):
    s = re.sub(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")[^"]*(")', rf'\g<1>{FONT}\2', s)
    s = re.sub(r'\s?w:(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', '', s)
    return s

def run(path, out=None):
    out = out or path
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1]); os.close(fd)
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            d = zi.read(it.filename)
            n = it.filename
            if n.endswith('.xml'):
                s = d.decode('utf-8')
                if n == 'xl/styles.xml':
                    s = fix_xlsx_styles(s)
                if n.startswith(('xl/theme/', 'ppt/', 'word/theme/', 'xl/charts/', 'xl/drawings/')):
                    s = fix_drawingml(s)
                if n.startswith('word/'):
                    s = fix_wordml(s)
                d = s.encode('utf-8')
            zo.writestr(it, d)
    shutil.move(tmp, out)
    print(f"applied {FONT}: {out}")

if __name__ == '__main__':
    if len(sys.argv) < 2: sys.exit(__doc__)
    run(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
