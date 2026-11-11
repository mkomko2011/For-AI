#!/usr/bin/env python3
"""藤川CSO向け方針資料（利用者編集版）のスライド7「当面の進め方と体制」を改定する。

- インフラ整備の表現を「CMS基盤の刷新（WordPressからAI機能搭載のSaaS型CMSへ）」に改める
- スライド7の各フォントを2pt上げ、はみ出さないよう配置と文言を調整する
他のスライドは変更しない。
Usage: revise_ints_site_policy_slide7.py <in.pptx> <out.pptx>
"""
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

NAVY = RGBColor(0x03, 0x38, 0x65)
BLUE = RGBColor(0x00, 0x99, 0xCC)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xBF, 0xBF, 0xBF)
FONT = "Meiryo UI"
L, W = 280991, 11691453


def style_run(run, size, bold=False, color=BLACK):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:latin"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def set_bullet(p, on, mar=228600):
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if on:
        pPr.set("marL", str(mar))
        pPr.set("indent", str(-mar))
        pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))
    else:
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))


def box(slide, x, y, w, h, paras, size, fill=None, line=LINE, color=BLACK,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullet=False):
    """paras: list of paragraphs; each paragraph is str or list of (text, bold)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.5)
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(90000)
    tf.margin_top = tf.margin_bottom = Emu(54000)
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(4 if bullet else 0)
        runs = [(para, None)] if isinstance(para, str) else para
        for text, b in runs:
            r = p.add_run()
            r.text = text
            style_run(r, size, bold=bool(b) if b is not None else (fill is not None and color == WHITE), color=color)
        set_bullet(p, bullet)
    return shp


prs = Presentation(sys.argv[1])
s = prs.slides[6]
assert s.shapes.title.text_frame.text.startswith("3. 当面の進め方と体制")

# 既存の本文図形（タイトル・メッセージ・ページ番号以外）を削除
for sh in list(s.shapes):
    if sh.is_placeholder:
        continue
    sh._element.getparent().remove(sh._element)

# メッセージ（16pt → 18pt、他スライドと統一）
for ph in s.placeholders:
    if ph.placeholder_format.idx == 11 or ph.placeholder_format.type == 2:
        if ph.top < 1000000:
            tf = ph.text_frame
            tf.clear()
            r = tf.paragraphs[0].add_run()
            r.text = ("来期のマーケ担当採用とCMS基盤の刷新までを土台づくりの期間と位置づけ、"
                      "本体マーコムとの合意と小規模な試行で投資判断の材料をそろえる")
            style_run(r, 18, bold=True, color=NAVY)
            set_bullet(tf.paragraphs[0], False)

# フェーズ（見出し・本文 16pt → 18pt）
gap = 120000
pw = (W - 2 * gap) // 3
top, hh, bh = 1600000, 440000, 2520000
phases = [
    ("11〜12月：整える", NAVY, [
        "掲載情報の整理と入退社フロー連動",
        "本体マーコムと機能分担・問い合わせ動線を協議",
        "現行WordPressの脆弱性・保守課題を情シス経由で解消",
    ]),
    ("1〜3月：試す", NAVY, [
        "事例・ソリューションを2〜3本試行（GRC含む）",
        "アクセス解析・問い合わせ把握を整備",
        "マーケ担当の役割・予算を設計",
        "CMS刷新の選択肢を比較検討（セキュリティ・保守負荷・AI活用）",
    ]),
    ("来期：本格運用", BLUE, [
        "マーケ担当を採用し、年間計画で定期発信",
        "問い合わせ数で効果を測り、独立サイトへの投資を判断",
        "CMS基盤の刷新（WordPressからAI機能搭載のSaaS型CMSへ移行）",
    ]),
]
for i, (head, fill, items) in enumerate(phases):
    x = L + i * (pw + gap)
    box(s, x, top, pw, hh, [head], 18, fill=fill, line=None, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, top + hh, pw, bh, items, 18, bullet=True)

# 役割表（16pt → 18pt）
rows = [
    ("役割（当面）", "担当"),
    ("サイト方針・コンテンツ計画・本体マーコム協議", [("Corp Planning（福島さん・小出）", False)]),
    ("掲載情報の管理", [("田嶋さん（入退社フロー内で削除等を実施）", False)]),
    ("ページ更新・公開", [("小澤さん・三枝さん", False), ("（手順書を整え、担当を広げる）", True)]),
    ("保守・本番反映", [("情シス・Normn", False)]),
]
cw = [5545726, 6145727]
y = top + hh + bh + 60000
rh = 380000
for ri, (c1, c2) in enumerate(rows):
    if ri == 0:
        box(s, L, y, cw[0], rh, [c1], 18, fill=NAVY, line=None, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        box(s, L + cw[0], y, cw[1], rh, [c2], 18, fill=NAVY, line=None, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    else:
        box(s, L, y, cw[0], rh, [c1], 18, anchor=MSO_ANCHOR.MIDDLE)
        box(s, L + cw[0], y, cw[1], rh, [c2], 18, anchor=MSO_ANCHOR.MIDDLE)
    y += rh
print("table bottom", y)

prs.save(sys.argv[2])
print("saved", sys.argv[2])
