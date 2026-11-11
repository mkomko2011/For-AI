#!/usr/bin/env python3
"""藤川CSO向け方針資料 v5 を作る（利用者編集版 v4 を入力とする）。

- スライド2：P.8「ご判断いただきたいこと」をベースに、資料全体を1枚にまとめた
  エグゼクティブサマリーへ差し替える
- スライド9：メッセージ18pt・表見出し16ptにして他スライドとそろえる
Usage: build_ints_site_policy_v5.py <in_v4.pptx> <out_v5.pptx>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ints_pptx_helpers import (BLACK, BLUE, NAVY, WHITE, L, W, LINE,  # noqa: E402
                                MSO_ANCHOR, PP_ALIGN, Presentation, box,
                                set_bullet, style_run)
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Pt  # noqa: E402
import copy  # noqa: E402

LIGHT = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor(0xF2, 0xF2, 0xF2)

prs = Presentation(sys.argv[1])
slides = prs.slides

# ---------- スライド9：フォントを他スライドにそろえる ----------
s9 = slides[8]
assert s9.shapes.title.text_frame.text.startswith("Appendix")
for sh in s9.shapes:
    if not sh.has_text_frame:
        continue
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is None:
                continue
            if sh.is_placeholder and sh.top < 1000000 and sh.placeholder_format.type != 1:
                r.font.size = Pt(18)          # メッセージ 16 → 18
            elif r.font.size.pt == 14:
                r.font.size = Pt(16)          # 表見出し 14 → 16


# ---------- 新エグゼクティブサマリー ----------
def set_title(slide, text):
    tf = slide.shapes.title.text_frame
    tf.clear()
    r = tf.paragraphs[0].add_run()
    r.text = text
    style_run(r, 20, bold=True, color=BLACK)
    for eff in r._r.findall(".//" + qn("a:effectLst")):
        eff.getparent().remove(eff)


layout = slides[7].slide_layout
s = slides.add_slide(layout)
set_title(s, "エグゼクティブサマリー")
for ph in list(s.placeholders):
    idx = ph.placeholder_format.idx
    if idx == 11:
        tf = ph.text_frame
        tf.clear()
        r = tf.paragraphs[0].add_run()
        r.text = ("掲載情報の整理は11/13までに完了させ、今後は「保守」と「発信」を切り分け、"
                  "発信は本体Biz Insightとの連携から始めることを提案し、次の4点のご判断をいただきたい")
        style_run(r, 18, bold=True, color=NAVY)
        set_bullet(tf.paragraphs[0], False)
    elif idx == 10:
        ph._element.getparent().remove(ph._element)

# ご報告（現状）
y = 1700000
box(s, L, y, 1500000, 620000, ["ご報告"], 16, fill=NAVY, line=None, color=WHITE,
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(s, L + 1500000, y, W - 1500000, 620000,
    ["退職者4名の掲載は11/4に削除済み、全掲載者の突合と社員数等の更新は11/13までに完了し、"
     "以降は入退社フローと連動して毎月末に差分を確認する"],
    16, fill=LIGHT, line=None, anchor=MSO_ANCHOR.MIDDLE)

# ご判断いただきたいこと（P.8ベース）＋事務局案
y += 620000 + 100000
cw = [560000, 5200000, W - 560000 - 5200000]
rh_head, rh = 400000, 700000
for x, w, t in ((L, cw[0], "No."), (L + cw[0], cw[1], "ご判断いただきたいこと"),
                (L + cw[0] + cw[1], cw[2], "事務局案")):
    box(s, x, y, w, rh_head, [t], 16, fill=NAVY, line=None, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
y += rh_head
rows = [
    ("1", "サイトの目的をB2Bのリード獲得・ナーチャリングに置くか",
     "採用面の役割は一定程度果たしたため、インバウンドで案件を獲得する機能へ転換する"),
    ("2", "保守と発信を切り分け、担い手を決めること",
     "保守は手順化してスタッフ層へ展開し、発信はINTS（Corp Planning→来期マーケ担当）がオーナーとなる"),
    ("3", "本体Biz Insightとの役割分担（A 集約／B 独立／C 連携）",
     "C 連携から始めて本体マーコムとの協議を開始し、独立サイトへの投資は来期に判断する"),
    ("4", "保守課題をMC・情シス経由で進めること（費用発生の可能性あり）",
     "現行WordPressの脆弱性対応を先行し、CMS基盤の刷新は1〜3月に比較検討して来期に移行する"),
]
for n, q, a in rows:
    box(s, L, y, cw[0], rh, [n], 18, fill=NAVY, line=None, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, L + cw[0], y, cw[1], rh, [q], 16, anchor=MSO_ANCHOR.MIDDLE)
    box(s, L + cw[0] + cw[1], y, cw[2], rh, [a], 16, anchor=MSO_ANCHOR.MIDDLE)
    y += rh

# 進め方
y += 100000
gap = 120000
pw = (W - 2 * gap) // 3
steps = [
    ("11〜12月：整える", "掲載整理・本体マーコムと協議", NAVY),
    ("1〜3月：試す", "事例試行・CMS刷新の比較検討", NAVY),
    ("来期：本格運用", "マーケ担当採用・CMS基盤の刷新", BLUE),
]
for i, (h, d, c) in enumerate(steps):
    x = L + i * (pw + gap)
    box(s, x, y, pw, 640000, [[(h, True)], [(d, False)]], 16, fill=c, line=None,
        color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
print("bottom", y + 640000)

# ページ番号
tree = s.shapes._spTree
for sp in s.slide_layout.shapes._spTree.iter(qn("p:sp")):
    ph = sp.find(".//" + qn("p:ph"))
    if ph is not None and ph.get("type") == "sldNum":
        new = copy.deepcopy(sp)
        new.find(".//" + qn("p:cNvPr")).set("id", str(900 + len(s.shapes)))
        tree.append(new)
        break

# 並べ替え：新サマリーを2枚目に入れ、旧サマリー（旧スライド2）を削除
sldIdLst = slides._sldIdLst
ids = list(sldIdLst)
new_id, old_summary = ids[-1], ids[1]
assert slides[1].shapes.title.text_frame.text == "エグゼクティブサマリー"
sldIdLst.remove(new_id)
sldIdLst.insert(1, new_id)
prs.part.drop_rel(old_summary.rId)
sldIdLst.remove(old_summary)

prs.save(sys.argv[2])
print("saved", sys.argv[2])
