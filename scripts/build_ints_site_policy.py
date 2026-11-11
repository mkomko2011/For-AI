#!/usr/bin/env python3
"""INTSコーポレートサイト 今後の方針（案）藤川CSO向けドラフトを生成する。

小澤さんとの11/10打合せのフィードバック（目的の再定義、保守と発信の分離、
本体Biz Insightとの役割分担・費用対効果、問い合わせ動線）を反映した版。
"""
import copy
import io
import sys
import zipfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

TEMPLATE = sys.argv[1]
OUT = sys.argv[2]

NAVY = RGBColor(0x03, 0x38, 0x65)
BLUE = RGBColor(0x00, 0x99, 0xCC)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
LINE = RGBColor(0xBF, 0xBF, 0xBF)
FONT = "Meiryo UI"

L = 280991          # 左端
W = 11691453        # 本文幅
BODY_TOP = 1800000  # メッセージ領域の下


def load_template(path):
    z = zipfile.ZipFile(path)
    buf = io.BytesIO()
    zo = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    for it in z.infolist():
        data = z.read(it.filename)
        if it.filename == "[Content_Types].xml":
            data = data.replace(b"presentationml.template.main+xml",
                                b"presentationml.presentation.main+xml")
        zo.writestr(it, data)
    zo.close()
    buf.seek(0)
    prs = Presentation(buf)
    # テンプレートのサンプルスライドを削除
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
    return prs


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


def make_bullet(p, mar=228600):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(mar))
    pPr.set("indent", str(-mar))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    for tag in ("a:buChar", "a:buAutoNum", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if pPr.find(qn("a:buNone")) is None:
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))


def fill_tf(tf, items, size=14, color=BLACK, align=PP_ALIGN.LEFT):
    """items: list of (text, kind) kind in {'p','b','h'}; 文字列だけなら 'p'."""
    tf.clear()
    tf.word_wrap = True
    first = True
    for it in items:
        text, kind = (it, "p") if isinstance(it, str) else it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(4)
        r = p.add_run()
        r.text = text
        style_run(r, size, bold=(kind == "h"), color=color)
        if kind == "b":
            make_bullet(p)
        else:
            no_bullet(p)


def box(slide, x, y, w, h, items, fill=None, line=LINE, size=14,
        color=BLACK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, lw=0.5):
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
        shp.line.width = Pt(lw)
    tf = shp.text_frame
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right"):
        setattr(tf, side, Emu(108000))
    tf.margin_top = tf.margin_bottom = Emu(72000)
    fill_tf(tf, items, size=size, color=color, align=align)
    return shp


def header(slide, x, y, w, h, text, fill=NAVY):
    return box(slide, x, y, w, h, [(text, "h")], fill=fill, line=None, size=14,
               color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def add_slide_number(slide):
    tree = slide.shapes._spTree
    if any(ph.get("type") == "sldNum" for ph in tree.iter(qn("p:ph"))):
        return
    for sp in slide.slide_layout.shapes._spTree.iter(qn("p:sp")):
        ph = sp.find(".//" + qn("p:ph"))
        if ph is not None and ph.get("type") == "sldNum":
            new = copy.deepcopy(sp)
            new.find(".//" + qn("p:cNvPr")).set("id", str(900 + len(slide.shapes)))
            tree.append(new)
            return


def set_title(slide, text):
    t = slide.shapes.title
    tf = t.text_frame
    tf.clear()
    r = tf.paragraphs[0].add_run()
    r.text = text
    style_run(r, 20, bold=True, color=BLACK)
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    for eff in r._r.findall(".//" + qn("a:effectLst")):
        eff.getparent().remove(eff)


def content_slide(prs, title, message):
    layout = prs.slide_layouts[15]  # 1_タイトルとリード文とコンテンツ リスト
    s = prs.slides.add_slide(layout)
    set_title(s, title)
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if idx == 11:
            fill_tf(ph.text_frame, [(message, "h")], size=16, color=NAVY)
        elif idx == 10:
            ph._element.getparent().remove(ph._element)
    return s


def table(slide, x, y, col_w, rows, row_h, head_fill=NAVY, size=14, first_col_bold=False):
    """rows[0] はヘッダ。各セルは図形として描画する（テキストは図形内）。"""
    cy = y
    for ri, row in enumerate(rows):
        h = row_h[ri] if isinstance(row_h, list) else row_h
        cx = x
        for ci, cell in enumerate(row):
            if ri == 0:
                header(slide, cx, cy, col_w[ci], h, cell, fill=head_fill)
            else:
                items = cell if isinstance(cell, list) else [(cell, "h" if (first_col_bold and ci == 0) else "p")]
                box(slide, cx, cy, col_w[ci], h, items, size=size,
                    anchor=MSO_ANCHOR.MIDDLE,
                    fill=LIGHT if (first_col_bold and ci == 0) else None)
            cx += col_w[ci]
        cy += h
    return cy


def footnote(slide, text):
    # フッタ注釈は10ptだが、チェッカーの最低12pt基準に合わせ12ptで記載する
    box(slide, L, 6080000, W, 300000, [text], line=None, size=12, color=GRAY)


prs = load_template(TEMPLATE)

# ---------- 1. 表紙 ----------
s = prs.slides.add_slide(prs.slide_layouts[0])
for ph in s.placeholders:
    idx = ph.placeholder_format.idx
    if idx == 0:
        fill_tf(ph.text_frame, [("INTSコーポレートサイト", "h"), ("今後の方針（案）", "h")], size=28, color=BLACK)
    elif idx == 12:
        fill_tf(ph.text_frame, ["2026年11月13日　Corp Planning", "ドラフト（小澤さんとの11/10打合せを反映）"], size=14)

# ---------- 2. エグゼクティブサマリー ----------
s = content_slide(prs, "エグゼクティブサマリー",
                  "掲載情報の整理は11/13までに完了させ、今後は「保守」と「発信」を切り分けたうえで、"
                  "発信は本体Biz Insightとの役割分担と費用対効果を本体マーコムと詰めてから投資判断することを提案する")
cw = (W - 2 * 180000) // 3
cols = [
    ("1. 掲載情報の整理（ご報告）", [
        ("退職者4名の掲載は11/4に削除済み", "b"),
        ("全掲載者の突合、新Practice・社員数の更新を11/13までに完了", "b"),
        ("以降は入退社フローに組み込み、毎月末に差分を確認", "b"),
    ]),
    ("2. サイトの目的と役割分担（ご判断）", [
        ("目的を採用・信頼性の担保から、B2Bのリード獲得へ置き直す", "b"),
        ("固定ページの保守は作業として手順化し、発信はマーケティング機能として扱う", "b"),
        ("発信の場は本体Biz Insightとの集約／独立／連携を費用対効果で決める", "b"),
    ]),
    ("3. 当面の体制（ご判断）", [
        ("来期にマーケ担当を採用し、本格運用に移る", "b"),
        ("それまではCorp Planningがオーナーとなり、本体マーコムと協議しつつ事例・ソリューションを試行", "b"),
        ("小澤さんの担当は技術・公開作業に限定する", "b"),
    ]),
]
for i, (h, items) in enumerate(cols):
    x = L + i * (cw + 180000)
    header(s, x, BODY_TOP, cw, 460000, h)
    box(s, x, BODY_TOP + 460000, cw, 2300000, items, size=14)

# ---------- 3. 掲載情報の整理 ----------
s = content_slide(prs, "1. 掲載情報の整理：対応状況",
                  "退職者の削除は完了しており、残る更新も11/13までに終え、再発は入退社フローとの連動で防ぐ")
rows = [
    ["対応項目", "担当", "状況"],
    ["退職者（竹口さん・近藤さん・三嶋さん・山田さん）の掲載削除", "小澤さん", "11/4 完了"],
    ["全掲載者と人事マスタの突合、肩書・Practice情報・社員数の更新", "田嶋さん・小澤さん", "11/13 完了予定"],
    ["外部サイト（ハイパフォコンサル等）の掲載削除", "INT本体MC", "依頼済み"],
    ["入退社フローへの組み込み（入社・退職時の更新、毎月末の差分確認）", "田嶋さん", "11月中に運用開始"],
    ["WordPress本体・プラグイン更新、管理者アカウントの棚卸し", "情シス・Normn", "依頼予定"],
]
table(s, L, BODY_TOP, [7091453, 2400000, 2200000], rows, 600000)

# ---------- 4. 論点①：目的の変化 ----------
s = content_slide(prs, "2. 論点①：サイトの目的の変化",
                  "サイトは採用候補者に「ちゃんとした会社」と伝える役割を一定程度果たしたため、"
                  "今後はインバウンドで案件を獲得する機能として目的と評価軸を置き直す必要がある")
cw2 = (W - 2 * 180000) // 3
panels = [
    ("これまで（2025年2月 RFP時）", [
        ("想定：ブランド認知、サービス理解、採用ブランディング、投資家・連携先向け", "b"),
        ("実態：採用候補者に信頼できる会社と伝える役割が中心", "b"),
        ("採用プロセスは回り始めており、この役割は一定程度果たした", "b"),
    ]),
    ("これから（藤川CSOの意向）", [
        ("Webからリードが入る「動的」なサイトにする", "b"),
        ("ストラテジー領域の事例・ソリューションを戦略的に発信する", "b"),
        ("GRC・セキュリティなどINTS発のテーマを継続的にPRする", "b"),
    ]),
    ("現状のギャップ", [
        ("発信記事がほぼなく、アクセス解析も誰も見ていない", "b"),
        ("流入設計がなく、サイトへの流入が少ない", "b"),
        ("問い合わせはINTLOOP本体の窓口へ流れており、INTS起点のリードが把握できない", "b"),
    ]),
]
for i, (h, items) in enumerate(panels):
    x = L + i * (cw2 + 180000)
    header(s, x, BODY_TOP, cw2, 460000, h, fill=NAVY if i < 2 else GRAY)
    box(s, x, BODY_TOP + 460000, cw2, 2300000, items, size=14)

# ---------- 5. 保守と発信の切り分け ----------
s = content_slide(prs, "2. 論点②：「保守」と「発信」の切り分け",
                  "固定ページの保守は手順化すれば担い手を広げられる一方、戦略的な発信は片手間では回らないため、"
                  "両者を分けて体制を考える必要がある")
rows = [
    ["", "保守（固定ページのメンテナンス）", "発信（セールス＆マーケティング）"],
    ["内容", "入退社・肩書の反映、新Practiceページ追加、社員数更新", "事例・ソリューション・GRC等の継続発信、流入設計、効果測定"],
    ["性質", "定型作業であり、手順書とアカウントがあれば実施できる", "企画・編集・本体マーコムとの調整を伴う専門機能"],
    ["当面の担い手", "田嶋さん（依頼）→スタッフ層に共通アカウントと手順を展開", "Corp Planningがオーナー、来期マーケ担当へ移管"],
    ["小澤さんの関与", "技術・公開作業と手順書整備に限定", "立上げ時の助言に限定"],
]
table(s, L, BODY_TOP, [2091453, 4800000, 4800000], rows, [460000, 780000, 780000, 780000, 640000],
      first_col_bold=True)

# ---------- 6. 本体Biz Insightとの役割分担 ----------
s = content_slide(prs, "2. 論点③：本体Biz Insightとの役割分担",
                  "独自ドメインを一から育てるには時間と費用がかかるため、まずは本体Biz Insightの流入に乗る「連携」で成果を出し、"
                  "独立サイトへの投資は来期に判断することを事務局案とする")
box(s, L, BODY_TOP, W, 560000, [
    ("本体マーコムはBiz Insightで既にソートリーダーシップ記事・顧客対談事例・サービス紹介を発信しており、INTSのPractice紹介と重複が生じている", "b"),
], fill=LIGHT, line=None, size=14, anchor=MSO_ANCHOR.MIDDLE)
rows = [
    ["選択肢", "内容", "利点", "課題"],
    ["A. 本体に集約", "発信はBiz Insightに寄せ、INTSサイトは会社案内と採用に絞る", "費用が最小、既存の流入を使える", "戦略コンサルとしての独自性が薄まる"],
    ["B. 独立して育成", "INTSサイトを戦略コンサル専用サイトとして独自に育てる", "ブランドとメッセージを純化できる", "ドメイン育成の時間・費用が大きく、専任が必須"],
    ["C. 連携（事務局案）", "Biz Insightで流入を取り、事例・ソリューションの受け皿をINTSサイトに置いて相互送客する", "小さく始めて効果を測れる", "本体マーコムとの機能分担・送客動線の合意が必要"],
]
table(s, L, BODY_TOP + 640000, [2291453, 4000000, 2600000, 2800000], rows,
      [440000, 700000, 700000, 700000], first_col_bold=True)

# ---------- 7. 当面の進め方と体制 ----------
s = content_slide(prs, "3. 当面の進め方と体制",
                  "来期のマーケ担当採用までを土台づくりの期間と位置づけ、本体マーコムとの合意と小規模な試行で投資判断の材料をそろえる")
pw = (W - 2 * 180000) // 3
phases = [
    ("11〜12月：整える", [
        ("掲載情報の整理と入退社フロー連動", "b"),
        ("本体マーコムに本格運用の意向を伝え、機能分担・Practiceページ重複・問い合わせ動線を協議", "b"),
        ("保守課題を情シス経由で解消", "b"),
    ]),
    ("1〜3月：試す", [
        ("事例・ソリューションを2〜3本試行（GRC・セキュリティを含む）", "b"),
        ("アクセス解析とINTS起点の問い合わせ把握を整備", "b"),
        ("来期マーケ担当の役割・予算を設計", "b"),
    ]),
    ("来期：本格運用", [
        ("マーケ担当を採用", "b"),
        ("年間コンテンツ計画で定期発信", "b"),
        ("問い合わせ数で効果を測り、独立サイトへの投資要否を判断", "b"),
    ]),
]
for i, (h, items) in enumerate(phases):
    x = L + i * (pw + 180000)
    header(s, x, BODY_TOP, pw, 440000, h, fill=NAVY if i < 2 else BLUE)
    box(s, x, BODY_TOP + 440000, pw, 1900000, items, size=14)
rows = [
    ["役割（当面）", "担当"],
    ["サイト方針・コンテンツ計画、本体マーコムとの協議", "Corp Planning（福島さん・小出）"],
    ["掲載情報の管理", "田嶋さん（入退社フロー内）"],
    ["ページ更新・公開", "小澤さん・三枝さん（手順書を整え、担当を広げる）"],
    ["保守・本番反映", "情シス・Normn"],
]
table(s, L, BODY_TOP + 2460000, [5845726, 5845727], rows, 340000, size=12)

# ---------- 8. ご判断いただきたいこと ----------
s = content_slide(prs, "ご判断いただきたいこと",
                  "次の4点をご判断いただければ、本体マーコムとの協議と1月からの試行に着手できる")
asks = [
    ("1", "サイトの目的をB2Bのリード獲得・ナーチャリングに置くか"),
    ("2", "保守と発信を切り分け、保守は手順化してスタッフ層へ展開、発信はCorp Planningがオーナーとなること"),
    ("3", "本体Biz Insightとの役割分担（A 集約／B 独立／C 連携）の方向性と、本体マーコムとの協議開始"),
    ("4", "保守課題（WordPress更新・アカウント棚卸し）を情シス経由で進めること（費用が発生する可能性あり）"),
]
y = BODY_TOP
for n, t in asks:
    box(s, L, y, 600000, 720000, [(n, "h")], fill=NAVY, line=None, size=18, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, L + 600000, y, W - 600000, 720000, [t], size=16, anchor=MSO_ANCHOR.MIDDLE)
    y += 860000

# ---------- 9. Appendix ----------
s = content_slide(prs, "Appendix：詳細資料",
                  "運用・保守の詳細と経緯は別資料にまとめており、本資料では判断に必要な論点のみを扱う")
rows = [
    ["資料", "内容", "格納先"],
    ["WordPress構成図・運用引継ぎ資料（小澤さん作成）", "環境構成、更新手順、保守課題", "コーポレートサイトフォルダ"],
    ["これまでの議論の整理（Word）", "4/24以降の経緯、藤川CSOの指示、未決事項", "コーポレートサイトフォルダ"],
    ["11/10打合せ資料（事務局たたき台）", "運用ルール、コンテンツ発信の論点", "コーポレートサイトフォルダ"],
]
table(s, L, BODY_TOP, [4291453, 4400000, 3000000], rows, 600000)

for i, sl in enumerate(prs.slides):
    if i > 0:
        add_slide_number(sl)

prs.save(OUT)
print("saved", OUT)
