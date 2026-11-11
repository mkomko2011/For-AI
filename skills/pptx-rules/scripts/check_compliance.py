#!/usr/bin/env python3
"""
check_compliance.py — INTLOOP/JCP パワーポイント準拠性チェッカー

SKILL.md に記載されたルールに対してスライドを自動検証し、違反一覧を出力する。

Usage:
    python3 check_compliance.py --input presentation.pptx
    python3 check_compliance.py --input presentation.pptx --output report.md

Exit codes:
    0: 準拠（エラー0件）
    1: エラーあり
    2: スクリプトエラー
"""

import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("[ERROR] python-pptx がインストールされていません。")
    print("  pip install python-pptx --break-system-packages")
    sys.exit(2)


# ===== ルール定義 =====

APPROVED_FONTS = {"Meiryo UI", "+mj-ea", "+mn-ea", "+mj-lt", "+mn-lt", ""}

# INTS ブランドカラー（theme1.xml から抽出）
APPROVED_COLORS_HEX = {
    "000000",  # Text
    "FFFFFF",  # Background
    "033865",  # Corporate Navy
    "7F7F7F",  # Gray
    "0099CC",  # Accent1
    "0087C5",  # Accent2
    "6666FF",  # Accent3
    "C00000",  # Accent4
    "FF9900",  # Accent5
    "FFC000",  # Accent6
    # グレースケール許容（濃淡）
    "595959", "808080", "A6A6A6", "BFBFBF", "D9D9D9", "F2F2F2",
    "262626", "404040",
}

MIN_FONT_SIZE_PT = 14.0

# 禁止表現
BANNED_PHRASES = [
    "以下の通り",
    "以下のとおり",
    "詳細は下記",
    "詳細は以下",
    "詳細は別紙",
    "下記参照",
]

# 手入力の行頭記号（箇条書き機能を使わず文字で入力したもの）
MANUAL_BULLET_CHARS = ("・", "•", "●", "■", "□", "◆", "◇", "-", "－", "※")

# 語尾
DESU_MASU_RE = re.compile(r"(です|ます|でした|ません|でしょう)。")
DEARU_DA_RE = re.compile(r"(である|だ|であった|ではない)。")


def color_distance(hex1, hex2):
    """2つの6桁HEX文字列の最大チャネル差を返す。"""
    try:
        r1, g1, b1 = int(hex1[0:2], 16), int(hex1[2:4], 16), int(hex1[4:6], 16)
        r2, g2, b2 = int(hex2[0:2], 16), int(hex2[2:4], 16), int(hex2[4:6], 16)
        return max(abs(r1 - r2), abs(g1 - g2), abs(b1 - b2))
    except Exception:
        return 999


def is_approved_color(hex_color, tolerance=5):
    """ブランドカラー範囲内か判定。"""
    if not hex_color:
        return True
    hex_color = hex_color.upper().lstrip("#")
    if len(hex_color) != 6:
        return True  # 解析不能な値はスキップ
    if hex_color in APPROVED_COLORS_HEX:
        return True
    for approved in APPROVED_COLORS_HEX:
        if color_distance(hex_color, approved) <= tolerance:
            return True
    return False


def extract_rgb(color_obj):
    """python-pptx の ColorFormat から HEX を抽出。取れなければ None。"""
    try:
        if color_obj.type is None:
            return None
        rgb = color_obj.rgb
        if rgb is None:
            return None
        return str(rgb).upper()
    except Exception:
        return None


def iter_text_shapes(slide):
    """スライド内の全テキスト持ちシェイプを走査（グループ内も再帰）。"""
    def recurse(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from recurse(shape.shapes)
            else:
                yield shape
    yield from recurse(slide.shapes)


def check_presentation(pptx_path):
    """プレゼン全体を検証し、issues リストを返す。"""
    prs = Presentation(pptx_path)
    issues = []  # (level, slide_no, rule, detail)

    # スライド間統一チェック用
    title_sizes_per_slide = {}
    message_sizes_per_slide = {}
    suffix_style_counts = Counter()
    all_slide_text = defaultdict(list)

    for slide_idx, slide in enumerate(prs.slides, start=1):
        # 独立テキストボックス検出用カウンタ
        plain_textbox_count = 0
        shape_with_text_count = 0

        slide_title_size = None
        slide_message_candidates = []

        for shape in iter_text_shapes(slide):
            if not shape.has_text_frame:
                continue

            text = shape.text_frame.text.strip()
            if not text:
                continue

            # R3: 独立テキストボックス検出（図形じゃない純粋Textbox）
            is_plain_textbox = (
                shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX
            )
            if is_plain_textbox:
                plain_textbox_count += 1
            else:
                shape_with_text_count += 1

            # パラグラフ・ランを走査
            for para in shape.text_frame.paragraphs:
                # R12: 手入力の行頭記号（箇条書き機能を使うこと）
                ptext = "".join(r.text for r in para.runs).lstrip()
                if ptext.startswith(MANUAL_BULLET_CHARS) and not ptext.startswith("※"):
                    issues.append((
                        "ERROR", slide_idx, "R12-手入力の行頭記号",
                        f"行頭記号を文字で入力（箇条書き機能を使用すること）: '{ptext[:30]}'"
                    ))
                for run in para.runs:
                    # R1: フォント
                    font_name = run.font.name or ""
                    if font_name and font_name not in APPROVED_FONTS:
                        issues.append((
                            "ERROR", slide_idx, "R1-フォント",
                            f"非Meiryo UIフォント検出: '{font_name}' (text: '{run.text[:30]}')"
                        ))

                    # R6: フォントサイズ
                    if run.font.size is not None:
                        pt_size = run.font.size.pt
                        if pt_size < MIN_FONT_SIZE_PT:
                            issues.append((
                                "ERROR", slide_idx, "R6-最低サイズ",
                                f"{pt_size:.1f}pt < {MIN_FONT_SIZE_PT:.0f}pt (text: '{run.text[:30]}')"
                            ))

                    # R2: 色
                    hex_color = extract_rgb(run.font.color)
                    if hex_color and not is_approved_color(hex_color):
                        issues.append((
                            "ERROR", slide_idx, "R2-カラー",
                            f"ブランド外カラー #{hex_color} (text: '{run.text[:30]}')"
                        ))

            # R8: 禁止表現
            for banned in BANNED_PHRASES:
                if banned in text:
                    issues.append((
                        "ERROR", slide_idx, "R8-禁止表現",
                        f"'{banned}' を検出"
                    ))

            # タイトル/メッセージ候補判定（プレースホルダで判定）
            try:
                if shape.is_placeholder:
                    ph_type = shape.placeholder_format.type
                    # PP_PLACEHOLDER.TITLE = 13, CENTER_TITLE = 15
                    if ph_type in (13, 15):
                        if shape.text_frame.paragraphs:
                            for run in shape.text_frame.paragraphs[0].runs:
                                if run.font.size:
                                    slide_title_size = run.font.size.pt
                                    break
            except Exception:
                pass

            # メッセージ候補（タイトル直下の長文を推定）
            top_emu = shape.top if shape.top is not None else 0
            if 500000 < top_emu < 1700000 and len(text) > 20:
                slide_message_candidates.append((text, shape))

            # 語尾収集
            if DESU_MASU_RE.search(text):
                suffix_style_counts["desu_masu"] += 1
            if DEARU_DA_RE.search(text):
                suffix_style_counts["dearu_da"] += 1

            all_slide_text[slide_idx].append(text)

        # R11: ページ番号（表紙＝1枚目以外は必須）
        if slide_idx > 1:
            has_num = False
            for el in slide.shapes._spTree.iter():
                tag = el.tag.split("}")[-1]
                if tag == "ph" and el.get("type") == "sldNum":
                    has_num = True
                    break
                if tag == "fld" and el.get("type") == "slidenum":
                    has_num = True
                    break
            if not has_num:
                issues.append((
                    "ERROR", slide_idx, "R11-ページ番号なし",
                    "表紙以外のスライドにはページ番号（スライド番号プレースホルダー）が必須"
                ))

        # R3: 独立テキストボックス比率チェック
        if plain_textbox_count > 0 and shape_with_text_count > 0:
            ratio = plain_textbox_count / (plain_textbox_count + shape_with_text_count)
            if ratio > 0.5:
                issues.append((
                    "ERROR", slide_idx, "R3-独立TextBox",
                    f"独立テキストボックスが {plain_textbox_count}個 / 図形内テキスト {shape_with_text_count}個"
                    f"（図形とテキストを分離している可能性が高い）"
                ))

        # R9: メッセージ3行超チェック
        for msg_text, msg_shape in slide_message_candidates:
            line_count = msg_text.count("\n") + 1
            if line_count > 3:
                issues.append((
                    "ERROR", slide_idx, "R9-メッセージ3行超",
                    f"{line_count}行のメッセージ候補: '{msg_text[:50]}...'"
                ))

        # R10: 複数メッセージ候補
        if len(slide_message_candidates) >= 2:
            issues.append((
                "WARN", slide_idx, "R10-複数メッセージ疑い",
                f"メッセージ候補が{len(slide_message_candidates)}つ検出"
            ))

        if slide_title_size:
            title_sizes_per_slide[slide_idx] = slide_title_size

    # R4: タイトルサイズ統一チェック
    if len(set(title_sizes_per_slide.values())) > 1:
        size_summary = dict(Counter(title_sizes_per_slide.values()))
        issues.append((
            "WARN", 0, "R4-タイトルサイズ不一致",
            f"タイトルサイズが複数検出: {size_summary}"
        ))

    # R7: 語尾混在チェック
    if suffix_style_counts["desu_masu"] > 0 and suffix_style_counts["dearu_da"] > 0:
        issues.append((
            "WARN", 0, "R7-語尾混在",
            f"です/ます {suffix_style_counts['desu_masu']}箇所 / "
            f"である/だ {suffix_style_counts['dearu_da']}箇所 が混在"
        ))

    return issues, len(prs.slides)


def format_report(issues, total_slides, pptx_path):
    """Markdown形式のレポート生成。"""
    errors = [i for i in issues if i[0] == "ERROR"]
    warns = [i for i in issues if i[0] == "WARN"]

    lines = []
    lines.append(f"# 準拠性チェックレポート\n")
    lines.append(f"- 対象: `{pptx_path}`")
    lines.append(f"- スライド数: {total_slides}")
    lines.append(f"- **エラー: {len(errors)}件 / 警告: {len(warns)}件**\n")

    status = "✅ 合格" if not errors else "❌ 不合格（エラーあり）"
    lines.append(f"## 判定: {status}\n")

    if errors:
        lines.append("## ❌ エラー（必須修正）\n")
        lines.append("| スライド | ルール | 内容 |")
        lines.append("|---|---|---|")
        for level, slide_no, rule, detail in errors:
            slide_str = f"全体" if slide_no == 0 else f"p.{slide_no}"
            lines.append(f"| {slide_str} | {rule} | {detail} |")
        lines.append("")

    if warns:
        lines.append("## ⚠️ 警告（確認推奨）\n")
        lines.append("| スライド | ルール | 内容 |")
        lines.append("|---|---|---|")
        for level, slide_no, rule, detail in warns:
            slide_str = f"全体" if slide_no == 0 else f"p.{slide_no}"
            lines.append(f"| {slide_str} | {rule} | {detail} |")
        lines.append("")

    if not errors and not warns:
        lines.append("## 🎉 問題なし\n")
        lines.append("すべてのルールに準拠しています。")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="INTLOOP/JCP パワポ準拠性チェッカー")
    parser.add_argument("--input", required=True, help="チェック対象の .pptx ファイル")
    parser.add_argument("--output", help="レポート出力先（省略時は標準出力）")
    parser.add_argument("--max-issues", type=int, default=100,
                        help="最大表示件数（同種エラーの乱立防止）")
    args = parser.parse_args()

    pptx_path = Path(args.input)
    if not pptx_path.exists():
        print(f"[ERROR] ファイルが見つかりません: {pptx_path}")
        sys.exit(2)

    try:
        issues, total_slides = check_presentation(pptx_path)
    except Exception as e:
        print(f"[ERROR] 解析中にエラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)

    # 同種エラーの大量発生を抑制
    if len(issues) > args.max_issues:
        truncated_note = f"\n...(同様の違反を多数検出。上位{args.max_issues}件のみ表示)"
        issues = issues[:args.max_issues]
    else:
        truncated_note = ""

    report = format_report(issues, total_slides, pptx_path.name) + truncated_note

    if args.output:
        Path(args.output).write_text(report, encoding="utf-8")
        print(f"[OK] レポートを保存: {args.output}")
    else:
        print(report)

    errors = [i for i in issues if i[0] == "ERROR"]
    sys.exit(0 if not errors else 1)


if __name__ == "__main__":
    main()
