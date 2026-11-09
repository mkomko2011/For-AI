---
name: intloop-office-font
description: "INTLOOP の Office ファイル共通フォントルール。Excel（.xlsx/.xlsm）・PowerPoint（.pptx）・Word（.docx）を作成・編集・修正するあらゆる依頼で必ず起動し、xlsx / pptx / docx / pptx-rules スキルと併用する。ユーザーから明示の指定がない限り、フォントは必ず Meiryo UI に統一する（既定スタイル・テーマ・全セル／全テキスト）。他スキルの既定フォント（Arial・Calibri 等）より本ルールを優先する。"
---

# INTLOOP Office フォントルール（Meiryo UI 統一）

## ルール（MUST）

1. **Excel・PowerPoint・Word のフォントは、ユーザーの明示指定がない限り必ず `Meiryo UI`。**
   - xlsx / pptx / docx スキルが推奨する Arial・Calibri 等の既定フォントより、本ルールを優先する。
   - 既存ファイルの編集でも、納品物は Meiryo UI に統一する（ユーザーが「元の書式のまま」と指定した場合を除く）。
2. **個々のセル・テキストだけでなく「既定」も Meiryo UI にする。**
   - Excel：`Normal` スタイル（既定フォント）、テーマフォント（major/minor の latin・ea）、全セル。
   - PowerPoint：テーマフォント（major/minor の latin・ea）、全ランのフォント（`latin`・`ea`）。
   - Word：既定スタイル（docDefaults の rFonts ascii/hAnsi/eastAsia）とテーマフォント。
3. **LibreOffice で再計算・変換した後は、必ずフォントを再適用する。**
   LibreOffice は保存時に既定フォントを Calibri に戻したり、日本語セルを代替フォント
   （WenQuanYi 等）に置き換えたりする。`recalc.py` の後に本スキルのスクリプトを実行すること。
4. **納品前に必ずチェックする。** `scripts/check_fonts.py` で Meiryo UI 以外が 0 件であることを確認してから納品する。

## 手順

```bash
# Excel：再計算（recalc.py）の後に実行。キャッシュ値を保ったまま styles.xml とテーマを書き換える
python scripts/apply_meiryo.py output.xlsx

# PowerPoint / Word も同じスクリプトで可（テーマ・既定・全ランを置換）
python scripts/apply_meiryo.py deck.pptx
python scripts/apply_meiryo.py doc.docx

# チェック（Meiryo UI 以外のフォントが残っていれば exit 1）
python scripts/check_fonts.py output.xlsx
```

openpyxl でブックを作る場合も、最初に既定フォントを設定しておく：

```python
from openpyxl.styles import Font
wb._named_styles['Normal'].font = Font(name='Meiryo UI', size=10)
```

python-pptx の場合は、ランごとに `run.font.name = "Meiryo UI"` に加え、
東アジア文字用に `a:ea typeface="Meiryo UI"` も設定する（`apply_meiryo.py` で一括適用可）。

## 関連する表記ルール

- 稼働率など「率」の％表示は、特に指定がない限り **小数第一位（`0.0%`）**。
  ただし小数第一位で丸めると値が誤って見えるもの（例：0.15% の逓増幅）は、ユーザーに確認する。
