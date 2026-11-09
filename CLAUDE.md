# このリポジトリでの作業ルール

## Office ファイルのフォント（必須）
- Excel・PowerPoint・Word を作成・編集するときは、ユーザーの明示指定がない限り、フォントを必ず **Meiryo UI** に統一する。
  - 全セル・全テキストだけでなく、既定スタイル（Excel の Normal、Word の docDefaults）とテーマフォントも Meiryo UI にする。
  - xlsx / pptx / docx スキルの既定フォント（Arial・Calibri 等）より、このルールを優先する。
- LibreOffice で再計算・変換した後は、フォントが置き換わるので必ず再適用する。
  - 適用：`python skills/intloop-office-font/scripts/apply_meiryo.py <file>`
  - 確認：`python skills/intloop-office-font/scripts/check_fonts.py <file>`（Meiryo UI 以外が 0 件であること）
- 詳細は `skills/intloop-office-font/SKILL.md` を参照。

## 数値表記
- 稼働率など「率」の％表示は、特に指定がない限り小数第一位（`0.0%`）。
  丸めると値が誤って見えるもの（例：0.15%）はユーザーに確認する。
