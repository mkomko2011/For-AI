# INTS UAT 自動確認

Talent / Review / Accounts / SFA の UAT（11/16〜20 週）のうち、閲覧と CSV 出力で確認できるケースを Playwright で自動確認する。

- 本番環境・共有データのため、**保存・提出・削除などデータを変更するボタンは押さない**（`lib.js` の `SAFE_DENY` でガード）
- 認証情報は環境変数 `INTS_UAT_EMAIL` / `INTS_UAT_PASSWORD` からだけ読む（ファイルやログには残さない。`.auth/` は git 管理外）

## 実行
```
node run.js            # 全ケース
node run.js RV-        # Review だけ
python3 build_report.py  # INTS_UAT_自動化分類と結果.xlsx を作り直す
```

## ファイル
- `cases.js` … 自動確認するケースと期待文言
- `INTS_UAT_自動化分類と結果.xlsx` … 全 197 件の分類（A 自動／B 自動化可／C 人の判断／D 下準備／E 全体影響）と結果
- `testdata/` … SmartHR CSV などのテストデータ、期待値の検算メモ
- `out/screens/` … 実行時のスクリーンショット
