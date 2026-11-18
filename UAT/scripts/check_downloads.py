# -*- coding: utf-8 -*-
"""UAT でダウンロードした CSV / JSON を、テストケースの期待結果に照らして自動検証する。

使い方:
    python check_downloads.py <ダウンロードファイルを集めたフォルダ>

ファイル名から対象ケースを判定する。判定できないファイルはスキップして一覧に出す。
committee CSV は委員の出力（RV-066）と HR の出力（RV-083）でファイル名が同じになり得るため、
ファイル名に "member" または "委員" を含めれば RV-066、"hr" を含めれば RV-083 として扱う。
"""
import csv
import io
import json
import re
import sys
from pathlib import Path

results = []


def report(case, fname, ok, msg):
    results.append((case, fname, "OK" if ok else "NG", msg))


def read_csv(p):
    raw = p.read_bytes()
    bom = raw.startswith(b"\xef\xbb\xbf")
    enc = None
    for e in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(e)
            enc = e
            break
        except UnicodeDecodeError:
            continue
    if enc is None:
        return None, None, "文字コードを判定できない"
    rows = list(csv.reader(io.StringIO(text)))
    # Excel で文字化けしない = UTF-8 BOM 付き、または Shift_JIS
    excel_ok = bom or enc == "cp932"
    note = "UTF-8(BOM付)" if bom else ("Shift_JIS" if enc == "cp932" else "UTF-8(BOMなし)")
    return rows, excel_ok, note


def flat(rows):
    return "\n".join(",".join(r) for r in rows)


def check_encoding(case, p, excel_ok, note):
    report(case, p.name, excel_ok, f"Excel で文字化けしないか: {note}")


def tm019(p):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report("TM-019", p.name, False, note)
    check_encoding("TM-019", p, ok, note)
    body = flat(rows)
    report("TM-019", p.name, "UAT-新入" in body and "太郎" in body, "名簿に UAT-新入 太郎 が含まれる")
    report("TM-019", p.name, "UAT-山本" in body, "名簿に UAT-山本 直樹（ICBP）が含まれる")
    report("TM-019", p.name, "Telecom" in body, "Practice（Telecom）の集計がある")


def tm071(p):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report("TM-071", p.name, False, note)
    check_encoding("TM-071", p, ok, note)
    body = flat(rows)
    report("TM-071", p.name, "受注済" in body and "有償" in body, "「合計（受注済・有償）」行がある")
    report("TM-071", p.name, "仮タグ" in body and ("見込" in body), "「合計（見込・計画・仮タグ込み）」行がある")
    for k in ("受注済", "見込", "計画", "仮タグ"):
        report("TM-071", p.name, k in body, f"確度「{k}」の案件行がある")
    months = re.findall(r"20\d\d-\d\d", ",".join(rows[0])) if rows else []
    report("TM-071", p.name, len(months) >= 2, f"月列が並んでいる（ヘッダーの月: {len(months)} 列）")


def tm074(p):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report("TM-074", p.name, False, note)
    check_encoding("TM-074", p, ok, note)
    hay = [r for r in rows if any("林" in c and "拓海" in c for c in r) or any(c == "林 拓海" for c in r)]
    report("TM-074", p.name, bool(hay), "林 拓海 の行がある")
    if hay:
        report("TM-074", p.name, any(re.fullmatch(r"90(\.0+)?%?", c.strip()) for c in hay[0]),
               f"林 拓海 が 90%（行: {','.join(hay[0])}）")
    report("TM-074", p.name, "UAT-山本" not in flat(rows) and "ICBP" not in flat(rows), "ICBP が含まれない")
    report("TM-074", p.name, "2026-10" in p.name, "ファイル名が ForestBaR_稼働率_2026-10.csv")


def tm079(p):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report("TM-079", p.name, False, note)
    check_encoding("TM-079", p, ok, note)
    need = ["姓", "名", "メールアドレス", "入社年月日", "部署", "役職", "雇用形態"]
    head = ",".join(rows[0]) if rows else ""
    missing = [k for k in need if k not in head]
    report("TM-079", p.name, not missing, "列（姓名・ヨミガナ・メール・入社年月日・部署・役職・雇用形態）" + (f" 不足: {missing}" if missing else ""))
    report("TM-079", p.name, "ヨミガナ" in head or "カナ" in head, "ヨミガナ列がある")
    jiro = [r for r in rows if any("次郎" in c for c in r)]
    report("TM-079", p.name, bool(jiro), "UAT-候補 次郎 の行がある")
    if jiro:
        line = ",".join(jiro[0])
        report("TM-079", p.name, "2027-01-01" in line or "2027/01/01" in line or "2027/1/1" in line, "入社年月日 2027-01-01")
        report("TM-079", p.name, "Technology" in line, "部署 Technology")
        report("TM-079", p.name, "正社員" in line, "雇用形態 正社員")


def committee(p, case):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report(case, p.name, False, note)
    check_encoding(case, p, ok, note)
    head = rows[0] if rows else []
    idx = next((i for i, h in enumerate(head) if "昇給率" in h), None)
    if case == "RV-066":
        report(case, p.name, idx is None, "委員の出力に「昇給率」列が無い")
    else:
        report(case, p.name, idx is not None, "HR の出力に「昇給率(参考 %)」列がある")
        if idx is not None:
            for name in ("渡辺", "山口"):
                rr = [r for r in rows if any(name in c for c in r)]
                val = rr[0][idx] if rr and len(rr[0]) > idx else "(行なし)"
                report(case, p.name, val.strip().rstrip("%") in ("7", "7.0"), f"{name} の昇給率 = {val}（期待 7）")


def rv082(p):
    rows, ok, note = read_csv(p)
    if rows is None:
        return report("RV-082", p.name, False, note)
    check_encoding("RV-082", p, ok, note)
    kondo = [r for r in rows if any("近藤" in c for c in r)]
    report("RV-082", p.name, bool(kondo) and "差戻" in ",".join(kondo[0]), "近藤 拓也 が「差戻し」")


def rv084(p):
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:  # noqa
        return report("RV-084", p.name, False, f"JSON として読めない: {e}")
    text = json.dumps(data, ensure_ascii=False)
    report("RV-084", p.name, "evaluationResults" in text, "evaluationResults がある")
    for k in ("committeeAdjusted", "adjustedTotal", "compositePct", "gradeByPct"):
        report("RV-084", p.name, k in text, f"{k} が含まれる")


def route(p):
    n = p.name.lower()
    if p.suffix.lower() == ".json":
        return rv084(p)
    if p.suffix.lower() != ".csv":
        return None
    if n.startswith("headcount_"):
        return tm019(p)
    if n.startswith("assignment_ledger_"):
        return tm071(p)
    if "forestbar" in n:
        return tm074(p)
    if n.startswith("smarthr_onboarding_"):
        return tm079(p)
    if "status" in n or "ステータス" in n:
        return rv082(p)
    if "member" in n or "委員" in n:
        return committee(p, "RV-066")
    if "hr" in n or "committee" in n or "評価統合" in n:
        return committee(p, "RV-083")
    return None


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    folder = Path(sys.argv[1])
    skipped = []
    for p in sorted(folder.iterdir()):
        if p.is_file():
            if route(p) is None and not any(r[1] == p.name for r in results):
                skipped.append(p.name)
    print(f"{'ケース':8} {'判定':4} 内容  [ファイル]")
    for case, fname, st, msg in results:
        print(f"{case:8} {st:4} {msg}  [{fname}]")
    if skipped:
        print("\n対象ケースを判定できずスキップ:", ", ".join(skipped))
    ng = sum(1 for r in results if r[2] == "NG")
    print(f"\n合計 {len(results)} 項目 / NG {ng}")


if __name__ == "__main__":
    main()
