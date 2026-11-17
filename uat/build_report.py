import json, re, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

res = json.load(open('out/results.json', encoding='utf-8'))
auto_ids = {r['id'] for r in res}

def ids(prefix, a, b): return [f"{prefix}-{i:03d}" for i in range(a, b + 1)]
ALL = ids('TM', 1, 86) + ids('RV', 1, 91) + ids('AC', 1, 20)

HUMAN = {  # 業務判断・内容の妥当性は人が見る
 'TM-037':'推薦内容の妥当性','TM-044':'マッチング結果・スコアの妥当性','TM-053':'ヒートマップの見え方','TM-055':'シグナル・要因の妥当性','TM-056':'後継候補の妥当性','TM-063':'Practice ビューの見え方',
 'RV-035':'昨年度比較・定性評価の読みやすさ','RV-036':'評価結果の開示内容の妥当性','RV-052':'定性評価の記入（評価者の判断）','RV-053':'昇格準備度の判断','RV-054':'VALUE 最終値の判断',
 'RV-074':'委員会確定値の判断','RV-075':'最終判定・昇格・報酬の判断','RV-077':'KPI 委員会評価の判断','AC-002':'テストデータの合意'}
PREP = {  # Claude が下準備・検算できる
 'TM-008':'退職テスト用の社員番号を一覧から抽出','TM-009':'列不足 CSV を生成','TM-010':'SmartHR 差分 CSV を生成','TM-067':'上限＋20% の境界値を事前計算',
 'RV-008':'計算例（67.1% → B）を Excel 式で検算','RV-020':'過去評価 CSV（FY25・林 拓海）を生成','RV-077':'KPI 3.5 → 62.5%・総合 % を検算','AC-011':'SmartHR CSV（3 行）を UTF-8 で生成'}
GLOBAL = {'TM-004':'全員のデータを初期化','TM-085':'全員のデータを初期化','TM-086':'ロールを元に戻す','RV-002':'全員のデータを初期化','RV-091':'全員のデータを初期化','AC-019':'テストユーザー停止','AC-020':'Huberitus へ削除依頼'}
ROLE_SWITCH = lambda i: (i.startswith('TM') and int(i[3:]) >= 22 and int(i[3:]) <= 63) or (i.startswith('RV') and (24 <= int(i[3:]) <= 70 or int(i[3:]) >= 86)) or (i.startswith('AC') and 13 <= int(i[3:]) <= 18)

def classify(i):
    if i in auto_ids or i + 'a' in auto_ids:
        return ('A 自動確認（実装済み・閲覧のみ）', '小出さんのアカウントで今すぐ実行可（データ変更なし）')
    if i in GLOBAL: return ('E 人が実行（全体に影響）', GLOBAL[i])
    if i in HUMAN: return ('C 人の判断が必要', HUMAN[i])
    if i in PREP: return ('D Claude が下準備・検算', PREP[i])
    if ROLE_SWITCH(i): return ('B 自動化可（ロール切替・データ変更あり）', 'ロール別のテスト用アカウント、または打鍵と同時実行の合意があれば自動化可')
    return ('B 自動化可（データ変更あり）', '本番の共有データを変更するため、実施タイミングを打鍵者と合わせて実行')

wb = Workbook()
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
COL = {'OK':'C6EFCE','NG':'FFC7CE','要目視':'FFEB9C','未実行':'D9D9D9'}

ws = wb.active; ws.title = '概要'
rows = [classify(i)[0] for i in ALL]
cats = sorted(set(rows))
ws.append(['INTS UAT 自動化・効率化の分類と自動確認結果']); ws['A1'].font = Font(bold=True, size=14)
ws.append([f'作成 {datetime.date.today()}｜対象 Talent 86・Review 91・Accounts 20 = {len(ALL)} 件'])
ws.append([])
ws.append(['分類', 'Talent', 'Review', 'Accounts', '合計'])
for c in cats:
    n = lambda p: sum(1 for i, r in zip(ALL, rows) if r == c and i.startswith(p))
    ws.append([c, n('TM'), n('RV'), n('AC'), n('')])
ws.append([])
ws.append(['自動確認の結果', *[f"{s}: {sum(1 for r in res if r['status']==s)}" for s in ['OK','NG','要目視','未実行']]])
ws.append(['※ 自動確認は閲覧・CSV 出力のみ。保存・提出・削除などデータを変更するボタンはプログラムで押せないようにしている'])
for c in range(1, 6): ws.cell(4, c).font = H; ws.cell(4, c).fill = HF
ws.column_dimensions['A'].width = 48
for c in 'BCDE': ws.column_dimensions[c].width = 12

ws2 = wb.create_sheet('ケース別分類')
ws2.append(['ケース No', '分類', '理由・補足', '自動確認の結果'])
rmap = {r['id']: r for r in res}
for i in ALL:
    c, why = classify(i)
    r = rmap.get(i) or rmap.get(i + 'a')
    ws2.append([i, c, why, r['status'] if r else ''])
    if r: ws2.cell(ws2.max_row, 4).fill = PatternFill('solid', fgColor=COL.get(r['status'], 'FFFFFF'))
for c in range(1, 5): ws2.cell(1, c).font = H; ws2.cell(1, c).fill = HF
for c, w in zip('ABCD', [11, 40, 60, 14]): ws2.column_dimensions[c].width = w
ws2.freeze_panes = 'A2'; ws2.auto_filter.ref = ws2.dimensions

ws3 = wb.create_sheet('自動確認結果')
ws3.append(['ID', '製品', '確認内容', '結果', '詳細・目視ポイント', '所要秒', '実行日時', 'スクリーンショット'])
for r in res:
    ws3.append([r['id'], r['app'], r['title'], r['status'], r.get('notes', ''), r['sec'], r['at'], (r.get('screenshot') or '').split('/uat/')[-1]])
    ws3.cell(ws3.max_row, 4).fill = PatternFill('solid', fgColor=COL.get(r['status'], 'FFFFFF'))
for c in range(1, 9): ws3.cell(1, c).font = H; ws3.cell(1, c).fill = HF
for c, w in zip('ABCDEFGH', [11, 9, 44, 9, 80, 8, 22, 28]): ws3.column_dimensions[c].width = w
for row in ws3.iter_rows(min_row=2):
    for cell in row: cell.alignment = Alignment(wrap_text=True, vertical='top')
ws3.freeze_panes = 'A2'
wb.save('INTS_UAT_自動化分類と結果.xlsx')
for c in cats: print(c, sum(1 for r in rows if r == c))
