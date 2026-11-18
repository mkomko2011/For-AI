"""日報の「作業指示者」などの自由記述から、社員（氏名・メール・役職）を特定する。

参照元（上から順に優先）
  1. ①人事マスタ_INTS_FY27.xlsx「①社員マスタ」…INTS在籍者。役職から管理職（M/SM/D/MD/役員）を判定
  2. 社員証発行者リスト.xlsx …INTLOOP全社（@intloop.com のみ使用。役職なし）
どちらも Microsoft 365 コネクタの read_resource で取得したタブ区切りテキストを渡す。
"""
import re
import datetime as dt

NORM = str.maketrans({'髙': '高', '﨑': '崎', '邉': '辺', '邊': '辺', '齋': '斎', '齊': '斉',
                      '濵': '浜', '德': '徳', '瀨': '瀬', '𠮷': '吉', '　': ' '})
MANAGER_GRADES = {'M', 'SM', 'D', 'MD', '役員'}


def norm(s):
    return s.translate(NORM).strip()


def load_ints_master(text, today):
    """人事マスタ（①社員マスタ）から在籍者を読み込む。退職日が today 以前の人は除外。"""
    header, people, seen = None, [], set()
    for line in text.split('\n'):
        c = line.split('\t')
        if c[0].startswith('確定用'):
            header = c
            continue
        if not header or len(c) < 20 or not c[4].strip():
            continue
        r = dict(zip(header, c))
        retired = r.get('退職日', '').strip()
        if retired:
            try:
                if dt.datetime.strptime(retired[:10].replace('-', '/'), '%Y/%m/%d') <= today:
                    continue
            except ValueError:
                pass
        mail = r['メールアドレス'].strip()
        if mail in seen:
            continue
        seen.add(mail)
        parts = re.split(r'\s+', norm(r['氏名']))
        people.append(dict(氏名=r['氏名'].strip(), 姓=parts[0], 名=parts[1] if len(parts) > 1 else '',
                           メール=mail, 役職=r['役職'], 管理職=r['役職'] in MANAGER_GRADES,
                           Division=r.get('組織1： 所属Division', ''), Practice=r.get('組織2： 所属Practice', ''),
                           出典='INTS人事マスタ'))
    return people


def load_intloop_list(text, known_mails, ints_surnames):
    """社員証発行者リスト（氏名に区切りなし）。INTS人事マスタにいない @intloop.com の人だけ追加。
    姓と名の境目は、INTSの姓リストと一般的な姓の長さ（2文字）で推定する。"""
    people = []
    for line in text.split('\n'):
        c = line.split('\t')
        if len(c) < 3 or not c[2].strip().endswith('@intloop.com'):
            continue
        name, mail = norm(c[1]), c[2].strip()
        if mail in known_mails or not name:
            continue
        sur = next((s for s in sorted(ints_surnames, key=len, reverse=True) if name.startswith(s) and len(name) > len(s)), None)
        if not sur:
            sur = name[:2] if len(name) >= 3 else name[:1]
        people.append(dict(氏名=c[1].strip(), 姓=sur, 名=name[len(sur):], メール=mail, 役職='',
                           管理職=False, Division='', Practice='', 出典='INTLOOP社員証リスト'))
    return people


def resolve(text, people):
    """自由記述から人を特定する。戻り値: (hits, unresolved)
    hits: [(person or {'氏名': '要確認: ...'}, 判定方法)]、unresolved: 特定できなかった「〇〇さん」"""
    t = norm(text)
    hits, used = [], []

    def take(i, j, e, how):
        if any(a < j and i < b for a, b in used):
            return
        used.append((i, j))
        hits.append((e, how))

    # 1) 氏名（姓名連結）で一致。INTS → 全社の順
    for e in people:
        if e['名']:
            for m in re.finditer(re.escape(e['姓'] + e['名']), t):
                take(m.start(), m.end(), e, '氏名一致')
    # 2) 姓（＋（名））＋敬称・役職
    by_sur = {}
    for e in people:
        by_sur.setdefault(e['姓'], []).append(e)
    tail = r'(?:[（(]([^）)]+)[）)])?\s*(?=さん|様|氏|MD|SM|D|M|$|[、,\s→⇒])'
    for s, group in by_sur.items():
        for m in re.finditer(re.escape(s) + tail, t):
            if any(a <= m.start() < b for a, b in used):
                continue
            cand = [e for e in group if not m.group(1) or e['名'].startswith(norm(m.group(1)))]
            ints = [e for e in cand if e['出典'] == 'INTS人事マスタ']
            for pool, how in ((ints, '姓一致'), ([e for e in ints if e['管理職']], '姓一致(管理職優先)'), (cand, '姓一致(全社)')):
                if len(pool) == 1:
                    take(m.start(), m.end(), pool[0], how)
                    break
            else:
                hits.append(({'氏名': '要確認: ' + '/'.join(e['氏名'] for e in (ints or cand))}, '同姓あり'))
                used.append((m.start(), m.end()))
    # 3) 名だけ（例：【敦さん】）。INTSで一意な場合のみ
    for e in people:
        if e['名'] and e['出典'] == 'INTS人事マスタ':
            same = [x for x in people if x['名'] == e['名'] and x['出典'] == 'INTS人事マスタ']
            if len(same) == 1:
                for m in re.finditer(r'(?<![一-龥々])' + re.escape(e['名']) + r'(?=さん)', t):
                    take(m.start(), m.end(), e, '名一致')
    unresolved = [m.group(1) for m in re.finditer(r'([一-龥々]{1,3})さん', t)
                  if not any(a <= m.start() < b for a, b in used)]
    return hits, unresolved
