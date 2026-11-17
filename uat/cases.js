// 自動確認するテストケース（閲覧・出力のみ。データは変更しない）
// 前提ロール: 小出（Talent/Review = コーポレート・HR 管理者、Accounts = 管理者）
// status: OK / NG / 要目視（自動で判定しきれない部分がある）
const { APPS, bodyText, checkTexts, safeClick, download } = require('./lib');

async function visit(page, app, p) {
  await page.goto(APPS[app] + p, { waitUntil: 'networkidle' });
  return bodyText(page);
}
function judge(chk, extraNote = '', partial = false) {
  const notes = [];
  if (chk.missing.length) notes.push('見つからない文言: ' + chk.missing.map(String).join(' / '));
  if (chk.unexpected.length) notes.push('出てはいけない文言: ' + chk.unexpected.map(String).join(' / '));
  if (extraNote) notes.push(extraNote);
  const status = chk.missing.length || chk.unexpected.length ? 'NG' : partial ? '要目視' : 'OK';
  return { status, notes: notes.join('｜') };
}
const ERR = [/Application error|Internal Server Error|Unhandled Runtime Error|500 \|/];

module.exports = [
  // ---------- ログイン画面（未ログイン） ----------
  { id: 'TM-001a', app: 'talent', auth: false, title: 'Talent ログイン画面の文言と再設定リンク',
    run: async (pg) => { const t = await visit(pg, 'talent', '/login');
      const href = await pg.locator('a[href*="dealreview"]').first().getAttribute('href').catch(() => null);
      return judge(checkTexts(t, ['DealReview と同じアカウント（メールアドレス + パスワード）でログインしてください', 'DealReview のログイン画面']), href ? `再設定リンク: ${href}` : 'DealReview へのリンクが無い'); } },
  { id: 'RV-001a', app: 'review', auth: false, title: 'Review ログイン画面の文言と再設定リンク',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/login'), ['DealReview と同じアカウント（メールアドレス + パスワード）でログインしてください', 'DealReview のログイン画面'])) },
  { id: 'AC-001a', app: 'accounts', auth: false, title: 'Accounts ログイン画面の文言',
    run: async (pg) => judge(checkTexts(await visit(pg, 'accounts', '/login'), ['INTS Accounts', 'DealReview のログイン画面'])) },
  { id: 'AUTH-GUARD', app: 'talent', auth: false, title: '未ログインで保護画面を開くとログイン画面へ戻る（4 製品）',
    run: async (pg) => {
      const bad = [];
      for (const [app, p] of [['talent', '/admin/users'], ['talent', '/assignments'], ['review', '/admin/fy'], ['review', '/review/committee'], ['accounts', '/users'], ['sfa', '/']]) {
        await pg.goto(APPS[app] + p, { waitUntil: 'networkidle' });
        if (!/\/login/.test(pg.url())) bad.push(`${app}${p} → ${pg.url()}`);
      }
      return { status: bad.length ? 'NG' : 'OK', notes: bad.join(' / ') };
    } },

  // ---------- INTS Talent（コーポレート） ----------
  { id: 'TM-001', app: 'talent', title: 'ログイン後ダッシュボード（組織ダッシュボード＋管理）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'talent', '/dashboard'), ['管理'], ERR), 'ヘッダーの氏名が DealReview の氏名か目視', true) },
  { id: 'TM-002', app: 'talent', title: 'ロール・人物紐付け一覧（4 名・注記）',
    run: async (pg) => { const t = await visit(pg, 'talent', '/admin/users');
      return judge(checkTexts(t, ['INTS Accounts', /福島/, /小出/, /田嶋/, /東/], ERR), 'ロール列（田嶋・東 = 社員（本人））を目視', true); } },
  { id: 'TM-007', app: 'talent', title: 'コーポレートのサイドバー（①〜⑥・管理）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'talent', '/dashboard'), [/①/, /②/, /③/, /④/, /⑤/, /⑥/, '管理'], ERR)) },
  { id: 'TM-018', app: 'talent', title: '共通マスタ（組織／役職／区分／ID 体系／SoR）',
    run: async (pg) => { let t = await visit(pg, 'talent', '/hr/masters'); const miss = [];
      for (const tab of ['組織', '役職', '区分', 'ID 体系', '正（SoR）マトリクス']) { if (!(await safeClick(pg, tab))) miss.push(tab); }
      await safeClick(pg, '区分'); t = await bodyText(pg);
      const chk = checkTexts(t, ['正社員', 'ICBP'], ERR);
      await safeClick(pg, '役職'); const t2 = await bodyText(pg);
      const icInt = /(^|\s)(IC|INT)(\s|$)/.test(t2);
      return judge(chk, [miss.length ? 'タブが見つからない: ' + miss.join(',') : '', icInt ? '役職に IC / INT らしき表記あり（要確認）' : ''].filter(Boolean).join('｜'), true); } },
  { id: 'TM-019', app: 'talent', title: '人員構成 CSV（BOM・ファイル名）',
    run: async (pg) => { await visit(pg, 'talent', '/org'); const d = await download(pg, /人員構成 CSV/);
      if (!d) return { status: 'NG', notes: '「人員構成 CSV」ボタンが見つからない' };
      const ok = /^headcount_\d{4}-\d{2}-\d{2}\.csv$/.test(d.name) && d.bom;
      return { status: ok ? '要目視' : 'NG', notes: `${d.name}｜BOM=${d.bom}｜行数=${d.rows}｜ヘッダー: ${d.header.slice(0, 120)}｜UAT データの有無は目視` }; } },
  { id: 'TM-021', app: 'talent', title: 'システム連携の契約・状態（モック／未接続）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'talent', '/admin/integration'), ['SmartHR', /DealReview|DR/, /SFA|仮タグ/, 'ForestBaR', /モック|未接続/], ERR)) },
  { id: 'TM-024', app: 'talent', title: '通知一覧の表示',
    run: async (pg) => judge(checkTexts(await visit(pg, 'talent', '/notifications'), ['通知'], ERR), '未読件数とヘッダーの一致は目視', true) },
  { id: 'TM-069', app: 'talent', title: '有償稼働率 速報タブの表示',
    run: async (pg) => { await visit(pg, 'talent', '/assignments'); const c = await safeClick(pg, '有償稼働率 速報');
      return judge(checkTexts(await bodyText(pg), [/速報/, /Practice/], ERR), c ? '' : 'タブが見つからない', true); } },
  { id: 'TM-071', app: 'talent', title: 'アサイン管理表 CSV（財務報告用）',
    run: async (pg) => { await visit(pg, 'talent', '/assignments'); const d = await download(pg, /CSV 出力/);
      if (!d) return { status: 'NG', notes: '「CSV 出力」ボタンが見つからない' };
      const okName = /^assignment_ledger_\d{4}-\d{2}\.csv$/.test(d.name);
      const hasKinds = ['受注済', '見込', '計画', '仮タグ'].filter((k) => !d.text.includes(k));
      return { status: okName && d.bom && !hasKinds.length ? 'OK' : 'NG', notes: `${d.name}｜BOM=${d.bom}｜行数=${d.rows}${hasKinds.length ? '｜確度の値が無い: ' + hasKinds.join(',') : ''}` }; } },
  { id: 'TM-075', app: 'talent', title: '月次締め：将来月は締められない（閲覧のみ）',
    run: async (pg) => { const t = await visit(pg, 'talent', '/utilization');
      return judge(checkTexts(t, [/月次締め|稼働実績/], ERR), '対象月 26/12 の選択は UI 確認後に自動化（現状は画面表示のみ確認）', true); } },

  // ---------- INTS Review（HR 管理者） ----------
  { id: 'RV-001', app: 'review', title: 'ログイン後マイページ（HR は管理メニューあり）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/review'), ['管理'], ERR), 'ヘッダーの氏名・ロールを目視', true) },
  { id: 'RV-005', app: 'review', title: 'HR 管理者のサイドバー（管理・評価ステータス・委員会・分析）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/review'), ['管理', '評価ステータス', '評価委員会', '評価分析'], ERR)) },
  { id: 'RV-006', app: 'review', title: 'FY27 年間スケジュール（閲覧のみ、保存しない）',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY27'); await safeClick(pg, '基本・年間スケジュール');
      return judge(checkTexts(await bodyText(pg), ['目標設定', '中間レビュー', '総合評価', '受付中'], ERR), '運用メモの保存は手動で実施', true); } },
  { id: 'RV-007', app: 'review', title: '評価モデル（階層ミックス・稼働率 KPI・閾値）の値',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY27'); await safeClick(pg, '評価モデル');
      return judge(checkTexts(await bodyText(pg), ['70', '30', '40', '60', '85', '80', '1.8', '4.2'], ERR), '閾値逆転時の保存不可は手動で確認', true); } },
  { id: 'RV-008', app: 'review', title: '達成度・判定ルールの既定値',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY27'); await safeClick(pg, '達成度・判定ルール（Excel 方式）');
      return judge(checkTexts(await bodyText(pg), ['150', '-35', '90', '70', '50', '20'].map((v) => v.replace('-', '')), ERR), '計算例（コンピ 3／88.2／85 → 67.1% B）は UI 確認後に自動入力', true); } },
  { id: 'RV-011', app: 'review', title: '評価対象・委員会（委員 5 名）',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY27'); await safeClick(pg, '評価対象・委員会');
      return judge(checkTexts(await bodyText(pg), ['森田', '高橋', '岡本', '松井', '石川', 'MD'], ERR)); } },
  { id: 'RV-012', app: 'review', title: 'カウンセラーアサイン（林 拓海 → 木下 慎一）',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY27'); await safeClick(pg, 'カウンセラーアサイン（FY 紐付け）');
      return judge(checkTexts(await bodyText(pg), ['林 拓海', '木下 慎一'], ERR), '林の行のタイトル C を目視', true); } },
  { id: 'RV-013', app: 'review', title: 'PJ アサイン・バリデーション（FY26・エラーで落ちない）',
    run: async (pg) => { await visit(pg, 'review', '/admin/fy'); await safeClick(pg, 'FY26'); await safeClick(pg, 'PJ アサイン・バリデーション');
      return judge(checkTexts(await bodyText(pg), ['エラー', '警告'], ERR)); } },
  { id: 'RV-064', app: 'review', title: '評価委員会 横断一覧（列の確認）',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/review/committee'), ['実績値', '委員会', 'KPI', '準備度', 'マトリクス', '定性', '最終', '昇格'], ERR), '判定分布・Practice 別平均の値は目視', true) },
  { id: 'RV-083', app: 'review', title: '委員会 CSV（HR は昇給率列あり）',
    run: async (pg) => { await visit(pg, 'review', '/review/committee'); const d = await download(pg, /CSV/);
      if (!d) return { status: 'NG', notes: 'CSV ボタンが見つからない' };
      return { status: d.header.includes('昇給率') ? 'OK' : 'NG', notes: `${d.name}｜BOM=${d.bom}｜行数=${d.rows}｜昇給率列=${d.header.includes('昇給率')}` }; } },
  { id: 'RV-067', app: 'review', title: '成長トレンド・スコア分布の表示',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/analytics'), [/成長トレンド|スコア分布/], ERR), 'グラフの内容は目視', true) },
  { id: 'RV-068', app: 'review', title: '評価マトリクス 3×4・キャリブレーション',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/analytics/talent-map'), [/キャリブレーション/, /昇進準備度|準備度/], ERR), '12 セルの配置は目視', true) },
  { id: 'RV-069', app: 'review', title: '評価ステータス（FY26）の表示',
    run: async (pg) => judge(checkTexts(await visit(pg, 'review', '/review/status'), [/期限超過|未提出|進捗/], ERR), '', true) },
  { id: 'RV-082', app: 'review', title: '評価ステータス CSV',
    run: async (pg) => { await visit(pg, 'review', '/review/status'); const d = await download(pg, /CSV/);
      if (!d) return { status: 'NG', notes: 'CSV ボタンが見つからない' };
      return { status: d.rows > 0 && d.bom ? 'OK' : 'NG', notes: `${d.name}｜BOM=${d.bom}｜行数=${d.rows}｜ヘッダー: ${d.header.slice(0, 120)}` }; } },
  { id: 'RV-PERF', app: 'review', title: '主要画面の表示時間（参考）',
    run: async (pg) => { const r = [];
      for (const p of ['/review', '/review/pj', '/review/overall', '/review/goals', '/review/kpi', '/review/checkins', '/review/midterm', '/review/committee', '/review/status']) {
        const t0 = Date.now(); await pg.goto(APPS.review + p, { waitUntil: 'networkidle' }); const t = await bodyText(pg);
        r.push(`${p} ${((Date.now() - t0) / 1000).toFixed(1)}s${ERR.some((e) => e.test(t)) ? ' ★エラー' : ''}`); }
      return { status: r.some((x) => x.includes('★')) ? 'NG' : 'OK', notes: r.join(' / ') }; } },

  // ---------- INTS Accounts（管理者） ----------
  { id: 'AC-003', app: 'accounts', title: 'ユーザー一覧（在籍・件数・バッジ）',
    run: async (pg) => { const t = await visit(pg, 'accounts', '/users'); const m = t.match(/(\d+)\s*名/);
      return judge(checkTexts(t, ['ユーザー', '新規登録', 'SmartHR', '監査ログ', 'プロダクト'], ERR), m ? `件数 ${m[1]} 名（102 名前後が期待）` : '件数表示が見つからない', true); } },
  { id: 'AC-004', app: 'accounts', title: '製品別の絞り込み（?product=sfa / talent）と CSV',
    run: async (pg) => { const a = await visit(pg, 'accounts', '/users?product=sfa'); const b = await visit(pg, 'accounts', '/users?product=talent');
      const chk = checkTexts(a + ' ' + b, ['INTS SFA の利用者', 'INTS Talent の利用者'], ERR);
      await visit(pg, 'accounts', '/users'); const d = await download(pg, /CSV 出力/);
      const note = d ? `${d.name}｜BOM=${d.bom}｜行数=${d.rows}｜ヘッダー: ${d.header.slice(0, 100)}` : 'CSV ボタンが見つからない';
      return judge(chk, note + '｜検索・状態フィルタは UI 確認後に自動化', true); } },
  { id: 'AC-012', app: 'accounts', title: '組織・役職マスタ／監査ログ／プロダクト（閲覧）',
    run: async (pg) => { const t1 = await visit(pg, 'accounts', '/masters'); const t2 = await visit(pg, 'accounts', '/audit'); const t3 = await visit(pg, 'accounts', '/products');
      return judge(checkTexts([t1, t2, t3].join(' '), ['Industry', 'Solution', 'DealReview', 'SFA', 'Talent', 'Review'], ERR), '役職 16 件・各製品の人数は目視', true); } },
];
