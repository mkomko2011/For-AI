// 使い方: node run.js [ケースIDの一部で絞り込み]
const fs = require('fs'); const path = require('path');
const { newBrowser, loggedInContext, shot, OUT } = require('./lib');
const cases = require('./cases');
(async () => {
  const filter = process.argv[2];
  const b = await newBrowser(); const ctxs = {}; const results = [];
  let anon = await b.newContext({ viewport: { width: 1440, height: 900 } });
  for (const c of cases.filter((x) => !filter || x.id.includes(filter))) {
    const t0 = Date.now(); let r; let page;
    try {
      let ctx = anon;
      if (c.auth !== false) { ctxs[c.app] = ctxs[c.app] || (await loggedInContext(b, c.app)); ctx = ctxs[c.app]; }
      page = await ctx.newPage();
      r = await c.run(page);
    } catch (e) {
      const m = String(e.message || e);
      r = m.startsWith('NO_CREDENTIALS') ? { status: '未実行', notes: '認証情報（INTS_UAT_EMAIL / INTS_UAT_PASSWORD）が未設定' }
        : m.startsWith('LOGIN_FAILED') ? { status: '未実行', notes: 'ログイン失敗: ' + m.slice(13, 150) }
        : { status: '要目視', notes: '自動確認が途中で止まった（画面構成の想定違いの可能性）: ' + m.slice(0, 150) };
    }
    if (page) { r.screenshot = await shot(page, c.id); await page.close(); }
    results.push({ id: c.id, app: c.app, title: c.title, ...r, sec: ((Date.now() - t0) / 1000).toFixed(1), at: new Date().toISOString() });
    console.log(`${r.status.padEnd(4)} ${c.id.padEnd(10)} ${c.title}  ${r.notes || ''}`);
  }
  fs.writeFileSync(path.join(OUT, 'results.json'), JSON.stringify(results, null, 2));
  await b.close();
})();
