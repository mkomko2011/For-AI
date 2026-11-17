// INTS UAT 自動確認：共通ヘルパー
// - 認証情報は環境変数 INTS_UAT_EMAIL / INTS_UAT_PASSWORD からのみ読む（ファイル・ログに残さない）
// - 本番環境・共有データのため、データを変更するボタンは押さない（SAFE_DENY でガード）
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const APPS = {
  talent: 'https://ints-talent.vercel.app',
  review: 'https://ints-review.vercel.app',
  accounts: 'https://ints-accounts.vercel.app',
  sfa: 'https://ints-sfa.vercel.app',
};

// これらの文言を含むボタンは誤って押さないよう拒否する
const SAFE_DENY = /保存|提出|登録|追加|削除|承認|却下|差戻|再投入|初期化|取り込|取込|適用|締め|反映|確定|停止|退職|再発行|送信|同期|ロック|解除|作成|コピーして|起案|合意|ログアウト/;

const OUT = path.join(__dirname, 'out');
const AUTH = path.join(__dirname, '.auth');
fs.mkdirSync(OUT, { recursive: true });
fs.mkdirSync(AUTH, { recursive: true });

function creds() {
  const email = process.env.INTS_UAT_EMAIL;
  const password = process.env.INTS_UAT_PASSWORD;
  if (!email || !password) return null;
  return { email, password };
}

async function newBrowser() {
  return chromium.launch();
}

// アプリごとにログインして storageState を保存（.auth/ は git 管理外）
async function loggedInContext(browser, app) {
  const state = path.join(AUTH, `${app}.json`);
  if (fs.existsSync(state)) {
    const ctx = await browser.newContext({ storageState: state, acceptDownloads: true, viewport: { width: 1440, height: 900 } });
    const p = await ctx.newPage();
    await p.goto(APPS[app] + '/', { waitUntil: 'networkidle' });
    const ok = !/\/login/.test(p.url());
    await p.close();
    if (ok) return ctx;
    await ctx.close();
  }
  const c = creds();
  if (!c) throw new Error('NO_CREDENTIALS');
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 900 } });
  const p = await ctx.newPage();
  await p.goto(APPS[app] + '/login', { waitUntil: 'networkidle' });
  await p.locator('input[type=email]').fill(c.email);
  await p.locator('input[type=password]').fill(c.password);
  await p.locator('button[type=submit]').click();
  await p.waitForURL((u) => !/\/login/.test(u.toString()), { timeout: 30000 }).catch(() => {});
  if (/\/login/.test(p.url())) {
    const msg = (await p.innerText('body')).replace(/\s+/g, ' ').slice(0, 200);
    await p.close(); await ctx.close();
    throw new Error('LOGIN_FAILED: ' + msg);
  }
  await ctx.storageState({ path: state });
  await p.close();
  return ctx;
}

async function bodyText(page) {
  await page.waitForLoadState('networkidle').catch(() => {});
  await page.waitForTimeout(800);
  return (await page.innerText('body').catch(() => '')).replace(/\s+/g, ' ');
}

// 期待文言のチェック。missing を返す
function checkTexts(text, expects = [], notExpects = []) {
  const missing = expects.filter((e) => !(e instanceof RegExp ? e.test(text) : text.includes(e)));
  const unexpected = notExpects.filter((e) => (e instanceof RegExp ? e.test(text) : text.includes(e)));
  return { missing, unexpected };
}

// 安全なクリック（タブ・フィルタ・CSV 出力など）。データ変更系の文言は拒否
async function safeClick(page, name, { exact = false } = {}) {
  if (SAFE_DENY.test(name)) throw new Error(`安全ガード: 「${name}」はデータ変更の可能性があるため押しません`);
  const cand = [
    page.getByRole('tab', { name, exact }),
    page.getByRole('button', { name, exact }),
    page.getByRole('link', { name, exact }),
    page.getByText(name, { exact }),
  ];
  for (const l of cand) {
    if (await l.count()) { await l.first().click(); await page.waitForLoadState('networkidle').catch(() => {}); await page.waitForTimeout(500); return true; }
  }
  return false;
}

// CSV 等のダウンロード。ボタン文言は出力系のみ許可
async function download(page, name) {
  if (!/CSV|出力|エクスポート|ダウンロード|Excel/.test(name)) throw new Error('出力系以外のボタンはダウンロード扱いしません: ' + name);
  const btn = page.getByRole('button', { name }).or(page.getByRole('link', { name }));
  if (!(await btn.count())) return null;
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 20000 }), btn.first().click()]);
  const file = path.join(OUT, 'downloads', dl.suggestedFilename());
  fs.mkdirSync(path.dirname(file), { recursive: true });
  await dl.saveAs(file);
  const buf = fs.readFileSync(file);
  const bom = buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
  const text = buf.toString('utf8').replace(/^﻿/, '');
  const lines = text.split(/\r?\n/).filter(Boolean);
  return { file, name: dl.suggestedFilename(), bom, header: lines[0] || '', rows: lines.length - 1, text };
}

async function shot(page, id) {
  const f = path.join(OUT, 'screens', `${id}.png`);
  fs.mkdirSync(path.dirname(f), { recursive: true });
  await page.screenshot({ path: f, fullPage: true }).catch(() => {});
  return f;
}

module.exports = { APPS, SAFE_DENY, OUT, creds, newBrowser, loggedInContext, bodyText, checkTexts, safeClick, download, shot };
