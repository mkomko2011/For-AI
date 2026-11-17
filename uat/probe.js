const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  const hosts = new Set(); const blocked=[];
  p.on('request', r => { try { hosts.add(new URL(r.url()).host); } catch {} });
  p.on('requestfailed', r => blocked.push(r.url()+' '+(r.failure()||{}).errorText));
  await p.goto('https://ints-talent.vercel.app/login', { waitUntil: 'networkidle' });
  console.log('TITLE', await p.title());
  console.log(await p.evaluate(() => [...document.querySelectorAll('input,button,a')].map(e => `${e.tagName} type=${e.type||''} name=${e.name||''} id=${e.id||''} ph=${e.placeholder||''} txt=${(e.innerText||'').trim().slice(0,40)} href=${e.getAttribute('href')||''}`).join('\n')));
  console.log('HOSTS', [...hosts]); console.log('FAILED', blocked);
  await p.screenshot({ path: 'login.png' });
  await b.close();
})();
