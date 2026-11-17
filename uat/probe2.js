const { chromium } = require('playwright');
const apps = {
  talent: ['https://ints-talent.vercel.app', ['/login','/dashboard','/admin/users','/assignments','/utilization']],
  review: ['https://ints-review.vercel.app', ['/login','/review','/admin/fy','/review/committee']],
  accounts: ['https://ints-accounts.vercel.app', ['/login','/users']],
  sfa: ['https://ints-sfa.vercel.app', ['/','/login']],
};
(async () => {
  const b = await chromium.launch();
  for (const [k,[base,paths]] of Object.entries(apps)) {
    for (const path of paths) {
      const p = await b.newPage();
      const r = await p.goto(base+path, { waitUntil: 'networkidle' }).catch(e=>null);
      const txt = (await p.innerText('body').catch(()=>'')).replace(/\s+/g,' ').slice(0,300);
      console.log(`[${k}] ${path} -> ${r&&r.status()} final=${p.url()}\n   ${txt}`);
      await p.close();
    }
  }
  await b.close();
})();
