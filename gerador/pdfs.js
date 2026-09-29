// Gera site/pdf/<id>.pdf com todos os painéis de cada cenário (Dashboard, Retrato, Curva física, Gantt).
// Uso: sirva a pasta site/ (python3 -m http.server 8766 -d site) e rode: node gerador/pdfs.js [porta]
const { chromium } = require(process.env.PW || '/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const port = process.argv[2] || 8766, site = path.join(__dirname, '..', 'site'), out = path.join(site, 'pdf');
  fs.mkdirSync(out, { recursive: true });
  const ids = fs.readdirSync(path.join(site, 'obras')).filter(f => f.endsWith('.json.card.json')).map(f => f.replace('.json.card.json', ''));
  const b = await chromium.launch();
  for (const id of ids) {
    const p = await b.newPage({ viewport: { width: 1360, height: 900 }, colorScheme: 'light' });
    const errs = []; p.on('pageerror', e => errs.push(e.message));
    await p.goto(`http://localhost:${port}/#${id}`); await p.waitForFunction(() => !document.getElementById('app').hidden && typeof DASH !== 'undefined' && DASH, null, { timeout: 60000 });
    await p.evaluate(() => { try { localStorage.clear() } catch (e) {} });
    const n = await p.evaluate(() => printAll()); await p.waitForTimeout(500);
    const f = path.join(out, id + '.pdf');
    await p.pdf({ path: f, format: 'A4', landscape: true, printBackground: true, scale: 0.72, margin: { top: '8mm', bottom: '10mm', left: '8mm', right: '8mm' },
      displayHeaderFooter: true, headerTemplate: '<span></span>',
      footerTemplate: '<div style="font:8px Arial;width:100%;padding:0 8mm;display:flex;justify-content:space-between;color:#666"><span class="title"></span><span>página <span class="pageNumber"></span> de <span class="totalPages"></span></span></div>' });
    console.log(id, n, 'painéis', Math.round(fs.statSync(f).size / 1024) + ' KB', errs.length ? errs : '');
    await p.close();
  }
  await b.close();
})();
