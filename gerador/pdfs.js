// Gera site/pdf/<id>--<painel>.pdf: um PDF por painel de cada cenário (Dashboard, Retrato, Curva física, Gantt, Esquemático).
// O botão ⬇ PDF da página baixa o PDF do painel aberto.
// Uso: sirva a pasta site/ (python3 -m http.server 8766 -d site) e rode: node gerador/pdfs.js [porta]
const { chromium } = require(process.env.PW || '/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const port = process.argv[2] || 8766, site = path.join(__dirname, '..', 'site'), out = path.join(site, 'pdf');
  fs.mkdirSync(out, { recursive: true });
  const ids = fs.readdirSync(path.join(site, 'obras')).filter(f => f.endsWith('.json.card.json')).map(f => f.replace('.json.card.json', ''));
  // PDFs antigos (um por cenário, ou de painéis que deixaram de existir) saem
  for (const f of fs.readdirSync(out)) if (f.endsWith('.pdf')) fs.unlinkSync(path.join(out, f));
  const b = await chromium.launch();
  const abre = async id => {
    const p = await b.newPage({ viewport: { width: 1360, height: 900 }, colorScheme: 'light' });
    const errs = []; p.on('pageerror', e => errs.push(e.message));
    await p.goto(`http://localhost:${port}/#${id}`);
    await p.waitForFunction(() => !document.getElementById('app').hidden && typeof DASH !== 'undefined' && DASH, null, { timeout: 60000 });
    await p.evaluate(() => { try { localStorage.clear() } catch (e) {} });
    return { p, errs };
  };
  for (const id of ids) {
    let { p } = await abre(id);
    const nomes = await p.evaluate(() => PAINEIS.map(n => [n, pdfSlug(n)]));
    await p.close();
    const feitos = [];
    for (const [nome, slug] of nomes) {
      const { p, errs } = await abre(id);
      await p.evaluate(n => printAll(n), nome); await p.waitForTimeout(400);
      const f = path.join(out, `${id}--${slug}.pdf`);
      await p.pdf({ path: f, format: 'A4', landscape: true, printBackground: true, scale: 0.72, margin: { top: '8mm', bottom: '10mm', left: '8mm', right: '8mm' },
        displayHeaderFooter: true, headerTemplate: '<span></span>',
        footerTemplate: '<div style="font:8px Arial;width:100%;padding:0 8mm;display:flex;justify-content:space-between;color:#666"><span class="title"></span><span>página <span class="pageNumber"></span> de <span class="totalPages"></span></span></div>' });
      feitos.push(`${slug} ${Math.round(fs.statSync(f).size / 1024)} KB${errs.length ? ' ERRO ' + errs.join('; ') : ''}`);
      await p.close();
    }
    console.log(id, '·', feitos.join(' · '));
  }
  await b.close();
})();
