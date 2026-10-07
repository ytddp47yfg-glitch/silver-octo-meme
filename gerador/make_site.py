"""Monta a página inicial com a lista de obras a partir de site/obras/*.json.card.json.

Gera:
  site/index.html           página completa (para abrir num servidor web)
  site/index_artifact.html  mesma página sem <html>/<head> (para publicar como Artifact no claude.ai)
"""
import glob, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(os.path.dirname(HERE), 'site')
cards = [json.load(open(f)) for f in sorted(glob.glob(os.path.join(SITE, 'obras', '*.json.card.json')))]
for c in cards: c['pdf'] = os.path.exists(os.path.join(SITE, 'pdf', c['id'] + '.pdf'))   # PDF dos painéis (gerador/pdfs.js)
for c in cards: c['xlsx'] = os.path.exists(os.path.join(SITE, 'xlsx', c['id'] + '.xlsx'))   # planilha Excel completa do cenário
# cartões com o mesmo 'obra' são cenários da mesma obra (ex.: padrão × Prevision) e viram um só cartão
obras = {}
for c in sorted(cards, key=lambda c: c.get('cen_ord', 0)):
    k = c.get('obra', c['id'])
    o = obras.setdefault(k, dict(c, id=k, cens=[]))
    o['cens'].append(c)
obras = sorted(obras.values(), key=lambda c: c['nome'])
body = open(os.path.join(HERE, 'multi_tpl.html')).read().replace('/*OBRAS*/[]', json.dumps(obras, ensure_ascii=False))
open(os.path.join(SITE, 'index_artifact.html'), 'w').write(body)
open(os.path.join(SITE, 'index.html'), 'w').write(
    '<!doctype html>\n<html lang="pt-BR"><head><meta charset="utf-8">'
    '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"></head><body>\n'
    + body + '\n</body></html>\n')
print('obras:', [(o['nome'], [c.get('cen') or '-' for c in o['cens']]) for o in obras])
