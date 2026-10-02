"""Leva as imagens da web para a página da grade (artifact) e re-embute a xlsx.

Uso: python3 pagina_imagens_web.py <pagina.html> <saida.html>
Lê especificacoes/imagens/imagens_produtos_web.json (campo asset_id = id do arquivo
enviado ao artifact; url /_blob/<id>) e:
- cria const WEB com produto, marca, imagem, página de origem e as células;
- botão "Imagem web" na aba Grade (por célula) e na aba Ambientes (por item);
- seção "Imagens da web" na aba Imagens;
- re-embute a xlsx em <script id="xlsxdata">.
Idempotente: substitui o bloco marcado /*WEB*/ se já existir.
"""
import sys, json, re, base64, unicodedata

SRC, OUT = sys.argv[1], sys.argv[2]
XLSX = 'especificacoes/CDJ-GRE-ARQ-AC-R00.xlsx'
WEBJ = 'especificacoes/imagens/imagens_produtos_web.json'
ABA = {'TÉRREO': 0, '2° PAVTO': 1, 'PILOTIS': 2, 'APTO TIPO': 3}
L0 = 13   # primeira linha de ambiente nas abas de pavimento

sys.path.insert(0, 'especificacoes/gerador')
USUARIA = set()
for ln in open('especificacoes/gerador/imagens_links.py', encoding='utf-8'):
    m = re.match(r"\s*\('(\w[\w ]*)', '([A-Z]+\d+)', 'CDJ-IMG", ln)
    if m:
        USUARIA.add((m.group(1).strip(), m.group(2)))

def col_idx(ref):
    letras = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n, int(ref[len(letras):])

web = []
for p in json.load(open(WEBJ, encoding='utf-8')):
    if not p.get('asset_id'):
        continue
    cells = []
    for c in dict.fromkeys(p['celulas']):
        aba, ref = c.split('!')
        aba = aba.strip()
        if (aba, ref) in USUARIA:
            continue
        col, row = col_idx(ref)
        cells.append([ABA[aba], row - L0, col - 1])
    web.append({'n': p['n'], 't': p['produto'], 'm': p['marca'], 'src': '/_blob/' + p['asset_id'],
                'pg': p.get('pagina_url') or '', 'ok': p.get('confianca') == 'alta', 'cells': cells})

s = open(SRC, encoding='utf-8').read()
s = re.sub(r'\n/\*WEB\*/.*?/\*FIMWEB\*/', '', s, flags=re.S)

js = '''
/*WEB*/
var WEB=%s;
var WEBC={};WEB.forEach((w,k)=>w.cells.forEach(([a,r,c])=>{(WEBC[a+'.'+r+'.'+c]=WEBC[a+'.'+r+'.'+c]||[]).push(k)}));
var nrm=s=>String(s).normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/[^A-Za-z0-9]/g,'').toUpperCase();
var WEBA={};WEB.forEach((w,k)=>w.cells.forEach(([a,r,c])=>{const row=X.sheets[a]&&X.sheets[a].rows[r];if(!row)return;const key=row[0]+'|'+nrm(X.cols[c]);(WEBA[key]=WEBA[key]||new Set()).add(k)}));
function wbtn(ks){return (ks||[]).map(k=>' <button type="button" class="ximg xweb" data-w="'+k+'" aria-label="Ver imagem da web: '+esc(WEB[k].t)+'">Imagem web'+(ks.length>1?' '+(ks.indexOf(k)+1):'')+'</button>').join('');}
function wcard(w){return '<div class="card"><div class="pics"><img alt="" src="'+w.src+'"></div><div class="cap"><b>WEB-'+String(w.n).padStart(2,'0')+' · '+esc(w.m)+'</b>'+esc(w.t)+'</div><div class="up"><span class="st '+(w.ok?'ok':'low')+'">'+(w.ok?'Produto confere':'Referência · conferir')+'</span>'+(w.pg?'<a class="upmsg" href="'+esc(w.pg)+'" target="_blank" rel="noopener">Página de origem</a>':'')+'</div></div>';}
function openWeb(k){lbIx=null;const w=WEB[k];document.getElementById('lbT').textContent='Imagem da web · referência';document.getElementById('lbG').innerHTML=wcard(w);lb.hidden=false;document.getElementById('lbX').focus();}
document.addEventListener('click',e=>{const b=e.target.closest('.xweb');if(b){e.preventDefault();e.stopImmediatePropagation();lastBtn=b;openWeb(+b.dataset.w)}},true);
document.getElementById('imgsweb').innerHTML=WEB.map(wcard).join('');
render();xrender();
/*FIMWEB*/''' % json.dumps(web, ensure_ascii=False)

# botões na grade (por célula) e em Ambientes (por item)
s = s.replace("+'</td>').join('')+'</tr>';});", "+(typeof WEBC!=='undefined'?wbtn(WEBC[XS+'.'+i+'.'+(c+1)]):'')+'</td>').join('')+'</tr>';});", 1)
s = s.replace("i.v.map(v=>'<p>'+fmt(v)+refBtn(a.n,v)+'</p>').join('')+'</div>')",
              "i.v.map(v=>'<p>'+fmt(v)+refBtn(a.n,v)+'</p>').join('')+(typeof WEBA!=='undefined'&&WEBA[a.n+'|'+nrm(i.k)]?'<p>'+wbtn([...WEBA[a.n+'|'+nrm(i.k)]])+'</p>':'')+'</div>')", 1)
assert 'wbtn(WEBC' in s and 'WEBA[a.n' in s
# seção na aba Imagens
if 'id="imgsweb"' not in s:
    s = s.replace('<div class="imgs" id="imgs"></div>\n</section>',
                  '<div class="imgs" id="imgs"></div>\n  <h4>Imagens da web (referência dos produtos especificados)</h4>\n'
                  '  <p class="muted">Fotos de fabricantes e lojas para os produtos com marca e código na grade. São referência: confira o item na amostra. As marcadas "conferir" são de modelo aproximado ou genérico.</p>\n'
                  '  <div class="imgs" id="imgsweb"></div>\n</section>', 1)
assert 'id="imgsweb"' in s
s = s.replace('O botão "Ver imagem" aparece ao lado do item que tem foto na tabela de acabamentos.',
              'O botão "Ver imagem" aparece ao lado do item que tem foto na tabela de acabamentos, e "Imagem web" nos produtos com foto de referência da web.')
# o bloco WEB entra depois de X (usa X, esc, lb) e antes do fim do script principal
i = s.index("const tabs=document.querySelectorAll('nav.tabs button');")
s = s[:i] + js.lstrip('\n') + '\n' + s[i:]
# xlsx embutida
b64 = base64.b64encode(open(XLSX, 'rb').read()).decode()
s = re.sub(r'(<script id="xlsxdata" type="application/octet-stream">)[^<]*(</script>)', lambda m: m.group(1) + b64 + m.group(2), s, count=1)
open(OUT, 'w', encoding='utf-8').write(s)
print(len(web), 'produtos,', sum(len(w['cells']) for w in web), 'células,', len(s) // 1024, 'KB')
