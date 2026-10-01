"""Lê as etiquetas de acabamento por ambiente nas plantas de arquitetura (PDF com coordenadas).

Etiqueta: ○ piso, △ parede/rodapé (uma ou mais), □ teto, logo abaixo do nome do ambiente e da área.
Uso: python3 arq_ambientes.py <planta.pdf> [...]  -> JSON [{prancha, ambiente, area, piso, parede, teto, xy}]
"""
import sys, json, re, os, pymupdf

def forma(w, segs):
    x0, y0, x1, y1 = w[0] - 9, w[1] - 9, w[2] + 9, w[3] + 9
    near = [r for r in segs if r.x0 >= x0 and r.x1 <= x1 and r.y0 >= y0 and r.y1 <= y1]
    if any(5.5 <= r.width <= 8.5 and 11 <= r.height <= 16 for r in near): return 'parede'
    if any(11 <= r.width <= 16 and r.height < 1 for r in near) and any(11 <= r.height <= 16 and r.width < 1 for r in near): return 'teto'
    small = sum(1 for r in near if r.width <= 2.6 and r.height <= 2.6)
    if small >= 8: return 'piso'
    return None

def ambientes(path):
    p = pymupdf.open(path)[0]
    W = p.get_text('words')
    segs = [d['rect'] for d in p.get_drawings() if d['rect'].width < 20 and d['rect'].height < 20]
    # quadro de especificações fica à direita; ignora essa faixa
    xq = min([w[0] for w in W if w[4] == 'CÓDIGO'] or [1e9]) - 20
    nums = [w for w in W if re.fullmatch(r'\d{2}', w[4]) and w[0] < xq and 5 < (w[3] - w[1]) < 9]
    tags = []
    for w in nums:
        f = forma(w, segs)
        if f: tags.append((w, f))
    # agrupa etiquetas na mesma linha (triângulos ficam ~2pt abaixo de círculos e quadrados)
    tags.sort(key=lambda t: t[0][0])
    grupos = []
    for w, f in tags:
        for g in grupos:
            if abs(g[0][0][1] - w[1]) < 5 and 0 < w[0] - max(x[0][2] for x in g) < 45:
                g.append((w, f)); break
        else:
            grupos.append([(w, f)])
    # rótulos de ambiente: nome em negrito logo acima de 'A=...'
    spans = [sp for bl in p.get_text('dict')['blocks'] for ln in bl.get('lines', []) for sp in ln['spans'] if sp['text'].strip()]
    bold = [sp for sp in spans if 'Bold' in sp['font'] or sp['flags'] & 16]
    labs = []
    for sp in spans:
        t = sp['text'].strip()
        if not re.match(r'A\s*=\s*[\d.,]+', t) or sp['bbox'][0] > xq: continue
        x0, y0, x1, y1 = sp['bbox']
        nome, top = [], y0
        for _ in range(3):
            ln = [b for b in bold if top - 14 < b['bbox'][3] <= top + 4 and b['bbox'][0] < x1 + 60 and b['bbox'][2] > x0 - 60]
            if not ln: break
            nome = sorted(ln, key=lambda b: b['bbox'][0]) + nome
            top = min(b['bbox'][1] for b in ln)
        n = ' '.join(b['text'].strip() for b in nome)
        if n: labs.append({'ambiente': n, 'area': re.sub(r'^A\s*=\s*', '', t).rstrip(' m') + ' m²', 'x': (x0 + x1) / 2, 'y': y1, 'top': top})
    out = []
    usados = set()
    for g in grupos:
        gx = (g[0][0][0] + g[-1][0][2]) / 2; gy = g[0][0][1]
        cand = [l for l in labs if (-5 < gy - l['y'] < 90 or 0 < l['top'] - gy < 70) and abs(gx - l['x']) < 120]
        fs = [f for w, f in g]
        if len(fs) < 2 or 'piso' not in fs: continue
        if not cand:
            out.append({'prancha': re.search(r'PR\d{3}', os.path.basename(path)).group(0), 'ambiente': '?', 'xy': [round(gx), round(gy)],
                        'piso': [w[4] for w, f in g if f == 'piso'], 'parede': [w[4] for w, f in g if f == 'parede'], 'teto': [w[4] for w, f in g if f == 'teto']}); continue
        l = min(cand, key=lambda l: min(abs(gy - l['y']), abs(l['top'] - gy)) + abs(gx - l['x']) * 0.7)
        rec = {'prancha': re.search(r'PR\d{3}', os.path.basename(path)).group(0), 'ambiente': l['ambiente'], 'area': l['area'],
               'piso': [w[4] for w, f in g if f == 'piso'], 'parede': [w[4] for w, f in g if f == 'parede'],
               'teto': [w[4] for w, f in g if f == 'teto'], 'xy': [round(gx), round(gy)]}
        out.append(rec)
    return out

if __name__ == '__main__':
    res = []
    for f in sys.argv[1:]: res += ambientes(f)
    print(json.dumps(res, ensure_ascii=False))
