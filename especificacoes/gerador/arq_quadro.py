"""Lê o Quadro de Especificações - Geral das plantas de arquitetura (PDF) com coordenadas.

Uso: python3 arq_quadro.py <planta.pdf>  -> imprime JSON {"PISO":{"01":desc}, "PAREDE":{...}, "TETO":{...}}
"""
import sys, json, pymupdf

def quadro(path):
    p = pymupdf.open(path)[0]
    W = p.get_text('words')
    hdr = [w for w in W if w[4] == 'CÓDIGO']
    x0 = min(w[0] for w in hdr) - 15
    top = min(w[1] for w in hdr)
    sq = [w for w in W if w[4] == '□' and w[0] > x0 - 5]
    W = [w for w in W if w[0] > x0 and w[1] > top + 5 and w[0] < x0 + 600]
    codes = sorted([w for w in W if w[0] < x0 + 45 and w[4].isdigit() and len(w[4]) == 2], key=lambda w: w[1])
    # fim do quadro: sexto código depois da segunda volta para 01
    g, l = 0, 0
    for i, c in enumerate(codes):
        if int(c[4]) <= l: g += 1
        if g > 2: codes = codes[:i]; break
        l = int(c[4])
    end = codes[-1][3] + 30
    lines = {}
    for w in W:
        if w[0] >= x0 + 45 and w[1] < end:
            lines.setdefault(round(w[1]), []).append(w)
    ys = sorted(lines)
    paras, cur = [], None
    for y in ys:
        txt = ' '.join(x[4] for x in sorted(lines[y], key=lambda x: x[0]))
        if txt in ('PISO', 'PAREDE', 'TETO', 'DESCRIÇÃO', 'ESPECIFICAÇÕES - GERAL'): continue
        if cur and y - cur['y1'] <= 12:
            cur['t'] += ' ' + txt; cur['y1'] = y
        else:
            cur = {'y0': y, 'y1': y, 't': txt}; paras.append(cur)
    # grupos: troca quando a numeração volta para 01
    out, grp, last = {'PISO': {}, 'PAREDE': {}, 'TETO': {}}, 0, 0
    names = ['PISO', 'PAREDE', 'TETO']
    for c in codes:
        n = int(c[4])
        if n <= last: grp += 1
        if grp > 2: break
        last = n
        cy = (c[1] + c[3]) / 2
        best = min(paras, key=lambda pa: 0 if pa['y0'] - 3 <= cy <= pa['y1'] + 12 else min(abs(cy - pa['y0']), abs(cy - pa['y1'] - 10)))
        out[names[grp]][c[4]] = best['t'].replace('hidrofulgante', 'hidrofugante')
    return out

if __name__ == '__main__':
    print(json.dumps(quadro(sys.argv[1]), ensure_ascii=False, indent=1))
