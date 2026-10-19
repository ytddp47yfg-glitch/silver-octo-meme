"""Painel MAPA FACHADA: elevação planificada das 4 faces (pavimentos 4º ao 28º) pintada por período de execução,
com legenda das datas e linha do tempo (mesmo layout do MAPA IMPERMEABILIZAÇÃO).

Datas: aba CRONOGRAMA ATIVIDADES do cenário (lote FACHADA: face × etapa F1 = 4º ao 13º / F2 = 14º ao 28º, cremalheira e
limpeza; a face A etapa 1 vem detalhada por balancim). Larguras das faces: m² por pavimento do estudo de balancins
(fachada_tasks.json), só para a proporção do desenho.
Uso: python3 gerador/fachada_mapa.py data2.json fachada_tasks.json saida.json   →   VW_FAC=saida.json no viewer.py
"""
import datetime as dt, json, re, sys
from collections import defaultdict

ABR = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
FACES = [("D", "Lateral esquerda (Baltazar)", "BALTAZAR"), ("C", "Frontal", "FRONTAL"), ("B", "Lateral direita (Jairo)", "JAIRO"),
         ("A", "Posterior", "POSTERIOR")]
CORES = ["#2166ac", "#4393c3", "#1b9e77", "#66a61e", "#e6ab02", "#d95f02", "#e7298a", "#7570b3", "#a6761d", "#666666"]
P1, P2 = (4, 13), (14, 28)   # pavimentos das etapas F1 e F2
d = lambda s: dt.date.fromisoformat(str(s)[:10])
fd = lambda x: x.strftime("%d/%m/%y")


def ler(data2, tasks):
    C = json.load(open(data2))["CRONO"]; n = max(int(k[1:]) for k in C)
    rows = [{c: C.get(f"{c}{r}") for c in "CDEFGK"} for r in range(5, n + 1)]
    fac = [r for r in rows if r["E"] == "FACHADA" or str(r["C"]).upper().startswith(("FACHADA", "LIMPEZA FACHADA"))]
    resumo = {r["C"]: r for r in fac if r["D"] == "-"}
    bal = [r for r in fac if str(r["D"]).startswith("Fixação de granito - balancim")]
    T = json.load(open(tasks))
    larg = defaultdict(float)   # m² por pavimento de cada face (estudo de balancins, etapa 1)
    for t in T:
        if str(t["etapa"]) == "1": larg[t["face"]] += t["m2"] / max(1, t["pav"])
    m2bal = {t["bal"]: t["m2"] for t in T if t["face"] == "A" and str(t["etapa"]) == "1"}
    cr = next(t for t in T if t["bal"] == "20")
    return resumo, bal, larg, m2bal, cr["m2"] / cr["pav"]


def montar(resumo, bal, larg, m2bal, l_cr):
    regs = []   # regiões do desenho / períodos da legenda
    def add(nome, face, et, r, itens=None):
        regs.append(dict(nome=nome, face=face, et=et, ini=d(r["F"]), fim=d(r["G"]), real=r["K"] or 0, itens=itens))
    for f, nm, chave in FACES:
        for et in ("F1", "F2"):
            r = next(v for k, v in resumo.items() if chave in k.upper() and f"{et} ({f})" in k)
            its = None
            if f == "A" and et == "F1" and bal:
                num = lambda b: re.search(r"(\d+)$", b["D"]).group(1)
                its = [dict(n="Balancim " + num(b), ini=d(b["F"]), fim=d(b["G"]), real=b["K"] or 0) for b in bal]
            add(f"Face {f} · {nm} · {'4º ao 13º' if et == 'F1' else '14º ao 28º'}", f, et, r, its)
    add("Cremalheira (faixa da fachada frontal, 4º ao 28º)", "C", "CR", resumo["FACHADA CREMALHEIRA"])
    lp = resumo.get("LIMPEZA FACHADA")
    regs.sort(key=lambda x: (x["ini"], x["fim"]))
    for i, g in enumerate(regs): g["cor"] = CORES[i % len(CORES)]
    return regs, lp


def svg(regs, larg, m2bal, l_cr):
    # elevação planificada: D | C (com a faixa da cremalheira à direita) | B | A; pavimentos 4º (baixo) ao 28º (topo)
    H, Y0, X0, ESC, GAP = 30, 30, 46, 3.4, 14
    reg = {(g["face"], g["et"]): g for g in regs}
    out, x = [], X0
    tot_h = (P2[1] - P1[0] + 1) * H
    yp = lambda p: Y0 + (P2[1] - p) * H     # topo do pavimento p
    def caixa(x, w, p0, p1, g, txt, sub="", fs=17):
        y, h = yp(p1), (p1 - p0 + 1) * H
        out.append(f'<rect x="{x:.1f}" y="{y}" width="{w:.1f}" height="{h}" fill="{g["cor"]}" stroke="#fff" stroke-width="2"><title>{txt}: {fd(g["ini"])} a {fd(g["fim"])}</title></rect>')
        cx, cy = x + w / 2, y + h / 2
        out.append(f'<text x="{cx:.1f}" y="{cy - 4:.1f}" text-anchor="middle" font-size="{fs}" font-weight="700" style="fill:#fff">{txt}</text>')
        if sub: out.append(f'<text x="{cx:.1f}" y="{cy + 13:.1f}" text-anchor="middle" font-size="{fs - 2}" style="fill:#fff">{sub}</text>')
    for f, nm, _ in FACES:
        w = larg[f] * ESC
        wc = l_cr * ESC if f == "C" else 0
        for et, (p0, p1) in (("F1", P1), ("F2", P2)):
            g = reg[(f, et)]
            if f == "A" and et == "F1" and g["itens"]:   # balancins da face A (01 à esquerda)
                its = sorted(g["itens"], key=lambda i: i["n"]); tm = sum(m2bal.get(i["n"][-2:], 1) for i in its); xb = x
                for i in its:
                    wb = w * m2bal.get(i["n"][-2:], 1) / tm
                    caixa(xb, wb, p0, p1, g, "B" + i["n"][-2:], f'{i["fim"].strftime("%d/%m")}', 13); xb += wb
            else:
                caixa(x, w - wc, p0, p1, g, f"{f}{et[1]}", f'{ABR[g["ini"].month - 1]}/{str(g["ini"].year)[2:]} – {ABR[g["fim"].month - 1]}/{str(g["fim"].year)[2:]}')
        if wc: caixa(x + w - wc, wc, P1[0], P2[1], reg[("C", "CR")], "CR", "", 14)
        out.append(f'<text x="{x + w / 2:.1f}" y="{Y0 + tot_h + 22}" text-anchor="middle" font-size="16" font-weight="700" fill="#0b6938">FACE {f}</text>')
        out.append(f'<text x="{x + w / 2:.1f}" y="{Y0 + tot_h + 39}" text-anchor="middle" font-size="13" fill="#555">{nm}</text>')
        x += w + GAP
    W = x - GAP + 10
    for p in range(P1[0], P2[1] + 1):
        out.append(f'<text x="{X0 - 6}" y="{yp(p) + H / 2 + 4}" text-anchor="end" font-size="12" fill="#666">{p}º</text>')
    for p in (P1[1], ):   # divisa das etapas
        out.append(f'<line x1="{X0 - 30}" x2="{W}" y1="{yp(p)}" y2="{yp(p)}" stroke="#333" stroke-dasharray="5 4" stroke-width="1.2"/>')
    out.append(f'<text x="{W}" y="{yp(P1[1]) - 5}" text-anchor="end" font-size="12" fill="#333">etapa 2 (14º ao 28º) ↑ · etapa 1 (4º ao 13º) ↓</text>')
    out.append(f'<rect x="{X0}" y="{Y0 + tot_h}" width="{W - X0 - 10}" height="4" fill="#999"/>')
    # o CSS do painel (.dz svg text) sobrepõe atributos de apresentação: tamanho e cor dos textos vão em style
    def estilo(m):
        t = m.group(0); fs = re.search(r'font-size="([\d.]+)"', t); fl = re.search(r'(?:style="fill:|fill=")(#[0-9a-fA-F]+)"', t)
        t = re.sub(r'\s(font-size|fill)="[^"]*"', "", t); t = re.sub(r'\sstyle="[^"]*"', "", t)
        return t[:-1] + f' style="font-size:{fs.group(1) if fs else 12}px;fill:{fl.group(1) if fl else "#333"}">'
    out = [re.sub(r"<text [^>]*>", estilo, o) for o in out]
    return f'<svg viewBox="0 0 {W:.0f} {Y0 + tot_h + 50}" xmlns="http://www.w3.org/2000/svg" font-family="Arial, Helvetica, sans-serif" role="img" aria-label="Elevação planificada da fachada por período">{"".join(out)}</svg>'


if __name__ == "__main__":
    resumo, bal, larg, m2bal, l_cr = ler(sys.argv[1], sys.argv[2])
    regs, lp = montar(resumo, bal, larg, m2bal, l_cr)
    ini = min(g["ini"] for g in regs); fim = max([g["fim"] for g in regs] + ([d(lp["G"])] if lp else []))
    seg = ini - dt.timedelta(days=ini.weekday())
    semanas, w = [], seg
    while w <= fim:
        semanas.append(dict(m=f"{ABR[w.month - 1]}/{str(w.year)[2:]}", n=(w.day - 1) // 7 + 1, d=w.isoformat())); w += dt.timedelta(days=7)
    wi = lambda x: (x - seg).days // 7
    per = []
    for k, g in enumerate(regs, 1):
        its = [dict(n=i["n"], m2=0, de=fd(i["ini"]), ate=fd(i["fim"]), i0=wi(i["ini"]), i1=wi(i["fim"])) for i in (g["itens"] or [])] \
            or [dict(n=g["nome"].split(" · ", 1)[1] if g["face"] != "C" or g["et"] != "CR" else "Faixa da cremalheira", m2=0, de=fd(g["ini"]), ate=fd(g["fim"]), i0=wi(g["ini"]), i1=wi(g["fim"]))]
        per.append(dict(k=k, cor=g["cor"], sis=g["nome"] + (f" · {g['real']:.0f}% realizado" if g["real"] else ""), itens=its, sist=[],
                        i0=wi(g["ini"]), i1=wi(g["fim"]), de=fd(g["ini"]), ate=fd(g["fim"]), m2=0, replan=""))
    if lp:
        per.append(dict(k=len(per) + 1, cor="#9e9e9e", sis="Limpeza da fachada (todas as faces)", itens=[dict(n="Limpeza final", m2=0, de=fd(d(lp["F"])), ate=fd(d(lp["G"])), i0=wi(d(lp["F"])), i1=wi(d(lp["G"])))],
                        sist=[], i0=wi(d(lp["F"])), i1=wi(d(lp["G"])), de=fd(d(lp["F"])), ate=fd(d(lp["G"])), m2=0, replan=""))
    out = dict(tit="Fachada · pavimentos 4º ao 28º", sub="elevação planificada por período de execução", per=per, semanas=semanas, svg=svg(regs, larg, m2bal, l_cr),
               fora="Embasamento (até o 3º pavimento) fora deste cronograma; limpeza final em todas as faces",
               nota="Datas pelo cronograma de atividades do Prevision (lote FACHADA, cenário Prevision Atual de 19/10/2026): cada face em duas etapas "
                    "(4º ao 13º e 14º ao 28º) e a faixa da cremalheira; a face A etapa 1 é detalhada por balancim (fixação de granito). "
                    "Larguras das faces proporcionais ao m² por pavimento do estudo de balancins; desenho esquemático, sem escala.")
    json.dump(out, open(sys.argv[3], "w"), ensure_ascii=False, separators=(",", ":"))
    for p in per: print(p["k"], p["de"], "→", p["ate"], p["sis"], len(p["itens"]))
