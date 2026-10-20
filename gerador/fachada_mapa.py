"""Painel MAPA FACHADA: elevação planificada das 4 faces (pavimentos 4º ao 28º) pintada por período de execução,
com legenda das datas e linha do tempo (mesmo layout do MAPA IMPERMEABILIZAÇÃO).

Datas: aba CRONOGRAMA ATIVIDADES do cenário (lote FACHADA: face × etapa F1 = 4º ao 13º / F2 = 14º ao 28º, cremalheira e
limpeza; uma região por face e etapa, sem detalhar balancins). Larguras das faces: m² por pavimento do estudo de balancins
(fachada_tasks.json), só para a proporção do desenho.
Uso: python3 gerador/fachada_mapa.py data2.json fachada_tasks.json saida.json   →   VW_FAC=saida.json no viewer.py
"""
import base64, datetime as dt, json, os, re, sys
from collections import defaultdict

ABR = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
FACES = [("D", "Lateral esquerda (Baltazar)", "BALTAZAR"), ("C", "Frontal", "FRONTAL"), ("B", "Lateral direita (Jairo)", "JAIRO"),
         ("A", "Posterior", "POSTERIOR")]
CORES = ["#2166ac", "#4393c3", "#1b9e77", "#66a61e", "#e6ab02", "#d95f02", "#e7298a", "#7570b3", "#a6761d", "#666666"]
P1, P2 = (4, 13), (14, 28)   # pavimentos das etapas F1 e F2
# elevações do projeto (imagens do Estudo de fachada R12): face → (imagem, recorte x, torre x, nível do topo (29º), nível do 4º,
# nível do térreo), em px da imagem; CR_PX = faixa da cremalheira (torre do elevador de obra) na elevação frontal
_IMG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imper", "ed-jardim_elevacoes_{}.png")
ELEV = {"A": ("A_B", (150, 478), (162, 465), 26, 808, 904), "B": ("A_B", (765, 913), (775, 912), 26, 808, 904),
        "C": ("C_D", (262, 646), (287, 627), 37, 818, 915), "D": ("C_D", (985, 1140), (997, 1137), 37, 818, 915)}
CR_PX = (565, 627)   # torre do elevador de obra + 3 vãos de vidro à esquerda
SUAVE = bytes(255 - (255 - v) * 42 // 100 for v in range(256))   # preto → cinza claro (42% do contraste)


def recorte(img, x0, x1, y0, y1):
    import pymupdf
    pm = pymupdf.Pixmap(_IMG.format(img)); doc = pymupdf.open(); pg = doc.new_page(width=pm.width, height=pm.height)
    pg.insert_image(pg.rect, pixmap=pm)
    px = pg.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), clip=pymupdf.Rect(x0, y0, x1, y1), colorspace=pymupdf.csGRAY)
    px = pymupdf.Pixmap(pymupdf.csGRAY, px.width, px.height, bytes(SUAVE[v] for v in px.samples), 0)   # linhas mais claras
    return "data:image/jpeg;base64," + base64.b64encode(px.tobytes("jpeg", jpg_quality=82)).decode()
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
            add(f"Face {f} · {nm} · {'4º ao 13º' if et == 'F1' else '14º ao 28º'}", f, et, r)   # por face (sem balancins)
    add("Cremalheira (faixa da fachada frontal, 4º ao 28º)", "C", "CR", resumo["FACHADA CREMALHEIRA"])
    lp = resumo.get("LIMPEZA FACHADA")
    regs.sort(key=lambda x: (x["ini"], x["fim"]))
    for i, g in enumerate(regs): g["cor"] = CORES[i % len(CORES)]
    return regs, lp


def svg(regs, larg, m2bal, l_cr):
    # elevação planificada: D | C (com a faixa da cremalheira à direita) | B | A; pavimentos 4º (baixo) ao 28º (topo)
    H, Y0, X0, GAP = 30, 70, 46, 22
    reg = {(g["face"], g["et"]): g for g in regs}
    out, x = [], X0
    tot_h = (P2[1] - P1[0] + 1) * H
    yg = Y0 + tot_h + (ELEV["A"][5] - ELEV["A"][4]) * tot_h / (ELEV["A"][4] - ELEV["A"][3])   # térreo no desenho
    yn = lambda dy: Y0 + tot_h + dy * tot_h / (ELEV["A"][4] - ELEV["A"][3])   # nível abaixo do 4º, em px do desenho
    yp = lambda p: Y0 + (P2[1] - p) * H     # topo do pavimento p
    def tinta(x, w, p0, p1, g):
        y, h = yp(p1), (p1 - p0 + 1) * H
        out.append(f'<rect x="{x:.1f}" y="{y}" width="{w:.1f}" height="{h}" fill="{g["cor"]}" fill-opacity=".72" style="mix-blend-mode:multiply">'
                   f'<title>{g["nome"]}: {fd(g["ini"])} a {fd(g["fim"])}</title></rect>')
    def rot(cx, p0, p1, txt, sub="", fs=18):
        cy = yp(p1) + (p1 - p0 + 1) * H / 2
        st = 'stroke="#222" stroke-width="3.5" stroke-linejoin="round" paint-order="stroke"'
        out.append(f'<text x="{cx:.1f}" y="{cy - 4:.1f}" text-anchor="middle" font-size="{fs}" font-weight="700" {st} style="fill:#fff">{txt}</text>')
        if sub: out.append(f'<text x="{cx:.1f}" y="{cy + 14:.1f}" text-anchor="middle" font-size="{fs - 3}" font-weight="700" {st} style="fill:#fff">{sub}</text>')
    for f, nm, _ in FACES:
        # elevação do projeto (estudo R12): recorte da face, escala única, níveis do desenho alinhados à grade dos pavimentos
        img, (a, b), (t0, t1), topo, base4, ter = ELEV[f]
        s = tot_h / (base4 - topo); y_img = yp(P2[1]) - (topo - 8) * s
        w = (b - a) * s; X = lambda px: x + (px - a) * s
        out.append(f'<image x="{x:.1f}" y="{y_img:.1f}" width="{w:.1f}" height="{(ter + 8 - 8) * s:.1f}" preserveAspectRatio="none" href="{recorte(img, a, b, 8, ter + 8)}"/>')
        tw = (t1 - t0) * s; tx = X(t0)
        if f == "C":   # faixa da cremalheira no trecho direito da frontal (torre de elevador de obra no desenho)
            c0, c1 = X(CR_PX[0]), X(CR_PX[1])
            partes = [(tx, c0 - tx), (c1, X(t1) - c1)]
        else:
            partes = [(tx, tw)]
        for et, (p0, p1) in (("F1", P1), ("F2", P2)):
            g = reg[(f, et)]
            for xx, ww in partes:
                if ww > 1: tinta(xx, ww, p0, p1, g)
            rot(tx + (c0 - tx if f == "C" else tw) / 2, p0, p1, f"{f}{et[1]}", f'{ABR[g["ini"].month - 1]}/{str(g["ini"].year)[2:]} – {ABR[g["fim"].month - 1]}/{str(g["fim"].year)[2:]}')
        if f == "C":
            tinta(c0, c1 - c0, P1[0], P2[1], reg[("C", "CR")]); rot((c0 + c1) / 2, P1[0], P2[1], "CR", "", 14)
        out.append(f'<text x="{x + w / 2:.1f}" y="{yg + 22}" text-anchor="middle" font-size="16" font-weight="700" fill="#0b6938">FACE {f}</text>')
        out.append(f'<text x="{x + w / 2:.1f}" y="{yg + 39}" text-anchor="middle" font-size="13" fill="#555">{nm}</text>')
        x += w + GAP
    W = WE = x - GAP + 10
    # planta baixa do pavimento tipo, em desenho separado (abaixo da elevação na tela; na coluna da legenda no A3): contorno
    # da "Planta de distribuição" do Estudo de Fachada (R12), em coordenadas do PDF; cada face é a sequência de vértices do contorno. Faixa externa = etapa 1, interna = etapa 2
    CONT = {"A": [(169, 175), (695, 175)], "B": [(695, 175), (699, 307), (724, 393)],
            "C": [(612, 397.3), (370, 399), (271, 451), (100, 451)], "CR": [(724, 393), (612, 397.3)],
            "D": [(100, 451), (93, 376), (143, 366), (169, 175)]}
    POLI = CONT["A"] + CONT["B"][1:] + CONT["CR"][1:] + CONT["C"][1:] + CONT["D"][1:-1]
    op = []
    PW = 600; k = PW / (724 - 93); PX, PY = 72, 112
    PH = (451 - 175) * k; T = 11
    pt = lambda q: f"{PX + (q[0] - 93) * k:.1f},{PY + (q[1] - 175) * k:.1f}"
    pol = " ".join(pt(q) for q in POLI)
    op.append(f'<text x="{PX + PW / 2}" y="{PY - 88}" text-anchor="middle" font-size="18" font-weight="700" fill="#0b6938">PLANTA BAIXA · PAVIMENTO TIPO</text>')
    op.append(f'<text x="{PX + PW / 2}" y="{PY - 68}" text-anchor="middle" font-size="13" fill="#555">contorno do Estudo de Fachada (R12) · sem escala</text>')
    op.append(f'<defs><clipPath id="pdentro"><polygon points="{pol}"/></clipPath>'
               f'<clipPath id="pfora"><path clip-rule="evenodd" d="M{PX - 60},{PY - 60}h{PW + 120}v{PH + 120}h-{PW + 120}z M{pol.replace(" ", " L")}z"/></clipPath></defs>')
    op.append(f'<polygon points="{pol}" fill="#f4f4f4"/>')
    for f, lin in CONT.items():
        cores = [("pfora", reg[("C", "CR")]["cor"]), ("pdentro", reg[("C", "CR")]["cor"])] if f == "CR" else \
                [("pfora", reg[(f, "F1")]["cor"]), ("pdentro", reg[(f, "F2")]["cor"])]
        for cp, cor in cores:
            op.append(f'<polyline points="{" ".join(pt(q) for q in lin)}" fill="none" stroke="{cor}" stroke-width="{2 * T}" '
                       f'stroke-linejoin="miter" clip-path="url(#{cp})"/>')
    op.append(f'<polygon points="{pol}" fill="none" stroke="#333" stroke-width="1.2"/>')
    cx, cy = PX + ((668 - 93) * k), PY + (395 - 175) * k
    op.append(f'<text x="{cx:.1f}" y="{cy + 30:.1f}" text-anchor="middle" font-size="13" font-weight="700" fill="#555">CR (cremalheira)</text>')
    op.append(f'<text x="{PX + PW / 2}" y="{PY + PH / 2:.1f}" text-anchor="middle" font-size="13" fill="#888">torre</text>')
    cy = PY + PH / 2   # faces A e C em cima e embaixo; D e B com o texto girado nas laterais, para a planta ocupar a largura
    lab = [("A", PX + (432 - 93) * k, PY - 38, "", "Posterior"), ("C", PX + (330 - 93) * k, PY + PH + 30, "", "Frontal")]
    for f, lx, ly, _, nm in lab:
        op.append(f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" font-size="20" font-weight="700" fill="#0b6938">FACE {f}</text>')
        op.append(f'<text x="{lx:.1f}" y="{ly + 19:.1f}" text-anchor="middle" font-size="15" fill="#555">{nm}</text>')
    for f, nm, xt, xn, r in (("D", "Lat. esquerda (Baltazar)", PX - 46, PX - 24, -90), ("B", "Lat. direita (Jairo)", PX + PW + 48, PX + PW + 26, 90)):
        op.append(f'<text x="{xt:.1f}" y="{cy:.1f}" transform="rotate({r} {xt:.1f} {cy:.1f})" text-anchor="middle" font-size="20" font-weight="700" fill="#0b6938">FACE {f}</text>')
        op.append(f'<text x="{xn:.1f}" y="{cy:.1f}" transform="rotate({r} {xn:.1f} {cy:.1f})" text-anchor="middle" font-size="15" fill="#555">{nm}</text>')
    yk = PY + PH + 66
    op.append(f'<rect x="{PX}" y="{yk}" width="26" height="10" fill="#bbb"/><rect x="{PX}" y="{yk + 14}" width="26" height="10" fill="#777"/>')
    op.append(f'<text x="{PX + 34}" y="{yk + 9}" font-size="14" fill="#555">faixa externa = etapa 1 (4º ao 13º)</text>')
    op.append(f'<text x="{PX + 34}" y="{yk + 23}" font-size="14" fill="#555">faixa interna = etapa 2 (14º ao 28º)</text>')
    op.append(f'<text x="{PX}" y="{yk + 47}" font-size="14" fill="#555">Elevação: faces na ordem D → C → B → A, contornando a torre a partir da lateral esquerda.</text>')
    WP, HP = PX + PW + 72, yk + 60
    W, HT = WE + 130, yg + 50
    for p in range(P1[0], P2[1] + 1):
        out.append(f'<text x="{X0 - 10}" y="{yp(p) + H / 2 + 4}" text-anchor="end" font-size="12" fill="#666">{p}º</text>')
    for p in (P1[1], ):   # divisa das etapas
        out.append(f'<line x1="{X0 - 30}" x2="{WE}" y1="{yp(p)}" y2="{yp(p)}" stroke="#333" stroke-dasharray="5 4" stroke-width="1.2"/>')
    out.append(f'<text x="{WE + 6}" y="{yp(P1[1]) - 6}" font-size="12" font-weight="700" fill="#333">↑ etapa 2 (14º–28º)</text>')
    out.append(f'<text x="{WE + 6}" y="{yp(P1[1]) + 15}" font-size="12" font-weight="700" fill="#333">↓ etapa 1 (4º–13º)</text>')
    for t, yy in (("3º", yn(17)), ("2º", yn(52)), ("térreo", yn(83))):
        out.append(f'<text x="{X0 - 10}" y="{yy + 4}" text-anchor="end" font-size="12" fill="#999">{t}</text>')
    out.append(f'<rect x="{X0 - 30}" y="{yg}" width="{WE - X0 + 30}" height="4" fill="#666"/>')
    nt = [("28º laje técnica", 28), ("26º–27º duplex", 26.5), ("4º–25º tipo", 20)]
    for t, pp in nt:
        out.append(f'<text x="{WE + 6}" y="{yp(pp) + H / 2 + 4:.1f}" font-size="12" fill="#777">{t}</text>')
    for t, yy in (("3º pilotis / lazer", yn(17)), ("2º garagem", yn(52)), ("térreo garagens", yn(83))):
        out.append(f'<text x="{WE + 6}" y="{yy + 4}" font-size="12" fill="#777">{t}</text>')
    # o CSS do painel (.dz svg text) sobrepõe atributos de apresentação: tamanho e cor dos textos vão em style
    def estilo(m):
        t = m.group(0); fs = re.search(r'font-size="([\d.]+)"', t); fl = re.search(r'(?:style="fill:|fill=")(#[0-9a-fA-F]+)"', t)
        t = re.sub(r'\s(font-size|fill)="[^"]*"', "", t); t = re.sub(r'\sstyle="[^"]*"', "", t)
        return t[:-1] + f' style="font-size:{fs.group(1) if fs else 12}px;fill:{fl.group(1) if fl else "#333"}">'
    out = [re.sub(r"<text [^>]*>", estilo, o) for o in out]; op = [re.sub(r"<text [^>]*>", estilo, o) for o in op]
    cab = '<svg viewBox="0 0 {:.0f} {:.0f}" xmlns="http://www.w3.org/2000/svg" font-family="Arial, Helvetica, sans-serif" role="img" aria-label="{}">'
    return (cab.format(W, HT, "Elevações da fachada pintadas por período") + "".join(out) + "</svg>",
            cab.format(WP, HP, "Planta baixa do pavimento tipo com as faces") + "".join(op) + "</svg>")


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
    el, pl = svg(regs, larg, m2bal, l_cr)
    out = dict(tit="Fachada · pavimentos 4º ao 28º", sub="elevação planificada por período de execução", per=per, semanas=semanas, svg=el, planta=pl,
               fora="Embasamento (até o 3º pavimento) fora deste cronograma; limpeza final em todas as faces",
               nota="Datas pelo cronograma de atividades do Prevision (lote FACHADA, cenário Prevision Atual de 19/10/2026): cada face em duas etapas "
                    "(4º ao 13º e 14º ao 28º) e a faixa da cremalheira, uma cor por face e etapa. "
                    "Desenho: elevações do projeto (imagens do Estudo de fachada R12), na ordem D, C, B, A. Planta baixa: contorno do pavimento tipo da Planta de distribuição do Estudo de Fachada (R12); "
                    "a cremalheira fica no trecho direito da frontal (balancim 20 do estudo).")
    json.dump(out, open(sys.argv[3], "w"), ensure_ascii=False, separators=(",", ":"))
    for p in per: print(p["k"], p["de"], "→", p["ate"], p["sis"], len(p["itens"]))
