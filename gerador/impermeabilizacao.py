"""Painel IMPERMEABILIZAÇÃO: planta do pavimento pintada por período de execução + legenda com as datas.

Entrada: o cronograma de impermeabilização (aba IMPERMEABILIZAÇÃO: linhas por ambiente/sistema, colunas por semana,
cabeçalho mês na linha 2 e nº da semana na linha 3) e a planta já pintada por gerador/imper_planta.js.
Uso: python3 gerador/impermeabilizacao.py crono.xlsx planta_cor.jpg saida.json   →   VW_IMP=saida.json no viewer.py
"""
import base64, json, sys
from openpyxl import load_workbook

MES = {"JULHO": (7, 2026), "AGOSTO": (8, 2026), "SETEMBRO": (9, 2026), "OUTUBRO": (10, 2026), "NOVEMBRO": (11, 2026),
       "DEZEMBRO": (12, 2026), "JANEIRO / 27": (1, 2027), "FEVEREIRO": (2, 2027)}
ABR = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
# períodos (cor da planta) → linhas do cronograma; cada linha = ambiente (aplicação da manta nas semanas marcadas)
PERIODOS = [
    dict(k=1, cor="#2166ac", sis="Sistema A · manta asfáltica simples (4 mm)", sist=[(10, "Regularização"), (11, "Aplicação da manta"), (12, "Proteção mecânica")], fb=11,
         itens=[(6, "Quadra de tênis"), (7, "Piso de apoio das quadras"), (8, "Circulação externa e varanda do espaço kids e gourmet")]),
    dict(k=2, cor="#4daf4a", sis="Sistema B · dupla manta asfáltica (3 + 4 mm)",
         itens=[(15, "Bar da piscina"), (19, "Deck da piscina descoberta"), (14, "Terraço do salão de festas"),
                (17, "Circulação externa e varanda do espaço kids e gourmet"), (18, "Circulação externa coberta/descoberta"),
                (20, "Circulação interna (trecho 1)"), (21, "Circulação interna (trecho 2)"), (22, "Estar da piscina")]),
    dict(k=3, cor="#ffaa00", sis="Sistema C · piscinas (opção 01, dupla manta + antieflorescência)",
         itens=[(34, "Piscina, piscina infantil e prainha (com paredes)"), (28, "Piscina raia, SPA e escada (com paredes)")]),
    dict(k=4, cor="#d73027", sis="Sistema D · dupla manta antiraiz + drenante (jardins)",
         itens=[(44, "Jardins sobre laje 01 a 06 (playground e ducha), beirais 01 e 02")]),
    dict(k=5, cor="#762a83", sis="Sistema B · dupla manta asfáltica (3 + 4 mm)", itens=[(16, "Quadra de tênis (faixa em dupla manta)")]),
]


def ler(arq):
    ws = load_workbook(arq, data_only=True).worksheets[0]
    sem = {}
    for c in range(6, ws.max_column + 1):
        m, n = ws.cell(2, c).value, ws.cell(3, c).value
        if m in MES and n not in (None, ""): sem[c] = (MES[m], int(n))
    def janela(r):
        cs = [c for c in sem if isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value > 0]
        return (sem[cs[0]], sem[cs[-1]]) if cs else None
    area = lambda r: ws.cell(r, 4).value if isinstance(ws.cell(r, 4).value, (int, float)) else None
    rot = lambda s: f"{s[1]}ª sem. {ABR[s[0][0] - 1]}/{str(s[0][1])[2:]}"
    semanas = [dict(m=f"{ABR[v[0][0] - 1]}/{str(v[0][1])[2:]}", n=v[1]) for c, v in sorted(sem.items())]
    idx = {v: i for i, v in enumerate(v for c, v in sorted(sem.items()))}
    per = []
    for P in PERIODOS:
        its = []
        for r, nome in P["itens"]:
            j = janela(r) or janela(P["fb"])   # ambientes do sistema A: semanas do sistema (linha "Aplicação da manta")
            its.append(dict(n=nome, m2=round(area(r) or 0, 1), de=rot(j[0]), ate=rot(j[1]), i0=idx[j[0]], i1=idx[j[1]]))
        sist = [dict(n=n, de=rot(janela(r)[0]), ate=rot(janela(r)[1]), i0=idx[janela(r)[0]], i1=idx[janela(r)[1]]) for r, n in P.get("sist", [])]
        i0, i1 = min(x["i0"] for x in its), max(x["i1"] for x in its)
        per.append(dict(k=P["k"], cor=P["cor"], sis=P["sis"], itens=its, sist=sist, i0=i0, i1=i1,
                        de=min(its, key=lambda x: x["i0"])["de"],
                        ate=max(its, key=lambda x: x["i1"])["ate"], m2=round(sum(x["m2"] for x in its), 1)))
    return per, semanas


if __name__ == "__main__":
    per, semanas = ler(sys.argv[1])
    img = "data:image/jpeg;base64," + base64.b64encode(open(sys.argv[2], "rb").read()).decode()
    out = dict(tit="Impermeabilização · 3º pavimento (lazer)", per=per, semanas=semanas, img=img,
               fora="Argamassa flexível (áreas molhadas) e tratamentos com resina/cristalizante: fora deste cronograma",
               nota="Datas = semanas de aplicação da manta em cada ambiente, pelo cronograma de impermeabilização (Crono_Imper_Ed_Jardim). "
                    "Cores da planta pelos sistemas do projeto de impermeabilização; no sistema B a planta mostra o período do conjunto "
                    "(set–nov/26) e a legenda traz a semana de cada ambiente. Não considera ralos, rodapés e detalhes.")
    json.dump(out, open(sys.argv[3], "w"), ensure_ascii=False, separators=(",", ":"))
    for p in per: print(p["k"], p["de"], "→", p["ate"], p["m2"], "m²", [(i["n"][:25], i["de"], i["ate"]) for i in p["itens"]])
