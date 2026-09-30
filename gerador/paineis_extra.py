"""Abas CURVA FÍSICA e GANTT da planilha, no mesmo visual dos painéis online.

curva_fisica(wb, F, obra): F = JSON de gerador/fisico.py (curvas de avanço físico por cenário).
gantt(wb, G, obra):        G = JSON de gerador/gantt.py (árvore PL → pacote → pavimento do cronograma do Prevision).
Os dois JSON são uma fotografia dos dados de planejamento na data em que a planilha foi gerada (anotado em cada aba).
"""
import datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.formula import ArrayFormula

F_ = "Calibri"
NAVY, INK, INK2, MUT, LINE, BG = "1C3B2A", "3C3A33", "8C877A", "8C877A", "E5DFD2", "F4F1EA"
TEAL, TEALH, HL = "F1F5F0", "C5D3C8", "D9E6DC"
fill = lambda h: PatternFill("solid", fgColor=h)
sd = Side(style="thin", color=LINE)
MES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def _d(s): return dt.date.fromisoformat(s if len(s) > 7 else s + "-01")


def _mlab(s): y, m = s[:7].split("-"); return f"{MES[int(m) - 1]}/{y[2:]}"


def _fundo(ws, max_col, max_row):
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for c in row:
            if type(c).__name__ != "MergedCell" and (c.fill is None or c.fill.fill_type is None): c.fill = fill(BG)


def _card(ws, c0, c1, r0, r1):
    for rr in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            cell = ws.cell(rr, cc); cell.fill = fill("FCFBF8")
            cell.border = Border(top=sd if rr == r0 else None, bottom=sd if rr == r1 else None,
                                 left=sd if cc == c0 else None, right=sd if cc == c1 else None)


def _titulo(ws, t, sub):
    ws["B2"] = t; ws["B2"].font = Font(name=F_, bold=True, size=20, color=NAVY)
    ws["B3"] = sub; ws["B3"].font = Font(name=F_, size=10, color=INK2)
    ws.row_dimensions[2].height = 34; ws.row_dimensions[3].height = 20


def _h2(ws, ref, txt):
    ws[ref] = txt; ws[ref].font = Font(name=F_, bold=True, size=10, color=NAVY)


# ============================================================ CURVA FÍSICA
def curva_fisica(wb, F, obra):
    ws = wb.create_sheet("CURVA FÍSICA", wb.sheetnames.index("RETRATO 2027") + 1)
    ws.sheet_view.showGridLines = False; ws.sheet_properties.tabColor = "1F7A52"
    cens = F["cens"]; nc = len(cens)
    cor = ["7FA88F", "1F7A52", "1C3B2A"]
    ic = F["meses"].index(F["corte"])
    fim = min(len(F["meses"]) - 1, max([next((i for i, x in enumerate(v) if x >= 0.9995), len(F["meses"]) - 1)
                                        for v in [F["prev"]] + [c["v"] for c in cens]]) + 2)
    widths = {"A": 2, "B": 34, "C": 14, "D": 14}
    for k in range(nc + 2): widths[L(3 + k)] = 16
    for c, w in widths.items(): ws.column_dimensions[c].width = w
    _titulo(ws, f"{obra}  ·  CURVA FÍSICA POR CENÁRIO",
            f"Avanço físico acumulado dos itens diretos ({F['npl']} PLs), ponderado pela projeção de cada PL  ·  Corte: {_mlab(F['corte'])}"
            f"  ·  Realizado no corte: {F['real'][ic] * 100:.2f}%".replace(".", ","))
    # ---- dados mensais (base do gráfico e das tabelas), embaixo
    D0 = 70
    _h2(ws, f"B{D0 - 1}", "DADOS MENSAIS (% ACUMULADO)")
    hdr = ["Mês", "Realizado", "Previsto (Prevision)"] + [c["nome"] for c in cens]
    for j, h in enumerate(hdr):
        c = ws.cell(D0, 2 + j, h); c.font = Font(name=F_, bold=True, size=9, color=NAVY); c.fill = fill(TEAL)
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", wrap_text=True)
    for i in range(fim + 1):
        r = D0 + 1 + i
        ws.cell(r, 2, _d(F["meses"][i])).number_format = "dd/mm/yyyy"
        vals = [F["real"][i], F["prev"][i]] + [c["v"][i] for c in cens]
        for j, v in enumerate(vals):
            c = ws.cell(r, 3 + j, v); c.number_format = "0.00%"
        for j in range(len(hdr)):
            ws.cell(r, 2 + j).font = Font(name=F_, size=9, color=INK); ws.cell(r, 2 + j).border = Border(bottom=sd)
    R1 = D0 + 1 + fim
    # ---- gráfico (estilo Curva S do Prevision): realizado em colunas + linhas
    _card(ws, 2, 3 + nc + 1, 5, 27)
    ws["B5"] = "Curva S"; ws["B5"].font = Font(name=F_, size=15, color=INK)
    ws["B6"] = "Visão geral  ›  realizado em colunas; previsto do Prevision e cenários em linhas"; ws["B6"].font = Font(name=F_, size=9, color=MUT)
    bar = BarChart(); bar.type = "col"; bar.gapWidth = 10
    bar.add_data(Reference(ws, min_col=3, min_row=D0, max_row=R1), titles_from_data=True)
    bar.series[0].graphicalProperties.solidFill = "C5D3C8"; bar.series[0].graphicalProperties.line.noFill = True
    bar.set_categories(Reference(ws, min_col=2, min_row=D0 + 1, max_row=R1))
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=4, max_col=4 + nc, min_row=D0, max_row=R1), titles_from_data=True)
    for k, s in enumerate(ln.series):
        s.smooth = True; s.marker.symbol = "none"
        s.graphicalProperties.line.solidFill = "C0A062" if k == 0 else cor[(k - 1) % 3]
        s.graphicalProperties.line.width = 38100 if k == nc else 25400
    bar += ln
    bar.y_axis.scaling.min = 0; bar.y_axis.scaling.max = 1; bar.y_axis.numFmt = "0.00%"; bar.y_axis.majorUnit = 0.25
    bar.x_axis.number_format = "dd/mm/yyyy"; bar.x_axis.tickLblSkip = 4; bar.x_axis.delete = False; bar.y_axis.delete = False
    bar.legend.position = "t"; bar.height = 10.5; bar.width = 31; bar.title = None
    ws.add_chart(bar, "B7")
    # ---- avanço no fim de cada ano (fórmulas sobre os dados mensais)
    A0 = 30
    _h2(ws, f"B{A0}", "AVANÇO FÍSICO ACUMULADO NO FIM DE CADA ANO")
    for j, h in enumerate(["Data", "Realizado", "Previsto (Prevision)"] + [c["nome"] for c in cens]):
        c = ws.cell(A0 + 1, 2 + j, h); c.font = Font(name=F_, bold=True, size=9, color="FFFFFF"); c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", vertical="center", wrap_text=True)
    ws.row_dimensions[A0 + 1].height = 30
    anos = sorted({int(m[:4]) for m in F["meses"][:fim + 1]})
    ano_rt = 2027
    rng = lambda col: f"${L(col)}${D0 + 1}:${L(col)}${R1}"
    for k, a in enumerate(anos):
        r = A0 + 2 + k
        ws.cell(r, 2, dt.date(a, 12, 31)).number_format = "dd/mm/yyyy"
        for j in range(nc + 2):
            col = 3 + j
            f = f"=IFERROR(INDEX({rng(col)},MATCH(DATE(YEAR($B{r}),12,1),{rng(2)},0)),INDEX({rng(col)},ROWS({rng(col)})))"
            if j == 0: f = f'=IF(DATE(YEAR($B{r}),12,1)>DATE({F["corte"][:4]},{int(F["corte"][5:7])},1),"—",{f[1:]})'
            c = ws.cell(r, col, f); c.number_format = "0.0%"; c.alignment = Alignment(horizontal="right")
        for j in range(nc + 3):
            c = ws.cell(r, 2 + j); c.fill = fill(HL if a == ano_rt else ("FFFFFF" if k % 2 else TEAL)); c.border = Border(bottom=sd)
            c.font = Font(name=F_, size=10, bold=(a == ano_rt), color=NAVY if a == ano_rt else INK)
    r = A0 + 2 + len(anos)
    ws.cell(r, 2, "Conclusão (100%)").font = Font(name=F_, bold=True, size=10, color=INK)
    for j in range(1, nc + 2):
        col = 3 + j; ref = f"{L(col)}{r}"
        ws[ref] = ArrayFormula(ref, f"=IFERROR(INDEX({rng(2)},MATCH(TRUE,{rng(col)}>=0.9995,0)),\"—\")")
        ws[ref].number_format = "mmm/yy"; ws[ref].alignment = Alignment(horizontal="right"); ws[ref].font = Font(name=F_, bold=True, size=10, color=INK)
    for j in range(nc + 3): ws.cell(r, 2 + j).fill = fill(TEAL); ws.cell(r, 2 + j).border = Border(bottom=sd)
    # ---- por PL em 31/12 do ano do retrato
    P0 = r + 3
    k_ano = F["anos"].index(ano_rt) if ano_rt in F["anos"] else 0
    _h2(ws, f"B{P0}", f"POR PL EM 31/12/{ano_rt}")
    for j, h in enumerate(["PL", "No corte", "Prevision"] + [c["nome"] for c in cens]):
        c = ws.cell(P0 + 1, 2 + j, h); c.font = Font(name=F_, bold=True, size=9, color="FFFFFF"); c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", vertical="center", wrap_text=True)
    ws.row_dimensions[P0 + 1].height = 30
    for i, p in enumerate(cens[0]["pls"]):
        rr = P0 + 2 + i
        vals = [p["n"], p["fc"], p["pdez"][k_ano]] + [c["pls"][i]["dez"][k_ano] for c in cens]
        for j, v in enumerate(vals):
            c = ws.cell(rr, 2 + j, v); c.number_format = "0.0%"; c.font = Font(name=F_, size=10, color=INK)
            c.fill = fill(TEAL if i % 2 == 0 else "FFFFFF"); c.border = Border(bottom=sd)
            c.alignment = Alignment(horizontal="left" if j == 0 else "right")
    rN = P0 + 2 + len(cens[0]["pls"])
    ws.cell(rN + 1, 2, ("Como a curva é montada. Realizado e Previsto (Prevision): % acumulado do cronograma físico-financeiro do Prevision "
                        "(aba CP RESUMO) de cada PL. Cenários: até o corte, o realizado; depois, o físico de cada PL avança na mesma proporção do "
                        "desembolso que falta no cenário (curva financeira da aba da PL). PL sem desembolso futuro segue o previsto do Prevision. "
                        "A obra é a média das PLs diretas ponderada pela projeção. Valores gerados a partir das planilhas dos três cenários "
                        f"(fotografia de {dt.date.today():%d/%m/%Y}; os dados mensais estão na linha {D0}).")).font = Font(name=F_, italic=True, size=8, color=MUT)
    _fundo(ws, 3 + nc + 2, R1 + 2)
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


# ============================================================ GANTT
def gantt(wb, G, obra):
    pos = wb.sheetnames.index("CURVA FÍSICA") + 1 if "CURVA FÍSICA" in wb.sheetnames else wb.sheetnames.index("RETRATO 2027") + 1
    ws = wb.create_sheet("GANTT", pos)
    ws.sheet_view.showGridLines = False; ws.sheet_properties.tabColor = "1F7A52"
    ws.sheet_properties.outlinePr.summaryBelow = False
    d0 = _d(G["ini"][:7]); df = _d(G["fim"][:7])
    M = (df.year - d0.year) * 12 + df.month - d0.month + 2
    C0 = 7   # 1ª coluna de mês (G)
    for c, w in {"A": 2, "B": 50, "C": 10, "D": 10, "E": 8, "F": 3}.items(): ws.column_dimensions[c].width = w
    ws.column_dimensions["F"].hidden = True
    for k in range(M): ws.column_dimensions[L(C0 + k)].width = 2.6
    ref = _d(G["ref"]); ret = _d(G["ret"])
    _titulo(ws, f"{obra}  ·  GANTT DO CRONOGRAMA",
            f"{G['n']:,} atividades do Prevision (situação em {ref:%d/%m/%y})  ·  PL → pacote de trabalho → pavimento  ·  "
            "use os botões + / − à esquerda para abrir e fechar os níveis".replace(",", "."))
    # legenda
    leg = [("C5D3C8", "Planejado (início → término)"), ("446756", "Realizado (% do Prevision)"), ("1F7A52", "Concluído"),
           ("F1E6CC", "Não descrito de forma expressa na curva do Prevision (nota no comentário da célula: passe o mouse)")]
    col = 3
    for cor, txt in leg:
        ws.cell(5, col).fill = fill(cor); ws.cell(5, col + 1, txt).font = Font(name=F_, size=9, color=INK2)
        col += 1 if len(txt) < 12 else 2
    ws["C5"].fill = fill("C5D3C8"); ws["D5"] = leg[0][1]
    # (a legenda cabe melhor em linhas separadas)
    for c in range(3, 20): ws.cell(5, c).value = None; ws.cell(5, c).fill = PatternFill()
    for i, (cor, txt) in enumerate(leg + [("B3392B", f"Linha tracejada vermelha: situação em {ref:%d/%m/%y}"), (NAVY, f"Linha tracejada verde-escura: {ret:%d/%m/%Y}")]):
        r = 4 + i if i < 3 else 4 + i - 3
        cc = 3 if i < 3 else C0 + 2
        cl = ws.cell(r, cc); cl.fill = fill(cor)
        if cor == "F1E6CC": cl.border = Border(left=Side(style="dashed", color="C0A062"), right=Side(style="dashed", color="C0A062"),
                                              top=Side(style="dashed", color="C0A062"), bottom=Side(style="dashed", color="C0A062"))
        t = ws.cell(r, cc + 1, txt); t.font = Font(name=F_, size=9, color="8C6D2F" if cor == "F1E6CC" else INK2)
    ws["B4"] = "Legenda"; ws["B4"].font = Font(name=F_, bold=True, size=9, color=INK2)
    # cabeçalho: anos (linha 8) e meses (linha 9, datas com formato de 1 letra)
    HY, HM, R0 = 8, 9, 10
    heads = {2: "PL / pacote / pavimento", 3: "Início", 4: "Término", 5: "Real."}
    for c in range(2, C0 + M):
        for r in (HY, HM): ws.cell(r, c).fill = fill(TEAL); ws.cell(r, c).font = Font(name=F_, bold=True, size=9, color=NAVY)
    for c, h in heads.items():
        ws.cell(HM, c, h).alignment = Alignment(horizontal="left" if c == 2 else "right", vertical="center")
    for k in range(M):
        m = dt.date(d0.year + (d0.month - 1 + k) // 12, (d0.month - 1 + k) % 12 + 1, 1)
        c = ws.cell(HM, C0 + k, m); c.number_format = "mmmmm"; c.alignment = Alignment(horizontal="center")
        c.font = Font(name=F_, size=8, color=NAVY)
        if m.month == 1 or k == 0: ws.cell(HY, C0 + k, str(m.year)).alignment = Alignment(horizontal="left")   # texto: transborda para as colunas vizinhas
    ws.row_dimensions[HM].height = 16
    # linhas
    r = R0; inf_ok = 0
    def linha(n, lv, inferido):
        nonlocal r
        ws.cell(r, 2, n["n"]).alignment = Alignment(indent=lv * 2)
        ws.cell(r, 3, _d(n["i"])).number_format = "dd/mm/yy"
        ws.cell(r, 4, _d(n["f"])).number_format = "dd/mm/yy"
        ws.cell(r, 5, round(n["pr"] / 100, 4)).number_format = "0%"
        if inferido: ws.cell(r, 6, "x")
        cor = "8C6D2F" if inferido else (NAVY if lv == 0 else INK)
        for c in range(2, 6): ws.cell(r, c).font = Font(name=F_, size=9 if lv else 10, bold=lv < 2, color=cor)
        for c in range(2, C0 + M): ws.cell(r, c).border = Border(bottom=sd)
        if lv: ws.row_dimensions[r].outlineLevel = lv
        if lv >= 2: ws.row_dimensions[r].hidden = True
        ws.row_dimensions[r].height = 16
        r += 1
    for pl in G["arv"]:
        rpl = r; linha(pl, 0, False)
        if pl.get("nx"):
            ws.cell(rpl, 2).comment = Comment(f"{pl['nx']} pacote(s) desta PL sem descrição expressa na curva do Prevision (em roxo).", "Ferramenta", width=260, height=60)
        for pk in pl.get("c", []):
            rp = r; linha(pk, 1, bool(pk.get("x")))
            if pk.get("x"):
                ws.cell(rp, 2).comment = Comment(pk.get("nt", ""), "Ferramenta", width=340, height=120); inf_ok += 1
            for lt in pk.get("c", []): linha(lt, 2, bool(pk.get("x")))
    RN = r - 1
    # barras: formatação condicional sobre as datas (planejado, realizado, concluído, inferido)
    rng = f"{L(C0)}{R0}:{L(C0 + M - 1)}{RN}"
    g = f"{L(C0)}${HM}"
    over = f"AND({g}<=$D{R0},EOMONTH({g},0)>=$C{R0})"
    real = f"{g}<=$C{R0}+($D{R0}-$C{R0})*$E{R0}"
    regras = [(f'AND($F{R0}="x",{over},{real})', "C0A062"), (f'AND($F{R0}="x",{over})', "F1E6CC"),
              (f"AND($E{R0}>=0.999,{over})", "1F7A52"), (f"AND({over},{real},$E{R0}>0)", "446756"), (over, "C5D3C8")]
    for fml, cor in regras:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[fml], fill=fill(cor), stopIfTrue=True))
    # linhas de data (situação do Prevision e 31/12 do ano do retrato): borda tracejada fixa na coluna do mês
    for dd, cor in ((ref, "B3392B"), (ret, NAVY)):
        k = (dd.year - d0.year) * 12 + dd.month - d0.month
        if 0 <= k < M:
            for rr in range(HY, RN + 1):
                c = ws.cell(rr, C0 + k); b = c.border
                c.border = Border(left=Side(style="mediumDashed", color=cor), bottom=b.bottom)
    ws.freeze_panes = ws.cell(R0, C0)
    ws.cell(RN + 2, 2, ("Fonte: aba CRONOGRAMA ATIVIDADES (export do Prevision). Datas e % realizado vêm do Prevision; nos níveis PL e pacote o % é a "
                        "média dos itens ponderada pela duração. O cronograma é o mesmo nos três cenários (os cenários mudam as curvas financeiras). "
                        f"Fotografia de {dt.date.today():%d/%m/%Y}.")).font = Font(name=F_, italic=True, size=8, color=MUT)
    _fundo(ws, 6, RN + 3)
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_title_rows = f"{HY}:{HM}"
    return ws
