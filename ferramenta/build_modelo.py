"""Gera a ferramenta de fluxo de caixa com as mesmas abas do modelo original (BOTÂNICO - MODELO 05)."""
import json, re, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1]
import os
DATA_DIR = os.environ.get("DATA_DIR", ".")      # pasta com data.json, data2.json e (opcional) cfg.json da obra
d1 = json.load(open(os.path.join(DATA_DIR, "data.json")))
d2 = json.load(open(os.path.join(DATA_DIR, "data2.json")))
CFG = json.load(open(os.path.join(DATA_DIR, "cfg.json"))) if os.path.exists(os.path.join(DATA_DIR, "cfg.json")) else {}
BLANK = os.environ.get("BLANK") == "1"          # planilha em branco para uma obra nova
_ih = os.path.join(os.path.dirname(os.path.abspath(__file__)), "incc_hist.json")
if os.path.exists(_ih):   # base histórica INCC-DI (FGV) desde ago/1994: substitui a série parcial de cada obra
    _hist = dict(json.load(open(_ih)))
    for _d, _v in d1["incc"]: _hist.setdefault(_d, _v)
    d1["incc"] = sorted(_hist.items())
OBRA_NOME = os.environ.get("OBRA_NOME", "NOVA OBRA" if BLANK else CFG.get("obra", "BOTÂNICO"))
if BLANK:
    for o in d1["orc"]: o["orc"] = None; o["proj"] = None
    d1["cur"] = {}
    for k in ("Tabela", "CURVA PREVISION", "CRONO", "DADOS BRUTOS", "AUX.ESTRUTURA", "RETRATO 2027", "AUX. CURVA"): d2[k] = {}
    d2["VALOR POR PL"] = {k: v for k, v in d2["VALOR POR PL"].items() if k.startswith("J")}
    d2["VALOR POR PL"]["B3"] = "ORÇAMENTO - " + OBRA_NOME
    d2["MODELO PARA BI"] = {k: v for k, v in d2["MODELO PARA BI"].items() if not k.startswith("A")}

F = "Calibri"
f_in = Font(name=F, color="0000FF")
f_calc = Font(name=F)
f_link = Font(name=F, color="008000")
f_b = Font(name=F, bold=True)
f_bl = Font(name=F, bold=True, color="008000")
f_h = Font(name=F, bold=True, color="FFFFFF")
f_t = Font(name=F, bold=True, size=14)
f_note = Font(name=F, italic=True, color="8C877A", size=9)
fill_h = PatternFill("solid", fgColor="1C3B2A")
fill_y = PatternFill("solid", fgColor="FFFF00")
fill_in = PatternFill("solid", fgColor="FFF2CC")
fill_tot = PatternFill("solid", fgColor="E8F0EA")
fill_past = PatternFill("solid", fgColor="EDEDED")
fill_red = PatternFill("solid", fgColor="F6E3DF")
fill_ok = PatternFill("solid", fgColor="E8F0EA")
thin = Side(style="thin", color="BFBFBF")
BRD = Border(top=thin, bottom=thin, left=thin, right=thin)
NUM = '#,##0;[Red](#,##0);"-"'
NUM2 = '#,##0.00;[Red](#,##0.00);"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
MES = "mmm/yy"
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

NM = 72                     # meses no horizonte
R0, R1 = 5, 5 + NM - 1      # linhas de meses nas abas de disciplina (5..76)
RTOT, RDIF = R1 + 2, R1 + 4 # 78, 80
NDB_COLS = 200              # colunas lidas da aba DADOS BRUTOS (B..GS)
NDB_ROWS = 500              # linhas lidas da aba DADOS BRUTOS (3..500)

VP = "'VALOR POR PL'"
INC = "'\"INCC-FGV\"'"
CORTE, INICIO, IDX, FIM, OBRA = (f"{VP}!$P$5", f"{VP}!$P$6", f"{VP}!$P$7", f"{VP}!$P$8", f"{VP}!$P$9")
INCC_A, INCC_B = f"{INC}!$A$2:$A$800", f"{INC}!$B$2:$B$800"

# ------------------------------------------------------------------ disciplinas
TABS = ["PROJETO", "SEGUROS", "TAX. DIVERSAS", "BARRACÃO DE OBRA", "LIG. PROVISÓRIAS", "CONSUMOS", "DIVERSOS",
        "TRANSPORTES", "CONSULTORIAS", "MDO", "EQUIP E FER", "SERV. INICIAIS", "CONTENÇÃO", "TERRAPLENAGEM",
        "FUND. PROF", "FUND. SUP", "ESTRUTURA", "ALVENARIA", "IMPERMEABILIZAÇÃO", "INST. ELÉTRICA", "INST. HIDRO",
        "INST. INC", "INST. ESP", "ELEVADORES ", "ESQUADRIAS", "SERRALHERIA", "MARCENARIA", "REV. PISOS",
        "REV. PAREDE", "REV. ACAB", "PAVIMENTAÇÃO", "REV. FACHADA", "FORRO", "LOUÇAS", "PINTURA", "PAISAGISMO",
        "LIMPEZA", "TAX. ADM"]
for _k, _v in CFG.get("tabs", {}).items(): TABS[int(_k)] = _v   # abas que mudam de obra para obra (ex.: PL 31)
ND = len(TABS)  # 38 (01..37 + taxa de administração)
erp = dict(d2["tab2erp"])
erp.update({"SEGUROS": "SEGUROS", "LIG. PROVISÓRIAS": "LIGAÇÕES PROVISÓRIAS AGUA, ENERGIA,TELEFONIA, ESGOTO",
            # correções: no original PISOS/PAREDES estavam trocados e PAISAGISMO puxava PAVIMENTAÇÃO
            "REV. PISOS": "REVESTIMENTOS DE MASSA EM PISOS", "REV. PAREDE": "REVESTIMENTOS DE MASSAS EM PAREDES",
            "PAISAGISMO": "PAISAGISMO"})
if CFG.get("erp"): erp = dict(CFG["erp"])
erp.update(CFG.get("erp_update", {}))
ERP2 = CFG.get("erp2", {})                         # 2º nome do ERP somado na mesma disciplina
TAXA_SO_INC = CFG.get("taxa_somente_incorrido", True)   # taxa de adm.: projeção = incorrido (padrão) ou valor informado
FATOR = CFG.get("fator_proj", {})   # reorçamento por PL: projeção ÷ orçamento (mesma base de reajuste)
META_MIN_INC = CFG.get("meta_min_incorrido", False)
CONT_MES = CFG.get("contingencia")   # mês (aaaa-mm-dd) da contingência = orçamento INCC − distribuído nas PLs
CONT = 59 if CONT_MES else None      # linha da contingência na aba MODELO PARA BI
CONC = CFG.get("conciliacao")        # incorrido informado pela medição financeira {rs, incc, fonte}
CONC_ROW = 58 if CONC else None
CONC_MES = dt.date.fromisoformat(CONC.get("mes", CFG.get("corte", "2026-08-01"))) if CONC else None   # mês do incorrido informado pela medição (padrão: corte)      # linha do ajuste de conciliação na aba MODELO PARA BI   # PL com gasto acima da projeção fecha no realizado (INCC)
CP_CODE = CFG.get("cp_code", {})                   # código do PL no Prevision, quando difere do orçamento
CP_PREV = CFG.get("cp_code_prev", {})   # PL que usa a curva PREVISTA de outro item do Prevision (o realizado continua o próprio)
RET_FIS = CFG.get("ret_fis_mais_fin", {})   # RETRATO: execução física = avanço financeiro + defasagem (p.u.)
RET_SIT = CFG.get("ret_situacao", {})       # RETRATO: texto fixo da situação em 31/12 (em vez do cronograma de atividades)
ERP_FUT_NAO = set(CFG.get("erp_futuro_nao", []))   # abas que ignoram lançamentos do ERP após o corte
orc = d1["orc"]            # 37 disciplinas + taxa (mesma ordem de TABS)
vp_old = d2["VALOR POR PL"]
VPROW = [11 + i for i in range(ND - 1)] + [48]   # linha de cada disciplina em VALOR POR PL
DHROW = [5 + i for i in range(ND)]               # linha em DADOS HISTORICOS / CP RESUMO


def q(name):
    return "'" + name.replace("'", "''") + "'"


def month_iso(i):  # i-ésimo mês do horizonte a partir de dez/2024 (pré-carga)
    return dt.date(2024 + (11 + i) // 12, (11 + i) % 12 + 1, 1).isoformat()


wb = Workbook()
wb.remove(wb.active)


def hdr(ws, row, col, labels, height=None):
    for j, t in enumerate(labels):
        if t is None: continue
        c = ws.cell(row, col + j, t)
        c.font = f_h; c.fill = fill_h; c.border = BRD; c.alignment = WRAP
    if height: ws.row_dimensions[row].height = height


def put(ws, ref, v, font=None, fmt=None, fill=None, border=True):
    c = ws[ref]; c.value = v
    if font is not None: c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = BRD
    return c


def copy_values(ws, cells, skip_empty=True):
    for k, v in cells.items():
        if skip_empty and v in (None, ""): continue
        c = ws[k]; c.value = v; c.font = f_calc


# ================================================================== LEIA-ME
ws = wb.create_sheet("LEIA-ME")
txt = [
    ("Fluxo de caixa da obra — mesmas abas do MODELO 05, agora com cálculo automático", f_t),
    ("", None),
    ("O que mudou em relação ao modelo original", f_b),
    ("• As abas são as mesmas. A diferença é que nada precisa de ajuste manual: em cada aba de disciplina, o saldo", None),
    ("  (META EM INCC − INCORRIDO EM INCC) é distribuído nos meses futuros e convertido em R$ pelo INCC do corte; a linha DIFERENÇA fica zerada.", None),
    ("• A coluna PREVISÃO (R$) agora é calculada. O que você digitava nela vai na coluna H 'PREVISÃO INFORMADA' e vira o peso", None),
    ("  de cada mês. Se a soma do que você digitou não fechar com o saldo, a ferramenta reescala (mostra o % de ajuste em K15).", None),
    ("• Regra da curva (célula K16 de cada aba): itens DIRETOS com avanço físico realizado no Prevision abaixo do limite", None),
    ("  (VALOR POR PL!K7, hoje 10%) seguem a CURVA DO PREVISION (colunas M a O da aba). Os demais usam a coluna H;", None),
    ("  sem coluna H, também usam o Prevision. Para forçar uma fonte numa aba, use a célula K15.", None),
    ("• DADOS HISTORICOS localiza cada grupo do ERP pelo NOME e cada mês pela DATA do cabeçalho (antes era por posição).", None),
    ("", None),
    ("Atualização mensal", f_b),
    ("1. \"INCC-FGV\": acrescente o índice do mês (mês no formato 01/mm/aaaa).", None),
    ("2. VALOR POR PL: mude a DATA DE CORTE (K5) para o último mês com realizado fechado. A meta de cada PL é em INCC (coluna H); a projeção em R$ (coluna L) é calculada.", None),
    ("3. DADOS BRUTOS: cole a tabela dinâmica do ERP a partir da célula A1 (mesmo formato de sempre: grupos nas linhas, meses nas colunas).", None),
    ("   Grupo destacado em vermelho = nome que não existe em DADOS HISTORICOS (colunas B/C). Cadastre-o ou coloque na lista de ignorados.", None),
    ("4. CURVA PREVISION: cole o export do Prevision a partir da célula A1 (opcional — só é usado quando a coluna H da disciplina está vazia).", None),
    ("5. Abas de disciplina: se quiser, digite a previsão por mês na coluna H (meses futuros).", None),
    ("6. VALOR POR PL, coluna N e bloco CONFERÊNCIA GERAL: tudo deve estar 'OK'.", None),
    ("7. RETRATO 2027: escolha o ano em D5 e o INCC em F5. As observações vêm da aba CRONOGRAMA ATIVIDADES (cole ali o export de atividades do Prevision).", None),
    ("", None),
    ("Convenção de cores", f_b),
    ("Texto azul = entrada (editar)", f_in),
    ("Texto preto = fórmula (não editar)", f_calc),
    ("Texto verde = puxa de outra aba (não editar)", f_link),
    ("Fundo amarelo = parâmetro principal", None),
    ("Linhas cinza nas abas de disciplina = meses já realizados (até a data de corte)", None),
    ("", None),
    ("Diferenças de abas em relação ao original", f_b),
    ("• Novas: LEIA-ME, SEGUROS e LIG. PROVISÓRIAS (essas duas disciplinas tinham projeção e realizado no ERP, mas nenhuma aba de distribuição).", None),
    ("• Removida: Planilha1 (era uma cópia com valores fixos do MODELO PARA BI).", None),
    ("• Tabela, PREMISSAS e AUX.ESTRUTURA foram copiadas como estavam (somente valores), para consulta.", None),
]
if BLANK:
    ex = Font(name=F, color="8C877A", italic=True)
    txt = [
        ("PLANILHA EM BRANCO — como começar uma obra nova", f_t),
        ("Preencha só as células/abas de ENTRADA (texto azul, fundo amarelo nos parâmetros). O resto é fórmula e se calcula sozinho.", None),
        ("Os exemplos em cinza mostram o formato esperado (valores do Botânico, só como referência).", None),
        ("", None),
        ("1. VALOR POR PL — parâmetros e orçamento", f_b),
        ("   P9 = nome da obra · K6 = início do fluxo (1º mês; o realizado anterior entra acumulado nele) · K5 = data de corte (último mês fechado)", None),
        ("   H6 = INCC base do orçamento · K7 = limite de avanço físico para usar o Prevision (padrão 10%)", None),
        ("   Por PL (linhas 11 a 48): F = orçamento base (R$) · H = meta em INCC (= F ÷ INCC base) · I = DIRETO/INDIRETO · L = projeção em R$ (calculada)", None),
        ("   Exemplo: 01 PROJETOS | F: 2.400.000,00 | L: 2.640.655,00 | I: INDIRETO", ex),
        ("2. \"INCC-FGV\" — índices mensais (já preenchidos até ago/2026; acrescente os meses seguintes)", f_b),
        ("   Exemplo: A = 01/09/2026 | B = 1.305,42 (número-índice do INCC-DI, não a variação %)", ex),
        ("3. DADOS BRUTOS — cole a tabela dinâmica do ERP a partir de A1 (Colar especial → Valores)", f_b),
        ("   Linha 2 = meses (datas) a partir de B2 · coluna A = grupos a partir da linha 3", None),
        ("   Exemplo: A2 'Rótulos de Linha' | B2 01/05/2023 | C2 01/06/2023 …   ·   A8 'ALVENARIA' | B8 12.345,67 | C8 8.910,11 …", ex),
        ("   Se o ERP da obra usar outros nomes de grupo, ajuste as colunas B/C da aba DADOS HISTORICOS.", None),
        ("4. CURVA PREVISION — cole o export 'Cronograma físico financeiro' do Prevision a partir de A1", f_b),
        ("   Exemplo: J2 '30/06/2025' (mês, sobre a coluna Previsto) · linha 3: Nível | Código EAP | Descrição | … | Base | Previsto | Realizado", ex),
        ("   Os códigos de 1º nível (coluna B: 01, 02 … 37) precisam ser iguais aos da coluna C de VALOR POR PL.", None),
        ("5. CRONOGRAMA ATIVIDADES — cole o export de atividades do Prevision a partir de A1", f_b),
        ("   Exemplo: linha 4: ID | Pacote de trabalho/tarefas | Serviço | Lote | Data de Início | Data de Término | … | Base | Previsto | Realizado", ex),
        ("   Revise o de-para pacote → PL (colunas R a U). Pacotes sem de-para aparecem em vermelho.", None),
        ("6. Abas de disciplina (opcional) — coluna H = previsão informada por mês; K15 = forçar a fonte da curva", f_b),
        ("   Padrões herdados das premissas: ELEVADORES e ESQUADRIAS com previsão informada; LOUÇAS com MDO pelo Prevision e material pela premissa.", None),
        ("7. Confira — VALOR POR PL, coluna N e bloco CONFERÊNCIA GERAL (tudo 'OK'); depois DASHBOARD e RETRATO 2027.", f_b),
        ("", None),
    ] + txt
for i, (t, f) in enumerate(txt, 1):
    ws.cell(i, 1, t).font = f or f_calc
for _i, (_t, _f) in enumerate(txt, 1):
    if _t.startswith("Fundo amarelo"): ws.cell(_i, 1).fill = fill_y
    if _t.startswith("Linhas cinza"): ws.cell(_i, 1).fill = fill_past
ws.column_dimensions["A"].width = 135
ws.sheet_view.showGridLines = False

# ================================================================== Tabela (oculta, cópia)
ws = wb.create_sheet("Tabela"); copy_values(ws, d2["Tabela"]); ws.sheet_state = "hidden"

# ================================================================== VALOR POR PL
ws = wb.create_sheet("VALOR POR PL")
put(ws, "B3", vp_old.get("B3", "ORÇAMENTO - BOTÂNICO"), Font(name=F, bold=True, size=13), border=False)
put(ws, "J3", "ACOMPANHAMENTO IEC", f_b, border=False)
put(ws, "B5", "INCC 10-25", f_b); put(ws, "C5", vp_old.get("C5"), f_in, "#,##0.00")
put(ws, "B6", "ORÇ EM INCC", f_b); put(ws, "C6", vp_old.get("C6"), f_in, NUM2)
put(ws, "B7", "ORÇ 10-25", f_b); put(ws, "C7", "=C6*C5", f_calc, NUM)
put(ws, "D5", "TOTAL", f_b); put(ws, "E5", "=E50", f_calc, "0.00%"); put(ws, "F5", "=F50", f_calc, NUM2); put(ws, "G5", "=G50", f_calc, NUM2)
put(ws, "D7", "TAXA DE ADMINISTRAÇÃO", f_b); put(ws, "E7", "=IF(F49=0,0,F48/F49)", f_calc, "0.00%")
put(ws, "F7", "=F48", f_calc, NUM2); put(ws, "G7", "=G48", f_calc, NUM2)
put(ws, "D8", "SUBTOTAL", f_b); put(ws, "E8", "=IF(F50=0,0,F49/F50)", f_calc, "0.00%"); put(ws, "F8", "=F49", f_calc, NUM2); put(ws, "G8", "=G49", f_calc, NUM2)
put(ws, "H5", "INCC BASE DO ORÇAMENTO", f_b)
put(ws, "H6", vp_old.get("H6"), f_in, "#,##0.00")
ws["H6"].comment = Comment("Índice usado para converter o orçamento (coluna F) em INCC (coluna G). Valor do arquivo original.", "Ferramenta")
put(ws, "J5", "DATA DE CORTE ->", f_b); put(ws, "K5", dt.date.fromisoformat(CFG.get("corte", "2026-08-01")), f_in, "dd/mm/yyyy", fill_y)
ws["K5"].comment = Comment("Último mês com realizado fechado. Meses até aqui = INCORRIDO; depois = PREVISÃO calculada. "
                           "Original: 'INCORRIDO ATÉ -> AGO/2026'.", "Ferramenta")
put(ws, "J6", "INÍCIO DO FLUXO ->", f_b); put(ws, "K6", dt.date.fromisoformat(CFG.get("inicio", "2024-12-01")), f_in, "dd/mm/yyyy", fill_y)
put(ws, "J7", "LIMITE DE AVANÇO FÍSICO P/ USAR PREVISION ->", f_b); put(ws, "K7", CFG.get("limite", 0.10), f_in, "0%", fill_y)
CENARIO = os.environ.get("CENARIO", CFG.get("cenario", "PREVISION"))
CEN_NOME = CFG.get("cenario_nome", "Prevision")
CEN_NOME_P = CFG.get("cenario_nome_padrao", "padrão")   # nome exibido do cenário PADRÃO (curvas digitadas)
# DESLOCAMENTO: reduz o desembolso de um ano (fator em VALOR POR PL!K9) e joga a diferença para o ano seguinte
_dl = CFG.get("desloc") or {}
DESLOC = dict(de=_dl["de"], para=_dl["para"], fator=_dl["fator"][CENARIO]) if CENARIO in _dl.get("fator", {}) else None   # nome exibido no subtítulo dos painéis
put(ws, "J8", "CENÁRIO DAS CURVAS ->", f_b); put(ws, "K8", CENARIO, f_in, fill=fill_y)
ws["K8"].comment = Comment("PREVISION: itens diretos com avanço físico abaixo do limite (K7) seguem a curva do Prevision; os demais, a curva digitada. "
                           "PADRÃO: todos seguem a curva digitada (coluna H das abas), como na versão original. "
                           "Nos dois cenários ficam as premissas especiais (elevadores, esquadrias, louças).", "Ferramenta")
_dvc = DataValidation(type="list", formula1='"PREVISION,PADRÃO"', allow_blank=False); ws.add_data_validation(_dvc); _dvc.add("K8")
if DESLOC:
    put(ws, "J9", f"FATOR SOBRE A PREVISÃO DE {DESLOC['de']} (diferença vai para {DESLOC['para']}) ->", f_b)
    put(ws, "K9", DESLOC["fator"], f_in, "0.000000", fill_y)
    ws["K9"].comment = Comment(f"Multiplica a previsão (saldo a desembolsar) de {DESLOC['de']} em todas as PLs. A parte retirada de cada PL vai para "
                               f"{DESLOC['para']}, proporcional à curva da própria PL em {DESLOC['para']} (sem curva nesse ano: mesmo mês do ano seguinte). "
                               "O total de cada PL (INCC e R$) não muda. 1 = sem deslocamento.", "Ferramenta")
ws["K7"].comment = Comment("Itens DIRETOS com avanço físico realizado (Prevision, no mês de corte) abaixo deste limite "
                           "usam a curva do Prevision para distribuir o saldo. Regra definida pelo usuário.", "Ferramenta")
ws["K6"].comment = Comment("1º mês do horizonte (72 meses). Todo o realizado anterior entra acumulado neste mês ('Acu até').", "Ferramenta")
params = [(5, "Corte (1º dia do mês)", "=DATE(YEAR(K5),MONTH(K5),1)", "dd/mm/yyyy"),
          (6, "Início (1º dia do mês)", "=DATE(YEAR(K6),MONTH(K6),1)", "dd/mm/yyyy"),
          (7, "INCC no mês de corte", f'=IFERROR(INDEX({INCC_B},MATCH(P5,{INCC_A},0)),IFERROR(INDEX({INCC_B},MATCH(P5,{INCC_A},1)),"FALTA INCC"))', "#,##0.00"),
          (8, "Último mês do horizonte", f"=DATE(YEAR(P6),MONTH(P6)+{NM - 1},1)", "dd/mm/yyyy")]
for r, lab, f, fmt in params:
    put(ws, f"O{r}", lab, f_calc); put(ws, f"P{r}", f, f_link if "!" in f else f_calc, fmt)
put(ws, "O9", "Obra", f_calc); put(ws, "P9", OBRA_NOME, f_in, fill=fill_y)
hdr(ws, 10, 2, ["CPU", "CÓDIGO", "DESCRIÇÃO / SERVIÇO", "PESO DO ORÇAMENTO", "VALOR EM R$\n(orçamento base)",
                "VALOR EM INCC\n(orçamento)", "META EM INCC\n(soma da\ndistribuição)", "TIPO", "PL", "ORÇAMENTO REAJUSTADO\n(INCC do corte)",
                "PROJEÇÃO (R$)", "VARIAÇÃO", "CONFERÊNCIA", "FONTE DA CURVA", "AVANÇO FÍSICO\n(Prevision, corte)",
                "CÓD. NO\nPREVISION", "CURVA PREVISTA\nDE OUTRO PL\n(cód. Prevision)", "FATOR DA\nPROJEÇÃO\n(reorçamento)"], height=45)
for i, t in enumerate(TABS):
    r = VPROW[i]; o = orc[i]
    old_r = 11 + i if i < ND - 1 else None
    put(ws, f"B{r}", o["cpu"] or None, f_in)
    put(ws, f"C{r}", f"{int(float(o['cod'])):02d}", f_in)
    put(ws, f"D{r}", o["disc"], f_in)
    put(ws, f"E{r}", f"=IF($F$50=0,0,F{r}/$F$50)", f_calc, "0.00%")
    put(ws, f"F{r}", None if o["orc"] is None else round(o["orc"], 2), f_in, NUM2)
    if o.get("incc") is not None:   # valor em INCC do orçamento original (planilha da obra)
        put(ws, f"G{r}", round(o["incc"], 6), f_in, NUM2)
    else:
        put(ws, f"G{r}", f"=IF($H$6=0,0,F{r}/$H$6)", f_calc, NUM2)
    if i == ND - 1 and TAXA_SO_INC:   # taxa: somente o incorrido → meta = incorrido em INCC
        put(ws, f"H{r}", f"={q(t)}!$K$8", f_link, NUM2)
    else:
        if META_MIN_INC:   # PL fecha no que foi efetivamente gasto: meta nunca abaixo do incorrido em INCC
            put(ws, f"H{r}", f'=MAX(G{r}*IF(S{r}="",1,S{r}),{q(t)}!$K$8)', f_calc, NUM2)
        else:
            put(ws, f"H{r}", f'=G{r}*IF(S{r}="",1,S{r})', f_calc, NUM2)
    put(ws, f"S{r}", FATOR.get(t), f_in, "0.0000")
    n_ = int(float(o["cod"]))
    put(ws, f"I{r}", "DIRETO" if 12 <= n_ <= 37 else "INDIRETO", f_in)
    put(ws, f"J{r}", vp_old.get(f"J{old_r}") if old_r else "TAXA DE ADMINISTRAÇÃO", f_in)
    put(ws, f"K{r}", f"=IF(ISNUMBER($P$7),G{r}*$P$7,0)", f_calc, NUM2)
    put(ws, f"L{r}", f"={q(t)}!$J$7", f_link, NUM2)   # projeção em R$ = resultado da distribuição em INCC
    put(ws, f"M{r}", f"=L{r}-K{r}", f_calc, NUM2)
    T = q(t)
    put(ws, f"N{r}", (f'=IF(D{r}="","",IF({T}!$K$8>{T}!$K$6+0.01,"⚠ Incorrido acima da meta em INCC",'
                      f'IF(AND({T}!$K$18>1,{T}!$K$17=0),"⚠ Saldo sem curva (preencha a coluna H da aba)",'
                      f'IF(ABS({T}!$D${RDIF})>0.01,"⚠ Verificar","OK"))))'), f_calc)
    put(ws, f"O{r}", f"={T}!$K$16", f_link)
    put(ws, f"P{r}", f"={T}!$K$14", f_link, "0.0%")
    put(ws, f"Q{r}", CP_CODE.get(t, f"{n_:02d}" if i < ND - 1 else "00"), f_in)
    put(ws, f"R{r}", CP_PREV.get(t), f_in)
ws["Q10"].comment = Comment("Código do PL no cronograma do Prevision (coluna B da aba CURVA PREVISION). "
                           "Normalmente igual ao código do orçamento; mude quando o Prevision numerar diferente.", "Ferramenta")
ws["L10"].comment = Comment("Calculada: incorrido em R$ + (meta em INCC − incorrido em INCC) × INCC do corte. "
                           "O controle da obra é em INCC: a distribuição de cada PL soma a coluna H.", "Ferramenta")
ws["G10"].comment = Comment("Valor em INCC do orçamento, copiado da aba VALOR POR PL da planilha original da obra "
                           "(nas obras novas: orçamento em R$ ÷ INCC base).", "Ferramenta")
ws["H10"].comment = Comment("Meta da distribuição em INCC = INCC do orçamento (G) × fator da projeção (S)"
                           + (" — nunca abaixo do incorrido em INCC (a PL fecha no que foi efetivamente gasto)" if META_MIN_INC else "") + ". A soma do fluxo em INCC de cada PL é igual a este valor. "
                           "Taxa de administração 'somente incorrido': meta = incorrido em INCC.", "Ferramenta")
put(ws, "D49", "TOTAIS (sem taxa de adm.)", f_b)
put(ws, "D50", "TOTAL COM TAXA DE ADMINISTRAÇÃO", f_b)
for col in "EFGHKLM":
    put(ws, f"{col}49", f"=SUM({col}11:{col}47)", f_b, "0.00%" if col == "E" else NUM2, fill_tot)
    put(ws, f"{col}50", f"=SUM({col}11:{col}48)", f_b, "0.00%" if col == "E" else NUM2, fill_tot)
# conferência geral
put(ws, "D53", "CONFERÊNCIA GERAL", f_h, fill=fill_h); put(ws, "E53", "RESULTADO", f_h, fill=fill_h)
put(ws, "F53", "STATUS", f_h, fill=fill_h); ws.merge_cells("F53:H53")
MRANGE = "'MODELO PARA BI'!$H$1:$CA$1"
DH = "'DADOS HISTORICOS'"
_HB = L(CI("CB") + NDB_COLS - 1)   # última coluna do bloco auxiliar de DADOS HISTORICOS
checks = [
    ("Disciplinas com pendência (coluna N)", '=COUNTIF(N11:N48,"⚠*")', '=IF(E54=0,"OK","⚠ Veja a coluna N")', "0"),
    ("Grupos do ERP sem correspondência", "='DADOS HISTORICOS'!C2", '=IF(E55=0,"OK","⚠ Veja nomes em vermelho em DADOS BRUTOS")', "0"),
    ("Valores do ERP fora do horizonte (R$)", "='DADOS HISTORICOS'!BX44", '=IF(ABS(E56)<1,"OK","⚠ Aumente o horizonte ou revise datas")', NUM),
    ("INCC do mês de corte", '=IF(ISNUMBER(P7),P7,0)', f'=IF(COUNTIF({INCC_A},P5)>0,"OK",IF(E57>0,"⚠ INCC do mês de corte não cadastrado: usando o último índice disponível","⚠ Inclua o índice na aba INCC-FGV"))', "#,##0.00"),
    ("Meses realizados sem INCC",
     (f"=SUMPRODUCT(({MRANGE}<=P5)*(COUNTIF({INCC_A},{MRANGE})=0))"
      f"+SUMPRODUCT(({DH}!$CB$1:${_HB}$1=0)*({DH}!$CB$4:${_HB}$4<>\"\")*({DH}!$CB$4:${_HB}$4<=P5)*(ABS({DH}!$CB$44:${_HB}$44)>0.5))"),
     '=IF(E58=0,"OK","⚠ Complete a aba INCC-FGV (há mês com desembolso sem índice)")', "0"),
    ("Fluxo total − meta total (INCC)", ("='MODELO PARA BI'!E63" + ("-'MODELO PARA BI'!E59" if CONT else "") + "-H50"), '=IF(H50=0,"⚠ Preencha o orçamento (coluna F)",IF(ABS(E59)<0.01,"OK","⚠ Distribuição em INCC diferente da meta"))', "#,##0.00"),
]
for k, (lab, f, st, fmt) in enumerate(checks):
    r = 54 + k
    put(ws, f"D{r}", lab, f_calc); put(ws, f"E{r}", f, f_link if "!" in f else f_calc, fmt); put(ws, f"F{r}", st, f_calc)
    ws.merge_cells(f"F{r}:H{r}")
for rng in ("N11:N48", "F54:F59"):
    top = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({top},1)="⚠"'], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], fill=fill_ok))
if CONC:
    put(ws, "D62", f"CONCILIAÇÃO COM A MEDIÇÃO FINANCEIRA (incorrido até {CONC_MES.month:02d}/{CONC_MES.year})", f_h, fill=fill_h)
    for col, t_ in zip("EFG", ("INFORMADO\n(medição)", "ERP\n(ferramenta)", "AJUSTE\n(informado − ERP)")):
        c = put(ws, f"{col}62", t_, f_h, fill=fill_h); c.alignment = WRAP
    ws.row_dimensions[62].height = 30
    _cm = f"DATE({CONC_MES.year},{CONC_MES.month},1)"
    _tr = lambda col: "+".join(f"SUMPRODUCT(({q(t)}!$B${R0}:$B${R1}<={_cm})*{q(t)}!${col}${R0}:${col}${R1})" for t in TABS)
    for r_, lab_, v_, col_, fm_ in ((63, "Incorrido (R$)", CONC["rs"], "C", NUM2), (64, "Incorrido (INCC)", CONC["incc"], "L", "#,##0.00")):
        put(ws, f"D{r_}", lab_, f_calc); put(ws, f"E{r_}", v_, f_in, fm_, fill_y)
        put(ws, f"F{r_}", "=" + _tr(col_), f_link, fm_); put(ws, f"G{r_}", f"=E{r_}-F{r_}", f_b, fm_)
    ws["E63"].comment = Comment("Fonte: " + CONC.get("fonte", "medição financeira") + f". O ajuste (coluna G) entra no mês da medição ({CONC_MES.month:02d}/{CONC_MES.year}) na aba "
                                "MODELO PARA BI (linha AJUSTE DE CONCILIAÇÃO): R$ na curva em R$ e INCC no total em INCC. "
                                "Como a meta de cada PL já é o custo total (incorrido + saldo), o mesmo INCC sai do saldo futuro das PLs (linha 65): sem dupla contagem.", "Ferramenta")
    # sem dupla contagem: o incorrido a mais da medição sai do saldo futuro das PLs, proporcional ao saldo de cada uma
    put(ws, "D65", "Redução do saldo futuro das PLs (proporcional ao saldo)", f_calc)
    put(ws, "F65", "=" + "+".join(f"MAX(0,{q(t)}!$K$6-{q(t)}!$K$8)" for t in TABS), f_link, "#,##0.00")
    put(ws, "G65", "=IF(F65>0,1-G64/F65,1)", f_b, "0.000000")
    ws["F65"].comment = Comment("Soma dos saldos positivos das PLs (meta − incorrido do ERP, em INCC).", "Ferramenta")
    ws["G65"].comment = Comment("Fator aplicado ao saldo em INCC de cada PL (aba da PL, K10) = 1 − ajuste em INCC ÷ soma dos saldos. "
                                "O incorrido que a medição tem a mais que o ERP já está dentro da meta das PLs; por isso sai do saldo a distribuir.", "Ferramenta")
from openpyxl.worksheet.datavalidation import DataValidation
dvt = DataValidation(type="list", formula1='"DIRETO,INDIRETO"', allow_blank=True); ws.add_data_validation(dvt); dvt.add("I11:I48")
ws.conditional_formatting.add("O11:O48", CellIsRule(operator="equal", formula=['"CURVA PREVISION"'], fill=PatternFill("solid", fgColor="E8F0EA"), font=Font(name=F, bold=True, color="1C3B2A")))
ws["R10"].comment = Comment("Opcional. Código de outro item do Prevision cuja curva PREVISTA este PL deve seguir "
                           "(ex.: hidráulica com o andamento da elétrica). Vazio = usa a própria linha. O realizado continua o do próprio PL.", "Ferramenta")
ws["S10"].comment = Comment("Reorçamento (projeção) da PL: meta em INCC = INCC do orçamento × fator. Fator = projeção ÷ orçamento, "
                           "os dois na mesma base de reajuste (" + CFG.get("fator_fonte", "planilha de medição financeira, aba ANÁLISE IEC: coluna W ÷ coluna H") + "). "
                           "Vazio = 1 (sem reorçamento).", "Ferramenta")
for col, w in zip("ABCDEFGHIJKLMNOPQRS", (2, 11, 8, 46, 11, 17, 15, 13, 11, 40, 18, 17, 15, 44, 24, 14, 10, 14, 12)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "E11"

# ---- estilo dos gráficos no padrão EPO (planilha de medição financeira): fundo creme, texto cinza-quente, grade bege
def estilo_grafico(ch_):
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
    txt = lambda: RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800, solidFill="8C877A", latin=DFont(typeface="Calibri"))),
                                        endParaRPr=CharacterProperties())])
    ch_.graphical_properties = GraphicalProperties(solidFill="FCFBF8", ln=LineProperties(solidFill="E5DFD2"))
    for sub in ch_._charts:
        for ax in (getattr(sub, "x_axis", None), getattr(sub, "y_axis", None)):
            if ax is None: continue
            ax.txPr = txt()
            ax.spPr = GraphicalProperties(ln=LineProperties(solidFill="DED8C9"))
            if ax.majorGridlines is not None: ax.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="EDE8DC"))
    if ch_.legend is not None: ch_.legend.txPr = txt()
# ---- tema dos painéis (padrão EPO): fundo bege, cartões creme com contorno
PBG, PLINE = "F4F1EA", "E5DFD2"
def card_box(ws_, c0, c1, r0, r1):
    cols = [L(c) for c in range(CI(c0), CI(c1) + 1)]
    sd = Side(style="thin", color=PLINE)
    for rr in range(r0, r1 + 1):
        for j, col in enumerate(cols):
            cell = ws_[f"{col}{rr}"]; cell.fill = PatternFill("solid", fgColor="FCFBF8")
            cell.border = Border(top=sd if rr == r0 else None, bottom=sd if rr == r1 else None,
                                 left=sd if j == 0 else None, right=sd if j == len(cols) - 1 else None)
def painel_fundo(ws_, max_col, max_row):
    """Pinta de cinza-claro (fundo do painel) as células sem preenchimento."""
    bgf = PatternFill("solid", fgColor=PBG)
    for row in ws_.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for c in row:
            if type(c).__name__ == "MergedCell": continue
            if c.fill is None or c.fill.fill_type is None: c.fill = bgf
# ================================================================== RETRATO 2027 (painel; conteúdo montado depois do MODELO PARA BI)
ws = wb.create_sheet("RETRATO 2027")
ret = d2["RETRATO 2027"]
BIROW = {}  # disciplina i -> linha em MODELO PARA BI (preenchido abaixo)
RH = 11     # linha do cabeçalho da tabela
layout = []  # (linha, índice da disciplina) ou (linha, ('sub', rótulo, código, [linhas]))
r = RH + 1
for i in range(11): layout.append((r, i)); r += 1
layout.append((r, ("sub", "INDIRETOS", "01 a 11", list(range(RH + 1, r))))); r += 1
a0 = r
for i in range(11, 16): layout.append((r, i)); r += 1
layout.append((r, ("sub", "SERVIÇOS INICIAIS, CONTENÇÃO, TERRAPLENAGEM E FUNDAÇÕES", "12 a 16", list(range(a0, r))))); r += 1
for i in range(16, 37): layout.append((r, i)); r += 1
RT_TOT = r          # total sem taxa
RET_ROWS = layout
RROW = {item: rr for rr, item in layout if isinstance(item, int)}
RROW[ND - 1] = RT_TOT + 1
ANO_RT = "'RETRATO 2027'!$D$5"
DATA_RT = "'RETRATO 2027'!$K$5"

# ================================================================== MODELO PARA BI
bi = wb.create_sheet("MODELO PARA BI")
biold = d2["MODELO PARA BI"]
hdr(bi, 1, 1, ["COD EMPREENDIMENTO", "COD GRUPO 1", "GRUPO 1", "VALOR ORÇADO\n(INCC)", "SOMA DA DISTRIBUIÇÃO\n(INCC)",
               "VARIAÇÃO\n(ORÇADO − DISTRIBUÍDO)"], height=45)
for k in range(NM):
    col = 8 + k
    f = f"={INICIO}" if k == 0 else f"=DATE(YEAR({L(col - 1)}1),MONTH({L(col - 1)}1)+1,1)"
    c = bi.cell(1, col, f); c.number_format = MES; c.font = f_h; c.fill = fill_h; c.alignment = WRAP
    bi.column_dimensions[L(col)].width = 10
BIL = L(8 + NM - 1)  # CA
name2tab = {}
for t in TABS:
    name2tab[t] = t
# grupo do BI -> aba (pela fórmula TRANSPOSE do original + novas abas)
bi_map = {"ALVENARIA": "ALVENARIA", "BARRACAO DE OBRA": "BARRACÃO DE OBRA", "CONSULTORIAS (OBRAS)": "CONSULTORIAS",
          "CONSUMO DE ÁGUA, ENERGIA, ETC.": "CONSUMOS", "CONTENÇÃO": "CONTENÇÃO", "DIVERSOS": "DIVERSOS",
          "ELEVADORES": "ELEVADORES ", "EQUIPAMENTOS E FERRAMENTAS": "EQUIP E FER", "ESQUADRIAS": "ESQUADRIAS",
          "ESTRUTURA EM CONCRETO ARMADO": "ESTRUTURA", "FORRO": "FORRO", "FUNDAÇÃO PROFUNDA": "FUND. PROF",
          "FUNDAÇÃO SUPERFICIAL": "FUND. SUP", "IMPERMEABILIZAÇÃO": "IMPERMEABILIZAÇÃO",
          "INSTALAÇÕES COMBATE INCÊNDIO": "INST. INC", "INSTALAÇÕES ELÉTRICAS": "INST. ELÉTRICA",
          "INSTALAÇÕES ESPECIAIS": "INST. ESP", "INSTALAÇÕES HIDROSANITÁRIAS": "INST. HIDRO",
          "LOUÇAS SANITÁRIAS": "LOUÇAS", "LIMPEZA FINAL DE OBRA": "LIMPEZA", "MARCENARIA": "MARCENARIA",
          "PAISAGISMO": "PAISAGISMO", "PAVIMENTAÇÃO": "PAVIMENTAÇÃO", "PINTURA": "PINTURA", "PROJETOS": "PROJETO",
          "REVESTIMENTO DIVERSOS - FACHADA": "REV. FACHADA", "REVESTIMENTOS DE ACABAMENTO": "REV. ACAB",
          "REVESTIMENTOS DE MASSA EM PISOS": "REV. PISOS", "REVESTIMENTOS DE MASSA EM PAREDES": "REV. PAREDE",
          "SEGUROS": "SEGUROS", "SERRALHERIA": "SERRALHERIA", "TAXA DE ADMINISTRAÇÃO": "TAX. ADM",
          "TERRAPLENAGEM": "TERRAPLENAGEM", "TRANSPORTES": "TRANSPORTES", "MÃO DE OBRA PRÓPRIA": "MDO",
          "TAXAS DIVERSAS": "TAX. DIVERSAS", "SERVIÇOS INICIAIS": "SERV. INICIAIS",
          "LIGAÇÕES PROVISÓRIAS AGUA, ENERGIA,TELEFONIA, ESGOTO": "LIG. PROVISÓRIAS"}
bi_map.update(CFG.get("bi_extra", {}))
bi_map = {k: v for k, v in bi_map.items() if v in TABS}
TAXA_BI = None
for r in range(2, 60):   # grupos do BI (2-57) + ajuste de conciliação (58) + contingência (59)
    for c in "ABC":
        v = biold.get(f"{c}{r}") if r <= 57 else None   # abaixo da 57 o original tem os totais antigos
        put(bi, f"{c}{r}", v, f_in)
    name = (biold.get(f"C{r}") or "").strip() if r <= 57 else ""
    t = bi_map.get(name)
    put(bi, f"E{r}", f"=SUM(H{r}:{BIL}{r})", f_calc, NUM2)
    put(bi, f"F{r}", f'=IF(D{r}="","",D{r}-E{r})', f_calc, NUM2)
    if t:
        i = TABS.index(t); BIROW[i] = r
        if t == "TAX. ADM": TAXA_BI = r
        put(bi, f"D{r}", f"={q(t)}!$K$5", f_link, NUM2)
        for k in range(NM):
            c = bi.cell(r, 8 + k, f"={q(t)}!E{R0 + k}"); c.font = f_link; c.number_format = NUM2
assert len(BIROW) == ND, (len(BIROW), [t for i, t in enumerate(TABS) if i not in BIROW])
if CONC:   # AJUSTE DE CONCILIAÇÃO: diferença entre o incorrido da medição financeira e o do ERP, no mês de corte
    assert CONC_ROW not in BIROW.values() and not bi[f"C{CONC_ROW}"].value, "linha do ajuste ocupada"
    put(bi, f"A{CONC_ROW}", biold.get("A2"), f_in); put(bi, f"B{CONC_ROW}", None, f_in)
    put(bi, f"C{CONC_ROW}", "AJUSTE DE CONCILIAÇÃO (medição)", Font(name=F, bold=True, color="8C6D2F"))
    put(bi, f"D{CONC_ROW}", None, f_calc, NUM2)
    for k in range(NM):
        c = bi.cell(CONC_ROW, 8 + k, f"=IF({L(8 + k)}$1=DATE({CONC_MES.year},{CONC_MES.month},1),{VP}!$G$64,0)"); c.font = f_calc; c.number_format = NUM2
    bi[f"C{CONC_ROW}"].comment = Comment("Incorrido informado pela medição financeira − incorrido do ERP (VALOR POR PL, linhas 63-64), "
                                         f"lançado em {CONC_MES.month:02d}/{CONC_MES.year} (mês da medição).", "Ferramenta")
if CONT:   # CONTINGÊNCIA: diferença entre o orçamento em INCC e o distribuído nas PLs, lançada num único mês
    assert CONT not in BIROW.values() and not bi[f"C{CONT}"].value, "linha da contingência ocupada"
    _cd = dt.date.fromisoformat(CONT_MES)
    put(bi, f"A{CONT}", biold.get("A2"), f_in); put(bi, f"B{CONT}", None, f_in)
    put(bi, f"C{CONT}", "CONTINGÊNCIA", Font(name=F, bold=True, color="8A6D3B"))
    put(bi, f"D{CONT}", None, f_calc, NUM2)
    _dist = "+".join(f"{q(t)}!$E${RTOT}" for t in TABS)
    for k in range(NM):
        c = bi.cell(CONT, 8 + k, f"=IF({L(8 + k)}$1=DATE({_cd.year},{_cd.month},1),MAX(0,{VP}!$G$50-({_dist})" + (f"-{VP}!$G$64" if CONC else "") + "),0)")
        c.font = f_calc; c.number_format = NUM2
    bi[f"C{CONT}"].comment = Comment(f"Orçamento em INCC (VALOR POR PL!G50) − soma distribuída nas PLs (projeções). "
                                     f"Lançada em {_cd.month:02d}/{_cd.year}; no fluxo em R$ entra pelo INCC do corte. "
                                     f"Variação negativa (projeções acima do orçamento) não gera contingência: fica zero.", "Ferramenta")
lab = [(61, "SOMA SEM TAXA DE ADM. (INCC)"), (62, "TAXA DE ADMINISTRAÇÃO (INCC)"), (63, "TOTAL (INCC)"),
       (65, "CURVA MENSAL (R$ nominal)"), (66, "CURVA ACUMULADA (R$ nominal)"), (67, "% ACUMULADO DA PROJEÇÃO")]
for r, t in lab:
    put(bi, f"C{r}", t, f_b)
for col in ["D", "E"] + [L(8 + k) for k in range(NM)]:
    put(bi, f"{col}61", f"=SUM({col}2:{col}59)-{col}{TAXA_BI}", f_b, NUM2, fill_tot)
    put(bi, f"{col}62", f"={col}{TAXA_BI}", f_b, NUM2, fill_tot)
    put(bi, f"{col}63", f"={col}61+{col}62", f_b, NUM2, fill_tot)
for k in range(NM):
    col = L(8 + k); rr = R0 + k
    put(bi, f"{col}65", "=" + "+".join(f"{q(t)}!C{rr}+{q(t)}!D{rr}" for t in TABS) + (f"+{col}{CONT}*{IDX}" if CONT else "") + (f"+IF({col}$1=DATE({CONC_MES.year},{CONC_MES.month},1),{VP}!$G$63,0)" if CONC else ""), f_calc, NUM)
    put(bi, f"{col}66", f"={col}65" if k == 0 else f"={L(7 + k)}66+{col}65", f_calc, NUM)
    put(bi, f"{col}67", f"=IF($E$66=0,0,{col}66/$E$66)", f_calc, PCT)
put(bi, "E65", f"=SUM(H65:{BIL}65)", f_b, NUM)
put(bi, "E66", f"={BIL}66", f_b, NUM)
put(bi, "C64", '="INCORRIDO ATÉ -> "&TEXT(MONTH(' + CORTE + '),"00")&"/"&YEAR(' + CORTE + ')', f_link)
bi.conditional_formatting.add(f"H2:{BIL}63", FormulaRule(formula=[f"H$1<={CORTE}"], fill=fill_past))
for col, w in zip("ABCDEFG", (12, 15, 40, 14, 14, 16, 2)): bi.column_dimensions[col].width = w
bi.freeze_panes = "H2"
ch = LineChart(); ch.title = "Curva S — acumulado (R$)"; ch.height = 8; ch.width = 26; ch.legend = None
ch.add_data(Reference(bi, min_col=8, max_col=8 + NM - 1, min_row=66), from_rows=True)
ch.set_categories(Reference(bi, min_col=8, max_col=8 + NM - 1, min_row=1)); ch.x_axis.number_format = MES
bi.add_chart(ch, "C70")

# ---- RETRATO 2027 como painel (agora que BIROW existe)
ws = wb["RETRATO 2027"]
ws.sheet_view.showGridLines = False; ws.sheet_view.zoomScale = 90
YR = f"YEAR({MRANGE})"
NAVY_, MUT_, INK_ = "1C3B2A", "8C877A", "3C3A33"
fl = lambda h: PatternFill("solid", fgColor=h)
sd_n = Side(style="medium", color=NAVY_); sd_t = Side(style="thin", color="EDE8DC")
CR = "'CRONOGRAMA ATIVIDADES'"
MESES_C = '"jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"'
for col, w in zip("ABCDEFGHIJKLMN", (2, 8, 40, 13, 13, 17, 13, 17, 9, 14, 78, 34, 2, 2)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["M"].hidden = True; ws.column_dimensions["N"].hidden = True
# faixa de título
put(ws, "B2", f'={OBRA}&"  ·  RETRATO DE "&D5', Font(name=F, bold=True, size=20, color=NAVY_), border=False)
put(ws, "B3", (f'="Execução física (Prevision), desembolso do ano (fluxo financeiro) e situação da obra em 31/12/"&D5'
               f'&"     ·     Corte: "&CHOOSE(MONTH({CORTE}),{MESES_C})&"/"&YEAR({CORTE})&"     ·     Cenário: "&IF({VP}!$K$8="PADRÃO","{CEN_NOME_P}","{CEN_NOME}")'),
    Font(name=F, size=10, color="8C877A"), border=False)
ws.row_dimensions[2].height = 34; ws.row_dimensions[3].height = 20
ws["B2"].alignment = Alignment(vertical="center")
# parâmetros
lab_ = Font(name=F, bold=True, size=9, color=MUT_)
put(ws, "B5", "ANO DO RETRATO →", lab_, border=False); ws.merge_cells("B5:C5"); ws["B5"].alignment = Alignment(horizontal="right")
put(ws, "D5", 2027, Font(name=F, bold=True, size=12, color="0000FF"), "0", fill_y)
put(ws, "E5", "INCC →", lab_, border=False); ws["E5"].alignment = Alignment(horizontal="right")
put(ws, "F5", f"={CORTE}", Font(name=F, bold=True, color="008000"), "mm/yyyy", fill_y)
put(ws, "G5", f'=IFERROR(INDEX({INCC_B},MATCH(DATE(YEAR(F5),MONTH(F5),1),{INCC_A},0)),IFERROR(INDEX({INCC_B},MATCH(DATE(YEAR(F5),MONTH(F5),1),{INCC_A},1)),"FALTA INCC"))', f_link, "#,##0.00")
ws["F5"].comment = Comment("Mês do INCC usado para converter INCC em R$. Por padrão acompanha a data de corte; pode digitar outro mês. "
                           "Se o índice do mês não estiver cadastrado, usa o último disponível.", "Ferramenta")
put(ws, "J5", "SITUAÇÃO EM →", lab_, border=False); ws["J5"].alignment = Alignment(horizontal="right")
put(ws, "K5", "=DATE(D5,12,31)", Font(name=F, bold=True, color=NAVY_), "dd/mm/yyyy", border=False)
ws["K5"].alignment = Alignment(horizontal="left")
# cartões
TR0, TR1 = RH + 1, RT_TOT - 1
cards = [("B", "C", "DESEMBOLSO NO ANO (R$ NOMINAL)", f"=F{RT_TOT + 2}", '"R$" #,##0', "com taxa de administração"),
         ("D", "E", "DESEMBOLSO NO ANO (INCC)", f"=G{RT_TOT + 2}", '#,##0.00', "pontos de índice"),
         ("F", "G", "EM R$ DO INCC SELECIONADO", f"=H{RT_TOT + 2}", '"R$" #,##0', "INCC × índice do mês escolhido"),
         ("H", "I", "EXECUÇÃO FÍSICA NO FIM DO ANO", f"=IFERROR(SUMPRODUCT(D{TR0}:D{TR1},N{TR0}:N{TR1})/SUMPRODUCT(ISNUMBER(D{TR0}:D{TR1})*N{TR0}:N{TR1}),0)", "0.0%", "itens diretos, ponderado pela projeção"),
         ("J", "J", "AVANÇO FINANCEIRO", f"=IFERROR(SUMPRODUCT(J{TR0}:J{TR1},M{TR0}:M{TR1})/SUM(M{TR0}:M{TR1}),0)", "0.0%", "acumulado no fim do ano"),
         ("K", "K", "MAIOR DESEMBOLSO DO ANO", None, None, None)]
for c0, c1, lab, f, fmt, note in cards:
    card_box(ws, c0, c1, 7, 9)
    put(ws, f"{c0}7", lab, Font(name=F, bold=True, size=8, color=MUT_), border=False)
    if f: put(ws, f"{c0}8", f, Font(name=F, bold=True, size=15, color=NAVY_), fmt, border=False)
    if note: put(ws, f"{c0}9", note, Font(name=F, size=8, color=MUT_), border=False)
    for rr in (7, 8, 9): ws[f"{c0}{rr}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    if c0 != c1:
        for rr in (7, 8, 9): ws.merge_cells(f"{c0}{rr}:{c1}{rr}")
dir_rows = [rr for rr, it in RET_ROWS if isinstance(it, int)]
# só linhas de PL (código com 2 caracteres); as linhas de subtotal ("01 a 11") ficam de fora
_mx = f'MAX(IF(LEN(B{TR0}:B{TR1})=2,F{TR0}:F{TR1}))'
ws["K8"] = ArrayFormula("K8", f"=INDEX(C{TR0}:C{TR1},MATCH({_mx},IF(LEN(B{TR0}:B{TR1})=2,F{TR0}:F{TR1}),0))")
ws["K8"].font = Font(name=F, bold=True, size=12, color=NAVY_)
ws["K9"] = ArrayFormula("K9", f'="R$ "&FIXED({_mx}/1000000,1)&" mi no ano"')
ws["K9"].font = Font(name=F, size=8, color=MUT_)
for rr in (8, 9): ws[f"K{rr}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[7].height = 22; ws.row_dimensions[8].height = 28; ws.row_dimensions[9].height = 18
# cabeçalho da tabela
heads = {"B": "Cód.", "C": "PL / serviço", "D": "Execução física\nfim do ano\n(Prevision)", "E": "Execução física\nno corte\n(Prevision)",
         "F": "Desembolso\nno ano\nR$ nominal", "G": "Desembolso\nno ano\nINCC", "H": "Desembolso no ano\nR$ do INCC\nselecionado",
         "I": "% do\nano", "J": "Avanço\nfinanceiro\nfim do ano",
         "K": "Situação em 31/12 (cronograma de atividades do Prevision)", "L": "Anotação do modelo anterior / justificativa"}
for col, h in heads.items():
    c = put(ws, f"{col}{RH}", h, Font(name=F, bold=True, color="FFFFFF", size=9), fill=fl(NAVY_), border=False)
    c.alignment = Alignment(horizontal="left" if col in "BCKL" else "right", vertical="center", wrap_text=True)
ws[f"K{RH}"] = f'="Situação em 31/12/"&D5&" (cronograma de atividades do Prevision)"'
ws.row_dimensions[RH].height = 44
TOTF = f"$F${RT_TOT + 2}"
CRM, CRO, CRF = f"{CR}!$M$5:$M$6004", f"{CR}!$O$5:$O$6004", f"{CR}!$F$5:$F$6004"
from openpyxl.worksheet.formula import ArrayFormula
old_notes = {}
for rr_old, it in [(5 + k, None) for k in range(40)]:
    pass
for k, (rr, item) in enumerate(RET_ROWS):
    bg = fl("F1F5F0" if k % 2 == 0 else "FFFFFF")
    if isinstance(item, int):
        i = item; t = q(TABS[i]); vr = VPROW[i]; br = BIROW[i]; cr = DHROW[i]
        put(ws, f"B{rr}", f"={VP}!C{vr}", f_link, fill=bg, border=False)
        put(ws, f"C{rr}", f"={VP}!D{vr}", f_link, fill=bg, border=False)
        put(ws, f"D{rr}", (f'=IF(OR({VP}!I{vr}<>"DIRETO",\'CP RESUMO\'!$C${cr}=""),"",IFERROR(INDEX(\'CP RESUMO\'!$D${cr}:${L(3 + NM)}${cr},'
                           f"MATCH(DATE($D$5,12,1),'CP RESUMO'!$D$4:${L(3 + NM)}$4,0))/100,\"\"))"), f_calc, "0.0%", bg, border=False)
        if TABS[i] in RET_FIS:   # execução física = avanço financeiro do fim do ano + defasagem (executa antes de pagar)
            put(ws, f"D{rr}", f"=MIN(1,J{rr}+{RET_FIS[TABS[i]]})", f_calc, "0.0%", bg, border=False)
            ws[f"D{rr}"].comment = Comment(f"Premissa do cenário: execução física = avanço financeiro do fim do ano + "
                                           f"{RET_FIS[TABS[i]] * 100:.1f} p.p. (executa antes de pagar). Não vem do Prevision.", "Ferramenta")
        put(ws, f"E{rr}", f'=IF({VP}!I{vr}<>"DIRETO","",{t}!$K$14)', f_link, "0.0%", bg, border=False)
        put(ws, f"F{rr}", f"=SUMPRODUCT((YEAR({t}!$B${R0}:$B${R1})=$D$5)*({t}!$C${R0}:$C${R1}+{t}!$D${R0}:$D${R1}))", f_calc, NUM2, bg, border=False)
        put(ws, f"G{rr}", f"=SUMPRODUCT(({YR}=$D$5)*'MODELO PARA BI'!$H${br}:${BIL}${br})", f_calc, NUM2, bg, border=False)
        put(ws, f"H{rr}", f"=IF(ISNUMBER($G$5),G{rr}*$G$5,0)", f_calc, NUM2, bg, border=False)
        put(ws, f"I{rr}", f"=IF({TOTF}=0,0,F{rr}/{TOTF})", f_calc, "0.0%", bg, border=False)
        put(ws, f"J{rr}", f"=IF({VP}!L{vr}=0,0,SUMPRODUCT(({t}!$B${R0}:$B${R1}<=DATE($D$5,12,1))*({t}!$C${R0}:$C${R1}+{t}!$D${R0}:$D${R1}))/{VP}!L{vr})", f_calc, "0.0%", bg, border=False)
        put(ws, f"M{rr}", f"={VP}!L{vr}", f_link, NUM2, border=False)
        put(ws, f"N{rr}", f'=IF({VP}!I{vr}="DIRETO",{VP}!L{vr},0)', f_link, NUM2, border=False)
        tot = f"COUNTIF({CRM},B{rr})"; con = f'COUNTIFS({CRM},B{rr},{CRO},"CONCLUÍDA")'; ini = f'COUNTIFS({CRM},B{rr},{CRO},"A INICIAR")'
        mn = f"_xlfn.MINIFS({CRF},{CRM},B{rr})"
        obs = (f'=IF({tot}=0,"",IF({con}={tot},"Concluído (todas as {""}"&{tot}&" atividades)",'
               f'IF({ini}={tot},"Não iniciado · início previsto em "&CHOOSE(MONTH({mn}),{MESES_C})&"/"&YEAR({mn}),'
               f'_xlfn.TEXTJOIN(" / ",TRUE,IF(({CR}!$S$5:$S$304=B{rr})*({CR}!$AA$5:$AA$304<>""),{CR}!$AA$5:$AA$304,"")))))')
        obs = obs.replace('(todas as {""}"&', '(todas as "&')
        ws[f"K{rr}"] = RET_SIT[TABS[i]] if TABS[i] in RET_SIT else ArrayFormula(f"K{rr}", obs)
        ws[f"K{rr}"].font = f_calc; ws[f"K{rr}"].fill = bg
        ws[f"K{rr}"].alignment = Alignment(wrap_text=True, vertical="top")
        old_r = 5 + ([x for x, _ in RET_ROWS].index(rr))
        put(ws, f"L{rr}", CFG.get("ret_notes", {}).get(TABS[i]) or ret.get(f"H{old_r}"), Font(name=F, color="0000FF", size=8), fill=bg, border=False)
        ws[f"L{rr}"].alignment = Alignment(wrap_text=True, vertical="top")
    else:
        _, label, code, rows = item
        a, b = rows[0], rows[-1]
        for col in "BCDEFGHIJKL":
            ws[f"{col}{rr}"].fill = fl("C5D3C8"); ws[f"{col}{rr}"].font = Font(name=F, bold=True, color=NAVY_)
        ws[f"B{rr}"] = code; ws[f"C{rr}"] = label
        ws[f"D{rr}"] = f"=IFERROR(SUMPRODUCT(D{a}:D{b},N{a}:N{b})/SUMPRODUCT(ISNUMBER(D{a}:D{b})*N{a}:N{b}),\"\")"
        ws[f"E{rr}"] = f"=IFERROR(SUMPRODUCT(E{a}:E{b},N{a}:N{b})/SUMPRODUCT(ISNUMBER(E{a}:E{b})*N{a}:N{b}),\"\")"
        for col in "FGH": ws[f"{col}{rr}"] = f"=SUM({col}{a}:{col}{b})"
        ws[f"I{rr}"] = f"=IF({TOTF}=0,0,F{rr}/{TOTF})"
        ws[f"J{rr}"] = f"=IFERROR(SUMPRODUCT(J{a}:J{b},M{a}:M{b})/SUM(M{a}:M{b}),0)"
        ws[f"M{rr}"] = f"=SUM(M{a}:M{b})"; ws[f"N{rr}"] = f"=SUM(N{a}:N{b})"
        for col, fm in zip("DEFGHIJ", ("0.0%", "0.0%", NUM2, NUM2, NUM2, "0.0%", "0.0%")): ws[f"{col}{rr}"].number_format = fm
    for col in "BCDEFGHIJKL":
        ws[f"{col}{rr}"].border = Border(bottom=sd_t)
    ws[f"C{rr}"].alignment = Alignment(wrap_text=True, vertical="top")
    for col in "BDEFGHIJ": ws[f"{col}{rr}"].alignment = Alignment(vertical="top", horizontal="left" if col == "B" else "right")
# totais
T1, T2, T3 = RT_TOT, RT_TOT + 1, RT_TOT + 2
sub_rows = [rr for rr, it in RET_ROWS if not isinstance(it, int)]
dir_after = [rr for rr, it in RET_ROWS if isinstance(it, int) and rr > sub_rows[-1]]
i = ND - 1; t = q(TABS[i]); br = BIROW[i]; vr = VPROW[i]
for col in "BCDEFGHIJKL":
    for rr in (T1, T3): ws[f"{col}{rr}"].fill = fl(NAVY_); ws[f"{col}{rr}"].font = Font(name=F, bold=True, color="FFFFFF")
    ws[f"{col}{T2}"].font = f_calc; ws[f"{col}{T2}"].border = Border(bottom=sd_t)
ws[f"C{T1}"] = "DESEMBOLSO DO ANO (sem taxa de administração)"
ws[f"C{T3}"] = "DESEMBOLSO DO ANO COM TAXA DE ADMINISTRAÇÃO"
sumexpr = lambda col: f"={col}{sub_rows[0]}+{col}{sub_rows[1]}+SUM({col}{dir_after[0]}:{col}{dir_after[-1]})"
for col in "FGH": ws[f"{col}{T1}"] = sumexpr(col)
ws[f"B{T2}"] = f"={VP}!C{vr}"; ws[f"C{T2}"] = f"={VP}!D{vr}"
ws[f"F{T2}"] = f"=SUMPRODUCT((YEAR({t}!$B${R0}:$B${R1})=$D$5)*({t}!$C${R0}:$C${R1}+{t}!$D${R0}:$D${R1}))"
ws[f"G{T2}"] = f"=SUMPRODUCT(({YR}=$D$5)*'MODELO PARA BI'!$H${br}:${BIL}${br})"
ws[f"H{T2}"] = f"=IF(ISNUMBER($G$5),G{T2}*$G$5,0)"
ws[f"M{T2}"] = f"={VP}!L{vr}"
for col in "FGH": ws[f"{col}{T3}"] = f"={col}{T1}+{col}{T2}"
for rr in (T1, T2, T3):
    ws[f"I{rr}"] = f"=IF({TOTF}=0,0,F{rr}/{TOTF})"
    for col, fm in zip("FGHI", (NUM2, NUM2, NUM2, "0.0%")): ws[f"{col}{rr}"].number_format = fm
    for col in "FGHI": ws[f"{col}{rr}"].alignment = Alignment(horizontal="right")
    ws.row_dimensions[rr].height = 20
ws.conditional_formatting.add(f"F{TR0}:H{T3}", CellIsRule(operator="lessThan", formula=["-0.5"], font=Font(name=F, bold=True, color="B3392B")))
# barras de dados nas colunas de execução física
from openpyxl.formatting.rule import DataBarRule
ws.conditional_formatting.add(f"D{TR0}:D{TR1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="1F7A52", showValue=True))
ws.conditional_formatting.add(f"I{TR0}:I{TR1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="99BCAB", showValue=True))
put(ws, f"B{T3 + 1}", ("Execução física = % acumulado 'Previsto' do cronograma físico-financeiro do Prevision (aba CP RESUMO) em dezembro do ano; "
                       "no corte = % 'Realizado'. Situação em 31/12 = atividades do cronograma (aba CRONOGRAMA ATIVIDADES) concluídas ou em execução "
                       "nessa data, pelo andar mais alto."), Font(name=F, italic=True, size=8, color=MUT_), border=False)
ws.freeze_panes = f"D{RH + 1}"
# alturas das linhas (Excel não ajusta sozinho texto quebrado): estimadas pelo tamanho do texto da última apuração
import math, os as _os
_lp = _os.path.join(DATA_DIR, "ret_lens.json")
_lens = json.load(open(_lp)) if (_os.path.exists(_lp) and not BLANK) else {}
for rr, _ in RET_ROWS:
    lc, lk, ll = _lens.get(str(rr), [30, 120, 0])
    lines = max(1, math.ceil(lc / 42), math.ceil(lk / 92), math.ceil(ll / 50))
    ws.row_dimensions[rr].height = max(16, 13.2 * lines + 3)
# gráficos
g_rows = [rr for rr, it in RET_ROWS if isinstance(it, int) and it >= 16]
c1 = BarChart(); c1.type = "bar"; c1.title = None; c1.legend = None
put(ws, f"B{T3 + 2}", "EXECUÇÃO FÍSICA PREVISTA NO FIM DO ANO (PREVISION)", Font(name=F, bold=True, size=10, color=NAVY_), border=False)
c1.height = 13; c1.width = 17; c1.gapWidth = 40
c1.add_data(Reference(ws, min_col=4, min_row=g_rows[0], max_row=g_rows[-1])); c1.set_categories(Reference(ws, min_col=3, min_row=g_rows[0], max_row=g_rows[-1]))
c1.series[0].graphicalProperties.solidFill = "1F7A52"; c1.x_axis.scaling.orientation = "maxMin"
c1.y_axis.numFmt = "0%"; c1.y_axis.scaling.min = 0; c1.y_axis.scaling.max = 1; c1.x_axis.delete = False; c1.y_axis.delete = False
ws.add_chart(c1, f"B{T3 + 3}")
c2 = BarChart(); c2.type = "bar"; c2.title = None; c2.legend = None
put(ws, f"H{T3 + 2}", "DESEMBOLSO NO ANO POR PL (R$ MILHÕES)", Font(name=F, bold=True, size=10, color=NAVY_), border=False)
c2.height = 13; c2.width = 17; c2.gapWidth = 40
c2.add_data(Reference(ws, min_col=6, min_row=g_rows[0], max_row=g_rows[-1])); c2.set_categories(Reference(ws, min_col=3, min_row=g_rows[0], max_row=g_rows[-1]))
c2.series[0].graphicalProperties.solidFill = "99BCAB"; c2.x_axis.scaling.orientation = "maxMin"
c2.y_axis.numFmt = '#,##0,,"M"'; c2.x_axis.delete = False; c2.y_axis.delete = False
ws.add_chart(c2, f"H{T3 + 3}")
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.sheet_properties.tabColor = "1F7A52"

# ================================================================== AUX. CURVA
ws = wb.create_sheet("AUX. CURVA")
aux = d2["AUX. CURVA"]
put(ws, "A1", f"={OBRA}", Font(name=F, bold=True, size=13, color="008000"), border=False)
put(ws, "C1", "Progresso mensal — curva inicial × realizado + previsto", f_b, border=False)
hdr(ws, 3, 1, ["Mês", "Curva inicial\n% mensal (informar)", "Curva inicial\n% acumulado", "Realizado\n(INCC)", "Previsto\n(INCC)",
               "Atual\n% mensal", "Atual\n% acumulado", None, "Curva inicial\n(INCC)", "Curva inicial\n(R$)",
               "Atual\n(R$ nominal)", "Atual acumulado\n(R$ nominal)", "Inicial acumulado\n(R$)"], height=45)
TOT_INCC = "'MODELO PARA BI'!$E$63"
for k in range(NM):
    r = 4 + k; col = L(8 + k)
    put(ws, f"A{r}", f"='MODELO PARA BI'!{col}1", f_link, MES)
    put(ws, f"B{r}", None, f_in, "0.00%")
    put(ws, f"C{r}", f"=B{r}" if k == 0 else f"=C{r - 1}+B{r}", f_calc, PCT)
    put(ws, f"D{r}", f"=IF(A{r}<={CORTE},'MODELO PARA BI'!{col}63,0)", f_link, NUM2)
    put(ws, f"E{r}", f"=IF(A{r}>{CORTE},'MODELO PARA BI'!{col}63,0)", f_link, NUM2)
    put(ws, f"F{r}", f"=IF({TOT_INCC}=0,0,(D{r}+E{r})/{TOT_INCC})", f_calc, PCT)
    put(ws, f"G{r}", f"=F{r}" if k == 0 else f"=G{r - 1}+F{r}", f_calc, PCT)
    put(ws, f"I{r}", f"=B{r}*{VP}!$G$50", f_calc, NUM2)
    put(ws, f"J{r}", f"=I{r}*IF(A{r}<={CORTE},IFERROR(INDEX({INCC_B},MATCH(A{r},{INCC_A},0)),{IDX}),{IDX})", f_calc, NUM)
    put(ws, f"K{r}", f"='MODELO PARA BI'!{col}65", f_link, NUM)
    put(ws, f"L{r}", f"='MODELO PARA BI'!{col}66", f_link, NUM)
    put(ws, f"M{r}", f"=J{r}" if k == 0 else f"=M{r - 1}+J{r}", f_calc, NUM)
notes = [v for k, v in sorted(aux.items(), key=lambda kv: int(re.sub(r"\D", "", kv[0]))) if k.startswith("AA")]
put(ws, "O3", "Anotações (do arquivo original)", f_h, fill=fill_h)
for k, v in enumerate(notes): put(ws, f"O{4 + k}", v, f_in, border=False)
ws.conditional_formatting.add(f"A4:M{3 + NM}", FormulaRule(formula=[f"$A4<={CORTE}"], fill=fill_past))
for col, w in zip("ABCDEFGHIJKLMNO", (9, 14, 12, 12, 12, 10, 12, 2, 12, 14, 14, 16, 16, 2, 60)): ws.column_dimensions[col].width = w
ws.freeze_panes = "B4"
lc = LineChart(); lc.title = "Curva S — inicial × atual (% acumulado)"; lc.height = 9; lc.width = 22
lc.add_data(Reference(ws, min_col=3, min_row=3, max_row=3 + NM), titles_from_data=True)
lc.add_data(Reference(ws, min_col=7, min_row=3, max_row=3 + NM), titles_from_data=True)
lc.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + NM)); lc.x_axis.number_format = MES
ws.add_chart(lc, "O12")

# ================================================================== "INCC-FGV"
ws = wb.create_sheet('"INCC-FGV"')
hdr(ws, 1, 1, ["Mês", "Índice", "(%) mês", "(%) no ano", "(%) 12 meses"])
for k in range(799):
    r = 2 + k
    if k < len(d1["incc"]):
        dd, v = d1["incc"][k]
        put(ws, f"A{r}", dt.date.fromisoformat(dd), f_in, "mm/yyyy"); put(ws, f"B{r}", v, f_in, "#,##0.00")
    else:
        ws[f"A{r}"].number_format = "mm/yyyy"; ws[f"B{r}"].number_format = "#,##0.00"; ws[f"A{r}"].font = f_in; ws[f"B{r}"].font = f_in
    if r > 2:
        put(ws, f"C{r}", f'=IF(OR(B{r}="",B{r - 1}=""),"",B{r}/B{r - 1}-1)', f_calc, "0.00%", border=False)
    put(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(B{r}/INDEX($B$2:$B$800,MATCH(DATE(YEAR(A{r})-1,12,1),$A$2:$A$800,0))-1,""))', f_calc, "0.00%", border=False)
    put(ws, f"E{r}", f'=IF(B{r}="","",IFERROR(B{r}/INDEX($B$2:$B$800,MATCH(DATE(YEAR(A{r})-1,MONTH(A{r}),1),$A$2:$A$800,0))-1,""))', f_calc, "0.00%", border=False)
put(ws, "G2", "Fonte: FGV (INCC-DI). Acrescente uma linha por mês.", f_note, border=False)
for col, w in zip("ABCDE", (10, 12, 10, 11, 12)): ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

# ================================================================== CURVA PREVISION (cole o export)
ws = wb.create_sheet("CURVA PREVISION")
copy_values(ws, d2["CURVA PREVISION"])
ws.sheet_properties.tabColor = "FFC000"

# ================================================================== CRONOGRAMA ATIVIDADES (cole o export de atividades do Prevision)
ws = wb.create_sheet("CRONOGRAMA ATIVIDADES")
for k, v in d2["CRONO"].items():
    c = ws[k]
    if isinstance(v, str) and re.match(r"\d{4}-\d\d-\d\dT", v): v = dt.datetime.fromisoformat(v)
    c.value = v; c.font = f_in
    if isinstance(v, dt.datetime): c.number_format = "dd/mm/yyyy"
NA0, NA1 = 5, 6004   # até 6.000 atividades
MAP0, MAP1 = 5, 304
DATA = DATA_RT
put(ws, "M1", "CÁLCULO (não editar) — situação de cada atividade na data do retrato", f_note, border=False)
put(ws, "M2", "Pacotes sem de-para →", f_calc, border=False)
put(ws, "O2", f'=COUNTIF(M{NA0}:M{NA1},"?")', f_calc, "0")
ws.conditional_formatting.add("O2", CellIsRule(operator="greaterThan", formula=["0"], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
put(ws, "M3", '="Data do retrato: "&TEXT(DAY(' + DATA + '),"00")&"/"&TEXT(MONTH(' + DATA + '),"00")&"/"&YEAR(' + DATA + ')', f_link, border=False)
hdr(ws, 4, 13, ["PL", "Nº pav.", "Situação na data"])
for r in range(NA0, NA1 + 1):
    ws[f"M{r}"] = f'=IF(C{r}="","",IFERROR(INDEX($S${MAP0}:$S${MAP1},MATCH(C{r},$R${MAP0}:$R${MAP1},0))&"","?"))'
    ws[f"N{r}"] = (f'=IF(E{r}="","",IF(ISNUMBER(SEARCH("SUBSOLO",E{r})),IFERROR(-VALUE(LEFT(E{r},FIND("º",SUBSTITUTE(E{r},"°","º"))-1)),-1),'
                   f'IF(LEFT(E{r},6)="TÉRREO",0,IFERROR(VALUE(LEFT(E{r},FIND("º",SUBSTITUTE(E{r},"°","º"))-1)),""))))')
    ws[f"O{r}"] = f'=IF(OR(F{r}="",G{r}=""),"",IF(G{r}<={DATA},"CONCLUÍDA",IF(F{r}<={DATA},"EM ANDAMENTO","A INICIAR")))'
    for col in "MNO": ws[f"{col}{r}"].font = f_calc
ws.conditional_formatting.add(f"M{NA0}:M{NA1}", CellIsRule(operator="equal", formula=['"?"'], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
hdr(ws, 4, 18, ["Pacote de trabalho (Prevision)", "PL", "Rótulo nas observações", "Exibir?", "Nº ativ.", "Concluídas",
                "Em execução", "Maior pav.\nconcluído", "Maior pav.\nem execução", "Texto (observação)"], height=32)
put(ws, "R3", "DE-PARA: PACOTE DO PREVISION → PL (edite as colunas azuis; PL vazio = pacote ignorado)", f_b, border=False)
lbl = lambda x: f'IF({x}<0,IF(COUNTIF($E${NA0}:$E${NA1},"SUBSOLO*")>0,"subsolo",-{x}&"º subsolo"),IF({x}=0,"térreo",{x}&"º pav."))'
C_, O_, N_ = f"$C${NA0}:$C${NA1}", f"$O${NA0}:$O${NA1}", f"$N${NA0}:$N${NA1}"
for k in range(MAP1 - MAP0 + 1):
    r = MAP0 + k
    if k < len(d2["PKG"]):
        p, code, labl, n, show = d2["PKG"][k]
        ws[f"R{r}"] = p; ws[f"S{r}"] = code or None; ws[f"T{r}"] = labl; ws[f"U{r}"] = show
    for col in "RSTU": ws[f"{col}{r}"].font = f_in
    ws[f"V{r}"] = f'=IF(R{r}="","",COUNTIF({C_},R{r}))'
    ws[f"W{r}"] = f'=IF(R{r}="","",COUNTIFS({C_},R{r},{O_},"CONCLUÍDA"))'
    ws[f"X{r}"] = f'=IF(R{r}="","",COUNTIFS({C_},R{r},{O_},"EM ANDAMENTO"))'
    ws[f"Y{r}"] = f'=IF(R{r}="","",IF(COUNTIFS({C_},R{r},{O_},"CONCLUÍDA",{N_},">=-99")=0,"",_xlfn.MAXIFS({N_},{C_},R{r},{O_},"CONCLUÍDA")))'
    ws[f"Z{r}"] = f'=IF(R{r}="","",IF(COUNTIFS({C_},R{r},{O_},"EM ANDAMENTO",{N_},">=-99")=0,"",_xlfn.MAXIFS({N_},{C_},R{r},{O_},"EM ANDAMENTO")))'
    ws[f"AA{r}"] = (f'=IF(OR(R{r}="",S{r}="",U{r}<>"SIM",V{r}=0),"",IF(W{r}+X{r}=0,"",T{r}&": "&'
                    f'IF(W{r}=V{r},"concluído",IF(W{r}=0,"em execução"&IF(Z{r}<>""," ("&{lbl(f"Z{r}")}&")",""),'
                    f'IF(Y{r}<>"","até "&{lbl(f"Y{r}")},W{r}&" de "&V{r}&" concluídas")&IF(Z{r}="","",IF(OR(Y{r}="",Z{r}>Y{r}),", executando "&{lbl(f"Z{r}")}," (restam serviços no "&{lbl(f"Z{r}")}&")"))))))')
    for col in ("V", "W", "X", "Y", "Z", "AA"): ws[f"{col}{r}"].font = f_calc
from openpyxl.worksheet.datavalidation import DataValidation
dv_ = DataValidation(type="list", formula1='"SIM,NÃO"', allow_blank=True); ws.add_data_validation(dv_); dv_.add(f"U{MAP0}:U{MAP1}")
for col, w in zip(["M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"],
                  (6, 7, 15, 2, 2, 44, 6, 30, 8, 8, 10, 10, 10, 10, 60)):
    ws.column_dimensions[col].width = w
ws.column_dimensions["C"].width = 40; ws.column_dimensions["E"].width = 28
ws.freeze_panes = "A5"
ws.sheet_properties.tabColor = "FFC000"

# ================================================================== CP RESUMO
ws = wb.create_sheet("CP RESUMO")
put(ws, "A1", "CRONOGRAMA DO PREVISION — % ACUMULADO PREVISTO POR PL (coluna 'Previsto' de cada mês)", f_b, border=False)
put(ws, "A2", "1º mês do Prevision →", f_calc, border=False)
put(ws, "B2", ("=IFERROR(IF(ISNUMBER('CURVA PREVISION'!J2),DATE(YEAR('CURVA PREVISION'!J2),MONTH('CURVA PREVISION'!J2),1),"
               "DATE(VALUE(RIGHT('CURVA PREVISION'!J2,4)),VALUE(MID('CURVA PREVISION'!J2,4,2)),1)),\"\")"), f_link, "mm/yyyy")
put(ws, "C2", "nº de meses →", f_calc, border=False)
put(ws, "D2", "=COUNTIF('CURVA PREVISION'!$I$2:$KZ$2,\"??/??/????\")+COUNT('CURVA PREVISION'!$I$2:$KZ$2)", f_link, "0")
hdr(ws, 4, 1, ["PL", "DESCRIÇÃO", "LINHA NO\nPREVISION"], height=32)
put(ws, "C3", "índice do mês →", f_note, border=False)
for k in range(NM):
    col = L(4 + k)
    c = put(ws, f"{col}4", f"='MODELO PARA BI'!{L(8 + k)}1", f_h, MES, fill_h)
    put(ws, f"{col}3", f'=IF($B$2="","",(YEAR({col}4)-YEAR($B$2))*12+MONTH({col}4)-MONTH($B$2))', f_note, "0", border=False)
    ws.column_dimensions[col].width = 8
CPL = L(4 + NM - 1)
for i in range(ND):
    r = DHROW[i]; vr = VPROW[i]
    put(ws, f"A{r}", f'=IF({VP}!R{vr}<>"",{VP}!R{vr},IF({VP}!Q{vr}="",{VP}!C{vr},{VP}!Q{vr}))', f_link); put(ws, f"B{r}", f"={VP}!D{vr}", f_link)
    put(ws, f"C{r}", f"=IFERROR(MATCH(A{r},'CURVA PREVISION'!$B$1:$B$3000,0),\"\")", f_link, "0")
    for k in range(NM):
        col = L(4 + k)
        put(ws, f"{col}{r}", (f"=IF(OR($C{r}=\"\",$B$2=\"\",$D$2=0),0,IF({col}$3<0,0,"
                              f"N(INDEX('CURVA PREVISION'!$A$1:$KZ$3000,$C{r},10+3*MIN({col}$3,$D$2-1)))))"), f_calc, "0.0")
RB = 42  # deslocamento do bloco REALIZADO
put(ws, "A45", "% ACUMULADO REALIZADO POR PL (coluna 'Realizado' de cada mês)", f_b, border=False)
hdr(ws, 46, 1, ["PL", "DESCRIÇÃO", "LINHA NO\nPREVISION"], height=32)
for k in range(NM):
    col = L(4 + k)
    put(ws, f"{col}46", f"={col}4", f_h, MES, fill_h)
for i in range(ND):
    r = DHROW[i] + RB
    vr = VPROW[i]
    put(ws, f"A{r}", f'=IF({VP}!Q{vr}="",{VP}!C{vr},{VP}!Q{vr})', f_link); put(ws, f"B{r}", f"=B{r - RB}", f_calc)
    put(ws, f"C{r}", f"=IFERROR(MATCH(A{r},'CURVA PREVISION'!$B$1:$B$3000,0),\"\")", f_link, "0")
    for k in range(NM):
        col = L(4 + k)
        put(ws, f"{col}{r}", (f"=IF(OR($C{r}=\"\",$B$2=\"\",$D$2=0),0,IF({col}$3<0,0,"
                              f"N(INDEX('CURVA PREVISION'!$A$1:$KZ$3000,$C{r},11+3*MIN({col}$3,$D$2-1)))))"), f_calc, "0.0")
ws.conditional_formatting.add(f"D{5 + RB}:{CPL}{4 + ND + RB}", FormulaRule(formula=[f"D$4>{CORTE}"], font=Font(name=F, color="BFBFBF")))
ws.column_dimensions["A"].width = 6; ws.column_dimensions["B"].width = 40; ws.column_dimensions["C"].width = 10
ws.freeze_panes = "D5"

# ================================================================== DADOS HISTORICOS
ws = wb.create_sheet("DADOS HISTORICOS")
H0 = CI("CB")                       # 1ª coluna do bloco auxiliar
H1 = H0 + NDB_COLS - 1
HA, HB = L(H0), L(H1)
put(ws, "A1", "INCORRIDO POR DISCIPLINA E MÊS (R$) — lido da aba DADOS BRUTOS pelo NOME do grupo e pela DATA do mês", f_b, border=False)
put(ws, "A2", "Grupos do ERP sem correspondência →", f_calc, border=False)
put(ws, "C2", (f"=SUMPRODUCT(('DADOS BRUTOS'!$A$3:$A${NDB_ROWS}<>\"\")*(COUNTIF($B$5:$C$42,'DADOS BRUTOS'!$A$3:$A${NDB_ROWS})=0)"
               f"*(COUNTIF($A$50:$A$69,'DADOS BRUTOS'!$A$3:$A${NDB_ROWS})=0))"), f_calc, "0")
ws.conditional_formatting.add("C2", CellIsRule(operator="greaterThan", formula=["0"], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
hdr(ws, 4, 1, ["ITEM", "NOME NO ERP (1)", "NOME NO ERP (2)"])
for k in range(NM):
    col = L(4 + k)
    put(ws, f"{col}4", f"='MODELO PARA BI'!{L(8 + k)}1", f_h, MES, fill_h)
    ws.column_dimensions[col].width = 11
put(ws, "D3", "Acu até →", f_note, border=False)
put(ws, "BX4", "FORA DO\nHORIZONTE", f_h, fill=fill_h); put(ws, "BY4", "TOTAL NO ERP", f_h, fill=fill_h)
put(ws, "BZ4", "INCC DO\nACUMULADO\n(mês a mês)", f_h, fill=fill_h)
ws["BZ4"].comment = Comment("Realizado até o 1º mês do horizonte (coluna D, 'Acu até'), convertido em INCC pelo índice de CADA mês "
                            "em que foi pago (aba INCC-FGV). É o valor usado na linha 'Acu até' das abas de disciplina.", "Ferramenta")
ws["BX4"].alignment = WRAP; ws["BY4"].alignment = WRAP; ws["BZ4"].alignment = WRAP
put(ws, f"{HA}2", "CÁLCULO AUXILIAR — soma por coluna da aba DADOS BRUTOS (não editar)", f_note, border=False)
for j in range(NDB_COLS):
    col = L(H0 + j); src = L(2 + j)
    put(ws, f"{col}3", f"=\"{src}\"", f_note, border=False)
    put(ws, f"{col}4", f"=IF(ISNUMBER('DADOS BRUTOS'!{src}$2),DATE(YEAR('DADOS BRUTOS'!{src}$2),MONTH('DADOS BRUTOS'!{src}$2),1),\"\")", f_link, MES)
    put(ws, f"{col}1", f'=IF({col}$4="",0,IFERROR(1/INDEX({INCC_B},MATCH({col}$4,{INCC_A},0)),0))', f_calc, "0.000000", border=False)
    put(ws, f"{col}44", f"=SUM({col}5:{col}42)", f_calc, NUM, border=False)
tab2erp = erp
for i, t in enumerate(TABS):
    r = DHROW[i]
    put(ws, f"A{r}", f"={VP}!D{VPROW[i]}", f_link)
    put(ws, f"B{r}", tab2erp[t], f_in); put(ws, f"C{r}", ERP2.get(t), f_in)
    for j in range(NDB_COLS):
        col = L(H0 + j); src = L(2 + j)
        rng = f"'DADOS BRUTOS'!{src}$3:{src}${NDB_ROWS}"
        lab_ = f"'DADOS BRUTOS'!$A$3:$A${NDB_ROWS}"
        put(ws, f"{col}{r}", (f"=IF(OR({col}$4=\"\",$B{r}=\"\"),0,SUMIF({lab_},$B{r},{rng})"
                              f"+IF($C{r}=\"\",0,SUMIF({lab_},$C{r},{rng})))"), f_calc, NUM, border=False)
    for k in range(NM):
        col = L(4 + k)
        crit = f"\"<=\"&{col}$4" if k == 0 else f"{col}$4"
        put(ws, f"{col}{r}", f"=SUMIF(${HA}$4:${HB}$4,{crit},${HA}{r}:${HB}{r})", f_calc, NUM)
    put(ws, f"BX{r}", f"=SUMIF(${HA}$4:${HB}$4,\">\"&{FIM},${HA}{r}:${HB}{r})", f_calc, NUM)
    put(ws, f"BY{r}", f"=SUM(${HA}{r}:${HB}{r})", f_calc, NUM)
    put(ws, f"BZ{r}", f'=SUMPRODUCT((${HA}$4:${HB}$4<>"")*(${HA}$4:${HB}$4<=$D$4),${HA}{r}:${HB}{r},${HA}$1:${HB}$1)', f_calc, NUM2)
put(ws, "A44", "TOTAL", f_b); put(ws, "A45", "TOTAL ACUMULADO", f_b)
for k in list(range(NM)) + ["BX", "BY", "BZ"]:
    col = L(4 + k) if isinstance(k, int) else k
    put(ws, f"{col}44", f"=SUM({col}5:{col}42)", f_b, NUM, fill_tot)
for k in range(NM):
    col = L(4 + k)
    put(ws, f"{col}45", f"={col}44" if k == 0 else f"={L(3 + k)}45+{col}44", f_b, NUM)
put(ws, "A49", "NOMES DO ERP A IGNORAR (títulos, subtotais, grupos fora do orçamento)", f_h, fill=fill_h)
ws.merge_cells("A49:C49")
IGN = CFG.get("ignore", ["APORTES", "CUSTOS COM OBRAS", "MATERIAIS E SERVIÇOS PRESTADOS", "Total Geral", "REEMBOLSO MATERIAIS E SERVIÇOS"])
for k, v in enumerate(IGN):
    put(ws, f"A{50 + k}", v, f_in)
for k in range(len(IGN), 20): put(ws, f"A{50 + k}", None, f_in)
ws.conditional_formatting.add(f"D5:BW42", FormulaRule(formula=[f"D$4<={CORTE}"], fill=fill_past))
ws.column_dimensions["A"].width = 44; ws.column_dimensions["B"].width = 44; ws.column_dimensions["C"].width = 24
for c in range(H0, H1 + 1): ws.column_dimensions[L(c)].outlineLevel = 1; ws.column_dimensions[L(c)].hidden = True
ws.column_dimensions["BZ"].width = 14
ws.freeze_panes = "D5"

# ================================================================== DADOS BRUTOS
ws = wb.create_sheet("DADOS BRUTOS")
for k, v in d2["DADOS BRUTOS"].items():
    c = ws[k]; c.value = v; c.font = f_in
    if re.match(r"[A-Z]+2$", k) and isinstance(v, float) and k != "A2":
        c.number_format = "mmm/yy"
    elif isinstance(v, float):
        c.number_format = NUM2
ws.conditional_formatting.add(f"A3:A{NDB_ROWS}", FormulaRule(
    formula=[f"AND($A3<>\"\",COUNTIF('DADOS HISTORICOS'!$B$5:$C$42,$A3)=0,COUNTIF('DADOS HISTORICOS'!$A$50:$A$69,$A3)=0)"],
    fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
ws.column_dimensions["A"].width = 50
ws.freeze_panes = "B3"
ws.sheet_properties.tabColor = "FFC000"

# ================================================================== PREMISSAS (oculta, cópia)
ws = wb.create_sheet("PREMISSAS"); copy_values(ws, d2["PREMISSAS"]); ws.sheet_state = "hidden"
ws.column_dimensions["B"].width = 90

# ================================================================== abas de disciplina
cur = d1["cur"]


SPLIT = "MDO PREVISION + MATERIAL PREMISSA"
FORCE = {"ELEVADORES ": "PREVISÃO INFORMADA", "ESQUADRIAS": "PREVISÃO INFORMADA", "LOUÇAS": SPLIT}
FORCE.update(CFG.get("force", {}))                 # fontes forçadas específicas da obra
AUXE = CFG.get("aux_estrutura")                     # aba AUX.ESTRUTURA resgatada da origem → curva da ESTRUTURA (gerador/aux_estrutura.py)
if AUXE: FORCE["ESTRUTURA"] = "PREVISÃO INFORMADA"
OBS = {"ELEVADORES ": "Previsão informada mantida (decisão do usuário). Premissa: pedido 1 ano antes, 12 pagamentos.",
       "ESQUADRIAS": "Previsão informada mantida (decisão do usuário). Premissa: 60% em 10 pagamentos + 40% por medição direta.",
       "LOUÇAS": "Saldo dividido: MDO segue a curva do Prevision; material segue a premissa (pedido antecipado, pagamento 28/56/84 dias)."}
if TAXA_SO_INC:
    OBS["TAX. ADM"] = "Somente o incorrido no ERP: a meta em INCC (VALOR POR PL!H48) é o próprio incorrido em INCC, sem previsão futura."
for _t, _n in CFG.get("notes", {}).items(): OBS[_t] = (OBS[_t] + " " if _t in OBS else "") + _n


def disc_sheet(i, t):
    split = FORCE.get(t) == SPLIT
    ws = wb.create_sheet(t)
    vr, dr = VPROW[i], DHROW[i]
    put(ws, "B2", f"={VP}!D{vr}", Font(name=F, bold=True, size=13, color="008000"), border=False)
    put(ws, "J2", "INCC CORTE →", f_b, border=False); put(ws, "K2", f"={IDX}", f_link, "#,##0.00")
    hdr(ws, 4, 2, ["ITEM", "INCORRIDO\n(R$)", "PREVISÃO\n(R$) — calculada", "INCC", "CURVA MÊS", "CURVA\nACUMULADA",
                   "PREVISÃO INFORMADA\n(R$ ou peso) ← digite", "ORIGEM", "VALOR R$", "INCC", "INCORRIDO\n(INCC do mês)",
                   "PREVISION\n% PREVISTO\nACUMULADO", "PREVISION\n% PREVISTO\nNO MÊS", "PREVISION\n% REALIZADO\nACUMULADO", "PESO USADO\n(curva da\nprevisão)"], height=58)
    ws["H4"].fill = PatternFill("solid", fgColor="8A6D3B")
    dsc = orc[i]["disc"]
    series = cur.get(dsc, {})
    for k in range(NM):
        r = R0 + k; cp = L(4 + k); cpp = L(3 + k)
        if k == 0:
            put(ws, f"A{r}", "Acu até →", f_note, border=False)
            put(ws, f"B{r}", f"={INICIO}", f_link, MES)
        else:
            put(ws, f"B{r}", f"=DATE(YEAR(B{r - 1}),MONTH(B{r - 1})+1,1)", f_calc, MES)
        put(ws, f"C{r}", f"=IF(AND($K$22=\"NÃO\",B{r}>{CORTE}),0,'DADOS HISTORICOS'!{cp}{dr})", f_link, NUM2)
        if split:
            _dform = (f'=IF(B{r}<={CORTE},0,IF($K$16="{SPLIT}",R{r}+U{r},'
                      f'IF($K$17>0,$K$18*P{r}/$K$17,0)))')
        else:
            _dform = f"=IF(B{r}<={CORTE},0,IF($K$17>0,$K$18*P{r}/$K$17,0))"
        if DESLOC:
            _de, _pa = DESLOC["de"], DESLOC["para"]
            put(ws, f"W{r}", _dform, f_calc, NUM2)
            _back = f"W{r - 12}" if k >= 12 else "0"
            put(ws, f"D{r}", (f"=IF(YEAR(B{r})={_de},W{r}*{VP}!$K$9,IF(YEAR(B{r})={_pa},"
                              f"W{r}+(1-{VP}!$K$9)*IF($W$3>0,$W$2*W{r}/$W$3,{_back}),W{r}))"), f_calc, NUM2)
        else:
            put(ws, f"D{r}", _dform, f_calc, NUM2)
        # L = incorrido em INCC: cada mês pelo seu índice (o acumulado anterior ao horizonte vem de DADOS HISTORICOS!BZ);
        #     lançamentos do ERP depois do corte pelo INCC do corte
        if k == 0:
            put(ws, f"L{r}", f"='DADOS HISTORICOS'!$BZ${dr}", f_link, NUM2)
        else:
            put(ws, f"L{r}", f"=C{r}/IF(B{r}<={CORTE},IFERROR(INDEX({INCC_B},MATCH(B{r},{INCC_A},0)),$K$2),$K$2)", f_calc, NUM2)
        # E = INCC do mês: incorrido em INCC + previsão (saldo em INCC convertido em R$ pelo INCC do corte)
        put(ws, f"E{r}", f"=L{r}+D{r}/$K$2", f_calc, NUM2)
        put(ws, f"F{r}", f"=IF($K$9=0,0,E{r}/$K$9)", f_calc, PCT)
        put(ws, f"G{r}", f"=F{r}" if k == 0 else f"=G{r - 1}+F{r}", f_calc, PCT)
        v = series.get(month_iso(k))
        put(ws, f"H{r}", v, f_in, NUM2, fill_in)
        put(ws, f"M{r}", f"='CP RESUMO'!{cp}{dr}", f_link, "0.00")
        put(ws, f"N{r}", 0 if k == 0 else f"=MAX(0,M{r}-M{r - 1})", f_calc, "0.00")
        put(ws, f"O{r}", f"=IF(B{r}<={CORTE},'CP RESUMO'!{cp}{dr + 42},\"\")", f_link, "0.00")
        put(ws, f"P{r}", (f'=IF(B{r}<={CORTE},0,IF(OR($K$16="CURVA PREVISION",$K$16="{SPLIT}"),N{r},'
                          f'IF($K$16="PREVISÃO INFORMADA",MAX(0,H{r}),0)))'), f_calc, "#,##0.00")
        if split:
            idx = k + 1
            def sh(kc):  # peso da MDO no mês de instalação correspondente à parcela
                return (f"IF(AND({idx}+{kc}>=1,{idx}+{kc}<={NM}),INDEX($N${R0}:$N${R1},{idx}+{kc}),0)")
            put(ws, f"R{r}", f'=IF(OR(B{r}<={CORTE},$K$12=0),0,$K$18*$K$26*N{r}/$K$12)', f_calc, NUM2)
            put(ws, f"S{r}", f"=({sh('$K$33')}+{sh('$K$34')}+{sh('$K$35')})/3", f_calc, "0.00")
            first = f"DATE(YEAR({CORTE}),MONTH({CORTE})+1,1)"
            put(ws, f"T{r}", f"=IF(B{r}<={CORTE},0,S{r}+IF(B{r}={first},SUMPRODUCT((B{R0}:B{R1}<={CORTE})*S{R0}:S{R1}),0))", f_calc, "0.00")
            put(ws, f"U{r}", f"=IF($K$32=0,0,$K$18*$K$27*T{r}/$K$32)", f_calc, NUM2)
    if DESLOC:
        hdr(ws, 4, 23, ["PREVISÃO (R$)\nantes do\ndeslocamento"], height=58)
        put(ws, "V2", f"Σ {DESLOC['de']} →", f_b, border=False)
        put(ws, "W2", f"=SUMPRODUCT((YEAR(B{R0}:B{R1})={DESLOC['de']})*W{R0}:W{R1})", f_calc, NUM2)
        put(ws, "V3", f"Σ {DESLOC['para']} →", f_b, border=False)
        put(ws, "W3", f"=SUMPRODUCT((YEAR(B{R0}:B{R1})={DESLOC['para']})*W{R0}:W{R1})", f_calc, NUM2)
        ws.column_dimensions["W"].width = 16
    blk = [(5, "VALOR ORÇADO (R$ base / INCC)", f"={VP}!F{vr}", f"={VP}!G{vr}"),
           (6, "META DA DISTRIBUIÇÃO (INCC)", None, f"={VP}!H{vr}"),
           (7, "PROJEÇÃO (R$) = incorrido + saldo INCC × INCC do corte", f"=C{RTOT}+D{RTOT}", None),
           (8, "INCORRIDO + COMPROMISSOS DO ERP (INCC)", None, f"=SUM(L{R0}:L{R1})"),
           (10, "SALDO A DISTRIBUIR (INCC)" + (" − conciliação" if CONC else ""), None, f"=IF(K6-K8>0,(K6-K8)*{VP}!$G$65,K6-K8)" if CONC else "=K6-K8"),
           (9, "TOTAL NO FLUXO", f"=C{RTOT}+D{RTOT}", f"=E{RTOT}"),
           (11, "Σ PREVISÃO INFORMADA (meses futuros)", None, f"=SUMPRODUCT((B{R0}:B{R1}>{CORTE})*H{R0}:H{R1})"),
           (12, "Σ CURVA PREVISION (meses futuros, %)", None, f"=SUMPRODUCT((B{R0}:B{R1}>{CORTE})*N{R0}:N{R1})"),
           (13, "TIPO DO ITEM", None, f"={VP}!I{vr}"),
           (14, "AVANÇO FÍSICO REALIZADO NO CORTE (Prevision)", None,
            f"=IFERROR(INDEX('CP RESUMO'!$D${dr + 42}:${L(3 + NM)}${dr + 42},MATCH({CORTE},'CP RESUMO'!$D$4:${L(3 + NM)}$4,0))/100,0)"),
           (15, "FORÇAR FONTE (opcional)", None, FORCE.get(t, "AUTOMÁTICO")),
           (16, "FONTE DA CURVA USADA", None,
            ((f'=IF(K15="{SPLIT}",IF(AND(K12>0,K32>0),"{SPLIT}",IF(K11>0,"PREVISÃO INFORMADA","SEM CURVA")),' if split else '=') +
             'IF(K15="CURVA PREVISION",IF(K12>0,"CURVA PREVISION","SEM CURVA"),'
             'IF(K15="PREVISÃO INFORMADA",IF(K11>0,"PREVISÃO INFORMADA","SEM CURVA"),'
             f'IF(AND(K13="DIRETO",K14<IF({VP}!$K$8="PADRÃO",0,{VP}!$K$7),K12>0),"CURVA PREVISION",'
             'IF(K11>0,"PREVISÃO INFORMADA",IF(K12>0,"CURVA PREVISION","SEM CURVA")))))' + (")" if split else ""))),
           (17, "Σ PESOS USADOS", None, f"=SUM(P{R0}:P{R1})"),
           (18, "SALDO A DISTRIBUIR (R$) = saldo INCC × INCC do corte", None, "=MAX(0,K10)*$K$2"),
           (19, "AJUSTE APLICADO À PREVISÃO INFORMADA", None, '=IF(AND(K16="PREVISÃO INFORMADA",K11>0),K18/K11-1,"")'),
           (20, "MOTIVO", None,
            ('=IF(K15<>"AUTOMÁTICO","Fonte escolhida manualmente em K15 (ver observações)",'
             f'IF(K16="CURVA PREVISION",IF(AND(K13="DIRETO",K14<IF({VP}!$K$8="PADRÃO",0,{VP}!$K$7)),"Item direto com avanço físico abaixo de "&TEXT({VP}!$K$7*100,"0")&"%","Sem previsão informada na coluna H"),'
             f'IF(K16="PREVISÃO INFORMADA",IF(K13="INDIRETO","Item indireto: usa a previsão informada",IF({VP}!$K$8="PADRÃO","Cenário padrão: usa a curva digitada","Avanço físico acima do limite: usa a previsão informada")),'
             '"Não há curva: preencha a coluna H ou revise o Prevision")))'))]
    for r, lab, j, k_ in blk:
        put(ws, f"I{r}", lab, f_b)
        if j: put(ws, f"J{r}", j, f_link if "!" in j else f_calc, NUM2)
        if k_:
            fm = "0.0%" if r in (14, 19) else ("0.00" if r == 12 else NUM2)
            if r in (6, 8, 10): fm = "#,##0.00"
            put(ws, f"K{r}", k_, f_in if r == 15 else (f_link if "!" in k_ else f_calc), fm, fill_y if r == 15 else None)
    ws["K15"].comment = Comment("AUTOMÁTICO aplica a regra (item direto com avanço físico abaixo do limite de VALOR POR PL!K7 usa o Prevision). "
                                "Escolha CURVA PREVISION ou PREVISÃO INFORMADA para forçar a fonte nesta aba.", "Ferramenta")
    from openpyxl.worksheet.datavalidation import DataValidation
    dvf = DataValidation(type="list", formula1=('"AUTOMÁTICO,CURVA PREVISION,PREVISÃO INFORMADA' + (f',{SPLIT}' if split else '') + '"'), allow_blank=False)
    ws.add_data_validation(dvf); dvf.add("K15")
    ws["K20"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("K20:K21"); ws.row_dimensions[20].height = 30
    put(ws, "I22", "CONSIDERAR LANÇAMENTOS DO ERP APÓS O CORTE?", f_b)
    put(ws, "K22", "NÃO" if t in ERP_FUT_NAO else "SIM", f_in, fill=fill_y)
    ws["K22"].comment = Comment("SIM: valores do ERP com data depois do corte entram no fluxo como estão (compromissos) e só o restante é distribuído. "
                                "NÃO: vale só o realizado até o corte; todo o saldo até a projeção é distribuído pela curva.", "Ferramenta")
    dve = DataValidation(type="list", formula1='"SIM,NÃO"', allow_blank=False); ws.add_data_validation(dve); dve.add("K22")
    put(ws, "I23", "Critério / observações:", f_b, border=False)
    put(ws, "I24", OBS.get(t), f_in, border=False)
    ws["I24"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("I24:K24"); ws.row_dimensions[24].height = 45
    if split:
        put(ws, "I25", "DIVISÃO MDO × MATERIAL", f_h, fill=fill_h); ws.merge_cells("I25:K25")
        rows_ = [(26, "% MÃO DE OBRA (segue a curva do Prevision)", 0.40, "0%", True),
                 (27, "% MATERIAL (segue a premissa)", "=1-K26", "0%", False),
                 (28, "ANTECEDÊNCIA DO PEDIDO (dias antes da instalação)", 180, "0", True),
                 (29, "1ª PARCELA (dias após o pedido)", 28, "0", True),
                 (30, "2ª PARCELA (dias após o pedido)", 56, "0", True),
                 (31, "3ª PARCELA (dias após o pedido)", 84, "0", True),
                 (32, "Σ PESOS DO MATERIAL (meses futuros)", f"=SUM(T{R0}:T{R1})", "0.00", False),
                 (33, "Parcela 1 paga X meses antes da instalação", "=ROUND((K28-K29)/30,0)", "0", False),
                 (34, "Parcela 2 paga X meses antes da instalação", "=ROUND((K28-K30)/30,0)", "0", False),
                 (35, "Parcela 3 paga X meses antes da instalação", "=ROUND((K28-K31)/30,0)", "0", False),
                 (36, "MDO no fluxo (R$)", f"=SUM(R{R0}:R{R1})", NUM2, False),
                 (37, "MATERIAL no fluxo (R$)", f"=SUM(U{R0}:U{R1})", NUM2, False)]
        for r_, lab_, v_, fm_, inp in rows_:
            put(ws, f"I{r_}", lab_, f_b)
            put(ws, f"K{r_}", v_, f_in if inp else f_calc, fm_, fill_y if inp else None)
        ws["K28"].comment = Comment("Premissa: 'Louças e metais: 40 a 180 dias no pedido'. Usado o prazo máximo (180 dias), "
                                    "cenário mais conservador para o caixa. Altere se o fornecedor tiver outro prazo.", "Ferramenta")
        ws["K29"].comment = Comment("Premissa: 'pagto - 28/56/84' (3 parcelas iguais contadas da data do pedido).", "Ferramenta")
        hdr(ws, 4, 18, ["MDO\n(R$)", "MATERIAL\npeso (pagamento\ndo pedido)", "MATERIAL\npeso ajustado\n(só futuro)", "MATERIAL\n(R$)"], height=58)
        for col in "RSTU": ws.column_dimensions[col].width = 13
        ws.column_dimensions["Q"].width = 2
        put(ws, f"R{RTOT}", f"=SUM(R{R0}:R{R1})", f_b, NUM2, fill_tot); put(ws, f"U{RTOT}", f"=SUM(U{R0}:U{R1})", f_b, NUM2, fill_tot)
    ws.conditional_formatting.add("K16", CellIsRule(operator="equal", formula=['"SEM CURVA"'], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
    ws.conditional_formatting.add("K16", CellIsRule(operator="equal", formula=['"CURVA PREVISION"'], fill=PatternFill("solid", fgColor="E8F0EA"), font=Font(name=F, bold=True, color="1C3B2A")))
    ws.conditional_formatting.add(f"H{R0}:H{R1}", FormulaRule(formula=['$K$16<>"PREVISÃO INFORMADA"'], font=Font(name=F, color="A6A6A6", strike=True)))
    ws.conditional_formatting.add(f"N{R0}:N{R1}", FormulaRule(formula=[f'AND(OR($K$16="CURVA PREVISION",$K$16="{SPLIT}"),$B{R0}>{CORTE})'], fill=PatternFill("solid", fgColor="E8F0EA")))
    ws.conditional_formatting.add("K16", CellIsRule(operator="equal", formula=[f'"{SPLIT}"'], fill=PatternFill("solid", fgColor="E8F0EA"), font=Font(name=F, bold=True, color="1C3B2A")))
    ws.conditional_formatting.add(f"H{R0}:H{R1}", FormulaRule(formula=['$K$16="PREVISÃO INFORMADA"'], fill=PatternFill("solid", fgColor="F6E3DF")))
    put(ws, f"B{RTOT}", "TOTAL", f_b)
    for col in "CDEFH":
        put(ws, f"{col}{RTOT}", f"=SUM({col}{R0}:{col}{R1})", f_b, PCT if col == "F" else NUM2, fill_tot)
    put(ws, f"B{RDIF}", "DIFERENÇA (fluxo em INCC − meta em INCC" + (", descontada a conciliação)" if CONC else ")"), f_b)
    put(ws, f"D{RDIF}", f"=E{RTOT}-K6" + ("+IF(K6-K8>0,K6-K8-K10,0)" if CONC else ""), f_b, "#,##0.00")
    ws.conditional_formatting.add(f"D{RDIF}", FormulaRule(formula=[f"ABS(D{RDIF})>0.01"], fill=fill_red, font=Font(name=F, bold=True, color="B3392B")))
    put(ws, f"L{RTOT}", f"=SUM(L{R0}:L{R1})", f_b, NUM2, fill_tot)
    ws.conditional_formatting.add(f"B{R0}:H{R1}", FormulaRule(formula=[f"$B{R0}<={CORTE}"], fill=fill_past))
    for col, w in zip("ABCDEFGHIJKLMNOP", (9, 10, 15, 16, 10, 10, 10, 18, 44, 15, 24, 12, 12, 12, 12, 13)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C5"
    return ws


for i, t in enumerate(TABS):
    if t == "TAX. ADM": continue
    disc_sheet(i, t)
    if t == "ESTRUTURA" and AUXE:
        sys.path.insert(0, os.environ.get("GERADOR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "gerador"))); import aux_estrutura
        _hf, _info = aux_estrutura.montar(wb, json.load(open(os.path.join(DATA_DIR, AUXE))), R0, NM, CFG.get("aux_estrutura_cfg"))
        for k in range(NM):
            c = wb["ESTRUTURA"][f"H{R0 + k}"]; c.value = _hf(k); c.font = f_link
        wb["ESTRUTURA"]["H4"].comment = Comment("PREVISÃO INFORMADA ligada à aba AUX.ESTRUTURA (quantitativos por trecho/pavimento "
                                                "× datas do cronograma de atividades). Para digitar à mão, apague a configuração "
                                                "aux_estrutura da obra.", "Ferramenta")
        print("AUX.ESTRUTURA:", _info, file=sys.stderr)
    elif t == "ESTRUTURA":
        ax = wb.create_sheet("AUX.ESTRUTURA"); copy_values(ax, d2["AUX.ESTRUTURA"])
# TAX. ADM segue a ordem do original (logo após PROJETO)
tx = disc_sheet(ND - 1, "TAX. ADM")
wb.move_sheet(tx, offset=wb.sheetnames.index("PROJETO") + 1 - wb.sheetnames.index("TAX. ADM"))


# ================================================================== DASHBOARD
ws = wb.create_sheet("DASHBOARD", 1)
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
BIM, BIV = "'MODELO PARA BI'!$H$1:$CA$1", "'MODELO PARA BI'!$H$65:$CA$65"
ANO_R = ANO_RT
MESES_PT = '"Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"'
BRL = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"-"'
BRL0 = '"R$" #,##0;[Red]-"R$" #,##0;"-"'
NAVY, INK, MUT = "1C3B2A", "3C3A33", "8C877A"
C_REAL, C_PROJ = "1F7A52", "99BCAB"          # par validado (azul = realizado, laranja = a realizar)
TEAL, TEAL2, TEALH = "F1F5F0", "FFFFFF", "C5D3C8"
fill = lambda h: PatternFill("solid", fgColor=h)
side_n = Side(style="medium", color=NAVY); side_t = Side(style="thin", color="EDE8DC"); side_w = Side(style="thin", color="FFFFFF")
for col, w in zip("ABCDEFGHIJKLMNOPQRST", (2, 11, 25, 20, 11, 22, 2, 22, 21, 2, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11)):
    ws.column_dimensions[col].width = w
# ---- faixa de título
put(ws, "B2", f'={OBRA}&"  ·  FLUXO DE CAIXA DA OBRA"', Font(name=F, bold=True, size=20, color=NAVY), border=False)
put(ws, "B3", (f'="Data de corte: "&CHOOSE(MONTH({CORTE}),{MESES_PT})&"/"&YEAR({CORTE})'
               f'&"     ·     Ano do retrato: "&{ANO_R}&"     ·     Cenário: "&IF({VP}!$K$8="PADRÃO","{CEN_NOME_P} (curvas digitadas)","{CEN_NOME}")&"     ·     R$ nominais, com taxa de adm."'),
    Font(name=F, size=10, color="8C877A"), border=False)
ws.row_dimensions[2].height = 34; ws.row_dimensions[3].height = 20
ws["B2"].alignment = Alignment(vertical="center"); ws["B3"].alignment = Alignment(vertical="top")
# ---- cartões
cards = [("B", "C", "Projeção total", "='MODELO PARA BI'!E66", BRL0, "com contingência" if CONT else "orçamento em INCC convertido (R$)"),
         ("D", "E", "Realizado até o corte", f"=SUMPRODUCT(({BIM}<={CORTE})*{BIV})", BRL0, '=FIXED(IF(B6=0,0,D6/B6)*100,1)&"% da projeção"'),
         ("F", "F", None, f"=SUMPRODUCT((YEAR({BIM})={ANO_R})*{BIV})", BRL0, "realizado + previsto"),
         ("H", "H", "Saldo a realizar", f"=SUM({BIV})-D6", BRL0, "após a data de corte"),
         ("I", "I", "Conferência", f'=IF(COUNTIF({VP}!$F$54:$F$59,"⚠*")=0,"✓ Tudo OK","⚠ "&COUNTIF({VP}!$F$54:$F$59,"⚠*")&" pendência(s)")', None, "detalhes em VALOR POR PL")]
for c0, c1, lab, f, fmt, note in cards:
    card_box(ws, c0, c1, 5, 7)
    put(ws, f"{c0}5", lab.upper() if lab else f'="DESEMBOLSO EM "&{ANO_R}', Font(name=F, bold=True, size=8, color=MUT), border=False)
    put(ws, f"{c0}6", f, Font(name=F, bold=True, size=14, color=NAVY), fmt, border=False)
    put(ws, f"{c0}7", note, Font(name=F, size=8, color=MUT), border=False)
    for rr in (5, 6, 7):
        ws[f"{c0}{rr}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    if c0 != c1:
        for rr in (5, 6, 7): ws.merge_cells(f"{c0}{rr}:{c1}{rr}")
ws.row_dimensions[5].height = 22; ws.row_dimensions[6].height = 30; ws.row_dimensions[7].height = 18
ws.conditional_formatting.add("I5:I7", FormulaRule(formula=['LEFT($I$6,1)="✓"'], fill=fill("E8F0EA")))
ws.conditional_formatting.add("I5:I7", FormulaRule(formula=['LEFT($I$6,1)="⚠"'], fill=fill("F6E3DF")))
ws.conditional_formatting.add("I6", FormulaRule(formula=['LEFT(I6,1)="✓"'], font=Font(name=F, bold=True, size=14, color="1F7A52")))
ws.conditional_formatting.add("I6", FormulaRule(formula=['LEFT(I6,1)="⚠"'], font=Font(name=F, bold=True, size=14, color="B3392B")))
# ---- tabela anual
T0 = 10
tb = lambda: Border(bottom=side_t)
put(ws, f"B{T0}", "REALIZADO (ATUAL)", Font(name=F, bold=True, size=12, color=NAVY), fill=fill(TEALH), border=False)
ws.merge_cells(f"B{T0}:F{T0}")
put(ws, f"B{T0 + 1}", f'=CHOOSE(MONTH(DATE(YEAR({CORTE}),MONTH({CORTE})+1,1)),{MESES_PT})&"/"&YEAR(DATE(YEAR({CORTE}),MONTH({CORTE})+1,1))',
    Font(name=F, bold=True, size=10, color=NAVY), fill=fill(TEALH), border=False)
ws.merge_cells(f"B{T0 + 1}:F{T0 + 1}")
ws[f"B{T0 + 1}"].comment = Comment("Mês do relatório = mês seguinte à data de corte.", "Ferramenta")
for rr in (T0, T0 + 1): ws[f"B{rr}"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[T0].height = 22
put(ws, f"H{T0 + 1}", "COMPOSIÇÃO DO ANO", Font(name=F, bold=True, size=8, color=MUT), border=False); ws.merge_cells(f"H{T0 + 1}:I{T0 + 1}")
ws[f"H{T0 + 1}"].alignment = Alignment(horizontal="center", vertical="bottom")
heads = {"B": "Ano", "C": "Anual\nR$", "D": "Acumulado\nR$", "E": "Anual\n%", "F": "Acumulado\n%", "H": "Realizado\nR$", "I": "A realizar\nR$"}
for col, h in heads.items():
    c = put(ws, f"{col}{T0 + 2}", h, Font(name=F, bold=True, color="FFFFFF"), fill=fill(NAVY), border=False)
    c.alignment = Alignment(horizontal="center" if col == "B" else "right", vertical="center", wrap_text=True)
ws[f"H{T0 + 2}"].fill = fill(C_REAL); ws[f"I{T0 + 2}"].fill = fill(C_PROJ)
ws.row_dimensions[T0 + 2].height = 32
NY = 7; Y0 = T0 + 3; YT = Y0 + NY
for k in range(NY):
    r = Y0 + k; bg = fill(TEAL if k % 2 == 0 else TEAL2)
    put(ws, f"B{r}", f"=YEAR({INICIO})+{k}", Font(name=F, bold=True, color=INK), "0", bg, border=False)
    put(ws, f"C{r}", f"=SUMPRODUCT((YEAR({BIM})=B{r})*{BIV})", f_calc, BRL, bg, border=False)
    put(ws, f"D{r}", f"=C{r}" if k == 0 else f"=D{r - 1}+C{r}", f_calc, BRL, bg, border=False)
    put(ws, f"E{r}", f"=IF($C${YT}=0,0,C{r}/$C${YT})", f_calc, "0.00%", bg, border=False)
    put(ws, f"F{r}", f"=IF($C${YT}=0,0,D{r}/$C${YT})", f_calc, "0.00%", bg, border=False)
    put(ws, f"H{r}", f"=SUMPRODUCT((YEAR({BIM})=B{r})*({BIM}<={CORTE})*{BIV})", f_calc, BRL, bg, border=False)
    put(ws, f"I{r}", f"=C{r}-H{r}", f_calc, BRL, bg, border=False)
    ws[f"B{r}"].alignment = Alignment(horizontal="center")
    for col in "BCDEFHI": ws[f"{col}{r}"].border = tb()
    ws.row_dimensions[r].height = 19
for col in "BCDEFHI":
    c = ws[f"{col}{YT}"]; c.fill = fill(NAVY); c.font = Font(name=F, bold=True, color="FFFFFF")
put(ws, f"B{YT}", "TOTAL", None, border=False); ws[f"B{YT}"].alignment = Alignment(horizontal="center")
ws[f"C{YT}"] = f"=SUM(C{Y0}:C{YT - 1})"; ws[f"C{YT}"].number_format = BRL
ws[f"E{YT}"] = f"=SUM(E{Y0}:E{YT - 1})"; ws[f"E{YT}"].number_format = "0.00%"
ws[f"H{YT}"] = f"=SUM(H{Y0}:H{YT - 1})"; ws[f"H{YT}"].number_format = BRL
ws[f"I{YT}"] = f"=SUM(I{Y0}:I{YT - 1})"; ws[f"I{YT}"].number_format = BRL
ws.row_dimensions[YT].height = 22
ws.conditional_formatting.add(f"B{Y0}:I{YT - 1}", FormulaRule(formula=[f"$B{Y0}={ANO_R}"], font=Font(name=F, bold=True, color=NAVY),
                                                             fill=fill("D9E6DC"), border=Border(top=side_n, bottom=side_n)))
put(ws, f"B{YT + 1}", "Linha destacada = ano do retrato (RETRATO 2027!D5).  Realizado = até a data de corte; a realizar = previsto depois dela.",
    Font(name=F, italic=True, size=8, color=MUT), border=False)
# ---- gráfico 1: realizado × a realizar por ano (colunas empilhadas, um só eixo)
ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.gapWidth = 60
ch.title = None; ch.height = 8.2; ch.width = 17
put(ws, f"K{T0 - 2}", "DESEMBOLSO ANUAL (R$ MILHÕES)", Font(name=F, bold=True, size=10, color=NAVY), border=False)
ch.add_data(Reference(ws, min_col=8, max_col=9, min_row=T0 + 2, max_row=YT - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=Y0, max_row=YT - 1))
from openpyxl.chart.series import SeriesLabel
for s_, colr, nm in zip(ch.series, (C_REAL, C_PROJ), ("Realizado", "A realizar")):
    s_.tx = SeriesLabel(v=nm)
    s_.graphicalProperties.solidFill = colr; s_.graphicalProperties.line.solidFill = "FFFFFF"
ch.y_axis.numFmt = '#,##0,,"M"'; ch.y_axis.majorGridlines.spPr = None; ch.legend.position = "b"
ch.x_axis.delete = False; ch.y_axis.delete = False
ws.add_chart(ch, f"K{T0 - 1}")
# ---- desembolso do ano do retrato por grupo
G0 = YT + 4
put(ws, f"B{G0}", f'="DESEMBOLSO EM "&{ANO_R}&" POR GRUPO"', Font(name=F, bold=True, size=10, color=NAVY), border=False)
gh = {"B": "Grupo", "D": "R$", "E": "% do ano", "F": "Avanço\nfinanceiro"}
for col in "BCDEF":
    c = put(ws, f"{col}{G0 + 1}", gh.get(col), Font(name=F, bold=True, color="FFFFFF"), fill=fill(NAVY), border=False)
    c.alignment = Alignment(horizontal="left" if col == "B" else "right", vertical="center", wrap_text=True, indent=1 if col == "B" else 0)
ws.merge_cells(f"B{G0 + 1}:C{G0 + 1}"); ws.row_dimensions[G0 + 1].height = 30
RT = "'RETRATO 2027'"
_rr = lambda a, b: (RROW[a], RROW[b])
def _g(lab, a, b):
    r0, r1 = _rr(a, b); return (lab, f"F{r0}:F{r1}", f"M{r0}:M{r1}", f"J{r0}:J{r1}")
groups = [_g("Indiretos (01 a 11)", 0, 10), _g("Infraestrutura (12 a 16)", 11, 15), _g("Estrutura (17)", 16, 16),
          _g("Vedações, instalações e elevadores (18 a 24)", 17, 23), _g("Acabamentos (25 a 37)", 24, 36),
          ("Taxa de administração", f"F{RROW[ND - 1]}:F{RROW[ND - 1]}", None, None)]
for k, (lab, rk, rm, rl) in enumerate(groups):
    r = G0 + 2 + k; bg = fill(TEAL if k % 2 == 0 else TEAL2)
    put(ws, f"B{r}", lab, f_calc, fill=bg, border=False); ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"].alignment = Alignment(indent=1)
    put(ws, f"D{r}", f"=SUM({RT}!{rk})", f_calc, BRL, bg, border=False)
    put(ws, f"E{r}", f"=IF($D${G0 + 8}=0,0,D{r}/$D${G0 + 8})", f_calc, "0.0%", bg, border=False)
    put(ws, f"F{r}", f"=IFERROR(SUMPRODUCT({RT}!{rl},{RT}!{rm})/SUM({RT}!{rm}),0)" if rm else None, f_calc, "0.0%", bg, border=False)
    for col in "BCDEF": ws[f"{col}{r}"].border = tb()
    ws.row_dimensions[r].height = 19
r = G0 + 8
for col in "BCDEF":
    c = ws[f"{col}{r}"]; c.fill = fill(NAVY); c.font = Font(name=F, bold=True, color="FFFFFF")
ws[f"B{r}"] = "TOTAL"; ws.merge_cells(f"B{r}:C{r}"); ws[f"B{r}"].alignment = Alignment(indent=1)
ws[f"D{r}"] = f"=SUM(D{G0 + 2}:D{G0 + 7})"; ws[f"D{r}"].number_format = BRL
ws[f"E{r}"] = f"=SUM(E{G0 + 2}:E{G0 + 7})"; ws[f"E{r}"].number_format = "0.0%"
ws.row_dimensions[r].height = 22
put(ws, f"B{r + 1}", "Avanço financeiro = acumulado até dezembro do ano ÷ projeção.", Font(name=F, italic=True, size=8, color=MUT), border=False)
put(ws, f"F{r + 1}", f'=IF(ABS(D{r}-F6)<1,"✓ confere com o cartão","⚠ diferente do cartão")', Font(name=F, italic=True, size=8, color=MUT), border=False)
bc = BarChart(); bc.type = "bar"; bc.title = None; bc.legend = None; bc.height = 8.2; bc.width = 17; bc.gapWidth = 50
bc.add_data(Reference(ws, min_col=4, min_row=G0 + 2, max_row=G0 + 7)); bc.set_categories(Reference(ws, min_col=2, min_row=G0 + 2, max_row=G0 + 7))
bc.series[0].graphicalProperties.solidFill = C_REAL; bc.series[0].graphicalProperties.line.solidFill = "FFFFFF"
bc.y_axis.numFmt = '#,##0,,"M"'; bc.x_axis.scaling.orientation = "maxMin"; bc.y_axis.majorGridlines = None
bc.x_axis.delete = False; bc.y_axis.delete = False
ws.add_chart(bc, f"K{G0 - 1}")
put(ws, f"K{G0 - 2}", f'="DESEMBOLSO EM "&{ANO_R}&" POR GRUPO (R$ MILHÕES)"', Font(name=F, bold=True, size=10, color=NAVY), border=False)
# ---- curva S (% acumulado da projeção, mês a mês)
S0 = G0 + 12
put(ws, f"B{S0}", "CURVA S — % ACUMULADO DA PROJEÇÃO", Font(name=F, bold=True, size=10, color=NAVY), border=False)
lc = LineChart(); lc.title = None; lc.height = 7.5; lc.width = 33; lc.legend = None
lc.add_data(Reference(wb["MODELO PARA BI"], min_col=8, max_col=8 + NM - 1, min_row=67), from_rows=True)
lc.set_categories(Reference(wb["MODELO PARA BI"], min_col=8, max_col=8 + NM - 1, min_row=1))
lc.series[0].graphicalProperties.line.solidFill = C_REAL; lc.series[0].graphicalProperties.line.width = 28575; lc.series[0].smooth = True
lc.y_axis.numFmt = "0%"; lc.y_axis.scaling.min = 0; lc.y_axis.scaling.max = 1; lc.x_axis.number_format = "mmm/yy"; lc.x_axis.tickLblSkip = 6
lc.x_axis.delete = False; lc.y_axis.delete = False
ws.add_chart(lc, f"B{S0 + 1}")
ws.sheet_properties.tabColor = "1F7A52"
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
painel_fundo(ws, 30, S0 + 20)
painel_fundo(wb["RETRATO 2027"], 13, RT_TOT + 40)
# painéis extras (Curva física e Gantt), quando o gerador recebe os dados
import paineis_extra
if os.environ.get("FIS_JSON"): paineis_extra.curva_fisica(wb, json.load(open(os.environ["FIS_JSON"])), CFG.get("obra", ""))
if os.environ.get("GANTT_JSON"): paineis_extra.gantt(wb, json.load(open(os.environ["GANTT_JSON"])), CFG.get("obra", ""))
if os.environ.get("ESQ"):   # aba ESQUEMÁTICO (corte por pavimento), gerador/esquematico.py
    sys.path.insert(0, os.environ.get("GERADOR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "gerador"))); import esquematico
    print("ESQUEMÁTICO:", esquematico.esquematico(wb, CFG.get("obra", ""), CFG.get("esq_marcos"), CFG.get("esq_subdiv"), CFG.get("esq_extra")), file=sys.stderr)

for n in ("VALOR POR PL", '"INCC-FGV"', "CURVA PREVISION", "DADOS BRUTOS"):
    wb[n].sheet_properties.tabColor = "FFC000"
for n in ("RETRATO 2027", "MODELO PARA BI", "AUX. CURVA", "CP RESUMO", "DADOS HISTORICOS"):
    wb[n].sheet_properties.tabColor = "1C3B2A"
wb.active = wb.sheetnames.index("DASHBOARD")
for w in wb.worksheets:
    w.sheet_view.tabSelected = (w.title == "DASHBOARD")
import os
if os.environ.get("ONLY_DASH"):
    wb.active = wb.sheetnames.index(os.environ["ONLY_DASH"])
    for w in wb.worksheets:
        w.sheet_view.tabSelected = (w.title == os.environ["ONLY_DASH"])
        if w.title != os.environ.get("ONLY_DASH"): w.sheet_state = "hidden"
# acabamento da planilha para download: sem linhas de grade e gráficos com curvas suavizadas e cantos arredondados
for w in wb.worksheets:
    w.sheet_view.showGridLines = False
    for ch_ in w._charts:
        ch_.roundedCorners = True
        estilo_grafico(ch_)
        for sub in ch_._charts:
            if isinstance(sub, LineChart):
                for s_ in sub.series: s_.smooth = True
wb.save(OUT)
print("ok", len(wb.sheetnames), "abas:", wb.sheetnames)
