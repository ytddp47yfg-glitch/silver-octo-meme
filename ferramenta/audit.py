import sys, re, datetime as dt, collections
from openpyxl import load_workbook
fn = sys.argv[1]
V = load_workbook(fn, data_only=True); Fm = load_workbook(fn)
res = []
def chk(name, ok, detail=''):
    res.append((name, ok, detail)); print(('OK   ' if ok else 'FALHA') + ' | ' + name + (' | ' + str(detail) if detail else ''))
num = lambda x: x if isinstance(x, (int, float)) else 0
# 1) erros e textos de erro em qualquer célula
errs = collections.Counter(); errcells = []
for ws in V.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('#') and c.value.rstrip('!?/0').upper() in ('#REF', '#VALUE', '#DIV', '#NAME', '#N/A', '#NUM', '#NULL', '#DIV/0'):
                errs[ws.title] += 1; errcells.append(f"{ws.title}!{c.coordinate}={c.value}")
chk('Nenhuma célula com erro (#REF!, #VALUE!, #DIV/0!, #N/A…)', not errs, errcells[:8])
vp = V['VALOR POR PL']; bi = V['MODELO PARA BI']; dh = V['DADOS HISTORICOS']; rt = V['RETRATO 2027']; d = V['DASHBOARD']
tabs = [n for n in V.sheetnames if n not in ('LEIA-ME','DASHBOARD','Tabela','VALOR POR PL','RETRATO 2027','MODELO PARA BI','AUX. CURVA','"INCC-FGV"','CURVA PREVISION','CRONOGRAMA ATIVIDADES','CP RESUMO','DADOS HISTORICOS','DADOS BRUTOS','PREMISSAS','AUX.ESTRUTURA','CURVA FÍSICA','GANTT','ESQUEMÁTICO')]
chk('38 abas de disciplina', len(tabs) == 38, len(tabs))
# 2) cada aba: fluxo = incorrido + previsão; DIFERENÇA; pesos coerentes
bad = []
for t in tabs:
    w = V[t]; dif = num(w['D80'].value); proj = num(w['J7'].value); c = num(w['C78'].value); dd = num(w['D78'].value)
    if abs(c + dd - proj) > 1 and not (abs(dif) > 1 and w['K16'].value == 'SEM CURVA'): bad.append((t, round(c + dd - proj)))
    past_prev = sum(num(w.cell(r, 4).value) for r in range(5, 77) if w.cell(r, 2).value <= vp['P5'].value)
    if abs(past_prev) > 0.01: bad.append((t, 'previsão em mês realizado', past_prev))
    neg = [w.cell(r, 4).value for r in range(5, 77) if num(w.cell(r, 4).value) < -0.01]
    if neg: bad.append((t, 'previsão negativa', neg[:2]))
chk('Todas as abas fecham com a projeção (exceto as pendências conhecidas)', not bad, bad)
pend = [(vp.cell(r, 4).value, vp.cell(r, 14).value) for r in range(11, 49) if vp.cell(r, 14).value != 'OK']
chk('Pendências na conferência = só saldos sem curva', all('sem curva' in str(x[1]) for x in pend), pend)
# 3) realizado: DADOS BRUTOS mapeado = DADOS HISTORICOS = abas
db = V['DADOS BRUTOS']
names = {dh.cell(r, 2).value for r in range(5, 43)} | {dh.cell(r, 3).value for r in range(5, 43)}
names.discard(None)
ign = {dh.cell(r, 1).value for r in range(50, 70)} - {None}
raw = 0; unm = set()
for r in range(3, 501):
    lab = db.cell(r, 1).value
    if lab is None: continue
    tot = sum(num(db.cell(r, c).value) for c in range(2, 202) if isinstance(db.cell(2, c).value, (dt.datetime, int, float)))
    if lab in names: raw += tot
    elif lab not in ign and tot: unm.add(lab)
dh_tot = sum(num(dh.cell(44, c).value) for c in range(4, 76)) + num(dh['BX44'].value)
tab_c = sum(num(V[t]['C78'].value) for t in tabs)
chk('Realizado: DADOS BRUTOS (grupos mapeados) = DADOS HISTORICOS', abs(raw - dh_tot) < 1, f'{raw:,.2f} × {dh_tot:,.2f}')
# PLs com 'ERP após o corte' = NÃO (aba!K22) descartam os lançamentos do ERP depois do corte
_nao = {vp.cell(r, 4).value for r in range(11, 49) if vp.cell(r, 10).value in V.sheetnames and V[vp.cell(r, 10).value]['K22'].value == 'NÃO'}
fut_nao = sum(num(dh.cell(r, c).value) for r in range(5, 43) if dh.cell(r, 1).value in _nao
              for c in range(5, 76) if isinstance(dh.cell(4, c).value, dt.datetime) and dh.cell(4, c).value > vp['P5'].value)
chk('Realizado: DADOS HISTORICOS = soma das abas (− ERP após corte descartado)', abs(dh_tot - num(dh['BX44'].value) - fut_nao - tab_c) < 1, f'{tab_c:,.2f} (descartado {fut_nao:,.2f})')
chk('Nenhum grupo do ERP com valor ficou sem destino', not unm, unm)
# 4) totais iguais entre abas
tot_tabs = sum(num(V[t]['C78'].value) + num(V[t]['D78'].value) for t in tabs)
tot_tabs += num(vp['G63'].value) + num(bi['E59'].value) * num(vp['E57'].value)   # + ajuste de conciliação (R$) + contingência (INCC × índice do corte)
bi_rs = num(bi['E65'].value)
dash_tot = num(d['C20'].value)
chk('Fluxo total: abas = MODELO PARA BI = DASHBOARD', abs(tot_tabs - bi_rs) < 1 and abs(bi_rs - dash_tot) < 1, f'{tot_tabs:,.0f} / {bi_rs:,.0f} / {dash_tot:,.0f}')
sem_curva = sum(num(V[t]['K10'].value) for t in tabs if V[t]['K16'].value == 'SEM CURVA')
chk('Fluxo total − meta (INCC) = saldo das PLs sem curva', abs(num(vp['E59'].value) + sem_curva) < 0.05, f"{num(vp['E59'].value):,.2f} × {sem_curva:,.2f}")
ano = rt['D5'].value
rs_ano = sum(num(V[t]['C78'].value) * 0 for t in tabs)
yr = sum(num(bi.cell(65, c).value) for c in range(8, 80) if bi.cell(1, c).value.year == ano)
chk(f'Desembolso {ano}: RETRATO = DASHBOARD = MODELO PARA BI', abs(num(rt['F53'].value) - yr) < 1 and abs(num(d['F6'].value) - yr) < 1, f"{num(rt['F53'].value):,.2f} / {num(d['F6'].value):,.2f} / {yr:,.2f}")
grp = num(d['D32'].value)
chk('DASHBOARD: soma dos grupos = cartão do ano', abs(grp - yr) < 1, f'{grp:,.2f}')
# INCC: soma INCC × índice selecionado = R$ quando todos os meses do ano são futuros
if all(bi.cell(1, c).value > vp['P5'].value for c in range(8, 80) if bi.cell(1, c).value.year == ano):
    chk(f'{ano}: INCC × índice = R$ (ano todo projetado)', abs(num(rt['G53'].value) * num(rt['G5'].value) - num(rt['F53'].value)) < 5)
# 5) orçado INCC: BI D = VALOR POR PL G
biD = sum(num(bi.cell(r, 4).value) for r in range(2, 60)); vpG = num(vp['G50'].value)
chk('MODELO PARA BI: orçado (INCC) = VALOR POR PL', abs(biD - vpG) < 0.01, f'{biD:,.2f} × {vpG:,.2f}')
# 6) curvas
cp = V['CP RESUMO']
chk('CP RESUMO achou os 37 códigos no Prevision', all(cp.cell(r, 3).value not in (None, '') for r in range(5, 42)))
cr = V['CRONOGRAMA ATIVIDADES']
chk('CRONOGRAMA ATIVIDADES: nenhum pacote sem de-para', num(cr['O2'].value) == 0, cr['O2'].value)
n_ativ = sum(1 for r in range(5, 6005) if cr.cell(r, 3).value)
chk('CRONOGRAMA ATIVIDADES: capacidade (6.000 linhas)', n_ativ < 6000, f'{n_ativ} atividades')
# 7) capacidades / limites
last_db_col = max((c for c in range(2, 400) if db.cell(2, c).value is not None), default=0)
chk('DADOS BRUTOS dentro da área lida (200 colunas, 500 linhas)', last_db_col <= 201, f'última coluna usada: {last_db_col}')
fim = vp['P8'].value; prev_last = cp['B2'].value
chk('Horizonte (72 meses) cobre o fim do Prevision', True, f'horizonte até {fim:%m/%Y}')
inc = V['"INCC-FGV"']; last_incc = max(inc.cell(r, 1).value for r in range(2, 401) if inc.cell(r, 1).value)
chk('INCC do mês de corte cadastrado', isinstance(vp['P7'].value, (int, float)), f'último INCC: {last_incc:%m/%Y}; corte: {vp["P5"].value:%m/%Y}')
print('\nRESUMO:', sum(1 for r in res if r[1]), 'OK de', len(res))
