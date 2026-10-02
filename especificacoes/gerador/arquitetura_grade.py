"""Acrescenta na grade as especificações das plantas de arquitetura (CDJ-ARQ-EX, Quadro de Especificações - Geral).

Usa arq_quadro.py (código -> descrição) e arq_ambientes.py (etiquetas ○ piso, △ parede/rodapé, □ teto por ambiente).
Regras: célula vazia -> preenche (laranja, INFO. REVISADA); igual ao que já está -> mantém;
diferente do que já está -> mantém as duas e pinta de vermelho claro (INFO. CONFLITANTE).

Uso: python3 arquitetura_grade.py <pasta com os PDFs> <grade.xlsx> <saida_log.json>
"""
import sys, os, glob, json, copy, re, difflib
import openpyxl
from openpyxl.styles import PatternFill, Font
sys.path.insert(0, os.path.dirname(__file__))
from arq_quadro import quadro
from arq_ambientes import ambientes

PASTA, XLSX, LOG = sys.argv[1:4]
pdfs = {re.search(r'PR\d{3}', f).group(0): f for f in glob.glob(os.path.join(PASTA, '*.pdf'))}
Q = quadro(pdfs['PR003'])
TAGS = []
for pr in ('PR003', 'PR004', 'PR005', 'PR007', 'PR009', 'PR011', 'PR013'):  # PR008 e PR012 (trecho 2) conferidas à parte em arq_conferencia.py
    TAGS += ambientes(pdfs[pr])

def tag(pr, amb, area=None, xy=None):
    for t in TAGS:
        if t['prancha'] != pr: continue
        if xy and abs(t['xy'][0] - xy[0]) < 3 and abs(t['xy'][1] - xy[1]) < 3: return t
        if not xy and t['ambiente'] == amb and (area is None or t.get('area') == area): return t
    raise KeyError((pr, amb, area, xy))

# linha da grade -> etiquetas das plantas (prancha, ambiente, área ou posição)
MAPA = {
 'TÉRREO ': {
  'GUARITA': [('PR003', 'GUARITA', '7,42 m²')],
  'LAVABO GUARITA': [('PR003', 'LAVABO', '1,89 m²')],
  'CIRCULAÇÃO GUARITA': [('PR003', 'CIRC.', '5,48 m²')],
  'CIRCULAÇÃO ELEVADOR DE SERVIÇO/ANTICÂRA': [('PR003', 'ANTECÂMARA', '5,78 m²')],
  'GARAGEM': [('PR005', None, None, (531, 2264))],
  'CASA DE BOMBAS': [('PR003', 'C. BOMBAS', '4,41 m²')],
  'CIRCULAÇÃO#1': [('PR003', 'CIRCULAÇÃO', '35,62 m²')],
  'CONCIERGE': [('PR003', 'CONCIERGE', '11,72 m²')],
  'DELIVERY': [('PR003', 'DELIVERY', '8,59 m²')],
  'CIRCULAÇÃO SERVIÇO': [('PR003', 'CIRC. SERV.', '29,56 m²')],
  'CAR WASH': [('PR005', 'CAR WASH', None)],
  'BICICLETÁRIO': [('PR005', 'BICICLETÁRIO', None)],
  'ZELADORIA': [('PR005', 'ZELADORIA', None)],
  'ARS 1': [('PR005', 'A.R.S.', '24,09 m²')],
  'ARS 2': [('PR005', 'A.R.S.', '14,93 m²')],
  'ARS 3': [('PR005', 'A.R.S.', '4,25 m²')],
  'SALA MOTORISTA': [('PR005', 'SALA MOTORISTA', None)],
  'DML': [('PR005', 'DML', '7,01 m²')],
  'VEST. P.C.D': [('PR005', 'VEST. P.C.D.', None)],
  'VEST. FEMININO': [('PR005', 'VESTIÁRIO FEM.', None)],
  'VEST. MASC.': [('PR005', 'VESTIÁRIO MASC.', None)],
  'ESTAR/COPA FUNCIONÁRIOS': [('PR005', 'ESTAR/COPA FUNC.', None)],
  'CIRCULAÇÃO#2': [('PR005', 'CIRCULAÇÃO', '24,85 m²')],
  'GERADOR': [('PR004', 'GERADOR', None)],
  'VAGAS BOX': [('PR005', 'VAGA BOX 01', None)],
  'QUADROS ELÉTRICOS': [('PR003', 'QUADROS ELÉTRICOS', None)],
  '+ACESSO E SAÍDA DE VEÍCULOS': [('PR003', 'ACESSO VEÍCULOS', '287,93 m²')],
  '+ESTACIONAMENTO VISITANTES': [('PR005', 'ESTACIONAMENTO VISTANTES', None)],
  '+RAMPA VEÍCULOS': [('PR005', 'RAMPA VEÍCULOS', None)],
  '+DEPÓSITO (A=8,85 m²)': [('PR005', 'DEPÓSITO', '8,85 m²')],
  '+SERVIÇO JUNTO AO PORTÃO (A=6,87 m²)': [('PR003', 'SERV.', '6,87 m²', (2447, 996))],
  '+CIRCULAÇÃO JUNTO À ESCADA (A=4,06 m²)': [('PR003', 'CIRC.', '4,06 m²')],
 },
 '2° PAVTO': {
  'CASA DE BOMBAS PISCINA 1': [('PR007', 'CASA DE BOMBAS PISCINA', None)],
  'SHAFT ELÉTRICA': [('PR007', 'SHAFT ELÉTRICA', None)],
  'CIRCULAÇÃO/ANTECÂMARA': [('PR007', 'CIRC.', '5,86 m²'), ('PR007', 'ANTECÂMARA', '5,78 m²')],
  'PRESSURIZAÇÃO/ANTECÂMARA': [('PR007', 'PRESSURIZAÇÃO', None), ('PR007', 'ANTECÂMARA', '3,49 m²')],
  'ÁREA TÉCNICA 1': [('PR007', 'ÁREA TÉCNICA', '13,08 m²')],
  'ÁREA TÉCNICA 2': [('PR007', 'ÁREA TÉCNICA', '8,14 m²')],
  'GARAGEM': [('PR009', 'GARAGEM', None)],
  'VAGAS BOX': [('PR009', 'VAGA BOX 01', None)],
  'TELECOM': [('PR009', 'TELECOM', None)],
  'SALA DE SEGURANÇA': [('PR009', 'SALA DE SEGURANÇA', None)],
  'CASA DE BOMBAS PISCINA 2 (FUNDO)': [('PR009', 'CASA DE BOMBAS PISCINA', None)],
  'DML': [('PR009', 'DML', None)],
  'RESERVATÓRIO INFERIOR': [('PR009', 'RESERVATÓRIO INFERIOR', None)],
  'HALL TÉCNICO/CIRCULAÇÃO': [('PR009', 'HALL TÉCNICO', None), ('PR009', 'CIRC.', '51,88 m²')],
  '+RAMPA VEÍCULOS': [('PR009', None, None, (528, 1602))],
  '+CIRCULAÇÃO ÁREA TÉCNICA (A=14,54 m²)': [('PR007', 'CIRC.', '14,54 m²')],
  '+CIRCULAÇÃO JUNTO À ESCADA (A=3,07 m²)': [('PR007', 'CIRC.', '3,07 m²')],
 },
 'PILOTIS': {
  'CIRCULAÇÃO/ANTECÂMARA ELEVADOR SERVIÇO': [('PR011', 'CIRC.', '5,76 m²')],
  'ANTECÂMARA': [('PR011', 'ANTECÂMARA', '5,78 m²')],
  'QUADRA TÊNIS': [('PR013', 'QUADRA DE TÊNIS', None)],
  'DEP JARDIM': [('PR013', 'DEP. JARDIM', None)],
  'DEPÓSITO': [('PR013', 'DEPÓSITO', None)],
  'QUADRA INFANTIL': [('PR013', 'QUADRA INFANTIL', None)],
  'ESCADA EXTERNA': [('PR013', None, None, (1686, 804))],
  '+CIRCULAÇÃO JUNTO À ESCADA (A=2,91 m²)': [('PR011', 'CIRC.', '2,91 m²')],
 },
}

def textos(t):
    out = {'B': [], 'F': [], 'H': [], 'I': []}
    for c in t['piso']: out['B'].append(f'ARQ-P{c} - ' + Q['PISO'][c])
    for c in t['parede']:
        if int(c) >= 19: out['F'].append(f'ARQ-RD{c} - ' + Q['PAREDE'][c])
        else: out['H'].append(f'ARQ-PA{c} - ' + Q['PAREDE'][c])
    for c in t['teto']:
        if c in Q['TETO']: out['I'].append(f'ARQ-T{c} - ' + Q['TETO'][c])
        else: out['I'].append(f'ARQ-T{c} - código não existe no quadro de tetos (01 a 06); conferir na planta')
    return out

def norm(s):
    s = re.sub(r'^\S+ - ', '', s).lower()
    return re.sub(r'[^a-z0-9à-ú ]', '', s)

def igual(a, b):
    # mesma numeração do quadro de arquitetura (a tabela v00 usa os mesmos números: "07 - ...", "01| T1| T8 - ...")
    num = re.search(r'ARQ-[A-Z]+(\d+)', a).group(1)
    pre = b.split(' - ')[0]
    if re.fullmatch(r'[\d| PAT]+', pre) and num in re.findall(r'\b\d{2}\b', pre):
        return True
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio() > 0.85

wb = openpyxl.load_workbook(XLSX)
ref = wb['PILOTIS']['K7']
REV_FILL, REV_COR = copy.copy(ref.fill), copy.copy(ref.font.color)
CONF_FILL, CONF_COR = PatternFill('solid', fgColor='FFFFC7CE'), 'FF9C0006'
log = []
for aba, mapa in MAPA.items():
    ws = wb[aba]
    last = max(r for r in range(13, ws.max_row + 1) if ws.cell(r, 1).value)
    linhas, vistos = {}, {}
    for r in range(13, last + 1):
        n = (ws.cell(r, 1).value or '').strip()
        if not n: continue
        vistos[n] = vistos.get(n, 0) + 1
        linhas[n if vistos[n] == 1 else f'{n}#{vistos[n]}'] = r
        if vistos[n] == 2: linhas[f'{n}#1'] = linhas[n]
    for chave, refs in mapa.items():
        if chave.startswith('+'):
            nome = chave[1:]
            r = next((x for x in range(13, ws.max_row + 1) if (ws.cell(x, 1).value or '').strip() == nome), None)
            if r is None:
                last += 1; r = last
                for c in range(1, 27):
                    src, dst = ws.cell(last - 1, c), ws.cell(r, c)
                    dst._style = copy.copy(src._style)
                    if c > 1: dst.value = None
                ws.cell(r, 1).value = nome
                ws.row_dimensions[r].height = ws.row_dimensions[last - 1].height
                log.append((aba.strip(), nome, 'CÔMODO/AMBIENTE', 'linha nova', 'planta de arquitetura'))
        else:
            r = linhas[chave]
        nome = ws.cell(r, 1).value.strip()
        novos = {'B': [], 'F': [], 'H': [], 'I': []}
        for pr, amb, area, *xy in refs:
            t = tag(pr, amb, area, xy[0] if xy else None)
            for col, vs in textos(t).items():
                for v in vs:
                    if v not in novos[col]: novos[col].append(v)
        for col, vs in novos.items():
            if not vs: continue
            cel = ws[f'{col}{r}']
            atual = (cel.value or '').strip()
            partes = [p for p in atual.split('\n\n') if p.strip()] if atual else []
            add = [v for v in vs if v not in partes]
            if not add: continue
            hdr = ' '.join(ws[f'{col}11'].value.split())
            if not partes:
                cel.value = '\n\n'.join(add)
                cel.fill = copy.copy(REV_FILL); f = copy.copy(cel.font); f.color = copy.copy(REV_COR); cel.font = f
                log.append((aba.strip(), nome, hdr, '—', ' | '.join(add)))
            else:
                diferentes = [v for v in add if not any(igual(v, p) for p in partes)]
                cel.value = '\n\n'.join(partes + [('CONFLITO – Arquitetura: ' + v) if v in diferentes else v for v in add])
                if diferentes:
                    cel.fill = CONF_FILL; f = copy.copy(cel.font); f.color = CONF_COR; cel.font = f
                    log.append((aba.strip(), nome, hdr, 'CONFLITO: ' + ' | '.join(partes), ' | '.join(add)))
                else:
                    log.append((aba.strip(), nome, hdr, ' | '.join(partes), 'mesma especificação na arquitetura: ' + ' | '.join(add)))
        ws.row_dimensions[r].height = None

# pendências
p = wb['PENDÊNCIAS']
r = p.max_row + 2
p.cell(r, 1, '5. ESPECIFICAÇÕES DAS PLANTAS DE ARQUITETURA (CDJ-ARQ-EX PR003–PR013, Quadro de Especificações - Geral)').font = Font(bold=True, size=12)
r += 1
for i, h in enumerate(['Nº', 'PAVIMENTO / AMBIENTE', 'COLUNA', 'ANTES', 'ARQUITETURA']):
    c = p.cell(r, 1 + i, h); c.font = Font(bold=True, color='FFFFFFFF'); c.fill = PatternFill('solid', fgColor='FF243F2E')
for n, (pav, nome, col, a, b) in enumerate(log, 1):
    r += 1
    for i, v in enumerate([n, f'{pav} – {nome}', col, a, b]):
        c = p.cell(r, 1 + i, v)
        if str(a).startswith('CONFLITO'): c.fill = CONF_FILL
r += 2
obs = [
 ('Legenda das plantas', 'Etiqueta ○ = piso, △ = parede (códigos 01 a 18) ou rodapé (19 a 25), □ = teto. Na grade os códigos aparecem como ARQ-P (piso), ARQ-PA (parede), ARQ-RD (rodapé) e ARQ-T (teto) para não confundir com os códigos do detalhamento de interiores.'),
 ('Ambientes "ACABAMENTOS VER INTERIORES"', 'Lobby, hall de elevadores, antecâmara/eclusa, I.S. P.C.D., vestíbulo (Térreo), halls do 2º pav. e a maior parte do Pilotis remetem ao projeto de interiores; nesses ambientes vale o detalhamento R04 já lançado.'),
 ('TÉRREO – CONCIERGE', 'A etiqueta de teto da planta PR003 traz □19, mas o quadro de tetos só vai de 01 a 06. Provável erro de etiqueta; confirmar com a arquitetura.'),
 ('TÉRREO – ARS 1, 2 e 3', 'A planta tem três A.R.S. (24,09 m², 14,93 m² e 4,25 m²) com a mesma especificação; a ordem 1-2-3 da grade foi assumida.'),
 ('TÉRREO – DML E JARDINAGEM', 'Sem etiqueta de acabamento na planta PR003.'),
 ('PR008 e PR012', 'Pranchas de revisão anterior (R11 e R13 de 2025) às demais; não foram usadas.'),
 ('Portas, janelas e guarda-corpos', 'Os quadros de esquadrias das plantas não foram lançados na grade nesta etapa.'),
]
for k, v in obs:
    r += 1; p.cell(r, 2, k); p.cell(r, 4, v)
wb.save(XLSX)
json.dump({'quadro': Q, 'mudancas': log, 'mapa': {a: {k: [list(x) for x in v] for k, v in m.items()} for a, m in MAPA.items()}}, open(LOG, 'w'), ensure_ascii=False, indent=1)
print(len(log), 'registros;', sum(1 for x in log if str(x[3]).startswith('CONFLITO')), 'conflitos')
