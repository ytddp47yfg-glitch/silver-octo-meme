"""Imagens dentro da própria aba da grade, com link nas células.

As fotos ficam à direita da grade (a partir da coluna AH, fora da área de
impressão), na mesma linha do ambiente. A célula que cita o código recebe a
linha "► VER IMAGEM (clique)" e um hyperlink interno para a célula da foto;
abaixo de cada foto há "◄ VOLTAR" para a célula de origem.
Não há aba separada de imagens. Rodar depois de formatacao.py e antes de
area_impressao.py (formatacao.py apaga as larguras das colunas extras).
"""
import sys, io, copy, json
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image

P = sys.argv[1] if len(sys.argv) > 1 else 'especificacoes/CDJ-GRE-ARQ-AC-R00.xlsx'
DIR = 'especificacoes/imagens/'
MARCA = '► VER IMAGEM (clique)'
COL0 = 34            # AH: primeira coluna de imagens (a grade vai até AF)
LARG = 72.0          # largura das colunas de imagem
BOX_W = 500          # px úteis para a foto
# (aba, célula que cita o item) -> (arquivo da imagem, legenda)
LINKS = [
    ('PILOTIS', 'B41', 'CDJ-IMG-01_ESPACO-TERMAS_P10.jpg', 'P10 - Piso elevado em mármore branco selecionado'),
    ('PILOTIS', 'H41', 'CDJ-IMG-02-04_ESPACO-TERMAS_PA49.jpg', 'PA49 - Mosaico Bisazza (painel e pilar)'),
    ('PILOTIS', 'U41', 'CDJ-IMG-03_ESPACO-TERMAS_D54.jpg', 'D54 - Chuveiro de embutir no teto Raindream, Roca'),
    ('PILOTIS', 'B55', 'CDJ-IMG-05_VESTIARIO-FEM-MASC-TERMAS_P10.jpg', 'P10 - Piso elevado em mármore branco selecionado'),
    ('PILOTIS', 'H55', 'CDJ-IMG-12_VESTIARIO-FEM-MASC-TERMAS_PA47.jpg', 'PA47 - Pintura com tinta mineral, cor 930, Terracor'),
    ('PILOTIS', 'J55', 'CDJ-IMG-07_VESTIARIO-FEM-MASC-TERMAS_D73.jpg', 'D73 - Bancada em mármore branco selecionado'),
    ('PILOTIS', 'K55', 'CDJ-IMG-06_VESTIARIO-FEM-MASC-TERMAS_D38.jpg', 'D38 - Espelho liso incolor 4mm'),
    ('PILOTIS', 'L55', 'CDJ-IMG-08_VESTIARIO-FEM-MASC-TERMAS_D72.jpg', 'D72 - Cuba de semi-encaixe slim, Deca'),
    ('PILOTIS', 'O55', 'CDJ-IMG-09_VESTIARIO-FEM-MASC-TERMAS_D69.jpg', 'D69 - Torneira com sensor Decalux, Deca'),
    ('PILOTIS', 'P55', 'CDJ-IMG-10_VESTIARIO-FEM-MASC-TERMAS_D22.jpg', 'D22 - Bacia Carrara / Panache e assento'),
    ('PILOTIS', 'S55', 'CDJ-IMG-11_VESTIARIO-FEM-MASC-TERMAS_D33.jpg', 'D33 - Monocomando para chuveiro Mix&Match, Docol'),
    ('PILOTIS', 'U55', 'CDJ-IMG-13_VESTIARIO-FEM-MASC-TERMAS_D70.jpg', 'D70 - Chuveiro 200 de parede Docol Eden'),
]

# imagens da web (02/10/2026): lista em especificacoes/imagens/imagens_produtos_web.json,
# arquivos CDJ-WEB-NN_<CODIGO>.jpg. As fotos enviadas pela usuária (LINKS) têm
# prioridade: a célula que já tem uma delas não recebe a imagem da web.
WEB = 'especificacoes/imagens/imagens_produtos_web.json'
ABAS = {a.strip(): a for a in ['TÉRREO ', '2° PAVTO', 'PILOTIS', 'APTO TIPO']}
NCOLS = 20           # colunas de imagem reservadas por linha (AH em diante)
MIN_H_FOTO = 120.0   # altura mínima (pt) das linhas que têm foto, para a foto não ficar minúscula

def legenda_web(p):
    t = f"WEB-{p['n']:02d} - {p['produto']} ({p['marca']})"
    t += '\nImagem de referência da web' + ('' if p.get('confianca') == 'alta' else ' - CONFERIR')
    return t

def entradas():
    """(aba, célula, arquivo, legenda, max_px), na ordem: fotos da usuária, depois web."""
    ent = [(a, r, arq, leg, 1600) for a, r, arq, leg in LINKS]
    usuaria = {(a, r) for a, r, *_ in LINKS}
    for p in json.load(open(WEB, encoding='utf-8')):
        if not p.get('arquivo'):
            continue
        for c in dict.fromkeys(p['celulas']):
            aba, ref = c.split('!')
            aba = ABAS.get(aba.strip(), aba)
            if (aba, ref) not in usuaria:
                ent.append((aba, ref, p['arquivo'], legenda_web(p), 900))
    return ent

def img_bytes(path, max_w=1600):
    im = Image.open(path).convert('RGB')
    if im.width > max_w:
        im = im.resize((max_w, int(im.height * max_w / im.width)))
    b = io.BytesIO(); im.save(b, 'JPEG', quality=85)
    return b.getvalue(), im.size

wb = openpyxl.load_workbook(P)
# guarda os bytes das imagens já existentes (cabeçalhos) antes de mexer
for ws in wb.worksheets:
    for im in ws._images:
        b = im._data(); im._data = (lambda b=b: b)

ENT = entradas()
abas = {a for a, *_ in ENT}
for aba in abas:
    ws = wb[aba]
    ws._images = [im for im in ws._images if im.anchor._from.col < COL0 - 1]   # tira fotos de rodadas anteriores
    for row in ws.iter_rows(min_col=COL0, max_col=COL0 + NCOLS):
        for t in row:
            if t.row > 11:
                t.value = None; t.hyperlink = None
    for c in range(COL0, COL0 + NCOLS):
        ws.column_dimensions[L(c)].width = LARG

fotos = {}       # (aba, linha, arquivo) -> [coluna, legenda, [células de origem], max_px]
proxima = {}
primeira = {}    # (aba, célula) -> célula da primeira foto
for aba, ref, arq, leg, mx in ENT:
    r = wb[aba][ref].row
    k = (aba, r, arq)
    if k not in fotos:
        col = proxima.get((aba, r), COL0); proxima[(aba, r)] = col + 1
        assert col < COL0 + NCOLS, (aba, r)
        fotos[k] = [col, leg, [], mx]
    fotos[k][2].append(ref)
    primeira.setdefault((aba, ref), []).append(f'{L(fotos[k][0])}{r}')

for (aba, ref), alvos in primeira.items():
    ws = wb[aba]; cel = ws[ref]
    v = str(cel.value or '')
    if MARCA not in v:
        cel.value = v.rstrip() + '\n\n' + MARCA
    cel.hyperlink = Hyperlink(ref=ref, location=f"'{aba}'!{alvos[0]}",
                              tooltip=('Ver imagem' + (f' ({len(alvos)} imagens lado a lado)' if len(alvos) > 1 else ''))[:255], display=None)
    f = copy.copy(cel.font); f.u = None; cel.font = f

# a linha "► VER IMAGEM" entra depois de formatacao.py: recalcula a altura das linhas
# com a mesma regra de lá, senão o texto centralizado fica cortado em cima e embaixo
import math
LARG_F = {1: 40.0, 3: 40.0, 4: 40.0, 5: 40.0, 7: 40.0, 8: 70.0, 24: 70.0, 26: 70.0, **{c: 31.0 for c in range(27, 33)}}
def _linhas(txt, col, sz):
    n = max(8, int((LARG_F.get(col, 48.0) * 7.0 - 16) / {11: 6.0, 12: 7.2}[sz]))
    return sum(max(1, math.ceil(len(par) / n)) for par in str(txt).split('\n'))
for aba, r in {(a, wb[a][ref].row) for a, ref in primeira}:
    ws = wb[aba]
    n = max([1] + [_linhas(ws.cell(r, c).value, c, 12 if c == 1 else 11) for c in range(1, 27) if ws.cell(r, c).value not in (None, '')])
    ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 0, min(409, max(60.0, n * 15.0 + 14)))

for (aba, r, arq), (col, leg, refs, mx) in fotos.items():
    ws = wb[aba]
    if (ws.row_dimensions[r].height or 0) < MIN_H_FOTO:
        ws.row_dimensions[r].height = MIN_H_FOTO
    alvo = f'{L(col)}{r}'
    t = ws[alvo]
    t.value = f'{leg}\n◄ VOLTAR para {", ".join(refs)}'
    t.hyperlink = Hyperlink(ref=alvo, location=f"'{aba}'!{refs[0]}", tooltip='Voltar para a grade')
    t.font = Font(name='Aptos Narrow', size=11, bold=True, color='243F2E')
    t.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
    t.fill = PatternFill('solid', fgColor='FFFFFF')
    # foto ajustada à caixa (largura da coluna x altura da linha menos a legenda)
    data, (w, h) = img_bytes(DIR + arq, mx)
    row_px = int((ws.row_dimensions[r].height or 60) * 96 / 72)
    linhas_leg = t.value.count('\n') + 1 + len(t.value) // 70
    box_h = max(80, row_px - 18 * linhas_leg - 14)
    k = min(BOX_W / w, box_h / h)
    dw, dh = int(w * k), int(h * k)
    xi = XLImage(io.BytesIO(data)); xi._data = (lambda b=data: b)
    xi.width, xi.height = dw, dh
    col_px = int(LARG * 7 + 5)
    off_x = max(0, (col_px - dw) // 2)
    xi.anchor = OneCellAnchor(_from=AnchorMarker(col=col - 1, colOff=pixels_to_EMU(off_x), row=r - 1, rowOff=pixels_to_EMU(6)),
                              ext=XDRPositiveSize2D(pixels_to_EMU(dw), pixels_to_EMU(dh)))
    ws.add_image(xi)
print(len(primeira), 'células com link,', len(fotos), 'fotos')
# título da área de imagens
for aba in abas:
    ws = wb[aba]
    h = ws.cell(11, COL0); h.value = 'IMAGENS DAS ESPECIFICAÇÕES (fora da área de impressão)'
    h.font = Font(name='Aptos Narrow', size=20, bold=True, color='FFFFFF'); h.fill = PatternFill('solid', fgColor='243F2E')
    h.alignment = Alignment(horizontal='left', vertical='center')
    rng = f'{L(COL0)}11:{L(COL0 + 9)}11'
    if rng not in [str(m) for m in ws.merged_cells.ranges]:
        ws.merge_cells(rng)

cr = wb['CONTROLE DE REVISÕES ']
for rr in range(12, 40):
    for c in range(1, 10):
        x = cr.cell(rr, c)
        if isinstance(x.value, str) and x.value.strip() in ('IMAGENS ESPECIFICAÇÕES', 'IMAGENS: LINK "VER IMAGEM" NAS CÉLULAS DA GRADE'):
            x.value = 'IMAGENS: AO LADO DA GRADE (LINK "VER IMAGEM")'
        if c == 4 and isinstance(x.value, str) and x.value.strip() == 'IMAGENS ESPECIFICAÇÕES':
            pass
for rr in range(24, 40):
    if str(cr.cell(rr, 4).value or '').strip() in ('IMAGENS ESPECIFICAÇÕES', 'IMAGENS: AO LADO DA GRADE (LINK "VER IMAGEM")'):
        cr.cell(rr, 4).value = 'IMAGENS ESPECIFICAÇÕES'
        cr.cell(rr, 5).value = ('Imagens das fichas TERMAS e VESTIÁRIOS - TERMAS da tabela de acabamentos v00. '
                                'Em 02/10/2026 passaram para a própria aba PILOTIS, à direita da grade (a partir da coluna AH); '
                                'o link "► VER IMAGEM (clique)" na célula leva à foto e "◄ VOLTAR" retorna.')
wb.save(P)
