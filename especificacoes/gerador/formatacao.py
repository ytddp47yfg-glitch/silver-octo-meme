"""Formatação uniforme das abas de pavimento da grade CDJ-GRE-ARQ-AC.

- larguras de coluna iguais ao modelo em todas as abas;
- fonte das células de dados (ambiente 12 negrito, demais 11) e alinhamento;
  (com 16 os textos longos passavam do limite de 409 pt de altura de linha);
- altura de cada linha calculada pelo texto, com mínimo comum;
- impressão como no modelo: A3 paisagem, mesma escala, uma folha por grupo
  (Revestimentos B:K, Louças e metais L:U, Complementares V:AF), com as linhas
  1 a 11 e a coluna A repetidas.
Rodar area_impressao.py depois.
"""
import sys, copy, math
import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.pagebreak import Break, ColBreak
from openpyxl.styles import Alignment, PatternFill

P = sys.argv[1] if len(sys.argv) > 1 else 'especificacoes/CDJ-GRE-ARQ-AC-R00.xlsx'
ABAS = ['TÉRREO ', '2° PAVTO', 'PILOTIS', 'APTO TIPO']
# larguras por coluna: textos longos (parede, portas, outros complementares) mais largos,
# itens curtos (soleira, filete, peitoril, sóculo) mais estreitos; as três faixas
# de impressão (A+B:K, A+L:U, A+V:AF) ficam com ~510 caracteres de largura
LARG = {1: 40.0, 3: 40.0, 4: 40.0, 5: 40.0, 7: 40.0, 8: 70.0, 24: 70.0, 26: 70.0,
        **{c: 31.0 for c in range(27, 33)}}
PADRAO = 48.0
ESCALA = 38          # mesma escala do modelo; maior faixa (A + B:K) ~520 caracteres de largura
MIN_H = 60.0
LINHA = 15.0          # altura de uma linha de texto em Aptos Narrow 11
PX_CHAR = 7.0         # px por unidade de largura de coluna
PX_LETRA = {11: 6.0, 12: 7.2}  # largura média de uma letra (px)

def cpl(col, sz):
    w = LARG.get(col, PADRAO) * PX_CHAR - 16
    return max(8, int(w / PX_LETRA[sz]))

def linhas(txt, n):
    t = 0
    for par in str(txt).split('\n'):
        t += max(1, math.ceil(len(par) / n))
    return t

wb = openpyxl.load_workbook(P)
for nome in ABAS:
    ws = wb[nome]
    ws.column_dimensions.clear()
    for c in range(1, 33):
        ws.column_dimensions[L(c)] = ColumnDimension(ws, index=L(c), width=LARG.get(c, PADRAO), customWidth=True)
    last = max(r for r in range(13, ws.max_row + 1) if ws.cell(r, 1).value)
    for r in range(13, last + 1):
        n = 1
        for c in range(1, 27):
            x = ws.cell(r, c)
            f = copy.copy(x.font)
            f.name = 'Aptos Narrow'; f.sz = 12 if c == 1 else 11; f.b = (c == 1)
            # primeira versão (R00): sem marcação de INFO. REVISADA (laranja); conflitos (vermelho) ficam
            if x.fill is not None and x.fill.fill_type and x.fill.fgColor.type == 'theme' and x.fill.fgColor.theme == 5:
                x.fill = PatternFill(fill_type=None)
                f.color = None
            x.font = f
            x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if x.value not in (None, ''):
                n = max(n, linhas(x.value, cpl(c, 12 if c == 1 else 11)))
        ws.row_dimensions[r].height = min(409, max(MIN_H, n * LINHA + 14))
    # impressão
    ps = ws.page_setup
    # linhas vazias abaixo dos dados: sem preenchimento laranja herdado
    for r in range(last + 1, ws.max_row + 1):
        for c in range(1, 33):
            x = ws.cell(r, c)
            if x.fill is not None and x.fill.fill_type and x.fill.fgColor.type == 'theme' and x.fill.fgColor.theme == 5:
                x.fill = PatternFill(fill_type=None)
    # escala fixa: o "ajustar a N páginas" do Excel ignora as quebras manuais
    ps.paperSize = ws.PAPERSIZE_A3; ps.orientation = 'landscape'; ps.scale = ESCALA
    ps.fitToWidth = None; ps.fitToHeight = None
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.print_options.horizontalCentered = True
    m = ws.page_margins; m.left = m.right = 0.2; m.top = m.bottom = 0.3; m.header = m.footer = 0.1
    ws.col_breaks.brk = []
    for b in (11, 21):
        ws.col_breaks.append(Break(id=b))
    ws.print_title_rows = '1:11'; ws.print_title_cols = 'A:A'
    alto = [r for r in range(13, last + 1) if ws.row_dimensions[r].height >= 409]
    print(nome, 'linhas no limite', alto, 'linhas', 13, '-', last, 'maior altura', max(ws.row_dimensions[r].height for r in range(13, last + 1)))
# IMAGENS: volta ao layout do modelo (escala fixa, quebra na coluna H)
for wi in [wb[n] for n in wb.sheetnames if n.startswith('IMAGENS')]:
  wi.sheet_properties.pageSetUpPr.fitToPage = False
  wi.page_setup.scale = ESCALA; wi.page_setup.fitToWidth = None; wi.page_setup.fitToHeight = None
  wi.print_options.horizontalCentered = True
# traços soltos do modelo abaixo dos dados criavam linhas/folhas em branco na impressão
for ws in wb.worksheets[1:6]:
    for row in ws.iter_rows(min_row=13):
        for x in row:
            if x.value == '-' and not any(ws.cell(x.row, c).value not in (None, '', '-') for c in range(1, 33)):
                print('limpo', ws.title, x.coordinate); x.value = None
# planilha sem linhas de grade
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
# mesmo rodapé (Página X de Y) em todas as abas
rod = copy.copy(wb['TÉRREO '].oddFooter.left)
for ws in wb.worksheets:
    if not ws.oddFooter.left.text:
        ws.oddFooter.left = copy.copy(rod)
wb.save(P)
