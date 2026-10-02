"""Terraço do Salão de Festas (Pilotis): piso ○08 da planta de arquitetura PR011 R13.

A extração automática (arq_ambientes.py) não pegou a etiqueta; conferida no PDF em 2026-10-02.
O detalhamento R04 não especifica acabamentos para o terraço (o grupo "VARANDA, TERRAÇO,
PISCINA E PRAINHA, BRASEIRO" da lista R04 só tem códigos sem descrição, #N/D).

Uso: python3 terraco_salao.py <grade.xlsx>
"""
import sys, openpyxl

XLSX = sys.argv[1]
PISO = ('ARQ-P08 - Piso elevado em granito cinza santa rosa, acabamento serrado, Espessura: 2cm. '
        'Dimensões e paginação conforme detalhamento. Amostra a ser aprovada pela GPA&A. '
        'Obs.: Resina, hidrofugante ou outros acabamentos aplicados no granito devem ser foscos e não '
        'podem alterar ou escurecer a cor original do material. (Planta PR011)')
NOTA = ('Tabela v00 e detalhamento R04 sem especificação de acabamento (a lista R04 só traz códigos '
        'sem descrição para "Varanda, terraço, piscina e prainha, braseiro"). A planta de arquitetura '
        'PR011 indica piso ○08 (ARQ-P08, lançado na grade) e guarda-corpos GC06/GC07 h=130cm. '
        'Sem indicação de rodapé, parede ou teto.')

wb = openpyxl.load_workbook(XLSX)
ws = wb['PILOTIS']
r = next(r for r in range(13, ws.max_row + 1) if (ws.cell(r, 1).value or '').strip() == 'TERRAÇO SALÃO DE FESTAS')
ws.cell(r, 2).value = PISO
pen = wb['PENDÊNCIAS']
for i in range(1, pen.max_row + 1):
    if (pen.cell(i, 2).value or '').strip() == 'TERRAÇO SALÃO DE FESTAS' and pen.cell(i, 1).value == 1:
        pen.cell(i, 4).value = NOTA
wb.save(XLSX)
print('PILOTIS B%d' % r)
