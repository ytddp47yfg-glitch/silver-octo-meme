"""Regra de conflito (R00): o detalhamento prevalece sobre a Tabela R01 e só é
conflito a divergência entre ARQUITETURA e DETALHAMENTO.

- Tabela R01 x detalhamento: vale o detalhamento; o texto da tabela sai da célula.
- Divergência interna do detalhamento (lista x pranchas, quadro x legenda): não é
  conflito; mantém o texto principal, a outra indicação vira "OBS." e vai para PENDÊNCIAS.
- Arquitetura x detalhamento: continua em vermelho.

Uso: python3 conflitos_regra.py <grade.xlsx>
"""
import sys, re
import openpyxl
from openpyxl.styles import PatternFill, Font

XLSX = sys.argv[1]
SEM = PatternFill(fill_type=None)

def aba(wb, nome):
    return next(wb[n] for n in wb.sheetnames if n.strip() == nome)

def normal(c):
    c.fill = SEM
    f = c.font
    c.font = Font(name=f.name, size=f.sz, bold=f.b, italic=f.i)

wb = openpyxl.load_workbook(XLSX)

# 1) Tabela R01 x detalhamento: prevalece o detalhamento
apto = aba(wb, 'APTO TIPO')
for ref in ('F13', 'F14', 'H15', 'J18'):
    c = apto[ref]
    c.value = re.sub(r'\n*CONFLITO – Tabela R01:[^\n]*', '', c.value or '').strip()
    normal(c)

# 2) divergências internas do detalhamento: OBS. sem vermelho
for ws, refs in ((apto, ('F17',)), (aba(wb, 'PILOTIS'), ('J35', 'O35', 'J36', 'O36'))):
    for ref in refs:
        c = ws[ref]
        v = c.value or ''
        v = v.replace('CONFLITO – Legenda de rodapé do DET PR001/PR002: h=6,5cm',
                      'OBS. – A legenda de rodapé do DET PR001/PR002 indica h=6,5cm (ver PENDÊNCIAS).')
        v = v.replace('CONFLITO – Pranchas 37 e 38 do detalhamento R04:',
                      'OBS. – As pranchas 37 e 38 do detalhamento R04 indicam')
        if 'Pranchas 37' in (c.value or '') and not v.endswith('(ver PENDÊNCIAS).'):
            v += ' (ver PENDÊNCIAS).'
        c.value = v
        normal(c)

# 3) PENDÊNCIAS: notas antigas reescritas
pen = wb['PENDÊNCIAS']
for row in pen.iter_rows():
    for c in row:
        v = c.value
        if not isinstance(v, str): continue
        if 'CONFLITO – Tabela R01:' in v:
            c.value = v.replace('CONFLITO – Tabela R01:', 'Tabela R01 divergente (prevalece o detalhamento):')
        elif 'CONFLITO – Legenda de rodapé' in v:
            c.value = ('RODAPÉ: divergência interna do detalhamento. O quadro do DET PR001 (e a Tabela R01) '
                       'indicam h=10cm; a legenda de rodapé do DET PR001/PR002 indica h=6,5cm. '
                       'Mantido h=10cm na grade. Confirmar.')
        elif v.startswith('Pranchas 37 e 38 citam torneira Argon'):
            c.value = ('Divergência interna do detalhamento R04: as pranchas 37 e 38 citam torneira Argon (Docol) '
                       'e bancada em granito Branco Siena com duas cubas esculpidas; a lista R04 traz D29 '
                       '(Decalux, Deca) e D36 (granito cinza Santa Rosa). Mantida a lista na grade, com OBS. Confirmar.')

# o rastreamento antigo marcava "CONFLITO:" em linhas que seguem válidas (Guarita)
wb.save(XLSX)
print('ok')
