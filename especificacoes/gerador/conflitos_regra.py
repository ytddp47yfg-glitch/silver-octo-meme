"""Regra de conflito (R00): o detalhamento prevalece sobre a Tabela R01 e só é
conflito a divergência entre ARQUITETURA e DETALHAMENTO.

- Tabela R01 x detalhamento: vale o detalhamento; o texto da tabela sai da célula.
- Prancha x tabela/quadro da própria prancha (arquitetura ou detalhamento): CONFLITO.
- Arquitetura x detalhamento: CONFLITO (vermelho, as duas versões na célula).

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

# 2) prancha x tabela/quadro da própria prancha (arquitetura ou detalhamento): CONFLITO em vermelho
#    (regra da usuária, 2026-10-02). A lista de especificações R04 conta como tabela do detalhamento.
FILL = PatternFill('solid', fgColor='FFFFC7CE')
def conflito(c):
    c.fill = FILL
    f = c.font
    c.font = Font(name=f.name, size=f.sz, bold=f.b, italic=f.i, color='FF9C0006')

TROCAS = [
 (r'(CONFLITO|OBS\.) – (A legenda|Legenda) de rodapé do DET PR001/PR002[^\n]*',
  'CONFLITO – Legenda de rodapé do DET PR001/PR002: h=6,5cm (o quadro da mesma prancha indica h=10cm)'),
 (r'(CONFLITO|OBS\.) – (As pranchas|Pranchas) 37 e 38 do detalhamento R04( indicam|:) ?([^\n]*?)( \(ver PENDÊNCIAS\)\.)?$',
  r'CONFLITO – Pranchas 37 e 38 do detalhamento R04: \4'),
 (r'^ARQ-T19 - código não existe no quadro de tetos \(01 a 06\); conferir na planta$',
  'ARQ-T19 - código indicado na etiqueta de teto da planta PR003.\n\nCONFLITO – Quadro de tetos da própria PR003: só traz os códigos 01 a 06; o T19 não existe no quadro.'),
]
for ws, refs in ((apto, ('F17',)), (aba(wb, 'PILOTIS'), ('J35', 'O35', 'J36', 'O36')), (aba(wb, 'TÉRREO'), ('I30',))):
    for ref in refs:
        c = ws[ref]
        v = c.value or ''
        for pat, rep in TROCAS:
            v = re.sub(pat, rep, v, flags=re.M)
        c.value = v
        conflito(c)

# 3) PENDÊNCIAS: notas antigas reescritas
pen = wb['PENDÊNCIAS']
for row in pen.iter_rows():
    for c in row:
        v = c.value
        if not isinstance(v, str): continue
        if 'CONFLITO – Tabela R01:' in v:
            c.value = v.replace('CONFLITO – Tabela R01:', 'Tabela R01 divergente (prevalece o detalhamento):')
        elif v.startswith('RODAPÉ: CONFLITO – Legenda de rodapé') or v.startswith('RODAPÉ: divergência interna'):
            c.value = ('RODAPÉ: CONFLITO na própria prancha (ver aba CONFLITOS). O quadro do DET PR001 indica h=10cm; '
                       'a legenda de rodapé do DET PR001/PR002 indica h=6,5cm.')
        elif v.startswith('Pranchas 37 e 38 citam torneira Argon') or v.startswith('Divergência interna do detalhamento R04'):
            c.value = ('CONFLITO no detalhamento R04 (ver aba CONFLITOS): as pranchas 37 e 38 citam torneira Argon (Docol) '
                       'e bancada em granito Branco Siena com duas cubas esculpidas; a lista R04 traz D29 '
                       '(Decalux, Deca) e D36 (granito cinza Santa Rosa).')
        elif v.startswith('A etiqueta de teto da planta PR003 traz'):
            c.value = ('CONFLITO na própria planta (ver aba CONFLITOS): a etiqueta de teto da PR003 traz □19, mas o quadro '
                       'de tetos da prancha só vai de 01 a 06. Confirmar com a arquitetura.')

# o rastreamento antigo marcava "CONFLITO:" em linhas que seguem válidas (Guarita)
wb.save(XLSX)
print('ok')
