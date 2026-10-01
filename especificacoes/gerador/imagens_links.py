"""Link "VER IMAGEM" nas células da grade (no lugar da aba IMAGENS ESPECIFICAÇÕES).

Cada imagem da tabela de acabamentos virou uma página própria (artifact no
claude.ai); a célula que cita o código recebe o hyperlink para essa página e
uma linha final "► VER IMAGEM (clique)". A aba IMAGENS ESPECIFICAÇÕES é removida.
"""
import sys, copy
import openpyxl
from openpyxl.worksheet.hyperlink import Hyperlink

P = sys.argv[1] if len(sys.argv) > 1 else 'especificacoes/CDJ-GRE-ARQ-AC-R00.xlsx'
MARCA = '► VER IMAGEM (clique)'
# (aba, célula) -> (página da imagem, legenda)
LINKS = {
    ('PILOTIS', 'B41'): ('https://claude.ai/artifact/NZBQfeagZDFKMCf9v9pWAm', 'P10 - Piso elevado em mármore branco selecionado'),
    ('PILOTIS', 'H41'): ('https://claude.ai/artifact/2HNnXmHTGT5Mf5nrp6VDAR', 'PA49 - Mosaico Bisazza (painel e pilar)'),
    ('PILOTIS', 'U41'): ('https://claude.ai/artifact/Nsju1SEAeUuHpeqhrho7ZV', 'D54 - Chuveiro de embutir no teto Raindream, Roca'),
    ('PILOTIS', 'B55'): ('https://claude.ai/artifact/BbjJbZZsEcyi4Lh2t9MDEK', 'P10 - Piso elevado em mármore branco selecionado'),
    ('PILOTIS', 'K55'): ('https://claude.ai/artifact/GqTn8Ry8w4KBFqzqfYQqgZ', 'D38 - Espelho liso incolor 4mm'),
    ('PILOTIS', 'J55'): ('https://claude.ai/artifact/Na8iftp8jvWND6WUo2sage', 'D73 - Bancada em mármore branco selecionado'),
    ('PILOTIS', 'L55'): ('https://claude.ai/artifact/HjzFWaHsE3azAbWUpo6LW8', 'D72 - Cuba de semi-encaixe slim, Deca'),
    ('PILOTIS', 'O55'): ('https://claude.ai/artifact/FYgsJpGSehavVp143f8qUW', 'D69 - Torneira com sensor Decalux, Deca'),
    ('PILOTIS', 'P55'): ('https://claude.ai/artifact/WZRA8L1T4HG7XttkqkkPZK', 'D22 - Bacia Carrara / Panache e assento'),
    ('PILOTIS', 'S55'): ('https://claude.ai/artifact/PEVzzaAXCkuSf4egcYg7cb', 'D33 - Monocomando para chuveiro Mix&Match, Docol'),
    ('PILOTIS', 'H55'): ('https://claude.ai/artifact/TPwApMaHcwuUYfYibo2hTh', 'PA47 - Pintura com tinta mineral, cor 930, Terracor'),
    ('PILOTIS', 'U55'): ('https://claude.ai/artifact/19N6usZwRGuEpnNqvBwxCU', 'D70 - Chuveiro 200 de parede Docol Eden'),
}

wb = openpyxl.load_workbook(P)
for (aba, ref), (url, leg) in LINKS.items():
    x = wb[aba][ref]
    v = str(x.value or '')
    if MARCA not in v:
        x.value = v.rstrip() + '\n\n' + MARCA
    x.hyperlink = Hyperlink(ref=ref, target=url, tooltip=('Abrir imagem: ' + leg)[:255])
    f = copy.copy(x.font); f.u = None; x.font = f   # mantém a fonte/cor da célula
    print(aba, ref, leg)
if 'IMAGENS ESPECIFICAÇÕES ' in wb.sheetnames:
    del wb['IMAGENS ESPECIFICAÇÕES ']
    print('aba IMAGENS ESPECIFICAÇÕES removida')
cr = wb['CONTROLE DE REVISÕES ']
for r in range(12, 40):
    for c in range(1, 10):
        x = cr.cell(r, c)
        if isinstance(x.value, str) and x.value.strip() == 'IMAGENS ESPECIFICAÇÕES':
            if c == 4:   # linha do controle de revisões
                cr.cell(r, 5).value = ('Imagens das fichas TERMAS e VESTIÁRIOS - TERMAS da tabela de acabamentos v00. '
                                       'Em 01/10/2026 a aba de imagens foi substituída pelo link "► VER IMAGEM (clique)" '
                                       'nas próprias células da grade.')
            else:        # sumário
                x.value = 'IMAGENS: LINK "VER IMAGEM" NAS CÉLULAS DA GRADE'
wb.save(P)
