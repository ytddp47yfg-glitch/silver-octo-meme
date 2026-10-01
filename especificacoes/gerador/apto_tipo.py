"""Cria a aba APTO TIPO na grade com a tabela JARDIM-ACABAMENTOS R01 e o detalhamento CDJ_DET_EX PR001-PR008 /
CDJ-ARQ-DT PR406-407 e PR721-728 (circulações dos andares). Conflitos: as duas versões na célula, em vermelho claro.

Uso: python3 apto_tipo.py <grade.xlsx>
"""
import sys, copy
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.drawing.image import Image as XLImage

XLSX = sys.argv[1]
T = 'TAB R01'          # tabela JARDIM-ACABAMENTOS_R01.xlsx
CONF = 'CONFLITO'

# ambiente -> {coluna: [textos]} ; texto iniciado por CONFLITO marca a célula
A = {
 'SALA / VARANDA DA SALA': {
  'B': ['Piso em mármore travertino romano bruto, 85x130cm, paginação conforme detalhamento. Rejunte Infinity Hard, cor Areia ou Palha (validar com o Setor de Produto). (DET PR001/PR002; Tab. R01)'],
  'F': ['Rodapé em mármore travertino romano bruto, semi-embutido, h=15cm. (DET PR001/PR002)',
        CONF + ' – Tabela R01: rodapé DE SOBREPOR em mármore travertino romano bruto, h=15cm'],
  'H': ['Pintura acrílica cor branco neve, acabamento acetinado. (DET PR001; Tab. R01)'],
  'T': ['Varanda: ralo linear oculto; caimento do piso para o ralo i=1%. (DET PR001)'],
  'Z': ['Varanda: mureta h=90cm. (DET PR001)'],
 },
 'LAVABO': {
  'B': ['Piso em mármore travertino romano bruto. Rejunte Infinity Hard, cor Areia ou Palha (validar com o Setor de Produto). (DET PR004; Tab. R01)'],
  'C': ['Soleira indicada em planta, material não especificado. (DET PR004)'],
  'F': ['Rodapé semi-embutido em mármore travertino romano bruto, h=15cm. (DET PR004)',
        CONF + ' – Tabela R01: rodapé DE SOBREPOR em mármore travertino romano bruto, h=15cm'],
  'H': ['Pintura acrílica cor branco neve, acabamento acetinado. (DET PR004; Tab. R01)'],
  'I': ['Forro em gesso acartonado, pintura acrílica acetinada cor branco neve, tabica metálica. PD 298cm. (DET PR004)'],
  'J': ['Bancada esculpida em mármore travertino romano resinado (cuba esculpida). (DET PR004; Tab. R01)'],
  'O': ['Torneira de mesa bica baixa para lavatório, cromado, linha Unic, cód. 1197.C90. Ref.: Deca. (DET PR004; Tab. R01)'],
  'P': ['Bacia convencional suspensa branca, linha LK, cód. P.232.17. Ref.: Deca. Caixa de descarga embutida para alvenaria. (DET PR004; Tab. R01)'],
  'R': ['Acabamento para registro de gaveta 1 1/4" e 1 1/2", cromado, Unic, cód. 4900.C90.GD. Ref.: Deca', 'Acabamento para caixa de descarga embutida, cromado, H Quadra Duo, cód. 4900.C.HQD.DUO. Ref.: Deca. (DET PR004)'],
  'T': ['Ralo oculto 15x15cm, tampa revestida com o mesmo mármore. (DET PR004)'],
  'W': ['Maçaneta Pado Sara, cromo acetinado, interna. (DET PR004)'],
  'X': ['Porta piso-forro, cor branca, alizar 10cm. (DET PR004)'],
 },
 'BANHO SUÍTE MASTER': {
  'B': ['11 - Piso em contrapiso regularizado, revestido em mármore travertino romano resinado. Rejunte creme ou marfim (ver padrão da chapa). Paginação conforme detalhamento. Amostra a ser aprovada pela GPA&A. (PR408; Tab. R01: "bruto resinado")'],
  'E': ['Peitoril a definir conforme projeto de esquadrias. (PR408)'],
  'F': ['23 - Rodapé em mármore travertino romano bruto semi-embutido, h=15cm; paginação segue alinhamento do piso. (PR408)'],
  'H': ['14 - Parede revestida em mármore travertino romano resinado. Paginação conforme detalhamento. Amostra a ser aprovada pela GPA&A. (PR408)',
        '1 - Parede em alvenaria emassada e regularizada, pintada com tinta acrílica cor branco neve, acabamento acetinado. Ref.: Suvinil. (PR408)',
        CONF + ' – Tabela R01: parede em mármore travertino romano BRUTO paginado'],
  'I': ['1 - Forro em gesso acartonado liso, tinta acrílica cor branco neve, acabamento fosco, junta de dilatação metálica branca. Ref.: Suvinil. Sanca (DET 20). (PR408)'],
  'J': ['35 - Bancada em mármore travertino romano resinado. Rejunte creme ou marfim (ver padrão da chapa). Dimensões conforme detalhamento. (PR408; Tab. R01)'],
  'L': ['33 - Cuba de apoio retangular Jader Almeida 670x495x120mm, branca, cód. L.310.17. Ref.: Deca. (PR408; Tab. R01)'],
  'O': ['36 - Misturador monocomando de mesa bica alta para lavatório, Level cromado, cód. 2885.C26. Ref.: Deca. (PR408; Tab. R01)'],
  'P': ['34 - Bacia convencional Carrara branca com caixa acoplada e acionamento duo, cód. P.606.17. Ref.: Deca. (PR408; Tab. R01)', '46 - Bidê 3 furos Carrara branco, cód. B.60.17. Ref.: Deca. (PR408; Tab. R01)'],
  'Q': ['37 - Ducha higiênica Level cromada, cód. 1984.C26.ACT.CR. (PR408)'],
  'R': ['49 - Acabamento para registro Level cromado, cód. 4900.C26.PQ. Ref.: Deca. (PR408)'],
  'S': ['44 - Acabamento de monocomando de chuveiro para alta e baixa pressão, Level cromado, cód. 4993.C26.CHU. (PR408; Tab. R01)'],
  'T': ['47 - Ralo linear oculto Royal 70cm', '48 - Ralo oculto 15x15cm Square. (PR408)'],
  'U': ['45 - Misturador para bidê Level cromado, cód. 1895.C26. (PR408; Tab. R01)'],
  'V': ['39 - Tomadas e interruptores Arteor quadrado branco. Ref.: Legrand. (PR408)'],
  'W': ['Maçaneta Sara cromo acetinado. Ref.: Pado. (PR408)'],
  'X': ['38 - Porta em marcenaria piso-forro, cor branca, alizar 10cm. (PR408)'],
 },
 'BANHO SUÍTES': {
  'B': ['Piso em mármore travertino romano bruto resinado, paginado. Rejunte Infinity Hard, cor Areia ou Palha (validar com o Setor de Produto). Desnível do box 1cm. (DET PR008; Tab. R01)'],
  'H': ['Parede em mármore travertino romano bruto: piso ao teto na área do box e barrado semi-embutido até 110cm nas demais paredes; acima, pintura acrílica cor branco neve, acetinado. (DET PR008; Tab. R01)'],
  'I': ['Forro em gesso acartonado, pintura acrílica acetinada cor branco neve, tabica metálica e sanca. PD 298cm. (DET PR008)'],
  'J': ['Bancada esculpida em mármore travertino romano resinado. (DET PR008; Tab. R01)'],
  'L': ['Cuba retangular de apoio 60cm com deck, branca, linha Slim, cód. L.13060.M.17. Ref.: Deca. (DET PR008; Tab. R01)'],
  'O': ['Misturador monocomando de mesa bica baixa para lavatório, cromado, linha Level, cód. 2875.C26. Ref.: Deca. (DET PR008; Tab. R01)'],
  'P': ['Bacia convencional Carrara branca com caixa acoplada, cód. KP.606.17. Ref.: Deca. (DET PR008; Tab. R01)'],
  'Q': ['Ducha higiênica com registro e derivação, cromada, linha Level. Ref.: Deca. (DET PR008; Tab. R01)'],
  'R': ['Acabamento para registro de gaveta 1 1/4" e 1 1/2", cromado, Level, cód. 4900.C26.GD. Ref.: Deca. (DET PR008)'],
  'S': ['Acabamento monocomando para chuveiro, cromado, linha Level, cód. 4993.C26.CHU. Ref.: Deca. (DET PR008; Tab. R01)'],
  'T': ['Ralo oculto 15x15cm e ralo linear oculto 90x5cm, tampas revestidas com a mesma pedra. (DET PR008)'],
 },
 'SUÍTES / QUARTOS / CIRCULAÇÃO / CLOSET / ESTAR ÍNTIMO': {
  'B': ['Piso em cumaru, assoalho piso pronto. Ref.: Indusparquet. (DET PR001/PR002; Tab. R01)'],
  'D': ['Transição madeira x pedra: perfil metálico (DET 05) ou filete de mármore (DET 07). (DET PR001)'],
  'F': ['Rodapé maciço com o mesmo acabamento do piso, h=10cm. Ref.: Indusparquet. (Tab. R01; quadro do DET PR001)',
        CONF + ' – Legenda de rodapé do DET PR001/PR002: h=6,5cm'],
  'H': ['Pintura acrílica cor branco neve. (Tab. R01)'],
 },
 'COPA NOTURNA': {
  'B': ['Piso em cumaru, assoalho piso pronto. Ref.: Indusparquet. (Tab. R01)'],
  'F': ['Rodapé maciço com o mesmo acabamento do piso, h=10cm. Ref.: Indusparquet. (Tab. R01)'],
  'J': ['Bancada esculpida em mármore travertino romano resinado. (DET PR003)',
        CONF + ' – Tabela R01: bancada "pintura acrílica, cor branco neve"'],
  'L': ['Cuba redonda CR-30, escovada, cód. 01018116 / 90010181016. Ref.: Mekal. (DET PR003; Tab. R01)'],
  'O': ['Torneira de mesa bica alta para lavatório, cromada, linha Link, cód. 1198.C.LNK. Ref.: Deca. (DET PR003; Tab. R01)'],
 },
 'COZINHA / DESPENSAS': {
  'B': ['Piso em granito branco siena escovado, 85x130cm. Rejunte Infinity Hard, cor Branco Pérola ou Palha (validar com o Setor de Produto). (DET PR001, PR406/407; Tab. R01: paginação a definir)'],
  'C': ['Soleira em granito branco siena. (PR406)'],
  'F': ['Rodapé semi-embutido em granito branco siena escovado, h=15cm (DET 22). (DET PR001, PR406/407; Tab. R01)'],
  'H': ['Rodabanca em granito branco siena escovado, h=60cm', 'Pintura acrílica; cor a definir (Tab. R01). (PR406/407)'],
  'J': ['Bancada e ilha em granito branco siena escovado; parte seca h=90cm, área molhada com desnível 1cm h=92cm; suporte em metalon chumbado na parede; emendas seguem a paginação do piso. (PR406/407; Tab. R01)'],
  'L': ['Cuba de embutir 600x400x230mm em aço inox, linha Retta. Ref.: Mekal ou similar. (Tab. R01; PR406 sem modelo)'],
  'O': ['Monocomando de mesa para cozinha, cromado, linha Level. Ref.: Deca ou similar. (Tab. R01; PR406 sem modelo)'],
  'X': ['Porta em marcenaria. (PR406/407)'],
 },
 'ÁREA DE SERVIÇO / CIRCULAÇÃO SERVIÇO': {
  'B': ['Piso em granito branco siena escovado, 85x130cm, ver paginação. Rejunte Infinity Hard, cor Branco Pérola ou Palha. (DET PR001, PR005; Tab. R01)'],
  'C': ['Soleira indicada em planta, material não especificado. (DET PR005)'],
  'F': ['Rodapé semi-embutido em granito branco siena escovado, h=15cm. (DET PR005)'],
  'H': ['Parede em Formica Branco Real L515, 308x125cm. (DET PR005/PR006; Tab. R01)'],
  'I': ['Forro de gesso com tabica metálica (DET 11) e cortineiro 25x16cm (DET 24). (DET PR006)'],
  'J': ['Bancada em granito branco siena escovado, molhada com 1cm de rebaixo, rodabanca h=60cm com topo polido; bancada seca com laterais polidas. (DET PR005/PR006; Tab. R01)'],
  'L': ['Tanque para lavanderia CT-60, escovado, cód. 01019616 / 90010196016, 2 unidades, instalados sob a pedra da bancada. Ref.: Mekal. (DET PR005; Tab. R01)'],
  'O': ['Torneira de parede com arejador para cozinha, cromada, linha Link, cód. 1159.C.LNK, 2 unidades. Ref.: Deca. (DET PR005; Tab. R01)'],
  'R': ['Acabamento para registro de gaveta 1 1/4" e 1 1/2", cromado, Link, cód. 4900.C.GD.LNK. Ref.: Deca. (DET PR005)'],
  'T': ['Ralo oculto 15x15cm, tampa revestida com a mesma pedra (o DET diz "mármore"; o piso é granito). (DET PR005)'],
  'W': ['Maçaneta Pado Sara, cromo acetinado, interna. Porta de segurança: fechadura digital FDE-600W. (DET PR005/PR006)'],
  'X': ['Porta piso-forro, cor branca, alizar 10cm', 'Porta piso-forro especial de segurança, cor branca, alizar 10cm. (DET PR006)'],
 },
 'BANHO SERVIÇO': {
  'B': ['Piso em granito branco siena escovado. (DET PR007; Tab. R01: paginação a definir)'],
  'D': ['Filete em granito na transição para o vinílico (DET 04). (DET PR007)'],
  'H': ['Revestimento Forma Branco AC retificado 32,5x59cm. Ref.: Eliane. (DET PR007; Tab. R01)'],
  'I': ['Forro em gesso acartonado, pintura acrílica acetinada cor branco neve, tabica metálica, sanca 25x16cm. PD 298cm. (DET PR007)'],
  'J': ['Bancada em granito branco siena escovado. (DET PR007; Tab. R01)'],
  'L': ['Cuba redonda de embutir 31cm branca, linha L, cód. L.41.17. Ref.: Deca. (DET PR007; Tab. R01)'],
  'O': ['Torneira de mesa bica baixa para lavatório, cromada, linha Link, cód. 1197.C.LNK. Ref.: Deca. (DET PR007; Tab. R01)'],
  'P': ['Bacia com caixa acoplada branca, linha Axis, cód. KP.470.17. Ref.: Deca. (DET PR007; Tab. R01)'],
  'Q': ['Ducha higiênica com registro e derivação, cromada, linha Link, cód. 1984.C.ACT.LNK.BR. Ref.: Deca. (DET PR007)'],
  'R': ['Acabamento para registro de gaveta e pressão até 1", cromado, Link, cód. 4900.C.PQ.LNK. Ref.: Deca. (DET PR007; Tab. R01)'],
  'T': ['Ralo linear oculto e ralo quadrado oculto. (DET PR007)'],
 },
 'QUARTO / ESTAR FUNCIONÁRIOS': {
  'B': ['Piso vinílico Austin, linha Urban, 18x121cm. Ref.: Durafloor. (DET PR001; Tab. R01)'],
  'D': ['Filete em granito na transição vinílico x granito (DET 04). (DET PR001)'],
  'F': ['Rodapé Easy Y.01 120. Ref.: Durafloor. (DET PR001; Tab. R01)'],
  'H': ['Pintura acrílica cor branco neve. (Tab. R01)'],
 },
 'VARANDA TÉCNICA': {
  'B': ['Piso em concreto. (DET PR001)'],
 },
 'HALL ELEVADOR SOCIAL (ANDAR)': {
  'B': ['Piso em mármore travertino romano bruto. (DET PR003)'],
  'F': ['Rodapé em mármore travertino romano bruto, semi-embutido, h=15cm. (DET PR003)'],
  'H': ['Parede pintada com tinta acrílica cor branco neve, acabamento acetinado', 'Aduela do elevador (2cm) em mármore travertino romano bruto. (DET PR003)'],
 },
 'HALL ELEVADOR SERVIÇO (ANDAR)': {
  'B': ['Piso em granito branco siena escovado. (DET PR003)'],
  'F': ['Rodapé em granito branco siena escovado, semi-embutido, h=15cm. (DET PR003)'],
  'H': ['Parede pintada com tinta acrílica cor branco neve, acabamento acetinado', 'Aduela do elevador (2cm) em granito branco siena escovado. (DET PR003)'],
 },
 'HALL SERVIÇO / ANTECÂMARA – GARDEN, TIPO E COBERTURA': {
  'B': ['Piso em granito ("revestimento conforme ambiente"). (PR721-728)'],
  'F': ['Rodapé em granito semi-embutido, seguir alinhamento do piso (DET 22). (PR722/724/726/728)'],
  'H': ['Parede pintada. (PR722/724/726/728)'],
  'I': ['Forro em gesso com sanca: PD 259cm (Garden e Tipo), 267,5cm (Cobertura 1º nível), 272,5cm (Cobertura 2º nível); junta de dilatação em perfil metálico "Z" (DET 09); trechos em laje de concreto aparente. (PR721-728)'],
  'X': ['PCF1, PCF2, PE-A, P13, P23, PA05, PCF3, P15A, P15C (vão 90x240cm); material não especificado nas pranchas de circulação. (PR721-728)'],
  'Z': ['Sirene de alarme e botoeira; hidrante nos halls conforme projeto PCI. (PR722-728)'],
 },
}

wb = openpyxl.load_workbook(XLSX)
src = wb['2° PAVTO']
if 'APTO TIPO' in wb.sheetnames: del wb['APTO TIPO']
ws = wb.copy_worksheet(src); ws.title = 'APTO TIPO'
wb.move_sheet(ws, offset=-(len(wb.sheetnames) - 1 - wb.sheetnames.index('PILOTIS')) + 1)
import io
for img in src._images:
    b = img._data(); img._data = (lambda b=b: b)
    ni = XLImage(io.BytesIO(b)); ni._data = (lambda b=b: b)
    ni.anchor = copy.deepcopy(img.anchor); ni.width, ni.height = img.width, img.height
    ws.add_image(ni)
for r in range(1, 10):
    for c in ws[r]:
        if isinstance(c.value, str) and '2°' in c.value: c.value = c.value.replace('2° PAVIMENTO', 'APARTAMENTO TIPO')
# limpa linhas do 2º pav
last = max(r for r in range(13, ws.max_row + 1) if ws.cell(r, 1).value)
modelo = 14
for r in range(13, max(last, 13 + len(A)) + 1):
    for c in range(1, 27):
        ws.cell(r, c).value = None
        ws.cell(r, c)._style = copy.copy(src.cell(modelo, c)._style)
ref = wb['PILOTIS']['K7']
REV_FILL, REV_COR = copy.copy(ref.fill), copy.copy(ref.font.color)
CONF_FILL = PatternFill('solid', fgColor='FFFFC7CE')
log = []
for i, (amb, cols) in enumerate(A.items()):
    r = 13 + i
    ws.cell(r, 1).value = amb
    ws.row_dimensions[r].height = None
    for col, vs in cols.items():
        cel = ws[f'{col}{r}']
        cel.value = '\n\n'.join(vs)
        if any(v.startswith(CONF) for v in vs):
            cel.fill = CONF_FILL; f = copy.copy(cel.font); f.color = 'FF9C0006'; cel.font = f
            log.append((amb, ' '.join(ws[f'{col}11'].value.split()), vs[-1]))
ws['B1'].value = 'GRADE DE ESPECIFICAÇÕES – APARTAMENTO TIPO\nEDIFÍCIO JARDIM'
for r in range(1, 13):
    for c in ws[r]:
        if isinstance(c.value, str) and c.value.startswith('GERAL —'):
            c.value = 'GERAL — APARTAMENTO TIPO: TABELA JARDIM-ACABAMENTOS R01 E DETALHAMENTO CDJ_DET_EX PR001-PR008, CDJ-ARQ-DT PR406/407 E PR721-728'
p = wb['PENDÊNCIAS']
r = p.max_row + 2
p.cell(r, 1, '6. APARTAMENTO TIPO (pasta DTL - DETALHAMENTO / APARTAMENTOS)').font = Font(bold=True, size=12)
itens = [(a, f'{c}: {t}') for a, c, t in log] + [
 ('BANHO SUÍTE MASTER', 'Prancha CDJ-ARQ-DT-PR408-ISSM_TIPO-R05 lida pelo PDF anexado (só imagem vetorial, sem texto); códigos da legenda da prancha.'),
 ('ÁREA DE SERVIÇO', 'Ralo com "tampa revestida com o mesmo mármore", mas o piso é granito branco siena. Confirmar.'),
 ('SOLEIRAS', 'Indicadas em planta no lavabo e na área de serviço sem material; só a cozinha define (granito branco siena).'),
 ('COZINHA', 'PR406/407 não trazem modelo de cuba e monocomando; valem os da tabela R01 (Mekal Retta e Deca Level).'),
 ('REJUNTES', 'DET indica "validar com o Setor de Produto"; PR001/PR002 "não liberado para execução, paginação poderá sofrer alterações".'),
]
for a, t in itens:
    r += 1; p.cell(r, 2, a); p.cell(r, 4, t)
wb.save(XLSX)
print(len(A), 'ambientes;', len(log), 'conflitos;', len(ws._images), 'imagens na aba')
