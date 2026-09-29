"""Dados do painel GANTT a partir da aba CRONOGRAMA ATIVIDADES (export do Prevision) da ferramenta.

Uso: python3 gantt.py arquivo_calc.xlsx saida.json

Árvore: PL (coluna M da aba, código do VALOR POR PL) → pacote de trabalho → lote/pavimento (linha-resumo do
Prevision, serviço '-') → tarefas (demais linhas do mesmo pacote e lote). Pais: início = menor início, término =
maior término, % = média dos filhos ponderada pela duração.
"""
import sys, json, datetime as dt
from openpyxl import load_workbook


def main():
    arq, out = sys.argv[1], sys.argv[2]
    wb = load_workbook(arq, data_only=True)
    ws = wb['CRONOGRAMA ATIVIDADES']; vp = wb['VALOR POR PL']
    nome_pl = {str(vp.cell(r, 3).value).strip(): vp.cell(r, 4).value for r in range(11, 49) if vp.cell(r, 3).value}
    ref = None
    rows = []
    for r in range(5, 6005):
        pac = ws.cell(r, 3).value
        if not pac: continue
        ini, fim = ws.cell(r, 6).value, ws.cell(r, 7).value
        if not isinstance(ini, dt.datetime) or not isinstance(fim, dt.datetime): continue
        ref = ref or ws.cell(r, 8).value
        num = lambda c: float(ws.cell(r, c).value) if isinstance(ws.cell(r, c).value, (int, float)) else 0.0
        rows.append(dict(pac=str(pac).strip(), srv=str(ws.cell(r, 4).value or '-').strip(), lote=str(ws.cell(r, 5).value or '').strip(),
                         i=ini.date(), f=fim.date(), pb=num(9), pp=num(10), pr=num(11), pl=str(ws.cell(r, 13).value or '').strip(),
                         pav=ws.cell(r, 14).value, sit=ws.cell(r, 15).value or ''))
    ordem = {}
    arvore = {}
    for x in rows:   # PL → pacote → lote → tarefas, na ordem do cronograma
        pl = arvore.setdefault(x['pl'], {})
        pk = pl.setdefault(x['pac'], {})
        lt = pk.setdefault(x['lote'], {'resumo': None, 'tarefas': []})
        if x['srv'] == '-' and lt['resumo'] is None: lt['resumo'] = x
        else: lt['tarefas'].append(x)
        ordem.setdefault(x['pl'], len(ordem))

    def folha(x, nome):
        return dict(n=nome, i=x['i'].isoformat(), f=x['f'].isoformat(), pp=round(x['pp'], 1), pr=round(x['pr'], 1),
                    pb=round(x['pb'], 1), s=x['sit'])

    def junta(nome, filhos, extra=None):
        i = min(c['i'] for c in filhos); f = max(c['f'] for c in filhos)
        dur = lambda c: max(1, (dt.date.fromisoformat(c['f']) - dt.date.fromisoformat(c['i'])).days)
        sw = sum(dur(c) for c in filhos)
        no = dict(n=nome, i=i, f=f, pp=round(sum(c['pp'] * dur(c) for c in filhos) / sw, 1),
                  pr=round(sum(c['pr'] * dur(c) for c in filhos) / sw, 1), pb=round(sum(c['pb'] * dur(c) for c in filhos) / sw, 1), c=filhos)
        if extra: no.update(extra)
        return no

    raiz = []
    for pl, pacs in sorted(arvore.items(), key=lambda kv: min(x['i'] for p in kv[1].values() for l in p.values()
                                                               for x in ([l['resumo']] if l['resumo'] else []) + l['tarefas'])):
        nos_pac = []
        for pac, lotes in pacs.items():
            nos_lote = []
            for lote, d in lotes.items():
                tarefas = [folha(t, t['srv']) for t in d['tarefas']]
                if d['resumo']:
                    no = folha(d['resumo'], lote or pac)
                    if tarefas: no['c'] = tarefas
                elif tarefas:
                    no = junta(lote or pac, tarefas)
                else: continue
                nos_lote.append(no)
            nos_lote.sort(key=lambda c: c['i'])
            nos_pac.append(nos_lote[0] if len(nos_lote) == 1 and not nos_lote[0].get('c') and nos_lote[0]['n'] in ('GERAL', pac)
                           else junta(pac, nos_lote))
            if nos_pac[-1]['n'] != pac: nos_pac[-1] = dict(nos_pac[-1], n=pac)
        nos_pac.sort(key=lambda c: c['i'])
        raiz.append(junta(nome_pl.get(pl, pl or 'SEM PL'), nos_pac, {'cod': pl}))
    corte = vp['P5'].value
    dados = dict(ref=(ref.date().isoformat() if isinstance(ref, dt.datetime) else None), corte=corte.strftime('%Y-%m'),
                 ret=f"{wb['RETRATO 2027']['D5'].value}-12-31", ini=min(x['i'] for x in rows).isoformat(),
                 fim=max(x['f'] for x in rows).isoformat(), n=len(rows), arv=raiz)
    json.dump(dados, open(out, 'w'), ensure_ascii=False, separators=(',', ':'))
    print(len(rows), 'atividades ·', len(raiz), 'PLs ·', sum(len(p['c']) for p in raiz), 'pacotes · referência', dados['ref'], '·', dados['ini'], '→', dados['fim'])


if __name__ == '__main__':
    main()
