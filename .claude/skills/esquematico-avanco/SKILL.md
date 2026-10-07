---
name: esquematico-avanco
description: Cria a aba ESQUEMÁTICO na planilha da ferramenta de fluxo de caixa: corte ilustrativo do prédio, pavimento × PL, colorido pela situação (concluído / em andamento / a executar), com a situação atual do cronograma e a previsão para uma data. Use quando pedirem esquemático, corte, "andamento ilustrativo das PLs", visão por pavimento ou algo parecido com o 3D colorido dos relatórios mensais de obra.
---

# Esquemático de avanço por pavimento (aba ESQUEMÁTICO)

Gera, dentro da planilha da ferramenta, um **corte esquemático** do prédio: uma linha por pavimento (do barrilete ao
térreo), uma coluna por PL, cada célula pintada pela situação da PL naquele pavimento. É a versão em planilha do
3D colorido dos relatórios mensais (verde = executado, cinza = a executar), mas calculada a partir do cronograma.

São dois cortes lado a lado:
1. **Situação em <data de referência do cronograma>**: % realizado do Prevision (coluna K da aba CRONOGRAMA ATIVIDADES).
2. **Previsão para <data>**: data digitável (célula C6, padrão = data do RETRATO). O saldo de cada tarefa segue linear
   da data de referência (ou do início planejado, se for depois) até o término planejado.

Tudo é fórmula: colar um export novo do Prevision na aba CRONOGRAMA ATIVIDADES e recalcular atualiza o esquemático.

## Pré-requisitos

- A aba **CRONOGRAMA ATIVIDADES** da ferramenta preenchida com o export de atividades do Prevision
  (colunas B:K = ID, pacote, serviço, lote, início, término, data de referência, base, previsto, realizado).
- O **de-para pacote → PL** (colunas R:S da mesma aba) sem pendências (`O2 = 0`). A coluna M traz a PL de cada atividade.
- O **nº do pavimento** sai do nome do lote (coluna N): "13º PAV. TIPO" → 13, "TÉRREO …" → 0, "1º SUBSOLO" → −1.
  Lotes sem número (GERAL, FACHADA, ELEVADORES) ficam fora do corte.

## Como gerar

No build da ferramenta (gera a aba junto com as outras):

```bash
ESQ=1 FIS_JSON=… GANTT_JSON=… bash rebuild.sh <pasta de dados> <saída>
```

`build_modelo.py` chama `gerador/esquematico.py` → `esquematico(wb, obra)` depois das abas CURVA FÍSICA e GANTT.
Avulso, numa planilha já gerada: `python3 gerador/esquematico.py ferramenta.xlsx saida.xlsx`. Esse modo relê e regrava a
pasta com openpyxl e perde detalhes de gráficos, então prefira o build. Depois de gerar:

1. Recalcule (LibreOffice: `recalc.py`) e reaplique os cantos arredondados (`cantos.py`).
2. Rode a auditoria. A aba ESQUEMÁTICO já está na lista de exclusões de `audit.py`; tem que dar 19 OK.
3. Confira visualmente: converta só a aba para PDF/PNG (copie a planilha com `data_only=True`, apague as outras abas,
   `soffice --convert-to pdf`) e verifique:
   - a estrutura sobe de baixo para cima;
   - o último pavimento concretado bate com o cronograma;
   - a faixa de infraestrutura e a de fachada fazem sentido.

## O que a aba mostra

| Elemento | Regra |
|---|---|
| Célula pavimento × PL | média das tarefas da PL no pavimento, ponderada pela duração. A linha-resumo do pacote/lote (serviço "-") só entra quando o lote não tem tarefas |
| Colunas (PLs) | PLs com atividades em pelo menos 3 pavimentos (`MIN_PAV`), na ordem do código |
| Faixa INFRAESTRUTURA (abaixo do térreo) | PLs cujo nome tem CONTENÇÃO, TERRAPLENAGEM ou FUNDAÇÃO, sobre todas as atividades |
| Faixa vertical FACHADA | PLs com FACHADA no nome (os lotes de fachada não têm pavimento) |
| Coluna PAVIMENTO | todas as PLs do pavimento, com barra de dados |
| Linha "% da PL nos pavimentos" | a PL em todos os pavimentos numerados |
| Cores | concluído ≥ 99,95% (verde 1F7A52) · 50–99% (dourado C0A062) · 1–49% (dourado claro E8D9B0) · 0% (cinza D9D4C7) · sem atividade (creme F4F1EA) |

As colunas auxiliares (peso e % na data do 2º corte, uma linha por atividade) ficam ocultas a partir da coluna BH.

## Ajustes comuns

- **Outra obra**: nada a configurar se os lotes do Prevision começarem pelo número do pavimento ("5º PAV …").
  Se a obra usar outro padrão (ex.: "PAV 05", "TORRE A – 5º"), ajuste `_pav()` em `gerador/esquematico.py` e a fórmula da
  coluna N da aba CRONOGRAMA ATIVIDADES (`build_modelo.py`), sempre os dois juntos.
- **Mais ou menos PLs no corte**: `MIN_PAV`. Para fixar uma lista, filtre `pls` em `_ler()`.
- **Rótulo dos pavimentos**: é o lote mais frequente do pavimento sem o prefixo "Nº PAV." (`_rotulo()`).
- **Datas**: C5 = maior data de referência das atividades. C6 = RETRATO 2027!K5, editável (célula amarela).

## Limites (diga isso ao usuário)

- É **ilustrativo, sem escala**: mostra pavimento × PL, não a geometria. Não substitui o levantamento em campo nem o 3D.
- **Planta** (setores / trechos num pavimento) precisa do desenho da obra: peça a planta (imagem ou PDF) com os
  setores marcados. Os pacotes já trazem SETOR A/B/C, TRECHO A/B/C e QUADRANTES; dá para montar blocos por setor
  com a mesma lógica (filtro pelo nome do pacote/lote), posicionados conforme a planta.
- A previsão é linear entre as datas do cronograma; não considera a curva de alocação de cada tarefa.
