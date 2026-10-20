---
name: esquematico-avanco
description: Cria a aba ESQUEMÁTICO na planilha da ferramenta de fluxo de caixa: corte ilustrativo do prédio, pavimento × PL, colorido pela situação (concluído / em andamento / a executar), com a situação atual do cronograma e a previsão para uma data. Use quando pedirem esquemático, corte, "andamento ilustrativo das PLs", visão por pavimento ou algo parecido com o 3D colorido dos relatórios mensais de obra.
---

# Esquemático de avanço por pavimento (aba ESQUEMÁTICO)

Gera, dentro da planilha da ferramenta, um **corte esquemático** do prédio: uma linha por pavimento (do barrilete ao
térreo), uma coluna por PL, cada célula pintada pela situação da PL naquele pavimento. É a versão em planilha do
3D colorido dos relatórios mensais (verde = executado, cinza = a executar), mas calculada a partir do cronograma.

Os cortes ficam lado a lado (padrão: três quadros):
1. **Situação em <data de referência do cronograma>**: % realizado do Prevision (coluna K da aba CRONOGRAMA ATIVIDADES).
2. **Marco(s)**: previsão em datas definidas na configuração da obra (`cfg.json` → `"esq_marcos": [["Marco", "2026-12-31"]]`;
   sem isso, 31/12 do ano da situação). Uma linha por marco a partir da célula C6, editável.
3. **Retrato**: previsão para a data do RETRATO (RETRATO 2027!K5), sempre o último quadro.

Nas previsões, o saldo de cada tarefa segue linear da data de referência (ou do início planejado, se for depois)
até o término planejado.

Tudo é fórmula: colar um export novo do Prevision na aba CRONOGRAMA ATIVIDADES e recalcular atualiza o esquemático.

## Pré-requisitos

- A aba **CRONOGRAMA ATIVIDADES** da ferramenta preenchida com o export de atividades do Prevision
  (colunas B:K = ID, pacote, serviço, lote, início, término, data de referência, base, previsto, realizado).
- O **de-para pacote → PL** (colunas R:S da mesma aba) sem pendências (`O2 = 0`). A coluna M traz a PL de cada atividade.
- O **nº do pavimento** sai do nome do lote (coluna N): "13º PAV. TIPO" → 13, "TÉRREO …" → 0, "1º SUBSOLO" → −1.
  Lotes sem número (GERAL, FACHADA, ELEVADORES) ficam fora do corte.

## Como gerar (todas as obras, mesmo caminho)

O esquemático faz parte do padrão de **todas** as obras e de **todos** os cenários (skill `padrao-obras`). Não gere à mão:

```bash
python3 ferramenta/regerar.py <obra>        # botanico | ed-jardim | torre-catharina (sem argumento = todas)
```

O `regerar.py` faz a 2ª passada de cada cenário com `ESQ=1` (aba ESQUEMÁTICO na planilha, via `gerador/esquematico.py`)
e grava o JSON do painel com `esquematico.py --json`, usando o que a obra tiver em `ferramenta/obras.json` → `"esquematico"`:
`acab` (cortes "acabamento finalizado até…") e `cortes` (arquivo em `gerador/cortes/<obra>.json`). Obra sem essas chaves
mostra só os quadros de PLs.

Configuração dos quadros: `cfg.json` do cenário em `ferramenta/obras/<obra>/<pasta>/` (`esq_marcos`, `esq_subdiv`,
`esq_extra`). Para uniformizar, use as mesmas chaves em todos os cenários da obra.

Conferência depois de gerar (o `regerar.py` já aborta se a auditoria não der 19 OK):
- a estrutura sobe de baixo para cima;
- o último pavimento concretado bate com o cronograma;
- a faixa de infraestrutura e a de fachada fazem sentido;
- o último quadro termina na data final do Prevision (Ed. Jardim: 30/04/2029).

Avulso, numa planilha já gerada: `python3 gerador/esquematico.py planilha.xlsx saida.xlsx` (relê e regrava com openpyxl e
perde detalhes de gráficos; prefira o `regerar.py`).

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

As colunas auxiliares (peso e % em cada marco, uma linha por atividade) ficam ocultas depois do último corte.

## Ajustes comuns

- **Outra obra**: nada a configurar se os lotes do Prevision começarem pelo número do pavimento ("5º PAV …").
  Se a obra usar outro padrão (ex.: "PAV 05", "TORRE A – 5º"), ajuste `_pav()` em `gerador/esquematico.py` e a fórmula da
  coluna N da aba CRONOGRAMA ATIVIDADES (`build_modelo.py`), sempre os dois juntos.
- **Mais ou menos PLs no corte**: `MIN_PAV`. Para fixar uma lista, filtre `pls` em `_ler()`.
- **Rótulo dos pavimentos**: é o lote mais frequente do pavimento sem o prefixo "Nº PAV." (`_rotulo()`).
- **Datas**: C5 = maior data de referência das atividades. C6… = marcos e, por último, RETRATO 2027!K5, todas editáveis (células amarelas).
- **Último quadro**: a data final do Prevision da obra (Ed. Jardim: `["Marco", "2029-04-30"]`), não 31/12.
- **Mais marcos**: acrescente pares em `esq_marcos` (ex.: `[["Marco", "2026-12-31"], ["Marco", "2027-06-30"]]`). Cada marco é mais um quadro. Para pôr o retrato no meio (ordem cronológica), inclua `["Retrato", "RETRATO"]` na posição desejada (ex.: `[["Marco","2026-12-31"],["Retrato","RETRATO"],["Marco","2028-12-31"],["Marco","2029-12-31"]]`).

## Limites (diga isso ao usuário)

- É **ilustrativo, sem escala**: mostra pavimento × PL, não a geometria. Não substitui o levantamento em campo nem o 3D.
- **Planta** (setores / trechos num pavimento) precisa do desenho da obra: peça a planta (imagem ou PDF) com os
  setores marcados. Os pacotes já trazem SETOR A/B/C, TRECHO A/B/C e QUADRANTES; dá para montar blocos por setor
  com a mesma lógica (filtro pelo nome do pacote/lote), posicionados conforme a planta.
- A previsão é linear entre as datas do cronograma; não considera a curva de alocação de cada tarefa.

## Quebrar um PL em colunas (cfg `esq_subdiv`)

`{"<código do PL>": [["RÓTULO", ["PACOTE 1", "PACOTE 2"]], ...]}` troca a coluna do PL por uma coluna por grupo de
pacotes do Prevision (só no esquemático; curvas e valores não mudam). Confira que todo pacote do PL na aba
CRONOGRAMA ATIVIDADES (de-para R:S) caiu em algum grupo. Ed Jardim: PL 30 (revestimentos de acabamento) em
REV. PEDRAS/BANCADAS (rev. parede, piso, porcelanatos, bancadas, piso elevado/externo), PISO MAD./RODAPÉ
(piso madeira/laminado/vinílico, rodapé, deck) e TRAT. PEDRAS (tratamento de granito/mármore, polimento).

## Coluna por pacote só em alguns cortes (cfg `esq_extra`)

`[["CHECK LIST", ["CHECKLIST"], "2028-01-01"]]` acrescenta a coluna (pacotes sem PL) só nos cortes de previsão com data
digitada a partir de 01/01/2028. O JSON do painel traz as colunas de cada quadro em `q.pls`.

## Corte "acabamento finalizado até a data" (painel)

`python3 gerador/esquematico.py --json calc.xlsx esq.json --acab "REVESTIMENTO PAREDE|REVESTIMENTOS PISO|REVESTIMENTO PORCELANATO PAREDE|REVESTIMENTOS PORCELANATO PISO" 2027-12-31 "do revestimento de parede e de piso (cerâmica/porcelanato)"`
grava `acab` no JSON e o painel ESQUEMÁTICO mostra, antes dos quadros, o corte por apartamento (pavimentos TIPO, GARDEN
e DUPLEX; o duplex 26º/27º é um apartamento; nº do apto = pavimento × 100): verde = término dos pacotes até a data,
amarelo = em execução na data, cinza = depois (com o mês). Ed Jardim, Prevision Atual: 700, 900, 1000, 1100, 1200,
1400, 1600 e 1800 (confere com as anotações do usuário de 19/10/2026).

Outros cortes no mesmo layout: `--cortes gerador/cortes/<obra>.json` (lista de {pacotes, ate, desde, rot, tit, res,
feito, leg}). Ed Jardim: check list em 2028 e em 2029 até 30/04/2029, fim do Prevision (`desde` = 1º/jan do ano; o que
concluiu antes aparece em verde-claro). Para outra obra, crie `gerador/cortes/<obra>.json` e aponte em `ferramenta/obras.json`.

## PDF em A3

O PDF do painel ESQUEMÁTICO (gerador/pdfs.js) sai em A3 paisagem: folha 1 = cabeçalho + cortes por apartamento lado a
lado; depois um quadro de PLs por folha, ampliado (zoom calculado para caber em 1480 × 1000 px). CSS em `html.a3`.
