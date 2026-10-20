---
name: mapas-obra
description: Painéis MAPA IMPERMEABILIZAÇÃO (planta pintada por período) e MAPA FACHADA (elevações do projeto pintadas por face e etapa, com planta baixa das faces), impressos em 2 folhas A3. Use quando pedirem mapa, planta ou elevação colorida por data/período, para atualizar os mapas do Ed. Jardim ou montá-los em outra obra.
---

# Mapas por período (impermeabilização e fachada)

Os dois painéis usam o mesmo layout (`renderMapa` em `gerador/multi_tpl.html`): desenho colorido + chave dos períodos;
folha 2 com a legenda de datas e a linha do tempo semanal. PDF em **2 folhas A3** (planta/elevação + chave | datas +
linha do tempo), fundo branco, quadros com a mesma altura. O `ferramenta/regerar.py` liga os painéis quando a obra
declara os insumos em `ferramenta/obras.json` → `"extras"`.

## MAPA IMPERMEABILIZAÇÃO

Insumos: cronograma de impermeabilização da obra (xlsx com sistemas × ambientes × semanas) e a planta do pavimento
(imagem). Hoje: Ed. Jardim, 3º pavimento (`gerador/imper/ed-jardim_planta_3pav*.{png,jpg}`).

1. Recolorir a planta pelas hachuras: `node gerador/imper_planta.js planta.png planta_cor.jpg` (classes de cor da planta → período; ajuste as
   regras de cor/área ao desenho da obra).
2. `python3 gerador/impermeabilizacao.py <crono.xlsx> <planta_cor.jpg> <saida.json>`: períodos (PERIODOS no topo do
   script: sistema, linhas do cronograma, início replanejado), etapas regularização → manta → proteção mecânica por
   ambiente (estimativas `REG_SEM`/`PROT_SEM` quando o cronograma não traz).
3. Guarde o JSON em `ferramenta/obras/<obra>/extras/imp.json` e declare `"impermeabilizacao"` em `extras`.

Regras já combinadas (Ed. Jardim): sistema A começa no 1º dia útil de mar/27; o período vai da regularização à proteção
mecânica (não só a manta).

## MAPA FACHADA

Insumos: lote FACHADA do CRONOGRAMA ATIVIDADES do cenário (faces × etapa F1 4º–13º / F2 14º–28º, cremalheira,
limpeza), as **elevações do projeto** e a **planta baixa** das faces. Ed. Jardim: elevações das imagens do Estudo de
fachada R12 (`gerador/imper/ed-jardim_elevacoes_{A_B,C_D}.png`), contorno da planta de distribuição do mesmo estudo.

`python3 gerador/fachada_mapa.py <data2.json do cenário> <fachada_tasks.json> <saida.json>` (o `regerar.py` roda por
cenário, com o `data2.json` de cada um). No script, por obra:
- `FACES` (ordem e nomes), `P1`/`P2` (pavimentos das etapas);
- `ELEV`: por face, imagem, recorte x, torre x, nível do topo, do 4º e do térreo (px da imagem) — meça as linhas de
  nível do desenho; `CR_PX` = faixa da cremalheira na frontal (Ed. Jardim: torre do elevador de obra + 3 vãos de vidro);
- `CONT`: vértices do contorno da planta por face.

Padrão visual: uma região **por face e etapa** (sem dividir por balancim); traço do desenho clareado (`SUAVE`, 42% do
contraste); cores dos períodos com multiply (opacidade .72); rótulos com contorno escuro; planta baixa em desenho
separado (abaixo da elevação na tela, no topo da coluna da legenda no A3); faces D e B com texto girado.

## Levar para outra obra

Peça ao usuário: cronograma de impermeabilização + planta do pavimento; elevações das fachadas (imagem/PDF) + planta
com as faces. Sem esses arquivos o mapa não existe para a obra; diga isso em vez de desenhar algo esquemático.
Depois: gere os JSON, declare em `extras`, `python3 ferramenta/regerar.py <obra>`, publique com a rotina do `CLAUDE.md`.
