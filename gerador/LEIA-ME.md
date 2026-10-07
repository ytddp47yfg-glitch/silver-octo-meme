# Gerador da página "Fluxo de caixa das obras"

A página (`site/index.html`) mostra uma tela inicial com uma lista de obras. Cada obra tem os
painéis DASHBOARD e RETRATO e todas as abas da planilha. Os dados de cada obra ficam num
arquivo separado em `site/obras/<id>.json`, que só é carregado quando a obra é escolhida.

## Incluir ou atualizar uma obra

1. Parta de `fluxo_caixa_MODELO_EM_BRANCO.xlsx` (mesma estrutura, sem dados; a aba LEIA-ME explica cada
   entrada) e preencha as abas de entrada da nova obra: VALOR POR PL, "INCC-FGV", DADOS BRUTOS, CURVA PREVISION e
   CRONOGRAMA ATIVIDADES.
2. Recalcule a planilha no LibreOffice para gravar os valores das fórmulas (o Excel já grava ao salvar):
   `soffice --headless --convert-to xlsx --outdir /tmp/recalc OBRA.xlsx`
3. Gere o arquivo da obra (`VW_ID` = identificador curto, sem espaços nem acentos):

   ```
   VW_VALUES=/tmp/recalc/OBRA.xlsx VW_FORMULAS=OBRA.xlsx VW_NOME="Nome da Obra" VW_ID=nome-da-obra \
     python3 gerador/viewer.py --json site/obras/nome-da-obra.json
   ```
4. Monte a página inicial: `python3 gerador/make_site.py`
5. Publique `site/index_artifact.html` como Artifact, mandando cada `site/obras/<id>.json`
   como arquivo de apoio no caminho `obras/<id>.json`.

Dependências: Python 3 com `openpyxl` e LibreOffice (para o passo 2).

## Painel CURVA FÍSICA (Ed Jardim)

Curvas de avanço físico por cenário, calculadas a partir das PLs (itens diretos, ponderados pela projeção):

    python3 gerador/fisico.py fis.json "Padrão=ej_padrao_calc.xlsx" "Prevision=ej_tool_calc.xlsx" "Prevision ajustado=ej_sim4_calc.xlsx"

Depois gere o JSON de cada cenário com `VW_FIS=fis.json` no ambiente do `viewer.py`; o painel aparece só nas obras que têm esses dados.

## Painel GANTT (todas as obras)

Gantt do cronograma de atividades do Prevision (aba CRONOGRAMA ATIVIDADES), em árvore PL → pacote → pavimento/tarefa:

    python3 gerador/gantt.py obra_calc.xlsx gantt.json

Depois gere o JSON de cada cenário com `VW_GANTT=gantt.json` no ambiente do `viewer.py`.

## PDF dos painéis

Com o site servido localmente (`python3 -m http.server 8766 -d site`), `node gerador/pdfs.js 8766` gera `site/pdf/<cenário>.pdf`
com todos os painéis do cenário (A4 paisagem; o Gantt repete a régua de meses em cada página e traz no final as notas dos
pacotes inferidos). Rode `python3 gerador/make_site.py` depois, para o botão "⬇ PDF" aparecer.

## Excel para download

Copie a planilha recalculada de cada cenário para `site/xlsx/<cenário>.xlsx` (mesmo nome do JSON em `site/obras/`) e rode
`python3 gerador/make_site.py`: o botão "⬇ Excel" aparece no cabeçalho do cenário. Para publicar no claude.ai (que não serve .xlsx),
gere também `site/xlsx/<cenário>.xlsx.b64.txt` (`base64 -w0 arquivo.xlsx > arquivo.xlsx.b64.txt`): a página monta o .xlsx a partir dele.
