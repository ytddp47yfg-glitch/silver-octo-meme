# Gerador da página "Fluxo de caixa das obras"

A página (`site/index.html`) mostra uma tela inicial com uma lista de obras. Cada obra tem os
painéis DASHBOARD e RETRATO e todas as abas da planilha. Os dados de cada obra ficam num
arquivo separado em `site/obras/<id>.json`, que só é carregado quando a obra é escolhida.

## Incluir ou atualizar uma obra

1. Parta de `fluxo_caixa_botanico_modelo.xlsx` (mesma estrutura de abas) e preencha as abas de
   entrada da nova obra: VALOR POR PL, "INCC-FGV", DADOS BRUTOS, CURVA PREVISION e
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
