---
name: padrao-obras
description: Padrão único das obras da EPO na ferramenta de fluxo de caixa (Botânico, Ed. Jardim, Torre Catharina e qualquer obra nova). Use sempre que for atualizar, regerar, incluir ou comparar obras/cenários, mexer em painéis, planilhas, PDFs ou na identidade visual, ou quando pedirem que "todas as obras fiquem iguais". Define o que toda obra tem, como regerar com um comando só e como conferir.
---

# Padrão das obras (todas iguais)

Toda obra passa pela **mesma sequência** e sai com o **mesmo conjunto** de planilhas, painéis e PDFs. O que muda de uma
obra para outra são só os dados (pasta do cenário) e os insumos opcionais declarados no registro.

## Onde fica cada coisa

| O quê | Onde |
|---|---|
| Registro de obras e cenários | `ferramenta/obras.json` |
| Dados de cada cenário | `ferramenta/obras/<obra>/<pasta>/` → `data.json`, `data2.json` (DADOS BRUTOS, CURVA PREVISION, CRONO…), `cfg.json`, `ret_lens.json` [, `aux_estrutura.json`] |
| Insumos opcionais da obra | `ferramenta/obras/<obra>/extras/` (ex.: `imp.json`, `fachada_tasks.json`) e `gerador/cortes/<obra>.json` |
| Gerador da planilha | `ferramenta/build_modelo.py` · auditoria `ferramenta/audit.py` · `ferramenta/rebuild.sh` (um cenário) |
| Comando único | `python3 ferramenta/regerar.py [obra …] [--sem-site]` |
| Saídas intermediárias (não versionar) | `ferramenta/saida/` |
| Site, painéis, PDFs | `gerador/` (`viewer.py`, `make_site.py`, `multi_tpl.html`, `pdfs.js`) → `site/` |
| Planilhas entregues | raiz do repo (`fluxo_caixa_<obra>_modelo[_CENÁRIO].xlsx`) e `site/xlsx/<id>.xlsx` + `.b64.txt` |

Requisitos: Python 3 com `openpyxl`, LibreOffice (recálculo via `/mnt/skills/public/xlsx/scripts/recalc.py`), Node com
Playwright (PDFs). Se o Python do sistema não tiver openpyxl, exporte `PYTHONPATH` para uma pasta com ele.

## O que TODA obra tem

Por cenário (Padrão = curvas digitadas, `CENARIO=PADRÃO`; Prevision e variações = curvas do Prevision):

1. **Planilha** com as abas de sempre + **CURVA FÍSICA**, **GANTT** e **ESQUEMÁTICO**, recalculada, sem erros, auditoria
   **19 OK de 19**.
2. **Painéis no site**: DASHBOARD, RETRATO 2027, CURVA FÍSICA, GANTT, ESQUEMÁTICO.
3. **Mapas** (MAPA IMPERMEABILIZAÇÃO, MAPA FACHADA) quando a obra tiver os insumos (planta/elevações e cronogramas
   específicos; hoje só o Ed. Jardim). Para ligar numa obra nova, siga a skill `mapas-obra`.
4. **PDF por painel** (botão ⬇ PDF baixa o do painel aberto): A4 paisagem; ESQUEMÁTICO e mapas em **A3** paisagem.
5. **⬇ Excel** com a planilha do cenário.

Se uma obra não tiver algo da lista, isso é uma diferença a resolver (falta de insumo → diga ao usuário qual arquivo
precisa), não uma variação aceitável.

## Regerar (sempre por aqui)

```bash
python3 ferramenta/regerar.py ed-jardim          # uma obra
python3 ferramenta/regerar.py                    # todas
```

O script, para cada obra: 1ª passada de todos os cenários → Gantt (`gerador/gantt.py`, cronograma do cenário
`gantt_de`) e curva física (`gerador/fisico.py`, todos os cenários, base `fis_base`) → 2ª passada com as abas extras →
recalibra os cenários com `fixar_2027` → JSON do esquemático e dos mapas → `viewer.py` → copia as planilhas. Depois:
`make_site.py` → `pdfs.js` → `make_site.py` → `conferir_publicacao.py preparar`. Aborta se algum cenário não der 19 OK
ou tiver erro de fórmula.

Em seguida **publique com a rotina do `CLAUDE.md`** (lotes do `site/.pendente.json` → reler → `conferir`) e faça commit
de `ferramenta/` (dados e cfg alterados), das planilhas, de `site/` e de `site/.publicado.json`.

## Valores fixos de 2027

Cenários com `"fixar_2027"` no registro (Ed. Jardim Padrão = 56.009.788,87; Prevision ajustado = 53.460.558,96) mantêm o
2027 da revisão anterior: a diferença vai para 2028 proporcionalmente (`cfg.desloc`). O `regerar.py` recalibra o
`cfg.desloc.fator` sozinho (secante, tolerância de meio centavo) a cada mudança de dados. Não mexa no fator à mão.

## Conferência obrigatória a cada build

- 2027 de **cada** cenário antes × depois (tabela na resposta), total projetado e realizado até o corte.
- Mudança só estética **não pode** alterar número nenhum.
- Auditoria 19 OK em todos os cenários (o script já cobra).
- PDFs: mapas com 2 folhas A3; esquemático com uma folha de cortes + uma por quadro.

## Identidade visual EPO (igual em todas as obras)

- Logo: `gerador/marca/epo_logo.svg` (vetor; o PNG antigo era esticado). Caixas do logo na proporção 1,1 : 1.
- Verdes: escuro `#0b6938`, sálvia `#a1c1aa` (só no quadradinho do logo), grafite `#2d2a26`.
- Linha sob os títulos: **só verde-escuro**, sem trecho claro no início.
- Impressão: fundo **branco** (nada de cinza), sem a borda externa da página; quadros lado a lado com a **mesma altura**.
- A3: cabeçalho com nome da obra + logo em toda folha; rodapé com "Versão de dd/mm/aaaa".
- Elevações/plantas de projeto: traço em cinza claro, cores dos períodos por cima (multiply).

## Incluir uma obra ou um cenário

1. Crie `ferramenta/obras/<obra>/<pasta>/` com `data.json`, `data2.json`, `cfg.json` (copie a estrutura de uma obra
   existente; `fluxo_caixa_MODELO_EM_BRANCO.xlsx` explica as abas de entrada) e `ret_lens.json`.
2. Acrescente a obra/cenário em `ferramenta/obras.json` (`id`, `cenario`, `descricao`, `ordem`, `pasta`,
   `planilha_raiz`, `gantt_de` [, `CENARIO`, `fixar_2027`]) e, na obra, `fis_base`, `esquematico`, `extras`.
3. `python3 ferramenta/regerar.py <obra>` e publique.

Atualizações recorrentes: incorrido do ERP → skill `incorrido-erp`; esquemático → `esquematico-avanco`;
mapas → `mapas-obra`.
Cronograma de atividades do Prevision → `gerador/prevision/crono.py`; CFF → `gerador/prevision/cff.py` (ver `entrega/PROCEDIMENTO.md`, seção 6.2).
