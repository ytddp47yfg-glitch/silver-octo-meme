# Pacote do fluxo de caixa EPO (Botânico, Ed Jardim, Torre Catharina): 08/10/2026

Tudo o que é preciso para regerar as planilhas e o painel. Comece pelo `RESUMO.md`, que traz o estado atual e as regras.
Em `PROJETO_INSTRUCOES.md` está o texto para as instruções de um Projeto no claude.ai.

## Estrutura

```
repo/                      cópia do repositório GitHub ytddp47yfg-glitch/silver-octo-meme
                           (branch claude/planilha-fluxo-financeiro-d5b771)
  CLAUDE.md                regras do projeto (rotina obrigatória de publicação conferida)
  gerador/                 viewer.py, make_site.py, multi_tpl.html, pdfs.js, gantt.py, fisico.py, esquematico.py,
                           aux_estrutura.py, conferir_publicacao.py, paineis_extra.py, LEIA-ME.md
  .claude/                 skill esquematico-avanco + gancho Stop (bloqueia o fim do turno se o painel não estiver conferido)
  site/                    painel (index_artifact.html, obras/*.json, pdf/), publicado em
                           https://claude.ai/artifact/9TpDGLVKmKfKstM6MiEAvy
  fluxo_caixa_*.xlsx       as 8 planilhas atuais (iguais às de download do painel)
ferramenta/                gerador das planilhas (não estava no repositório)
  build_modelo.py          gera a planilha de um cenário a partir de uma pasta de dados
  audit.py                 auditoria (19 verificações), cantos.py (cantos arredondados dos gráficos após recalcular)
  paineis_extra.py         abas CURVA FÍSICA e GANTT da planilha
  incc_hist.json           série do INCC-DI (até set/2026)
  rebuild.sh               gera + recalcula + audita um cenário
  regerar_tudo.sh          regera as 8 planilhas e o site e prepara a publicação
  tools/                   scripts de apoio (dados brutos do ERP, comparação do Prevision etc.)
  obras/
    botanico/              data.json, data2.json, cfg.json, ret_lens.json, aux_estrutura.json (Padrão e Prevision)
    ed-jardim/data         Padrão e Prevision;  ed-jardim/sim4 = Prevision ajustado;  fis.json, *_gantt.json, *_esq.json
    torre-catharina/data   Visão Set/26 e Prevision;  torre-catharina/sim = Simulação (versão 1)
origens/                   planilhas originais (MODELO 05), medições financeiras (IEC), DADOS do ERP (out/26),
                           cronogramas do Ed Jardim (07/10/26), relatório de exemplo (pptx do esquemático)
```

## Pasta de dados de um cenário

| Arquivo | Conteúdo |
|---|---|
| `data.json` | Orçamento, projeção e curvas digitadas por PL, INCC. |
| `data2.json` | DADOS BRUTOS do ERP, CURVA PREVISION, CRONOGRAMA ATIVIDADES (Prevision), VALOR POR PL, MODELO PARA BI etc. |
| `cfg.json` | Premissas da obra (ver `RESUMO.md`). |

Campos do `cfg.json`:
- **Datas:** `corte`, `inicio`.
- **Ligação com o ERP:** `erp`/`erp2`/`ignore` (de-para das abas para os grupos do ERP).
- **Projeção:** `fator_proj` (fator IEC por PL), `meta_min_incorrido`.
- **Ajustes:** `contingencia`, `conciliacao` {rs, incc, fonte, mes}, `force` (fonte de curva forçada), `desloc` (fator K9 do 2027 no Ed Jardim).
- **Estrutura:** `aux_estrutura` e `aux_estrutura_cfg.simulacao` (datas simuladas).
- **Esquemático:** `esq_marcos` (marcos do esquemático).
- **Nomes de exibição:** `cenario_nome` / `cenario_nome_padrao`.

## Como regerar

Requisitos:
- Python 3 com openpyxl.
- LibreOffice (`soffice`), para recalcular as fórmulas.
- Node 20+ com Playwright (Chromium), só para os PDFs do painel.

Um cenário (de dentro de `ferramenta/`):

```bash
bash rebuild.sh obras/torre-catharina/data saida/cat_tool                 # Prevision
CENARIO=PADRÃO bash rebuild.sh obras/torre-catharina/data saida/cat_padrao  # Visão Set/26
```

Tudo (8 planilhas + painel): `bash regerar_tudo.sh`. Depois publique com a rotina do `repo/CLAUDE.md`: preparar → publicar → reler → conferir.

Sobre o recálculo:
- O `rebuild.sh` recalcula com o script `recalc.py` da skill xlsx do Claude (variável `RECALC`).
- Fora do ambiente do Claude, aponte `RECALC` para um script equivalente. Outra opção é abrir o `_calc.xlsx` no Excel e salvar.

Conferências obrigatórias a cada build:
- auditoria 19 OK;
- 2027 de cada cenário (o do Ed Jardim Padrão e Prevision ajustado é fixado);
- projeção total.

Nunca hospede o painel com dados da empresa fora da página privada do claude.ai sem autorização.
