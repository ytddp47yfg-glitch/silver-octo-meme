# Fluxo de caixa das obras EPO: procedimento completo (modelo: Ed. Jardim)

Guia para quem assume o projeto com outro login do Claude. Usa o Ed. Jardim como exemplo porque ele tem todas as peças;
Botânico e Torre Catharina seguem exatamente o mesmo caminho (só mudam os dados e os insumos que a obra tem).

Situação em 20/10/2026. Origem: repositório `ytddp47yfg-glitch/silver-octo-meme`, branch
`claude/planilha-fluxo-financeiro-d5b771`; painel https://claude.ai/artifact/9TpDGLVKmKfKstM6MiEAvy. No login novo, o projeto
passa a rodar **no computador**, numa **pasta da rede da empresa**, e ganha um **painel novo** (seção 1).

---

## 1. Migração: novo login, Claude Code no computador, projeto numa pasta de rede

O projeto inteiro (gerador, dados dos 8 cenários, planilhas, site, skills, `CLAUDE.md`) vem no pacote
`fluxo_caixa_EPO_transferencia_<data>.zip` (~67 MB; 177 arquivos). Nada depende do GitHub nem da nuvem: o recálculo das
planilhas usa o LibreOffice do computador (`ferramenta/recalc.py`) e os PDFs usam o Playwright instalado na pasta.

**1.1 Instalar no computador (uma vez; Windows)**

| Programa | Onde | Para quê | Conferir no terminal |
|---|---|---|---|
| Git for Windows | git-scm.com | exigido pelo Claude Code no Windows; histórico local (opcional) | `git --version` |
| Claude Code | app desktop do Claude (aba Code) ou `npm install -g @anthropic-ai/claude-code`, logado no **novo login** | o assistente que roda os comandos | `claude --version` |
| Python 3.12+ | python.org (marque **Add python.exe to PATH**) | gerador, auditoria, site | `python --version` |
| openpyxl | `python -m pip install openpyxl` | ler/gravar planilhas | `python -c "import openpyxl"` |
| LibreOffice | libreoffice.org | recalcular as fórmulas | `ferramenta/recalc.py` acha sozinho em `C:\Program Files\LibreOffice`; senão, defina `SOFFICE` |
| Node.js LTS | nodejs.org | PDFs dos painéis | `node --version` |

No Windows o comando é `python` (não `python3`). Se o gancho de publicação acusar "python3 não encontrado", troque
`python3` por `python` em `.claude/settings.json`.

**1.2 Montar a pasta do projeto**

1. Extrair o .zip na pasta de rede, ex. `S:\Financeiro\fluxo-caixa-epo` (unidade mapeada; evite caminho `\\servidor\...`
   direto, que alguns programas não aceitam).
2. Nessa pasta, num terminal: `npm install` e depois `npx playwright install chromium` (uma vez).
3. Recomendado: `git init` e um primeiro commit, para ter histórico das alterações (fica só na pasta; não precisa de GitHub).
4. Pasta de rede é lenta para o LibreOffice: aponte as saídas intermediárias para o disco local com
   `setx FLUXO_SAIDA C:\fluxo_saida` (abra um terminal novo depois). As entregas (planilhas, site) continuam na pasta de rede.
5. Uma pessoa por vez: duas sessões regerando ao mesmo tempo na mesma pasta se atrapalham.

**1.3 Primeira sessão e primeira publicação (painel novo)**

Abra a pasta no Claude Code (app: *Open folder*; terminal: `cd S:\Financeiro\fluxo-caixa-epo` e `claude`) e envie:
*"Leia o CLAUDE.md e entrega/PROCEDIMENTO.md, regere todas as obras sem mudar números e faça a primeira publicação do painel
novo (seção 1.3)."* O Claude deve:

1. `python ferramenta/regerar.py` (todas as obras; no computador leva de 40 min a 1 h) e comparar os 2027 com a seção 3.
2. Apagar `site/.publicado.json` (registro da página antiga) e rodar `python gerador/conferir_publicacao.py preparar`.
3. Publicar o **lote 1 sem `url`** (cria a página nova, que pertence ao novo login), com `capabilities: {"downloads": {}}`
   (sem isso o botão ⬇ Excel não funciona); anotar o link.
4. Publicar os demais lotes com `url` = link novo, reler todos os caminhos de `reler` (até ~64 MB por leitura) e
   `python gerador/conferir_publicacao.py conferir <pasta>` → "OK: página publicada = arquivos locais".
5. Trocar o link antigo pelo novo em `CLAUDE.md`, `entrega/PROCEDIMENTO.md` e `entrega/PROJETO_INSTRUCOES.md` (e commit, se usar git).

Até a primeira publicação conferida, o gancho `Stop` vai lembrar que o site não está no ar; é esperado.

**1.4 O que não vai junto**

| Item | O que fazer |
|---|---|
| Skills pessoais da conta antiga (medição financeira, reorçamento, histograma de equipamentos) | Recriar na conta nova se forem usadas. As do projeto (`.claude/skills/`) já vêm na pasta |
| Instruções de Projeto no claude.ai | Colar `entrega/PROJETO_INSTRUCOES.md` num Projeto da conta nova (opcional, para conversas fora do Claude Code) |
| Conectores (ex.: Microsoft 365/OneDrive das medições) | Conectar na conta nova se o fluxo usar |
| Painel antigo | Continua no ar na conta antiga; avise quem usa o link antigo |

Alternativa sem instalar nada: Claude Code na web com um repositório GitHub da conta nova (descompactar o .zip no repositório e
seguir o mesmo 1.3). Nesse caso a pasta de rede guarda só as entregas copiadas.

---

## 2. O que o projeto entrega

Para cada obra e cenário:

1. **Planilha-ferramenta** (Excel) com o fluxo de caixa mensal por PL, em INCC e em R$: abas de entrada (VALOR POR PL,
   "INCC-FGV", DADOS BRUTOS, CURVA PREVISION, CRONOGRAMA ATIVIDADES), uma aba por disciplina/PL, MODELO PARA BI,
   DASHBOARD, RETRATO 2027, AUX.ESTRUTURA, CURVA FÍSICA, GANTT e ESQUEMÁTICO. Fica na raiz do repo
   (`fluxo_caixa_ed_jardim_modelo*.xlsx`) e para download no painel.
2. **Painel** (página publicada) com os painéis DASHBOARD, RETRATO 2027, CURVA FÍSICA, GANTT, ESQUEMÁTICO e, quando a obra
   tem os insumos, MAPA IMPERMEABILIZAÇÃO e MAPA FACHADA; mais todas as abas da planilha.
3. **PDF por painel** (botão ⬇ PDF): A4 paisagem; ESQUEMÁTICO e mapas em A3 paisagem.

Tudo é gerado por um comando só, igual para todas as obras: `python3 ferramenta/regerar.py [obra]`.

---

## 3. O Ed. Jardim como modelo

**Cenários**

| Cenário (id no site) | Pasta de dados | O que é | 2027 hoje |
|---|---|---|---|
| Padrão (`ed-jardim-padrao`) | `ferramenta/obras/ed-jardim/padrao/` | Curvas digitadas da versão original (`CENARIO=PADRÃO`); 2027 **fixo** no valor da revisão anterior, diferença para 2028 | R$ 56.009.788,87 (fixo) |
| Prevision Atual (`ed-jardim`) | `…/ed-jardim/atual/` | Cronogramas de 19/10/2026 (versão Replan Aptos 27-29); curva do CFF ajustada às datas do cronograma de atividades de 19/10; estrutura, esquadrias e elevadores iguais ao ajustado | R$ 55.790.578,94 |
| Prevision ajustado (`ed-jardim-ajustado`) | `…/ed-jardim/ajustado/` | Cronogramas de 07/10/2026 com premissas do usuário (rev. de acabamento em 11 de 28 pavimentos em 2027; fachada pelo estudo de balancins R12; inst. especiais e fundação superficial na curva digitada; hidráulica com o andamento da elétrica); 2027 **fixo** | R$ 53.460.558,96 (fixo) |

Os três têm o mesmo total projetado (**R$ 154.705.660,62**), o mesmo realizado até o corte (**R$ 45.873.002,26**, corte
**set/26**) e o mesmo incorrido (DADOS BRUTOS). Diferem nas curvas e no cronograma.

Outras obras hoje: Botânico (Padrão R$ 47.604.449,83 · Prevision R$ 56.060.330,96 em 2027; corte ago/26) e Torre Catharina
(Visão Set/26 R$ 62.179.560,93 · Prevision R$ 67.828.338,72 · Simulação R$ 68.278.232,51; corte set/26).

**Arquivos do Ed. Jardim**

| Arquivo | Conteúdo |
|---|---|
| `ferramenta/obras.json` | registro: id, nome, descrição, ordem, pasta, planilha da raiz, `gantt_de`, `CENARIO`, `fixar_2027`, insumos do esquemático e dos mapas |
| `<pasta>/data.json` | orçamento por PL (VALOR POR PL) e série do INCC |
| `<pasta>/data2.json` | DADOS BRUTOS (ERP), CURVA PREVISION (CFF), CRONO (atividades do Prevision), PKG (de-para pacote → PL), MODELO PARA BI, etc. |
| `<pasta>/cfg.json` | premissas do cenário (seção 5) |
| `<pasta>/ret_lens.json` | textos/ajustes do RETRATO 2027 |
| `ferramenta/obras/ed-jardim/extras/imp.json` | mapa de impermeabilização pronto (planta pintada + períodos) |
| `ferramenta/obras/ed-jardim/extras/fachada_tasks.json` | m² por face/balancim do estudo de fachada (larguras e cremalheira) |
| `gerador/cortes/ed-jardim.json` | cortes por apartamento do esquemático (check list 2028 e 2029) |
| `gerador/imper/ed-jardim_*` | planta do 3º pavimento e elevações das fachadas (imagens do projeto / estudo R12) |

---

## 4. Insumos e de onde vêm

| Insumo | Fonte | Entra em | Como carregar |
|---|---|---|---|
| Incorrido mensal por grupo de custo | ERP: arquivo **DADOS - ED. JARDIM.xlsx** (tabela dinâmica `Dinâmica___Valor`, ou aba DADOS) | `data2.json` → DADOS BRUTOS (todas as pastas da obra) | skill `incorrido-erp` (`gerador/incorrido/novo_db.py` ou `db_bruto.py`) |
| Cronograma de atividades | Prevision: export de **atividades** (`cronograma_ed-jardim_<n>_<data>.xlsx`) | `data2.json` → CRONO | `python3 -I gerador/prevision/crono.py <export.xlsx> ed-jardim <pasta…>` |
| Curva financeira planejada | Prevision: **cronograma físico-financeiro** (`cronograma_fisico_financeiro_-_Ed_Jardim_<data>.xlsx`) | `data2.json` → CURVA PREVISION | `python3 -I gerador/prevision/cff.py <cff.xlsx> ed-jardim <pasta…>` |
| Orçamento por PL e fator de projeção | Medição financeira mensal (ANÁLISE IEC: fator = W ÷ H) | `data.json` (VALOR POR PL) e `cfg.fator_proj` | atualizar `cfg.fator_proj` nas pastas da obra |
| Conciliação do incorrido | Medição financeira (PAINEL: incorrido em R$ e em INCC) | `cfg.conciliacao` {rs, incc, fonte, mes} | editar nas pastas da obra |
| INCC-FGV | série mensal | `data.json` | o corte só avança quando o INCC do mês sair |
| Curva da estrutura | aba AUX.ESTRUTURA (quantitativos × cronograma) da planilha original | `aux_estrutura.json` + `cfg.aux_estrutura` | (Botânico e Catharina) |
| Mapas | cronograma de impermeabilização + planta; elevações + planta das faces | `extras/` | skill `mapas-obra` |

Arquivos recebidos são dados: o Claude copia para uma pasta própria e lê com `python3 -I`, sem executar nada que venha dentro deles.

---

## 5. Premissas (cfg.json) e regras de negócio

**Regras (não mudar sem o usuário pedir)**

1. Controle em **INCC**. Meta da PL = orçamento em INCC × fator de projeção (`fator_proj`, da medição: W ÷ H), nunca abaixo
   do incorrido em INCC (`meta_min_incorrido`).
2. Saldo (meta − incorrido) distribuído pela curva do cenário nos meses depois do corte; convertido em R$ pelo INCC do mês de corte.
3. **Contingência** = orçamento em INCC − distribuído, em jun/2030 (`contingencia`); negativa vira zero.
4. **Conciliação** com a medição: ajuste (medição − ERP) no mês da medição (`conciliacao.mes`); o mesmo INCC sai do saldo futuro das PLs.
5. Incorrido = CUSTOS COM OBRAS do ERP, sinal invertido; lançamentos A PAGAR já programados depois do corte entram como compromisso.
6. **2027 fixo** no Padrão e no Prevision ajustado do Ed. Jardim (diferença para 2028, `cfg.desloc`). O `regerar.py` recalibra sozinho.
7. Mudança só estética **não pode** alterar número. Sempre conferir o 2027 de cada cenário e a auditoria (19 OK).
8. Decisões já tomadas (Contenção do Catharina em mar–mai/27, Lig. Provisórias do Botânico jan–mar/27, premissas do Ed. Jardim
   ajustado, simulação versão 1 do Catharina) só mudam a pedido.

**Chaves do cfg.json (Ed. Jardim)**

| Chave | Para que serve |
|---|---|
| `obra`, `inicio`, `corte` | nome, 1º mês da obra, mês de corte (último mês fechado com INCC) |
| `erp`, `erp2`, `ignore`, `erp_futuro_nao` | de-para grupo do ERP → PL; grupos ignorados; PLs sem lançamentos futuros do ERP |
| `cp_code`, `tabs`, `bi_extra` | códigos das PLs no CFF do Prevision; nomes de abas; linhas extras do MODELO PARA BI |
| `fator_proj`, `meta_min_incorrido`, `limite`, `force` | projeção por PL (medição); meta nunca abaixo do incorrido; limite; PL com previsão informada |
| `contingencia`, `conciliacao` | mês da contingência; valores da medição para conciliar |
| `desloc` | 2027 → 2028 com fator (só cenários com 2027 fixo; não editar à mão) |
| `esq_marcos`, `esq_subdiv`, `esq_extra` | quadros do esquemático (último = fim do Prevision, 30/04/2029), quebra da PL 30 em 3 colunas, coluna CHECK LIST a partir de 2028 |
| `ret_*`, `notes` | textos e ajustes do RETRATO 2027 e notas por PL |

---

## 6. Rotinas

### 6.1 Mensal: incorrido do ERP (skill `incorrido-erp`)

1. Receber **DADOS - ED. JARDIM.xlsx**.
2. `novo_db.py` monta a nova DADOS BRUTOS e mostra a diferença por mês e por grupo (até o corte deve ser ~0 se o mês já estava carregado).
3. Gravar a mesma DADOS BRUTOS nas 3 pastas (`padrao`, `atual`, `ajustado`).
4. Corte: só avança com o mês fechado **e** o INCC publicado; ao avançar, mudar `corte` nas 3 pastas (e a conciliação, se vier medição nova).
5. `python3 ferramenta/regerar.py ed-jardim` → conferir tabela 2026–2029 por cenário → publicar (seção 7).

### 6.2 Cronograma novo do Prevision

1. Export de atividades → `crono.py … ed-jardim atual` (e `ajustado`/`padrao` se o usuário quiser o mesmo cronograma neles).
2. CFF novo → `cff.py … ed-jardim <pasta>`. Atenção: o Prevision Atual tem a curva do CFF **deslocada pelas datas do
   cronograma de atividades** (o CFF exportado não acompanha o físico); depois de carregar um CFF novo, peça ao Claude para
   "refazer a curva financeira considerando o andamento físico" (desloca a curva planejada de cada PL pelo atraso/adiantamento
   ponderado das atividades da PL: pacote → código da ferramenta → `cp_code` → código do CFF).
3. Pacote novo sem PL: a auditoria acusa ("CRONOGRAMA ATIVIDADES: pacote sem de-para"); completar o de-para (PKG) e regerar.
4. Conferir o esquemático (estrutura sobe de baixo para cima; último quadro = fim do Prevision) e regerar/publicar.

### 6.3 Esquemático (skill `esquematico-avanco`)

Quadros: situação na data de referência do cronograma, 31/12/2026, RETRATO (31/12/2027), 31/12/2028 e 30/04/2029 (fim do
Prevision). Cortes por apartamento antes dos quadros: acabamento (revestimento de parede e piso) até 31/12/2027 e check list
em 2028 e até 30/04/2029. PL 30 quebrada em REV. PEDRAS/BANCADAS, PISO MAD./RODAPÉ e TRAT. PEDRAS (só no esquemático).

### 6.4 Mapas (skill `mapas-obra`)

- **Impermeabilização**: planta do 3º pavimento pintada por período; sistema A começa no 1º dia útil de mar/27; período vai
  da regularização à proteção mecânica.
- **Fachada**: elevações do projeto (estudo R12) pintadas por face e etapa (4º–13º e 14º–28º), cremalheira na frontal
  cobrindo a torre do elevador de obra e 3 vãos de vidro, planta baixa com as faces A (Posterior), B (Lateral direita/Jairo),
  C (Frontal), D (Lateral esquerda/Baltazar).

### 6.5 Incluir obra ou cenário

Criar `ferramenta/obras/<obra>/<pasta>/` (copiar a estrutura do Ed. Jardim; `fluxo_caixa_MODELO_EM_BRANCO.xlsx` explica as
abas de entrada), registrar em `ferramenta/obras.json` e rodar `python3 ferramenta/regerar.py <obra>`.

---

## 7. Regerar e publicar (sempre, a cada mudança)

```bash
python3 ferramenta/regerar.py ed-jardim       # ou sem argumento = todas as obras (~40 min)
```

O script gera, recalcula (LibreOffice) e audita cada cenário (aborta se não der **19 OK** ou se houver erro de fórmula),
monta Gantt, curva física, esquemático e mapas, gera o site, os PDFs e os downloads, e roda `conferir_publicacao.py preparar`.

Publicação conferida (regra do `CLAUDE.md`, sem exceção):

1. `python3 gerador/conferir_publicacao.py preparar` → grava `site/.pendente.json` com os lotes.
2. Publicar cada lote: Artifact publish com `file_path = site/index_artifact.html`, `url` = página do painel, `files` = lote.
   Se recusar por arquivo "não lido", listar os arquivos da página (`scope: files`) e publicar de novo.
3. Reler da página todos os caminhos de `reler` (Artifact read com `paths` e `out_dir`; no máximo ~64 MB por leitura, divida se precisar).
4. `python3 gerador/conferir_publicacao.py conferir <pasta>` → tem que dar "OK: página publicada = arquivos locais".
5. Commit e push (inclui `site/.publicado.json`, as planilhas da raiz e `ferramenta/` se dados mudaram).

O gancho `Stop` impede encerrar o turno com o site diferente do que está confirmado no ar; outro gancho cobra commit/push.

---

## 8. Conferências que o usuário espera ver na resposta

- 2027 de **cada** cenário antes → depois; total projetado; realizado até o corte; e, quando muda dado, a tabela 2026–2029.
- Auditoria 19 OK em todos os cenários; nenhuma fórmula com erro.
- Mudança estética: números idênticos.
- Link do painel e confirmação "conferidos no ar: N de N".

---

## 9. Identidade visual EPO

Logo vetorial (`gerador/marca/epo_logo.svg`, proporção 1,1 : 1); verde-escuro `#0b6938` nas linhas de título (sem trecho
claro); impressão com fundo branco, sem borda externa, quadros lado a lado com a mesma altura; A3 com nome da obra + logo em
toda folha e "Versão de dd/mm/aaaa" no rodapé; desenhos de projeto com traço cinza claro e cores por cima. Planilhas sem
linhas de grade, padrão de verdes EPO, fórmulas de data independentes do idioma do Excel.

---

## 10. Como pedir (exemplos)

- "Atualize o incorrido do Ed. Jardim com este arquivo" (anexar DADOS - ED. JARDIM.xlsx).
- "Carregue este cronograma do Prevision no cenário Prevision Atual" (anexar export de atividades e, se houver, o CFF).
- "Refaça a curva financeira considerando o andamento físico."
- "Avance o corte para out/26" (depois do INCC de outubro).
- "Regere todas as obras e publique."
- "Monte o mapa de fachada do Torre Catharina" (anexar elevações e planta com as faces).

---

## 11. Limites e pendências conhecidas

- Mapas só existem para o Ed. Jardim: as outras obras precisam dos desenhos e cronogramas específicos.
- O esquemático é ilustrativo (pavimento × PL), previsão linear entre as datas do cronograma.
- O deslocamento da curva do Prevision Atual pelo físico ainda é feito sob pedido (não há script dedicado).
- Corte do Botânico em ago/26 (dados de 07/10); Ed. Jardim e Catharina em set/26.
