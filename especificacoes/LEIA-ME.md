# Grade de especificações · Ed. Cidade Jardim (CDJ-GRE-ARQ-AC)

Ponto de partida para continuar o trabalho em outra sessão ou outro login. Situação em 26/10/2026.

## O que é

Levantamento das especificações de acabamento do Edifício Cidade Jardim, no formato da planilha modelo
CDJ-GRE-ARQ-AC (grade por ambiente: revestimentos, louças e metais, complementares).

- **Planilha atual:** `especificacoes/CDJ-GRE-ARQ-AC-R00.xlsx`. Abas: CONTROLE DE REVISÕES, TÉRREO,
  2° PAVTO, PILOTIS, APTO TIPO, CONFLITOS e PENDÊNCIAS. As imagens dos itens ficam na própria grade,
  com o link interno VER IMAGEM / VOLTAR.
- **Página de visualização:** https://claude.ai/artifact/Y847BF9UXVc7Dssm2etyFw (abas Ambientes, Grade,
  Imagens e Pendências; botão para baixar a planilha). Ela pertence ao login que a criou. Para outro login
  usar, compartilhe pelo menu **Share** da página com permissão de edição.
- **Branch:** `claude/youthful-lovelace-hqvphy`.

## Fontes usadas

| Fonte | Onde está |
|---|---|
| Modelo da grade CDJ-GRE-ARQ-AC-R00 (vazio) | `fonte/entrada/CDJ-GRE-ARQ-AC-R00_MODELO.xlsx` |
| Tabela de acabamentos Ed. Cidade Jardim v00 (19/06/2026) | `fonte/entrada/Tabela_de_acabamentos_Ed._Cidade_Jardim_v00_19jun2026.pdf` |
| Detalhamento de interiores R04 (lista CDJ-AIN-EX-DC001-ESPECIFIC-R04 e pranchas PR001 a PR042) | **não está no repositório**; usado em outra sessão |
| Plantas de arquitetura CDJ-ARQ-EX R11 a R14 | **não está no repositório**; usado em outra sessão |
| Apto tipo: JARDIM-ACABAMENTOS R01 e detalhamento | **não está no repositório**; usado em outra sessão |
| Imagens em alta resolução enviadas pelo usuário | `imagens/CDJ-IMG-01` a `13` (versões finais) e `fonte/gerador_base/imagens_usuario/` (originais) |
| Imagens de referência da web dos produtos | `imagens/CDJ-WEB-*` e `imagens/imagens_produtos_web.json` |

Para refazer as etapas do R04, da arquitetura ou do apto tipo, é preciso enviar de novo os documentos que não
estão no repositório.

## Como a planilha foi montada

1. **Grade base (Pilotis)**: `fonte/gerador_base/build.py` lê o modelo e os dados extraídos da tabela v00
   (`recs.json`) e gera a grade R00 só com o Pilotis, a aba de imagens e as pendências.
   Uso, a partir da raiz do repositório: `python3 especificacoes/fonte/gerador_base/build.py`.
   O resultado vai para `fonte/gerador_base/saida/` e **não** sobrescreve a planilha atual.
2. **Etapas seguintes** (scripts em `especificacoes/gerador/`, cada um altera a planilha recebida como argumento):
   `levantamento_r04.py` (detalhamento R04), `arquitetura_grade.py`, `arq_*.py` e `terraco_salao.py`
   (plantas de arquitetura), `apto_tipo.py` (aba APTO TIPO), `conflitos*.py` (aba CONFLITOS),
   `imagens_links.py` (imagens na grade), `formatacao.py`, `area_impressao.py` e `previa_impressao.py`
   (impressão) e `revisao_geral.py` (textos, sumário e pendências).
   Os registros de cada etapa estão em `gerador/*_log.json` e `gerador/conferencia_etiquetas_arquitetura.csv`.
3. **Página**: o HTML publicado está em `fonte/pagina/grade-cidade-jardim_publicada.html`.
   `gerador/pagina_imagens_web.py` atualiza esse HTML com as imagens web e a planilha embutida.
   Para republicar no mesmo link, use o Artifact com a URL da página.
   A página tem um banco de dados (coleção `imagens`) e um armazenamento de arquivos, usados pelo botão
   "Enviar imagem em alta resolução" da aba Imagens.

Dependências: Python 3 com `openpyxl` e `Pillow`; LibreOffice só para renderizar ou conferir a planilha.

## O que falta

- **Pendências de especificação**: aba PENDÊNCIAS, seção 1 (17 itens). Principais: qual é o "mármore branco
  selecionado"; rejunte do mosaico New Sterlizia; espessura do PA48; forro T7 "a definir"; bacia Deca Carrara
  ou Kohler Panache e o assento; código do chuveiro D31; tons "a definir" das quadras.
- **Conflitos entre projetos**: aba CONFLITOS (8 células). Exemplos: Guarita (piso, parede), teto do Concierge,
  bancada e torneira dos vestiários do Pilotis, rodapé do apto tipo. Cada um aguarda a definição da arquitetura.
- **Ambientes sem especificação em nenhuma fonte**: aba PENDÊNCIAS, seção 2. São Circulação Salão de Festas,
  Apoio Cozinha, Circulação Lavabos, Playground, Circulação Massagem/Beleza e Lounge Quadras.
- **Imagens**: os 13 itens das fichas Termas e Vestiário Termas já têm imagem do usuário (CDJ-IMG-01 a 13).
  As imagens web são de referência. As marcadas "conferir" são de modelo aproximado.

## Observação sobre rede

O ambiente de nuvem usado até aqui bloqueava os sites dos fabricantes (Deca, Docol, Kohler, Roca, Bisazza,
lojas). Na sessão nova, libere esses domínios em **Network access** antes de criar a sessão, se for preciso
baixar imagens ou fichas técnicas.
