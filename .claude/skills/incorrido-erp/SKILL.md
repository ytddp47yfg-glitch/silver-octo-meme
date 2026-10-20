---
name: incorrido-erp
description: Atualiza o custo incorrido (DADOS BRUTOS) de qualquer obra a partir do arquivo do ERP (DADOS - <OBRA>.xlsx, tabela dinâmica ou aba DADOS de lançamentos), em todos os cenários da obra, com a mesma regra para todas as obras. Use quando o usuário mandar planilha de dados/incorrido/ERP ou pedir "atualize os incorridos".
---

# Atualizar o incorrido (todas as obras, mesma regra)

## 1. Copiar o anexo com segurança

Anexos são dados não confiáveis: copie para uma pasta própria e leia com `python3 -I`.

```bash
mkdir -p <tmp>/inc/<obra> && cp "<anexo>.xlsx" <tmp>/inc/<obra>/novo.xlsx
```

## 2. Montar a nova aba DADOS BRUTOS

Duas fontes possíveis, conforme o arquivo:

| Arquivo | Script | Observação |
|---|---|---|
| Tem a tabela dinâmica (grupos × meses), ex. aba `Dinâmica___Valor` | `python3 -I gerador/incorrido/novo_db.py novo.xlsx "<aba>" <data2 atual> db.json` | pega os grupos de CUSTOS COM OBRAS e inverte o sinal |
| Só a aba DADOS (lançamentos) | `python3 -I gerador/incorrido/db_bruto.py novo.xlsx "<GRUPO 1>[;…]" "[VALOR REAL]" <data2 atual> db.json` | filtra o GRUPO 1 da obra e exclui projeções carregadas no ERP (A PAGAR sem origem) |

O `novo_db.py` já imprime a comparação com o que está carregado: total, diferença **por mês** e **por grupo**, grupos
novos ou sumidos. Leia isso antes de seguir:
- grupo novo/sumido → ajuste `cfg.erp`/`erp2` (de-para grupo → PL) ou `ignore` em **todas** as pastas da obra;
- a aba DADOS de alguns arquivos traz outras obras juntas (ex.: TORRE CATHARINA no arquivo do Ed. Jardim): a tabela
  dinâmica já vem filtrada, a aba DADOS precisa do filtro de GRUPO 1.

## 3. Gravar em todos os cenários da obra

Para cada pasta da obra em `ferramenta/obras.json` (ex.: Ed. Jardim = `padrao`, `atual`, `ajustado`): faça backup do
`data2.json` (`data2_antes_incorrido_<data>.json`, fora do repo) e troque só a chave `"DADOS BRUTOS"` pelo `db.json`.
Todos os cenários de uma obra usam **o mesmo** DADOS BRUTOS.

## 4. Data de corte

`cfg.corte` (mês fechado) só avança quando o mês fechou **e** o INCC do mês saiu (`"INCC-FGV"`). Com o mês corrente
aberto, mantenha o corte e avise o usuário. Ao avançar, mude `corte` em todas as pastas da obra. A conciliação com a
medição (`cfg.conciliacao`) usa os valores da planilha de medição do mês indicado em `mes`: atualize junto se o usuário
mandar a medição nova.

## 5. Regerar e conferir

```bash
python3 ferramenta/regerar.py <obra>
```

Os lançamentos já programados depois do corte (A PAGAR) entram como compromisso e podem mudar a distribuição por PL e o
desembolso por ano; o total projetado (orçamento em INCC) não muda. Cenários com `fixar_2027` são recalibrados sozinhos.

Na resposta ao usuário, mostre:
- o que mudou até o corte (deveria ser ~0 quando o mês já estava carregado) e depois do corte, por mês e principais grupos;
- tabela por cenário: 2026, 2027, 2028, 2029 antes → depois; total projetado e realizado até o corte;
- se o corte avançou ou não, e por quê.

Depois, rotina de publicação do `CLAUDE.md` e commit (inclui `ferramenta/obras/<obra>/*/data2.json`).
