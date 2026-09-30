# Ed Jardim — levantamento de especificações do fluxo de caixa

Levantamento das especificações (parâmetros, regras de cálculo, premissas por PL e fontes de dados) usadas nos três cenários do fluxo de caixa do Ed Jardim, lido diretamente das planilhas do repositório:

| Cenário | Arquivo | JSON do site |
|---|---|---|
| Padrão | `fluxo_caixa_ed_jardim_modelo_PADRAO.xlsx` | `site/obras/ed-jardim-padrao.json` |
| Prevision | `fluxo_caixa_ed_jardim_modelo.xlsx` | `site/obras/ed-jardim.json` |
| Prevision ajustado | `fluxo_caixa_ed_jardim_modelo_PREVISION_AJUSTADO.xlsx` | `site/obras/ed-jardim-ajustado.json` |

Valores em R$ nominais (com taxa de administração), salvo indicação de INCC.

## 1. Parâmetros gerais

| Parâmetro | Valor | Onde |
|---|---|---|
| Orçamento base | Orçamento maio/25 — Ed Jardim (INCC 10-25) | `VALOR POR PL!B3` |
| Orçamento total (R$ base) | R$ 142.983.236,35 | `VALOR POR PL!F50` |
| Orçamento total (INCC) | 121.324,93 pontos | `VALOR POR PL!G50` |
| INCC base do orçamento | 1.178,3886 | `VALOR POR PL!H6` |
| Taxa de administração orçada | R$ 2.075.142,69 / 1.748 INCC (1,47%) | `VALOR POR PL!F7:G7` |
| Data de corte | ago/2026 | `VALOR POR PL!K5` |
| Início do fluxo | dez/2024 | `VALOR POR PL!K6` |
| Último mês do horizonte | nov/2030 | `VALOR POR PL!P8` |
| INCC no mês de corte | 1.296,889 | `VALOR POR PL!P7` |
| Limite de avanço físico p/ usar o Prevision | 100% (todo item direto com avanço < 100% pode seguir o Prevision) | `VALOR POR PL!K7` |
| Fator de redução do saldo (conciliação) | 0,992238 | `VALOR POR PL!G65` |
| Ano do retrato | 2027 (situação em 31/12/2027) | `RETRATO 2027!D5` |

### Parâmetros que variam por cenário

| Parâmetro | Padrão | Prevision | Prevision ajustado |
|---|---|---|---|
| Cenário das curvas (`VALOR POR PL!K8`) | PADRÃO | PREVISION | PREVISION |
| Fator sobre a previsão de 2027 (`VALOR POR PL!K9`) | 0,988190 | — (sem fator) | 0,991181 |

## 2. Resultado por cenário

| Indicador | Padrão | Prevision | Prevision ajustado |
|---|---|---|---|
| Projeção total (com contingência) | R$ 154.459.552,47 | R$ 154.459.552,47 | R$ 154.459.552,47 |
| Realizado até o corte (medição) | R$ 42.586.104,68 | R$ 42.586.104,68 | R$ 42.586.104,68 |
| Desembolso em 2027 | R$ 56.009.788,87 | R$ 64.514.201,29 | R$ 53.460.558,96 |
| Saldo a realizar após o corte | R$ 111.873.447,79 | R$ 111.873.447,79 | R$ 111.873.447,79 |
| Conferência | ✓ Tudo OK | ✓ Tudo OK | ✓ Tudo OK |

### Desembolso anual (R$)

| Ano | Padrão | Prevision | Prevision ajustado |
|---|---|---|---|
| 2024 | 5.778.713,93 | 5.778.713,93 | 5.778.713,93 |
| 2025 | 18.489.705,84 | 18.489.705,84 | 18.489.705,84 |
| 2026 | 34.809.925,84 | 35.516.870,27 | 34.039.717,25 |
| 2027 | 56.009.788,87 | 64.514.201,29 | 53.460.558,96 |
| 2028 | 31.766.718,43 | 26.393.249,86 | 33.826.476,73 |
| 2029 | 4.883.554,60 | 1.045.666,32 | 6.143.234,81 |
| 2030 | 2.721.144,95 | 2.721.144,95 | 2.721.144,95 |
| **Total** | **154.459.552,47** | **154.459.552,47** | **154.459.552,47** |

2024 e 2025 são integralmente realizados; em 2026 o realizado até o corte é R$ 18.317.684,91 e o restante é previsto. 2030 contém a contingência (jun/2030).

## 3. Regras de cálculo

1. **Fluxo controlado em INCC.** A meta de cada PL é um valor em INCC. O saldo (meta − incorrido em INCC) é distribuído pela curva escolhida e convertido em R$ pelo INCC do corte (1.296,889). A projeção em R$ é `incorrido + saldo INCC × INCC do corte`.
2. **Orçamento em INCC.** Coluna G de VALOR POR PL = INCC do orçamento original (total 121.324,93).
3. **Reorçamento (Análise IEC).** Fator da projeção (col. S) = projeção ÷ orçamento reajustado (ANÁLISE IEC: W ÷ H). Meta em INCC = INCC orçado × fator. Equipamentos e ferramentas ficou com fator 1 (a fórmula da planilha de controle estava errada; confirmado pelo usuário).
4. **PL que já gastou acima da projeção fecha no realizado.** Meta = MAX(INCC orçado × fator; incorrido em INCC).
5. **Realizado.** Vem do ERP (DADOS BRUTOS: Dinâmica___Valor, custos com obras, sinal invertido); os meses anteriores ao horizonte são convertidos pelo INCC de cada mês (base INCC-DI desde ago/1994).
6. **Conciliação com a medição financeira.** Incorrido informado pela medição: R$ 42.586.104,68 / 35.062 INCC; ERP: R$ 41.898.485,13 / 34.438,89 INCC. O ajuste (R$ 687.619,55 / 623,11 INCC) entra no MODELO PARA BI no mês de corte e **sai do saldo futuro das PLs** (fator 0,992238 sobre os saldos positivos, `K10` de cada aba), sem dupla contagem.
7. **Contingência.** MAX(0; orçamento em INCC − soma distribuída), lançada em jun/2030: 2.098,21 INCC (≈ R$ 2,72 mi). O total em INCC do MODELO PARA BI fica igual ao orçamento.
8. **Taxa de administração.** Somente o incorrido no ERP (meta = incorrido em INCC), sem previsão futura; projeção R$ 2.075.142,69.
9. **Escolha da curva de cada PL (col. O / aba K16).**
   - Indiretos (01 a 11): previsão informada (curva digitada).
   - Diretos: no cenário **Padrão**, curva digitada; nos cenários **Prevision**, curva do Prevision quando o avanço físico no corte é menor que o limite (K7 = 100%); acima do limite, previsão informada.
   - Qualquer PL pode ter a fonte forçada em `K15` da aba (ver seção 5).
   - `VALOR POR PL!R` permite a um PL seguir a curva prevista de outro item do Prevision (o realizado continua o próprio).
10. **Lançamentos do ERP após o corte** (`K22` da aba): SIM por padrão; NÃO em DIVERSOS.
11. **Fator de 2027 (K9), cenários Padrão e Prevision ajustado.** A previsão de 2027 de cada PL é multiplicada pelo fator e a parte retirada vai para 2028, proporcional à curva da PL em 2028, mantendo o total da obra. Serve para fixar o 2027 nos valores da revisão anterior: R$ 56.009.788,87 (Padrão) e R$ 53.460.558,96 (Prevision ajustado). O fator precisa ser recalibrado sempre que os DADOS BRUTOS mudam.

## 4. Especificação por PL

Fonte da curva por cenário: PAD = Padrão, PRV = Prevision, AJU = Prevision ajustado. *Digitada* = previsão informada (col. H da aba).

| Cód. | PL | Tipo | Orçado R$ (base) | Orçado INCC | Fator reorç. | Meta INCC | Projeção R$ | Físico no corte | Cód. Prevision | Curva PAD | Curva PRV | Curva AJU |
|---|---|---|---:|---:|---:|---:|---:|---:|---|---|---|---|
| 01 | PROJETOS | INDIRETO | 2.720.799,48 | 2.308,92 | 0,9853 | 2.274,91 | 2.560.901,01 | 0,0% | 01 | Digitada | Digitada | Digitada |
| 02 | SEGURO DE OBRA | INDIRETO | 122.200,00 | 103,70 | 0,9247 | 95,89 | 112.496,08 | 0,0% | 02 | Digitada | Digitada | Digitada |
| 03 | TAXAS DIVERSAS | INDIRETO | 150.000,00 | 127,29 | 1 | 127,29 | 161.870,94 | 0,0% | 03 | Digitada | Digitada | Digitada |
| 04 | BARRACAO DE OBRA | INDIRETO | 995.153,99 | 844,50 | 1,1497 | 970,91 | 1.189.305,88 | 0,0% | 04 | Digitada | Digitada | Digitada |
| 05 | LIGAÇÕES PROVISÓRIAS AGUA, ESGOTO, ENERGIA,TELEFONE, INTERNET | INDIRETO | 50.000,00 | 42,43 | 1 | 42,43 | 54.161,04 | 0,0% | 05 | Digitada | Digitada | Digitada |
| 06 | CONSUMO DE ÁGUA, ENERGIA, TELEFONE, INTERNET | INDIRETO | 385.000,00 | 326,72 | 1,3754 | 449,37 | 566.543,37 | 0,0% | 06 | Digitada | Digitada | Digitada |
| 07 | DIVERSOS | INDIRETO | 131.000,00 | 111,17 | 1,9129 | 212,66 | 264.678,70 | 0,0% | 07 | Digitada | Digitada | Digitada |
| 08 | FRETES E TRANSPORTES | INDIRETO | 742.720,00 | 630,28 | 0,9936 | 626,24 | 797.227,41 | 0,0% | 08 | Digitada | Digitada | Digitada |
| 09 | CONSULTORIAS (OBRAS) | INDIRETO | 805.441,86 | 683,51 | 1,0510 | 718,38 | 900.329,29 | 0,0% | 09 | Digitada | Digitada | Digitada |
| 10 | MÃO DE OBRA PRÓPRIA | INDIRETO | 14.511.646,08 | 12.314,82 | 1,0717 | 13.198,23 | 16.647.692,28 | 0,0% | 10 | Digitada | Digitada | Digitada |
| 11 | EQUIPAMENTOS E FERRAMENTAS | INDIRETO | 4.368.148,80 | 3.706,88 | 1 | 3.706,88 | 4.725.154,59 | 0,0% | 11 | Digitada | Digitada | Digitada |
| 12 | SERVIÇOS INICIAIS | DIRETO | 595.242,68 | 505,13 | 0,9498 | 489,09 | 558.350,21 | 90,0% | 12 | Digitada | Prevision | Prevision |
| 13 | CONTENÇÕES | DIRETO | 1.384.152,81 | 1.174,61 | 0,8603 | 1.071,60 | 1.252.186,97 | 100,0% | 13 | Digitada | Digitada | Digitada |
| 14 | TERRAPLENAGEM | DIRETO | 2.940.699,35 | 2.495,53 | 0,8079 | 2.036,81 | 2.371.480,27 | 89,0% | 14 | Digitada | Prevision | Prevision |
| 15 | FUNDAÇÃO PROFUNDA | DIRETO | 5.345.899,12 | 4.536,62 | 0,7565 | 3.672,58 | 4.368.243,12 | 100,0% | 15 | Digitada | Digitada | Digitada |
| 16 | FUNDAÇÃO SUPERFICIAL | DIRETO | 4.052.793,15 | 3.439,27 | 0,7933 | 2.728,51 | 3.371.658,86 | 88,2% | 16 | Digitada | Prevision | Digitada |
| 17 | ESTRUTURA EM CONCRETO | DIRETO | 20.273.268,87 | 17.204,23 | 0,9753 | 16.778,97 | 21.237.887,95 | 67,0% | 17 | Digitada | Digitada | Digitada |
| 18 | ALVENARIA | DIRETO | 3.009.694,16 | 2.554,08 | 1,1036 | 2.818,72 | 3.626.816,03 | 31,0% | 19 | Digitada | Prevision | Prevision |
| 19 | IMPERMEABILIZAÇÃO | DIRETO | 3.799.651,26 | 3.224,45 | 0,7924 | 2.555,21 | 3.290.324,64 | 0,0% | 20 | Digitada | Prevision | Prevision |
| 20 | INSTALAÇÕES ELÉTRICAS | DIRETO | 8.532.606,85 | 7.240,91 | 1,0014 | 7.251,39 | 9.322.902,37 | 8,0% | 21 | Digitada | Prevision | Prevision |
| 21 | INSTALAÇÕES HIDRÁULICAS | DIRETO | 5.180.198,40 | 4.396,00 | 0,9918 | 4.360,03 | 5.602.741,92 | 0,0% | 22 (curva do 21) | Digitada | Prevision | Prevision |
| 22 | INSTALAÇÕES COMBATE INCÊNDIO | DIRETO | 647.524,80 | 549,50 | 1,0976 | 603,13 | 776.378,60 | 0,0% | 23 | Digitada | Prevision | Prevision |
| 23 | INSTALAÇÕES ESPECIAIS | DIRETO | 5.085.820,25 | 4.315,91 | 1 | 4.315,91 | 5.553.813,78 | 0,0% | 24 | Digitada | Prevision | Digitada |
| 24 | ELEVADORES | DIRETO | 2.639.985,62 | 2.240,34 | 0,8362 | 1.873,26 | 2.413.916,17 | 0,0% | 25 | Digitada | Digitada | Digitada |
| 25 | ESQUADRIAS | DIRETO | 12.959.647,29 | 10.997,77 | 1 | 10.997,77 | 14.152.236,60 | 0,0% | 26 | Digitada | Digitada | Digitada |
| 26 | SERRALHERIA | DIRETO | 445.157,53 | 377,77 | 1 | 377,77 | 486.166,02 | 0,0% | 27 | Digitada | Prevision | Prevision |
| 27 | MARCENARIA | DIRETO | 1.728.851,77 | 1.467,13 | 1 | 1.467,13 | 1.887.939,47 | 0,0% | 28 | Digitada | Prevision | Prevision |
| 28 | REVESTIMENTOS DE MASSA EM PISOS | DIRETO | 1.122.595,50 | 952,65 | 1 | 952,65 | 1.225.895,93 | 0,5% | 29 | Digitada | Prevision | Prevision |
| 29 | REVESTIMENTOS DE MASSA EM PAREDES | DIRETO | 2.524.242,57 | 2.142,11 | 1 | 2.142,11 | 2.756.537,27 | 15,1% | 30 | Digitada | Prevision | Prevision |
| 30 | REVESTIMENTOS DE ACABAMENTO | DIRETO | 17.569.932,69 | 14.910,13 | 1 | 14.910,13 | 19.186.705,30 | 0,0% | 31 | Digitada | Prevision | Digitada |
| 31 | ESTRUTURA METÁLICA | DIRETO | 195.186,00 | 165,64 | 1 | 165,64 | 213.146,88 | 0,0% | 18 | Digitada | Prevision | Prevision |
| 32 | REVESTIMENTO DIVERSOS - FACHADA | DIRETO | 7.404.956,91 | 6.283,97 | 1 | 6.283,97 | 8.088.432,56 | 0,0% | 32 | Digitada | Prevision | Digitada |
| 33 | FORROS DIVERSOS | DIRETO | 2.024.194,97 | 1.717,77 | 1 | 1.717,77 | 2.210.459,94 | 0,0% | 33 | Digitada | Prevision | Prevision |
| 34 | LOUÇAS E METAIS | DIRETO | 1.131.097,39 | 959,87 | 1 | 959,87 | 1.235.180,16 | 0,0% | 34 | MDO Prev. + material premissa | MDO Prev. + material premissa | MDO Prev. + material premissa |
| 35 | PINTURAS | DIRETO | 2.190.870,85 | 1.859,21 | 1 | 1.859,21 | 2.391.424,41 | 5,9% | 35 | Digitada | Prevision | Prevision |
| 36 | JARDINEIRAS E PAISAGISMO | DIRETO | 2.411.990,63 | 2.046,86 | 1 | 2.046,86 | 2.617.588,75 | 11,4% | 36 | Digitada | Prevision | Prevision |
| 37 | LIMPEZA FINAL DE OBRA | DIRETO | 729.722,03 | 619,25 | 1 | 619,25 | 796.870,53 | 0,0% | 37 | Digitada | Prevision | Prevision |
| 38 | TAXA DE ADMINISTRAÇÃO | INDIRETO | 2.075.142,69 | 1.748,00 | 1 | 1.748,16 | 2.075.142,69 | 0,0% | 00 | Sem curva | Sem curva | Sem curva |
| | **Total** | | **142.983.236,35** | **121.324,93** | | **119.226,72** | **151.050.787,97** | | | | | |

A diferença entre a soma das projeções das PLs e a projeção total da obra é o ajuste de conciliação mais a contingência.

### Desembolso em 2027 por PL (R$)

| Cód. | PL | Padrão | Prevision | Prevision ajustado |
|---|---|---:|---:|---:|
| 01 | PROJETOS | 290.943,73 | 294.420,99 | 291.824,55 |
| 02 | SEGURO DE OBRA | 8.595,56 | 8.698,29 | 8.621,58 |
| 03 | TAXAS DIVERSAS | 109.358,18 | 110.665,19 | 109.689,25 |
| 04 | BARRACAO DE OBRA | 118.064,06 | 119.372,65 | 118.395,54 |
| 05 | LIGAÇÕES PROVISÓRIAS AGUA, ESGOTO, ENERGIA,TELEFONE, INTERNET | 11.145,29 | 11.278,49 | 11.179,03 |
| 06 | CONSUMO DE ÁGUA, ENERGIA, TELEFONE, INTERNET | 176.988,78 | 179.104,09 | 177.524,61 |
| 07 | DIVERSOS | 38.632,92 | 39.094,65 | 38.749,88 |
| 08 | FRETES E TRANSPORTES | 468.511,32 | 474.110,80 | 469.929,71 |
| 09 | CONSULTORIAS (OBRAS) | 336.713,10 | 340.737,37 | 337.732,48 |
| 10 | MÃO DE OBRA PRÓPRIA | 4.906.247,67 | 4.964.606,10 | 4.921.030,34 |
| 11 | EQUIPAMENTOS E FERRAMENTAS | 1.127.225,97 | 1.140.583,45 | 1.130.609,53 |
| 12 | SERVIÇOS INICIAIS | 24.316,70 | 24.316,70 | 24.316,70 |
| 13 | CONTENÇÕES | 0,00 | 0,00 | 0,00 |
| 14 | TERRAPLENAGEM | 0,00 | 0,00 | 0,00 |
| 15 | FUNDAÇÃO PROFUNDA | 0,00 | 0,00 | 0,00 |
| 16 | FUNDAÇÃO SUPERFICIAL | 227.082,38 | 873.692,29 | 227.744,98 |
| 17 | ESTRUTURA EM CONCRETO | 4.008.854,12 | 4.055.950,15 | 4.020.783,93 |
| 18 | ALVENARIA | 1.684.542,15 | 1.543.385,97 | 1.529.964,00 |
| 19 | IMPERMEABILIZAÇÃO | 2.452.193,35 | 2.705.057,88 | 2.681.217,05 |
| 20 | INSTALAÇÕES ELÉTRICAS | 3.721.250,22 | 3.939.404,90 | 3.904.740,64 |
| 21 | INSTALAÇÕES HIDRÁULICAS | 3.492.620,83 | 3.975.305,32 | 2.354.569,58 |
| 22 | INSTALAÇÕES COMBATE INCÊNDIO | 561.605,79 | 443.574,01 | 439.662,22 |
| 23 | INSTALAÇÕES ESPECIAIS | 3.542.348,48 | 4.593.004,00 | 3.553.072,76 |
| 24 | ELEVADORES | 764.100,17 | 773.232,43 | 766.413,45 |
| 25 | ESQUADRIAS | 7.059.182,25 | 7.143.551,08 | 7.080.553,57 |
| 26 | SERRALHERIA | 126.967,07 | 24.111,86 | 23.899,23 |
| 27 | MARCENARIA | 0,00 | 688.720,32 | 682.646,64 |
| 28 | REVESTIMENTOS DE MASSA EM PISOS | 1.066.358,60 | 986.801,88 | 978.099,47 |
| 29 | REVESTIMENTOS DE MASSA EM PAREDES | 1.637.162,17 | 1.570.989,55 | 1.557.307,36 |
| 30 | REVESTIMENTOS DE ACABAMENTO | 10.132.172,11 | 14.144.439,14 | 7.471.161,36 |
| 31 | ESTRUTURA METÁLICA | 0,00 | 0,00 | 0,00 |
| 32 | REVESTIMENTO DIVERSOS - FACHADA | 4.059.138,07 | 5.581.307,66 | 4.817.501,12 |
| 33 | FORROS DIVERSOS | 873.741,33 | 924.856,44 | 916.700,32 |
| 34 | LOUÇAS E METAIS | 369.644,11 | 374.061,96 | 370.763,18 |
| 35 | PINTURAS | 1.231.138,61 | 1.143.604,89 | 1.133.583,81 |
| 36 | JARDINEIRAS E PAISAGISMO | 1.244.429,74 | 1.248.450,28 | 1.237.510,61 |
| 37 | LIMPEZA FINAL DE OBRA | 138.514,05 | 73.710,52 | 73.060,49 |
| 38 | TAXA DE ADMINISTRAÇÃO | 0,00 | 0,00 | 0,00 |
| | **Total** | **56.009.788,87** | **64.514.201,29** | **53.460.558,96** |

## 5. Premissas especiais por PL

| PL | Cenários | Especificação |
|---|---|---|
| Seguros; Lig. provisórias | todos | Sem aba no arquivo original: projeção = orçamento reajustado pelo INCC do corte; saldo em curva linear até dez/2028 (premissa dos indiretos). **A confirmar.** |
| Diversos | todos | ERP tinha R$ 45.630 lançados para set-out/26, acima da projeção. Vale só o realizado até o corte; o saldo até a projeção vai pela curva digitada (K22 = NÃO). |
| Equipamentos e ferramentas | todos | Reorçamento com fator 1 (erro de fórmula na ANÁLISE IEC; confirmado pelo usuário). |
| Estrutura em concreto | todos | Curva digitada forçada também nos cenários Prevision (decisão do usuário). |
| Elevadores | todos | Curva digitada forçada. Premissa: pedido 1 ano antes, 12 pagamentos. |
| Esquadrias | todos | Curva digitada forçada. Premissa: 60% em 10 pagamentos + 40% por medição direta. |
| Louças e metais | todos | Saldo dividido: 40% MDO pela curva do Prevision; 60% material pela premissa — pedido 180 dias antes da instalação, parcelas a 28/56/84 dias do pedido (pagas 5, 4 e 3 meses antes da instalação). No fluxo: MDO R$ 494.072,06; material R$ 741.108,10. |
| Rev. de pisos / paredes | todos | Correção do original: ambos liam a mesma linha do ERP (R$ 380.967 contados duas vezes); o valor fica só em PAREDES. |
| Estrutura metálica | todos | PL 31 do Ed Jardim (no Botânico é Pavimentação); item 18 no Prevision. Início previsto em abr/2028. |
| Taxa de administração | todos | Somente incorrido; sem previsão futura. |
| Fundação superficial | AJU | Volta para a curva digitada (simulação pedida pelo usuário). |
| Instalações especiais | AJU | Volta para a curva digitada (simulação pedida pelo usuário). |
| Instalações hidrossanitárias | AJU | Curva prevista igual ao andamento das instalações elétricas no Prevision (VALOR POR PL!R = 21); o realizado continua o da hidráulica. |
| Revestimentos de acabamento | AJU | 11 dos 28 pavimentos em 2027 (39,3% do valor, R$ 6,90 mi) no formato mensal do Prevision; os 17 restantes a ≈0,92 pav./mês (R$ 575 mil/mês) de jan/2028 a jul/2029. RETRATO: execução física 2027 = avanço financeiro + 0,7 p.p.; pavimentos em 2027: Pilotis (parcial), 700, 900, 1000, 1200 e 1800 (liberados para finalizar); 1600, 1700, 1900, 2000 e 2500 (em negociação). |
| Revestimento de fachada | AJU | Curva pelo Estudo de fachada R12 (cenário 07): m² de cada frente de balancim pelos dias úteis (seg-sáb), out/2026 a abr/2028 — 8,5% em 2026, 62,8% em 2027, 28,7% em 2028. O custo do estudo é só dos balancins; a curva dá o formato do saldo. RETRATO: execução física = financeiro + 0,7 p.p. |

### Premissas gerais registradas na aba PREMISSAS

- MDO e equipamentos: histograma.
- Indiretos: curva linear.
- Atividades diretas: andamento do Prevision.
- Elevador: pedido 1 ano antes, 12 pagamentos.
- Esquadrias (ex.: Confiança): 60% em 10 pagamentos (modelo Astro) + 40% por medição direta.
- Louças e metais: pedido 40 a 180 dias antes; pagamento 28/56/84.
- Acabamentos: pagamento em 3×.
- A aba ainda diz "taxa de administração é 10% sobre o custo do mês", mas o modelo usa só o incorrido (regra 8 da seção 3).

## 6. Fontes de dados

| Dado | Fonte | Referência |
|---|---|---|
| Orçamento por PL | Orçamento maio/25 (MODELO 05 do Ed Jardim) | VALOR POR PL |
| Realizado | ERP — Dinâmica___Valor, custos com obras (41 grupos, jul/2015 a jun/2029, total R$ 48.653.312,02; inclui compromissos após o corte) | DADOS BRUTOS |
| Incorrido conciliado | Medição financeira: R$ 42.586.104,68 / 35.062 INCC | VALOR POR PL!E63:E64 |
| Reorçamentos | ANÁLISE IEC (W ÷ H), 20 PLs com fator | VALOR POR PL!S |
| Curva e avanço físico | Prevision — cronograma físico-financeiro de 28/09/2026 (desde ago/2024, 55 meses) | CURVA PREVISION, CP RESUMO |
| Cronograma de atividades | Prevision — export de 28/09/2026 (2.253 atividades, de-para pacote → PL) | CRONOGRAMA ATIVIDADES |
| Índice | INCC-DI (FGV), série oficial desde ago/1994 | "INCC-FGV" |
| Curva da fachada (AJU) | Estudo de balancins R12, cenário 07 | aba REV. FACHADA |

## 7. Painéis e saídas

- **Site** (`site/index.html`): cartão do Ed Jardim com os 3 cenários; painéis DASHBOARD, RETRATO 2027, CURVA FÍSICA e GANTT, mais todas as abas.
- **Curva física** (`gerador/fisico.py`): avanço físico por PL (realizado e previsto do Prevision; cenário = físico no corte + desembolso restante), obra ponderada pela projeção dos itens diretos.
- **Gantt** (`gerador/gantt.py`): árvore PL → pacote → pavimento a partir do CRONOGRAMA ATIVIDADES; 42 pacotes do Ed Jardim com associação inferida (sem correspondência expressa na EAP do Prevision), sinalizados com nota.
- **PDF** (`site/pdf/ed-jardim*.pdf`): A4 paisagem com todos os painéis do cenário.
- **Planilhas**: padrão estético da EPO (paleta da medição financeira, Calibri, sem linhas de grade, gráficos suavizados com cantos arredondados). A versão Prevision ajustado tem ainda as abas CURVA FÍSICA e GANTT.

## 8. Pontos de atenção / a confirmar

1. **Seguros e Ligações provisórias**: projeção e curva linear até dez/2028 marcadas como "Confirmar" nas próprias abas.
2. **Fator de 2027 (K9)**: depende dos DADOS BRUTOS; a última atualização do ERP (compromissos de set-nov/2026) já exigiu recalibração (0,988190 Padrão / 0,991181 Prevision ajustado). Qualquer nova carga do ERP exige recalibrar para manter o 2027 fixado.
3. **Cenário Prevision sem fator de 2027**: o cenário Prevision puro não tem K9, por isso seu 2027 (R$ 64.514.201,29) fica bem acima dos outros dois.
4. **Aba PREMISSAS desatualizada** sobre a taxa de administração (10% do custo do mês × somente incorrido no modelo).
5. **Rótulo "INCORRIDO ATÉ" no MODELO PARA BI (C64)**: usa `TEXT(...;"mm/aaaa")`, formato em português. Recalculado no LibreOffice em inglês, aparece como "05/Shabbat"; no Excel em pt-BR mostra "08/2026". Só afeta o texto do rótulo, não os valores.
6. **Gantt**: 42 pacotes com associação PL ↔ pacote inferida; vale validar o de-para com o planejamento.
