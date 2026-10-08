# Fluxo de caixa das obras EPO: resumo do trabalho (situação em 08/10/2026)

Controle de fluxo de caixa de três obras (Botânico, Ed Jardim, Torre Catharina). Cada obra tem uma planilha-ferramenta
gerada em Python a partir da planilha original ("MODELO 05") e um painel web publicado no claude.ai:
https://claude.ai/artifact/9TpDGLVKmKfKstM6MiEAvy (dashboard, retrato 2027, curva física, Gantt e esquemático; botões
⬇ Excel e ⬇ PDF de cada painel).

## Situação atual por cenário

| Obra | Cenário | Corte | Projeção total | Realizado | Desembolso 2027 |
|---|---|---|---|---|---|
| Botânico | Padrão | ago/26 | R$ 160,42 mi | R$ 21,28 mi | R$ 47,60 mi |
| Botânico | Prevision | ago/26 | R$ 160,42 mi | R$ 21,28 mi | R$ 56,06 mi |
| Ed Jardim | Padrão | set/26 | R$ 154,71 mi | R$ 45,87 mi | R$ 56,01 mi (fixado) |
| Ed Jardim | Prevision | set/26 | R$ 154,71 mi | R$ 45,87 mi | R$ 65,54 mi |
| Ed Jardim | Prevision ajustado | set/26 | R$ 154,71 mi | R$ 45,87 mi | R$ 53,46 mi (fixado) |
| Torre Catharina | Visão Set/26 (antigo Padrão) | set/26 | R$ 183,96 mi | R$ 47,47 mi | R$ 62,18 mi |
| Torre Catharina | Prevision | set/26 | R$ 183,96 mi | R$ 47,47 mi | R$ 67,83 mi |
| Torre Catharina | Simulação (versão 1) | set/26 | R$ 183,96 mi | R$ 47,47 mi | R$ 68,28 mi |

Todos os cenários: auditoria 19/19 OK.

## Como a ferramenta calcula (regras de negócio)

- **Controle em INCC.** Meta de cada PL = orçamento em INCC × fator de projeção (fator IEC, da medição financeira,
  aba ANÁLISE IEC: coluna W ÷ coluna H). A meta nunca fica abaixo do que já foi gasto em INCC (meta mínima = incorrido).
- **Saldo** (meta − incorrido, em INCC) distribuído nos meses depois do corte pela curva do cenário e convertido em R$
  pelo INCC do mês de corte.
  - Padrão: curvas digitadas na planilha original.
  - Prevision: curvas de andamento do cronograma Prevision.
- **Contingência** = orçamento em INCC − total distribuído, lançada em jun/2030. Se der negativa, fica zero
  ("para variações negativas não há contingência").
- **Conciliação com a medição financeira**:
  - O ajuste = incorrido da medição − incorrido do ERP, lançado no mês da medição.
  - O mesmo INCC sai do saldo futuro das PLs (opção A, sem dupla contagem).
  - Está fixada em ago/26, o mês da medição (cfg `conciliacao.mes`), mesmo com o corte em set/26.
- **Incorrido**: aba DADOS do ERP.
  - Filtros: grupo da obra, PL GRUPO G3 = "CUSTOS COM OBRAS", sinal invertido.
  - Exclui projeções carregadas no ERP: "A PAGAR" sem origem.
- **INCC-DI**: série em `incc_hist.json`. Set/2026 = 1.299,742 (0,22% sobre ago/26), derivado da variação publicada.
  Conferir no FGV/IBRE.
- **Ed Jardim, Padrão e Prevision ajustado**: o 2027 é fixado nos valores da revisão anterior (R$ 56.009.788,87 e
  R$ 53.460.558,96).
  - O fator K9 (cfg `desloc`) reduz o INCC de 2027 e joga a diferença para 2028, proporcionalmente.
  - A cada atualização, recalibrar: k = (meta 2027 − Σ C de 2027) ÷ Σ W de 2027 (colunas da aba de cada PL).
- **AUX.ESTRUTURA** (Botânico e Catharina): aba resgatada da planilha original e usada como fonte da curva da ESTRUTURA,
  com as datas vindas do cronograma de atividades.
  - Catharina: custo por trecho no mês de término; retenção técnica de 7%; aço material 88% em 2 parcelas + 12% de corte
    e dobra; pagamento no mês seguinte.
  - Botânico: custo de cada pavimento rateado pelos dias de execução.
- **Simulação (Catharina, versão 1)**: novas datas de término dos trechos de estrutura (equipes Zálem, Jhony e equipe
  única). Ficam na coluna BA "TÉRMINO SIMULADO" da AUX.ESTRUTURA.
  - O 2º e o 3º pav. Olympus foram lidos como 29/01/2027 e 26/02/2027 (na tabela de origem estavam 2026 e 2926).
- **Catharina, "Visão Set/26"**: base na planilha REV SÉRGIO 02, com os fatores de projeção da medição.
  - Os fatores da própria planilha (VALOR POR PL!I ÷ (F × 1,09)) estão guardados em cfg `fator_proj_planilha`. Pendente:
    confirmar com o Sérgio se o 1,09 é o reajuste em INCC (mai/25 → set/26 = 9,10%).
- **Decisões do usuário, por obra:**
  - Catharina: Contenção em 3 parcelas iguais (mar, abr, mai/27).
  - Botânico: Lig. Provisórias em jan, fev e mar/27.
  - Ed Jardim: premissas do Prevision ajustado.
    - Rev. de acabamento em 11 de 28 pavimentos em 2027.
    - Fachada pelo estudo de balancins.
    - Inst. especiais e fundação superficial na curva digitada.
    - Hidráulica com o andamento da elétrica.
- **Mudanças só estéticas não podem alterar números.** A cada build, confira o 2027 de cada cenário e a auditoria (19 OK).

## Painéis e extras

- Curva física por cenário (Ed Jardim), Gantt do cronograma com pacotes "não descritos de forma expressa" na curva do
  Prevision marcados.
- Esquemático por pavimento (Ed Jardim, Prevision ajustado), com três quadros: situação atual, marco de 31/12/2026 e
  retrato de 31/12/2027.
- Visual da planilha no padrão EPO (verdes), sem linhas de grade, gráficos com linhas suavizadas.

## Pendências e próximos passos sugeridos

1. Botânico: levar o corte para set/26 (incorrido de setembro já está nos dados brutos).
2. Confirmar o 1,09 com o Sérgio (Catharina).
3. Conferir o INCC-DI de set/2026 no FGV.
4. Esquemático e painel ESQUEMÁTICO nas demais obras/cenários (só existe no Ed Jardim ajustado).
