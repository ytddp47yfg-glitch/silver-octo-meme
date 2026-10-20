#!/bin/bash
# Gera, recalcula e audita a planilha de UM cenário (ferramenta/regerar.py chama este script para todas as obras).
# Uso: [CENARIO=PADRÃO] [ESQ=1 FIS_JSON=... GANTT_JSON=...] bash ferramenta/rebuild.sh <pasta de dados> <saída sem .xlsx>
# Requer python3 + openpyxl e LibreOffice; RECALC = script de recálculo (padrão: recalc.py da skill xlsx do Claude).
set -e
F=$(cd "$(dirname "$0")" && pwd)
export GERADOR=${GERADOR:-$F/../gerador}
RECALC=${RECALC:-/mnt/skills/public/xlsx/scripts/recalc.py}
mkdir -p "$(dirname "$2")"
DATA_DIR=$1 python3 "$F/build_modelo.py" "$2.xlsx" >/dev/null
cp "$2.xlsx" "${2}_calc.xlsx"; CALC=$(realpath "${2}_calc.xlsx")
if [ -f "$RECALC" ]; then (cd "$(dirname "$RECALC")" && python3 "$RECALC" "$CALC" 590 2>/dev/null | grep -E '"status"|total_errors'); else echo "AVISO: sem recálculo automático ($RECALC)"; fi
python3 "$GERADOR/cantos.py" "${2}_calc.xlsx"; python3 "$F/audit.py" "${2}_calc.xlsx" 2>/dev/null | grep -E "FALHA|RESUMO|Desembolso 2027"
