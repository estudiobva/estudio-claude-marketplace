#!/bin/bash
# revisar-cruce.sh: corre el cruce de ventas (solapa "cruce ventas") de todos los periodos de una
# sociedad y resume cada mes: Zetti - Portal IVA, comprobantes de otros periodos, estados y las
# observaciones que NO son corrimientos ni redondeos (las que hay que mirar).
#
#   bash scripts/revisar-cruce.sh <CUIT> [AAAAMM ...]      (sin periodos: todos los de la subcarpeta)
#   EXCEPCION="2:NC B": la misma excepcion de depuracion que en correr-periodos.sh.
#   SUBCARPETA=controles por defecto, igual que correr-periodos.sh.
#
# Correrlo cuando ya estan armados todos los meses: el cruce mira los meses vecinos para
# separar los cortes de mes de las diferencias reales.
CUIT=$(echo "$1" | tr -cd '0-9'); shift
[ ${#CUIT} = 11 ] || { echo "Uso: bash revisar-cruce.sh <CUIT> [AAAAMM ...]"; exit 2; }
SUBCARPETA=${SUBCARPETA-controles}
AQUI="$(cd "$(dirname "$0")" && pwd)"
PY=${BVA_PYTHON:-python}
export PYTHONIOENCODING=utf-8
SOCDIR=$(node "$AQUI/carpeta-sociedad.js" "$CUIT") || exit 4
SOCDIR=$(cygpath -u "$SOCDIR" 2>/dev/null || echo "$SOCDIR")
SOC=$(basename "$SOCDIR"); FARM=$(echo "$SOC" | sed -E "s/^ *[0-9]+ *- *//; s/ *- *$CUIT *$//")
BASE="$SOCDIR/01-Impuestos Mensuales/${SUBCARPETA:+$SUBCARPETA/}"
if [ $# -gt 0 ]; then PERS="$*"; else PERS=$(ls "$BASE" | grep -E '^[0-9]{6}$' | sort); fi
TMP=$(mktemp -d)
for PER in $PERS; do
  PA="$BASE$PER/$PER- papeles de trabajo- $FARM.xlsx"
  [ -f "$PA" ] || { echo "$FARM $PER: no hay papel"; continue; }
  (cd "$AQUI" && $PY cruce_3_fuentes.py --papel="$PA" ${EXCEPCION:+"--excepcion=$EXCEPCION"} > "$TMP/x.json" 2> "$TMP/err.txt") || { echo "$FARM $PER FALLO"; tail -3 "$TMP/err.txt"; continue; }
  $PY - "$PA" "$TMP/x.json" "$FARM $PER" <<'EOF'
import json, sys
import openpyxl
d = json.load(open(sys.argv[2], encoding='utf-8'))
ws = openpyxl.load_workbook(sys.argv[1], read_only=True)['cruce ventas']
on, out, est, tot = False, [], {}, None
for r in ws.iter_rows(max_col=20, values_only=True):
    if r[0] == 'Zetti - Portal IVA':
        tot = r[6]
    if r[0] == 'Fecha':
        on = True
        continue
    if on and r[0] == 'TOTAL':
        break
    if on and r[0]:
        est[r[3]] = est.get(r[3], 0) + 1
        obs = str(r[4])
        if 'No es una diferencia' not in obs and 'redondeo' not in obs and 'Corrimiento' not in str(r[3]):
            out.append('    ' + str(r[0])[:10] + ' ' + str(r[1])[:40] + ' | ' + obs[:300])
print(sys.argv[3], 'Zetti-Portal', tot, 'otros periodos', d.get('otros_periodos'), est)
for x in out[:40]:
    print(x)
if len(out) > 40:
    print(f'    ... y {len(out) - 40} renglones mas en la solapa "cruce ventas"')
EOF
done
rm -rf "$TMP"
