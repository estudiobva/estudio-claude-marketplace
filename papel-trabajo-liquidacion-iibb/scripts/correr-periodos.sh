#!/bin/bash
# correr-periodos.sh: etapa 1 del papel de ventas (TSH / IIBB) de una sociedad, para uno o varios periodos.
#
#   bash scripts/correr-periodos.sh <CUIT> AAAAMM [AAAAMM ...]
#   SUBCARPETA=controles bash scripts/correr-periodos.sh 30711132925 202509 202510
#
# Por cada periodo, en 01-Impuestos Mensuales/<SUBCARPETA>/AAAAMM de la sociedad:
#   1. Ventas Zetti (5.6.5) -> "AAAAMM- papeles de trabajo- <Farmacia>.xlsx" (solapas Ventas Zetti y depurado)
#   2. Libro IVA Ventas de la DDJJ PRESENTADA (Portal IVA, --presentada) -> "AAAAMM - PORTAL IVA - VENTAS.xlsx"
#      Si ARCA no tiene DDJJ presentada del periodo, baja el BORRADOR del Libro IVA Ventas (importa desde ARCA).
#   3. Mis Comprobantes emitidos -> "AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx"
#   4. armar_papel.py: pega Portal y MC en el papel, TOTAL de MC, depura PV y cuadro resumen
#      (abre y guarda el papel una sola vez).
# El cruce se corre DESPUES, con revisar-cruce.sh, para que cada mes vea a sus vecinos.
#
# SUBCARPETA: por defecto "controles". SUBCARPETA="" deja el papel en 01-Impuestos Mensuales/AAAAMM.
# La carpeta del periodo se crea si no existe.
# Requiere los plugins zetti-subdiario-iva-ventas, portal-iva-descarga y cruce-ventas-3-fuentes del
# marketplace estudiobva-skills, con sus dependencias (setup-bva).
CUIT=$(echo "$1" | tr -cd '0-9'); shift
[ ${#CUIT} = 11 ] && [ $# -gt 0 ] || { echo "Uso: bash correr-periodos.sh <CUIT> AAAAMM [AAAAMM ...]"; exit 2; }
SUBCARPETA=${SUBCARPETA-controles}
AQUI="$(cd "$(dirname "$0")" && pwd)"
PY=${BVA_PYTHON:-python}
export PYTHONIOENCODING=utf-8

plugin() {  # ultima version instalada; si no esta instalado, la copia del marketplace
  local d; d=$(ls -d "$HOME/.claude/plugins/cache/estudiobva-skills/$1"/*/ 2>/dev/null | sort -V | tail -1)
  [ -n "$d" ] || d="$HOME/.claude/plugins/marketplaces/estudiobva-skills/$1/"
  [ -d "$d" ] || { echo "Falta el plugin $1 (instalarlo con setup-bva)" >&2; exit 3; }
  echo "${d%/}"
}
P_PORTAL=$(plugin portal-iva-descarga) || exit 3
P_MC=$(plugin cruce-ventas-3-fuentes) || exit 3
plugin zetti-subdiario-iva-ventas >/dev/null || exit 3

SOCDIR=$(node "$AQUI/carpeta-sociedad.js" "$CUIT") || exit 4
SOCDIR=$(cygpath -u "$SOCDIR" 2>/dev/null || echo "$SOCDIR")
SOC=$(basename "$SOCDIR")
FARM=$(echo "$SOC" | sed -E "s/^ *[0-9]+ *- *//; s/ *- *$CUIT *$//")
SALIDAS=${BVA_SALIDAS_PATH:-$HOME/Documents/BVA-salidas}
echo "Sociedad: $SOC  |  Farmacia: $FARM  |  Subcarpeta: ${SUBCARPETA:-(ninguna)}"

for PER in "$@"; do
  MMAAAA="${PER:4:2}/${PER:0:4}"
  D="$SOCDIR/01-Impuestos Mensuales/${SUBCARPETA:+$SUBCARPETA/}$PER"
  PA="$D/$PER- papeles de trabajo- $FARM.xlsx"
  echo "===== $PER ====="; date +%H:%M
  mkdir -p "$D"

  ok=0
  for i in 1 2 3; do
    node "$AQUI/bajar-ventas-zetti.js" --cuit=$CUIT --periodo=$PER ${SUBCARPETA:+--subcarpeta=$SUBCARPETA} --si-existe=reemplazar 2>&1 | tail -4
    [ -f "$PA" ] && sleep 20 && $PY -c "import zipfile,sys; sys.exit(zipfile.ZipFile(sys.argv[1]).testzip() is not None)" "$PA" 2>/dev/null && ok=1 && break
    echo "papel ausente o danado, reintento"
  done
  [ $ok = 1 ] || { echo "FALLO ZETTI $PER"; continue; }

  SP="$SALIDAS/portal-iva/$SOC - $PER - PORTAL IVA - VENTAS.xlsx"
  SM="$SALIDAS/mis-comprobantes/$PER - MIS COMPROBANTES - EMITIDOS.xlsx"
  rm -f "$SP" "$SM"
  OUTP=$(cd "$P_PORTAL" && node scripts/portal-iva-descarga.js --cuit=$CUIT --periodo=$MMAAAA --presentada --libros=ventas --no-archivar 2>&1)
  echo "$OUTP" | tail -6
  # El proceso nacio para DDJJ sin presentar: si no hay DDJJ presentada, se usa el borrador (importa desde ARCA).
  if [ ! -f "$SP" ] && echo "$OUTP" | grep -q sin_presentacion; then
    echo "Sin DDJJ IVA presentada para $PER: bajo el borrador del Libro IVA Ventas"
    (cd "$P_PORTAL" && node scripts/portal-iva-descarga.js --cuit=$CUIT --periodo=$MMAAAA --libros=ventas --no-archivar 2>&1 | tail -6)
  fi
  [ -f "$SP" ] || { echo "FALLO PORTAL $PER (si el aviso dice que el libro no coincide con la DDJJ presentada, consultar antes de seguir)"; continue; }
  NP="$P_MC/node_modules"; [ -d "$NP" ] || NP="$P_PORTAL/node_modules"
  (export NODE_PATH="$(cygpath -w "$NP" 2>/dev/null || echo "$NP")"; cd "$P_MC" && node scripts/mis-comprobantes-emitidos.js --cuit=$CUIT --periodo=$MMAAAA --no-archivar --si-existe=reemplazar 2>&1 | tail -4)
  [ -f "$SM" ] || { echo "FALLO MC $PER"; continue; }
  cp "$SP" "$D/$PER - PORTAL IVA - VENTAS.xlsx" && cp "$SM" "$D/$PER - MIS COMPROBANTES - EMITIDOS.xlsx"

  (cd "$AQUI" && $PY armar_papel.py "$PA" "$D/$PER - PORTAL IVA - VENTAS.xlsx" "$D/$PER - MIS COMPROBANTES - EMITIDOS.xlsx" 2>&1 | cut -c1-300; exit "${PIPESTATUS[0]}") \
    && $PY -c "import zipfile,sys; sys.exit(zipfile.ZipFile(sys.argv[1]).testzip() is not None)" "$PA" 2>/dev/null && echo "OK PASOS 1-4,6 $PER" || echo "FALLO PASOS $PER"
done
date +%H:%M
