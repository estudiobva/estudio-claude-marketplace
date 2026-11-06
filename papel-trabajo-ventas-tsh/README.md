# papel-trabajo-ventas-tsh

Etapa 1 del papel de trabajo mensual de ventas de las sociedades del grupo, que es la base de TSH e IIBB por jurisdicción.

Por cada período deja en `01-Impuestos Mensuales/controles/AAAAMM` (o en la carpeta del mes) estos archivos:

- `AAAAMM- papeles de trabajo- <Farmacia>.xlsx`, con estas solapas:
  - Ventas Zetti
  - Ventas Zetti (depurado)
  - Ventas portal iva (con cuadro resumen)
  - Ventas MCR (con TOTAL)
  - cruce ventas
- `AAAAMM - PORTAL IVA - VENTAS.xlsx`, el Libro IVA Ventas de la DDJJ **presentada**.
- `AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx`

## Requisitos

Plugins del marketplace `estudiobva-skills`, instalados con sus dependencias (`setup-bva`):
- `portal-iva-descarga`: Libro IVA de la DDJJ presentada.
- `cruce-ventas-3-fuentes`: descarga de Mis Comprobantes emitidos.

El login y el reporte de Zetti vienen en `lib/` (copia de `zetti-subdiario-iva-ventas`), con su propio `package.json` (Playwright). Además: Python con `openpyxl` (`requirements.txt`), Node 20 o más, la unidad compartida montada (o `BVA_UNIDAD_PATH` en el `.env`) y la planilla de claves configurada (`.env.example`). Ya no hace falta Git Bash.

## Uso

```bash
# Arma los papeles de cada mes y al final corre el cruce (en segundo plano; con sociedades grandes tarda)
node scripts/papel-ventas.js --cuit=<CUIT> --periodos=202509,202510,202511 > salida.json 2> progreso.log
node scripts/papel-ventas.js --cuit=<CUIT> --desde=202601 --hasta=202609

# Solo el cruce y el resumen de todos los meses de la subcarpeta
node scripts/papel-ventas.js --cuit=<CUIT> --solo-cruce
```

- Para guardar en la carpeta del mes en lugar de `controles`: `--subcarpeta=` (vacío).
- Reanudable: saltea los meses con el papel completo (`--si-existe=reemplazar` para rehacerlos). Antes de reemplazar respalda el papel en `BVA-salidas/papel-ventas-tsh/respaldos`.
- No rehace un papel que ya tiene solapas de la etapa 2 (compras), salvo con `--forzar`.
- Navegador oculto; `--ver` para verlo.
- Python: `BVA_PYTHON` (por defecto `python3` en Mac, `python` en Windows). Salidas: `BVA_SALIDAS_PATH`, por defecto `~/Documents/BVA-salidas`.
- Nunca escribe en ARCA ni en Zetti.

La habilidad `papel-trabajo-ventas-tsh` (en `skills/`) le explica a Claude cómo correrlo y cómo distinguir corrimientos y cortes de mes de las diferencias reales.

## Notas

- `armar_papel.py` abre y guarda el papel una sola vez. `totales_mcr.py` toma `max_row` una sola vez. Con las sociedades de ~77.000 comprobantes por mes, cada mes pasó de unas 3 horas a minutos.
- `zetti_papel_trabajo.py` tolera hasta 50 centavos cruzados entre exento y gravado en el total del reporte de Zetti, siempre que el Total dé exacto.
- `cruce_3_fuentes.py` detecta un cierre Z cargado con un PV inexistente cuyo número sigue la secuencia de otro PV (ej. Z con PV 41 en vez de 14). Lo cruza con el PV correcto y deja un aviso.
