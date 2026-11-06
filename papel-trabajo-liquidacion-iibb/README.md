# papel-trabajo-liquidacion-iibb

Etapa 1 del papel de trabajo mensual de ventas de las sociedades del grupo, que es la base de TSH e IIBB por jurisdicción.

Por cada período deja en `01-Impuestos Mensuales/controles/AAAAMM` (o en la carpeta del mes) estos archivos:

- `AAAAMM- papeles de trabajo- <Farmacia>.xlsx`, con estas solapas:
  - Ventas Zetti
  - Ventas Zetti (depurado)
  - Ventas portal iva (con cuadro resumen)
  - Ventas MCR (con TOTAL)
  - cruce ventas
- `AAAAMM - PORTAL IVA - VENTAS.xlsx`, el Libro IVA Ventas de la DDJJ **presentada**; si todavía no se presentó, el del borrador (importado desde ARCA).
- `AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx`

## Requisitos

Plugins del marketplace `estudiobva-skills`, con sus dependencias instaladas (`setup-bva`):
- `zetti-subdiario-iva-ventas`: login y reporte 5.6.5 de Zetti.
- `portal-iva-descarga`: Libro IVA de la DDJJ presentada y rutas de la unidad compartida.
- `cruce-ventas-3-fuentes`: descarga de Mis Comprobantes emitidos.

Además: Python con `openpyxl` (`requirements.txt`), Node 20 o más, Git Bash, la unidad compartida montada (o `BVA_UNIDAD_PATH` en el `.env`) y la planilla de claves configurada.

## Uso

```bash
# 1. Armar los papeles (en segundo plano; con sociedades grandes tarda)
bash scripts/correr-periodos.sh 30711132925 202509 202510 202511 > log.txt 2>&1

# 2. Cruce de todos los meses (cuando están todos armados)
bash scripts/revisar-cruce.sh 30711132925
```

- Para guardar en la carpeta del mes en lugar de `controles`, usar `SUBCARPETA=""`.
- Python: `BVA_PYTHON`, por defecto `python`.
- Salidas temporales: `BVA_SALIDAS_PATH`, por defecto `~/Documents/BVA-salidas`.

La habilidad `papel-trabajo-liquidacion-iibb` (en `skills/`) le explica a Claude el proceso completo y cómo distinguir corrimientos y cortes de mes de las diferencias reales.

## Notas

- `armar_papel.py` abre y guarda el papel una sola vez. `totales_mcr.py` toma `max_row` una sola vez. Con Salud y Belleza (~77.000 comprobantes por mes), cada mes pasó de unas 3 horas a minutos.
- `zetti_papel_trabajo.py` tolera hasta 50 centavos cruzados entre exento y gravado en el total del reporte de Zetti, siempre que el Total dé exacto (Sanar 12 202602).
- `cruce_3_fuentes.py` detecta un cierre Z cargado con un PV inexistente cuyo número sigue la secuencia de otro PV (Sanar 10 202603: Z 1830 con PV 41 en vez de 14). Lo cruza con el PV correcto y deja un aviso.
