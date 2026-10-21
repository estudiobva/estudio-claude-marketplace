# auditor-eerr

Auditor de los Estados de Resultados (EERR 2026) de las sociedades. Corre 15 controles sobre cada
`EERR ... .xlsx` de `0 - ESTUDIO - 00000000000/EERR 2026/` y documenta los hallazgos en la hoja **Auditorias**
del mismo archivo.

```
python scripts/auditor_eerr.py --nombre="<sociedad>" [--hasta=AAAAMM] [--simular]
python scripts/auditor_eerr.py --todas [--hasta=AAAAMM] [--simular] [--procesos=4]
```

- **Controles:** período, meses completos, fórmulas, resultado / ventas, margen bruto, gastos / ventas, ventas vs
  Portal IVA y Zetti, existencias, NCR vs NC comerciales, compras clasificadas = Portal IVA compras, sueldos vs F.931,
  IIBB vs DJ ARBA, IVA prorrateo, tasas municipales, gastos fijos y financieros. Detalle en la skill.
- **Hoja Auditorias:** la columna Estado va en B y queda fija al scrollear (desde 1.1.0; las hojas viejas se
  reordenan solas).
- **Estados:** Pendiente → Resuelto (solo, cuando deja de detectarse). Aprobado (equipo) no se vuelve a informar
  mientras el valor no cambie. Historial de ejecuciones a la derecha de la hoja.
- **Requisitos:** Python 3 con `openpyxl` y `pypdf` (`requirements.txt`) y Microsoft Excel (Windows: PowerShell +
  COM; Mac: AppleScript) para recalcular cada archivo antes de devolverlo a la unidad.
- **Configuración:** `~/.fisco-ar/.env` (`BVA_UNIDAD_PATH`, `BVA_SALIDAS_PATH`) y, opcional, umbrales en
  `~/.fisco-ar/auditor_eerr_config.json`.
- **Seguridad:** respaldo de cada EERR antes de pisarlo (`BVA-salidas/auditor-eerr/_respaldos/`); no pisa un archivo
  que cambió en la unidad mientras se auditaba. No entra a ningún sitio ni usa credenciales.
