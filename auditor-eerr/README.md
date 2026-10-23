# auditor-eerr

Auditor de los Estados de Resultados (EERR 2026) de las sociedades. Corre 21 controles sobre cada
`EERR ... .xlsx` de `0 - ESTUDIO - 00000000000/EERR 2026/` y documenta los hallazgos en la hoja **Auditorias**
del mismo archivo.

```
python scripts/auditor_eerr.py --nombre="<sociedad>" [--hasta=AAAAMM] [--simular]
python scripts/auditor_eerr.py --todas [--hasta=AAAAMM] [--simular] [--procesos=4]
```

- **Controles:** período, meses completos, fórmulas, resultado / ventas, margen bruto, gastos / ventas, ventas vs
  Portal IVA y Zetti, existencias, NCR vs NC comerciales, compras clasificadas = Portal IVA compras, sueldos vs F.931,
  IIBB vs DJ ARBA, IVA prorrateo, tasas municipales, gastos fijos y financieros; y del papel de trabajo (desde 1.1.0):
  débito/crédito fiscal vs DDJJ, signo de las NC, cómputo de NRF, prorrateo, venta vs compra exenta/gravada y NC de
  compras vs libro. Detalle en la skill.
- **Hoja Auditorias:** la columna Estado va en B y queda fija al scrollear (desde 1.1.0; las hojas viejas se
  reordenan solas).
- **Estados:** Pendiente → Resuelto (solo, cuando deja de detectarse). Aprobado (equipo) no se vuelve a informar
  mientras el valor no cambie. Historial de ejecuciones a la derecha de la hoja.
- **Hallazgos a mano (desde 1.2.0):** se pueden agregar hallazgos a mano en la hoja Auditorias (una fila con control,
  mes, concepto o descripción): se conservan en cada auditoría y no se resuelven ni se reabren solos; su estado lo
  cambia solo el equipo. Si se cargan sin estado entran como Pendiente.
- **Desde 1.2.0:** hoja **Utilidades** (utilidad bruta y operativa por mes, con fórmulas sobre el EERR) y alerta de
  pérdida operativa; gastos fijos detectados en los últimos 6 meses (cuentas y proveedores) con alerta en el mes que
  faltan; Zetti vs Portal IVA sin tolerancia; NCR con neto gravado o IVA en el libro de compras = rojo.
- **Requisitos:** Python 3 con `openpyxl` y `pypdf` (`requirements.txt`) y Microsoft Excel (Windows: PowerShell +
  COM; Mac: AppleScript) para recalcular cada archivo antes de devolverlo a la unidad.
- **Configuración:** `~/.fisco-ar/.env` (`BVA_UNIDAD_PATH`, `BVA_SALIDAS_PATH`) y, opcional, umbrales y excepciones por
  sociedad (`"excepciones": {"<sociedad o CUIT>": [números de control]}`) en `~/.fisco-ar/auditor_eerr_config.json`.
- **Seguridad:** respaldo de cada EERR antes de pisarlo (`BVA-salidas/auditor-eerr/_respaldos/`); no pisa un archivo
  que cambió en la unidad mientras se auditaba. No entra a ningún sitio ni usa credenciales.
