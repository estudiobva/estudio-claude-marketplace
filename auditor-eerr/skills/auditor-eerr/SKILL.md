---
name: auditor-eerr
description: "Audita los Estados de Resultados (EERR 2026) de una o todas las sociedades con 21 controles (período, meses completos, fórmulas, resultado, margen bruto, gastos/ventas, ventas vs Portal IVA y Zetti, existencias, NCR vs NC comerciales, compras clasificadas = Portal IVA compras, sueldos vs F.931, IIBB vs DJ ARBA, IVA prorrateo, tasas, gastos fijos y financieros, y del papel de trabajo: débito/crédito fiscal vs DDJJ, signo de las NC, cómputo de NRF, prorrateo, venta vs compra exenta/gravada y NC de compras vs libro) y documenta cada hallazgo en la hoja 'Auditorias' del mismo EERR, con fecha, descripción, links y estado Pendiente/Resuelto/Aprobado. Usar SIEMPRE que pidan \"auditá el EERR de [sociedad]\", \"auditar los EERR\", \"correr el auditor\", \"qué errores tiene el EERR\", \"control del estado de resultados\", \"KPIs de auditoría\" o \"revisar la solapa Auditorias\"."
---

# Auditor de EERR

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/auditor_eerr.py --nombre="<sociedad>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/auditor_eerr.py --nombre="<sociedad>" --simular
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/auditor_eerr.py --todas
```

En Mac usar `python3` (o el `BVA_PYTHON` del `.env`).

| Argumento | Para qué |
|---|---|
| `--nombre` | Sociedad, como figura en el nombre del EERR (`EERR 01-2026 a 12-2026 - <sociedad>.xlsx`). |
| `--todas` | Todos los EERR de la carpeta. Además marca importes idénticos entre sociedades (planillas copiadas). Tarda varios minutos: correrlo en segundo plano. **Solo si el usuario lo pide expresamente.** |
| `--hasta=AAAAMM` | Último mes a auditar. Default: el último con la DDJJ de IVA vencida (desde el día 26, el mes anterior; si no, el anterior a ese). |
| `--simular` | No toca la unidad: deja la copia auditada en `BVA-salidas/auditor-eerr/simulacion/`. Usarlo si el usuario solo quiere ver qué encuentra. |

Lee los EERR de `0 - ESTUDIO - 00000000000/EERR 2026/` y, de cada sociedad, su carpeta en la unidad: libros
`AAAAMM - PORTAL IVA - VENTAS/COMPRAS.xlsx`, cruce Zetti (hoja "cruze" del papel del mes), papel de liquidaciones
(`202600-liquidaciones-IVA, IIBB y Tasa Municipal-X.xlsx`, hojas `MM-AAAA`), DDJJ de IVA (`AAAAMM-IVA-DDJJ-*.pdf`), F.931 (`02-Sueldos y F.931`), DJ de anticipo de ARBA
(`AAAAMM-IIBB-DJ-*.pdf`), conciliación de NCR (`AAAAMM-COMPRAS-SIN-NCR-*.xlsx`) y el Maestro de Proveedores.

**Necesita Microsoft Excel** en la computadora: recalcula cada archivo antes de devolverlo a la unidad (openpyxl no
calcula fórmulas). Si Excel no puede, no audita ese archivo. Antes de pisar cada EERR deja un respaldo en
`BVA-salidas/auditor-eerr/_respaldos/AAAAMMDD-HHMM/`; si el archivo cambió en la unidad mientras se auditaba, no lo
pisa y deja la copia en `no-guardados/`. Pedir que el EERR esté cerrado en Excel mientras corre.

## Hoja "Auditorias"

Una fila por hallazgo: N°, **Estado** (columna B, fija al scrollear junto con N°), nivel (**Rojo** = no cierra contra
una fuente o falta algo indispensable; **Amarillo** = desvío a revisar), fecha de detección, control, mes, concepto,
descripción, valor EERR, valor de control, diferencia, link a la celda del EERR, link al documento fuente, fecha de
estado, observaciones y última auditoría. A la derecha, el historial de cada ejecución. Las hojas escritas con la 1.0.0
(Estado en la columna M) se reordenan solas en la próxima auditoría, sin perder estados, observaciones ni links.

- Nuevo → **Pendiente**. Pendiente que deja de detectarse → **Resuelto** (solo).
- **Aprobado** (lo pone el equipo) → no se vuelve a informar mientras el valor no cambie; si cambia, entra como nuevo.
- Resuelto que vuelve a detectarse → se reabre como Pendiente.

## Los 21 controles

1. Período: 12 meses seguidos, nombre del archivo, CUIT y carpeta; meses no cerrados con datos.
2. Meses completos: ventas, CMV y sueldos en cada mes; importes repetidos entre meses (y entre sociedades con `--todas`).
3. Fórmulas: errores, filas de totales y resultado con fórmula, rango del TOTAL de gastos, TOTAL = suma de filas.
4. Resultado / ventas: pérdidas 2+ meses seguidos; utilidad mayor al 25%.
5. Margen bruto: banda 20–40% (provisoria), ±5 pts vs mes anterior y vs mediana.
6. Cada gasto / ventas: salto de más de 20% vs mediana y vs mes anterior, con materialidad.
7. Ventas vs Portal IVA (tolerancia 1%) y cruce Zetti: lo que queda sin explicar en la fila TOTAL DIFERENCIAS de la hoja "cruze"
   del papel del mes (`AAAAMM-papeldetrabajo-X.xlsx`, la deja cruze-zetti-portaliva-ventas); si no, el cruce de 3 fuentes o el
   archivo de cruce suelto. Papeles con "no va" / "no usar" en el nombre no se leen.
8. Existencias inicial y final.
9. NCR vs NC comerciales (desvío de 30% vs su mediana); NCR no encontradas; meses sin conciliar.
10. Compras clasificadas por cuenta = Portal IVA compras; cada rubro llega a su fila; proveedores sin clasificar o reclasificados en el Maestro; peso de GASTOS VARIOS.
11. Sueldos vs F.931 (remuneraciones + 351 + 352 + 312 + 028); costo laboral / ventas.
12. IIBB vs impuesto determinado de la DJ de ARBA; alícuota efectiva 1,5–5%.
13. IVA prorrateo vs la hoja de control "IVA prorrateo".
14. Seguridad e Higiene, matrícula, publicidad y marquesinas.
15. Gastos fijos faltantes, posibles facturas duplicadas y financieros sin cargar.

Controles del papel de trabajo (lista de Meli, Ezequiel y Jaz del 14/10/2026). Leen la hoja `MM-AAAA` del papel de
liquidaciones del equipo por etiqueta (no el "- AUTOMATICO"); los `.xls` no se leen (aviso en el historial).

16. Débito y crédito fiscal del papel = DDJJ de IVA presentada (PDF del mes o `BVA-salidas/ddjj-iva`); crédito / débito
    que se aparta 25% de su mediana o crédito mayor que el débito cuando no es lo habitual.
17. Signo de las NC: ninguna nota de crédito (ventas o compras) en positivo; también filas con código de NC y nombre de factura.
18. Cómputo de NRF: las NCR conciliadas (`COMPRAS-SIN-NCR`) están dentro de las NC de compras del papel y coinciden con
    "NCR eliminadas"; meses con conciliación sin el bloque "NCR eliminadas / TOTAL NC".
19. Prorrateo del papel: % de venta exenta = Ventas Netas del papel y = Portal IVA ventas (1 punto); IVA no computable =
    % exento × IVA de servicios; IVA de servicios sin alícuota en la clasificación (queda fuera del prorrateo); fila IVA
    Prorrateo del EERR vs papel.
20. Venta vs compra: % exento de las ventas vs de las compras de mercadería (15 puntos); compras / ventas exentas y gravadas
    que se apartan 30% de su mediana.
21. NC de compras del papel = libro de compras del Portal IVA (con o sin las NCR ya eliminadas del libro), con el detalle
    NCR / NC comerciales.

Excepciones por sociedad (una sola auditoría, criterio de Javi): en `auditor_eerr_config.json`,
`"excepciones": {"<sociedad o CUIT>": [19]}` saltea esos controles para esa sociedad (queda anotado en el historial).

Umbrales en `CFG` del script; se pisan con `~/.fisco-ar/auditor_eerr_config.json` (sobrevive a las actualizaciones).

## Respuesta al usuario

Por sociedad: cantidad de nuevos, siguen pendientes, resueltos solos y aprobados omitidos; después los **rojos**
(lo que no cierra contra una fuente) y un resumen de los amarillos por control. No pegar la hoja entera. Recordar que
el estado se cambia en la columna Estado (B) de la hoja "Auditorias".

**Plugin instalado:** la primera vez y después de cada actualización, `python -m pip install -r requirements.txt`
(o "revisá el setup del estudio"). Usa el mismo `~/.fisco-ar/.env` que los demás plugins (`BVA_UNIDAD_PATH`; en
Windows por defecto `I:\Unidades compartidas\BVA - Sociedades Farmaceuticas`).
