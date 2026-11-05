---
name: armado-eerr
description: "Arma y actualiza los Estados de Resultados (EERR) mensuales de las sociedades en '0 - ESTUDIO - 00000000000/EERR 2026' desde las fuentes de la unidad: ventas y compras del Portal IVA (compras imputadas por proveedor con el Maestro de Proveedores), IIBB de la DJ de anticipo de ARBA, sueldos del F.931, IVA prorrateo, Seguridad e Higiene y Publicidad del papel de trabajo. Completa solo lo que está vacío (no pisa lo que cargó el equipo), registra cada carga con su fuente en la hoja 'Armado' y crea el EERR del ejercicio nuevo cuando cierra el anterior. Usar SIEMPRE que pidan \"armar el EERR de [sociedad]\", \"actualizar los EERR con [mes]\", \"cargar septiembre en los EERR\", \"pasar el mes a los EERR\", \"hacer el EERR del ejercicio nuevo\", \"EERR de [sociedad] que cerró en [mes]\" o \"completar los estados de resultados\". No confundir con auditor-eerr (controla un EERR ya armado) ni con papel-trabajo-automatico (arma el papel de liquidaciones)."
---

# Armado de EERR

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/armar_eerr.py --nombre="<sociedad>" --simular
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/armar_eerr.py --nombre="<sociedad>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/armar_eerr.py --todas
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/armar_eerr.py --nombre="<sociedad>" --crear
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/armar_eerr.py --todas --crear
```

En Mac usar `python3` (o el `BVA_PYTHON` del `.env`).

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | Sociedad, como en el nombre de su carpeta en la unidad (`N - <sociedad> - CUIT`). |
| `--todas` | Todos los EERR de la carpeta (con `--crear`: solo las sociedades cuyo ejercicio ya cerró). Tarda: correrlo en segundo plano. |
| `--hasta=AAAAMM` | Último mes a cargar. Default: el último con la DDJJ de IVA vencida (como el auditor). |
| `--desde=AAAAMM` | Actualizar: primer mes que se revisa (default: los 3 últimos meses hasta `--hasta`). Crear: primer mes del ejercicio, solo si la sociedad no tiene EERR anterior. |
| `--pisar=AAAAMM,...` | Rehace esos meses aunque tengan datos (el valor anterior queda en un comentario de la celda). Usarlo cuando llega el libro de la DDJJ presentada o se corrige el Maestro. **Solo si el usuario lo pide.** |
| `--crear` | Arma el EERR del ejercicio siguiente al último que tiene la sociedad, desde la plantilla limpia (`assets/Plantilla EERR.xlsx`), y lo completa hasta `--hasta`. |
| `--simular` | No toca la unidad: deja el resultado en `BVA-salidas/armado-eerr/simulacion/`. Hacerlo primero si el usuario no pidió guardar. |

## Antes de correr: que estén las fuentes del mes

1. **Libros del Portal IVA** (ventas y compras) en la carpeta del mes: `portal-iva-descarga`, mejor con `--presentada`.
   Si para compras solo está el libro sin NCR de la conciliación o un borrador, se usa igual pero se avisa: pueden faltar
   Edesur, AySA o Metrogas que el equipo carga a mano antes de presentar.
2. **DJ de anticipo de ARBA** (`AAAAMM-IIBB-DJ-*.pdf` o `AAAAMM-IIBB-DDJJ-*.pdf`): `constancias-anticipo-arba`.
   AGIP y Convenio Multilateral no tienen fuente automática: IIBB queda vacío y se carga a mano.
3. **F.931** (`02-Sueldos y F.931/AAAA/AAAAMM-F931-DDJJ-*.pdf`): `f931-descarga-arca`.
4. **Papel de trabajo** del equipo (`AAAA00-liquidaciones-...xlsx`, hoja `MM-AAAA`) para la alícuota de Seguridad e
   Higiene y la PyP. Si el mes no está cargado, usa la alícuota del último mes que la tenga; si ningún papel la tiene,
   la tasa efectiva del EERR (queda **ESTIMADO** en el comentario de la celda).

## Qué hace

Por cada mes y concepto, **solo si la celda del EERR está vacía**:

| Fila del EERR | Fuente |
|---|---|
| Ventas exentas y no gravadas / gravadas | libro de ventas → hoja "Ventas Portal IVA" (por tipo de comprobante) + SUMIFS |
| CMV y gastos de compras | libro de compras imputado con el Maestro → hoja "Compras imputadas" (proveedor × mes) + SUMIFS |
| IIBB | impuesto determinado de la DJ de ARBA |
| IVA Prorrateo | IVA de compras que no son CMV × % de venta exenta → hoja "IVA prorrateo" |
| Sueldos y Cargas Sociales | F.931 → hoja "Sueldos" (sección VIII; rectificativa con VIII en cero: I, II y VI) |
| Seguridad e Higiene | ventas gravadas × alícuota del papel + conceptos fijos |
| Publicidad | PyP del papel + compras imputadas a Publicidad |

Criterios de las compras (acordados el 01/10/2026): **manda el Maestro de Proveedores**. Importe = no gravado + exento
+ netos; sin IVA discriminado (B, C, tiques), total − percepciones. Bloqueados y bienes de uso (`COMPRA DE EQUIPOS`)
quedan fuera del resultado; estado Revisar, sin Maestro o cuenta sin fila en el EERR → GASTOS VARIOS (se avisa).

Hoja **"Armado"**: una fila por carga (fecha, modo, mes, concepto, acción, valor, archivo de origen). En amarillo lo que
quedó sin fuente y los proveedores a revisar. Lo que ya tenía dato no se registra.

**Necesita Microsoft Excel**: recalcula cada archivo antes de devolverlo a la unidad. Antes de pisar un EERR deja un
respaldo en `EERR 2026/_respaldos/AAAAMMDD-HHMM/`; si el archivo cambió en la unidad mientras se armaba, no lo pisa y
deja la copia en `BVA-salidas/armado-eerr/no-guardados/`. Pedir que el EERR esté cerrado en Excel mientras corre.

## Después

Correr `auditor-eerr` sobre la sociedad (o pedirle al usuario si lo corre): controla el EERR armado contra las mismas
fuentes y deja los hallazgos en la hoja "Auditorias".

## Respuesta al usuario

Por sociedad: archivo, si quedó guardado / creado / simulado, cuántos conceptos se cargaron, y **lo que quedó sin
fuente por mes** (con qué herramienta se baja lo que falta). Mencionar los avisos de libro de compras no presentado y los
proveedores a revisar en el Maestro. No pegar la hoja Armado entera.

**Plugin instalado:** la primera vez y después de cada actualización, `python -m pip install -r requirements.txt`
(o "revisá el setup del estudio"). Usa el mismo `~/.fisco-ar/.env` que los demás plugins (`BVA_UNIDAD_PATH`,
`BVA_SALIDAS_PATH`).
