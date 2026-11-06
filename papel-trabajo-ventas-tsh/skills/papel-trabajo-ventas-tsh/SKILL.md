---
name: papel-trabajo-ventas-tsh
description: Etapa 1 del papel de trabajo mensual de ventas de una farmacia/sociedad del grupo (base para TSH e IIBB por jurisdicción). Por cada período baja las ventas de Zetti (5.6.5), el Libro IVA Ventas de la DDJJ presentada (Portal IVA) y Mis Comprobantes emitidos, arma el papel "AAAAMM- papeles de trabajo- <Farmacia>.xlsx" con las solapas Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva (con cuadro resumen), Ventas MCR (con TOTAL) y cruce ventas, y explica las diferencias reales entre Zetti y el Portal. Usar SIEMPRE que pidan "hacé [sociedad], períodos ...", "papel de trabajo de ventas", "etapa 1", "ventas para TSH / IIBB", "guardalo en controles", "cruce de ventas Zetti Portal Mis Comprobantes" para uno o varios meses. La etapa 2 (DDJJ de IVA, compras) es otra habilidad: papel-trabajo-ddjj-iva.
---

# Papel de trabajo de ventas (etapa 1: TSH / IIBB)

Todo corre con un script de Playwright con login automático (planilla de claves del estudio). No pedir ni escribir contraseñas.

## Datos que hacen falta

- **Sociedad**: nombre o CUIT. El script busca la carpeta en la unidad compartida por CUIT. Si te dan solo el nombre, buscá la carpeta (`<n> - <Nombre> - <CUIT>`) y confirmá el CUIT si hay más de una parecida.
- **Períodos**: AAAAMM, uno o varios.
- **Dónde guardar**: por defecto `01-Impuestos Mensuales/controles/AAAAMM`. Si piden la carpeta del mes normal, usá `--subcarpeta=`.

## Correr el script

Un solo script de Playwright hace todo: por cada período baja Zetti (5.6.5), el Libro IVA Ventas de la DDJJ presentada y Mis Comprobantes emitidos, arma el papel y, con todos los meses listos, corre el cruce y resume cada mes.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-ventas.js --cuit=<CUIT> --periodos=202509,202510,202511 > "$HOME/Documents/BVA-salidas/papel-ventas-tsh/ultimo.json" 2> "$HOME/Documents/BVA-salidas/papel-ventas-tsh/ultimo.log"
```

- Rango: `--desde=202601 --hasta=202609` en lugar de `--periodos`.
- Carpeta: por defecto `01-Impuestos Mensuales/controles/AAAAMM`. Para la carpeta del mes: `--subcarpeta=` (vacío).
- Correrlo **en segundo plano** (unos 3 o 4 minutos por mes en una farmacia común; 20 o 30 en las de 15.000 a 80.000 comprobantes por mes) y pedir que la computadora no se suspenda. El progreso queda en el `.log`.
- Navegador siempre oculto; `--ver` solo si el usuario pide ver la ventana.
- Es reanudable: un mes con el papel completo se saltea (`--si-existe=reemplazar` para rehacerlo). Antes de reemplazar un papel lo copia a `BVA-salidas/papel-ventas-tsh/respaldos`.
- **No pisa la etapa 2:** si el papel ya tiene solapas de compras (Compras Zetti, cruce compras…) no lo rehace. Solo con `--forzar` y el OK del usuario.
- Solo el cruce (por ejemplo después de corregir algo a mano): `--solo-cruce` (sin `--periodos` toma todos los meses de la subcarpeta). Para no correr el cruce: `--sin-cruce`.
- `--excepcion="2:NC B"` se pasa a la depuración de PV y al cruce.

Leer el JSON (también queda en `BVA-salidas/papel-ventas-tsh/<CUIT>-<fecha>.json`):
- `meses[]`: por período `ok`, y si falló, `paso` (`zetti`, `portal`, `mc`, `armado`, `papel`) y `error`. Un fallo de Zetti o de ARCA suele ser temporal: volver a correr solo ese período.
- **Fallo `portal` porque el libro no coincide con la DDJJ presentada**: no forzarlo. Avisar al usuario con los dos importes. Ojo: el control compara el neto gravado de facturas y tiques **sin restar las NC**, y no es el neto del cuadro resumen.
- `armado`: control del cuadro resumen (`cuadro_ok`, `cuadro_diferencia`), códigos no reconocidos y lo eliminado en la depuración.
- `cruce[]`: por período `zetti_menos_portal`, `estados`, `otros_periodos` y `para_revisar` (las observaciones que no son corrimientos ni redondeos).
- Código de salida 3: ARCA pidió captcha o rechazó la clave; el lote se cortó. No reintentar en loop.

Qué hace `armar_papel.py` (paso 4 de cada mes, una sola apertura y guardado del papel):
- Pega `Ventas portal iva` y `Ventas MCR`.
- Agrega el TOTAL de Mis Comprobantes.
- Depura en `Ventas Zetti (depurado)` los puntos de venta que no están en el Portal (PV 1, 2, 9999…, fallidos de Zetti que no van a ARCA).
- Arma el cuadro resumen de ventas debajo del Portal.

El cruce mira el mes anterior y el siguiente para separar los cortes de mes: por eso va al final, con todos los meses armados.

## Investigar y reportar

Reglas:
- **Corrimientos de fecha no son diferencia.** Un cierre Z que Zetti carga otro día que ARCA (hasta 10 días, varios días juntos, tramos que suman 0) se compensa dentro del mes. Se menciona solo como aclaración. Solo es diferencia real lo que cambia el total del mes.
- **Cortes de mes** (comprobantes del último día que el otro sistema pone en el mes vecino) tampoco son diferencia si el vecino los tiene.
- Esquema 2025: ARCA informa un renglón por cierre Z, numerado con el número de Z. El cruce los cruza por número. Desde 2026 ARCA trae los tiques uno por uno.
- Para cada diferencia real, buscar la causa en las solapas y dar comprobante, fecha e importe. Causas típicas:
  - Z cargado dos veces en Zetti.
  - Z que falta en ARCA o en Zetti.
  - Z con el mismo total y distinta composición.
  - Comprobante duplicado.
  - Z con PV mal tipeado: el cruce lo detecta solo si el número sigue la secuencia de otro PV.
  - ARCA con PV o números cruzados.
  - Bloques de facturas que faltan en Zetti.
- No volver a correr sociedades ya terminadas salvo que lo pidan.

Reporte al usuario: una tabla por período con `Zetti − Portal IVA` real y su causa. Después, las aclaraciones (corrimientos, cortes, composición) y lo que se eliminó en la depuración (cantidad e importe por PV).

## Archivos

| Script | Qué hace |
|---|---|
| `papel-ventas.js` | Orquesta todo: los cuatro pasos por período, el cruce y el resumen. Es lo único que hay que correr. |
| `bajar-ventas-zetti.js` + `zetti_papel_trabajo.py` | Ventas Zetti al papel; controla CUIT, fechas y el total del reporte, con tolerancia de 50 centavos para centavos cruzados entre exento y gravado. |
| `armar_papel.py` | copiar_solapa ×2, totales_mcr, depurar_pv y cuadro_resumen en una pasada. Cada uno también se puede correr suelto. |
| `cruce_3_fuentes.py --papel=<xlsx>` + `resumen_cruce.py` | Solapa `cruce ventas` y su resumen. |

Usa los plugins `portal-iva-descarga` (Libro IVA de la DDJJ presentada) y `cruce-ventas-3-fuentes` (Mis Comprobantes), que tienen que estar instalados. Nunca escribe en ARCA ni en Zetti.
