---
name: papel-trabajo-liquidacion-iibb
description: Etapa 1 del papel de trabajo mensual de ventas de una farmacia/sociedad del grupo (base para TSH e IIBB por jurisdicción). Por cada período baja las ventas de Zetti (5.6.5), el Libro IVA Ventas del Portal IVA (DDJJ presentada o, si no está presentada, el borrador) y Mis Comprobantes emitidos, arma el papel "AAAAMM- papeles de trabajo- <Farmacia>.xlsx" con las solapas Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva (con cuadro resumen), Ventas MCR (con TOTAL) y cruce ventas, y explica las diferencias reales entre Zetti y el Portal. Usar SIEMPRE que pidan "hacé [sociedad], períodos ...", "papel de trabajo de ventas", "etapa 1", "ventas para TSH / IIBB", "guardalo en controles", "cruce de ventas Zetti Portal Mis Comprobantes" para uno o varios meses. La etapa 2 (DDJJ de IVA, compras) es otra habilidad: papel-trabajo-ddjj-iva.
---

# Papel de trabajo de ventas (etapa 1: TSH / IIBB)

Todo corre con login automático (planilla de claves del estudio). No pedir ni escribir contraseñas.

## Datos que hacen falta

- **Sociedad**: nombre o CUIT. La carpeta se busca en la unidad compartida por CUIT (`node scripts/carpeta-sociedad.js <CUIT>`). Si te dan solo el nombre, buscá la carpeta (`<n> - <Nombre> - <CUIT>`) y confirmá el CUIT si hay más de una parecida.
- **Períodos**: AAAAMM, uno o varios.
- **Dónde guardar**: por defecto `01-Impuestos Mensuales/controles/AAAAMM`. Si piden la carpeta del mes normal, usá `SUBCARPETA=""`.

## Paso 1: armar los papeles (todos los períodos)

```bash
bash "${CLAUDE_PLUGIN_ROOT}/scripts/correr-periodos.sh" <CUIT> 202509 202510 ... 202607
```

- Correrlo en segundo plano y con el log a un archivo. Con sociedades grandes tarda: unos 3 o 4 minutos por mes en una farmacia común, y hasta 20 o 30 en las de 15.000 a 80.000 comprobantes por mes (La Recova, Salud y Belleza).
- Pedir que la computadora no se suspenda mientras corre. Si se duerme, un paso puede quedar colgado: mirar el log y, si no avanza, cortar el proceso y volver a correr ese período.
- Cada período termina con `OK PASOS 1-4,6 AAAAMM` o con `FALLO ZETTI|PORTAL|MC|PASOS AAAAMM`. Un fallo de Zetti o de ARCA suele ser temporal: volver a correr solo ese período.
- **Sin DDJJ de IVA presentada** (`sin_presentacion`): el script baja solo el borrador del Libro IVA Ventas, importando desde ARCA, y sigue. No frenar a preguntar: el proceso nació para DDJJ sin presentar. Si después se presenta con otros importes, volver a correr el mes.
- **FALLO PORTAL porque el libro no coincide con la DDJJ presentada**: no forzarlo. Avisar al usuario con los dos importes. Ojo: el control compara el neto gravado de facturas y tiques **sin restar las NC**, y no es el neto del cuadro resumen. Puede ser un renglón que el control leyó mal: mirar el Excel antes de decir que hay diferencia.

Qué hace `armar_papel.py` (paso 4 de cada mes, una sola apertura y guardado del papel):
- Pega `Ventas portal iva` y `Ventas MCR`.
- Agrega el TOTAL de Mis Comprobantes.
- Depura en `Ventas Zetti (depurado)` los puntos de venta que no están en el Portal (PV 1, 2, 9999…, fallidos de Zetti que no van a ARCA).
- Arma el cuadro resumen de ventas debajo del Portal.

## Paso 2: cruce de todos los meses

```bash
bash "${CLAUDE_PLUGIN_ROOT}/scripts/revisar-cruce.sh" <CUIT>
```

Se corre cuando están todos los meses: el cruce mira el mes anterior y el siguiente para separar los cortes de mes. Escribe la solapa `cruce ventas` en cada papel y resume por mes:
- `Zetti-Portal`: la diferencia del total del mes.
- Los estados.
- Las observaciones que no son corrimientos ni redondeos.

## Paso 3: investigar y reportar

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

- **Cada diferencia real sin compensar se deja escrita en el papel**, no solo en el chat. Va en la columna Observacion del renglon en la solapa `cruce ventas` (punto 4, "Detalle de las diferencias"). Se escribe con el mismo detalle que el reporte: importe y columna (exento, gravado, IVA), comprobantes con PV-numero, fecha e importe, que paso en cada fuente y los importes de Zetti contra ARCA. Si el renglon del dia vecino es un corrimiento, reescribir tambien su observacion para que diga con que Z compensa.
- **Resaltar en verde `99FFCC` (fila entera)** los comprobantes que originan esa diferencia, en las solapas de las fuentes: Ventas Zetti (depurado), Ventas portal iva y Ventas MCR. Va para las diferencias sin compensar, las que forman el Zetti - Portal del mes, y para los Z con el mismo total y distinta composicion (cambian exento, gravado e IVA). En esos Z se resalta la fila del Z en el depurado y todos los tiques de ese dia en Portal y MCR. No va para corrimientos ni cortes de mes. La observacion de un Z con distinta composicion se escribe asi: "Zetti: gravado $X e IVA $Y. ARCA: ...", y el estado es "Distinta composicion (mismo total)". Nunca se rotula como redondeo.
- Estos cambios se hacen con Excel (COM desde PowerShell), no con openpyxl. Asi no se pierden las imagenes, notas ni resaltados que haya agregado el usuario. No volver a correr `cruce_3_fuentes.py` sobre un papel que el usuario ya toco, porque reescribe la solapa entera.

Reporte al usuario: una tabla por período con `Zetti − Portal IVA` real y su causa. Después, las aclaraciones (corrimientos, cortes, composición) y lo que se eliminó en la depuración (cantidad e importe por PV).

## Archivos

| Script | Qué hace |
|---|---|
| `correr-periodos.sh` | Orquesta los pasos por período. |
| `bajar-ventas-zetti.js` + `zetti_papel_trabajo.py` | Ventas Zetti al papel; controla CUIT, fechas y el total del reporte, con tolerancia de 50 centavos para centavos cruzados entre exento y gravado. |
| `armar_papel.py` | copiar_solapa ×2, totales_mcr, depurar_pv y cuadro_resumen en una pasada. Cada uno también se puede correr suelto. |
| `cruce_3_fuentes.py --papel=<xlsx>` | Solapa `cruce ventas`. |
| `revisar-cruce.sh` | Cruce de todos los meses y resumen. |
| `carpeta-sociedad.js` | Carpeta de la sociedad por CUIT. |
