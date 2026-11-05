---
name: papel-trabajo-ddjj-iva
description: Etapa 2 del papel de trabajo mensual de una farmacia/sociedad (la etapa 1 es la de ventas para TSH e IIBB). Prepara los papeles de trabajo para presentar la DDJJ de IVA sobre el MISMO papel `AAAAMM- papeles de trabajo- <Sociedad>.xlsx` que ya trae Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva (con cuadro resumen), Ventas MCR y cruce ventas. Paso 1 - baja el Subdiario de IVA Compras de Zetti (5.6.4) y lo pega como solapa "Compras Zetti" después de "cruce ventas". Paso 2 - baja el Libro IVA Compras del Portal IVA y lo pega como "Compras portal iva" al lado. Paso 3 - cruza compras Zetti vs Portal IVA en la solapa "cruce compras". Paso 4 - filtra en "Compras Zetti" las facturas de AySA, EDESUR y METROGAS. Paso 5 - las agrega al final de "Compras portal iva" y vuelve a totalizar. Paso 6 - titulo PROVEEDORES en la columna AG de "Compras portal iva". Paso 7 - clasifica los proveedores con el Maestro de Proveedores y arma el cuadro resumen de compras. Usar SIEMPRE que pidan "preparar el papel de IVA de [sociedad] [periodo]", "DDJJ de IVA", "liquidación de IVA", "bajar las compras de Zetti y pegarlas en el papel", "etapa 2 del papel de trabajo". No se corre antes del día 11 del mes siguiente al período.
---

# Papel de trabajo para la DDJJ de IVA (etapa 2)

Se trabaja sobre el papel de trabajo del mes que arrancó la etapa 1 (ventas para TSH e IIBB, pasos 1 a 6).
Se arma paso a paso: mostrar el resultado de cada paso y esperar su OK antes del siguiente.

## Datos de entrada

- Sociedad y período (AAAAMM). El CUIT sale de la carpeta de la sociedad en la unidad compartida
  (`<unidad>\Unidades compartidas\BVA - Sociedades Farmaceuticas\<N> - <Sociedad> - <CUIT>`; la unidad es I: en
  las máquinas del estudio, o la de `BVA_UNIDAD_PATH` en el `.env`).
- Papel: `<Sociedad>/01-Impuestos Mensuales/AAAAMM/AAAAMM- papeles de trabajo- <Sociedad>.xlsx`
  (mismo formato de nombre para todas las sociedades, ej. `202609- papeles de trabajo- Sanar 1.xlsx`).
  Si el papel no existe o no tiene la solapa "cruce ventas" (en papeles viejos, "cruce"), la etapa 1 no está terminada: avisar y no seguir.
- Fecha: la etapa 2 no se corre antes del día 11 del mes siguiente al período (IVA 202609 → desde el 11/10/2026).

## Scripts

En `scripts/` del plugin (se corren con `cd "${CLAUDE_PLUGIN_ROOT:-.}" && ...`). Necesita Node (Playwright) y
Python con openpyxl; `setup-bva` instala las dependencias. Ninguno escribe en ARCA ni en Zetti.

- `bajar-compras-zetti.js` — 5.6.4 (formName `rpt_subdiario_iva_compra`), sociedad completa, emisión del 1 al
  último día, CSV → Excel con `zetti_compras_excel.py`. Login automático con el usuario de Zetti de la planilla
  de claves (`lib/` copiada de zetti-subdiario-iva-ventas). Solo lee Zetti.
- `cruce_compras.py` (paso 3), `filtrar_servicios.py` (paso 4), `agregar_servicios_portal.py` (paso 5),
  `clasificar_proveedores.py` y `cuadro_compras.py` (paso 7).
- El paso 2 usa el script del plugin **portal-iva-descarga** (tiene que estar instalado).
- `copiar_solapa.py` — copia la primera hoja de un Excel como solapa del papel (si ya existe la reemplaza en su lugar).
- `fuente.py` — letras: Aptos Narrow 11 para lo que sale de Zetti, Arial 10 para lo que sale de ARCA. Nunca Calibri.
- Importes siempre con formato número `#,##0.00` (no contabilidad).
- Si la DDJJ de IVA ya está presentada, lo que se baje del Portal IVA va con `--presentada` (skill portal-iva-descarga).

## Paso 1 — Compras Zetti

1. Bajar las compras (con `--si-existe=reemplazar` si ya había una descarga anterior):

   ```bash
   cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/bajar-compras-zetti.js --cuit=<CUIT> --periodo=AAAAMM --si-existe=reemplazar
   ```

   Deja `AAAAMM - ZETTI - COMPRAS.xlsx` en `01-Impuestos Mensuales/AAAAMM/` (no en IVA), hoja "Compras Zetti":
   fila 1 razón social, CUIT y "Desde: DD-MM-AAAA  Hasta: DD-MM-AAAA"; fila 2 títulos (Fecha, TC, M, Nro. Comp. en
   D = PV y E = número, Proveedor, CUIT, R, Exen, Grav, IVA, P/R IIBB, P/R IVA, P/R IG, Total); comprobantes en el
   orden original de Zetti; fila TOTAL (I a O). Varias sociedades a la vez: `--cuits="A;B"` (un solo login).

   Controlar el JSON: `ok: true`, `control_total_reporte: "ok"` (la suma coincide con la fila Total: del reporte),
   `desde`/`hasta` cubren el mes completo y `fechas_fuera_de_periodo` vacío. Si hay `aviso_nro`, informar
   cuáles comprobantes quedaron con el número entero en la columna D (ej. NC de Monroe que Zetti trae como
   `00NC-A-3116-`): se dejan tal cual salen de Zetti.

2. Pegarlas en el papel como solapa "Compras Zetti", a continuación de "cruce ventas":

   ```bash
   cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/copiar_solapa.py" "<carpeta del mes>/AAAAMM - ZETTI - COMPRAS.xlsx" "<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx" "Compras Zetti"
   ```

   Antes de guardar, copiar el papel al scratchpad como respaldo. Si da `PermissionError`, el papel está abierto
   en Excel: pedir que lo cierre y volver a correr (no guardar con otro nombre).

3. Informar: cantidad de comprobantes por TC, totales (Exen, Grav, IVA, P/R IIBB, P/R IVA, P/R IG, Total),
   control contra el total del reporte y el orden de solapas resultante.

## Paso 2 — Compras portal iva

1. Bajar el Libro IVA Compras del Portal IVA con el script del plugin portal-iva-descarga, desde la carpeta de
   la última versión instalada:

   ```bash
   cd "$(ls -d ~/.claude/plugins/cache/estudiobva-skills/portal-iva-descarga/*/ | sort -V | tail -1)" && node scripts/portal-iva-descarga.js --cuit=<CUIT> --periodo=AAAAMM --libros=compras --no-archivar --si-existe=reemplazar
   ```

   - **Caso normal (período abierto):** sin `--presentada`. El script importa desde ARCA al borrador del Libro
     Compras, espera la importación y baja el CSV. Si ARCA ya tiene acuse de la DDJJ del período, el script corta:
     NO usar `--aunque-presentada` por iniciativa propia; avisar y preguntar.
   - **Período ya presentado:** agregar `--presentada` (solo lectura, última secuencia, no toca el borrador). Controla
     el neto gravado contra la vista previa de la DDJJ y, si no coincide, no guarda. Piloto Sanar 1 202609 se hizo así
     (F2083 DJ Original presentada el 19/10/2026).

   Con `--no-archivar` el Excel queda en `~/Documents/BVA-salidas/portal-iva/<carpeta> - AAAAMM - PORTAL IVA - COMPRAS.xlsx`.
   Copiarlo como `AAAAMM - PORTAL IVA - COMPRAS.xlsx` en `01-Impuestos Mensuales/AAAAMM/` (al lado del papel, no en IVA),
   igual que las ventas de la etapa 1.

   Controlar en el JSON: `ok: true`, `filasFueraDePeriodo: 0` y, con `--presentada`, `control_ddjj` (IVA del CSV vs DDJJ;
   diferencias de centavos son redondeo).

2. Pegarlo tal cual como solapa "Compras portal iva", a continuación de "Compras Zetti" (respaldo previo del papel):

   ```bash
   cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/copiar_solapa.py" "<carpeta del mes>/AAAAMM - PORTAL IVA - COMPRAS.xlsx" "<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx" "Compras portal iva"
   ```

   Sale como de ARCA: Arial 10, `#,##0.00`, títulos en fila 1 inmovilizados con filtro, fila TOTAL (=SUM) y nota al pie
   (las NC vienen en negativo, el total ya es neto). Mismo formato que "Ventas portal iva".

3. Informar: cantidad de comprobantes, DDJJ (si estaba presentada), control contra la DDJJ y orden de solapas.

## Paso 3 — cruce compras

Cruza "Compras Zetti" contra "Compras portal iva" y escribe (o reemplaza) la solapa "cruce compras" a continuación
de "Compras portal iva". Es el equivalente al "cruce" de ventas, con dos diferencias:
**de Zetti no se elimina nada** y **no se usa Mis Comprobantes**. Respaldo previo del papel, y papel cerrado en Excel.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/cruce_compras.py" --papel="<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx"
```

Cómo cruza: 1) CUIT + PV + número; 2) lo que sobra, por CUIT + clase (factura/NC) + mismo total ("Numero distinto",
ej. las NC de Monroe que Zetti guarda como `00NC-A-3116-`); 3) el resto queda "Solo en Portal IVA" / "Solo en Zetti".
Columnas: Exento/No grav., Gravado, IVA (suma por alícuota), Perc. IIBB, Perc. IVA, Otros (Zetti P/R IG; Portal Otros
Imp. Nac. + municipales + internos + otros tributos), Total.

**Comprobantes C (facturas, NC y ND C, tique factura C) = OK:** no tienen IVA. El Portal IVA los expone solo en
Importe Total (defecto del sistema, no es un error) y Zetti los carga en Exento o No gravado. El script toma el total
del Portal como Exento/No grav., quedan OK y se informa en Avisos. Nunca marcarlos como diferencia.

Contenido de la solapa (letra Arial, `#,##0.00`, mismos colores que el cruce de ventas):
1. Totales de cada fuente por grupo (Facturas y ND, NC, Liquidaciones, Comprobantes X).
2. Diferencia Zetti − Portal IVA por columna.
3. Qué explica las diferencias, por estado; la suma da el punto 2.
4. Detalle comprobante por comprobante (Portal / Zetti / diferencia) con observación.
5. Percepciones con el mismo total pero en otra columna o que no cierran, agrupado por proveedor.
6. Diferencias de total por proveedor.

Estados y observaciones que arma solo: Solo en Portal IVA (anulaciones FC/NC dentro del Portal, comprobantes en $0
o por centavos, liquidaciones de bancos, NC de Suizo/Monroe → revisar si es NCR), Solo en Zetti (comprobantes X no
fiscales), Duplicado en Zetti, Importe distinto, Numero distinto, Portal IVA sin desglose (solo trae el total, salvo comprobantes C),
Distinta apertura, Percepcion en otra columna (ojo con percepciones de IVA que en el Portal quedan en Otros Imp. Nac.:
no computan en la DDJJ) y Percepcion de Zetti no cierra (columnas de Zetti que no suman el total).

Informar: diferencia total Zetti − Portal, los estados con su importe y lo más relevante para revisar.

## Paso 4 — Servicios en Compras Zetti (AySA, EDESUR, METROGAS)

Filtrar en la solapa "Compras Zetti" las facturas de AySA (agua), EDESUR (luz) y METROGAS (gas). Respaldo previo del
papel, y papel cerrado en Excel.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/filtrar_servicios.py" --papel="<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx"
```

Los busca por CUIT (AySA 30-70956507-5, EDESUR 30-65511651-2, METROGAS 30-65786367-6) y por nombre, deja aplicado el
filtro de la fila 2 en la columna Proveedor y oculta el resto de los comprobantes (la fila TOTAL queda visible).
No cambia datos. Informar cada factura (fecha, comprobante, total) y qué servicio no tiene factura en Zetti ese mes.
Suelen ser las que no vienen con la importación del Portal IVA (en el cruce de compras figuran como "Solo en Zetti").

## Paso 5 — Agregar los servicios a Compras portal iva

Agregar las facturas de AySA / EDESUR / METROGAS filtradas en el paso 4 al final de "Compras portal iva" (debajo del
último comprobante) y volver a totalizar. Respaldo previo del papel, y papel cerrado en Excel.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/agregar_servicios_portal.py" --papel="<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx"
```

- **Alícuota de IVA:** antes de copiar se chequea con qué porcentaje está cargado el IVA en Zetti (IVA / Grav: 2,5 /
  5 / 10,5 / 21 / 27 %; los servicios suelen ser 27 %) y el neto y el IVA van a las columnas de esa alícuota. Si no da
  ninguna alícuota, el script corta sin grabar: avisar.
- **Registro de la alícuota en Compras Zetti (obligatorio):** debajo del TOTAL de "Compras Zetti" queda el bloque
  "Control alicuota IVA servicios": un renglón por factura con proveedor, comprobante, fila y alícuota en la columna
  Proveedor, y en la columna IVA la fórmula `=+J<fila>*0,27` (Grav × alícuota), como se controla a mano.
  La alícuota varía: la mayoría al 27 %, otras al 21 %, y a veces todo exento (queda "todo exento / no gravado, sin
  IVA" con 0). Las no tomadas por mal cargadas también figuran ahí. Si se vuelve a correr, el bloque se reemplaza.
- **Importes negativos sin ser NC** (suele pasar en el exento/no gravado): NO se agrega al Portal; en "Compras Zetti",
  al costado de su fila, queda "No lo tomamos porque esta mal cargado".
- Columnas como las carga ARCA: Exen → Importe Exento, P/R IIBB → Perc. IIBB, P/R IVA → Perc. IVA, P/R IG → Per. Otros
  Imp. Nac., Total Neto Gravado y Total IVA. Si ya está en el Portal (CUIT + PV + número), no se duplica.
- No se escribe nada a la derecha de las filas agregadas: la columna AG de "Compras portal iva" queda libre (se usa
  para la clasificación del paso 7). La fila TOTAL (=SUM) se rehace hasta la
  última fila nueva y la nota de abajo suma los agregados.
- Después, volver a correr el cruce del paso 3 (`cruce_compras.py`): esas facturas pasan a OK.

Informar: cada factura agregada con su alícuota, neto, IVA y total; las no tomadas por mal cargadas; el nuevo total.

## Paso 6 — Columna PROVEEDORES en Compras portal iva

En "Compras portal iva", celda AG1, el título `PROVEEDORES` con el mismo formato que los
demás títulos (copiar el estilo de AF1: Arial negrita blanca sobre azul 1F4E79), ancho ~22, y extender el filtro de la
fila 1 hasta AG. La columna queda vacía debajo del título (se completa en el paso siguiente).

```python
ws['AG1'] = 'PROVEEDORES'; ws['AG1']._style = copy(ws['AF1']._style)
ws.auto_filter.ref = 'A1:AG' + <ultima fila del filtro actual>
```

## Paso 7 — Clasificar proveedores y cuadro resumen (Compras portal iva)

Dos partes. **Si no queda ningún NO CLASIFICADO, la parte 2 se arma
directo, sin preguntar.** Solo se frena cuando falta clasificar algún proveedor (ver parte 1).

**Padrón:** `0 - ESTUDIO - 00000000000\Maestro de Proveedores.xlsx` de la unidad compartida (el script lo busca en
`BVA_UNIDAD_PATH`, I: y G:; si no, `--maestro=<ruta>`),
solapa "MAESTRO", títulos en la fila 2, datos desde la 3: A = CUIT, B = razón social, E = CUENTA CONTABLE (la
clasificación). El resto de las columnas se ignora.

Antes de empezar, verificar la distribución de columnas de "Compras portal iva" (F = Nro. Doc. Vendedor, H = Importe
Total, K = No Gravado, L = Exento, N a S = percepciones/impuestos, T = Neto 0 %, Y/Z = 10,5 %, AA/AB = 21 %,
AC/AD = 27 %, AG = PROVEEDORES). Si es otra, avisar antes de seguir.

### Parte 1 — Clasificar proveedores

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/clasificar_proveedores.py" --papel="<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx"
```

Busca el CUIT de F en el padrón (sin guiones ni ceros a la izquierda) y escribe la cuenta contable en AG. Los que no
están (o están sin cuenta) van como `NO CLASIFICADO`.
- **Si hay NO CLASIFICADOS:** PARAR y preguntar cómo se clasifica cada uno (CUIT, nombre, cantidad de
  comprobantes). Con su respuesta, escribir esa cuenta en AG y recién ahí seguir con la parte 2.
- **Si no hay ninguno:** seguir directo con la parte 2, sin preguntar.

### Parte 2 — Cuadro resumen

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python "scripts/cuadro_compras.py" --papel="<carpeta del mes>/AAAAMM- papeles de trabajo- <Sociedad>.xlsx"
```

Corta sin grabar si queda algún NO CLASIFICADO. Arma el cuadro dos filas debajo de la nota del Portal: G = tipo,
H = grupo, I a Q = los nueve conceptos, R = TOTAL; debajo el bloque de otros conceptos (G a J) y el CONTROL. Fórmulas
SUMIFS vivas sobre las filas 2 a la última de comprobantes; no hay Excel para recalcular, así que el script verifica
calculando lo mismo en Python e imprime los totales y la diferencia del control (sale con código 2 si es de $1 o más).
Si se vuelve a correr, borra lo que hay debajo de la nota y lo rehace. Informar los totales del cuadro, los
otros conceptos y el control. (Piloto Sanar 1 202609: diferencia −0,23.)

Abajo de los datos (dejando un par de filas libres), cuadro por tipo de comprobante y MERCADERIA/SERVICIOS.
Regla: AG = "CMV" es MERCADERIA; todo lo demás es SERVICIOS.

Quince filas fijas, en este orden, aunque queden en cero: 1 Factura A / SERVICIOS, 1 Factura A / MERCADERIA,
63 Liquidaciones A / SERVICIOS, 4 Recibo A / SERVICIOS, 81 Tique Factura A / SERVICIOS, 81 Tique Factura A /
MERCADERIA, 2 Nota de Débito A / SERVICIOS, 6 Factura B / SERVICIOS, 11 Factura C / SERVICIOS, 15 Recibo C /
SERVICIOS, 111 Tique Factura C / SERVICIOS, 3 Nota de Crédito A / MERCADERIA, 3 Nota de Crédito A / SERVICIOS,
8 Nota de Crédito B / MERCADERIA, 13 Nota de Crédito C / SERVICIOS. Una combinación tipo/grupo que no está en la
lista (ej. factura C de un proveedor CMV): agregar la fila y avisar.

Nueve columnas de conceptos: Neto Gravado 0 % = T, Neto Gravado 10,5 = Y, Neto Gravado 21 = AA, Neto Gravado 27 = AC,
No Gravado = K, Exento = L, IVA 10,5 = Z, IVA 21 = AB, IVA 27 = AD.
Comprobantes sin IVA discriminado (tipos 6, 8, 11, 13, 15 y 111): además, sumar a No Gravado el Importe Total (H)
menos las columnas N a S. Tipo cargado como texto (ej. "FC", servicios públicos): sumarlo en "1 - Factura A /
SERVICIOS" con un criterio más en la fórmula, SIN modificar la columna B.

SUMIFS vivas (no valores pegados). Formato contable con negativos en rojo. En las filas de NC (3, 8 y 13), tipo y
MERCADERIA/SERVICIOS en rojo aunque la fila quede en cero. Fila TOTAL PORTAL IVA al pie con el total general.

**Bloque aparte:** columnas N a S (percepciones, IIBB, municipales, IVA, internos y otros tributos) NO van en el
cuadro; van en un bloque separado debajo, abiertas por MERCADERIA / SERVICIOS / TOTAL.

**Control:** al final, cuadro + otros conceptos contra el Importe Total de H, con la diferencia. Recalcular y
verificar que la diferencia sea de centavos antes de entregar. Contar los totales del cuadro en el chat. No tocar
otras solapas.

Aclaración: que el No Gravado dé negativo no es un error (se compensa con el exento del mismo comprobante); no
marcarlo como observación.

## Próximos pasos

Se agregan en nuevas versiones del plugin.
