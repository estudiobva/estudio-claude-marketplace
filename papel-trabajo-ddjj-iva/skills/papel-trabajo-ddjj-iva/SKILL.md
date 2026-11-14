---
name: papel-trabajo-ddjj-iva
description: Etapa 2 del papel de trabajo mensual de una farmacia/sociedad (la etapa 1 es la de ventas para TSH e IIBB). Prepara los papeles de trabajo para presentar la DDJJ de IVA sobre el MISMO papel `AAAAMM- papeles de trabajo- <Sociedad>.xlsx` que ya trae Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva (con cuadro resumen), Ventas MCR y cruce ventas. Paso 1 - baja el Subdiario de IVA Compras de Zetti (5.6.4) y lo pega como solapa "Compras Zetti" después de "cruce ventas". Paso 2 - baja el Libro IVA Compras del Portal IVA y lo pega como "Compras portal iva" al lado. Paso 3 - cruza compras Zetti vs Portal IVA en la solapa "cruce compras". Paso 4 - filtra en "Compras Zetti" las facturas de AySA, EDESUR y METROGAS. Paso 5 - las agrega al final de "Compras portal iva" y vuelve a totalizar. Paso 6 - titulo PROVEEDORES en la columna AG de "Compras portal iva". Paso 7 - clasifica los proveedores con el Maestro de Proveedores y arma el cuadro resumen de compras. Paso 8 - completa la solapa "imputacion compras" (resumen CMV / SERVICIOS y detalle por cuenta contable). Lo hace de punta a punta un script de Playwright (scripts/papel-ddjj-iva.js) con login automático. Usar SIEMPRE que pidan "preparar el papel de IVA de [sociedad] [periodo]", "DDJJ de IVA", "liquidación de IVA", "bajar las compras de Zetti y pegarlas en el papel", "etapa 2 del papel de trabajo". No se corre antes del día 11 del mes siguiente al período.
---

# Papel de trabajo para la DDJJ de IVA (etapa 2)

Se trabaja sobre el papel de trabajo del mes que armó la etapa 1 (ventas para TSH e IIBB, plugin `papel-trabajo-ventas-tsh`).
Lo hace un solo script de Playwright, con login automático (planilla de claves del estudio). No pedir ni escribir contraseñas.

## Datos de entrada

- Sociedad (nombre o CUIT) y período (AAAAMM). El script busca la carpeta de la sociedad por CUIT en la unidad compartida
  (`BVA_UNIDAD_PATH` del `.env`). Si te dan solo el nombre, buscá la carpeta (`<N> - <Sociedad> - <CUIT>`).
- Papel: `<Sociedad>/01-Impuestos Mensuales/AAAAMM/AAAAMM- papeles de trabajo- <Sociedad>.xlsx`. Si la etapa 1 quedó en
  `controles`, agregar `--subcarpeta=controles`; o pasar la ruta con `--papel=`. Si no existe o no tiene "cruce ventas",
  el script corta: la etapa 1 no está terminada.
- Fecha: no se corre antes del día 11 del mes siguiente al período (el script corta; `--antes-del-11` solo si el usuario lo pide).
- El papel tiene que estar **cerrado en Excel**.

## Correr el script

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=AAAAMM > "$HOME/Documents/BVA-salidas/papel-ddjj-iva/ultimo.json" 2> "$HOME/Documents/BVA-salidas/papel-ddjj-iva/ultimo.log"
```

Corre los 8 pasos en orden (abajo, qué hace cada uno). Antes de tocar el papel lo copia a `BVA-salidas/papel-ddjj-iva/respaldos`.
Navegador siempre oculto; `--ver` solo si el usuario pide ver la ventana. Nunca escribe en Zetti.

**De dónde sale el Libro IVA Compras (paso 2), `--portal=`:**
- `auto` (por defecto): el de la DDJJ **presentada**, en solo lectura y controlado contra la vista previa de la DDJJ.
  Si el período no está presentado, el script para con código 4 y `falta_decidir`. **Preguntar al usuario** y volver a correr
  con `--desde-paso=2` y una de estas:
  - `--portal=iva`: el `AAAAMM - PORTAL IVA - COMPRAS.xlsx` que ya está en la carpeta IVA del mes (el que bajó el circuito NCR).
  - `--portal=<ruta.xlsx>`: otro archivo.
  - `--portal=borrador`: importa desde ARCA al borrador del Libro Compras y lo baja. **Escribe en el borrador de ARCA** (no presenta).
    Si ya se eliminaron NCR del borrador (conciliacion-ncr), importar puede volver a traerlas: usarlo solo con el OK del usuario.

**Códigos de salida:** 0 OK · 1 falló un paso (`paso_fallido`, `pasos.<paso>.error`) · 2 papel no listo o argumentos ·
3 ARCA pidió captcha o rechazó la clave (no reintentar en loop) · 4 falta una decisión (`falta_decidir`) · 5 el control del
cuadro difiere en $1 o más.

**NO CLASIFICADOS (paso 7):** si algún proveedor no está en el Maestro (o está sin cuenta), el script para con código 4 y lista
CUIT, nombre, comprobantes y total en `pasos.7_clasificacion.no_clasificados`. Preguntar al usuario la cuenta de cada uno y volver a correr:

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=AAAAMM --desde-paso=7 --cuentas="<CUIT>=<CUENTA>;<CUIT>=<CUENTA>"
```

Eso escribe la cuenta solo en el papel; recordarle al usuario que la cargue en el Maestro de Proveedores para la próxima.

**Retomar:** `--desde-paso=N` (y `--hasta-paso=N`) corre solo esos pasos, por ejemplo después de corregir algo a mano.

**Informar al terminar**, con el JSON (también queda en `BVA-salidas/papel-ddjj-iva/<CUIT>-<AAAAMM>-<fecha>.json`):
1. Compras Zetti: comprobantes por TC, totales, control contra el total del reporte (`control_total_reporte`), `aviso_nro` si hay.
2. Compras portal iva: modo (presentada / archivo / borrador), DDJJ, comprobantes, control contra la DDJJ (centavos = redondeo).
3. Cruce compras: diferencia Zetti − Portal por columna, los estados con su importe y lo más relevante para revisar.
4-5. Servicios: cada factura de AySA / EDESUR / METROGAS (fecha, comprobante, total), cuáles se agregaron al Portal con su
   alícuota, las no tomadas por mal cargadas y qué servicio no tiene factura en Zetti.
7. Clasificación por cuenta y cuadro: totales del cuadro, otros conceptos y el control (diferencia de centavos = OK).

## Qué hace cada paso

Scripts en `scripts/`; cada uno se puede correr suelto (`python scripts/<x>.py --papel="<papel>"`) e imprime un JSON.
Letras: Aptos Narrow 11 para lo que sale de Zetti, Arial 10 para lo que sale de ARCA (`fuente.py`). Nunca Calibri.
Importes con formato número `#,##0.00` (no contabilidad).

### Paso 1 — Compras Zetti (`bajar-compras-zetti.js` + `copiar_solapa.py`)

Reporte 5.6.4 (formName `rpt_subdiario_iva_compra`), sociedad completa, emisión del 1 al último día, CSV → Excel con
`zetti_compras_excel.py`. Deja `AAAAMM - ZETTI - COMPRAS.xlsx` al lado del papel (no en IVA), hoja "Compras Zetti":
fila 1 razón social, CUIT y "Desde/Hasta"; fila 2 títulos (Fecha, TC, M, Nro. Comp. en D = PV y E = número, Proveedor,
CUIT, R, Exen, Grav, IVA, P/R IIBB, P/R IVA, P/R IG, Total); comprobantes en el orden original de Zetti; fila TOTAL (I a O).
La pega como solapa "Compras Zetti" a continuación de "cruce ventas".
Controles: `control_total_reporte: "ok"` (la suma coincide con la fila Total: del reporte), mes completo, sin fechas fuera
del período. `aviso_nro`: comprobantes con el número entero en la columna D (ej. NC de droguería que Zetti trae como
`00NC-A-3116-`): se dejan tal cual salen de Zetti.

### Paso 2 — Compras portal iva (plugin `portal-iva-descarga` + `copiar_solapa.py`)

Según `--portal` (ver arriba). Con la DDJJ presentada: última secuencia, solo lectura, controla el neto gravado contra la
vista previa de la DDJJ y, si no coincide, no guarda. Copia el libro como `AAAAMM - PORTAL IVA - COMPRAS.xlsx` al lado del
papel (no en IVA) y lo pega tal cual como "Compras portal iva": Arial 10, `#,##0.00`, títulos en fila 1 inmovilizados con
filtro, fila TOTAL (=SUM) y nota al pie (las NC vienen en negativo, el total ya es neto).

### Paso 3 — cruce compras (`cruce_compras.py`)

Cruza "Compras Zetti" contra "Compras portal iva" en la solapa "cruce compras". Como el cruce de ventas, con dos
diferencias: **de Zetti no se elimina nada** y **no se usa Mis Comprobantes**.

Cómo cruza: 1) CUIT + PV + número; 2) lo que sobra, por CUIT + clase (factura/NC) + mismo total ("Numero distinto");
3) el resto queda "Solo en Portal IVA" / "Solo en Zetti". Columnas: Exento/No grav., Gravado, IVA (suma por alícuota),
Perc. IIBB, Perc. IVA, Otros (Zetti P/R IG; Portal Otros Imp. Nac. + municipales + internos + otros tributos), Total.

**Comprobantes que el Portal trae solo en Importe Total = OK** (facturas C, facturas B y cualquier otro tipo): el
Portal IVA los expone solo en Importe Total (defecto del sistema, no es un error) y Zetti los carga en Exento o No
gravado; el cuadro resumen del paso 7 los suma en No Gravado. El script toma el total del Portal como Exento/No grav.,
quedan OK y se informa en Avisos (cantidad por tipo). Nunca marcarlos como diferencia.

Contenido de la solapa: 1) totales de cada fuente por grupo (Facturas y ND, NC, Liquidaciones, Comprobantes X);
2) diferencia Zetti − Portal IVA por columna; 3) qué explica las diferencias, por estado (la suma da el punto 2);
4) detalle comprobante por comprobante con observación; 5) percepciones con el mismo total pero en otra columna o que
no cierran, por proveedor; 6) diferencias de total por proveedor.

Estados: Solo en Portal IVA (anulaciones FC/NC dentro del Portal, comprobantes en $0 o por centavos, liquidaciones de
bancos, NC de Suizo/Monroe → revisar si es NCR), Solo en Zetti (comprobantes X no fiscales), Duplicado en Zetti, Importe
distinto, Numero distinto, Portal IVA sin desglose (salvo comprobantes C), Distinta apertura, Percepcion en otra columna
(ojo con percepciones de IVA que en el Portal quedan en Otros Imp. Nac.: no computan en la DDJJ) y Percepcion de Zetti
no cierra (columnas de Zetti que no suman el total).

### Paso 4 — Servicios en Compras Zetti (`filtrar_servicios.py`)

Filtra en "Compras Zetti" las facturas de AySA (agua), EDESUR (luz) y METROGAS (gas), por CUIT (AySA 30-70956507-5,
EDESUR 30-65511651-2, METROGAS 30-65786367-6) y por nombre: deja el filtro de la fila 2 en la columna Proveedor y oculta
el resto (la fila TOTAL queda visible). No cambia datos. Suelen ser las que no vienen con la importación del Portal IVA
(en el cruce figuran como "Solo en Zetti").

### Paso 5 — Servicios agregados a Compras portal iva (`agregar_servicios_portal.py`)

- **Alícuota de IVA:** IVA / Grav de Zetti (2,5 / 5 / 10,5 / 21 / 27 %; los servicios suelen ser 27 %): el neto y el IVA
  van a las columnas de esa alícuota. Si no da ninguna, corta sin grabar: avisar.
- **Registro de la alícuota en Compras Zetti:** debajo del TOTAL queda el bloque "Control alicuota IVA servicios", un
  renglón por factura con proveedor, comprobante, fila y alícuota, y en la columna IVA `=+J<fila>*0,27` (Grav × alícuota).
  A veces todo exento ("todo exento / no gravado, sin IVA" con 0). Si se vuelve a correr, el bloque se reemplaza.
- **Importes negativos sin ser NC:** NO se agrega; en "Compras Zetti", al costado, "No lo tomamos porque esta mal cargado".
- Columnas como las carga ARCA: Exen → Importe Exento, P/R IIBB → Perc. IIBB, P/R IVA → Perc. IVA, P/R IG → Per. Otros
  Imp. Nac., Total Neto Gravado y Total IVA. Si ya está en el Portal (CUIT + PV + número), no se duplica.
- La columna AG queda libre; la fila TOTAL (=SUM) se rehace hasta la última fila nueva y la nota suma los agregados.
- Si agregó alguna, el script vuelve a correr el cruce: esas facturas pasan a OK.

### Paso 6 — Columna PROVEEDORES (`titulo_proveedores.py`)

AG1 = `PROVEEDORES` con el estilo de AF1 (Arial negrita blanca sobre azul 1F4E79), ancho 22, filtro de la fila 1 hasta AG.

### Paso 7 — Clasificar proveedores y cuadro resumen (`clasificar_proveedores.py`, `asignar_cuentas.py`, `cuadro_compras.py`)

**Padrón:** `0 - ESTUDIO - 00000000000/Maestro de Proveedores.xlsx` de la unidad compartida (`BVA_UNIDAD_PATH`; si no,
`--maestro=<ruta>` al script suelto), solapa "MAESTRO", títulos en la fila 2: A = CUIT, E = CUENTA CONTABLE.
Busca el CUIT de F en el padrón y escribe la cuenta en AG; los que no están (o sin cuenta) van como `NO CLASIFICADO`
y el script para (ver arriba). Distribución esperada de "Compras portal iva": F = Nro. Doc. Vendedor, H = Importe Total,
K = No Gravado, L = Exento, N a S = percepciones/impuestos, T = Neto 0 %, Y/Z = 10,5 %, AA/AB = 21 %, AC/AD = 27 %,
AG = PROVEEDORES; si es otra, los scripts cortan.

Cuadro resumen: Corta sin grabar si queda algún NO CLASIFICADO. Arma el cuadro dos filas debajo de la nota del Portal: G = tipo,
H = grupo, I a Q = los nueve conceptos, R = TOTAL; debajo el bloque de otros conceptos (G a J) y el CONTROL. Fórmulas
SUMIFS vivas sobre las filas 2 a la última de comprobantes; no hay Excel para recalcular, así que el script verifica
calculando lo mismo en Python e imprime los totales y la diferencia del control (sale con código 2 si es de $1 o más).
Si se vuelve a correr, borra lo que hay debajo de la nota y lo rehace. Informar los totales del cuadro, los
otros conceptos y el control. (Piloto 09/2026: diferencia −0,23.)

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

### Paso 8 — imputacion compras (`imputacion_compras.py`)

Completa la solapa "imputacion compras" (si el papel ya la trae como plantilla, respeta rótulos y formato; si no, la
crea después de "Compras portal iva") con SUMIFS vivas contra "Compras portal iva" (filas 2 a la última de
comprobantes), según la clasificación de la columna AG (paso 7):

- **Resumen global (filas 3 a 5):** CMV (AG = "CMV"), SERVICIOS (todo lo demás) y Total. Columnas: Comprobantes,
  Neto Gravado (T + U + W + Y + AA + AC, los netos de cada alícuota; no AE, que en filas agregadas a mano suele quedar
  vacía), No Gravado (K, y en los tipos sin IVA discriminado 6, 8, 11, 13, 15 y 111 además H − N a S, como el cuadro del
  paso 7), Exento (L), IVA (V + X + Z + AB + AD), Percepciones y otros impuestos (N a S, todo junto) e Importe
  Total (H): neto + no gravado + exento + IVA + percepciones = Importe Total en cada fila. Si la plantilla tiene
  Importe Total en G, se agrega la columna de percepciones en G y el total pasa a H.
- **Detalle por cuenta contable (fila 9 en adelante):** Comprobantes y "gravado + exento" = K + L + T + U + Y + AA + AC
  (más el mismo agregado de los tipos sin IVA, H − N a S, si no las facturas C quedarían en cero). **Nunca incluye
  percepciones ni otros impuestos.** Se mantienen las cuentas de la
  plantilla aunque den cero y se **agregan las que aparezcan en AG y falten**, en orden alfabético; al pie, fila Total.

Controles (calculados en Python, sale con código 2 si no cierran): comprobantes del detalle = fila SERVICIOS, y
neto + no gravado + exento + IVA + N a S = Importe Total (centavos). Informar el resumen, las cuentas agregadas y las
que quedaron en cero. (Piloto Quilmes Oeste 09/2026: se agregó GASTOS DE MANTENIMIENTO; control 0,09.)

## Próximos pasos

Se agregan en nuevas versiones del plugin.
