---
name: portal-iva-vulcamoia
description: "Vuelca las descargas del Portal IVA / Libro IVA Digital de ARCA (zips libro_iva_periodo_AAAAMM con comprobantes_compras.csv, comprobantes_ventas.csv, comprobantes_compras_impo_bienes.csv y libro_iva.xml) en los templates input_portal COMPRAS e input_portal VENTAS, clasifica cada proveedor con el padron 'Clasificacion Proveedores' (VULCAMOIA S.A. u otra sociedad con la misma estructura), arma el Resumen por Clasificacion mensual y el Control con Portal IVA, y deja 'input_portal_COMPRAS Completo.xlsx' e 'input_portal_VENTAS Completo.xlsx' en la carpeta 04-Balances y EECC/<año>/PORTAL IVA de la sociedad. Usar SIEMPRE que pidan \"procesar el Portal IVA de Vulcamoia\", \"cargar los zips del libro IVA en los input\", \"input_portal COMPRAS / VENTAS\", \"clasificar los proveedores de Vulcamoia\", \"resumen por clasificacion\", \"control con Portal IVA\" o mencionen zips libro_iva_periodo. No confundir con portal-iva-descarga (baja los libros de ARCA de las farmacias) ni con clasificacion-compras (clasifica con el Maestro de Proveedores del estudio)."
---

# Portal IVA → Inputs COMPRAS / VENTAS con clasificación de proveedores

Skill de Melina, pasada a Claude Code. Flujo que se hace todos los meses (o por tanda de
meses): se bajan del Portal IVA de ARCA los zips del Libro IVA Digital por período, se vuelcan
en dos plantillas Excel, se clasifican las compras por concepto con el padrón de proveedores de
la sociedad y se dejan controles que cierran contra el propio Libro IVA.

Todo pasa en la carpeta de la sociedad en la unidad compartida:
`<sociedad>/04-Balances y EECC/<año>/PORTAL IVA`. No entra a ARCA.

## Entradas (en esa carpeta)

- Uno o más `libro_iva_periodo_AAAAMM*.zip`. Cada zip trae `comprobantes_compras.csv`,
  `comprobantes_ventas.csv`, a veces `comprobantes_compras_impo_bienes.csv` (despachos) y
  `libro_iva.xml`. Si hay dos zips del mismo período, se usa el más reciente (y se avisa).
- Plantilla `input_portal COMPRAS.xlsx`: hoja PORTAL IVA, encabezado en fila 2 (Fecha …
  Clasificacion), datos desde fila 3, formato hasta fila 4186.
- Plantilla `input_portal VENTAS.xlsx`: hoja PORTAL IVA con meses como fechas en B5:M5, filas 6/7
  FACTURAS/NOTAS DE CREDITO, bloque CONTROL PORTAL IVA en filas 14, 15, 19, 20 y Comprobación
  en fila 23.
- Padrón `Clasificacion Proveedores ... .xlsx`: columnas NOM_PROVE, IDENTIFTRI (CUIT),
  CLASIFICACION. Si hay más de uno, se usa el más reciente (y se avisa).

El script reconoce como plantilla el `input_portal ... COMPRAS/VENTAS.xlsx` que todavía no fue
procesado (sin la hoja "Control con Portal IVA" / "Comprobantes Ventas") e ignora los "Completo"
y "Copia de". Si encuentra más de una, frena y las lista: preguntarle al usuario cuál es y pasarla
con `--compras` / `--ventas`. Si alguna tiene otra estructura, revisarlo con el usuario antes de
correr: el script asume esas posiciones.

## Paso 1: procesar

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/portal-iva-vulcamoia.py procesar --solo-revisar
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/portal-iva-vulcamoia.py procesar
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/portal-iva-vulcamoia.py procesar --si-existe=reemplazar
```

En Mac es `python3` (o `BVA_PYTHON` del `.env`).

| Argumento | Para qué |
|---|---|
| `--sociedad` | Nombre o CUIT de la carpeta de la sociedad. Default: Vulcamoia. |
| `--anio` | Año de `04-Balances y EECC/<año>/PORTAL IVA`. Default: el más nuevo que tenga PORTAL IVA. |
| `--carpeta` | La carpeta PORTAL IVA explícita (en vez de `--sociedad`/`--anio`). |
| `--compras`, `--ventas`, `--padron` | Archivos explícitos, si la detección no alcanza. |
| `--solo-revisar` | Procesa y muestra los controles sin guardar nada en la carpeta. |
| `--si-existe` | `error` (default) o `reemplazar`: los "Completo" anteriores van a `_anteriores/` con fecha y hora. |

Conviene correr primero con `--solo-revisar` y mostrarle al usuario el resultado.

Salida en la carpeta: `input_portal_COMPRAS Completo.xlsx` (PORTAL IVA + Control con Portal IVA +
Resumen por Clasificación + Proveedores sin clasificar) e `input_portal_VENTAS Completo.xlsx`
(PORTAL IVA completo + Comprobantes Ventas). Las plantillas y el padrón no se tocan. Los libros
quedan con recálculo al abrir: Excel calcula las fórmulas cuando se abren.

Por stdout, un JSON con: `entradas` (qué archivos usó), filas cargadas, despachos, `controles`
(por mes, calculados en Python con la misma lógica que las fórmulas), `sin_clasificar` (CUIT,
nombre, comprobantes, importe), variantes del padrón unificadas, CUIT con clasificación ambigua,
`guardados`, `anteriores`, `carpeta_trabajo` (temporal) y `avisos`.

## Qué hace y por qué (criterios acordados)

**COMPRAS – hoja PORTAL IVA**

- Una fila por comprobante, compras locales + despachos de importación (tipo 66: Punto de Venta
  vacío, número de despacho en "Número Desde"), ordenado por fecha.
- Las NC conservan el signo negativo que trae el portal, así las sumas netean solas.
- Columna L = G+H+I, salvo tipo C (11 y 13) donde toma el Importe Total (K). Los comprobantes C
  llegan del portal con neto/no gravado/exento en cero; son "Operaciones que no generan Crédito
  Fiscal" y se incluyen en L y en la clasificación.
- Columna M: clasificación buscada por CUIT en el padrón. Si un CUIT tiene varias
  clasificaciones, se toma la más frecuente (p. ej. EDESUR tenía "Seguros" una vez y "Servicio de
  luz…" 44). Lo no encontrado queda "SIN CLASIFICAR" en naranja: nunca dejar la celda vacía.

**Padrón**: las variantes de escritura se unifican al leerlo (Mercaderias/Mercaderia →
Mercaderías; Gasto de mantenimiento → Gastos de mantenimiento; Gastos Varios → Gastos varios;
Rodado → Rodados), así el resumen no parte una misma clasificación en varias filas. El archivo del
padrón no se modifica. Si el JSON reporta `padron_otras_variantes_posibles`, mencionárselas al
usuario y, si acepta, agregarlas al diccionario `VARIANTES` de `scripts/procesar_portal_iva.py`
(es un cambio del plugin: va por PR).

**COMPRAS – hoja "Control con Portal IVA"**: meses 1–12 + Total; bloques Neto Gravado
(Operaciones de Compras / NC de Compras) y Exento o No Gravado (Operaciones / NC / Operaciones que
no generan Crédito Fiscal), valores del libro_iva.xml en azul con comentario del campo de origen,
fila total = Neto − NC + Exento − NC + No CF. Debajo: "Total según detalle" (SUMIFS de la columna
L) y "Diferencia" (rojo si |dif| > 1). Mapeo XML: neto = totales_compras + bienes_usados +
importacion_bienes; NC = totales_compras_mlcf_restituir; exento = total_exento_no_gravado(_nc);
no CF = total_compras_no_fiscal.

**COMPRAS – hoja "Resumen por Clasificación"**: filas = clasificaciones (ordenadas por total),
columnas = meses del año + Total + %; fórmulas SUMIFS sobre L y M hasta la fila 4186 para que las
cargas futuras entren solas; al pie, control contra el total de la columna L y monto sin
clasificar. La "Diferencia (debe ser 0)" es menos lo que quedó SIN CLASIFICAR: da 0 recién cuando
se clasifican todos.

**VENTAS**: agrega la hoja "Comprobantes Ventas" (detalle); filas 6/7 con SUMIFS por mes
(facturas y ND / NC tipos 3, 8, 13 en positivo); bloque de control desde el XML: 14 = ventas RI +
consumidor final + otros conceptos, 15 = restituciones, 19 = exento/no gravado + exportación,
20 = exento NC. Consumidor final se suma en la fila 14 porque la fórmula de la fila 21 de la
plantilla no toma la 16. La fila 23 (Comprobación) debe dar 0.

## Paso 2: verificar antes de entregar

En el JSON:

- `controles.compras_cierran` y `controles.ventas_cierran` en `true` (equivalen a la fila 16 del
  Control de COMPRAS y a la fila 23 de VENTAS en 0 todos los meses).
- `controles.resumen_diferencia` = − `sin_clasificar_importe`.

Si algo no cierra, investigar antes de entregar (tipos de comprobante nuevos, un CSV faltante, un
campo del XML distinto; los CSV quedan en `carpeta_trabajo`). Las diferencias de centavos en IVA
son redondeo de ARCA.

## Paso 3: proveedores sin clasificar

Si quedaron proveedores sin clasificar, informar la lista (CUIT, nombre, cantidad, importe) y
ofrecer buscarlos en la web. Si el usuario acepta:

1. Buscar por nombre (las búsquedas por CUIT suelen no devolver nada útil). Para cada uno, anotar
   la actividad, una clasificación de las que ya existen en el padrón, confianza (Alta/Media/Baja)
   y fuente. Apoyarse en análogos del padrón (p. ej. otras concesionarias viales → Viáticos y
   movilidad; A.N.A → Mercaderías para despachos; terminales portuarias → Almacenamiento). No
   inventar: si no hay información, dejarlo sin clasificar y sugerir consultar la constancia de
   inscripción de ARCA o preguntar internamente qué se le compró. Con personas humanas, limitarse
   a la actividad económica.
2. Escribir un JSON con las sugerencias en `carpeta_trabajo` (no en la unidad):
   `[{"cuit": 30..., "clasificacion": "...", "nombre": "...", "actividad": "...", "confianza": "Alta", "fuente": "..."}]`
   y aplicarlas resaltadas en el COMPRAS Completo:

   ```bash
   cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/portal-iva-vulcamoia.py clasificar --json=<sugerencias.json> --modo=sugerida
   ```
3. Cuando el usuario confirme (o pase él mismo las clasificaciones, incluso en una imagen o
   tabla), correr con `--modo=confirmada`: saca el resaltado, completa la hoja "Proveedores sin
   clasificar" con Estado = "Agregado al padrón" y agrega los CUIT al padrón con el formato
   existente, después de guardar una copia del padrón en `_anteriores/`. Usar fuente "Usuario"
   cuando la clasificación la dio el usuario. Es el único paso que modifica el padrón: pedir el OK
   antes.
4. `clasificar` reconstruye el Resumen y deja el libro para recalcular al abrir. Para tener los
   controles otra vez en el JSON, volver a correr `procesar --si-existe=reemplazar` (toma el
   padrón ya actualizado).

## Entrega

Resumir en pocas líneas: comprobantes cargados, estado de los controles (cierran / diferencias),
clasificaciones principales del resumen, lo que quedó pendiente y dónde quedaron los archivos.
Responder en español rioplatense, en prosa breve, con el tono de un colega del estudio contable.
