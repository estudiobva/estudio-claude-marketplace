---
name: facturacion-arca
description: Emite facturas electrónicas de servicios en ARCA (Comprobantes en línea / RCEL) para cualquier emisor — monotributista (Factura C) o responsable inscripto (Factura A/B) — con un script de Playwright que se loguea solo con la planilla de claves del estudio. Toma los datos de la planilla mensual de facturación (o del chat) y arma la planilla modelo para pedirle los datos al cliente. Usar SIEMPRE que pidan facturar, emitir comprobantes, "hacer las facturas" de alguien, mencionen RCEL, CAE, Factura A/B/C o un CUIT emisor, o pidan la planilla para que un cliente pase los datos — incluso de forma indirecta ("hay que facturarle a X", "armá lo de este mes de Y"). ESCRIBE EN ARCA: cada factura sale con CAE y no se puede deshacer.
---

# Facturación en ARCA con script

`scripts/facturacion-arca.js` entra a ARCA con la clave del emisor (planilla de
claves del estudio o `~/.fisco-ar/.env`, igual que anticipos-sct), abre
Comprobantes en línea, elige la empresa del emisor y carga cada factura.

## Seguridad (no negociable)

- Son comprobantes reales con CAE y **no se pueden deshacer**: sólo se anulan con
  nota de crédito.
- **Nunca** correr `--emitir` sin haber mostrado antes el resumen (tabla + total)
  y tener el OK explícito del usuario **para ese código**. Un OK anterior no vale
  para un código nuevo.
- Las claves salen de la planilla de claves o del `.env`. Si alguien pega una
  clave en el chat: no usarla, no repetirla, no guardarla, y recomendar cambiarla.
- Navegador oculto salvo que el usuario pida verlo (`--ver`).

## Paso 0 — Emisor

Los datos del emisor van en la hoja **Emisor** de su planilla: nombre como figura
en ARCA, CUIT, condición IVA, punto de venta, actividad (código), alícuota (sólo
RI) y carpeta de los PDF. El tipo sale solo:

- Monotributo → **Factura C** (sin IVA).
- Responsable Inscripto → **Factura A** a RI o monotributistas; **Factura B** al
  resto. Es el caso de más riesgo: mostrá neto, IVA y total.

Si el emisor no tiene clave en la planilla de claves ni en el `.env`, el script
lo dice (`sin_credenciales`): hay que cargarla ahí, nunca pasarla por el chat.

Si el usuario necesita pedirle los datos a alguien, copiá
`assets/Plantilla Facturacion ARCA.xlsx`, precompletá Emisor y los clientes
conocidos y entregásela. Cada mes se copia la hoja "Plantilla" con el nombre del mes.

## Paso 1 — Validar y mostrar el resumen

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="<Mes AAAA>" --validar
```

No entra a ARCA. Toma las filas con CUIT y `¿Emitida?` distinto de "Sí" y
controla: CUIT de 11 dígitos válido y cargado en Clientes, fechas completas (no se
inventan), período hasta ≥ desde, vencimiento ≥ fecha, fecha a ±10 días de hoy,
importe > 0, IVA = neto × alícuota. Si hay `errores`, pasáselos al usuario tal
cual y esperá que corrija la planilla.

Mostrá la `tabla` (markdown) del JSON con el total y los `avisos` (p. ej. una
condición de venta tomada de la habitual del cliente), y pedí la aprobación.

Si los datos vienen del chat (una sola factura), el script acepta argumentos en
vez de planilla: `--cuit --pv --condicion=mono|ri --actividad --alicuota
--receptor --cond-receptor=ri|mono|exento|cf --desc --importe (neto) --fecha
--desde --hasta --vto --cond-venta`. Pedí cada dato: nada de defaults inventados.

## Paso 2 — Dry-run (recomendado siempre con un emisor nuevo)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="<Mes AAAA>"
```

Recorre cada factura hasta la pantalla de confirmación y **no confirma**.
Devuelve, por factura, la razón social que trajo el padrón (`razonSocialArca`) y
una captura. Si la planilla dice otro nombre, sale un `aviso`: mostralo.

## Paso 3 — Emitir, con el OK del usuario

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="<Mes AAAA>" --emitir=<código>
```

El código es el del resumen aprobado. Si la planilla cambió, no coincide y no
emite nada (`codigo_no_coincide`): volver al paso 1.

Por cada factura el script controla la pantalla de confirmación (CUIT receptor,
total, neto/IVA en la A, fecha, tipo) antes de confirmar, busca el comprobante
nuevo en RCEL → Consultas, baja el PDF como `AAAAMM - [Emisor] - [Cliente].pdf`
y marca la fila con N° de comprobante y `¿Emitida?` = "Sí".

## Si algo falla

- **`verificar_a_mano` / `corte` después de confirmar: NO reintentar.** Puede
  haberse emitido igual y un reintento duplica la factura. Que el usuario mire
  RCEL → Consultas; si salió, anotarla en la planilla y borrar esa entrada de
  `~/Documents/BVA-salidas/facturas/.registro-<cuit>.json`.
- Un error antes de confirmar (campo que no se encontró, dato rechazado) no
  generó nada: se ve en `resultados` y en los dumps de
  `~/Documents/BVA-salidas/facturas/_debug/`.
- `avisoPlanilla`: la factura salió pero la planilla estaba abierta; anotar el
  número a mano. El registro ya la tiene y no se reemite.
- `captcha` / `credenciales_invalidas`: no reintentar; reintentar bloquea la clave.

## Cierre

- Una factura con CAE por cada fila aprobada, ni más ni menos; la suma emitida
  (`totalEmitido`) igual al total del resumen aprobado.
- Resumen final en tabla: cliente, tipo, N° de comprobante, CAE, importe y total.
- Entregar los PDF (están en `carpetaPdf`).
