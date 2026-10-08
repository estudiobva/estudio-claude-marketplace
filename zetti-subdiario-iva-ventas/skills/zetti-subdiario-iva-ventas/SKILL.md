---
name: zetti-subdiario-iva-ventas
description: "Baja de Zetti (T&S Web) el Subdiario de IVA Ventas (reporte 5.6.5) de una, varias o todas las sociedades para un mes (por defecto el mes anterior), lo deja limpio y ordenado por tipo de comprobante en Excel y lo archiva en la unidad compartida como 'AAAAMM - ZETTI - VENTAS.xlsx' en la carpeta IVA del mes, al lado de 'AAAAMM - PORTAL IVA - VENTAS.xlsx'. Corre un script de Playwright que se loguea solo con el usuario de Zetti de la planilla de claves del estudio: no hace falta tener la clave guardada en Chrome. Usar SIEMPRE que pidan \"bajá el IVA ventas de [sociedad] [mes]\", \"subdiario de IVA ventas\", \"reporte 5.6.5\", \"rpt_subdiario_iva_venta\", \"ventas Zetti del mes\" o \"bajar Zetti de todas las sociedades\". Para cruzar contra el Portal IVA usar cruce-ventas-zetti-portal-iva."
---

# Subdiario de IVA Ventas desde Zetti (T&S Web)

Un script de Playwright entra a T&S Web con el usuario del estudio, busca la
sociedad **por CUIT**, genera el reporte **5.6.5 Reporte de IVA ventas** en CSV
para el mes, lo pasa a Excel y lo archiva en la unidad compartida. El CSV se
borra: queda solo el `.xlsx`. Solo lee: no cambia nada en Zetti.

Se trabaja a **mes vencido**: sin `--periodo` toma el mes anterior al de hoy.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-ventas.js --nombre="<sociedad>"                    # mes anterior
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-ventas.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-ventas.js --nombres="<A>;<B>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-ventas.js --todas --periodo=09/2026 --si-existe=saltear
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-ventas.js --listar     # entidades de T&S Web con su CUIT
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un solo login para todas. |
| `--todas` | Todas las sociedades de Zetti que tienen carpeta en la unidad. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. Siempre del dia 1 al ultimo dia del mes. |
| `--si-existe` | Si ya esta el `.xlsx`: `abortar` (default), `saltear` o `reemplazar`. |
| `--entidad` | Nombre exacto de la entidad de T&S Web, solo si el CUIT no alcanza para elegirla. |
| `--no-archivar` | Deja el Excel en `~/Documents/BVA-salidas/zetti/` y no toca el Drive. |
| `--ver` | Navegador visible. Solo si el usuario lo pide: por defecto va oculto. |

Si falta la sociedad, preguntar antes de correr. Si falta el mes, usar el mes
anterior y decirlo.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Usa el mismo `~/.fisco-ar/.env` que los demas plugins (`CLAVES_ORGANISMOS_PATH`
apuntando a la planilla de la unidad, `BVA_UNIDAD_PATH` en Mac).

## Login

El usuario y la clave de Zetti estan en `Claves_Organismos.xlsx`, grupo de
columnas **ZETTI** (Usuario | Clave), fila **"Zetti T&S Web (todas las
sociedades)"**: es un solo usuario para todo el grupo. Si una sociedad necesitara
otro usuario, se carga en su propia fila y tiene prioridad. Alternativa puntual:
`ZETTI_USUARIO` / `ZETTI_CLAVE` en el `.env`.

Claude nunca escribe ni muestra la clave. Si hay que cambiarla, la carga la persona
en la planilla (o con `python3 scripts/cargar-clave-zetti.py "<ruta a Claves_Organismos.xlsx>"`,
que la pide sin mostrarla). Ojo: Excel convierte una clave numerica con 0 adelante
en numero y le saca el 0; las columnas ZETTI estan en formato texto para evitarlo.

## Como elige la sociedad en Zetti

T&S Web tiene ~140 entidades: las **Sociedades** ("SOC ...") y sus **Locales**
(una sociedad puede tener varios). La lista no trae CUIT, pero la API de nodos si:
el script arma el mapa CUIT → entidades y usa la de mas arriba del arbol (la
Sociedad), que abarca todos los locales. Por eso no hay que elegir entre
"S64 - SOY ..." y "SOC ...": el resultado es el de toda la sociedad.

Formulario del 5.6.5 que completa: Nodo el de la sociedad, Emision Desde
`01/MM/AAAA`, Emision Hasta el ultimo dia del mes, Periodo IVA vacio, Tipo `Todos`,
Valor a discriminar el que viene, todos los tildes sin tildar, Orden `Tipo
Comprobante`, Filtro `Mostrar todos`, Formato `CSV`.

## Controles antes de archivar

No archiva (y lo informa con `ok: false`) si:
- el CUIT del encabezado del reporte no es el de la sociedad pedida;
- las fechas del reporte o de algun comprobante no son del periodo;
- la suma de los comprobantes no da el **Total** del reporte al centavo. En Exento
  y Gravado se toleran hasta $0,05 cruzados (el reporte redondea con decimales
  internos): sale `control_total_reporte: "redondeo"` con el detalle;
- hay filas que no son comprobantes ni el total, o el reporte vino vacio.

## Que entrega

`[Sociedad]/01-Impuestos Mensuales/AAAAMM/IVA/AAAAMM - ZETTI - VENTAS.xlsx` (la
misma carpeta que el Portal IVA: con carpeta de año va en `AAAA/AAAAMM/IVA/`, y si
existe `AAAAMM-IVA` se usa esa). Hoja **"Ventas Zetti"** con el mismo formato que
la solapa "Ventas Zetti" del papel de trabajo mensual, para copiarla tal cual:

- Fila 1: `<razon social>  -  CUIT NN-NNNNNNNN-N  -  Subdiario de IVA Ventas MM/AAAA`.
- Fila 2: Fecha | TC | M | Nro. Comp. | (sin titulo) | Cliente | CUIT | RESP | Exen | Grav | IVA | P.IB | Total.
- Datos desde la fila 3. El comprobante va partido: D = punto de venta y E =
  numero, los dos como numero (`0006-00000070` -> 6 y 70).
- Fechas como fecha, importes con formato contable `$`. Ordenado por TC, letra
  (A antes que B) y fecha.
- Sin fila de totales, sin P.IVA y sin el codigo 901/902 del reporte (el papel no
  los tiene). Si P.IVA viniera con importe, el JSON trae `aviso_piva` (el Total
  lo incluye).

Validado contra el papel de Sanar 64 08/2026: las 60 filas iguales celda por
celda (ese papel no tenia el 31/08 porque se exporto hasta el 30/08).

## Respuesta al usuario

En 2-3 lineas por sociedad: sociedad, periodo, cantidad de comprobantes por TC y
totales de Exen, Grav, IVA y Total (del JSON), y donde quedo el archivo. Si hubo
`redondeo`, `aviso_piva` o `aviso_nro`, mencionarlo en una linea. Si una sociedad fallo, decir el motivo
(`error` / `code`) en vez de adivinar:

- `sin_credenciales` / `clave_rechazada`: falta o esta mal la fila ZETTI de la planilla.
- `sin_entidad`: ese CUIT no esta en T&S Web (no usa Zetti o no esta habilitado para el usuario).
- `ya_existe`: ya estaba el archivo; preguntar si reemplazar.
- `control`: el reporte no paso un control; no se archivo nada.
- `sin_descarga` / `reporte`: T&S Web no respondio; reintentar mas tarde (si falla 2-3 veces, frenar y consultar).

Solo ventas. Para cruzar contra el Portal IVA usar `cruce-ventas-zetti-portal-iva`.
