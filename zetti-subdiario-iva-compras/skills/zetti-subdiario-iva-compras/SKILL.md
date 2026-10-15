---
name: zetti-subdiario-iva-compras
description: "Baja de Zetti (T&S Web) el Subdiario de IVA Compras (reporte 5.6.4) de una, varias o todas las sociedades para un mes (por defecto el mes anterior), lo pasa a Excel con el formato exacto de la solapa 'Compras Zetti' del papel de trabajo y lo archiva en la unidad compartida como 'AAAAMM - ZETTI - COMPRAS.xlsx' en la carpeta IVA del mes, al lado de 'AAAAMM - PORTAL IVA - COMPRAS.xlsx'. Corre un script de Playwright que se loguea solo con el usuario de Zetti de la planilla de claves del estudio. Usar SIEMPRE que pidan \"bajá las compras de Zetti de [sociedad] [mes]\", \"subdiario de IVA compras\", \"reporte 5.6.4\", \"rpt_subdiario_iva_compra\", \"compras Zetti del mes\", \"la solapa Compras Zetti del papel\" o \"bajar compras de Zetti de todas las sociedades\". Para ventas usar zetti-subdiario-iva-ventas; para el Libro IVA Compras de ARCA, portal-iva-descarga."
---

# Subdiario de IVA Compras desde Zetti (T&S Web)

Un script de Playwright entra a T&S Web con el usuario del estudio, busca la
sociedad **por CUIT**, genera el reporte **5.6.4 Reporte IVA Compra** en CSV
para el mes, lo pasa a Excel con el formato de la solapa "Compras Zetti" de los
papeles de trabajo y lo archiva en la unidad compartida. El CSV se borra: queda
solo el `.xlsx`. Solo lee: no cambia nada en Zetti.

Se trabaja a **mes vencido**: sin `--periodo` toma el mes anterior al de hoy.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-compras.js --nombre="<sociedad>"                    # mes anterior
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-compras.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-compras.js --nombres="<A>;<B>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-compras.js --todas --periodo=09/2026 --si-existe=saltear
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/zetti-iva-compras.js --listar     # entidades de T&S Web con su CUIT
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un solo login para todas. |
| `--todas` | Todas las sociedades de Zetti que tienen carpeta en la unidad. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. Siempre del dia 1 al ultimo dia del mes (fecha de emision). |
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

Igual que `zetti-subdiario-iva-ventas`: usuario y clave de Zetti en
`Claves_Organismos.xlsx`, grupo de columnas **ZETTI**, fila **"Zetti T&S Web
(todas las sociedades)"** (o `ZETTI_USUARIO` / `ZETTI_CLAVE` en el `.env`).
Claude nunca escribe ni muestra la clave; si hay que cambiarla, la carga la
persona (`python3 scripts/cargar-clave-zetti.py "<ruta a Claves_Organismos.xlsx>"`
la pide sin mostrarla).

## Como elige la sociedad en Zetti

Por CUIT, con la API de nodos de T&S Web: usa la entidad de mas arriba del arbol
(la Sociedad, "SOC ..."), que abarca todos sus locales.

Formulario del 5.6.4 que completa: Nodo el de la sociedad, Emision Desde
`01/MM/AAAA`, Emision Hasta el ultimo dia del mes, Periodo IVA y Tipo Entidad
vacios, Filtro `Mostrar todos`, todos los tildes sin tildar (Excluir Importados,
Resumido, Ocultar...), Formato `CSV`.

## Controles antes de archivar

No archiva (y lo informa con `ok: false`) si:
- el CUIT del encabezado del reporte no es el de la sociedad pedida;
- las fechas del reporte o de algun comprobante no son del periodo;
- la suma de los comprobantes no da el **Total** del reporte al centavo (en las
  demas columnas se toleran $0,05 de redondeo: `control_total_reporte: "redondeo"`);
- hay IVA pero Gravado suma 0 (el CSV trae los importes de Grav una columna a la
  derecha de su titulo: si eso cambiara, se corta en vez de archivar mal);
- hay filas que no son comprobantes ni el total, o el reporte vino vacio.

Avisa sin cortar:
- `aviso_duplicados`: el mismo comprobante (TC, letra, numero, CUIT e importe)
  aparece mas de una vez en Zetti. Pasa con NC de Suizo y Monroe. Se deja tal
  cual viene (como en el papel), pero hay que mencionarlo.
- `aviso_pig`: P/R IG (percepciones de Ganancias) con importe; el papel no tiene
  esa columna (el Total la incluye).

## Que entrega

`[Sociedad]/01-Impuestos Mensuales/AAAAMM/IVA/AAAAMM - ZETTI - COMPRAS.xlsx` (la
misma carpeta que el Portal IVA). Hoja **"Compras Zetti"** igual a la solapa
"Compras Zetti" que arma el equipo en los papeles (el reporte abierto en Excel
sin las columnas vacias), para copiarla tal cual:

- Fila 1: A `T&S Web` | D razon social | F CUIT. Filas 2-3: D `Fecha des:` /
  `Fecha has:` con la fecha en E. Fila 4: `  Subdiario de IVA Compras`. Fila 5:
  `Razón Social:`.
- Fila 6: Fecha | TC | M | Nro. Comp. | Proveedor | CUIT | R | Exen | Grav | IVA |
  P/R IIBB | P/R IVA | Total.
- Datos desde la fila 7, en el orden del reporte (por fecha). Como al abrir el
  CSV en Excel: los CUIT sin guiones y los numeros sin guion quedan como numero,
  el resto como texto; importes de 1.000 o mas con `#,##0.00`, el resto General.
- Fila de total sin rotulo debajo de los datos (Exen a Total, el del reporte).
- Sin P/R IG (el papel no la tiene).

Validado contra los papeles 08/2026 de 14 sociedades: encabezado, columnas,
formatos y comprobantes iguales; las unicas diferencias son comprobantes que se
cargaron (sobre todo NC de Suizo y Monroe) o se corrigieron en Zetti despues de
que el equipo bajo el reporte, y la Ñ que en el papel quedo mal leida ("GRAÃ‘A").

## Respuesta al usuario

En 2-3 lineas por sociedad: sociedad, periodo, cantidad de comprobantes por TC y
letra, totales de Exen, Grav, IVA y Total (del JSON), y donde quedo el archivo.
Si hubo `redondeo`, `aviso_duplicados` o `aviso_pig`, mencionarlo en una linea.
Si una sociedad fallo, decir el motivo (`error` / `code`) en vez de adivinar:

- `sin_credenciales` / `clave_rechazada`: falta o esta mal la fila ZETTI de la planilla.
- `sin_entidad`: ese CUIT no esta en T&S Web (no usa Zetti o no esta habilitado para el usuario).
- `ya_existe`: ya estaba el archivo; preguntar si reemplazar.
- `control`: el reporte no paso un control; no se archivo nada.
- `sin_descarga` / `reporte`: T&S Web no respondio; reintentar mas tarde (si falla 2-3 veces, frenar y consultar).

Solo compras. Ventas: `zetti-subdiario-iva-ventas`.
