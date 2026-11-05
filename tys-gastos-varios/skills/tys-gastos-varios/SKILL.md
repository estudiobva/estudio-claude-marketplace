---
name: tys-gastos-varios
description: "Baja de Zetti (T&S Web) el Reporte Seguimiento de Valores (4.5.2) del tipo GAVA - Gastos Varios de una, varias o todas las sociedades, lo parte por mes, lo cruza con el EERR (Matricula Colegio Farmaceuticos, Seguridad e Higiene, Publicidad, Municipalidad, Certificacion de Balances, Metrogas, Edesur, Agua, Alquileres) y guarda cada mes en la carpeta del mes de la sociedad como 'AAAAMM - ZETTI - GASTOS VARIOS.xlsx'. Corre un script de Playwright que se loguea solo con el usuario de Zetti de la planilla de claves del estudio. Usar SIEMPRE que pidan \"baja los gastos varios de [sociedad]\", \"seguimiento de valores GAVA\", \"reporte 4.5.2\", \"traeme de T&S los gastos de [sociedad] de [periodo]\", \"gastos varios de Zetti\" o completar en el EERR las filas de Seguridad e Higiene, Publicidad, Municipalidad, Matricula Colegio o Certificacion de Balances a partir de T&S. Para el Subdiario de IVA usar zetti-subdiario-iva-compras / -ventas."
---

# Gastos Varios (GAVA) desde T&S Web → resumen y cruce con el EERR

Un script de Playwright entra a T&S Web con el usuario del estudio, busca la
sociedad **por CUIT**, genera el **4.5.2 Reporte Seguimiento de Valores** con el
tipo de valor **56 GAVA - GASTOS VARIOS** en CSV, lo controla, lo parte por mes
(fecha de emision), lo cruza con el EERR y guarda **un archivo por mes en la
carpeta del mes de la sociedad**:

```
[Sociedad]/01-Impuestos Mensuales/[AAAA/]AAAAMM/AAAAMM - ZETTI - GASTOS VARIOS.xlsx
```

(dentro de la carpeta del anio si la sociedad la tiene; crea la del mes si falta).
Solo lee Zetti y el EERR: no cambia nada en ninguno de los dos.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --nombre="<sociedad>"     # ejercicio del EERR hasta el mes anterior
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --nombre="<sociedad>" --periodo=10/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --nombre="<sociedad>" --desde=01/01/2026 --hasta=30/09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --nombres="<A>;<B>" --periodo=10/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --todas --periodo=10/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/tys-gastos-varios.js --listar     # entidades de T&S Web con su CUIT
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un solo login para todas. |
| `--todas` | Todas las sociedades de Zetti que tienen carpeta en la unidad. |
| `--periodo` | Un mes: `MM/AAAA` o `AAAAMM`. |
| `--desde` / `--hasta` | Rango de emision `DD/MM/AAAA` (un archivo por cada mes con valores). |
| (sin periodo) | El ejercicio del EERR de la sociedad (el que contiene el mes anterior) **hasta el mes anterior**; sin EERR, el anio calendario hasta el mes anterior. |
| `--si-existe` | Si el mes ya tiene el archivo: `reemplazar` (default: el archivo lo genera solo este script y asi queda al dia, p.ej. INGRESADO que pasaron a PAGADO), `saltear` o `abortar`. |
| `--no-archivar` | No toca la unidad: deja los Excel en `--salida`. |
| `--entidad` | Nombre exacto de la entidad de T&S Web (p.ej. un solo local en vez de la Sociedad). |
| `--sin-cruce` | No busca ni lee el EERR. |
| `--salida` | Con `--no-archivar`: carpeta de salida (default `~/Documents/BVA-salidas/gastos-varios`). |
| `--ver` | Navegador visible. Solo si el usuario lo pide: por defecto va oculto. |

Si falta la sociedad, preguntar. Si falta el periodo, se usa el ejercicio hasta
el mes anterior: decirlo. Los meses sin gastos varios no generan archivo
(`meses_sin_valores`).

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Usa el mismo `~/.fisco-ar/.env` que los demas plugins (`CLAVES_ORGANISMOS_PATH`,
`BVA_UNIDAD_PATH` en Mac).

## Login

Igual que los demas plugins de Zetti: usuario y clave en `Claves_Organismos.xlsx`,
grupo de columnas **ZETTI**, fila **"Zetti T&S Web (todas las sociedades)"** (o
`ZETTI_USUARIO` / `ZETTI_CLAVE` en el `.env`). Claude nunca escribe ni muestra la
clave; si hay que cambiarla la carga la persona
(`python3 scripts/cargar-clave-zetti.py "<ruta a Claves_Organismos.xlsx>"`).

## Que hace en T&S Web (lo mismo que se hacia a mano)

1. Login y lectura de las entidades con su CUIT; usa la **Sociedad** ("SOC ..."),
   que abarca todos los locales (con `--entidad` se puede pedir un local).
2. Abre el **4.5.2** (`rpt_seguimiento_valores`) con el nodo de la sociedad y
   verifica que el Nodo del formulario sea el de esa entidad.
3. **Tipos Valor Asociados**: lupa "Buscar (c2)", filtro Nombre = `GAVA`, tilda
   **56 GAVA - GASTOS VARIOS**, "Seleccionar (c2)" y controla que quede en la grilla.
4. **Emision Desde / Hasta** con el periodo; el resto de los filtros vacios y sin
   tildes; Orden **Valor**; Formato **CSV**; Aceptar.

## Controles

Se baja **un solo reporte** para todo el rango y se controla entero antes de
partirlo por mes. Corta sin guardar nada (`ok: false`, `code: control`, deja el
CSV "REPORTE GAVA CON ERROR" en `~/Documents/BVA-salidas/gastos-varios`) si:
- el detalle de un concepto no suma su total, o los conceptos no suman el total general;
- algun valor tiene fecha de emision fuera del rango, o el reporte salio con otras fechas;
- hay filas que no reconoce (la codificacion de un valor puede tener letras, p.ej. "6547U65": eso se acepta).

T&S deja el **total general en blanco** cuando pasa los $ 1.000 M: en ese caso
se controla concepto por concepto y se avisa (`aviso_total`).

## Reglas (de la habilidad original)

- **IGNORADO**: fuera del total (suelen duplicar un PAGADO); en el Detalle tachado
  y con el PAGADO del mismo monto si lo hay.
- **INGRESADO**: adentro, marcado "pendiente de pago".
- Mes = mes de la **Fecha Em.**; se marca cuando el vencimiento cae en otro mes.

## Que entrega

Por cada mes con valores, `AAAAMM - ZETTI - GASTOS VARIOS.xlsx` en la carpeta del mes:
- **Resumen**: conceptos del mes con cantidad de valores, total (sin IGNORADO),
  INGRESADO e IGNORADO; en amarillo los que van al EERR y su fila destino.
- **Detalle**: Concepto, Codificacion, Tipo Valor, Monto, Fecha Em., Fecha Venc.,
  Nodo, Estado, Observacion (IGNORADO, pendiente de pago, vence en otro mes).
- **Cruce EERR** (si hay EERR): por fila del EERR, T&S del mes vs valor calculado
  del EERR y diferencia (en rojo > $1). Si T&S coincide con el EERR del mes
  anterior lo dice ("se paga el mes siguiente", tipico de Seguridad e Higiene).

Mapeo concepto T&S → fila EERR:

| Concepto en T&S | Fila EERR |
|---|---|
| COLEGIO ... | Matricula Colegio Farmaceuticos |
| CONSEJO PROFESIONAL CS ECONOMICAS | Certificacion Balances (no existe en el EERR: se informa) |
| GTOS IMPOSITIVOS - TASA PUBLICIDAD Y PROPAGANDA | Publicidad |
| GTOS IMPOSITIVOS - TASA SEGURIDAD E HIGIENE | Seguridad e Higiene |
| MUNICIPAL ... | Municipalidad / ABL (no existe en el EERR: se informa) |
| GTOS SUELDOS - ADEF - CS | *a confirmar* (Sueldos y Cargas Sociales): se muestra, no se compara |
| GV - JUICIOS LABORALES | *a confirmar* (Indemnizacion por Despido): se muestra, no se compara |
| METROGAS / EDESUR / AYSA / ALQUILER / CAFAR | fila homonima |

El resto queda como "no mapeado" en el Resumen.

Los valores del EERR se leen calculados (el ultimo guardado en Excel). Si el EERR
tiene formulas sin valor calculado lo avisa (`aviso_eerr`).

## Volcar al EERR

El script **no escribe** en el EERR. Si piden volcar: no pisar celdas con formula
(Matricula, Publicidad y servicios salen de 'Compras imputadas' por SUMIFS:
informar en vez de escribir), trabajar sobre una copia salvo que pidan modificar
el original y dejar comentario "Tomado de T&S 4.5.2 GAVA, emision DD/MM/AAAA–DD/MM/AAAA".

## Respuesta al usuario

Una linea por sociedad: periodo, total GAVA (sin IGNORADO), meses guardados (y
sin valores), conceptos mapeados con su total, diferencias principales contra
el EERR y cantidad/total de INGRESADO e IGNORADO a revisar. Si fallo, el motivo:

- `sin_credenciales` / `clave_rechazada`: falta o esta mal la fila ZETTI de la planilla.
- `sin_entidad`: ese CUIT no esta en T&S Web.
- `sin_carpeta`: la sociedad no tiene carpeta en la unidad.
- `ya_existe`: con `--si-existe=abortar`, el mes ya tenia el archivo.
- `sin_periodo`: el ejercicio todavia no tiene meses vencidos; pedir `--periodo`.
- `control`: el reporte no paso un control; mirar el CSV que quedo al lado.
- `sin_descarga` / `reporte`: T&S Web no respondio; reintentar (si falla 2-3 veces, frenar y consultar).
