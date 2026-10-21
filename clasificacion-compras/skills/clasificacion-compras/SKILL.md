---
name: clasificacion-compras
description: "Clasifica las compras del mes de una o varias sociedades con el Maestro de Proveedores del estudio: a cada comprobante del Libro Compras del Portal IVA (ya sin las NCR de Suizo y Monroe si esta la conciliacion) le pone la cuenta contable de su proveedor y el rubro del EERR, y deja en la carpeta IVA del mes 'AAAAMM - COMPRAS CLASIFICADAS.xlsx' con el detalle, el resumen por rubro, el total por proveedor y los proveedores que faltan en el Maestro o estan sin confirmar. Usar SIEMPRE que pidan \"clasificar las compras de [sociedad] [mes]\", \"clasificacion de compras\", \"imputar las compras\", \"compras por cuenta / por rubro\", \"cuanto es CMV y cuanto gastos\", \"que proveedores no estan en el Maestro\" o \"compras para el EERR\". Va despues de portal-iva-descarga y de conciliacion-ncr. No confundir con padron-proveedores (da de alta proveedores en el padron) ni con conciliacion-ncr (saca las NCR)."
---

# Clasificacion de compras con el Maestro de Proveedores

Toma el Libro Compras del mes y lo clasifica proveedor por proveedor con el
**Maestro de Proveedores** (`0 - ESTUDIO - 00000000000/Maestro de Proveedores.xlsx`,
hoja MAESTRO). Es el mismo criterio con que se armaron los EERR 2026: comparado contra la
hoja "Compras imputadas" de los EERR da igual al peso, rubro por rubro.

Solo lee: no entra a ARCA y no modifica el Maestro.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/clasificar-compras.py --nombre="<sociedad>" --periodo=08/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/clasificar-compras.py --nombres="<A>;<B>" --periodo=08/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/clasificar-compras.py --cuit=<CUIT> --desde=01/2026 --hasta=08/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/clasificar-compras.py --todas --periodo=08/2026 --solo-revisar
```

En Mac es `python3` en vez de `python`.

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` (con `;`), `--todas` | Sociedad(es). El nombre es el de la carpeta, sin numero ni CUIT; si coincide con varias, el script corta y las lista. |
| `--periodo` o `--desde` + `--hasta` | `MM/AAAA` o `AAAAMM`. Con rango deja un archivo por mes. |
| `--fuente` | `auto` (default), `sin-ncr` (exige la conciliacion NCR) o `portal` (el libro tal cual). |
| `--compras=<ruta>` | Usar ese archivo de compras (una sociedad, un mes). |
| `--si-existe` | `error` (default), `reemplazar` (el anterior va a `_anteriores/`) o `saltear`. |
| `--solo-revisar` | Clasifica y muestra el resultado sin escribir nada. |
| `--maestro=<ruta>` | Otro Maestro (default el de la unidad, o `BVA_MAESTRO_PROVEEDORES` del `.env`). |

Antes de escribir, conviene correr con `--solo-revisar` y mostrarle al usuario el
resultado: comprobantes, CMV, gastos, excluido y proveedores sin Maestro.

## De donde salen las compras

Con `--fuente=auto`, por sociedad y mes:

1. **Conciliacion NCR** (`[AAAAMM]-COMPRAS-SIN-NCR-*.xlsx`, hoja "Compras S.NCR", en la
   carpeta `ncr/` del mes): es el libro ya depurado de las Notas de Credito de Recupero de
   Suizo y Monroe. Se prefiere siempre.
2. Si no hay conciliacion, el **libro del Portal IVA** (`[AAAAMM] - PORTAL IVA - COMPRAS.xlsx`,
   o el CSV/ZIP de ARCA) en la carpeta IVA del mes. El resultado lleva el aviso de que
   las NCR no estan depuradas: el CMV incluye esas notas de credito. Decirselo al usuario.

Si falta el libro, primero hay que bajarlo con `portal-iva-descarga` (`--libros=compras`), y
para depurar las NCR, correr `conciliacion-ncr`. El orden del mes es:
**portal-iva-descarga → conciliacion-ncr → clasificacion-compras**.

## Reglas de clasificacion

- **El Maestro manda.** La cuenta es la columna CUENTA CONTABLE del proveedor (por CUIT).
- Estado `Propuesta (...)`: se usa la cuenta propuesta y el proveedor sale en pendientes
  para confirmarlo.
- Estado `Revisar`: va a GASTOS VARIOS y sale en pendientes.
- Estado `Bloqueado` (CUIT limitada RG 3832/16): **EXCLUIDO - PROVEEDOR BLOQUEADO**, fuera del
  resultado.
- Cuenta `COMPRA DE EQUIPOS - ...`: **EXCLUIDO - BIEN DE USO**, fuera del resultado.
- CUIT que no esta en el Maestro: `SIN MAESTRO`, va a GASTOS VARIOS y sale en pendientes
  resaltado en amarillo. **No inventar la cuenta**: se da de alta en el Maestro (skill
  `padron-proveedores`, o se le pasa la lista a Joaquin) y despues se vuelve a correr con
  `--si-existe=reemplazar`.
- **Importe imputado** (lo que va al EERR): con IVA discriminado, no gravado + exento +
  netos gravados; sin IVA discriminado (Factura B/C, tiques), el total menos percepciones
  e impuestos. Las notas de credito vienen en negativo del Portal IVA y restan solas.
- Cada cuenta del Maestro se lleva a su fila del EERR (`RUBRO` en el script; p. ej.
  LUZ → Edesur, GAS → Metrogas, PAPELERA → IMPRENTA Y PAPELERIA). Una cuenta nueva que no
  este en `RUBRO` cae en GASTOS VARIOS: avisar para agregarla.

## Salida

`[AAAAMM] - COMPRAS CLASIFICADAS.xlsx` en la carpeta IVA del mes (la misma donde
`portal-iva-descarga` deja el libro), con estas hojas:

| Hoja | Que tiene |
|---|---|
| Compras clasificadas | El libro completo, mismas columnas y orden del Portal IVA, mas Cuenta contable, Rubro EERR, Estado Maestro e Importe imputado. |
| Resumen por rubro | Rubro EERR × cuenta: comprobantes, importe imputado, IVA y total. Fila TOTAL. |
| Por proveedor | Un renglon por CUIT con su cuenta, rubro, estado e importes. |
| Pendientes Maestro | Proveedores sin Maestro (amarillo) o sin clasificacion firme, con que hacer. |
| Control | Totales, CMV / gastos / excluido, verificacion (tiene que dar 0), fuente, fecha del Maestro y avisos. |

El JSON por stdout trae lo mismo resumido (`cmv`, `gastos`, `excluido`, `sin_maestro`,
`pendientes_confirmar`, `avisos`). Con `--todas`, las carpetas sin libro del mes
(personas, sociedades sin IVA) van a `sin_libro_del_mes` y no cuentan como error.

## Reportar al usuario

Por sociedad y mes: fuente usada (con o sin NCR depuradas), comprobantes, CMV, gastos,
excluido, y la lista de proveedores **sin Maestro** con su importe. Si hubo avisos,
nombrarlos. No hace falta listar los pendientes de confirmar uno por uno: estan en la hoja.
