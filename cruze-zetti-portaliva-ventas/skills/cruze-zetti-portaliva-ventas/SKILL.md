---
name: cruze-zetti-portaliva-ventas
description: "Circuito completo de ventas del mes de una o varias sociedades: baja el Libro IVA Ventas del Portal IVA (ARCA), baja el Subdiario de IVA Ventas de Zetti (T&S Web, 5.6.5), vuelca los dos en las hojas 'Ventas Portal IVA' y 'Ventas Zetti' del papel de trabajo del mes y arma la hoja 'cruze' con las diferencias. Todo con login automatico (planilla de claves) y navegador oculto. Usar SIEMPRE que pidan \"cruzar las ventas de [sociedad] [mes]\", \"cruze Zetti vs Portal\", \"armar las ventas del papel de trabajo\", \"bajar Portal y Zetti y cruzar\", \"control de ventas del mes\". Para bajar solo Zetti usar zetti-subdiario-iva-ventas; solo Portal, portal-iva-descarga."
---

# Cruze ventas Zetti vs Portal IVA (circuito completo)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/cruze-zetti-portaliva-ventas.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/cruze-zetti-portaliva-ventas.js --nombres="<A>;<B>" --periodo=09/2026
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. |
| `--si-existe` | `abortar` (default) o `reemplazar`: si las hojas del papel ya tienen datos. |
| `--rebajar` | Vuelve a bajar Portal IVA y Zetti aunque ya esten en la carpeta. |
| `--papel=<ruta>` | Papel a usar si en la carpeta del mes hay mas de uno. |

## Que hace, por sociedad

1. **Portal IVA**: usa `AAAAMM - PORTAL IVA - VENTAS.xlsx` de la carpeta IVA si se bajo
   despues del dia 5 del mes siguiente; si no, lo baja con `portal-iva-descarga.js
   --libros=ventas` (importa al borrador). Si la DDJJ ya esta presentada, baja el libro
   presentado (`--presentada`).
2. **Zetti**: igual con `AAAAMM - ZETTI - VENTAS.xlsx` y `zetti-iva-ventas.js`.
3. **Papel de trabajo** del mes, en la carpeta `AAAAMM` (o `AAAA/AAAAMM`): el `.xlsx`
   que diga "papel" (no los de agentes ni los "NO VA"). Si no hay, crea
   `AAAAMM-papeldetrabajo-<SOCIEDAD>.xlsx`. Si el del equipo es `.xls`, crea el `.xlsx`
   aparte y lo avisa. Escribe, con el formato compacto del papel de trabajo mensual:
   - **Ventas Zetti**: A1 titulo, fila 2 titulos, datos desde la 3 (PV y numero en D y E).
   - **Ventas Portal IVA**: el CSV del Libro IVA Ventas de ARCA pasado a Excel, sin perder
     nada: todas las columnas en el mismo orden (comprador, alicuotas, percepciones, etc.),
     encabezado en la fila 1 y una fila por comprobante, con fechas y numeros como tales.
     Verificado celda por celda contra el CSV crudo. (Sin la fila TOTAL ni la nota que trae
     el archivo de la carpeta IVA.)
   - **cruze**: diferencias Zetti vs Portal (cruzar_ventas.py: facturas/NC una a una, tiques
     por dia y punto de venta contra la Z). Las NC B a Consumidor Final del controlador
     fiscal (PV chicos) van a **NC controlador (fuera)**: no estan en el Portal porque van
     dentro del Z.
   Si el papel existe se respalda en `~/Documents/BVA-salidas/_respaldos/papeles/` y solo
   se reescriben esas hojas (el resto, por ejemplo Liquidacion IVA o COMPRAS, queda igual).
   Si alguna ya tiene datos, corta salvo `--si-existe=reemplazar`.

## Respuesta al usuario

Por sociedad: donde quedo el papel (nuevo o actualizado), total Zetti, total Portal,
diferencia, NC del controlador separadas y las filas del cruze con su patron (ver la
skill cruce-ventas-zetti-portal-iva para interpretar: corrimiento de fecha, solo en un
lado, diferencia de importe, redondeo). Si conviene, escribir la Observacion de las
filas en la hoja cruze.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Usa el mismo `~/.fisco-ar/.env` que los demas plugins (`CLAVES_ORGANISMOS_PATH` a la
planilla de la unidad; en Mac, `BVA_UNIDAD_PATH`). Credenciales: ARCA por sociedad y la
fila "Zetti T&S Web (todas las sociedades)" de Claves_Organismos.xlsx.

Validado 08/10/2026: hoja Ventas Zetti identica celda por celda a un papel armado a mano y
Ventas Portal IVA identica al CSV crudo de ARCA (08/2026); en 09/2026 la diferencia total quedo explicada por las NC
del controlador y un corrimiento de fecha de un cierre Z.
