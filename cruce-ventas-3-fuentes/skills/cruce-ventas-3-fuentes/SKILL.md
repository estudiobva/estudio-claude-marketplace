---
name: cruce-ventas-3-fuentes
description: "Ventas del mes de una o varias sociedades desde tres fuentes y su cruce: baja Mis Comprobantes emitidos (ARCA), el Libro IVA Ventas del Portal IVA y el Subdiario de IVA Ventas de Zetti (5.6.5), los archiva en la carpeta IVA del mes y arma 'AAAAMM - CRUCE VENTAS MC PORTAL ZETTI.xlsx' con las diferencias explicadas (falta en una fuente, importe distinto, corrimiento de fecha de tiques, corte de mes). Login automatico con la planilla de claves y navegador oculto. Usar SIEMPRE que pidan \"cruzar Mis Comprobantes, Portal IVA y Zetti\", \"cruce de ventas de las tres fuentes\", \"por que no cierran las ventas de [sociedad]\", \"bajar Mis Comprobantes emitidos\" o \"comprobantes emitidos de [mes] en Excel\". Para escribir las hojas del papel de trabajo usar cruze-zetti-portaliva-ventas (este no toca el papel)."
---

# Cruce de ventas: Mis Comprobantes + Portal IVA + Zetti

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/cruce-ventas-3-fuentes.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/cruce-ventas-3-fuentes.js --nombres="<A>;<B>" --periodo=09/2026
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Default: el mes anterior. |
| `--rebajar` | Baja las tres fuentes aunque ya esten en la carpeta. |
| `--si-existe` | `reemplazar` (default) o `abortar`: el archivo del cruce. |
| `--tolerancia` | Diferencia aceptada por componente (default 0.5). |

Solo Mis Comprobantes:

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/mis-comprobantes-emitidos.js --nombre="<sociedad>" --periodo=09/2026
```

(`--si-existe=abortar|saltear|reemplazar`, `--no-archivar` deja el Excel en
`~/Documents/BVA-salidas/mis-comprobantes`.)

## Que hace, por sociedad

1. **Mis Comprobantes** emitidos del mes → `AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx`,
   hoja "Mis Comprobantes Emitidos": el CSV de ARCA completo, mismas columnas y mismo
   orden. Ojo: Mis Comprobantes trae las NC en **positivo** (el Portal las trae negativas)
   y, desde 01/01/2026, cada tique del controlador fiscal de nueva tecnologia.
2. **Portal IVA** → `AAAAMM - PORTAL IVA - VENTAS.xlsx`. Si la DDJJ ya esta presentada,
   baja el libro presentado (`--presentada`).
3. **Zetti** → `AAAAMM - ZETTI - VENTAS.xlsx`.
4. **Cruce** → `AAAAMM - CRUCE VENTAS MC PORTAL ZETTI.xlsx`, hojas Resumen, Diferencias,
   Detalle y NC controlador (fuera).

Cada fuente se reusa si ya esta en la carpeta IVA y se bajo a partir del dia 6 del mes
siguiente; si no, se baja. Si una no se puede bajar (por ejemplo, la sociedad no usa
Zetti) cruza las otras dos y lo avisa. Si ARCA pide captcha o rechaza la clave, corta.

## Como se cruza

- Facturas, NC y ND: por (tipo ARCA, punto de venta, numero). El TC + letra de Zetti
  (FV A, NC B...) se traduce al codigo de ARCA.
- Tiques: sumados por (fecha, punto de venta) contra el Z de Zetti. Si dos dias se
  compensan, es **corrimiento de fecha** (el Z quedo con la fecha del dia siguiente).
- IVA del Portal = suma de "Importe IVA x%" (el "Total IVA" viene redondeado por ARCA).
- **Corte de mes** (solo Zetti contra ARCA): el numero que falta queda fuera del rango
  que Zetti tiene en el mes para ese tipo y PV → Zetti lo fecho en el mes de al lado.
- Las NC B del controlador fiscal viejo que Zetti lista pero ya estan dentro del Z van a
  "NC controlador (fuera)".

## Respuesta al usuario

Por sociedad: total de cada fuente, donde quedo el cruce y las diferencias agrupadas por
motivo (Falta en X / Solo en X / Importe distinto / Corrimiento de fecha / Corte de mes).
Una diferencia redonda (1.000.000, 1.800.000) suele ser un error de carga en Zetti.
Si una fuente se bajo antes de que termine el mes, avisar que puede estar incompleta.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
