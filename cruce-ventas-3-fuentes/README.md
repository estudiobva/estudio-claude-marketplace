# cruce-ventas-3-fuentes

Ventas del mes de una o varias sociedades desde tres fuentes, y el cruce entre ellas:

1. **Mis Comprobantes** (ARCA), emitidos: `mis-comprobantes-emitidos.js` →
   `AAAAMM - MIS COMPROBANTES - EMITIDOS.xlsx` (el CSV de ARCA completo).
2. **Portal IVA**, Libro IVA Ventas: `portal-iva-descarga.js --libros=ventas` (o el libro
   presentado si la DDJJ ya está presentada) → `AAAAMM - PORTAL IVA - VENTAS.xlsx`.
3. **Zetti**, Subdiario de IVA Ventas 5.6.5: `zetti-iva-ventas.js` → `AAAAMM - ZETTI - VENTAS.xlsx`.
4. **Cruce**: `cruzar_ventas_3.py` → `AAAAMM - CRUCE VENTAS MC PORTAL ZETTI.xlsx`
   (Resumen / Diferencias / Detalle / NC controlador).

Todo queda en la carpeta IVA del mes. Cada fuente se reusa si ya está en la carpeta y se
bajó a partir del día 6 del mes siguiente. Si una fuente no se puede bajar, cruza las otras
dos y lo avisa. No toca el papel de trabajo (eso es `cruze-zetti-portaliva-ventas`).
Solo lectura en ARCA y en Zetti.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

## Uso

```bash
node scripts/cruce-ventas-3-fuentes.js --nombre="<sociedad>" --periodo=09/2026
node scripts/cruce-ventas-3-fuentes.js --nombres="<A>;<B>" --periodo=09/2026 --rebajar
node scripts/mis-comprobantes-emitidos.js --nombre="<sociedad>" --periodo=09/2026
```

Detalle en `skills/cruce-ventas-3-fuentes/SKILL.md`.
