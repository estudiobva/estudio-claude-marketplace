# cruze-zetti-portaliva-ventas

Circuito completo de ventas del mes de una o varias sociedades, en un solo comando:

1. **Portal IVA**: usa `AAAAMM - PORTAL IVA - VENTAS.xlsx` de la carpeta IVA si se bajó
   después del día 5 del mes siguiente; si no, lo baja (`portal-iva-descarga.js
   --libros=ventas`, o el libro presentado si la DDJJ ya está presentada).
2. **Zetti**: igual con `AAAAMM - ZETTI - VENTAS.xlsx` (`zetti-iva-ventas.js`).
3. **Papel de trabajo** del mes: escribe las hojas `Ventas Zetti`, `Ventas Portal IVA`,
   `cruze` y `NC controlador (fuera)`. Si el papel existe lo respalda y solo toca esas
   hojas; si no, crea `AAAAMM-papeldetrabajo-<SOCIEDAD>.xlsx`.
4. **Cruze**: `cruzar_ventas.py` (facturas/NC una a una, tiques por día y punto de venta
   contra el cierre Z).

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

## Uso

```bash
node scripts/cruze-zetti-portaliva-ventas.js --nombre="<sociedad>" --periodo=09/2026
node scripts/cruze-zetti-portaliva-ventas.js --nombres="<A>;<B>" --periodo=09/2026 --si-existe=reemplazar
```

Detalle en `skills/cruze-zetti-portaliva-ventas/SKILL.md`.
