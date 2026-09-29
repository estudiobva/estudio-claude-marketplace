# portal-iva-descarga

Baja del **Portal IVA de ARCA** el Libro IVA **Compras** y **Ventas** de una o varias
sociedades para un periodo, los convierte a Excel y los archiva en la unidad
compartida:

```
[Sociedad]/01-Impuestos Mensuales/AAAAMM/IVA/AAAAMM - PORTAL IVA - COMPRAS.xlsx
[Sociedad]/01-Impuestos Mensuales/AAAAMM/IVA/AAAAMM - PORTAL IVA - VENTAS.xlsx
```

Por libro: importa desde ARCA al borrador, espera a que la importacion termine,
baja el CSV, lo pasa a Excel (fila TOTAL con formulas) y borra el CSV. Es el paso 1
del circuito `conciliacion-ncr`, para los dos libros con un solo login. Nunca
presenta la DDJJ.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
node scripts/portal-iva-descarga.js --nombre="<sociedad>"                    # mes anterior
node scripts/portal-iva-descarga.js --nombre="<sociedad>" --periodo=08/2026
node scripts/portal-iva-descarga.js --nombres="<A>;<B>" --periodo=08/2026 --si-existe=saltear
node scripts/portal-iva-descarga.js --cuit=<CUIT> --libros=ventas
```

Se trabaja a mes vencido: sin `--periodo` toma el mes anterior. `--ver` muestra el
navegador. `--no-archivar` deja los Excel en `~/Documents/BVA-salidas/portal-iva/`
(se cambia con `BVA_SALIDAS_PATH`).

## Reglas

- No entra si la DDJJ IVA del periodo ya tiene acuse en la carpeta del mes
  (`--aunque-presentada` para forzarlo).
- Corta si ARCA tiene abierto el borrador de otro periodo: descartarlo es decision
  de una persona.
- No pisa un Excel existente salvo `--si-existe=reemplazar`.
- Un libro que falla no frena al otro; un captcha o una clave rechazada cortan el lote.

## Estructura

- `scripts/portal-iva-descarga.js`: el flujo, una o varias sociedades.
- `scripts/csv_a_excel.py`: CSV de ARCA a Excel.
- `lib/portal-iva.js`: pantallas del Portal IVA (compartido con conciliacion-ncr).
- `lib/rutas-bva.js`: carpetas de la unidad compartida.
