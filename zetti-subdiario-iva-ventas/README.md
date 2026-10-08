# zetti-subdiario-iva-ventas

Baja de **Zetti (T&S Web)** el **Subdiario de IVA Ventas** (reporte 5.6.5) de una,
varias o todas las sociedades para un mes, lo pasa a Excel limpio y ordenado por
tipo de comprobante y lo archiva en la unidad compartida, al lado del Libro IVA
Ventas del Portal IVA:

```
[Sociedad]/01-Impuestos Mensuales/AAAAMM/IVA/AAAAMM - ZETTI - VENTAS.xlsx
```

El login es automático: usuario y clave de Zetti en `Claves_Organismos.xlsx`
(grupo de columnas ZETTI, fila "Zetti T&S Web (todas las sociedades)"). La
sociedad se busca por CUIT en T&S Web y se usa la entidad Sociedad, que abarca
todos sus locales. Solo lee: no cambia nada en Zetti.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # o usar ~/.fisco-ar/.env, el mismo de los demas plugins
```

## Uso

```bash
node scripts/zetti-iva-ventas.js --nombre="<sociedad>"                    # mes anterior
node scripts/zetti-iva-ventas.js --nombre="<sociedad>" --periodo=09/2026
node scripts/zetti-iva-ventas.js --todas --periodo=09/2026 --si-existe=saltear
node scripts/zetti-iva-ventas.js --listar
```

Detalle de argumentos, controles y salida en `skills/zetti-subdiario-iva-ventas/SKILL.md`.
