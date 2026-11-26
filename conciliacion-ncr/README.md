# conciliacion-ncr

Circuito de Notas de Credito de Recupero (NCR) de Suizo Argentina y Monroe Americana:
las baja del SharePoint de las farmacias, las concilia contra las compras del Portal IVA
(ARCA) y saca las conciliadas del borrador del Libro Compras. Nunca presenta la DDJJ.

## Instalación

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
# 1. El paquete de NCR del mes se baja solo del SharePoint (Playwright oculto):
node scripts/sharepoint-login.js --carpeta="NC por NRF/AAAA"                    # probar el ingreso (solo lectura)
node scripts/sharepoint/bajar-bundle-ncr.js --periodo=MM/AAAA                   # -> ~/Downloads/SP-NCR-AAAAMM.bundle
node scripts/descargar-ncr.js --nombre="<sociedad>" --periodo=MM/AAAA          # NCR a la carpeta ncr/ (baja el paquete si falta; --bajar lo renueva)
python3 scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=MM/AAAA --solo-revisar
python3 scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=MM/AAAA               # deja la conciliacion
node scripts/eliminar-ncr.js --nombre="<sociedad>" --periodo=MM/AAAA           # simulacion
node scripts/eliminar-ncr.js --nombre="<sociedad>" --periodo=MM/AAAA --ejecutar
node scripts/ncr-lote.js --periodo=MM/AAAA --nombres="A;B;C"                    # 1 + 2 + 3 + simulacion, varias
```

## Reglas

- Se elimina solo lo conciliado al centavo y al día. Las NC sin respaldo, las no
  encontradas y las diferencias se informan y no se tocan.
- `eliminar-ncr.js` sin `--ejecutar` solo simula. `ncr-lote.js` nunca elimina.
- El lote no entra a ARCA si la DDJJ de IVA del período ya está presentada.
- El lote reutiliza las compras que ya están en la carpeta si se bajaron después del día 5
  del mes siguiente; si se bajaron antes son parciales y no se usan (ver SKILL.md).
- Si ARCA o Microsoft (SharePoint) rechazan la clave o piden captcha, no reintenta.
- No pisa archivos: lo anterior va a `ncr/_anteriores/`.

## Estructura

- `scripts/sharepoint-login.js` + `lib/sharepoint-login.js`: ingreso al SharePoint por Playwright oculto
  (cuenta Microsoft del estudio, grupo SHAREPOINT de la planilla de claves; URL en `SHAREPOINT_SITIO`).
  Guarda la sesión en `~/.bva-playwright/sharepoint-state.json`. Si Microsoft pide verificación en dos
  pasos o captcha, corta sin reintentar.
- `scripts/sharepoint/bajar-bundle-ncr.js`: arma el paquete del mes desde el SharePoint (solo lectura).
- `scripts/sharepoint/bundle-ncr.js`: método anterior (pegado en la pestaña del SharePoint en Chrome);
  arma el mismo paquete.
- `scripts/descargar-ncr.js`: reparte las NCR a la carpeta `ncr/` de la sociedad.
- `scripts/conciliar-ncr.py`: la conciliación (solo openpyxl).
- `scripts/eliminar-ncr.js` + `scripts/leer_conciliacion_ncr.py`: eliminación en el Libro Compras.
- `scripts/ncr-lote.js`: varias sociedades hasta la simulación.
- `scripts/portal-iva-descarga.js` + `scripts/csv_a_excel.py`: compras del Portal IVA (paso 1),
  copia del plugin `portal-iva-descarga`.
- `lib/portal-iva.js`, `lib/rutas-bva.js`: Portal IVA y carpetas de la unidad compartida.
- `lib/arca-login.js`, `lib/env.js`, `lib/claves.js`: login en ARCA y credenciales.
- `lib/ncr-alias.json`: nombre de farmacia en Monroe -> `CLIENTE` en Suizo.
- `skills/conciliacion-ncr/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche).
