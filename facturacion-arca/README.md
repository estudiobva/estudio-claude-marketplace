# facturacion-arca

Emite facturas de servicios en **Comprobantes en Línea (RCEL) de ARCA** para
cualquier emisor: Monotributo (Factura C) o Responsable Inscripto (Factura A a
RI/monotributistas, B al resto). Los datos salen de la planilla mensual
(`assets/Plantilla Facturacion ARCA.xlsx`: hojas Emisor, Clientes y una por mes).

> ⚠️ **Escribe en ARCA.** Cada factura sale con CAE y sólo se anula con nota de crédito.

## Instalación

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # o ~/.fisco-ar/.env: ruta y contraseña de la planilla de claves
```

El emisor tiene que estar en la planilla de claves o en el `.env`
(`ARCA_<cuit>_LOGIN` / `ARCA_<cuit>_PASSWORD`). El login es el mismo de anticipos-sct.

## Uso

```bash
node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="Octubre 2026" --validar          # sin ARCA: controles, tabla, total y código
node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="Octubre 2026"                    # dry-run: hasta la confirmación, no confirma
node scripts/facturacion-arca.js --planilla="<ruta.xlsx>" --hoja="Octubre 2026" --emitir=<código>  # emite lo aprobado
```

`--filas=5,7` limita a algunas filas; `--ver` muestra el navegador. Una sola
factura sin planilla: ver la cabecera de `scripts/facturacion-arca.js`.

Salida en `~/Documents/BVA-salidas/facturas/` (o `BVA_SALIDAS_PATH`):

- PDF `AAAAMM - [Emisor] - [Cliente].pdf` en la carpeta del emisor (hoja Emisor) o en `facturas/<EMISOR>/`.
- Capturas del dry-run en `facturas/<EMISOR>/dry-run/`, dumps de cada paso en `facturas/_debug/`.
- `.registro-<cuit>.json`: lo emitido, para no reemitir nunca una fila.

## Reglas

- El código de `--emitir` es un hash de lo pendiente: si la planilla cambia después de aprobar, no emite.
- Elige la empresa del emisor en RCEL por CUIT/nombre y la verifica en el encabezado.
- Espera la razón social del padrón para el CUIT del receptor; avisa si no coincide con la planilla.
- Antes de confirmar controla en pantalla CUIT del receptor, total (y neto/IVA en la A), fecha y tipo.
- Después de confirmar busca el comprobante nuevo en RCEL → Consultas. Si no lo encuentra, corta el lote
  y deja la fila en `verificar`: hay que mirar a mano antes de seguir.
- Marca la fila (`N° comprobante`, `¿Emitida?` = Sí). Si la planilla está abierta, lo avisa; el registro igual la protege.

## Estado

Probado sin ARCA (controles, código, registro, marcado de la planilla) el 29/09/2026.
**Falta el primer dry-run contra RCEL**: algunos ids de campos (fecha del comprobante,
alícuota, condición de venta) pueden necesitar ajuste.

## Estructura

- `scripts/facturacion-arca.js`: controles, asistente de RCEL, emisión, PDF y registro.
- `scripts/facturacion_planilla.py`: lee la planilla y marca las filas emitidas.
- `lib/rcel.js`: navegación de RCEL.
- `lib/arca-login.js` y `lib/env.js` / `lib/claves.js`: login en ARCA y credenciales.
- `skills/facturacion-arca/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche); `lib/rcel.js`
parte del de fisco-ar y lo amplía el estudio.
