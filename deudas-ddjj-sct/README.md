# deudas-ddjj-sct

Controla en el **Sistema de Cuentas Tributarias (SCT) de ARCA** de cada sociedad:

- las **deudas que no son anticipos** (solapa Deudas: saldos de DDJJ, multas,
  intereses, retenciones), y
- las **DDJJ pendientes de presentación** (solapa del mismo nombre).

Los anticipos quedan afuera a propósito: los releva el plugin `anticipos-sct`, con
el mismo criterio, así que entre los dos se cubre toda la solapa Deudas sin
duplicar. Solo lee: nunca modifica, presenta ni paga nada.

## Instalación

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
node scripts/deudas-ddjj-sct-lote.js                  # todas las sociedades de la planilla de claves
node scripts/deudas-ddjj-sct-lote.js --solo=7,12      # posiciones de la planilla
node scripts/deudas-ddjj-sct-lote.js --consolidar     # rearma el consolidado sin entrar a ARCA
node scripts/deudas-ddjj-sct.js --nombre="<sociedad>" # una sociedad
```

`--ver` muestra el navegador. El lote es reanudable: lo ya relevado hoy se saltea
(`--rehacer` lo vuelve a relevar).

Salida en `~/Documents/BVA-salidas/deudas-ddjj/AAAA-MM-DD/` (se cambia con `--salida`
o `BVA_SALIDAS_PATH`): **un Excel consolidado**,
`Deudas_DDJJ_SCT_CONSOLIDADO_AAAA-MM-DD.xlsx`, con las hojas

- **Control**: una fila por sociedad (Sin observaciones / Con deudas / DDJJ pendientes /
  Deudas y DDJJ pendientes / ERROR / Omitida).
- **Deudas**: Sociedad, CUIT, Impuesto, Concepto, Subconcepto, Período, Capital,
  Int. Resarcitorio, Int. Punitorio, Total, Vencimiento, Estado, Observación.
- **DDJJ pendientes**: Sociedad, CUIT, Impuesto, Concepto, Subconcepto, Período,
  Vencimiento, Días de atraso (el SCT no informa importes).

`--por-sociedad` agrega un Excel suelto por sociedad; el script de una sociedad siempre
genera el suyo.

## Reglas

- El lote loguea una vez por apoderado. Si ARCA rechaza la clave o pide captcha, no reintenta.
- Verifica el contribuyente activo en el SCT antes de leer cada sociedad.
- Corta si leyó menos filas de las que declara la solapa.
- Si la solapa de pendientes no se abre y queda Vencimientos a la vista, lo detecta y no
  informa "sin pendientes" por error.

## Estructura

- `scripts/deudas-ddjj-sct-lote.js`: todas, con consolidado.
- `scripts/deudas-ddjj-sct.js`: una sociedad.
- `scripts/deudas_ddjj_a_excel.py`: los Excel.
- `lib/sct.js`: navegación y lectura del SCT (compartida con `anticipos-sct`).
- `lib/arca-login.js` y `lib/env.js` / `lib/claves.js`: login en ARCA y credenciales.
- `skills/deudas-ddjj-sct/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche).
