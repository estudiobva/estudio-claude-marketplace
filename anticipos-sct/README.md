# anticipos-sct

Releva los anticipos de **Ganancias y Bienes Personales** de cada sociedad en el
**Sistema de Cuentas Tributarias (SCT) de ARCA**, solapa Deudas, y los vuelca a
Excel: número de anticipo, capital, intereses, vencimiento y estado. Solo lee:
nunca modifica, presenta ni paga nada.

## Instalación

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
node scripts/anticipos-sct.js --nombre="<sociedad>"    # una sociedad
node scripts/anticipos-sct.js --cuit=<cuit> --crudo    # guarda además todas las filas de Deudas
node scripts/anticipos-sct-lote.js                     # todas las sociedades de la planilla de claves
node scripts/anticipos-sct-lote.js --solo=7,12         # posiciones de la planilla
node scripts/anticipos-sct-lote.js --consolidar        # rearma el consolidado sin entrar a ARCA
```

`--ver` muestra el navegador. El lote es reanudable: lo ya relevado hoy se saltea
(`--rehacer` lo vuelve a relevar).

Salida en `~/Documents/BVA-salidas/anticipos/` (se cambia con `--salida` o `BVA_SALIDAS_PATH`):

- `Anticipos_SCT_<SOCIEDAD>_AAAA-MM-DD.xlsx`: Sociedad, Impuesto, Nro. Anticipo,
  Capital, Int. Resarcitorio, Int. Punitorio, Total, Fecha de Vencimiento, Estado.
- Lote: una subcarpeta por día con un Excel por sociedad y
  `Anticipos_SCT_CONSOLIDADO_AAAA-MM-DD.xlsx` (hojas Anticipos y Control).

## Reglas

- El lote loguea una vez por apoderado. Si ARCA rechaza la clave o pide captcha, no reintenta.
- Verifica el contribuyente activo en el SCT antes de leer cada sociedad.
- Corta si leyó menos filas de las que declara la solapa Deudas.
- Un anticipo de un impuesto que no sea Ganancias o Bienes Personales no se descarta:
  sale marcado en rojo como novedad.
- Una sociedad sin anticipos igual genera su Excel, con la fila "Sin anticipos".

## Estructura

- `scripts/anticipos-sct.js`: una sociedad.
- `scripts/anticipos-sct-lote.js`: todas, con consolidado.
- `scripts/anticipos_a_excel.py`: los Excel.
- `lib/sct.js`: navegación y lectura del SCT.
- `lib/arca-login.js` y `lib/env.js` / `lib/claves.js`: login en ARCA y credenciales.
- `skills/anticipos-sct/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche).
