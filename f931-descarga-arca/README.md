# f931-descarga-arca

Baja los **F.931 ya presentados** de una sociedad desde **Declaración en Línea de
ARCA** y entrega un PDF por período, nombrado `AAAAMM.pdf`. Se loguea solo con la
planilla de claves del estudio (el mismo login que `anticipos-sct`). Solo lee:
nunca genera ni presenta una DDJJ.

## Instalación

```bash
npm install
npx playwright install chromium
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
node scripts/f931.js --nombre="<sociedad>" --listar                                  # lo presentado
node scripts/f931.js --nombre="<sociedad>" --periodo=05/2026                         # un período
node scripts/f931.js --nombre="<sociedad>" --desde=04/2025 --hasta=03/2026 --meses=12 # un ejercicio
node scripts/f931.js --cuit=<cuit> --desde=202601 --hasta=202606 --original          # originales, no rectificativas
```

`--ver` muestra el navegador. `--meses=N` frena si el rango no tiene exactamente
N períodos.

Salida en `~/Documents/BVA-salidas/f931/<CUIT>/` (se cambia con `--salida` o
`BVA_SALIDAS_PATH`): un `AAAAMM.pdf` por período, el formulario real renderizado
en una hoja A4 apaisada.

## Reglas

- Login con la clave de la planilla (o `ARCA_<cuit>_LOGIN/PASSWORD` del `.env`).
  Si ARCA rechaza la clave o pide captcha, corta sin reintentar.
- Elige la sociedad en el desplegable de Declaración en Línea y corta si el CUIT
  no figura (falta la relación en ARCA).
- Ante una rectificativa baja la secuencia vigente (la más alta); `--original`
  baja la 000.
- Un período sin DDJJ presentada se informa en `faltantes`, no es un error.
- Verifica que cada PDF tenga una sola página.

## Estructura

- `scripts/f931.js`: una sociedad, uno o varios períodos.
- `lib/f931.js`: navegación de Declaración en Línea y render a PDF.
- `lib/arca-login.js` y `lib/env.js` / `lib/claves.js`: login en ARCA y credenciales.
- `skills/f931-descarga-arca/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche).
