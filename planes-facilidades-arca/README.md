# planes-facilidades-arca

Releva los planes de facilidades de pago de cada sociedad en **Mis Facilidades (ARCA)**
y arma un reporte: cuotas pagadas, impagas y a vencer, CBU declarado y obligaciones
incluidas en cada plan. Solo lee: nunca presenta, reformula ni paga nada.

## Instalación

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # completar la ruta y la contraseña de la planilla de claves
```

## Uso

```bash
node scripts/planes-facilidades.js                    # todas las sociedades de la planilla de claves
node scripts/planes-facilidades.js --login=<apoderado> # solo las que entran con ese apoderado
node scripts/planes-facilidades.js --nombre="<sociedad>"
node scripts/planes-facilidades.js --historico         # incluye planes cancelados/caducos
node scripts/planes-facilidades.js --consolidar        # rearma Excel y PDF sin entrar a ARCA
node scripts/planes-facilidades.js --ver               # navegador visible
```

Salida en `descargas/planes/AAAA-MM-DD/` (excluida del repo):

- `REPORTE_PLANES_FACILIDADES_AAAAMMDD.pdf`: primero las sociedades que no se pudieron
  revisar, después las que tienen cuotas impagas y al final todas, plan por plan.
- `INFORME_PLANES_FACILIDADES_AAAAMMDD.xlsx`: hojas INFORME, CUOTAS IMPAGAS, PLANES
  VIGENTES, OBLIGACIONES, DETALLE CUOTAS y CONTROL.
- Un Excel por sociedad.

## Reglas

- Loguea una vez por apoderado. Si ARCA rechaza la clave o pide captcha, no reintenta.
- Verifica el CUIT del encabezado antes de leer cada sociedad.
- Si una pantalla no carga, la cuota queda "sin dato": nunca se informa como impaga.
- Controla que la suma de cuotas pagadas cierre con el "Total Pagado" de ARCA.

## Estructura

- `scripts/planes-facilidades.js`: el lote (login, recorrido y salida).
- `scripts/planes-reporte.js`: el PDF.
- `scripts/planes_a_excel.py`: los Excel.
- `lib/facilidades.js`: lectura de las pantallas de Mis Facilidades.
- `lib/arca-login.js` y `lib/env.js` / `lib/claves.js`: login en ARCA y credenciales.
- `skills/planes-facilidades-arca/SKILL.md`: la habilidad.

`lib/arca-login.js`, `args.js`, `launch.js`, `env.js`, `claves.js`, `claves-dump.js`
y `xlsx-min.js` vienen del plugin fisco-ar (MIT, Javier Gradiche).
