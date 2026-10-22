# recibo-pyp-quilmes

Baja de **ARQui** (Municipio de Quilmes) el **recibo actual de Derechos de
Publicidad y Propaganda** (marquesinas) de una, varias o todas las sociedades,
controla que el CUIT y la cuenta del recibo sean los de la sociedad y lo archiva
donde el equipo guarda las boletas de marquesina:

```
[Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-boleta marquesina- <Sociedad>.pdf
```

El login es automático: usuario y clave del grupo MUNICIPAL (Quilmes) de
`Claves_Organismos.xlsx`. El número de cuenta sale del bloque PUBLICIDAD Y
PROPAGANDA de `Tasas Municipales CLAVES.xlsx`, por CUIT. Solo lee y descarga: no
paga ni vincula cuentas.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
cp .env.example .env   # o usar ~/.fisco-ar/.env, el mismo de los demas plugins
```

## Uso

```bash
node scripts/recibo-pyp-quilmes.js --nombre="<sociedad>"
node scripts/recibo-pyp-quilmes.js --nombres="<A>;<B>"
node scripts/recibo-pyp-quilmes.js --todas            # + hoja "Control planilla"
node scripts/recibo-pyp-quilmes.js --listar           # cuentas de ARQui vs planilla de tasas
```

Detalle de argumentos, controles, salida y los bloqueos (403) de ARQui en
`skills/recibo-pyp-quilmes/SKILL.md`.
