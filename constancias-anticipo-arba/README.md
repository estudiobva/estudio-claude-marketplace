# constancias-anticipo-arba

Baja de ARBA las constancias de las DDJJ de anticipo de Ingresos Brutos **ya presentadas**
de una o varias sociedades (IBPresentaciones → Reimpresión → Constancia anticipo), el PDF
original de "Imprimir resumen", y las archiva en la unidad compartida:

```
[Sociedad]/01-Impuestos Mensuales/[AAAAMM]/IIBB/AAAAMM-IIBB-DJ-<TOKEN>.pdf
```

El `<TOKEN>` es el que ya usan los archivos de esa sociedad (no se deduce del nombre de la
carpeta). También deja el PDF y un JSON con el detalle (gravados, no gravados, deducciones,
impuesto determinado, monto imponible) en `~/Documents/BVA-salidas/arba-ddjj/<CUIT>/`, que
usa el generador de papeles de trabajo. Nunca pisa: si la DJ del mes ya está, la deja.

Login con la CIT de cada sociedad (`Claves_Organismos.xlsx`), un intento por sociedad y sin
reintentos (ARBA bloquea la CIT tras varios intentos fallidos). Solo lectura.

## Instalacion

```bash
npm install
npx playwright install chromium
```

## Uso

```bash
node scripts/constancias-anticipo-arba.js --nombre="<sociedad>" --desde=202601 --hasta=202608
node scripts/constancias-anticipo-arba.js --nombres="<A>;<B>" --periodo=09/2026
```

Detalle en `skills/constancias-anticipo-arba/SKILL.md`.
