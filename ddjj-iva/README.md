# ddjj-iva

Lee las DDJJ de IVA (F.2051) **ya presentadas** de una o varias sociedades para un rango
de meses, desde el Portal IVA de ARCA: Declaraciones juradas presentadas → "Declaración
Jurada de IVA" → última secuencia del período → vista previa.

Por período guarda en `~/Documents/BVA-salidas/ddjj-iva/<CUIT>/` (`$BVA_SALIDAS_PATH`):

- `AAAAMM.pdf`: la vista previa de la DDJJ.
- `AAAAMM.json`: secuencia, transacción, fecha de presentación y la determinación (débito
  y crédito fiscal, saldo técnico anterior, saldo de libre disponibilidad anterior,
  retenciones y percepciones, saldos del período).

Con `--archivar` copia el PDF a la carpeta IVA del mes como `AAAAMM-IVA-DDJJ-<TOKEN>.pdf`,
solo si ahí no hay ya una DDJJ de IVA. Solo lectura: nunca rectifica ni genera VEP.

## Instalacion

```bash
npm install
npx playwright install chromium
```

## Uso

```bash
node scripts/ddjj-iva.js --nombre="<sociedad>" --desde=202601 --hasta=202608
node scripts/ddjj-iva.js --cuit=<CUIT> --periodo=09/2026 --archivar
node scripts/ddjj-iva.js --nombres="<A>;<B>" --desde=202601 --hasta=202608
```

Detalle en `skills/ddjj-iva/SKILL.md`.
