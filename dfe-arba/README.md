# dfe-arba

Revisa el **Domicilio Fiscal Electrónico (DFE) de ARBA** de cada sociedad de la planilla de
claves (o de las que se pidan): entra con la CIT y lee la bandeja completa
(`SIIntimaciones/dfe/inicioDomicilioFiscalElectronico.do`).

- **Intimaciones y notificaciones**: se notifican formalmente; tienen fecha de puesta a
  disposición, de notificación y de lectura.
- **Comunicaciones**: avisos informativos (por ejemplo "Baja Retención A-122-R") que ARBA no
  notifica nunca.

**Solo lee el listado.** Nunca abre "Visualizar": abrir una comunicación le pone fecha de
leído y, si todavía no estaba notificada, puede adelantar la notificación.

Antes cruza el **Maestro de Sociedades** (`0 - ESTUDIO - 00000000000/Maestro de Sociedades.xlsx`,
hoja MAESTRO) con la planilla de claves: sociedades sin CIT de ARBA, filas que están en una
planilla y no en la otra, y un mismo CUIT con nombres distintos (las sucursales con el mismo
CUIT no se marcan).

Sale un Excel `DFE_ARBA_CONSOLIDADO_AAAA-MM-DD.xlsx` (Control, Comunicaciones, Maestro vs
Claves) en `~/Documents/BVA-salidas/dfe-arba/AAAA-MM-DD/` (`$BVA_SALIDAS_PATH`). Es reanudable
en el día.

## Instalacion

```text
/plugin install dfe-arba@estudiobva-skills
```

Dentro de la carpeta del plugin:

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

En el `.env` (o `~/.fisco-ar/.env`): `CLAVES_ORGANISMOS_PATH`, `CLAVES_ORGANISMOS_PASSWORD` y
`BVA_UNIDAD_PATH` (para el Maestro). Ver `.env.example`.

## Uso

```bash
node scripts/dfe-arba.js                     # todas las que tienen CIT de ARBA
node scripts/dfe-arba.js --cuit=<CUIT>       # una
node scripts/dfe-arba.js --solo=3,7,12       # posiciones de la planilla
node scripts/dfe-arba.js --consolidar        # rearma el Excel sin entrar a ARBA
```

Detalle en `skills/dfe-arba/SKILL.md`.
