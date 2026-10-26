# deudas-arba

El equivalente en ARBA de `deudas-ddjj-sct`: para cada sociedad de la planilla de claves (o
para las que se pidan) releva en ARBA:

- **Deuda por objeto** (Gestionar mis planes): IIBB, agente de retención y de percepción,
  partidas inmobiliarias; deuda prejudicial y judicial, saldo a favor y cantidad de DDJJ que
  ARBA da por no presentadas.
- **Agentes, período por período** (ARPresentaciones): DDJJ cerradas contra pagos registrados,
  para ver qué quincena o mes quedó con saldo. Se hace cuando ARBA marca deuda o DDJJ sin
  presentar (o para todos con `--detalle-agentes`).
- **DDJJ de anticipo de IIBB** (IBPresentaciones): los meses vencidos sin DDJJ, cuando ARBA
  marca alguna o no puede calcular la situación de IIBB.
- **Estado del agente** (consulta pública de ARBA): Activo / Inactivo / Suspendido. Se consulta
  también para las sociedades sin CIT.
- **Notificaciones DFE**: la cantidad que avisa ARBA.

Sale un Excel `Deudas_DDJJ_ARBA_CONSOLIDADO_AAAA-MM-DD.xlsx` (Control, Situación por objeto,
Agentes por período, DDJJ sin presentar) en `~/Documents/BVA-salidas/deudas-arba/AAAA-MM-DD/`
(`$BVA_SALIDAS_PATH`). Es reanudable en el día. Solo lee: nunca genera boletas, planes,
liquidaciones ni DDJJ.

AGIP no se automatiza: el Portal del Contribuyente bloquea el navegador automático.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

## Uso

```bash
node scripts/deudas-arba.js                     # todas las de la planilla
node scripts/deudas-arba.js --cuit=<CUIT>       # una
node scripts/deudas-arba.js --solo=3,7,12       # posiciones de la planilla
node scripts/deudas-arba.js --consolidar        # rearma el Excel sin entrar a ARBA
```

Detalle en `skills/deudas-arba/SKILL.md`.
