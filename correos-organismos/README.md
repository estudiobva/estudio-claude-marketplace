# correos-organismos

Releva el correo electrónico declarado de todas las sociedades de la planilla de claves (o
de las que se pidan) para saber si sigue activo, si hay que cambiarlo o si no hay nada:

- **ARCA, Sistema Registral**: correo principal y si está verificado, teléfono, adhesión al
  Domicilio Fiscal Electrónico y la lista completa de correos.
- **ARCA, DFE** > "Asociar temas a un correo electrónico": a qué correo llegan los avisos.
- **ARBA**, Datos de contacto: correos con carácter y fecha de verificación, teléfonos y el
  correo con que se registró la CIT.
- **AGIP**: no se automatiza (el firewall de AGIP bloquea el navegador automático); queda la
  columna para completar a mano.

Sale un Excel `CORREOS_ORGANISMOS_AAAA-MM-DD.xlsx` (Resumen, Detalle, Correos únicos,
Control) en `~/Documents/BVA-salidas/correos/AAAA-MM-DD/` (`$BVA_SALIDAS_PATH`). Es
reanudable en el día. Solo lee: nunca modifica correos, teléfonos ni adhesiones.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

## Uso

```bash
node scripts/correos-organismos.js                          # todas
node scripts/correos-organismos.js --nombres="<A>;<B>"
node scripts/correos-organismos.js --organismos=arba
node scripts/correos-organismos.js --consolidar             # rearma el Excel sin entrar
```

Detalle en `skills/correos-organismos/SKILL.md`.
