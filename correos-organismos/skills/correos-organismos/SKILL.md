---
name: correos-organismos
description: "Releva el correo electronico declarado de todas las sociedades (o de las que se pidan) en ARCA (Sistema Registral: correo principal, verificacion, telefono, adhesion al DFE, lista de correos; Domicilio Fiscal Electronico: a que correo llegan los avisos) y en ARBA (Datos de contacto y correo de la CIT), y arma un Excel con lo declarado, lo que hay que revisar (correos de terceros, sin verificar, dominios sin correo) y lo que falta. Login automatico con la planilla de claves, solo lectura. Usar SIEMPRE que pidan \"relevar los correos declarados\", \"a que mail le llegan las notificaciones de ARCA/ARBA\", \"domicilio fiscal electronico de las sociedades\", \"que correo tiene [sociedad] en ARBA\", \"hay que cambiar el mail en ARCA\" o \"control de correos de los organismos\"."
---

# Correos declarados en ARCA y ARBA

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/correos-organismos.js
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/correos-organismos.js --nombres="<A>;<B>"
```

| Argumento | Para que |
|---|---|
| `--nombres="A;B"` / `--solo=3,7,12` | Filtran por nombre o por posicion en la planilla de claves. Sin filtro: todas. |
| `--organismos` | `arca`, `arba` o `arca,arba` (default). |
| `--rehacer` | Vuelve a relevar lo que ya salio bien hoy (por defecto se saltea: es reanudable). |
| `--consolidar` | No entra a ningun sitio: rearma el Excel con lo relevado en el dia. |
| `--dominio-estudio` | Dominio propio del estudio (default `estudiobva.com`): los demas se marcan como externos. |
| `--salida` | Carpeta base (default `~/Documents/BVA-salidas/correos`), una subcarpeta por dia. |

## Que hace

- **ARCA**: un login por apoderado (no por sociedad: loguear muchas veces hace rebotar y un
  rechazo repetido bloquea la clave de todo el grupo). Por sociedad lee el Sistema
  Registral y, una vez por login, el DFE > "Asociar temas a un correo electronico".
- **ARBA**: un intento por sociedad con su CIT, sin reintentos (ARBA bloquea la CIT tras
  varios intentos fallidos).
- **AGIP**: no se automatiza; las sociedades con clave de AGIP quedan para relevar a mano.

Excel `CORREOS_ORGANISMOS_AAAA-MM-DD.xlsx`: Resumen por sociedad, Detalle de cada correo,
Correos unicos (en cuantas sociedades aparece cada uno) y Control.

## Que se puede saber y que no

Se ve si hay correo, si el organismo lo da por verificado, si coincide entre organismos,
si es del estudio o de terceros y si el dominio tiene servidor de correo (DNS). **No** se
puede saber si un buzon de terceros sigue leyendose: esos quedan marcados para confirmar.

## Respuesta al usuario

Cuantas se relevaron, cuales hay que revisar (correo de terceros, sin verificar, sin
correo, distinto entre ARCA y ARBA), las CIT rechazadas y las que quedan para AGIP a mano.
Si piden guardar el Excel en la unidad, copiarlo a `0 - ESTUDIO - 00000000000/`.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
