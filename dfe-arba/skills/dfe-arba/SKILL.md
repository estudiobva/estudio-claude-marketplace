---
name: dfe-arba
description: "Revisa el Domicilio Fiscal Electronico (DFE) de ARBA de todas las sociedades (o las que se pidan): entra con la CIT de la planilla de claves y lista todas las intimaciones, notificaciones y comunicaciones de la bandeja, con fecha de puesta a disposicion, notificacion y lectura, impuesto, operativo y expediente, sin abrir ninguna. Ademas cruza el Maestro de Sociedades con la planilla Claves_Organismos (sin CIT, filas faltantes, CUIT repetidos). Excel consolidado, solo lectura, Playwright con login automatico. Usar SIEMPRE que pidan \"revisar el domicilio fiscal electronico de ARBA\", \"DFE de ARBA\", \"notificaciones de ARBA\", \"intimaciones de ARBA\", \"que le llego a [sociedad] en ARBA\", \"bandeja de ARBA\" o \"comunicaciones sin leer en ARBA\". No confundir con correos-organismos (el correo declarado), deudas-arba (deuda y DDJJ sin presentar; solo cuenta los avisos DFE) ni con arca-ventanilla (ARCA)."
---

# Domicilio Fiscal Electronico de ARBA

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/dfe-arba.js
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/dfe-arba.js --cuit=<CUIT>
```

| Argumento | Para que |
|---|---|
| `--cuit=<CUIT>` / `--solo=3,7,12` | Una sociedad, o posiciones de la planilla de claves. Sin filtro: todas las que tienen CIT de ARBA. |
| `--dias=30` | Ventana de "recibidas en los ultimos N dias" (default 30). |
| `--rehacer` | Vuelve a relevar lo que ya salio bien hoy (por defecto se saltea: es reanudable). |
| `--consolidar` | No entra a ARBA: rearma el Excel con lo relevado en el dia. |
| `--salida` | Carpeta base (default `~/Documents/BVA-salidas/dfe-arba`), una subcarpeta por dia. |

Correr siempre con el navegador oculto (sin `--ver`) salvo que el usuario pida verlo. El lote
completo (~60 sociedades) tarda ~15 minutos: correrlo en segundo plano.

## Que hace

- **Un login por sociedad** con su CIT. Si ARBA la rechaza **no se reintenta** y se guarda una
  huella de la clave en `cit-rechazadas.json`; tampoco se prueban las CIT que `deudas-arba` ya
  registro como rechazadas con la misma clave (cada intento fallido suma para el bloqueo).
- Las sociedades que ademas son **agentes de recaudacion** pasan por "Seleccione un rol" despues
  de la clave: no es un rechazo. Se elige Contribuyente (la bandeja es la misma con los dos roles).
- Lee **todas** las filas de cada tabla de la bandeja (no solo la pagina visible), con el estado
  leida/notificada. **Nunca abre "Visualizar".**
- Cruza el Maestro de Sociedades con la planilla de claves.

Excel `DFE_ARBA_CONSOLIDADO_AAAA-MM-DD.xlsx`:
- **Control**: una fila por sociedad (Maestro + planilla de claves) con el estado del
  relevamiento, total en bandeja, sin leer, intimaciones/notificaciones sin leer y sin notificar
  aun, recibidas en los ultimos N dias y la ultima comunicacion. En rojo lo reciente y los errores.
- **Comunicaciones**: todas, de la mas nueva a la mas vieja, en rojo las de los ultimos N dias.
- **Maestro vs Claves**: diferencias entre las dos planillas.

## Como leer los resultados

- **Comunicaciones** (tipo) son avisos informativos ("Baja Retencion A-122-R", "Contribuyentes
  con gastos superiores a los ingresos", "ARBA Movil"): nunca se notifican y no son pendientes.
  Lo que importa son las **Intimaciones** y **Notificaciones**: multas automaticas, resoluciones
  sancionatorias, requerimientos de fiscalizacion remota (FIRE), DDJJ adeudadas, recupero de deuda.
- "Sin notificar aun" en una intimacion o notificacion = ARBA la puso a disposicion y todavia no
  corrio el plazo: abrirla la daria por notificada en ese momento. Avisar antes de que alguien la abra.
- Los titulos vienen de ARBA con letras acentuadas faltantes ("Retenci n"): es asi en el sitio.
- Bandeja vacia: ARBA muestra solo la bienvenida; sale como "Bandeja vacia", no como error.

## Respuesta al usuario

Cuantas se revisaron; las intimaciones/notificaciones recientes o sin leer por sociedad, con
titulo y fecha de notificacion (agrupar las que se repiten, como los FIRE mensuales); las que
todavia no estan notificadas; CIT rechazadas (actualizar Claves_Organismos), las que no tienen
CIT y las diferencias Maestro vs Claves. Si piden guardar el Excel en la unidad, copiarlo a
`0 - ESTUDIO - 00000000000/`.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
