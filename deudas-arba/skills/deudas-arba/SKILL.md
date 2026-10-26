---
name: deudas-arba
description: "Controla en ARBA, para todas las sociedades (o las que se pidan), la deuda por objeto (Ingresos Brutos, agente de retencion y de percepcion, partidas inmobiliarias), las DDJJ que ARBA da por no presentadas, lo declarado contra lo pagado como agente de recaudacion periodo por periodo, los meses sin DDJJ de anticipo de IIBB, el estado oficial del agente (activo/suspendido) y las notificaciones DFE, en un Excel consolidado. Es el equivalente en ARBA de deudas-ddjj-sct (ARCA). Login automatico con la CIT de la planilla de claves, solo lectura. Usar SIEMPRE que pidan \"deudas de ARBA\", \"deuda de agentes ARBA\", \"agentes de recaudacion con saldo\", \"DDJJ pendientes en ARBA\", \"DDJJ no presentadas de IIBB\", \"el mismo control de ARCA pero para ARBA\", \"esta suspendido el agente\" o \"que debe [sociedad] en ARBA\". No confundir con constancias-anticipo-arba (baja las constancias de DDJJ ya presentadas) ni con deudas-ddjj-sct (ARCA)."
---

# Deudas y DDJJ sin presentar en ARBA

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-arba.js
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-arba.js --cuit=<CUIT>
```

| Argumento | Para que |
|---|---|
| `--cuit=<CUIT>` / `--solo=3,7,12` | Una sociedad, o posiciones de la planilla de claves. Sin filtro: todas. |
| `--anios=5` | Años hacia atras para el detalle por periodo (default 5). |
| `--detalle-agentes` | Cruza declarado vs pagado de TODOS los agentes. Por defecto solo cuando ARBA marca deuda, DDJJ sin presentar o no pudo calcular la situacion. |
| `--rehacer` | Vuelve a relevar lo que ya salio bien hoy (por defecto se saltea: es reanudable). |
| `--consolidar` | No entra a ARBA: rearma el Excel con lo relevado en el dia. |
| `--salida` | Carpeta base (default `~/Documents/BVA-salidas/deudas-arba`), una subcarpeta por dia. |

Correr siempre con el navegador oculto (sin `--ver`) salvo que el usuario pida verlo. El
lote completo tarda ~15 minutos; correrlo en segundo plano.

## Que hace

- **Un login por sociedad** con su CIT. Si ARBA la rechaza **no se reintenta**: se guarda una
  huella de la clave en `cit-rechazadas.json` y mientras la planilla tenga la misma clave no se
  vuelve a probar (cada intento fallido suma para el bloqueo de la CIT, que el equipo usa a mano).
- **Gestionar mis planes**: deuda prejudicial/judicial, saldo a favor y DDJJ sin presentar por
  objeto, tal como las calcula ARBA. Sellos no se calcula al entrar (queda "No consultado").
- **ARPresentaciones** (agentes): DDJJ cerradas contra pagos registrados por año; un periodo con
  saldo > $ 1 es deuda. La quincena o el mes que todavia no vencio sale como "A vencer" (el
  vencimiento es aproximado). Para un agente cesado con DDJJ sin presentar se informa el
  periodo siguiente a la ultima DDJJ (casi siempre el del cese).
- **IBPresentaciones** (IIBB): meses vencidos sin DDJJ de anticipo. Si no hay ninguna DDJJ en
  todo el rango, se informa como "probablemente Convenio Multilateral", no como N omisiones.
- **Estado del agente**: consulta publica de ARBA (sin clave), tambien para las sociedades sin CIT.

Excel `Deudas_DDJJ_ARBA_CONSOLIDADO_AAAA-MM-DD.xlsx`: Control (una fila por sociedad),
Situacion por objeto, Agentes por periodo y DDJJ sin presentar.

## Como leer los resultados

- **"El Agente se encuentra suspendido" en A-122R no es un dato**: la pantalla lo muestra
  siempre que puede leer los datos del agente. El estado real es la columna "Estado agente".
- **"IMPUESTO INVALIDO" / "No se pudo calcular la situacion del objeto"** en IIBB: ARBA no da la
  deuda de ese CUIT. Con vinculo CONVENIO es Convenio Multilateral (ver SIFERE); en el resto se
  revisan igual las DDJJ de anticipo.
- **"CUIT con estado de Cta. en proceso de actualizacion"**: pasa despues de rectificar; volver
  a correr esa sociedad al dia siguiente (`--cuit=... --rehacer`).
- **Inmobiliario**: puede incluir cuotas del año en curso, y una partida con cotitulares aparece
  en cada uno.
- La deuda de agente de Gestionar incluye la quincena/mes que vence hoy o despues: el Control lo
  aclara en Observaciones.

## Respuesta al usuario

Cuantas se relevaron, cuales tienen deuda (separando lo vencido de lo que todavia no vencio y
el inmobiliario), DDJJ sin presentar con el periodo, agentes suspendidos, CIT rechazadas (hay
que actualizarlas en Claves_Organismos), las que no tienen CIT y las que hay que volver a
correr. AGIP no se automatiza: el Portal del Contribuyente bloquea el navegador automatico.
Si piden guardar el Excel en la unidad, copiarlo a `0 - ESTUDIO - 00000000000/`.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
