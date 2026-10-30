---
name: seguimiento-pedidos
description: "Lee los mails con el cliente y el Slack de quien la corre, carga en la Hoja de cálculo de Google compartida 'Seguimiento de pedidos Estudio-Cliente' (0 - ESTUDIO) lo que el estudio le pidió al cliente, lo que el cliente le pidió al estudio y las tareas internas que salieron de Slack, actualiza estados y reclamos sin duplicar, y le manda a la persona sus pendientes por mensaje privado de Slack. Usar SIEMPRE que pidan \"actualizá la planilla de pedidos\", \"qué tengo pendiente\", \"mis pendientes\", \"cargá los pedidos del cliente\", \"seguimiento de pedidos\", \"qué le pedimos al cliente\", \"qué me pidieron por Slack\" o \"mandame mis pendientes\". Necesita los conectores de Gmail y Slack activos."
---

# Seguimiento de pedidos Estudio ↔ Cliente

Cada persona la corre con **su** mail y **su** Slack. Lo que encuentra va a la planilla
común; al final le llega a esa persona un mensaje privado de Slack con lo que tiene abierto.

Planilla: Hoja de cálculo de Google **Seguimiento de pedidos Estudio-Cliente**
(`0 - ESTUDIO - 00000000000`). Hojas: **Estudio → Cliente** (EC), **Cliente → Estudio** (CE),
**Tareas internas** (TI), más Resumen, Equipo, Config y Corridas.

La planilla tiene un Apps Script publicado como app web (`apps-script/Codigo.gs`). El script
de la skill le habla con `PEDIDOS_URL` y `PEDIDOS_TOKEN` del `.env`: no abre ni descarga
ningún archivo. Gmail y Slack se leen con los conectores de Claude.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/pedidos.py probar
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/pedidos.py estado --email=<mail> --slack=<ID>
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/pedidos.py aplicar <ops.json> --solo-revisar
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/pedidos.py aplicar <ops.json>
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/pedidos.py pendientes --persona="<Nombre>" --formato=slack
```

En Mac es `python3`. Si `probar` dice que faltan `PEDIDOS_URL`/`PEDIDOS_TOKEN`, pedírselos a
Joaquín (están en el README del plugin cómo se generan) y agregarlos a `~/.fisco-ar/.env`.
**No** pegar el token en el chat ni en Slack.

## Paso 1 — Quién corre y desde cuándo

1. Mail propio: Gmail `search_threads` con `in:sent` y `pageSize: 1` → el `sender`.
   Slack ID propio: lo dice la descripción de las herramientas de Slack
   ("Current logged in user's user_id is …").
2. `pedidos.py estado --email=<mail> --slack=<ID>`. Devuelve `persona`, `desde`/`hasta`
   (desde la última corrida de esa persona; la primera vez, los días de Config),
   `config.dominios_cliente`, `config.canales_slack`, `equipo` (nombres, mails, IDs),
   `cliente` (personas del cliente), `claves_cargadas` y `abiertos`.
3. Si dice que no sabe quién sos, la persona tiene que agregarse en la hoja **Equipo**
   (Nombre, Email, Slack ID). No inventar el nombre.

## Paso 2 — Mails con el cliente

Por cada dominio de `config.dominios_cliente`:

```
search_threads  query: after:AAAA/MM/DD (from:<dominio> OR to:<dominio> OR cc:<dominio>)
```

(`desde` en formato `AAAA/MM/DD`, `pageSize: 50`, paginar). Leer cada hilo con
`get_thread` y `messageFormat: PLAIN_TEXT`.

- **Mail propio al cliente pidiendo algo** → EC `agregar`. Un pedido con varios puntos
  numerados = **una fila por punto** (clave `gmail:<threadId>#<n>`). `a_quien` = la persona
  del cliente a la que va ese punto (si el mail lo asigna) o el primer destinatario.
- **El cliente contesta con lo pedido** → `actualizar` la fila abierta (está en `abiertos`):
  `estado` Completado o Respondido parcial, `fecha_recibido`, y en `notas` qué mandó.
  `clave_origen`: `gmail:<threadId>/<messageId>`.
- **Se vuelve a pedir lo mismo** (mail propio de seguimiento en el hilo o uno nuevo) →
  `reclamo` sobre esa fila, con `fecha`, `detalle` y `clave_origen` del mensaje. **No**
  agregar una fila nueva.
- **El cliente pide algo al estudio** → CE `agregar` (`pedido_por` = quien escribe,
  `responsable` = quien tiene que hacerlo; si no está claro, quien corre la skill).
  Cuando se entrega (mail propio con lo pedido) → `actualizar` con `fecha_entrega` y Completado.

## Paso 3 — Slack

Usar `slack_search_public_and_private` con `include_context: false` (en el formato
detallado viene `Message_ts` y el ID del canal). Slack toma `after:` sin incluir ese día:
usar el día anterior a `desde`.

1. Privados y grupos: `filters: "to:<@ID> after:<desde-1>"`.
2. Menciones: `keywords: ["<@ID>"]`, `filters: "after:<desde-1>"`.
3. Lo que escribió la persona (para ver qué prometió o qué ya cerró):
   `filters: "from:<@ID> after:<desde-1>"`.
4. Canales de `config.canales_slack`: `slack_search_channels` para el ID y
   `slack_read_channel` con `oldest` = `desde` en segundos. Ojo: devuelve las páginas desde
   lo más viejo; seguir el cursor hasta el final para llegar a lo de hoy.

Qué cargar:

- **Alguien le pide algo a la persona** ("¿podés…?", "agregale…", "revisá…") → TI
  `agregar` con `responsable` = la persona, `pedido_por` = quien lo pidió, `origen`
  "Slack #canal" o "Slack DM <nombre>", `link` = permalink.
- **Tareas que se reparten en un canal** (resúmenes de reunión, "X: hacer tal cosa") →
  una fila TI por tarea con su responsable, aunque no sea quien corre la skill
  (clave `slack:<canal>/<ts>#<n>`).
- **Pedidos del cliente que se comentan en Slack** ("me lo pidió <alguien del cliente>", "el cliente quiere…")
  → CE con `canal` según cómo llegó, `responsable` = a quien se lo asignan.
- **Ya está hecho** (respondió, "listo", "enviado", PR mergeado) → `estado` Completado y
  `fecha_cierre`/`fecha_entrega`. Si no se puede saber, `A confirmar` con la duda en notas.
- No cargar charla, saludos, avisos de reuniones ni preguntas que se respondieron en el
  mismo momento sin dejar nada por hacer.

## Paso 4 — Cargar en la planilla

Armar `ops.json` en una carpeta temporal:

```json
{"persona": "<Nombre>", "desde": "AAAA-MM-DD", "hasta": "AAAA-MM-DD",
 "ops": [
  {"accion": "agregar", "hoja": "EC", "clave_origen": "gmail:<threadId>#1",
   "datos": {"fecha_pedido": "AAAA-MM-DD", "a_quien": "...", "sociedad": "...", "tema": "EERR",
             "que": "...", "canal": "Mail", "estado": "Pendiente", "link": "<viewUrl>", "notas": "..."}},
  {"accion": "agregar", "hoja": "CE", "clave_origen": "slack:<canal>/<ts>",
   "datos": {"fecha_pedido": "...", "pedido_por": "...", "canal": "Slack", "responsable": "...", "que": "..."}},
  {"accion": "agregar", "hoja": "TI", "clave_origen": "slack:<canal>/<ts>",
   "datos": {"origen": "Slack #general", "tarea": "...", "responsable": "...", "pedido_por": "...",
             "fecha_pedido": "...", "fecha_maxima": "...", "link": "<permalink>"}},
  {"accion": "actualizar", "id": "E-003", "clave_origen": "gmail:<threadId>/<messageId>",
   "datos": {"estado": "Completado", "fecha_recibido": "AAAA-MM-DD", "notas": "..."}},
  {"accion": "reclamo", "id": "E-003", "fecha": "AAAA-MM-DD", "detalle": "Reclamado por mail",
   "clave_origen": "gmail:<threadId>/<messageId>"}
 ]}
```

Reglas:

- **`clave_origen` siempre** (`gmail:<threadId>#n`, `gmail:<threadId>/<messageId>`,
  `slack:<canalID>/<ts>` o `slack:<canalID>/<ts>#n`). Con eso el script no carga dos veces
  lo mismo, aunque lo corran dos personas que vieron el mismo mensaje. No agregar lo que ya
  está en `claves_cargadas`.
- Nombres de responsables y de pedido_por: **los de `equipo`** (para el estudio) y **los de
  `cliente`** (para el cliente), escritos igual. Si aparece alguien del cliente que no está,
  usar su nombre y avisar para agregarlo a Listas.
- Estados: Pendiente, En curso, Respondido parcial, Completado, Cancelado, A confirmar.
  Temas y canales: los de `temas` y `canales`.
- Fechas `AAAA-MM-DD`. `fecha_maxima` solo si alguien la dijo; no inventarla.
- `que` / `tarea`: una frase que se entienda sola. **No copiar el mail entero ni datos
  sensibles** (claves, CBU, montos de terceros que no hagan falta).
- `link`: el `viewUrl` del hilo de Gmail o el permalink de Slack.

Correr primero con `--solo-revisar` y mostrarle a la persona la lista de lo que se va a
agregar, actualizar y reclamar (ID, hoja, qué, responsable, estado). Con su OK, correr sin
`--solo-revisar`. La planilla aplica los cambios de a una corrida por vez (si dos personas
la corren juntas, la segunda espera) y queda todo en el historial de versiones de Google.

## Paso 5 — Mandar los pendientes

`pedidos.py pendientes --persona="<Nombre>" --formato=slack` arma el mensaje (vencidos
primero, después lo que vence en 2 días, agrupado en: le pediste al cliente / te pidió el
cliente / tareas internas, con link a la planilla).

Mostrarle el mensaje a la persona y mandarlo con `slack_send_message` a **su propio** Slack
ID como `channel_id` (le llega como mensaje privado propio). **Nunca** mandárselo a otra
persona: si alguien pide mandarle los pendientes a otro, mostrarle el texto para que lo
reenvíe o pedir confirmación explícita para ese envío puntual.

## Reportar

Cuántos agregados, actualizados, reclamos y omitidos (ya cargados), los errores si hubo,
el link de la planilla y el mensaje de pendientes que se mandó. Si quedaron filas en
"A confirmar", nombrarlas: son las que la persona tiene que revisar a mano.
