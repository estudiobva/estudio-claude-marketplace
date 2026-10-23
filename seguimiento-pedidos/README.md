# seguimiento-pedidos

Hoja de cálculo de Google compartida con los pedidos **Estudio ↔ Cliente** y las tareas
internas, alimentada por una skill que lee el mail y el Slack de quien la corre.

## Qué hace

1. Mira quién la corre (mail de Gmail y Slack ID → hoja **Equipo** de la planilla) y desde
   cuándo leer (su última corrida, hoja **Corridas**).
2. Lee con los conectores de Claude:
   - Gmail: los hilos con los dominios del cliente (hoja **Config**).
   - Slack: mensajes privados y de grupo, menciones, lo que escribió la persona y los
     canales de Config.
3. Carga en la planilla, sin duplicar (columna oculta **Clave origen**):
   - **Estudio → Cliente**: lo que pedimos, con reclamos y fecha de recibido.
   - **Cliente → Estudio**: lo que pide el cliente, con responsable y fecha de entrega.
   - **Tareas internas**: lo que se pide o reparte por Slack.
4. Le manda a la persona sus pendientes por mensaje privado de Slack (a sí misma).

## Cómo está armado

- `apps-script/Codigo.gs` + `apps-script/appsscript.json`: el Apps Script que vive dentro de
  la Hoja de cálculo de Google. Arma las hojas, y como **app web** recibe los pedidos del
  script (estado, aplicar, pendientes, importar). Escribe con bloqueo: si dos personas corren
  la skill al mismo tiempo, la segunda espera.
- `scripts/pedidos.py`: cliente en Python (solo biblioteca estándar) que le habla a la app
  web con `PEDIDOS_URL` y `PEDIDOS_TOKEN`. No descarga ni abre la planilla.
- `skills/seguimiento-pedidos/SKILL.md`: qué lee Claude en Gmail y Slack y qué cuenta como
  pedido.

## Instalación de la planilla (una sola vez, la hace Joaquín)

1. Abrir la Hoja de cálculo de Google **Seguimiento de pedidos Estudio-Cliente** en
   `0 - ESTUDIO - 00000000000` (o crear una vacía con ese nombre).
2. En la planilla: **Extensiones › Apps Script**. Borrar lo que haya en `Código.gs` y pegar
   `apps-script/Codigo.gs`. En **Configuración del proyecto** tildar "Mostrar el archivo de
   manifiesto appsscript.json" y pegar `apps-script/appsscript.json`. Guardar.
3. Elegir la función `instalar` y **Ejecutar**. Google pide autorizar el acceso a la planilla
   (solo a esta planilla). Al terminar, el **Registro de ejecución** muestra el token.
4. **Implementar › Nueva implementación › App web**: ejecutar como **Yo**, acceso
   **Cualquier persona**. Copiar la URL (termina en `/exec`).
5. En cada computadora, en `~/.fisco-ar/.env`:
   ```
   PEDIDOS_URL=<la URL /exec>
   PEDIDOS_TOKEN=<el token>
   ```
   El token es como una clave: va solo en el `.env`, nunca en chats, Slack ni el repo. Si se
   filtra, correr `regenerarToken` en el editor y actualizar los `.env`.
6. Completar las hojas **Equipo** (Nombre, Email, Slack ID), **Listas** (columna Cliente) y
   **Config** (dominios del cliente, canales de Slack). O migrar la planilla Excel anterior:
   `python scripts/pedidos.py importar --xlsx="<ruta de la planilla .xlsx>"` (necesita
   `openpyxl`; solo funciona con la Google Sheet vacía).
7. Verificar: `python scripts/pedidos.py probar`.

Si se cambia `Codigo.gs`, pegar la versión nueva y en **Implementar › Administrar
implementaciones** editar la implementación con **Versión nueva** (la URL no cambia).

## Uso

En Claude: "actualizá la planilla de pedidos y mandame mis pendientes".

```bash
python scripts/pedidos.py probar
python scripts/pedidos.py estado --email=<mail> --slack=<ID>
python scripts/pedidos.py aplicar ops.json --solo-revisar
python scripts/pedidos.py aplicar ops.json
python scripts/pedidos.py pendientes --persona="<Nombre>" --formato=slack
```

## Cuidados

- No cambiar los encabezados de la fila 2 ni borrar la columna oculta **Clave origen**: la app
  web busca las columnas por nombre y usa la clave para no duplicar.
- Cada cambio queda en el historial de versiones de la planilla (Archivo › Historial de versiones).
- Los pendientes se mandan solo a quien corre la skill.
