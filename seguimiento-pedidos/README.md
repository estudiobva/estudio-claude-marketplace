# seguimiento-pedidos

Planilla compartida de pedidos **Estudio ↔ Cliente** y tareas internas, alimentada por una
skill que lee el mail y el Slack de quien la corre.

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

El script `scripts/pedidos.py` solo lee y escribe la planilla; la lectura de mails y Slack y
la decisión de qué es un pedido la hace Claude siguiendo `skills/seguimiento-pedidos/SKILL.md`.

## Requisitos

- Python 3 con `openpyxl` (`pip install -r requirements.txt`).
- Conectores de **Gmail** y **Slack** activos en Claude.
- La planilla en la unidad compartida (o `PEDIDOS_PLANILLA` en el `.env`). Si no existe:
  `python scripts/pedidos.py crear`, y completar las hojas **Equipo** (Nombre, Email,
  Slack ID), **Listas** (columna Cliente) y **Config** (dominios del cliente, link, canales).

## Uso

En Claude: "actualizá la planilla de pedidos y mandame mis pendientes".

```bash
python scripts/pedidos.py estado --email=<mail> --slack=<ID>
python scripts/pedidos.py aplicar ops.json --solo-revisar
python scripts/pedidos.py aplicar ops.json
python scripts/pedidos.py pendientes --persona="<Nombre>" --formato=slack
```

## Cuidados

- La planilla es `.xlsx`: no convertirla a Hoja de cálculo de Google (el script escribe el
  archivo). Si está abierta en Excel en la misma compu, el script avisa y no guarda.
- Antes de guardar deja un respaldo en `BVA-salidas/seguimiento-pedidos/respaldos/` y no
  escribe si la planilla cambió mientras procesaba.
- Los pendientes se mandan solo a quien corre la skill.
