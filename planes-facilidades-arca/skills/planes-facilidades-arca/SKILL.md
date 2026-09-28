---
name: "planes-facilidades-arca"
description: "Releva los planes de facilidades de pago vigentes de una sociedad en Mis Facilidades de ARCA (cuotas pagadas, pendientes, intereses) y los vuelca a un Excel. Usar cuando pidan relevar planes de pago, Mis Facilidades, cuotas de planes RG 5321, o deuda en plan de [sociedad]."
---

# Relevamiento de planes de facilidades de pago (Mis Facilidades - ARCA)

Objetivo: para una o varias sociedades, obtener todos los planes de facilidades de pago con su situación, y el detalle cuota por cuota (pagadas y pendientes), y dejarlo en un Excel por sociedad.

## Requisitos

- El usuario inicia sesión en ARCA con CUIT y Clave Fiscal a mano antes de arrancar (representante con las sociedades asociadas). Esta skill NO maneja credenciales.
- Usar las herramientas de Claude in Chrome (`mcp__claude-in-chrome__*`). Cargarlas en un solo ToolSearch: tabs_context_mcp, navigate, computer, read_page, get_page_text, find, form_input, tabs_create_mcp.
- Leer datos con `get_page_text` / `read_page`, no con capturas de pantalla.

## Entrada

- Sociedad(es): CUIT o nombre. Si dan nombre, buscar el CUIT en el Maestro de Sociedades o preguntar.
- Alcance: por defecto solo planes con Situación "Vigente". Si piden histórico, incluir también los "Plan Cancelado"/caducos.

## Pasos

1. `tabs_context_mcp` y ubicar la pestaña del Portal de Clave Fiscal (portalcf.cloud.afip.gob.ar/portal/app/). Si no hay sesión, pedirle al usuario que inicie sesión y avise.
2. En el buscador "¿Qué necesitás?" escribir "Mis Facilidades" y entrar al servicio (suele abrir una pestaña nueva en serviciossegsoc.afip.gob.ar/.../MisFacilidadesNet/). Volver a llamar `tabs_context_mcp` para tomar el tabId nuevo.
3. Pantalla "Seleccione una de las Cuits asociadas": elegir el CUIT de la sociedad en el desplegable (`form_input`) y confirmar. Verificar en la barra superior "Contribuyente: <CUIT> - <RAZÓN SOCIAL>" que sea la sociedad correcta.
4. En "Presentaciones Enviadas" (seguimiento_presentacion.aspx) leer la tabla completa: Presentación, Número, Cuotas, Tipo, Consolidado($), Estado, Situación. Si hay paginado, recorrer todas las páginas.
5. Por cada plan en alcance:
   a. Click en "Detalle" de esa fila → pantalla "Seguimiento de la presentación" (nuevos_planes.aspx). Confirmar Nro de plan, Fecha de consolidación y Tipo de plan.
   b. "Ver Pagos" → tabla "Pago Realizado" (detalle_pagos.aspx): Cuota N°, Capital, Interés Financiero, Interés Resarcitorio, Total, Fecha Venc., Pago, Estado de Cuota, y la fila Total Pagado. Una cuota puede ocupar DOS filas (primer intento impago + pago posterior con interés resarcitorio): registrar ambas filas, marcando la segunda como reintento, y tomar como fecha de pago efectiva la de la fila pagada.
   c. Volver y entrar a "Plan de Pago" para traer el cronograma completo, incluidas las cuotas futuras (necesario para saldo pendiente y próxima cuota). Un plan recién consolidado puede no tener pagos aún (Total Pagado $0,00): igual se relevan sus cuotas desde Plan de Pago.
   d. "Volver" hasta la lista de presentaciones y seguir con el siguiente plan.
6. Para otra sociedad: volver a la pantalla de selección de CUIT (o reingresar a Mis Facilidades) y repetir desde el paso 3.

## Normalización

- Importes en formato argentino (1.877.766,52) → número (1877766.52).
- Fechas DD/MM/AAAA.
- No inventar datos: si una pantalla no carga o falta un dato, dejarlo vacío y anotarlo en Observaciones.

## Salida: Excel por sociedad

Nombre: `planes_facilidades_[SOCIEDAD]_[AAAAMMDD].xlsx` (usar el skill xlsx para armarlo).

Hoja RESUMEN (una fila por plan): Nro plan, Fecha presentación, Tipo, Tipo de plan, Cant. cuotas, Consolidado, Situación, Cuotas pagadas, Capital pagado, Intereses pagados (financiero + resarcitorio), Total pagado, Saldo de capital (Consolidado − Capital pagado), Próxima cuota (N°, vencimiento, importe), Cuotas pagadas fuera de término.
Fila de totales al pie, con fórmulas.

Hoja DETALLE (una fila por cuota/fila de pago): Nro plan, Cuota N°, Capital, Interés financiero, Interés resarcitorio, Total, Fecha venc., Pago, Estado de cuota, Reintento (S/N).

Hoja OBSERVACIONES: pagos fuera de término, planes sin pagos, cuotas impagas vencidas, diferencias entre la suma de cuotas y el Total Pagado que informa ARCA.

Control: para cada plan, la suma del capital de las cuotas del cronograma debe igualar el Consolidado; y la suma de la hoja DETALLE pagada debe coincidir con "Total Pagado" de ARCA. Informar cualquier diferencia.

## Cierre

Entregar el Excel y resumir en el chat, por sociedad: cantidad de planes vigentes, saldo de capital total y próximos vencimientos. Si el usuario lo pide, subir el archivo a la carpeta de la sociedad en la unidad compartida.