---
name: recibo-pyp-quilmes
description: "Baja de ARQui (Agencia de Recaudación del Municipio de Quilmes) el recibo actual de Derechos de Publicidad y Propaganda (marquesinas, carteles) de una, varias o todas las sociedades, controla que el CUIT y la cuenta del recibo sean los de la sociedad, lo archiva en la carpeta TSH del mes como 'AAAAMM-boleta marquesina- <Sociedad>.pdf' y resume cuenta, clase, área, importe, vencimiento y Nº de recibo en un Excel. Corre un script de Playwright que se loguea solo con el usuario de ARQui de la planilla de claves del estudio: no hace falta entrar a mano. Usar SIEMPRE que pidan \"bajá el recibo de PyP de [sociedad]\", \"recibo/boleta de publicidad y propaganda\", \"boleta de marquesina de Quilmes\", \"tasa de publicidad Quilmes\", \"recibos de marquesinas de todas las sociedades\", \"cuánto paga de marquesina\" o mencionen ARQui y Publicidad y Propaganda. No paga nada: para pagar, la persona entra a ARQui."
---

# Recibo actual de Publicidad y Propaganda — ARQui Quilmes

Un script de Playwright entra a ARQui con el usuario de la planilla de claves, va a
**Publicidad y Propaganda > Recibo Actual**, elige la cuenta, baja el PDF, lo lee,
controla que sea de la sociedad y lo archiva. Navegador oculto. Solo lee y
descarga: nunca usa "Botón de pago" / "Mis pagos online", no presenta DDJJ y no
vincula cuentas con el "+".

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/recibo-pyp-quilmes.js --nombre="<sociedad>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/recibo-pyp-quilmes.js --nombres="<A>;<B>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/recibo-pyp-quilmes.js --todas
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/recibo-pyp-quilmes.js --listar
```

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un login para todas. |
| `--todas` | Todas las cuentas de PyP vinculadas a los usuarios de ARQui. La sociedad sale del CUIT impreso en cada recibo. Suma la hoja **Control planilla**. |
| `--cuenta` | Nº de cuenta (sin dígito verificador) para una sola sociedad, si no está en la planilla de tasas. |
| `--listar` | Solo lista las cuentas del desplegable de ARQui contra la planilla de tasas. No baja nada. |
| `--no-archivar` | Deja los PDF solo en la carpeta de salida; no toca la unidad. |
| `--reemplazar` | Pisa el PDF si ya estaba en TSH (por defecto lo deja como está). |
| `--ver` | Navegador visible. Solo si el usuario lo pide. |

Si falta la sociedad, preguntar antes de correr. `--todas` tarda unos 3–4 minutos
(hay una pausa entre recibos a propósito, ver "Bloqueos").

**Plugin instalado:** la primera vez (y después de cada actualización):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Usa el mismo `~/.fisco-ar/.env` que los demás plugins (`CLAVES_ORGANISMOS_PATH`,
`CLAVES_ORGANISMOS_PASSWORD`, y `BVA_UNIDAD_PATH` en Mac).

## De dónde sale cada dato

- **Login:** `Claves_Organismos.xlsx`, grupo **MUNICIPAL** de la fila de la sociedad
  (Usuario | Clave | Municipio = Quilmes). Casi todas comparten el usuario del
  estudio; un par de sociedades tienen uno propio. Si la sociedad no tiene fila de
  Quilmes, se usa el usuario de Quilmes más repetido. Override puntual:
  `ARQUI_USUARIO` / `ARQUI_CLAVE` en el `.env`. Claude nunca escribe ni muestra la clave.
- **Nº de cuenta:** `Tasas Municipales CLAVES.xlsx` (carpeta `0 - ESTUDIO - 00000000000`
  de la unidad, u otra ruta con `TASAS_MUNICIPALES_PATH`), hoja **QUILMES**, bloque
  **PUBLICIDAD Y PROPAGANDA**, buscando **por CUIT** (los nombres de la planilla no
  coinciden con los de las carpetas: abreviados, con espacios raros o distintos
  para el mismo CUIT). Se ignoran los bloques Seguridad e Higiene y Marquesinas:
  a veces traen otro número. En ese bloque las cuentas van sin dígito
  verificador, igual que en el desplegable de ARQui.
- **Si la sociedad no está en el bloque** (o la fila tiene `#N/A` / vacío), el script
  busca su recibo entre las cuentas del desplegable que la planilla no le asigna a
  nadie, controlando el CUIT de cada PDF. Si la encuentra la usa y avisa "cargarla
  en la planilla"; si no, queda `sin cuenta`. Nunca adivina un número.

## Controles

No archiva (estado `error`, código `control`) si el CUIT impreso en el recibo no es
el de la sociedad o la cuenta del recibo no es la elegida (más su dígito
verificador, ej. `500123456-7`). Si la cuenta de la planilla no está en el
desplegable del usuario queda `no_vinculada`: no se agrega sola; vincularla con el
"+" de ARQui lo decide y lo hace la persona.

## Qué entrega

- PDF archivado en `[Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-boleta marquesina- <Sociedad>.pdf`
  (AAAAMM = el mes del recibo; con carpeta de año va en `AAAA/AAAAMM/TSH/`). Es el
  nombre que ya usa el equipo. Si la sociedad tiene más de una cuenta, se agrega
  ` - <cuenta>` al nombre. Si el archivo ya estaba, no se pisa (`ya estaba`).
- Copia de los PDF y resumen `RECIBOS PYP QUILMES <fecha> <hora>.xlsx` en
  `~/Documents/BVA-salidas/recibo-pyp-quilmes/AAAA-MM-DD/` (`BVA_SALIDAS_PATH`):
  hoja **Recibos** (Sociedad, CUIT, Cuenta, Destinatario, Clase, Área, Período,
  Vencimiento, Vencido, Importe, Nº recibo, Estado, Observación, Archivo) y, con
  `--todas`, hoja **Control planilla** (cada cuenta de ARQui contra la planilla de
  tasas: OK, falta en la planilla, CUIT distinto, en la planilla pero no vinculada).
- JSON por stdout con lo mismo y `totales`.

El "recibo actual" es el del mes en curso: si ya pasó el vencimiento, ARQui igual
lo da y el resumen lo marca `Vencido`.

## Bloqueos de ARQui (403 Forbidden)

El firewall del municipio corta todo (login incluido) con **403 Forbidden** durante
unos minutos. Lo que lo dispara (probado el 22/10/2026): entrar con otro usuario
mientras sigue abierta la sesión del usuario general después de muchas descargas.
Loguear varios usuarios seguidos con la sesión anterior cerrada y cookies limpias
no lo dispara. Por eso el script tiene **una sola sesión abierta**: al cambiar de
usuario cierra la anterior, y hace primero los usuarios propios (pocas cuentas) y
al final el general. Además no recarga la pantalla entre cuentas, no baja imágenes
y hace una pausa `ARQUI_PAUSA_MS` (2,5 s) entre recibos. Si
igual aparece, espera `ARQUI_ESPERA_BLOQUEO_MS` (3 min) y reintenta una vez; si
sigue, corta el resto con código `bloqueo` y `totales.bloqueado: true`. En ese caso
**no relanzar enseguida**: esperar 5–10 minutos y volver a correr solo las que
faltaron.

## Respuesta al usuario

Una línea por sociedad: Sociedad | Cuenta | Importe | Vencimiento (y si está vencido) |
Nº de recibo, y dónde quedó el archivo. Después, en una línea cada uno, lo que no
salió y por qué, sin adivinar:

- `sin cuenta`: no está en la planilla de tasas ni entre las cuentas de ARQui del usuario.
- `no_vinculada`: la cuenta de la planilla no está en el desplegable de ese usuario de ARQui.
- `sin recibo`: ARQui no devolvió recibo para la cuenta y muestra el motivo. Al 22/10/2026,
  4 cuentas del desplegable del usuario general dan "La cuenta ingresado no existe."
  (probablemente dadas de baja). No es una falla del script.
- `control`: el recibo no es de esa sociedad o de esa cuenta; no se archivó.
- `login`: ARQui no aceptó el usuario/clave de la planilla.
- `bloqueo`: ver arriba.

Con `--todas`, resumir también la hoja Control planilla: qué cuentas faltan en la
planilla de tasas y cuáles tienen otro CUIT, para que el equipo la corrija.
