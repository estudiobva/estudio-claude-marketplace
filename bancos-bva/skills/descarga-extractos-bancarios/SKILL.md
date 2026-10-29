---
name: descarga-extractos-bancarios
description: Descarga extractos bancarios (PDF/Excel) del home banking de empresas de los 9 bancos del Estudio BVA (Galicia, Comafi, Santander, Hipotecario, Macro, Nación, BBVA/Francés, Patagonia, Provincia), los renombra y los deja en ~/Documents/BVA-salidas/extractos/[N] - [Sociedad] - [CUIT]/[BANCO]/[AAAA]/ (o, con --unidad, en 03-Bancos y Conciliaciones/[BANCO]/[AAAA]/ de la sociedad). Todo lo hace el script de Playwright scripts/extractos-banco.js (login solo con Claves_Bancos.xlsx cifrado, captura de descargas, renombrado y archivado sin pisar nada), que corre Claude igual que los scripts de ARCA. Usar SIEMPRE que pidan "descargar extractos de [banco] para [sociedad]", "bajá los resúmenes de [sociedad] de enero a junio", "extractos de todos los bancos de [sociedad]", "automatizar descarga bancaria", o quieran avanzar banco por banco/sociedad por sociedad. Es el paso PREVIO a extractico-bva (depura los PDF) y a extractos-bancarios (ordena el Excel depurado).
---

# Descarga de extractos bancarios

Hoy esto se hace a mano, banco por banco: **1 día completo por 1 a 1,5
sociedades**. Circuito: `descarga-extractos-bancarios` → `extractico-bva` →
`extractos-bancarios`.

Según el Maestro de Sociedades, **68 de las 79 entidades operan con los mismos 7
bancos** (Galicia, Comafi, Hipotecario, Nación, Francés/BBVA, Macro, Santander);
una entidad suma Patagonia y Provincia, y 10 entidades no tienen bancos cargados.

## El comando

Todo corre desde la raíz del plugin. Si es la primera vez en la PC, usar antes
la skill `bancos-setup` (instala y corre el diagnóstico).

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/extractos-banco.js --nombre="<sociedad>" --banco=todos --desde=202601 --hasta=202606
```

| Flag | Qué hace |
|---|---|
| `--cuit` / `--nombre` | Sociedad. El nombre es el de la carpeta de la unidad; si es ambiguo el script corta y lista los CUIT. |
| `--banco` | Uno, varios separados por coma, o `todos` (los que el Maestro marca con "Sí" para esa sociedad). |
| `--desde` / `--hasta` / `--periodo` | `AAAAMM` o `MM/AAAA`. Máximo 24 meses por corrida. |
| `--sociedad-claves` | Solo Galicia: nombre de la sociedad en el bloque Galicia del archivo de claves, si difiere del de la carpeta. |
| `--salida` | Carpeta destino. Default `~/Documents/BVA-salidas/extractos/` (o `BVA_SALIDAS_PATH`). |
| `--unidad` | Archiva directo en la unidad: `[Sociedad]/03-Bancos y Conciliaciones/[BANCO]/[AAAA]/`. |
| `--ver` | Navegador visible. **Solo si el usuario lo pide.** Sin `--ver` corre SIEMPRE oculto, también con `--control`. |
| `--control` | Claude navega el banco (oculto) con `scripts/control.js`. Obligatorio mientras un banco no tenga descarga automática (hoy ninguno): sin `--control` ni `--ver`, esos bancos se saltean con `requiere_ventana`. |
| `--grabar` | Guarda en `~/bva-inspeccion/` los clics de la persona dentro del banco (nunca lo tipeado), para automatizar la navegación de ese banco. |

Qué hace, por cada banco:

1. **Inicia sesión** con `Claves_Bancos.xlsx`. Intenta **una sola vez**; si el
   banco rechaza la clave corta con `credenciales_invalidas` y no reintenta. Si
   pide token/código o captcha, lo pide en la consola **y en un recuadro con botón
   "Listo" dentro de la ventana del banco** (sirve cualquiera de los dos), y espera
   a que la persona lo resuelva.
2. **Descarga.** Si el banco tiene descarga automática (`descargar` en su
   adaptador) la hace solo. Si no — hoy ninguno la tiene —, la consola indica
   qué bajar y la persona navega en la ventana; el script toma cada descarga y
   también los PDF que el banco abre en el visor.
3. **Renombra.** El período sale del nombre que pone el banco
   (`Extracto_Cuentas_Galicia_2026_01_23.pdf` → enero, semana 4); si no se
   puede leer o cae fuera de lo pedido, lo pregunta en la consola.
4. **Guarda** en `~/Documents/BVA-salidas/extractos/[Sociedad]/[BANCO]/[AAAA]/`
   (con `--unidad`, en `03-Bancos y Conciliaciones/[BANCO]/[AAAA]/`; todos los meses
   del año juntos) sin pisar nada: si el mismo contenido ya está (con cualquier
   nombre, también en una carpeta vieja `[AAAAMM]/`) lo saltea; si hay otro
   archivo distinto con el mismo nombre agrega `_2`. Rechaza "PDF" que no son
   PDF (páginas de error o de sesión vencida).
5. **Reporta** en JSON: guardados, ya existentes, rechazados, descartados,
   **períodos sin archivo** y, por banco, `error` con `code` y `que_hacer`.
   Nunca imprime credenciales.

**Quién lo corre:** Claude, igual que `anticipos-sct` y los demás scripts de
ARCA. El login lo hace el script con la planilla de claves: Claude nunca ve ni
tipea las claves. **El navegador corre siempre oculto** (pedido del usuario,
29/09/2026): Claude nunca agrega `--ver` por su cuenta, solo si el usuario pide
ver la ventana. Mientras no haya descarga automática, Claude corre con
`--control` en segundo plano y navega el banco oculto. Si un banco pide token o
captcha, oculto no se puede resolver: el banco corta con `segundo_factor` /
`captcha`, se sigue con el próximo y se le avisa al usuario (que decide si
correr ese banco con `--ver`).

## Modo control: Claude navega y descarga (`--control`)

Para los bancos que todavía no tienen descarga automática:

1. Claude corre el script con `--control` en segundo plano, oculto (el login lo
   sigue haciendo el script):
   `node scripts/extractos-banco.js --cuit=<CUIT> --banco=MACRO --desde=AAAAMM --hasta=AAAAMM --control --grabar`
2. Cuando la terminal muestra `CONTROL <BANCO> adentro — navegador oculto disponible en http://127.0.0.1:9333`,
   Claude maneja ese navegador con `scripts/control.js` (`captura`, `elementos`, `clic`,
   `elegir`, `completar`, `listo`, `responder`) y descarga los períodos pedidos. El puerto
   es solo local (127.0.0.1). **Si la línea termina en `· ya en <sección> de "<EMPRESA>": bajar
   desde ahi`, el script ya eligió la empresa y abrió la sección de resúmenes (tabla de abajo):
   Claude baja desde esa pantalla y NO va a Movimientos ni a otra parte del banco**, ni cambia
   de empresa. Solo si la terminal muestra `AVISO no llegue solo a la seccion`, Claude sigue la
   ruta de la tabla desde la empresa ya elegida.
3. Claude baja cada resumen con `control.js descargar --sel=... --banco=X --cuit=... --periodo=AAAAMM`
   **después de verificar el período dentro del PDF** (la fecha de emisión no siempre es el
   mes: Nación, Provincia y Patagonia emiten a fin de mes o principios del siguiente).
   Si hay dudas, primero sin `--periodo` (queda en `~/bva-inspeccion/bandeja/`), se revisa y
   después `control.js archivar --archivo=<bandeja> --banco=X --cuit=... --periodo=AAAAMM`.
   Con `--control`, `extractos-banco.js` **no** vuelve a tomar esas descargas ni pregunta
   períodos; lo que archiva `control.js` queda en un registro local
   (`<tmp>/bva-control-<puerto>.jsonl`) y el script lo suma a `guardados` / `ya_estaban` /
   `sin_archivo` de su JSON. Un PDF que el banco abre en el visor y que no está archivado
   va a la bandeja y aparece en `en_bandeja`. Al terminar, Claude toca `listo` y el script
   pasa al banco siguiente.

Límites de `control.js`: no escribe en campos de clave/token/PIN/código, avisa si la
página sale del sitio del banco y no cierra la sesión (la cierra `extractos-banco.js` al
terminar cada banco). Dentro del banco, solo navegar y descargar resúmenes: nada de
transferencias, pagos ni configuración.

### Sección fija de extractos por banco

Pedido del equipo (22/10/2026): que entre siempre a la misma parte del banco. Después del
login el script (`lib/bancos/seccion-extractos.js`) elige la empresa de la sociedad **por
CUIT** (si no la encuentra corta con `empresa_no_encontrada`: nunca entra a otra) y abre:

| Banco | Sección | Qué se baja ahí | Período |
|---|---|---|---|
| Macro | menú de la cuenta → **Resúmenes** ("Listado de resúmenes") | `a#collectionTable1022_DETAIL` (id repetido: `--n`, el último es el más nuevo) → `Resumen.pdf` = Resumen General de todas las cuentas | "Periodo del Extracto" dentro del PDF; la fecha de la fila es el cierre (30/09/2026 = septiembre) |
| Nación | cuenta → Más opciones → **Extracto de cuenta** → Extracto Mensual | `#downloadPDFPeriod0` = el más nuevo ("Mostrar más" para los viejos) | "PERIODO: 31/08/2026 AL 30/09/2026" en el PDF; la emisión 01/10 = septiembre |
| Patagonia | Cuentas → **Resumen electrónico** → "Resúmenes" de la cuenta | `#dateFieldDesde/Hasta` → Buscar → `#collectionTableMonthly_DETAIL` | emisión 30/09 = septiembre |
| Provincia | Cuentas → Mis Cuentas → Ver movimientos (CC) → **Extractos Electrónicos** | `button[id^=pdf-]` (el último es el más nuevo) | emisión 01/10 = septiembre; una caja de ahorro puede tener extractos diarios en vez de mensuales |

Sin sección fija todavía (Claude navega desde el inicio): Galicia, Hipotecario, Comafi,
Santander, BBVA. El ingreso es siempre por el **link de la columna LINKS de
Claves_Bancos.xlsx** (el destino del hipervínculo, no el texto "BNA Digital"), si es del
mismo banco; la terminal lo dice: `entrando por <url> (link de Claves_Bancos)`.

### Códigos de error del login

| code | Qué pasó |
|---|---|
| `credenciales_invalidas` / `bloqueado` | El banco rechazó los datos o bloqueó el usuario. No reintentar. |
| `clave_vencida` | El banco pide cambiar la clave (Santander, Galicia por empresa). La cambia la persona a mano. |
| `antibots` | Quedó en la pantalla de espera sin dejar entrar (BBVA "Verificando tus credenciales"). |
| `captcha` / `segundo_factor` | Oculto no se puede resolver: el usuario decide si correr ese banco con `--ver`. |
| `sitio_rechazo` | El banco cortó la conexión (ej. Santander oculto: `ERR_HTTP2_PROTOCOL_ERROR`). |
| `sesion_activa` | Quedó abierta la sesión de una corrida anterior (Macro: 10 min; Provincia: "Tu usuario ya se encuentra conectado"). Esperar y volver a correr ese banco. |
| `empresa_no_encontrada` | El usuario del banco no tiene esa sociedad o el nombre no se emparejó: no se entra a otra empresa. |

## Login por banco

Relevado el 24/09/2026 sobre la pantalla pública de cada banco
(`lib/bancos/<banco>.js`, motor común en `lib/bancos/comun.js`):

| Banco | URL | Datos (columna de Claves_Bancos.xlsx) | Particularidades |
|---|---|---|---|
| Galicia | empresas.bancogalicia.com.ar/login | Usuario, Clave | Login propio de cada sociedad (bloque Galicia: 50 de 60 con datos; la fila de login único está vacía). |
| Comafi | ebanking.comafiempresas.com.ar/login | Usuario (USUARIO), Clave de Acceso (ACCESO/DNI), Clave de Identificación (CLAVE ID) | — |
| Santander | empresas.santander.com.ar/login | Usuario, Contraseña | La portada ya menciona "Token": se mira solo el texto nuevo. |
| Hipotecario | empresas.hipotecario.com.ar/OB/login | DNI (ACCESO/DNI), Usuario, Contraseña | Aviso "Entendido" que se cierra solo. |
| Macro | macro.com.ar/biempresas | Usuario → Continuar → Clave → INGRESÁ | Dos pasos, ambos vistos en la corrida real del 24/09/2026 (el paso 2 muestra la imagen de seguridad). Desloguea al cambiar de sociedad. |
| Nación | digital.bna.com.ar/loginStep1 (BNA Digital, el link del archivo de claves) | DNI (ACCESO/DNI), Usuario → Continuar → Clave | Dos pasos; el segundo no se pudo ver sin usuario real. **No** es Nación Empresa 24 (bee3.redlink), que es otra plataforma. |
| BBVA (Francés) | netcash.bbva.com.ar | Código de empresa, Código de usuario, Clave | **Mapeo a confirmar**: se asume empresa = ACCESO/DNI, usuario = USUARIO. |
| Patagonia | ebankempresas.bancopatagonia.com.ar | Tipo+número de documento (ACCESO/DNI, 11 dígitos), Usuario → Clave | Dos pasos, segundo no visto. |
| Provincia | bancoprovincia.bancainternet.com.ar/spa-empresas | Usuario, Clave | — |

Probados contra réplicas de sus formularios y, contra el banco real: Macro, Nación,
Patagonia y Provincia (28/10/2026, una sociedad, ingreso + sección + descarga + salida).

Login suelto (sin descargar), para diagnosticar:
`cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/banco-login.js --banco=MACRO --quedate` (con `--inspeccionar`
guarda captura y estructura de la pantalla de inicio).

## Credenciales

`Claves_Bancos.xlsx` vive en `0 - ESTUDIO - 00000000000/` y está **cifrado**
(`CLAVES_BANCOS_PASSWORD` en `~/.fisco-ar/.env`, o `CLAVES_ORGANISMOS_PASSWORD`,
que hoy es la misma) — el mismo `.env` de fisco-ar. La unidad compartida se
detecta sola en Windows (`I:\`) y en Mac (Google Drive para escritorio).

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/claves-bancos.js --faltantes
```

Muestra los valores enmascarados (`***(8)` = 8 caracteres). Claude nunca usa
`--revelar`, nunca abre el `.xlsx` y nunca pega credenciales en un chat,
archivo o log.

Estructura de la hoja `BANCOS` (21/09/2026): filas 2–10, un banco por fila con
login único (COMAFI, PATAGONIA, PROVINCIA, HIPOTECARIO, NACION, FRANCES, MACRO,
SANTANDER, GALICIA); desde la fila 11, bloque Galicia con login por sociedad
(columnas desalineadas con su encabezado — `lib/claves-bancos.js` mapea por
etiqueta). 10 sociedades del bloque Galicia no tienen usuario o clave
(`node scripts/claves-bancos.js --faltantes` las lista). La copia vieja `extracos-bancarios/claves.xlsx` **no se usa**.

**Galicia: cómo se encuentra el login de cada sociedad.** El bloque Galicia no
trae CUIT, así que se busca por **nombre exacto** de la carpeta (37 de 55
sociedades). Las que figuran distinto se resuelven con `lib/galicia-alias.json`
(CUIT → nombre en el archivo), que **solo se usa si la equivalencia está
confirmada** (`"confirmado": true`). Nunca por parecido: un nombre corto de carpeta
se parece al nombre largo de OTRA sociedad del bloque y entraría con su usuario.
Al 24/09/2026: 10 equivalencias confirmadas; 45 de las 55 sociedades con Galicia
tienen login resuelto (el resto: sin fila, fila duplicada o sin usuario/clave).

`galicia-alias.json` tiene datos de clientes y **no se sube al repo**: en un
plugin instalado va en `~/.fisco-ar/galicia-alias.json` (formato:
`{ "<cuit>": { "nombreEnClaves": "<nombre en el bloque>", "confirmado": true } }`).

## Destino y nombres

```
~/Documents/BVA-salidas/extractos/  (default; --salida o BVA_SALIDAS_PATH)
  [N] - [Sociedad] - [CUIT]/        ej. "8 - Sociedad Ejemplo - 30000000007"
    [BANCO]/                        GALICIA/COMAFI/SANTANDER/HIPOTECARIO/MACRO/NACION/BBVA/PATAGONIA/PROVINCIA
      [AAAA]/
        [sociedad]_[banco]_[AAAA-MM]_mensual.pdf

Con --unidad:
[N] - [Sociedad] - [CUIT]/
  03-Bancos y Conciliaciones/
    [BANCO]/                        carpeta existente, o GALICIA/COMAFI/SANTANDER/HIPOTECARIO/MACRO/NACION/BBVA/PATAGONIA/PROVINCIA
      [AAAA]/                       ej. 2026 — todos los meses del año juntos
        [sociedad]_[banco]_[AAAA-MM]_mensual.pdf
        [sociedad]_[banco]_[AAAA-MM]_semana[N].pdf     (Galicia)
```

- `[sociedad]` = nombre de la carpeta sin número ni CUIT, en minúsculas con `_`
  (`sociedad_ejemplo`).
- Solo con `--unidad` — carpeta de banco: se reusa la que exista sin distinguir mayúsculas
  (`Galicia`, `bco. Galicia-cta 3393-5`, `Frances` = BBVA). Si hay dos
  candidatas, el script pregunta cuál.
- Falta `03-Bancos y Conciliaciones`: se crea en sociedades (CUIT 30/33/34); en
  personas humanas el script corta y hay que confirmarlo a mano.
- El staging del piloto (`automatizacion_bva/extracos-bancarios/extractos-descargados/`,
  298 archivos de 26 sociedades) **no es destino**; migrarlo es una tarea aparte.
- El histórico con otros nombres (`AAAAMM_GALICIA_SOCIEDAD_[cuenta].pdf`,
  `Extracto_Cuentas_Galicia_AAAA_MM (N).pdf`) no se renombra.

## Cómo trabaja Claude con esta skill

0. Primera vez en la PC: skill `bancos-setup`. Si ya está, correr igual
   `node scripts/doctor.js` y no seguir si da `❌`.
1. Confirmar sociedad (con CUIT), bancos y períodos. Si el nombre es ambiguo,
   mostrar candidatos con CUIT.
2. Revisar `claves-bancos.js --faltantes`: si a la sociedad le faltan claves de
   un banco, avisar y sacarlo del pedido.
3. Correr `extractos-banco.js` en segundo plano (Bash con `run_in_background`).
   Primera vez con un banco: agregar `--grabar`. Una sociedad por corrida.
4. Avisar a la persona que se abre la ventana del banco: ahí elige la sociedad,
   baja los resúmenes pedidos y toca "Listo" (token o captcha también ahí).
5. Leer el JSON que devuelve y reportar: archivos guardados por banco y mes,
   **períodos sin archivo**, rechazados y errores (`code` + qué hacer). Ante
   `credenciales_invalidas` o `bloqueado`, **no** sugerir reintentar enseguida.
6. Si hubo `--grabar`, ofrecer convertir la grabación en descarga automática
   para ese banco (`descargar()` en su adaptador).

## Reglas de seguridad

- Las credenciales las maneja solo el script (como en los scripts de ARCA).
  Claude no las lee, no las tipea y nunca usa `--revelar`.
- Captchas, tokens y preguntas de seguridad los resuelve la persona.
- Si un banco falla, el script guarda captura y texto de la pantalla en
  `~/bva-inspeccion/<banco>-error-*.png/.txt`: pedir esa captura para diagnosticar.
- No se modifica, renombra ni elimina nada existente (ni en la salida ni en la unidad).
- No inventar ni adivinar CUIT o carpeta.
- En el home banking solo consultar y descargar resúmenes: nada de
  transferencias, pagos ni configuración.
- De a una sociedad por corrida.

## Pendientes

- **Primera corrida real de cada banco** (con `--grabar`), y confirmar el mapeo
  de columnas de BBVA y el segundo paso de Macro y Patagonia.
- Descarga automática por banco a partir de las grabaciones.
- Migrar el staging del piloto; eliminar la copia vieja de `claves.xlsx`.
- AMEX aparece en 4 sociedades pero no está en el archivo de claves ni en el
  Maestro: definir si entra.

## Historial

- 21/09/2026: rutas y lectura de credenciales corregidas.
- 24/09/2026: destino único, reglas de nombres y carpetas, rutas Mac/Windows.
- 24/09/2026: login por Playwright para los 9 bancos y `extractos-banco.js`
  (descarga asistida + renombrado + archivado), probados contra réplicas.
- 28/09/2026: funciona como los scripts de ARCA (`anticipos-sct`): lo corre
  Claude, salida en `~/Documents/BVA-salidas/extractos/` (`--salida`, `--unidad`),
  carpeta por año `[BANCO]/[AAAA]/`, `--ver` en vez de `--oculto`, errores con
  `code` + `que_hacer`.
- 29/09/2026 (2): `--control` sin duplicados (el script ya no re-toma lo que baja
  `control.js`; registro compartido para el JSON), Provincia pregunta el período cuando
  la fecha del nombre es de los días 1-5 (mensual = mes anterior, diario = ese mes),
  BBVA ya no se da por "adentro" en el spinner (`antibots`), nuevos `clave_vencida` y
  `sitio_rechazo`. Probado contra un banco local falso.
- 29/09/2026: navegador SIEMPRE oculto salvo `--ver` pedido por el usuario
  (antes se abría visible solo con `--control` o bancos sin descarga
  automática). `--control` funciona oculto; `banco-login.js` también es oculto
  por defecto (`--ver` reemplaza a `--oculto`).
- 28/10/2026: ingreso por el hipervínculo real de la columna LINKS de Claves_Bancos
  (antes se leía el texto de la celda: "BNA Digital", "Banco Comafi") y la terminal dice
  por qué URL entra; sección fija de extractos con elección de empresa por CUIT (Macro,
  Nación, Patagonia, Provincia); el script cierra la sesión del banco al terminar cada uno;
  `sesion_activa` también para Provincia. Probado con el período 09/2026 en los 4 bancos.
