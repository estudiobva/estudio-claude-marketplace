---
name: f931-descarga-arca
description: Obtiene los formularios F.931 ya presentados de una sociedad desde "Declaración en Línea" de ARCA/AFIP (serviciossegsoc.afip.gob.ar/djproforma) y entrega un PDF por período, nombrado AAAAMM.pdf. Corre un script de Playwright que se loguea solo con la planilla de claves del estudio; no hace falta entrar a ARCA a mano. Usá esta skill SIEMPRE que pidan "imprimir las DDJJ o los 931 de [sociedad]", "bajá los F.931 del período X al Y", "consultar declaraciones juradas generadas", "necesito los 931 del ejercicio", "entrá a [sociedad] y traeme los formularios 931", o cuando mencionen Declaración en Línea, djproforma, dj_generadas, ver_formulario, SICOSS o el menú "INGRESE...". NO confundir con la skill f931-sueldos, que toma PDFs de F.931 que YA existen y los vuelca en input_931.xlsx; esta va a buscarlos a ARCA. Si el pedido es "cargá los 931 en el Excel", esa es f931-sueldos; si es "conseguime los 931", es esta.
---

# F.931: bajarlos de Declaración en Línea (ARCA)

Un script de Playwright entra a ARCA con la clave de la planilla del estudio,
elige la sociedad en Declaración en Línea, lee el listado de DDJJ presentadas y
genera un PDF por período. No hace falta loguearse a mano.

El PDF es el formulario real: el botón "Imprimir" del sitio dispara
`window.print()`, pero Playwright renderiza la página con `page.pdf()` sin pasar
por ese diálogo. Sale en una hoja A4 apaisada, sin los botones.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/f931.js --nombre="<sociedad>" --listar
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/f931.js --nombre="<sociedad>" --desde=01/2026 --hasta=12/2026 --meses=12
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/f931.js --cuit=<cuit> --periodo=05/2026
```

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | La sociedad. El script resuelve login y representación solo. |
| `--desde` / `--hasta` | Rango inclusivo (`MM/AAAA` o `AAAAMM`). Sin `--hasta`, baja solo `--desde`. |
| `--periodo` | Atajo para un solo período. |
| `--meses=N` | **Aborta si el rango no tiene exactamente N períodos.** Ver abajo. |
| `--listar` | Solo lista lo presentado, no baja nada. |
| `--original` | Ante una rectificativa, bajar igual la original (sec 000). |
| `--salida` | Carpeta destino. Default `~/Documents/BVA-salidas/f931/[CUIT]/` (o `BVA_SALIDAS_PATH`). |
| `--cuit-login` / `--password` | Forzar otro login (raro: la planilla ya dice quién entra por cada sociedad). |
| `--ver` | Navegador visible. Solo si el usuario lo pide. |

**Plugin instalado:** la primera vez (y después de cada actualización):
`npm install && npx playwright install chromium`. Las credenciales van en
`~/.fisco-ar/.env` (copiar `.env.example`), que sobrevive a las actualizaciones.

## El login

Es el mismo que usan `anticipos-sct` y `deudas-ddjj-sct`: `lib/env.js` busca la
sociedad por nombre o CUIT en `Claves_Organismos.xlsx` (ruta y contraseña en el
`.env`), toma el CUIT de login (el propio o el del apoderado) y la clave de ARCA,
y `lib/arca-login.js` entra en los dos pasos (CUIT → clave). Después abre
Declaración en Línea desde el portal y elige la sociedad en el desplegable de
representadas.

Si una sociedad no está en la planilla, se puede cargar en el `.env` como
`ARCA_<cuit>_LOGIN=<cuit que entra>` y `ARCA_<cuit>_PASSWORD=<clave>`: el `.env`
tiene prioridad sobre la planilla.

## Antes de correr: las dos confirmaciones que importan

**La sociedad.** Los nombres se parecen muchísimo ("SOY MAGA SANAR 5" vs "MAGA
SANAR 5" vs "SOY MAGA SUR SANAR 6"). Si hay más de un candidato razonable,
mostrale la lista con los CUIT y que elija. El script además falla explícito si
el CUIT no está en el desplegable de Declaración en Línea.

**El rango.** El usuario suele decir el rango *y* la cantidad esperada de meses
("de 202504 a 202605, tienen que ser 12"). **Pasá siempre `--meses=N` cuando
diga un número.** Ese ejemplo son 14 períodos, no 12, y el script frena:

```
El rango 202504-202605 tiene 14 periodos, no 12.
```

Descubrirlo después significa rehacer o tirar trabajo, y el usuario casi siempre
tiene en la cabeza un ejercicio concreto (abril–marzo, junio–mayo).

## Qué hace el script

Login → Declaración en Línea → elegir la sociedad en el desplegable → pasar los
avisos → **listado de DDJJ generadas** → para cada período del rango, abrir el
formulario y renderizarlo a PDF.

Visitar el listado **no es opcional** aunque después se navegue por URL: sin ese
estado de sesión, `ver_formulario.aspx` responde *"No puede acceder a la página
solicitada"*. El script ya lo hace en el orden correcto.

### Rectificativas

El listado trae el historial completo (puede pasar las 170 DDJJ). Cuando un
período tiene original y rectificativa, **el script baja la secuencia vigente**
—la más alta— y lo avisa por stderr:

```
con rectificativa: 202607 (Rectificativa (1)) — bajo la VIGENTE
```

Si el usuario quiere la original, `--original`. Preguntale cuál quiere cuando
aparezcan rectificativas en el rango: puede cambiar el importe.

## Qué reportar

Del JSON: sociedad y CUIT, cantidad de períodos, **`faltantes`** (un mes sin DDJJ
presentada es un hallazgo, no un error del script), **`conRectificativa`**, y el
detalle con empleados y Suma de Rem. 1 por período.

Sumá lo que te llame la atención de los datos —saltos de nómina, un mes con
importes muy distintos, períodos faltantes—: eso es justo lo que el usuario va a
mirar después.

Mandá los PDF con `SendUserFile` y, si son más de tres o cuatro, sumá un ZIP.

## Control de calidad

El script verifica cada PDF y avisa si alguno no salió en **una sola página**.
Si eso pasa, mirá el PDF antes de entregarlo: puede ser una sociedad con una
nómina más grande que la escala prevista.

## Cuando falla

| Código | Qué hacer |
|---|---|
| `credenciales_invalidas` | **No reintentar**: bloquea la clave fiscal. Revisar la fila de la planilla. |
| `captcha` | Cortar y avisar: alguien entra una vez a ARCA a mano con esa clave y se reintenta. |
| `rango_no_coincide` | El chequeo de `--meses`. Confirmar el rango con el usuario. |
| `servicio_no_disponible` | Declaración en Línea no abrió desde el portal. Suele ser ARCA; reintentar una vez más tarde. |
| `error` | Leer el mensaje. |

Si el CUIT no aparece en el desplegable, falta la relación en ARCA y **eso lo
resuelve el usuario**.

### Detalles del sitio que ya están resueltos en el código

- Hay que entrar **por el tile** del portal; la URL directa no autentica.
- La pantalla de control de acceso usa `select[id*="ddlCUIT"]` (el `value` es el
  CUIT sin guiones) y un `input[type=image][id*="btnAceptar"]`. Los logins con
  un solo CUIT no muestran el desplegable.
- Después hay **una o dos pantallas de avisos**, cada una con su ACEPTAR.
- **Visitar `dj_generadas.aspx` fija el estado de sesión** que `ver_formulario`
  necesita.
- La secuencia sale del número entre paréntesis de la columna Secuencia
  (`Original (0)` → `SecDJVig=000`, `Rectificativa (1)` → `001`).
- Sin CSS de impresión el PDF sale en **dos páginas, la segunda en blanco**. El
  script inyecta `@page { size: A4 landscape }`, usa `scale: 0.9` y oculta
  Imprimir/CERRAR (son imágenes con `onclick`, no botones).

## Lo que el script no hace

- **Nunca genera ni presenta una DDJJ.** Solo consulta y descarga.
- No archiva en el Drive: no hay convención documentada de carpeta para los
  F.931 (a diferencia de Portal IVA). Salen a `~/Documents/BVA-salidas/f931/[CUIT]/`
  y de ahí se entregan o se mueven a donde corresponda.
- No recorre todas las sociedades de una: se corre por sociedad.
