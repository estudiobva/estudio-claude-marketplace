---
name: conciliacion-ncr
description: "Concilia las compras del Portal IVA de ARCA de una sociedad contra las Notas de Credito de Recupero (NCR) de Suizo Argentina y Monroe Americana, y deja el resultado como archivo nuevo en la carpeta ncr/ del periodo. Cubre el circuito completo: bajar las compras del Portal IVA y las NCR del SharePoint de las farmacias (NC por NRF), conciliar, y eliminar del borrador del Libro Compras las NCR conciliadas. Usar SIEMPRE que pidan \"conciliar NCR\", \"conciliar [sociedad] [periodo]\", \"avanza/procesa con [sociedad]\", \"compras sin NCR\", \"eliminar/borrar las NCR del portal\", \"depurar compras\", o mencionen Notas de Credito de Recupero, NCR de Suizo o Monroe, archivos tipo 202608-SUIZO NCR.xlsx o exportaciones por farmacia. Tambien si reportan una diferencia entre el total de compras y el conciliado, NCR \"no encontradas\" o notas de credito de las droguerias sin respaldo en NCR."
---

# Conciliacion Portal IVA vs NCR (Suizo + Monroe)

Saca de las compras del Portal IVA de ARCA las notas de credito que ya estan
respaldadas por una NCR de la drogueria, y deja el resto en un archivo nuevo.
Validado end-to-end contra el periodo 08/2026: en todas las NCR conciliadas el
importe y la fecha coincidieron al centavo y al dia.

Se trabaja de a una sociedad y un periodo. La base es el Portal IVA — la hoja
MCR **no** participa de este flujo.

Antes de escribir el .xlsx de salida, leer la skill `xlsx`.

## Circuito completo

La conciliacion es el paso 3 de 4. Los pasos 1, 2 y 4 se corren con los scripts
de Playwright del estudio (`C:\Users\<usuario>\bva-playwright`, espejo en
`0 - ESTUDIO - 00000000000/automatizacion_bva/playwright/`) y necesitan una
sesion de Claude Code en la maquina, con la unidad compartida montada. Desde
Cowork solo se puede hacer el paso 3: los otros se hacen a mano o se piden a
quien tenga los scripts.

| Paso | Que | Como |
|---|---|---|
| 1 | Compras del Portal IVA | `node scripts/portal-iva-descarga.js --nombre="<Sociedad>" --periodo=MM/AAAA --libros=compras` (el lote lo hace solo si faltan) |
| 2 | NCR del SharePoint | `bundle-ncr.js` en Chrome + `node scripts/descargar-ncr.js` (ver abajo) |
| 3 | Conciliacion | `python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=MM/AAAA` (ver **Script**) |
| 4 | Eliminar las NCR del Libro Compras | `node scripts/eliminar-ncr.js` (ver abajo) |
| 1+2+3+simulacion del 4, varias sociedades | Lote | `node scripts/ncr-lote.js --periodo=MM/AAAA --nombres="A;B;C"` (ver **Lote**) |

**Plugin instalado:** los scripts viven en la raiz del plugin. Correr cada comando
con `cd "${CLAUDE_PLUGIN_ROOT:-.}" && ...`. La primera vez (y despues de cada
actualizacion del plugin): `npm install && npx playwright install chromium &&
python3 -m pip install -r requirements.txt`. Las credenciales van en
`~/.fisco-ar/.env` (copiar `.env.example`), que sobrevive a las actualizaciones.
El resumen de `ncr-lote.js` y los logs de simulacion de `eliminar-ncr.js` quedan en
`~/Documents/BVA-salidas/` (se cambia con `BVA_SALIDAS_PATH` en el `.env`).

**Orden obligatorio: el paso 4 va despues de la ULTIMA importacion desde ARCA
del periodo.** "Importar desde ARCA" trae todo lo que no esta en el libro, asi
que una importacion posterior vuelve a meter las NCR eliminadas (se arregla
corriendo 3 y 4 de nuevo). Por eso tampoco tiene sentido depurar un mes abierto.

**ARCA admite un solo borrador de Libro IVA por sociedad.** Si hay otro periodo
sin presentar, el script del paso 1 corta con "Tiene un borrador sin presentar,
correspondiente a otro periodo". Descartarlo borra lo cargado en ese mes: nunca
hacerlo sin que el usuario lo autorice expresamente para ese periodo.

### Paso 1 — Compras del Portal IVA

El script entra con la clave del Excel de claves, cambia la representada, abre
el borrador, importa desde ARCA y baja el CSV. **La importacion de ARCA es
asincronica**: el script espera a que la tarea figure "Procesada" antes de bajar.
Si en algun momento se usa un CSV bajado de otra forma, controlar que las fechas
cubran el mes entero: un CSV bajado antes de tiempo sale cortado (paso en 08/2026:
salio con dos tercios de los comprobantes, hasta el dia 20) y la
conciliacion da cientos de "NCR no encontradas" falsas.

Dejar el CSV en la carpeta `ncr/` del periodo.

Tambien sirve el Excel del plugin `portal-iva-descarga` (`AAAAMM - PORTAL IVA - COMPRAS.xlsx`
en la carpeta IVA): `conciliar-ncr.py` lo lee igual que el CSV y corta en la fila en
blanco antes del TOTAL. Si la sociedad usa carpeta de año (`01-Impuestos Mensuales/2026/AAAAMM`),
los scripts usan el mes que esta adentro.

### Paso 2 — NCR del SharePoint

Las droguerias las suben al SharePoint de las farmacias:
`soytufarmaciams.sharepoint.com/sites/EstudioContable` → Documentos →
`NC por NRF/<AAAA>/<MES>/SUIZO/` (un archivo del grupo) y `.../MONROE/` (~64,
uno por farmacia de todas las sociedades). **Solo se entra desde el perfil de
Chrome "Trabajo" (estudiobva)**, como invitado con la cuenta de Google que
autorizaron las farmacias: no hay claves para loguear un script. El SharePoint
es del cliente: solo se lee.

1. En la pestaña del SharePoint, con Claude in Chrome, correr
   `scripts/sharepoint/bundle-ncr.js` con `window.__NCR_PERIODO = 'AAAAMM'`.
   Baja Suizo y Monroe del mes en **una** descarga (Chrome frena las multiples):
   `~/Downloads/SP-NCR-AAAAMM.bundle`. Un paquete sirve para todas las sociedades.
2. `node scripts/descargar-ncr.js --nombre="<Sociedad>" --periodo=MM/AAAA`

`descargar-ncr` elige los archivos de Monroe **por el CUIT que trae cada archivo
en la columna `Cuit`**, no por el nombre (cambia de un mes a otro: `F ALEMANA` en
julio, `FRANCO ALEMANA ` en agosto). Resuelve el `CLIENTE` de Suizo de cada
farmacia (igual, o por `lib/ncr-alias.json`), controla los pies y guarda en
`ncr/` sin pisar. Su campo `clientes_suizo` es el ultimo argumento del script de
conciliacion. Si informa `sin_resolver_en_suizo`, **preguntar** a que `CLIENTE`
corresponde y agregarlo a `ncr-alias.json` recien con la respuesta.
`--solo-revisar` no escribe nada; `--actualizar` reemplaza lo que cambio y deja
lo anterior en `ncr/_anteriores/`.

Si no existe la carpeta del mes en el SharePoint, las droguerias todavia no
subieron las NCR: frenar.

### Paso 4 — Eliminar las NCR del Libro Compras

Es el paso que mas tiempo le lleva al estudio a mano (cientos de NCR por mes
en las sociedades grandes). `eliminar-ncr.js` toma la hoja **Eliminados** de la salida de esta
skill y las saca del borrador del Libro Compras de a una, con el mismo pedido
que el tachito rojo de cada fila (`ajax.do?f=eliminarComprobante&id=<idReg>`).
Nunca usa "Eliminar todos" ni presenta nada.

```
node scripts/eliminar-ncr.js --nombre="<Sociedad>" --periodo=MM/AAAA                          # simulacion
node scripts/eliminar-ncr.js --nombre="<Sociedad>" --periodo=MM/AAAA --ejecutar --limite=5    # piloto
node scripts/eliminar-ncr.js --nombre="<Sociedad>" --periodo=MM/AAAA --ejecutar               # el resto
```

Siempre en ese orden, y **pedir confirmacion explicita al usuario antes de cada
corrida con `--ejecutar`**. Con el navegador visible si el usuario lo pide
(`PLAYWRIGHT_HEADLESS=false`).

- Elimina una NCR solo si en el libro hay **exactamente una** fila con el mismo
  tipo (3), CUIT emisor, PV y numero, y coinciden importe y fecha. Lo demas se
  informa (`diferencia`, `ambigua`) y no se toca.
- "NC sin respaldo" y "NCR no encontradas" no se eliminan nunca.
- Controla antes que el libro este en Borrador, sea del CUIT y periodo correctos
  y que la grilla de ARCA siga con la estructura conocida. Los libros grandes
  (~5.000+ comprobantes) ARCA los pagina desde el servidor: el script los lee
  por tandas y valida contra el "de un total de N" de la pantalla.
- Al terminar recarga el libro desde ARCA y verifica. El cierre esperado es
  `cuadraConConciliacion: true`: el libro queda con la misma cantidad y el mismo
  total al centavo que la hoja `Compras S.NCR`.
- Es reanudable: las ya eliminadas salen como `no_esta_en_libro`. El log queda
  en `ncr/[AAAAMM]-ELIMINACION-NCR-[fecha-hora].csv`, archivo nuevo cada vez.

La simulacion tiene que dar 0 `diferencia` y 0 `ambigua`. Si no, reportar y no
ejecutar.

## Donde estan los archivos

Todo se lee del Shared Drive de sociedades: via el conector de Google Drive en
Cowork, o directo del disco (`I:\Unidades compartidas\...`, la letra puede
cambiar) en Claude Code. No se piden archivos adjuntos.

```
[N - Sociedad - CUIT]/01-Impuestos Mensuales/[AAAAMM]/[AAAAMM]-IVA/ncr/
```

La carpeta `ncr/` no esta siempre en el mismo lugar: segun la sociedad es
`[AAAAMM]/[AAAAMM]-IVA/ncr`, `[AAAAMM]-IVA/ncr`, `[AAAAMM]/IVA/ncr` o
`[AAAAMM]/ncr` (algunas se movieron a mano asi en 08/2026). Los
scripts usan la que exista (`carpetaNcr()` en `lib/rutas-bva.js`) y cortan si hay
mas de una.

Dentro de esa carpeta `ncr/` conviven las tres fuentes:

1. **Compras del Portal IVA** — un .zip tipo
   `comprobantes_periodo_[AAAAMM]_compras_[AAAAMMDD]_[HHMM].zip`, con un unico
   CSV adentro, o el CSV suelto (`[AAAAMM] - PORTAL IVA - COMPRAS.csv`, el que
   deja el paso 1). Puede estar en `ncr/` o un nivel arriba, en `[AAAAMM]-IVA/`.
2. **NCR de Suizo** — `[AAAAMM]-SUIZO NCR.xlsx`. Es un archivo del grupo
   entero, no de la sociedad.
3. **NCR de Monroe** — **un archivo .xlsx por farmacia**, nombrado con el
   nombre de la farmacia (`ARENALES.xlsx`, `VITAL.xlsx`, `FRANCO ALEMANA.xlsx`).

Si falta alguna de las tres, frenar y decir cual falta. No conciliar contra una
sola drogueria salvo que el usuario lo pida expresamente.

## Una sociedad = varias farmacias

Esto es lo que mas confunde. **Monroe exporta por farmacia, no por sociedad.**
Una sociedad puede tener una docena de archivos de Monroe, uno por sucursal,
todos con el mismo CUIT: son sucursales de la misma sociedad. Conciliar una sociedad exige **unir todos los archivos de
farmacia que haya en su carpeta**.

Suizo, en cambio, trae **una sola hoja** con las ~63 farmacias del grupo juntas
y una columna `CLIENTE` para separarlas. Hay que **filtrar por las farmacias de
esta sociedad**. El nombre de cada archivo de Monroe da la lista; el valor de
`CLIENTE` en Suizo a veces esta abreviado:

| Archivo Monroe | `CLIENTE` en Suizo |
|---|---|
| FRANCO ALEMANA | `F ALEMANA` |
| resto | igual al nombre del archivo |

Vocabulario completo de `CLIENTE` en el archivo de Suizo, para resolver
cualquier sociedad: `ARENALES, ARGENTINA, ARTESANAL, AV MAYO, BELGRANO,
BERAZATEGUI, BERNAL, CALLAO, CASEROS, CONGRESO, DIAMANDY, DLF, EL ALTO,
F ALEMANA, LAIGLON, LUNA, MAGA SHOP, MEGASOL, N GENERACION, N VENTURA, NORTE,
NUEVA 50, NUEVA NORTE, OBRAS, Q FACTORY, Q OESTE, RECOVA, SAINT ETIENNE,
SALUD, SAN NICOLAS, SANAR 1, SANAR 10, SANAR 11, SANAR 12, SANAR 2, SANAR 3,
SANAR 4, SANAR 5, SANAR 52, SANAR 59, SANAR 6, SANAR 60, SANAR 61, SANAR 62,
SANAR 63, SANAR 64, SANAR 65, SANAR 68, SANAR 7, SANAR 8, SANAR 9, SOCIAL
ONCE, SOL WILDE, SOLANO, SOY DEL SOL, SOY DEL SUR, SOY FAR, SOY QUILMES,
TOFANELLI, VARELA, VITAL, ZAPIOLA, ZEUS`

Si un archivo de Monroe no resuelve contra ningun `CLIENTE`, **preguntar** a
que farmacia corresponde. No adivinar por parecido de texto.

## Las claves de matcheo (validadas, no suponer otras)

En los dos casos se matchea contra el CSV del Portal IVA por
`Tipo de Comprobante` = **3** (Nota de Credito A), y se compara **PV y numero
como enteros** — el CSV no trae ceros a la izquierda, asi que padear rompe el
match.

| Drogueria | Punto de Venta | Numero | CUIT emisor a validar |
|---|---|---|---|
| **Monroe** | 4 digitos antes del guion de `Numero Formateado` (3116) | 8 digitos despues del guion | `30517059095` |
| **Suizo** | columna `terminal` (503) | columna `numero` | `30516968431` |

Validar el CUIT del emisor no es opcional: LUMARLE SA tambien emite notas de
credito tipo 3 que **no** entran en esta conciliacion.

Una vez matcheado, controlar que importe y fecha coincidan. En la corrida de
referencia coincidieron en el 100%; cualquier diferencia es una senal de alarma
que hay que reportar, no absorber.

## Trampas de parseo

Cada fuente tiene los tipos al reves de lo que uno espera. Estas cinco
explican casi todos los errores:

- **CSV del Portal IVA**: encoding **ISO-8859-1** (no UTF-8), separador `;`,
  decimal **coma** y punto de miles. Fecha en **ISO `YYYY-MM-DD`**.
- **`Importe Total` de Monroe es TEXTO con punto decimal y sin separador de
  miles** (`-79047.4`). Castear con `float(x)` directo: aplicarle el criterio
  argentino (punto = miles) infla el total unas 90 veces.
- **`fecha` de Suizo es TEXTO `DD/MM/YYYY`**; su `total` si es float real.
- **Los archivos de Monroe traen una fila vacia penultima y una fila de pie**
  (`Cantidad de Registros: N >>> Fin de la consulta <<<`). Filtrar por
  `Tipo Linea == 'Detalle'`, que es lo robusto.
- La `Fecha` de Monroe es datetime real pero se **muestra** como `mm-dd-yy`. No
  leerla a ojo del Excel.

## Los cuatro controles

Correrlos siempre y reportarlos aunque no los pidan.

1. **Contador del pie de cada archivo de Monroe.** El pie declara cuantos
   registros devolvio la consulta DMA. Si no cierra contra las filas leidas, la
   exportacion salio truncada y hay que rehacerla. En la corrida de referencia
   fallaron 5 de 12 archivos (faltaban 21 filas). Listar archivo por archivo
   con el delta. El truncado viene **de origen**: el archivo del SharePoint es
   identico (verificado 22/09/2026), asi que hay que pedir que la farmacia o
   Monroe lo reexporten; no sirve volver a bajarlo.
2. **NCR no encontradas en el Portal IVA.** Revisar la fecha antes de
   sospechar del proceso: una NCR con fecha del mes siguiente simplemente no
   pertenece a este periodo (asi se explicaron las 2 de la corrida de
   referencia). Nunca eliminar a ciegas.
3. **NC de las droguerias sin respaldo en NCR.** Es el hallazgo mas importante
   y el mas frecuente: notas de credito tipo 3 de Suizo o Monroe que estan en
   el Portal IVA pero en ningun archivo NCR (en la corrida de
   referencia fueron cientos de comprobantes). **No se eliminan** — se informan con detalle e
   importe total para que el usuario decida. La causa probable es que las
   exportaciones NCR salen filtradas por concepto (Monroe: solo
   `NC Dtos Acuerdos OS`; Suizo: solo bonos PAMI/Farmalink/Preserfar), asi que
   habria otras categorias de NCR sin exportar. Antes de darlo por perdido,
   verificar que no aparezcan bajo otro `CLIENTE` de Suizo.
4. **Verificacion aritmetica.** `total Portal IVA − total eliminados = total
   resultante`, recalculado sobre las filas que quedaron, no arrastrado. Si no
   cierra al centavo, no entregar el archivo.

## Criterio de eliminacion (decidido, no cambiar sin pedirlo)

**Se eliminan solo las NCR efectivamente conciliadas.** Todo comprobante que no
pudo probarse contra un archivo NCR se queda en compras y se informa. Es
deliberadamente conservador: el proceso nunca saca algo que no pudo demostrar.
Por eso `conciliar-ncr.py` tampoco elimina una NCR que matchea por clave pero con
otro importe o fecha (el script embebido la elimina y solo la informa): va a la
hoja `Diferencias` y se queda en compras.

## Salida

Archivo **nuevo** en la misma carpeta `ncr/`, sin tocar ninguna de las fuentes:

```
[AAAAMM]-COMPRAS-SIN-NCR-[SOCIEDAD].xlsx
```

Con estas hojas:

| Hoja | Contenido |
|---|---|
| `Compras S.NCR` | el CSV del Portal IVA menos las NCR conciliadas |
| `Eliminados` | las NCR conciliadas, con la drogueria y la farmacia de origen |
| `NCR no encontradas` | NCR de los archivos que no aparecen en el Portal IVA |
| `NC sin respaldo` | NC tipo 3 de Suizo/Monroe en el Portal sin NCR que las respalde |
| `Diferencias` | solo si hay: NCR que matchean pero con otro importe o fecha, repetidas, o contra mas de una fila del Portal. **No se eliminan** |
| `Control` | totales, verificacion aritmetica, fuentes usadas y pie de cada archivo de Monroe |

`eliminar-ncr.js` lee solo `Eliminados` y `Compras S.NCR` (via
`leer_conciliacion_ncr.py`): no cambiar esos nombres ni las columnas `_drog`,
`_farmacia`, `_cuit`, `_pv`, `_nro`, `_imp`, `_fecha`, `portal_importe`,
`portal_fecha`.

## Script

**En Claude Code, con los scripts del estudio: usar `scripts/conciliar-ncr.py`.**
Encapsula las reglas validadas (da identico al script de abajo contra
las sociedades conciliadas en 08/2026) y ademas:

- Ubica solo la carpeta `ncr/` y las compras (CSV, zip o Excel del Portal IVA;
  busca en `ncr/`, en la carpeta IVA y en la del mes). `--portal=` para forzar uno.
- Arma los `CLIENTE` de Suizo desde los archivos de Monroe + `lib/ncr-alias.json`,
  y corta si alguno no resuelve (`--clientes=` para forzarlos).
- Controla que cada archivo de Monroe sea del CUIT de la sociedad.
- Escribe los importes como numeros y las fechas como fechas (el script de abajo
  los deja como texto con coma cuando la fuente es el CSV).
- Escribe en un temporal, relee el Excel y controla cantidad y total antes de
  dejarlo en `ncr/`. Si la verificacion no da 0, no escribe nada.
- No pisa: si ya hay una conciliacion, corta; `--si-existe=reemplazar` manda la
  anterior a `ncr/_anteriores/`, `--si-existe=saltear` la deja.
- Solo necesita openpyxl (no pandas).

```
python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=MM/AAAA --solo-revisar   # concilia y reporta, no escribe
python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=MM/AAAA                  # deja la salida en ncr/
```

Devuelve un JSON con los mismos campos del formato de respuesta (totales,
`verificacion`, `ncr_no_encontradas_detalle`, `diferencias_detalle`,
`nc_sin_respaldo_por_drogueria`, `suizo_sin_respaldo_bajo_otro_cliente`,
`pies_archivos_monroe`, `avisos`).

**En Cowork, sin los scripts:** escribir este a disco y correrlo. No
reimplementar la logica a mano.

```python
# reconciliar_ncr.py <portal.csv> <suizo.xlsx> <dir_monroe> <salida.xlsx> <CLIENTE1,CLIENTE2,...>
import sys, json, glob, os, re, warnings
import pandas as pd
warnings.filterwarnings('ignore')

CUIT_MONROE, CUIT_SUIZO, TIPO_NC_A = '30517059095', '30516968431', 3
portal_csv, suizo_xlsx, dir_monroe, salida = sys.argv[1:5]
clientes_suizo = [c.strip().upper() for c in sys.argv[5].split(',') if c.strip()]

# --- Portal IVA -------------------------------------------------------------
p = pd.read_csv(portal_csv, sep=';', encoding='ISO-8859-1', dtype=str).fillna('')
num = lambda s: pd.to_numeric(s.str.replace('.','',regex=False).str.replace(',','.',regex=False),
                              errors='coerce').fillna(0.0)
p['_imp']   = num(p['Importe Total'])
p['_tipo']  = pd.to_numeric(p['Tipo de Comprobante'], errors='coerce')
p['_pv']    = pd.to_numeric(p['Punto de Venta'], errors='coerce')
p['_nro']   = pd.to_numeric(p['Número de Comprobante'], errors='coerce')
p['_cuit']  = p['Nro. Doc. Vendedor'].str.replace(r'\D','',regex=True)
p['_fecha'] = pd.to_datetime(p['Fecha de Emisión'], errors='coerce')
total_portal = round(p['_imp'].sum(), 2)

# --- NCR Monroe: un archivo por farmacia -----------------------------------
mon, pies = [], []
for f in sorted(glob.glob(os.path.join(dir_monroe, '*.xlsx'))):
    base = os.path.splitext(os.path.basename(f))[0].upper()
    if 'SUIZO' in base or base.startswith('~$'):
        continue
    d = pd.read_excel(f, dtype=str).fillna('')
    if 'Tipo Linea' not in d.columns:
        continue
    pie = d[d['Tipo Linea'].str.contains('Cantidad de Registros', na=False)]
    declarado = int(re.search(r'(\d+)', pie.iloc[0, 0]).group(1)) if len(pie) else None
    d = d[d['Tipo Linea'] == 'Detalle'].copy()
    pies.append({'archivo': base, 'leidas': len(d), 'declaradas': declarado,
                 'ok': declarado is None or declarado == len(d)})
    partes = d['Numero Formateado'].str.split('-', n=1, expand=True)
    d['_pv']  = pd.to_numeric(partes[0], errors='coerce')
    d['_nro'] = pd.to_numeric(partes[1], errors='coerce')
    d['_imp'] = pd.to_numeric(d['Importe Total'], errors='coerce')   # punto decimal
    d['_fecha'] = pd.to_datetime(d['Fecha'], errors='coerce')
    d['_drog'], d['_farmacia'], d['_cuit'] = 'MONROE', base, CUIT_MONROE
    mon.append(d)
mon = pd.concat(mon, ignore_index=True) if mon else pd.DataFrame()

# --- NCR Suizo: una hoja, filtrar por CLIENTE ------------------------------
s = pd.read_excel(suizo_xlsx, sheet_name=0)
s = s[s['CLIENTE'].astype(str).str.strip().str.upper().isin(clientes_suizo)].copy()
s['_pv']  = pd.to_numeric(s['terminal'], errors='coerce')
s['_nro'] = pd.to_numeric(s['numero'], errors='coerce')
s['_imp'] = pd.to_numeric(s['total'], errors='coerce')
s['_fecha'] = pd.to_datetime(s['fecha'], format='%d/%m/%Y', errors='coerce')
s['_drog'], s['_farmacia'], s['_cuit'] = 'SUIZO', s['CLIENTE'], CUIT_SUIZO

cols = ['_drog', '_farmacia', '_cuit', '_pv', '_nro', '_imp', '_fecha']
ncr = pd.concat([mon[cols] if len(mon) else pd.DataFrame(columns=cols), s[cols]],
                ignore_index=True)

# --- matcheo: tipo 3 + pv + nro, acotado al CUIT del emisor ----------------
idx = {}
for i, r in p[p['_tipo'] == TIPO_NC_A].iterrows():
    idx.setdefault((r['_cuit'], r['_pv'], r['_nro']), []).append(i)

filas_elim, elim, no_enc, difs = set(), [], [], []
for _, r in ncr.iterrows():
    hit = idx.get((r['_cuit'], r['_pv'], r['_nro']))
    if not hit:
        no_enc.append(r.to_dict()); continue
    i = hit[0]; filas_elim.add(i)
    reg = dict(r)
    reg['portal_importe'], reg['portal_fecha'] = p.at[i, '_imp'], p.at[i, '_fecha']
    if round(abs(p.at[i, '_imp'] - r['_imp']), 2) > 0.01 or p.at[i, '_fecha'] != r['_fecha']:
        difs.append(reg)
    elim.append(reg)

# --- NC de las droguerias sin respaldo: se informan, NO se eliminan --------
mask = (p['_tipo'] == TIPO_NC_A) & p['_cuit'].isin([CUIT_MONROE, CUIT_SUIZO])
sin_respaldo = p[mask & ~p.index.isin(filas_elim)]

resto = p[~p.index.isin(filas_elim)]
total_elim = round(p.loc[sorted(filas_elim), '_imp'].sum(), 2)
total_resto = round(resto['_imp'].sum(), 2)

drop = [c for c in p.columns if c.startswith('_')]
with pd.ExcelWriter(salida, engine='openpyxl') as w:
    resto.drop(columns=drop).to_excel(w, sheet_name='Compras S.NCR', index=False)
    pd.DataFrame(elim).to_excel(w, sheet_name='Eliminados', index=False)
    pd.DataFrame(no_enc).to_excel(w, sheet_name='NCR no encontradas', index=False)
    sin_respaldo.drop(columns=drop).to_excel(w, sheet_name='NC sin respaldo', index=False)

print(json.dumps({
  'portal_comprobantes': len(p), 'portal_total': total_portal,
  'ncr_leidas': len(ncr), 'eliminados': len(filas_elim), 'eliminados_total': total_elim,
  'resultantes': len(resto), 'resultantes_total': total_resto,
  'verificacion': round(total_portal - total_elim - total_resto, 2),
  'ncr_no_encontradas': len(no_enc), 'difs_importe_o_fecha': len(difs),
  'nc_sin_respaldo': len(sin_respaldo),
  'nc_sin_respaldo_total': round(sin_respaldo['_imp'].sum(), 2),
  'pies_archivos_monroe': pies,
}, indent=2, default=str))
```

`verificacion` tiene que dar **0.0**. Si no, no entregar el archivo.

## Lote

Para varias sociedades del mismo periodo, una vez armado el paquete del
SharePoint:

```
node scripts/ncr-lote.js --periodo=MM/AAAA --nombres="<sociedad A>;<sociedad B>"
node scripts/ncr-lote.js --periodo=MM/AAAA --cuits=<CUIT A>,<CUIT B> --sin-simular
```

Por sociedad corre compras del Portal IVA → `descargar-ncr` → `conciliar-ncr` →
`eliminar-ncr` en **simulacion**, y devuelve una tabla con lo que se eliminaria, los hallazgos y el
comando exacto para eliminar. **Nunca elimina** (rechaza `--ejecutar`): la
eliminacion sigue siendo de a una sociedad y con confirmacion del usuario.

- **Compras (paso 1):** primero las busca en la carpeta del periodo, con la misma
  regla que `conciliar-ncr`. El equipo suele bajar el Portal IVA y liquidar IVA
  semanas despues, asi que si estan se usan y no se entra a ARCA. Se mira el dia
  en que se bajo el archivo: durante el mes del periodo o hasta el dia 5 del mes
  siguiente es **parcial**; despues del 5 es el mes cerrado.
  - Parcial y es `AAAAMM - PORTAL IVA - COMPRAS.xlsx` (el del script): se vuelve a
    bajar y se reemplaza.
  - Parcial con otro nombre (lo subio el equipo): la sociedad queda con error y el
    archivo no se toca. Borrarlo o renombrarlo, o `--aceptar-compras-parciales`.
  - Si faltan, se bajan con `portal-iva-descarga.js --libros=compras`, salvo que la
    DDJJ ya este presentada o se pase `--no-bajar-compras`.
  - La columna `compras_origen` del resumen dice de donde salieron y cuando se bajaron.
- Reconcilia solo si `descargar-ncr` trajo algo nuevo o reemplazado (o con
  `--reconciliar`); si no, usa la conciliacion que ya esta.
- No entra a ARCA con una sociedad si en la carpeta del mes hay un acuse de IVA
  del periodo (`AAAAMM-IVA-ACUSE-*.pdf`): con la DDJJ presentada, entrar por
  "Nueva declaracion jurada" puede abrir una rectificativa.
- Si ARCA rechaza una clave o pide captcha, no entra mas a ARCA en ese lote
  (varias sociedades comparten apoderado y reintentar bloquea la clave).
- Headless por defecto; `--ver` para el navegador visible.
- Deja `~/Documents/BVA-salidas/ncr-lote/AAAAMM-<fecha-hora>.json` y `.csv`.

## Flujo

1. Identificar sociedad y periodo. Confirmar si el nombre es ambiguo — hay
   confusiones frecuentes: "Soy Maga Mega Sol" -> **Megasol**, "Sanar 51" ->
   **Solano 51**, "Soy 60" -> **Sanar 60**, "Sanar 57" -> suele ser **Sanar 68**.
2. Si faltan fuentes y hay scripts disponibles, correr los pasos 1 y 2 del
   circuito (compras y NCR). Si ya hay otro borrador abierto en ARCA, frenar y
   preguntar.
3. Ubicar la carpeta `ncr/` y listar sus archivos. Identificar el CSV/zip de
   compras, el de Suizo, y los de farmacia de Monroe.
4. En Cowork, bajar los archivos y descomprimir el zip. Para bajar de Drive usar
   `download_file_content` (devuelve base64); **no** usar `read_file_content`,
   que trunca los .xlsx en silencio. En Claude Code se leen directo del disco.
5. Armar la lista de `CLIENTE` de Suizo: la da `clientes_suizo` de
   `descargar-ncr`, o si no a partir de los nombres de archivo de Monroe (con la
   tabla de alias). Reportar los que no resuelvan.
6. Correr el script: `scripts/conciliar-ncr.py` en Claude Code (primero
   `--solo-revisar`), o el embebido en Cowork (necesita pandas y openpyxl).
7. Dejar la salida en la misma carpeta `ncr/` y reportar.
8. Si el usuario quiere depurar el Libro Compras, paso 4 del circuito:
   simulacion → confirmacion → piloto de 5 → revision del usuario en ARCA →
   confirmacion → el resto. Reportar el cierre contra `Compras S.NCR`.

## Formato de respuesta

```
Listo, **[Sociedad] [AAAAMM]** conciliada.

| Concepto | Comprobantes | Importe |
|---|---|---|
| Compras Portal IVA | X | $X |
| NCR conciliadas (eliminadas) | X | $X |
| Compras S.NCR | X | $X |
| **Verificación aritmética** | | ✅ **$0,00** |

NCR leídas: X (Monroe X en Y farmacias + Suizo X)

**Hallazgos:**
- ✅/⚠️ NCR no encontradas en el Portal IVA: X [detalle con fecha — si son del mes siguiente, decirlo]
- ✅/⚠️ NC de Suizo/Monroe sin respaldo en NCR: X por $X [NO eliminadas]
- ✅/⚠️ Exportaciones de Monroe truncadas: [archivo: leídas vs declaradas]
- ✅/⚠️ Diferencias de importe o fecha en las conciliadas: X
```

Ojo con el signo: los importes eliminados son negativos, asi que sacarlos
**aumenta** el total de compras. No es un error.

Despues de eliminar en el Portal IVA (paso 4):

```
**Libro Compras [Sociedad] [AAAAMM] depurado.**

| | Comprobantes | Total |
|---|---|---|
| Libro antes | X | $X |
| NCR eliminadas | −X | |
| **Libro ahora** | **X** | **$X** |
| Hoja Compras S.NCR | X | $X |

✅/⚠️ Cierre contra la conciliacion (cantidad y total al centavo) · ninguna eliminada sigue en el libro
Log: ncr/[AAAAMM]-ELIMINACION-NCR-[fecha-hora].csv · el borrador queda sin presentar
```

## Limitaciones conocidas

- **El CSV del Portal IVA no trae CUIT del receptor.** No hay forma de
  verificar desde el archivo que las compras sean de esta sociedad; depende de
  con que cuenta fiscal se exporto. Si hace falta certeza, cruzar contra Mis
  Comprobantes Recibidos, que si trae el CUIT del comprador.
- **Las NC de Monroe sin respaldo no se pueden atribuir a una farmacia**: el
  CSV del Portal no trae `Cod Cliente`.
- Los codigos internos no son intercambiables: el `ctacte` de Suizo (1081, 39,
  49...) y el `Cod Cliente` de Monroe (285405, 285408, 285409...) son sistemas
  de codificacion distintos.
- Si en algun mes el archivo de Suizo viniera con **varias hojas** en vez de
  una, hay que ajustar el parseo: el script lee la primera hoja.