---
name: saldos-farmacias-iva-iibb
description: "Arma los saldos de IVA (F.2051 de ARCA) e IIBB (DJ anticipo de ARBA R-606M, SICOL de AGIP o CM03 de Convenio Multilateral) del mes de cada sociedad para la planilla FARMACIAS SALDO del SharePoint (columnas C a I: saldo tecnico, SLD, SLD utilizado, pagos IVA, IIBB provincia, IIBB CABA, pagos IIBB), con controles, y prepara las DDJJ para subirlas a la carpeta 'DDJJ IVA-IIBB MM-AAAA' del SharePoint. Usar SIEMPRE que pidan \"cargar saldos de IVA e IIBB\", \"completar FARMACIAS SALDO\", \"pasar los saldos de [sociedad]\", \"saldos del mes\", \"subir las DDJJ de IVA e IIBB al SharePoint\" o \"avanza con [sociedad]\" en esta tanda. Solo lee la unidad: no entra a ARCA ni sube nada sin confirmacion."
---

# Saldos de IVA e IIBB por farmacia

Saca de las DDJJ del mes de cada sociedad los saldos que van en la planilla
`FARMACIAS SALDO MM-AAAA` (SharePoint Estudio Contable, carpeta "saldos y pagos de IVA e
IIBB"), y deja lista la carpeta de DDJJ para subir al mismo lugar.

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/saldos-iva-iibb.py --periodo=MM/AAAA
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/saldos-iva-iibb.py --periodo=MM/AAAA --cuits="<CUIT>;<CUIT>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && python scripts/saldos-iva-iibb.py --periodo=MM/AAAA --valores=<valores.json> --preparar-subida
```

En Mac es `python3` en vez de `python`. Dependencias: `pip install -r requirements.txt`.

| Argumento | Para que |
|---|---|
| `--periodo` | `MM/AAAA` o `AAAAMM`: el periodo de las DDJJ, no el de presentacion. |
| `--cuit` / `--cuits` (con `;`), `--nombre` / `--nombres` | Solo esas sociedades. Sin nada: todas las sociedades (CUIT 30/33) de la unidad. |
| `--valores=<json>` | Lo leido a mano de las DJ que son imagen (ver abajo). |
| `--preparar-subida` | Copia las DJ usadas a `<salida>/DDJJ IVA-IIBB MM-AAAA/`. No sube nada. |
| `--salida` | Default `<BVA_SALIDAS_PATH>/saldos-iva-iibb` (`~/Documents/BVA-salidas/...`). |

Salida: `SALDOS IVA IIBB MM-AAAA.xlsx` (una fila por sociedad con CUIT, C a I, tipo de IIBB,
estado, avisos y los PDF usados) y un JSON por stdout con `faltan_dj`,
`controles_que_no_cierran`, `pendientes` y la carpeta para el SharePoint.

## Paso a paso

1. Correr el script con el periodo. Mostrarle al usuario cuantas sociedades quedaron
   completas, cuales tienen DJ faltante y los avisos (rectificativas, CUIT o periodo que no
   coincide, controles que no cierran).
2. **DJ en imagen.** Las que no tienen texto, y todas las de AGIP y CM03, vienen en
   `pendientes` con el PNG en `<salida>/imagenes/AAAAMM/` y la lista de `campos`. Leer cada
   imagen y escribir `<salida>/valores-AAAAMM.json`:

   ```json
   {"<CUIT>": {"iva":  {"st_arca": 0, "st_contribuyente": 0, "sld_anterior": 0,
                        "retenciones": 0, "sld": 0, "a_pagar": 0},
               "iibb": {"tipo": "ARBA", "saldo_dj": -1000000.00, "saldo_anterior": 5000000.00,
                        "acumulado": -6000000.00}}}
   ```

   - IVA (seccion "Determinacion de la posicion mensual"): saldo tecnico a favor de ARCA,
     saldo tecnico a favor del contribuyente, SLD del periodo anterior neto de usos,
     retenciones/percepciones/pagos a cuenta, SLD a favor del contribuyente del periodo y
     saldo de impuesto a favor de ARCA. Lo que no figura va 0. Si el saldo tecnico es a
     favor del contribuyente, el SLD del periodo es el ultimo renglon (no cortar la imagen antes).
   - ARBA (R-606M): "Saldo de la DJ", "Saldo a favor anterior" y "Saldo acumulado al cierre",
     **con el signo del PDF** (negativo = a favor del contribuyente).
   - AGIP (SICOL, Rubro II): `anticipo`, `retenciones`, `percepciones`, `pagos_a_cuenta`,
     `saldos_a_favor`, `otros_creditos`, `favor_dgr`, `favor_contribuyente`, `total_a_pagar`.
   - CM03 (cuadro "Jurisdicciones incluidas"): `ba_favor`, `ba_pagar`, `caba_favor`,
     `caba_pagar`, `total_a_pagar` y en `otras` las demas jurisdicciones con saldo.
   - El `tipo` de IIBB lo dice el encabezado del PDF, no el nombre del archivo (hay DJ de
     AGIP guardadas como "IIBB-DJ").
   - Importes con punto decimal, sin separador de miles.

   Volver a correr con `--valores`. El script recalcula C a I y los controles: si alguno no
   cierra, revisar la lectura de esa imagen antes de seguir.
3. **Faltantes.** Avisar enseguida que DJ no estan en la unidad para que el usuario las baje;
   esas columnas quedan vacias (no 0). Una sociedad sin DJ ni este mes ni el anterior sale
   con el aviso "ver si va en la planilla".
4. **Cargar la planilla.** Ver "Carga en la planilla".
5. **Subir las DDJJ.** Ver "Subir las DDJJ a SharePoint". Siempre con OK del usuario.

## Que va en cada columna

IVA (F.2051):
- C SALDO TECNICO = saldo tecnico a favor del contribuyente; si es a favor de ARCA, 0.
- D SLD = SLD a favor del contribuyente del periodo; si la DJ termina a pagar, 0.
- E SLD Utilizado Comp. = `min(SLD anterior, max(0, saldo tecnico ARCA - retenciones))`: las
  retenciones se aplican primero.
- F Pagos IVA = saldo de impuesto a favor de ARCA (coincide con el VEP).
- Control: `F - D = saldo tecnico ARCA - SLD anterior - retenciones`. No usar
  "SLD anterior - E = SLD final": las retenciones que sobran se suman al SLD.

IIBB:
- ARBA: acumulado negativo → G = el importe en positivo, I = 0; positivo (a pagar) → G = 0,
  I = el importe. Control: `acumulado = saldo DJ - saldo anterior` (ARBA redondea a 10 centavos).
- AGIP: H = subtotal a favor del contribuyente, I = total a pagar, G = 0 (confirmar que la
  sociedad no presenta tambien en ARBA). Control: anticipo - deducciones = a favor DGR - a
  favor del contribuyente.
- CM03: G = Buenos Aires a favor, H = CABA a favor, I = total a pagar (= VEP). Las demas
  jurisdicciones no tienen columna: van en los avisos.
- Jurisdiccion donde la sociedad no presenta → 0. DJ que falta → vacio.

## Carga en la planilla

La planilla del mes es `FARMACIAS SALDO MM-AAAA` (hoja `MM-AAAA`). Si no existe, se parte
de la copia de la anterior ("... Copiar") renombrando archivo y hoja; si la copia trae
valores del mes anterior en C a I, avisar antes de borrarlos. Puede que alguien ya la este
cargando: mirar la fecha de modificacion y preguntar antes de escribir.

Emparejar SIEMPRE por CUIT (columna B): los nombres de la planilla, de la carpeta y de la DJ
no coinciden. No sobrescribir celdas con valor sin avisar.

1. Conector de Microsoft 365/SharePoint con escritura sobre Excel, si hay.
2. Si no, Excel Online en el navegador (perfil con sesion en el SharePoint):
   - Zoom 50% para ver C a I; ir a cada celda por el Cuadro de nombres (triple clic,
     referencia, Enter, esperar ~3 s).
   - Escribir con pulsaciones de tecla (`key`, ej. `1 8 9 2 6 5 2 7 . 8 0 Tab`), no con
     `type`; un valor por accion con 2 s de espera; la ultima celda con `Return`.
   - Verificar cada fila con una captura; las capturas llegan con atraso.
3. Si nada de eso esta disponible, el Excel de salida ya es la tabla para pegar por CUIT.

## Subir las DDJJ a SharePoint

Destino: `Estudio Contable / Documentos / saldos y pagos de IVA e IIBB / DDJJ IVA-IIBB MM-AAAA`
(carpeta plana, sin subcarpetas por sociedad; los PDF con el mismo nombre que en la unidad).

1. Correr con `--preparar-subida`: deja en `<salida>/DDJJ IVA-IIBB MM-AAAA/` las DJ de IVA e
   IIBB usadas (sin acuses ni VEP).
2. Mostrarle al usuario la carpeta y la cantidad de archivos y **esperar su OK**: es el sitio
   que se comparte con el cliente.
3. Si la carpeta del mes no existe en el SharePoint, crearla (Nuevo → Carpeta). Si existe,
   listarla para no duplicar; si un archivo ya esta, no reemplazarlo sin preguntar.
4. Subir: arrastrar los archivos, o con el navegador ubicar el `input type=file` de la
   biblioteca y pasarle los PDF con la herramienta de subida, en tandas de menos de 10 MB (no
   hacer clic en "Cargar": abre el selector nativo).
5. Listar la carpeta al terminar y comparar contra la cantidad esperada.

Para listar una carpeta sin la interfaz:
`https://<tenant>.sharepoint.com/sites/EstudioContable/_api/web/GetFolderByServerRelativeUrl('/sites/EstudioContable/Shared Documents/saldos y pagos de IVA e IIBB/DDJJ IVA-IIBB MM-AAAA')?$expand=Files&$select=Files/Name`

## Al terminar

Resumen corto: cuantas sociedades completas, que DJ faltan (por impuesto), avisos,
nombre de la planilla, si se cargo y si se subieron las DDJJ (cuantas y a que carpeta).
