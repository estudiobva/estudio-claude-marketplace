# papel-trabajo-ddjj-iva

Etapa 2 del papel de trabajo mensual de cada sociedad: prepara las compras para presentar la DDJJ de IVA
sobre el **mismo** papel `AAAAMM- papeles de trabajo- <Sociedad>.xlsx` que ya armó la etapa de ventas
(Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva, Ventas MCR y cruce ventas).

Lo hace de punta a punta un script de Playwright, `scripts/papel-ddjj-iva.js`, con login automático (desde Claude:
"preparar el papel de IVA de <Sociedad> 202609"). Nunca escribe en Zetti; en ARCA solo si se pide `--portal=borrador`.
Antes de tocar el papel lo respalda en `BVA-salidas/papel-ddjj-iva/respaldos`.

| Paso | Qué hace | Script |
|---|---|---|
| 1 | Baja de Zetti el Subdiario de IVA Compras (5.6.4) y lo pega como solapa **Compras Zetti** | `bajar-compras-zetti.js`, `copiar_solapa.py` |
| 2 | Baja el Libro IVA Compras del Portal IVA (abierto o DDJJ presentada) y lo pega como **Compras portal iva** | plugin `portal-iva-descarga`, `copiar_solapa.py` |
| 3 | Cruza Zetti vs Portal IVA en la solapa **cruce compras** (sin eliminar nada de Zetti), con observaciones por comprobante | `cruce_compras.py` |
| 4 | Filtra en Compras Zetti las facturas de AySA, EDESUR y METROGAS | `filtrar_servicios.py` |
| 5 | Las agrega al final de Compras portal iva con su alícuota (y deja el control de alícuota debajo de Compras Zetti); las mal cargadas no se toman | `agregar_servicios_portal.py` |
| 6 | Título PROVEEDORES en la columna AG de Compras portal iva | `titulo_proveedores.py` |
| 7 | Clasifica los proveedores con el Maestro de Proveedores (y las cuentas que se pasen con `--cuentas` para los NO CLASIFICADOS) y arma el cuadro resumen de compras (SUMIFS, otros conceptos y control) | `clasificar_proveedores.py`, `asignar_cuentas.py`, `cuadro_compras.py` |

## Requisitos

- `setup-bva` (instala Node/Playwright y openpyxl con `package.json` y `requirements.txt`).
- Plugin `portal-iva-descarga` instalado (paso 2).
- `.env` (o `~/.fisco-ar/.env`) con la planilla de claves: usuario de Zetti y claves de ARCA. Ver `.env.example`.
- Unidad compartida montada (I: por defecto, o `BVA_UNIDAD_PATH`). El Maestro de Proveedores se busca en
  `0 - ESTUDIO - 00000000000/Maestro de Proveedores.xlsx`.
- El papel tiene que estar **cerrado en Excel** al correr cada paso.

## Uso

```
node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609
node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202609 --desde-paso=7 --cuentas="<CUIT proveedor>=HONORARIOS"
node scripts/papel-ddjj-iva.js --cuit=<CUIT> --periodo=202610 --desde-paso=2 --portal=iva
```

- Libro IVA Compras (`--portal`): por defecto el de la DDJJ **presentada** (solo lectura, controlado contra la DDJJ). Si el
  período no está presentado, el script para y pide elegir: `--portal=iva` (el que ya está en la carpeta IVA del mes),
  `--portal=<ruta.xlsx>` o `--portal=borrador` (importa al borrador de ARCA; si ya se eliminaron NCR del borrador, puede
  volver a traerlas).
- NO CLASIFICADOS: para con código 4 y la lista; se retoma con `--desde-paso=7 --cuentas="CUIT=CUENTA;..."` (solo el papel,
  no el Maestro).
- `--subcarpeta=controles` si el papel de la etapa 1 quedó en controles, o `--papel=<ruta>`.
- No corre antes del día 11 del mes siguiente al período (`--antes-del-11` para forzarlo).
- Códigos de salida: 0 OK, 1 falló un paso, 2 papel no listo, 3 captcha/clave de ARCA, 4 falta una decisión, 5 control del cuadro ≥ $1.
- Navegador oculto; `--ver` para verlo.

Cada paso también se puede correr suelto (`python scripts/<x>.py --papel="<papel>"`).
Cada script imprime un JSON con el resultado. El detalle de cada paso y sus reglas está en
`skills/papel-trabajo-ddjj-iva/SKILL.md`.

Piloto: período 202609 (control del cuadro: diferencia −0,23).
