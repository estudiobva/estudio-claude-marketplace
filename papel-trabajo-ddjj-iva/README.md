# papel-trabajo-ddjj-iva

Etapa 2 del papel de trabajo mensual de cada sociedad: prepara las compras para presentar la DDJJ de IVA
sobre el **mismo** papel `AAAAMM- papeles de trabajo- <Sociedad>.xlsx` que ya armó la etapa de ventas
(Ventas Zetti, Ventas Zetti (depurado), Ventas portal iva, Ventas MCR y cruce ventas).

Se corre paso a paso desde Claude ("preparar el papel de IVA de Sanar 1 202609"). Nunca escribe en ARCA ni en
Zetti: solo lee y trabaja sobre el Excel del papel, con respaldo previo.

| Paso | Qué hace | Script |
|---|---|---|
| 1 | Baja de Zetti el Subdiario de IVA Compras (5.6.4) y lo pega como solapa **Compras Zetti** | `bajar-compras-zetti.js`, `copiar_solapa.py` |
| 2 | Baja el Libro IVA Compras del Portal IVA (abierto o DDJJ presentada) y lo pega como **Compras portal iva** | plugin `portal-iva-descarga`, `copiar_solapa.py` |
| 3 | Cruza Zetti vs Portal IVA en la solapa **cruce compras** (sin eliminar nada de Zetti), con observaciones por comprobante | `cruce_compras.py` |
| 4 | Filtra en Compras Zetti las facturas de AySA, EDESUR y METROGAS | `filtrar_servicios.py` |
| 5 | Las agrega al final de Compras portal iva con su alícuota (y deja el control de alícuota debajo de Compras Zetti); las mal cargadas no se toman | `agregar_servicios_portal.py` |
| 6 | Título PROVEEDORES en la columna AG de Compras portal iva | (en la skill) |
| 7 | Clasifica los proveedores con el Maestro de Proveedores y arma el cuadro resumen de compras (SUMIFS, otros conceptos y control) | `clasificar_proveedores.py`, `cuadro_compras.py` |

## Requisitos

- `setup-bva` (instala Node/Playwright y openpyxl con `package.json` y `requirements.txt`).
- Plugin `portal-iva-descarga` instalado (paso 2).
- `.env` (o `~/.fisco-ar/.env`) con la planilla de claves: usuario de Zetti y claves de ARCA. Ver `.env.example`.
- Unidad compartida montada (I: por defecto, o `BVA_UNIDAD_PATH`). El Maestro de Proveedores se busca en
  `0 - ESTUDIO - 00000000000/Maestro de Proveedores.xlsx`.
- El papel tiene que estar **cerrado en Excel** al correr cada paso.

## Uso directo

```
node scripts/bajar-compras-zetti.js --cuit=30708929502 --periodo=202609 --si-existe=reemplazar
python scripts/copiar_solapa.py "<mes>/202609 - ZETTI - COMPRAS.xlsx" "<mes>/202609- papeles de trabajo- Sanar 1.xlsx" "Compras Zetti"
python scripts/cruce_compras.py --papel="<mes>/202609- papeles de trabajo- Sanar 1.xlsx"
python scripts/filtrar_servicios.py --papel="..."
python scripts/agregar_servicios_portal.py --papel="..."
python scripts/clasificar_proveedores.py --papel="..."
python scripts/cuadro_compras.py --papel="..."
```

Cada script imprime un JSON con el resultado. El detalle de cada paso y sus reglas está en
`skills/papel-trabajo-ddjj-iva/SKILL.md`.

Piloto: Sanar 1, período 202609 (control del cuadro: diferencia −0,23).
