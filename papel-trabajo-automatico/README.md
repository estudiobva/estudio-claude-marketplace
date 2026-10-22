# papel-trabajo-automatico

Arma el papel de trabajo unificado de liquidaciones (`AAAA00-liquidaciones-IVA, IIBB y Tasa
Municipal-<SOC>.xlsx`) de los meses **ya presentados**, desde las fuentes oficiales, y lo
guarda **aparte** como `... - AUTOMATICO.xlsx`, al lado del papel del equipo (que no se toca).

Molde: el papel de la propia sociedad (se copia y se reemplazan valores; por mes se ubican
los bloques por etiqueta, se agrandan si hace falta y se regeneran las fórmulas). Si la
sociedad no tiene papel unificado, usa el de otra del mismo municipio, vaciado.

| Bloque | Fuente |
|---|---|
| Ventas por tipo de comprobante | `AAAAMM - PORTAL IVA - VENTAS.xlsx` de la carpeta IVA (plugin `portal-iva-descarga`). |
| Compras por tipo y clasificación | `AAAAMM - PORTAL IVA - COMPRAS.xlsx` + Maestro de Proveedores (CMV → MERCADERIA; el resto → SERVICIOS AL x%). |
| IVA: retenciones y saldos anteriores | DDJJ F.2051 presentada (plugin `ddjj-iva`, o el PDF `*IVA*DDJJ*` de la carpeta del mes). |
| IIBB ARBA | DJ de anticipo (plugin `constancias-anticipo-arba`, o el PDF `*IIBB*DJ*` de la carpeta del mes). |
| Seguridad e Higiene | Base = ventas gravadas; alícuota y concepto fijo del papel del equipo. |
| Publicidad y propaganda | Lo cargado por el equipo (es manual). |

Lo que no tiene fuente queda en **amarillo con un comentario** que dice qué falta.

## Instalacion

```bash
python3 -m pip install -r requirements.txt
```

## Uso

```bash
python3 scripts/generar_papel.py --cuit=<CUIT> --hasta=202609
python3 scripts/generar_papel.py --todas --hasta=202609 --dir=<carpeta>
```

Detalle en `skills/papel-trabajo-automatico/SKILL.md`.
