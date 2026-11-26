# portal-iva-vulcamoia

Skill de Melina, pasada a Claude Code. Toma los zips del Libro IVA Digital que se bajan del Portal
IVA de ARCA (`libro_iva_periodo_AAAAMM*.zip`) y llena las plantillas `input_portal COMPRAS` e
`input_portal VENTAS` de la sociedad: detalle de comprobantes, clasificación de cada proveedor con
el padrón "Clasificacion Proveedores", Resumen por Clasificación mensual y Control con Portal IVA
contra el `libro_iva.xml`. Pensada para Vulcamoia; sirve para otra sociedad con la misma
estructura de carpeta y plantillas. No entra a ARCA.

## Instalación

En las máquinas del equipo lo instala `setup-bva`. A mano:

```
/plugin install portal-iva-vulcamoia --marketplace estudiobva/estudio-claude-marketplace
```

```bash
python3 -m pip install --user -r requirements.txt
cp .env.example .env   # opcional: ruta de la unidad compartida
```

## Uso

```bash
python3 scripts/portal-iva-vulcamoia.py procesar --solo-revisar
python3 scripts/portal-iva-vulcamoia.py procesar
python3 scripts/portal-iva-vulcamoia.py procesar --si-existe=reemplazar
python3 scripts/portal-iva-vulcamoia.py procesar --sociedad="<nombre o CUIT>" --anio=2026
python3 scripts/portal-iva-vulcamoia.py clasificar --json=sugerencias.json --modo=sugerida
python3 scripts/portal-iva-vulcamoia.py clasificar --json=sugerencias.json --modo=confirmada
```

Lee de `<sociedad>/04-Balances y EECC/<año>/PORTAL IVA` y deja ahí
`input_portal_COMPRAS Completo.xlsx` e `input_portal_VENTAS Completo.xlsx`. No pisa nada: con
`--si-existe=reemplazar` los anteriores van a `_anteriores/`. El padrón solo cambia con
`clasificar --modo=confirmada`, que antes guarda una copia en `_anteriores/`.

Los controles (Control con Portal IVA y Comprobación de VENTAS) se calculan en Python y salen en
el JSON; los libros quedan con recálculo al abrir en Excel. No hace falta LibreOffice.

## Estructura

- `scripts/portal-iva-vulcamoia.py`: punto de entrada (carpeta, archivos, controles, guardado).
- `scripts/procesar_portal_iva.py`: el armado de COMPRAS y VENTAS (script original de la skill).
- `scripts/agregar_clasificaciones.py`: aplica clasificaciones sugeridas o confirmadas (original).
- `skills/portal-iva-vulcamoia/SKILL.md`: la habilidad.
