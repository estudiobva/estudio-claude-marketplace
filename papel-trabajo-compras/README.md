# papel-trabajo-compras

Completa las compras del mes en el papel de trabajo de cada sociedad, en el mismo archivo
(`[Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM-papeldetrabajo-*.xlsx`). Es el circuito que se
venia haciendo a mano en cada papel de 09/2026 (Soy 59, Soy 50, Tofanelli, Mega Sol, Soy Varela,
Sanar 5, Sanar 7, Sanar 3, Soy 51 Solano), convertido en scripts de Playwright con navegador oculto.

```
node scripts/papel-trabajo-compras.js --nombre="<sociedad>" --periodo=09/2026
```

## Hojas que agrega (antes de `liquidacion`)

| Hoja | Contenido |
|---|---|
| Compras Zetti | Subdiario de IVA Compras de Zetti (5.6.4), tal cual. |
| Compras Portal IVA | Libro Compras actual de ARCA (sin NCR) clasificado con el Maestro (AG:AL) y, al final, el cuadro de **Compras** + **Prorrateo IVA Compras de Servicios**. |
| Cruce compras | Zetti vs Portal por CUIT + tipo + letra + PV + numero (con 4 pasos de respaldo), con observaciones. |
| Resumen cruce compras | Estados, control contra las hojas (tiene que dar 0), proveedores con diferencias, hallazgos y metodologia. |
| Clasificacion compras | Por rubro del EERR (Mercaderia / Servicios) y por proveedor, con los pendientes del Maestro en amarillo. |

## Scripts

| Script | Que hace |
|---|---|
| `scripts/papel-trabajo-compras.js` | Orquestador: papel, conciliacion NCR, Portal, Zetti, armado, recalculo con Excel, controles y archivo. |
| `scripts/leer-borrador-compras.js` | Exporta el Libro Compras del borrador actual **sin importar** (o el de la DDJJ presentada). Solo lectura. |
| `scripts/zetti-iva-compras.js` | Copia del de `zetti-subdiario-iva-compras` (5.6.4), con `--no-archivar`. |
| `scripts/armar_papel_compras.py` | Arma las 5 hojas. Corta si queda alguna NCR, si las hojas ya existen o si aparece un tipo de comprobante nuevo. |
| `scripts/verificar_papel.py` | Controles despues del recalculo: si no dan, no se toca la unidad. |
| `scripts/clasificar_compras.py` | Copia de `clasificacion-compras` (Maestro de Proveedores), importada como modulo. |

`lib/` es la misma de los otros plugins (`lib/zetti.js` con el boton "Ingresar" de oct-2026).

## Requisitos

- Windows con Excel (el papel se recalcula por COM para que los controles lean valores reales).
- `.env` con la planilla de claves (ver `.env.example`); `BVA_ARCA_LOGIN_NOMBRE` si alguna sociedad
  no tiene la representada con su propia clave.
- Antes: `cruze-zetti-portaliva-ventas` (papel con `liquidacion`) y `conciliacion-ncr` (NCR eliminadas).

Nunca escribe en ARCA ni en Zetti. La unica escritura en la unidad es el papel (y el Zetti del mes si la
carpeta IVA no lo tenia). Nota: openpyxl descarta las extensiones de validacion de datos del papel al
guardarlo (las validaciones comunes quedan).
