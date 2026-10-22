---
name: cuadro-marcha-prestamos
description: Baja el cuadro de marcha (plan de cuotas / cronograma de amortización) de los préstamos vigentes de TODAS las sociedades del Estudio BVA en todos los bancos, con login automático desde Claves_Bancos.xlsx, y deja por préstamo el PDF + la planilla de cuotas y una planilla resumen con todos los préstamos. Lo corre Claude con el script de Playwright scripts/cuadro-marcha.js (navegador oculto). Usar SIEMPRE que pidan "cuadro de marcha", "plan de cuotas del préstamo", "cronograma del préstamo", "préstamos vigentes de las sociedades", "cuánto deben a los bancos", "deuda bancaria por sociedad" o el detalle de préstamos para una calificación bancaria.
---

# Cuadro de marcha de préstamos

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/cuadro-marcha.js
```

| Flag | Qué hace |
|---|---|
| `--banco` | `todos` (default) o varios separados por coma: hoy `NACION`, `MACRO`, `PATAGONIA`, `PROVINCIA`. |
| `--cuit` / `--nombre` | Solo esa sociedad (nombre de la carpeta de la unidad). |
| `--salida` | Default `~/Documents/BVA-salidas/cuadro-marcha/` (o `BVA_SALIDAS_PATH`). |
| `--unidad` | Archiva en `[Sociedad]/03-Bancos y Conciliaciones/[BANCO]/Prestamos/`. |
| `--relevar` | Guarda texto + captura de la pantalla de préstamos de cada empresa en `~/bva-inspeccion/` (para ajustar un banco). |
| `--ver` | Navegador visible. **Solo si el usuario lo pide.** |

Por banco: entra **una sola vez** (los bancos con login único muestran todas las
sociedades), recorre las empresas, las empareja con las carpetas de la unidad
(`lib/prestamos/sociedades.js`: Maestro + nombre de carpeta; si no puede, quedan
en `sin_sociedad` y se guardan en `_sin_sociedad/`; alias en
`~/.fisco-ar/prestamos-alias.json` = `{ "NOMBRE EN EL BANCO": "CUIT" }`), entra a
Préstamos y por cada préstamo vigente guarda:

- `[sociedad]_[banco]_prestamo-[nro]_cuadro-marcha_[AAAA-MM-DD].pdf` + `.csv` (Macro: plan de cuotas completo).
- Provincia no publica el plan de cuotas: `_detalle_` y `_cuotas_` (histórico) en PDF; el contrato mutuo se intenta y, si el banco falla, queda como aviso.

Al final escribe `cuadro-marcha-<fecha>.json` y `cuadro-marcha-<fecha>.csv` (una fila por préstamo:
banco, sociedad, número, capital adeudado, próxima cuota, vencimiento, cuotas pagas, archivos, avisos/errores).
Nunca pisa archivos; siempre cierra la sesión del banco.

## Estado por banco (09/10/2026)

| Banco | Estado |
|---|---|
| Nación (BNA+) | OK, 52 empresas por "Ambientes". Sin préstamos al 08/10. NUNCA `page.goto` adentro: corta la sesión. |
| Macro | OK, 55 empresas. SALIR en una empresa vuelve al selector; en el selector cierra. Si queda una sesión abierta, el banco bloquea 10 min (`sesion_activa`). Plan de cuotas sin descarga: PDF de la pantalla + CSV. |
| Patagonia | OK, 5 empresas (con CUIT). Una de ellas: `sin_funcionalidad` (no tiene contratada la consulta). |
| Provincia | OK, una sola empresa. Sin plan de cuotas online; contrato mutuo da error 500 del banco. |
| Hipotecario | Falta la navegación. El 08/10 la clave fue rechazada ("Los datos ingresados son incorrectos"): revisar Claves_Bancos.xlsx antes de reintentar. |
| Galicia | Falta la navegación. Login por sociedad; dos sociedades con **clave vencida** el 09/10. |
| BBVA, Comafi, Santander | Sin navegación: BBVA y Comafi bloquean el navegador automatizado (antibots/captcha); Santander pide cambiar la clave. |

Para sumar un banco: `lib/prestamos/<banco>.js` (ver `lib/prestamos/index.js`) y relevar con `--relevar`.

## Errores

Cada error trae `code` y `que_hacer`: `credenciales_invalidas` y `bloqueado` no se reintentan;
`sesion_activa` = esperar 10 minutos; `clave_vencida` la cambia una persona; `sin_conexion` =
se cortó internet o la PC se suspendió (mantenerla despierta en corridas largas);
`sin_relevar` = la pantalla de préstamos tiene algo que el script no sabe leer (se guarda la
captura igual).
