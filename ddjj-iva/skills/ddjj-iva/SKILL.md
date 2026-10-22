---
name: ddjj-iva
description: "Lee las DDJJ de IVA (F.2051) YA PRESENTADAS de una o varias sociedades en el Portal IVA de ARCA para un mes o un rango de meses: guarda la vista previa en PDF y los importes de la determinacion (debito fiscal, credito fiscal, saldo tecnico anterior, saldo de libre disponibilidad anterior, retenciones, saldos del periodo) y, si se pide, archiva el PDF en la carpeta IVA del mes. Login automatico con la planilla de claves, navegador oculto, solo lectura. Usar SIEMPRE que pidan \"traer las DDJJ de IVA presentadas\", \"bajar la DJ de IVA de [sociedad] [mes]\", \"cuanto dio el IVA presentado\", \"saldo a favor de IVA del mes anterior\", \"debito y credito de la DDJJ\", \"falta el PDF de la DDJJ de IVA en la carpeta\". No confundir con portal-iva-descarga --presentada (baja los LIBROS de la DDJJ presentada, no la determinacion)."
---

# DDJJ de IVA presentadas (F.2051)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-iva.js --nombre="<sociedad>" --periodo=09/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-iva.js --cuit=<CUIT> --desde=202601 --hasta=202608
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-iva.js --nombres="<A>;<B>" --periodo=09/2026 --archivar
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | Un mes, `MM/AAAA` o `AAAAMM`. |
| `--desde` / `--hasta` | Rango `AAAAMM`. Sin periodo ni rango: el mes anterior. |
| `--archivar` | Copia el PDF a la carpeta IVA del mes como `AAAAMM-IVA-DDJJ-<TOKEN>.pdf`, solo si ahi no hay ya un archivo `*IVA*DDJJ*` (no pisa lo del equipo). |
| `--token` | El `<TOKEN>` del nombre archivado. Default: el nombre de la carpeta de la sociedad sin espacios y en mayusculas. Solo con una sociedad. |
| `--salida` | Otra carpeta local para PDF y JSON (solo con una sociedad). |

## Que hace

Un login por sociedad. Por cada mes: Portal IVA → Declaraciones juradas presentadas →
"Declaracion Jurada de IVA" → la **ultima secuencia** del periodo (si hay rectificativas,
la mas reciente) → CONTINUAR hasta la vista previa. Lee la determinacion y guarda en
`~/Documents/BVA-salidas/ddjj-iva/<CUIT>/` `AAAAMM.pdf` y `AAAAMM.json` (`df`, `cf`,
`st_ant`, `st_arca`, `st`, `sld_ant`, `ret`, `sld`, `saldo_arca`, `secuencia`,
`transaccion`, `presentada_el`). Un mes sin DDJJ presentada sale con `falta: true`.
Cada mes se reintenta hasta 3 veces. Unos 8 minutos por sociedad para 8 meses.

Si ARCA pide captcha o rechaza la clave, corta el lote (seguir bloquea la clave). Si el
apoderado de la planilla no representa a la sociedad, cargar `ARCA_<cuit>_LOGIN` y
`ARCA_<cuit>_PASSWORD` en el `.env` con el de quien si la representa.

## Respuesta al usuario

Por sociedad y mes: secuencia, debito, credito, saldo tecnico, retenciones y saldo final
(a favor del contribuyente o de ARCA), y donde quedo el PDF. Avisar los meses sin DDJJ y
los que dieron error.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium`.
