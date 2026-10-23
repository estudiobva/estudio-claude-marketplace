---
name: constancias-anticipo-arba
description: "Baja de ARBA (app.arba.gov.ar/IBPresentaciones, Reimpresion > Constancia anticipo) las constancias de las DDJJ de anticipos mensuales de Ingresos Brutos YA PRESENTADAS de una o varias sociedades, con los datos del detalle (gravados, deducciones, impuesto determinado), y las archiva en [Sociedad]/01-Impuestos Mensuales/AAAAMM/IIBB/AAAAMM-IIBB-DJ-<TOKEN>.pdf. Script de Playwright con login automatico (CIT de la planilla de claves), navegador oculto, solo lectura. Usar SIEMPRE que pidan \"imprimir las constancias de anticipo de [sociedad]\", \"bajar los resumenes de IIBB del periodo X al Y\", \"reimpresion de constancia anticipo\", \"las DDJJ cerradas de anticipos mensuales\", \"los 12 resumenes del ejercicio\", \"la DJ de ARBA de [mes]\" o rehacerlas porque se perdieron. NO confundir con arba-alicuotas-iibb ni con cruce-alicuotas-agip-arba (alicuotas)."
---

# Constancias de anticipo de IIBB (ARBA)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/constancias-anticipo-arba.js --nombre="<sociedad>" --desde=202601 --hasta=202612
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/constancias-anticipo-arba.js --nombres="<A>;<B>" --periodo=09/2026
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | Un mes, `MM/AAAA` o `AAAAMM`. |
| `--desde` / `--hasta` | Rango `AAAAMM`. Sin periodo ni rango: el mes anterior. |
| `--no-archivar` | Deja los PDF solo en `~/Documents/BVA-salidas/arba-ddjj/<CUIT>/`. |
| `--token` | El `<TOKEN>` del nombre archivado, si la sociedad todavia no tiene ningun `AAAAMM-IIBB-DJ-<TOKEN>` (preguntarle al equipo cual usa; no inventarlo). Solo con una sociedad. |

**Antes de correr:** contar los periodos del rango. Si el usuario dijo una cantidad (por
ejemplo "los 12 del ejercicio") y el rango da otra, preguntar cual toma.

## Que hace, por sociedad

1. Entra con la CIT de la sociedad (columnas ARBA de `Claves_Organismos.xlsx`). Un solo
   intento: si ARBA rechaza, corta esa sociedad con `arba_login_rechazado` y sigue con la
   proxima (reintentar bloquea la CIT). Controla que ARBA haya entrado con ese CUIT.
2. Reimpresion → Constancia anticipo (la primera vez por el menu: directo a la URL queda
   la pantalla vacia). Busca el rango entero y lee la tabla completa.
3. Por cada mes abre la DJ (si hay rectificativa, la presentada mas reciente), lee el
   detalle y baja el PDF de `resumenPDF.do` (devuelve la DJ que quedo abierta).
4. Archiva en `[mes]/IIBB/AAAAMM-IIBB-DJ-<TOKEN>.pdf` (crea `IIBB` si falta). Si ya hay una
   `AAAAMM-IIBB-DJ-*` en el mes o en `IIBB/`, no la toca ("ya existía"). El `<TOKEN>` sale de
   los archivos que ya tiene la sociedad; si no hay ninguno, `archivo_error: sin_token`.

Sociedades sin CIT (Convenio Multilateral o AGIP) salen con `sin_cit`.

## Respuesta al usuario

Por sociedad: la razon social con que entro ARBA, por mes el tipo (Original/Rectificativa),
el impuesto determinado y donde quedo el PDF (o "ya existía"). Avisar meses sin DJ,
rectificativas, CIT rechazadas y sociedades sin token.

**Plugin instalado:** la primera vez (y despues de cada actualizacion):
`npm install && npx playwright install chromium`.
