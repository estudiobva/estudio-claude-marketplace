---
name: deudas-ddjj-sct
description: Controla en el Sistema de Cuentas Tributarias (SCT) de ARCA de todas las sociedades (o de una) si hay deudas que NO son anticipos (saldos de DDJJ, multas, intereses, retenciones) y DDJJ pendientes de presentación, y lo vuelca en un Excel consolidado con hojas Control, Deudas y DDJJ pendientes. Corre un script de Playwright que se loguea solo con la planilla de claves del estudio. Usar SIEMPRE que pidan "deudas del SCT", "deudas que no sean anticipos", "faltas de presentación", "DDJJ pendientes", "DDJJ no presentadas", "declaraciones omitidas", "qué debe [sociedad] en ARCA" o "control de cuentas tributarias". No confundir con anticipos-sct (esa releva SOLO los anticipos de Ganancias y Bienes Personales) ni con planes-facilidades-arca (cuotas de planes de pago).
---

# Deudas (sin anticipos) y DDJJ pendientes en el SCT

Un script de Playwright entra a ARCA con la clave de la planilla del estudio, se
posiciona en cada sociedad y lee dos solapas de la portada del SCT: **Deudas**
(descartando los anticipos) y **DDJJ pendientes de presentación**. No hace falta
loguearse a mano.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-ddjj-sct-lote.js              # todas las sociedades
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-ddjj-sct-lote.js --solo=7,12  # posiciones de la planilla
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-ddjj-sct-lote.js --consolidar # rearma el consolidado sin entrar a ARCA
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deudas-ddjj-sct.js --nombre="<sociedad>"
```

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | La sociedad (script de una). |
| `--solo` | Posiciones de la planilla de claves, separadas por coma (lote). |
| `--rehacer` | El lote saltea lo ya relevado hoy; esto lo vuelve a relevar. |
| `--por-sociedad` | El lote genera además un Excel suelto por sociedad. |
| `--salida` | Carpeta destino. Default `~/Documents/BVA-salidas/deudas-ddjj/` (o `BVA_SALIDAS_PATH`). |
| `--crudo` | (script de una) guarda además el JSON con todas las filas de las dos solapas. |
| `--ver` | Navegador visible. Para lotes largos conviene headless (sin `--ver`). |

El lote completo tarda del orden de una hora. Si es la primera vez en una máquina,
probar antes con `--solo=` y dos o tres sociedades.

**Plugin instalado:** la primera vez (y después de cada actualización):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Las credenciales van en `~/.fisco-ar/.env` (copiar `.env.example`), que sobrevive a
las actualizaciones.

## Qué entrega

Un Excel consolidado, `Deudas_DDJJ_SCT_CONSOLIDADO_[AAAA-MM-DD].xlsx`:

- **Control**: una fila por sociedad, incluidas las que fallaron (ERROR, con el
  motivo) y las omitidas. Resultado: Sin observaciones / Con deudas /
  DDJJ pendientes / Deudas y DDJJ pendientes.
- **Deudas**: una fila por deuda, con Capital | Int. Resarcitorio | Int. Punitorio |
  Total | Vencimiento | Estado | Observación. Las vencidas en rojo, fila TOTAL.
  La Observación aclara "Multa", "Capital cancelado: quedan intereses",
  "Diferencia menor a $ 1" o "Sin fecha de vencimiento".
- **DDJJ pendientes**: Impuesto | Concepto | Subconcepto | Período | Vencimiento |
  Días de atraso. El SCT no informa importes para estas.

## Qué cuenta como deuda "no anticipo"

Es el complemento exacto del criterio de `anticipos-sct` (`esAnticipo` en
`lib/sct.js`): lo que uno cuenta como anticipo, el otro no lo cuenta como deuda.
Anticipo = `Concepto 191 - ANTICIPOS`, o número en `Ant/Cuota` que no sea una DDJJ.
Todo lo demás de la solapa Deudas entra, de cualquier impuesto.

## Por qué frena (no son falsos positivos)

1. **Contribuyente equivocado** (`contribuyente_incorrecto`): el SCT recuerda la
   última sociedad y puede abrir en otra. Se confirma el CUIT activo antes de leer.
2. **Lectura incompleta** (`lectura_incompleta`): se leyeron menos filas de las que
   declara el contador de la solapa.
3. **Solapa incorrecta** (`solapa_incorrecta`): lo visible no es la tabla esperada.
   Deudas siempre tiene Saldo; DDJJ pendientes tiene Establecimiento y Vencimiento,
   y no tiene Detalle ni Fecha Vencimiento (esas son de Vencimientos). Si la solapa
   está en 0 no se abre: se espera, se reintenta y recién ahí se da por vacía.

## Detalles del SCT ya resueltos

- Una tabla vacía trae una fila "No se encontraron resultados": no es un registro.
- Los contadores de las solapas pueden tardar en cargar o no leerse; se ignoran si
  la tabla tiene filas.
- El resto (rebote al login, iframe Vue, selector por texto, "Mostrar Todos" más
  SIGUIENTE) es igual que en `anticipos-sct`.

## Seguridad

- Esto **solo lee**. Nunca modifica, elimina, presenta ni paga nada en ARCA.
- No inventar montos ni fechas. Si una sociedad falla, queda como ERROR en Control:
  no se da por "sin observaciones".
- Ante credenciales rechazadas corta de una: reintentar bloquea la clave fiscal.
- Si aparece un captcha, corta: hay que resolverlo a mano y volver a correr.

## Cierre

Entregar el consolidado y resumir en el chat: sociedades con deudas (cantidad,
vencidas, total), sociedades con DDJJ pendientes (qué obligación y período), y las
que dieron ERROR. Separar las deudas que son solo intereses o centavos de las que
tienen capital.
