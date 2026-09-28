---
name: anticipos-sct
description: Releva los anticipos de Ganancias y Bienes Personales de una sociedad (o de todas) en el Sistema de Cuentas Tributarias (SCT) de ARCA y los vuelca en un Excel por sociedad con número de anticipo, capital, intereses, fecha de vencimiento y estado (vencido/no vencido). Corre un script de Playwright que se loguea solo con la planilla de claves del estudio. Usar SIEMPRE que pidan "relevar anticipos", "SCT", "Sistema de Cuentas Tributarias", "anticipos de Ganancias", "anticipos de Bienes Personales", "anticipos vencidos", o "avanza con anticipos de [sociedad]". Los datos se buscan en la solapa DEUDAS del SCT (no en Vencimientos). No confundir con calendario-vencimientos (esa carga fechas de vencimiento futuras en el calendario general del grupo); esta releva anticipos reales ya generados por ARCA, sociedad por sociedad, con su importe exacto.
---

# Relevamiento de anticipos en el SCT (Ganancias y Bienes Personales)

Un script de Playwright entra a ARCA con la clave de la planilla del estudio, se
posiciona en la sociedad, lee la solapa DEUDAS del SCT y arma el Excel. No hace
falta loguearse a mano.

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/anticipos-sct.js --nombre="<sociedad>"
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/anticipos-sct.js --cuit=<cuit> --crudo
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/anticipos-sct-lote.js              # todas las sociedades
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/anticipos-sct-lote.js --solo=7,12  # posiciones de la planilla
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/anticipos-sct-lote.js --consolidar # rearma el consolidado sin entrar a ARCA
```

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | La sociedad (script de una). |
| `--solo` | Posiciones de la planilla de claves, separadas por coma (lote). |
| `--rehacer` | El lote saltea lo ya relevado hoy; esto lo vuelve a relevar. |
| `--salida` | Carpeta destino. Default `~/Documents/BVA-salidas/anticipos/` (o `BVA_SALIDAS_PATH`). |
| `--crudo` | Guarda además el JSON con TODAS las filas de Deudas. Para entender un caso raro. |
| `--ver` | Navegador visible. Para lotes largos conviene headless (sin `--ver`). |

Si es la primera vez en una máquina, hacer el piloto con una sociedad, mostrar el
Excel y recién después correr el lote.

**Plugin instalado:** la primera vez (y después de cada actualización):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Las credenciales van en `~/.fisco-ar/.env` (copiar `.env.example`), que sobrevive a
las actualizaciones.

## Qué entrega

`Anticipos_SCT_[SOCIEDAD]_[AAAA-MM-DD].xlsx`, con las columnas

| Sociedad | Impuesto | Nro. Anticipo | Capital | Int. Resarcitorio | Int. Punitorio | Total | Fecha de Vencimiento | Estado |

Los vencidos van en rojo y negrita. Hay fila TOTAL con `=SUM()` en cada columna
de importes. Al pie, una nota con la fecha de relevamiento, el CUIT y el saldo
total del SCT: la suma de la columna Total tiene que coincidir con ese saldo
cuando todas las deudas son anticipos.

Un anticipo pagado fuera de término queda en Deudas con capital $ 0 y solo
intereses: el Estado dice **"Vencido - solo intereses"**, para que no se lea como
un anticipo impago.

Si la sociedad **no tiene anticipos**, el Excel se genera igual con una fila
"Sin anticipos": saltearla en silencio hace que después nadie sepa si se revisó.

El lote deja además `Anticipos_SCT_CONSOLIDADO_[AAAA-MM-DD].xlsx`, con la hoja
Anticipos (una fila por anticipo, mismas columnas más CUIT y Novedad) y la hoja
Control (una fila por sociedad, incluidas las que fallaron).

## Cómo identifica un anticipo

En la solapa Deudas, un anticipo tiene `Concepto = 191 - ANTICIPOS` y el número
en la columna `Ant/Cuota`. Los impuestos esperados son los de Ganancias
(sociedades o personas humanas) y Bienes Personales. Las retenciones y
percepciones que también dicen "Ganancias" (SICORE, art. 79) no cuentan.

Si aparece un anticipo de **otro** impuesto, **no se descarta**: va al Excel
marcado en rojo y se reporta aparte como novedad, para que el usuario confirme.

El SCT es por CUIT: el Bienes Personales de los socios personas humanas solo
aparece si se releva el CUIT de esa persona.

El estado se decide comparando la fecha de vencimiento contra la de hoy, no por
el color de la fila (el SCT pinta de rojo las vencidas, pero eso es presentación
y puede cambiar).

## Tres cosas que el script verifica y por las que frena

Si alguna dispara, **no es un falso positivo**:

1. **Contribuyente equivocado** (`contribuyente_incorrecto`). El SCT recuerda la
   última sociedad elegida y a veces abre directo en la portada **sin mostrar el
   selector**. Si uno asume que eso ya es la sociedad pedida, releva los
   anticipos de otra empresa y el Excel sale plausible pero equivocado. El
   script confirma el contribuyente activo y corta si no coincide.
2. **Lectura incompleta** (`lectura_incompleta`). La solapa dice cuántas deudas
   hay ("Deudas 15"). Si se leyeron menos filas que ese número, corta: la tabla
   pagina de a 10 y se perderían anticipos.
3. **Solapa incorrecta** (`solapa_incorrecta`). La tabla de Deudas siempre tiene
   columna Saldo; la de Vencimientos no. Si lo visible no tiene Saldo, se está
   leyendo Vencimientos —que no trae importes— y corta.

## Detalles del SCT que ya están resueltos

- Hay que entrar **por el tile** del portal.
- El SCT **rebota al login** cada tanto. No son las credenciales: el script
  vuelve al portal y reintenta hasta 3 veces.
- El contenido vive en un **iframe Vue** (`…/scripts/vue/homeContribuyente/`),
  con `b-table` de BootstrapVue.
- El `value` de las opciones del selector de contribuyente es un índice: se
  elige por texto. Los logins con un solo CUIT no tienen selector.
- El selector de tamaño de página incluye **"Todos"**. Además el script acumula
  página por página con SIGUIENTE, por si el selector falla.
- El contador "Deudas 0" puede tardar en actualizarse.

## Seguridad

- Esto **solo lee**. Nunca modifica, elimina, presenta ni paga nada en ARCA.
- No inventar ni asumir montos, fechas ni números de anticipo. Si el SCT no
  carga o la sociedad no tiene anticipos, se reporta — no se completa con
  supuestos ni se saltea en silencio.
- Ante credenciales rechazadas corta de una: reintentar bloquea la clave fiscal.
- Si aparece un captcha, corta: hay que resolverlo a mano y volver a correr.

## Anticipos no vencidos

Deudas trae lo adeudado, incluidos los anticipos ya generados que todavía no
vencieron. Pero para un CUIT con **Deudas en 0** se vio que sus anticipos
próximos figuraban solo en la solapa **Vencimientos**, que no tiene columna de
importe. Si lo que se busca es "todos los anticipos del año", confirmar con el
usuario antes de dar por cerrado un relevamiento.

## Cierre

Entregar el Excel y resumir en el chat, por sociedad: cantidad de anticipos,
vencidos, próximos vencimientos con importe, e intereses pendientes si los hay.
