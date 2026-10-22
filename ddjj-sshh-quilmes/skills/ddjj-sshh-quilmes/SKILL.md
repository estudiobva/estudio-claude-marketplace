---
name: ddjj-sshh-quilmes
description: "Controla en ARQui (Agencia de Recaudación del Municipio de Quilmes) la DDJJ mensual de la Tasa de Seguridad e Higiene (TISEH, Régimen General) de una, varias o todas las sociedades: compara lo declarado en la actividad 477320 con el papel de trabajo del mes (hoja liquidacion, Neto Gravado de Ventas Netas) y con la DJ de anticipo IIBB de ARBA, avisa las que no tienen DDJJ con el monto a declarar, y archiva la boleta (Recibo de Pago) en la carpeta TSH del mes como 'AAAAMM-SEH-<soc>.pdf'. Corre un script de Playwright que se loguea solo con el usuario de ARQui de la planilla de claves del estudio: no hace falta entrar a mano. Usar SIEMPRE que pidan \"controlá la DDJJ de Seguridad e Higiene de [sociedad] [mes]\", \"SSHH Quilmes\", \"controlá Seguridad e Higiene de todas\", \"TISEH\", \"bajá la boleta de SEH\", \"liquidación mensual de Tasa Comercial\", \"qué falta presentar de SSHH\" o mencionen ARQui y Seguridad e Higiene. NO confundir con recibo-pyp-quilmes (recibo de Publicidad y Propaganda / marquesinas, otra tasa de ARQui). No paga ni presenta nada."
---

# DDJJ mensual Seguridad e Higiene — ARQui Quilmes

Un script de Playwright entra a ARQui con el usuario de la planilla de claves, consulta las DDJJ de Seguridad e
Higiene del período, baja la Declaración y el Recibo de Pago, compara lo declarado con el papel de trabajo y con
ARBA, y archiva la boleta. Navegador oculto. Solo consulta y descarga: nunca usa "Nuevo" / "Generar Comprobante"
(no presenta DDJJ), ni pagos, ni "Adherir cuenta".

## Correrlo

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-sshh-quilmes.js --nombre="<sociedad>" --periodo=AAAAMM
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-sshh-quilmes.js --nombres="<A>;<B>" --periodo=AAAAMM
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/ddjj-sshh-quilmes.js --todas --periodo=AAAAMM
```

| Argumento | Para qué |
|---|---|
| `--nombre` / `--cuit` | Una sociedad (titular de la planilla de claves, o CUIT). |
| `--nombres` / `--cuits` | Varias, separadas por `;`. Un login para todas. |
| `--todas` | Todas las cuentas de SSHH vinculadas al usuario de ARQui; la sociedad sale del CUIT de sus DDJJ. Suma la hoja **Control planilla**. |
| `--periodo` | `AAAAMM` o `MM/AAAA`. Por defecto, el mes anterior. |
| `--no-archivar` | Deja la boleta solo en la carpeta de salida; no toca la unidad. |
| `--reemplazar` | Pisa la boleta del mes si ya estaba en la unidad (por defecto la deja). |
| `--ver` | Navegador visible. Solo si el usuario lo pide. |

Si falta la sociedad o el período, preguntar antes de correr ("SSHH de septiembre" = `--periodo=202609`).
`--todas` (unas 33 cuentas) tarda unos 15–20 minutos: hay una pausa a propósito entre pedidos (ver "Bloqueos"). Correrlo en segundo plano.

**Plugin instalado:** la primera vez (y después de cada actualización):
`npm install && npx playwright install chromium && python3 -m pip install -r requirements.txt`.
Usa el mismo `~/.fisco-ar/.env` que los demás plugins (`CLAVES_ORGANISMOS_PATH`, `CLAVES_ORGANISMOS_PASSWORD`, y
`BVA_UNIDAD_PATH` en Mac).

## De dónde sale cada dato

- **Login:** `Claves_Organismos.xlsx`, grupo **MUNICIPAL** de la fila de la sociedad (Municipio = Quilmes), o el
  usuario de Quilmes más repetido. Override: `ARQUI_USUARIO` / `ARQUI_CLAVE` en el `.env`. Claude nunca escribe ni
  muestra la clave.
- **Cuenta:** `Tasas Municipales CLAVES.xlsx`, hoja **QUILMES**, bloque **SEGURIDAD E HIGIENE**, por CUIT (los
  nombres de la planilla no son los de las carpetas). Los bloques de Marquesinas y Publicidad son otras tasas.
- **Lo declarado:** ARQui > Consultar (DDJJ del año) y la **Declaración** de la DDJJ del mes (si hay rectificativas,
  la última): actividades, alícuota, ingreso, tasa, vencimiento y Nº de recibo.
- **Papel:** `[Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM-papeldetrabajo-*.xlsx` (o dentro de la carpeta de año),
  hoja `liquidacion`, fila "Ventas Netas (s/ Notas de Crédito)", columna Neto Gravado. Es la venta de 477320 (los
  medicamentos van exentos). Se compara la parte entera.
- **ARBA:** la DJ de anticipo IIBB del mes que esté en la carpeta del mes (cualquier `*iibb*dj*.pdf` que no sea el
  acuse, también en `IIBB/` o `IVA/`), línea 477320. SSHH vence (~16 del mes siguiente) antes de que se presente IIBB,
  así que en los primeros días puede no estar todavía: el control queda contra el papel. Si la DJ está escaneada
  (PDF sin texto) no se puede leer y se avisa.

## Estados

- `OK`: declarado = papel y = ARBA (lo que haya), solo con 477320. Menos de $1 de diferencia no cuenta (hay quien
  declara con centavos o redondea).
- `diferencia`: no coincide o hay otras actividades declaradas. **No se toca:** rectificar lo decide la persona.
- `sin DDJJ`: no hay DDJJ del mes en ARQui. La observación trae el monto a declarar (papel, parte entera) y la tasa
  estimada al 1,3%.
- `sin control`: hay DDJJ pero no hay papel ni ARBA para comparar.
- `sin cuenta` / `sin CUIT`: no está en la planilla de tasas, o (con `--todas`) la cuenta no tiene DDJJ ni en el año
  ni en el anterior y no está en la planilla.
- `error` (con `codigo`): `no_vinculada` (la cuenta no está en el usuario de ARQui: vincularla con "Adherir cuenta"
  lo decide la persona), `control` (la Declaración es de otro CUIT o período), `bloqueo`, `login`.

## Qué entrega

- Boleta archivada en `[Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-SEH-<soc>.pdf`, con el mismo `<soc>` que
  el equipo ya usa para esa sociedad (ej. `202609-SEH-sanar65.pdf`). Si ya había una boleta del mes en la carpeta del
  mes, no se pisa (`ya estaba`).
- En `~/Documents/BVA-salidas/ddjj-sshh-quilmes/AAAA-MM-DD/` (`BVA_SALIDAS_PATH`): la Declaración y la boleta de cada
  cuenta y el resumen `DDJJ SSHH QUILMES AAAAMM <fecha> <hora>.xlsx` con la hoja **Control** (Sociedad, CUIT, Cuenta,
  Estado, fecha, carácter, declarado, otras actividades, tasa, papel, ARBA, diferencias, boleta, vencimiento, Nº de
  recibo, observación) y, con `--todas`, la hoja **Control planilla**.
- JSON por stdout con lo mismo y `totales`.

## Bloqueos de ARQui (403 Forbidden)

El firewall del municipio corta todo con **403 Forbidden** unos minutos si recibe muchos pedidos o logins seguidos.
El script hace un solo login por usuario, pausa `ARQUI_PAUSA_MS` (2,5 s) entre pedidos y, ante un 403, espera
`ARQUI_ESPERA_BLOQUEO_MS` (5 min) y reintenta una vez; si sigue, corta con código `bloqueo` y `totales.bloqueado:
true`. En ese caso **no relanzar enseguida**: esperar 5–10 minutos y volver a correr solo las que faltaron.

## Respuesta al usuario

Una línea por sociedad: Sociedad | Cuenta | Estado | Declarado | Papel | ARBA | Tasa | Vencimiento, y dónde quedó la
boleta. Después, en una línea cada una, las `diferencia` (con el detalle), las `sin DDJJ` (con el monto a declarar)
y los errores, sin adivinar. Con `--todas`, resumir también la hoja Control planilla.
