# bancos-bva

Plugin de Claude del Estudio BVA para los extractos bancarios, con la misma
estructura que **fisco-ar**: scripts de Playwright + skills que los usan.

| Skill | Qué hace |
|---|---|
| `bancos-setup` | Instala Playwright/Chromium y corre el diagnóstico (`scripts/doctor.js`). |
| `descarga-extractos-bancarios` | Login al home banking de empresas, descarga, renombrado y archivado en `~/Documents/BVA-salidas/extractos/[Sociedad]/[BANCO]/[AAAA]/` (con `--unidad`, en `03-Bancos y Conciliaciones/[BANCO]/[AAAA]/`). |

```bash
npm install
node scripts/doctor.js
node scripts/extractos-banco.js --nombre="<sociedad>" --banco=todos --desde=202601 --hasta=202606
```

- Bancos con login: Galicia, Comafi, Santander, Hipotecario, Macro, Nación,
  BBVA, Patagonia, Provincia (`lib/bancos/`). Estado real de cada uno: ver abajo.
- Credenciales: `Claves_Bancos.xlsx` (cifrado) vía el mismo `.env` de fisco-ar
  (`~/.fisco-ar/.env`). Ningún script imprime claves.
- Unidad compartida: se detecta sola en Windows (`I:\`) y Mac (Google Drive).

Estructura:

```
.claude-plugin/   plugin.json
lib/              motor (bancos/, extractos.js, claves-bancos.js, rutas-bva.js, + vendorizados de fisco-ar)
scripts/          extractos-banco.js, banco-login.js, claves-bancos.js, doctor.js
skills/           bancos-setup, descarga-extractos-bancarios
```

## Salida

`~/Documents/BVA-salidas/extractos/[N] - [Sociedad] - [CUIT]/[BANCO]/[AAAA]/[sociedad]_[banco]_[AAAA-MM]_mensual.pdf`
(`--salida` o `BVA_SALIDAS_PATH`). Con `--unidad`, directo en la unidad compartida:
`[Sociedad]/03-Bancos y Conciliaciones/[BANCO]/[AAAA]/`. Nunca pisa: si el contenido ya
está lo saltea, si hay otro distinto con el mismo nombre agrega `_2`.

## Estado por banco (primera corrida real, 28-29/09/2026)

Ningún banco tiene todavía descarga automática (`descargar()` en su adaptador): el
script loguea y después la navegación se hace en la ventana (a mano o con
`--control` + `scripts/control.js`).

| Banco | Login | Descarga | Detalle |
|---|---|---|---|
| Macro | OK | OK con `--control` | Selector de empresa: el clic va al botón `#section0_repeatN_actionButtonVerify` (tapa el texto). Cuenta → menú oculto `#menuTableWidget<id> #menu_resumenes`. Filas `a#collectionTable1022_DETAIL` (id repetido: `--n`). El PDF se llama `Resumen.pdf` (sin período). Hay que esperar ~6 s entre descargas (la página recarga). Online solo 13 meses. |
| Nación (BNA+) | OK | OK con `--control` | Cambio de empresa: `#loggedUserMenu` → Ambientes. Cuenta → Más opciones → `[id='account.last.summary.title.label']` → Mostrar más → `#downloadPDFPeriodN`. La "fecha de emisión" de la lista **no** mapea fijo al mes: hay que leer el PERIODO dentro del PDF. La sesión vence a los pocos minutos sin actividad. |
| BBVA (Francés) | **Bloqueado** | — | El formulario se completa (se arregló que la página vaciaba "Código de empresa" al inicializar). Al enviar, `POST aso-live-senda.bbva.com.ar/TechArchitecture/mx/grantingTicket/V02` responde **403** con una página HTML "Algo salió mal — Reference ID: 0.xxxxxxxx.<epoch>.xxxxxxxx" (formato de bloqueo de Akamai), antes de validar la clave. Idéntico con y sin `?260828` en la URL (`BVA_IGNORAR_LINK_EXCEL=1` fuerza la del adaptador). La pantalla queda en "Verificando tus credenciales". |
| Comafi | **Bloqueado** | — | Cloudflare Turnstile. El script ahora lo detecta y espera a la persona, pero el widget responde "La verificación falló" aun tildado a mano. |
| Galicia | Rechazado | — | "Los datos ingresados son incorrectos" con el login de la sociedad probada. Probablemente el dato del archivo de claves: confirmarlo entrando a mano. No se reintenta. |
| Santander | Frenado | — | "Tu contraseña venció": postergar vuelve al login. Hay que renovar la clave. |
| Hipotecario | OK | **Sin resolver** | Selector de empresa por CUIT (`[id='<cuit>']`). La UI nueva no muestra sección de resúmenes: solo Últimos movimientos con "Elegir fecha" y "Descargar". Falta saber dónde están los resúmenes oficiales. |
| Patagonia | Sin probar | — | Ninguna sociedad probada lo tiene. Segundo paso del login nunca visto. |
| Provincia | Sin probar | — | Ninguna sociedad probada lo tiene. |

El motor común (`lib/bancos/comun.js`) ahora: usa el link de `Claves_Bancos.xlsx` si es del
mismo host que el adaptador; revisa y rellena los campos que la página vacía; detecta
captchas (Turnstile, reCAPTCHA, hCaptcha) y los deja a la persona — **nunca** los resuelve
solo ni con servicios externos.
