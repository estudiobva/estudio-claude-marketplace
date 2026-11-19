# Transcribir el comparativo del balance anterior

Instructivo para un subagente (uno por sociedad). En el prompt del subagente pasar: nombre, CUIT, cierre del balance a transcribir, ruta del config, carpeta de la sociedad, ruta del PDF y una carpeta de trabajo propia, y la ruta del plugin (el subagente no siempre tiene `CLAUDE_PLUGIN_ROOT`: reemplazala en los comandos).

Transcribí a mano el comparativo del balance anterior de UNA sociedad a su archivo JSON de configuración. Respondé en español.

DATOS DE ESTA SOCIEDAD: ver al final ("SOCIEDAD").

CÓMO LEER EL PDF (escaneado, sin texto): con python3 y pypdfium2 renderizá páginas a PNG y miralas con la herramienta Read. Helper: `python "${CLAUDE_PLUGIN_ROOT}/scripts/herramientas/zoom.py" "<pdf>" <carpeta>/z "<pag>:<rot>:y0:y1:x0:x1"` (pag 1-based; rot en grados para PIL, -90 para páginas apaisadas que se leen de abajo hacia arriba; y0..x1 fracciones 0–1; varias specs separadas por coma; genera <carpeta>/z_p<pag>.png). Primero mirá todas las páginas en baja resolución para ubicar cada estado; después zoom para leer cifras exactas con centavos. Formato argentino 1.234.567,89; paréntesis = negativo. Usá SOLO la columna del ejercicio del balance (la más reciente), no la del ejercicio anterior.

MODELO DE FORMATO: cualquier config que ya tenga "comparativo" en la carpeta de configs (EERR 2026/Config Balance) y el esqueleto de abajo (copiá su estructura; el archivo de tu sociedad ya existe con nombre, cuit, cierre, etc.: CONSERVÁ "nombre", "cuit" y "cierre" tal cual y completá el resto). Campos:
- razon_social (como dice el balance), registro (Nº de registro DPPJ / IGJ, string), ejercicio_nro = número del ejercicio del PDF + 1 (string), fuente_comparativo = "<ruta relativa desde la carpeta de la sociedad al PDF> (bajado de PUB de ARCA y transcripto a mano el DD/MM/AAAA)".
- bancos: [[nombre, saldo, ""]] uno por banco de la Nota de Caja y Bancos (sin la caja). Nombres normalizados: "Banco BBVA (Francés) Cta. Cte.", "Banco Comafi Cta. Cte.", "Banco Galicia Cta. Cte.", "Banco Galicia Cta. Cte. (2)", "Banco Galicia Más (ex HSBC) Cta. Cte." (HSBC = Galicia Más), "Banco Macro Cta. Cte.", "Banco Nación Cta. Cte.", "Banco Santander Cta. Cte.", "Banco Hipotecario Cta. Cte.", "Banco Provincia Cta. Cte.", "Banco Patagonia Cta. Cte.", etc. Saldos negativos dentro de caja y bancos van negativos acá.
- fci_2025: {"GALICIA": x} si el FCI es HSBC o Galicia; otros: "FRANCES", "COMAFI", o el nombre del banco en mayúsculas. Si hay plazos fijos u otras inversiones que no son FCI, van igual en fci_2025 con una clave descriptiva (ej. "PLAZO FIJO NACION").
- deudas_bancarias: saldos deudores de bancos en deudas financieras [[nombre (saldo deudor), saldo, ""]]. Préstamos bancarios → nota2 "j.prestamos".
- ganancias_2025: {"nota": "No está la DDJJ Ganancias en la unidad. Balance anterior: saldo a favor $X; provisión $Y"}.
- mapa_anexo2: líneas del EERR → cuenta del Anexo 2, SOLO si el destino existe en tu anexo2: "Gastos de transporte"/"Combustibles" → la cuenta de transporte o "Movilidad y viáticos" si existe; "Matricula Colegio Farmaceuticos" → "Matrícula Colegio Farmacéuticos"; "Gastos varios" → "Gastos varios"; "Uniformes" → "Uniformes"; "Agua" → "Agua"; "Seguridad e Higiene" → "Seguridad e Higiene".
- comparativo.nota2 por código: a.1 caja; c.1 deudores por ventas (si hay varias cuentas de créditos por ventas, c.1 = la principal y el resto a cuentas_extra "c"); d.1 IVA saldo técnico a favor; d.2 IVA libre disponibilidad; d.3 saldo a favor Ganancias; d.4 saldo a favor IIBB; d.5 crédito imp. débitos y créditos; d.6 retenciones Ganancias; d.7 anticipos Ganancias; e.1 mercadería; f.1 quebranto (0 si no hay); f.2 impuesto diferido no corriente; g.1 proveedores; h.1 sueldos y cargas a pagar; h.2/h.3/h.4 0 salvo esas líneas; i.1 IIBB a pagar; i.2 IVA a pagar; i.3 provisión Ganancias; i.4 tasas municipales a pagar; j.prestamos; k.1 impuesto diferido pasivo. Cuenta que no encaja (otros créditos, bienes de uso netos si existen, anticipos a proveedores, otras deudas, planes de pago, provisiones, etc.) → comparativo.cuentas_extra {letra: [[cuenta, saldo, fuente sugerida]]} con la letra del rubro donde está en el balance (para rubros que no son a–k, usá la letra más cercana y aclaralo). Si un rubro no trae detalle por cuenta, comparativo.rubros_sin_detalle {letra: total}.
- er (12): [ventas, CMV(neg), ganancia bruta, gastos adm(neg), gastos com(neg), gastos fin(neg), resultado ordinario, resultados financieros y por tenencia incl. RECPAM, otros ingresos/egresos, resultado antes de IG, impuesto a las ganancias(neg), resultado del ejercicio]. RECPAM en línea aparte → posición 7; otros ingresos → 8; si el resultado antes de IG no viene impreso, calculalo.
- pn_inicio: saldos al CIERRE del EPN (fila "Saldos al cierre"): capital, ajuste, reserva_legal, otras_reservas, rna. Suma = PN al cierre.
- pn_2025: [PN al inicio del ejercicio, distribuciones/dividendos/otros movimientos (negativo si restan), resultado del ejercicio, PN al cierre].
- flujo (22, en este orden): [efectivo al inicio, efectivo al cierre, aumento/disminución, resultado del ejercicio, depreciaciones+bajas BU, var créditos por ventas, var créditos fiscales, var bienes de cambio, var créditos fiscales no corrientes, var deudas comerciales, var deudas sociales, var deudas fiscales, var deudas financieras, var otras deudas no corrientes, variación RECPAM/RFyT, flujo operativo, compras de bienes de uso, flujo de inversión, RFyT generado por efectivo (financiación), deudas financieras+distribución de utilidades (financiación), flujo de financiación, aumento/disminución neta]. Renglones extra: sumalos al más afín y anotalo. Copiá los signos como están impresos.
- anexo2: {cuenta: [adm, com, fin, total]} con nombres normalizados ("Comisiones de tarjetas", "Gastos de transporte", "Gastos bancarios", "Imprenta y Papelería", "Servicios de luz, gas, teléfono, internet", "IVA Crédito Fiscal no computable", "Matrícula Colegio Farmacéuticos", "Amortizaciones de bienes de uso", etc.). Imputaciones raras del original se respetan.
- anexo3: [existencia inicial, compras, existencia final].

CONTROL OBLIGATORIO: corré `python "${CLAUDE_PLUGIN_ROOT}/scripts/herramientas/chk_config.py" <tu config>`. Todo debe dar 0 (±0,02) y los totales de a), b), d), i), j) deben coincidir con el estado de situación (el chk no suma cuentas_extra salvo en i: si usaste cuentas_extra en otros rubros, verificá esa suma vos). Si algo no cierra, volvé a hacer zoom. Si una diferencia viene del propio PDF (redondeo o error impreso), copiá lo impreso y anotalo. No inventes números.
Si el PDF NO es un balance (o está incompleto / ilegible), no escribas el config y explicalo.

REGLAS: escribí SOLO tu archivo de config (UTF-8, ensure_ascii=False, indent=1) y archivos en tu carpeta de trabajo. No modifiques nada de la unidad. No corras otros scripts.

RESPUESTA FINAL (corta): razón social, registro, ejercicio siguiente, PN al cierre, resultado, salida completa del chk_config.py, y datos dudosos / cuentas_extra / rubros_sin_detalle.

VERIFICACIÓN DE CIERRE Y PDF EQUIVOCADO: el cierre del balance a transcribir es el indicado en "SOCIEDAD". Si el PDF es de otro cierre, o es solo la constancia de presentación de PUB (1 página, F.8250), buscá en la carpeta de la sociedad (04-Balances y EECC/2025, 04-Balances y EECC/2026 y sus subcarpetas, incluida "PUB") otro PDF que sea el balance completo de ESE cierre y usalo (decilo en la respuesta y usá su ruta en fuente_comparativo). Si no hay ninguno, NO escribas el config y explicalo.

## Esqueleto del config (valores de ejemplo)

```json
{
 "nombre": "Ejemplo", "cuit": "30000000000", "cierre": "2026-12-31",
 "razon_social": "Farmacia Ejemplo S.C.S.", "registro": "123.456", "ejercicio_nro": "21",
 "fuente_comparativo": "04-Balances y EECC/2025/EECC Legalizados/EECC_30000000000_2025_0.pdf (bajado de PUB de ARCA y transcripto a mano el DD/MM/AAAA)",
 "bancos": [["Banco Galicia Cta. Cte.", 1000.0, ""], ["Banco Nación Cta. Cte.", -50.0, ""]],
 "fci_2025": {"GALICIA": 5000.0},
 "deudas_bancarias": [["Banco Hipotecario Cta. Cte. (saldo deudor)", 10.0, ""]],
 "ganancias_2025": {"nota": "No está la DDJJ Ganancias en la unidad. Balance anterior: saldo a favor $X; provisión $Y"},
 "mapa_anexo2": {"Gastos de transporte": "Gastos de transporte", "Uniformes": "Uniformes"},
 "comparativo": {
  "nota2": {"a.1": 0, "c.1": 0, "d.1": 0, "d.2": 0, "d.3": 0, "d.4": 0, "d.5": 0, "d.6": 0, "d.7": 0, "e.1": 0,
            "f.1": 0, "f.2": 0, "g.1": 0, "h.1": 0, "h.2": 0, "h.3": 0, "h.4": 0, "i.1": 0, "i.2": 0, "i.3": 0, "i.4": 0,
            "j.prestamos": 0, "k.1": 0},
  "cuentas_extra": {"i": [["Planes de facilidades a pagar", 0, "Mis Facilidades ARCA"]]},
  "rubros_sin_detalle": {},
  "er": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  "pn_inicio": {"capital": 0, "ajuste": 0, "reserva_legal": 0, "otras_reservas": 0, "rna": 0},
  "pn_2025": [0, 0, 0, 0],
  "flujo": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  "anexo2": {"Alquileres y Expensas": [0, 0, 0, 0], "Sueldos y cargas sociales": [0, 0, 0, 0]},
  "anexo3": [0, 0, 0]
 }
}
```

Casos ya vistos: provisión de Ganancias restada dentro de créditos fiscales (va en `cuentas_extra` "d" en negativo); planes de pago dentro de deudas fiscales (`cuentas_extra` "i"); aportes irrevocables en el PN (sumarlos a `ajuste` y anotarlo en `nota_pn`); FCI o préstamo sin banco (clave descriptiva y avisar); bancos con saldo negativo dentro de caja y bancos (negativos en `bancos`).
