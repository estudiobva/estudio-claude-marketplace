---
name: ddjj-sshh-quilmes
description: Controla y carga la DDJJ mensual de la Tasa de Seguridad e Higiene (TISEH, Régimen General) de las sociedades con cuenta en ARQui Quilmes (arquivirtual.quilmes.gov.ar), siempre con la actividad 477320. El monto para cargar sale del papel de trabajo del mes (hoja liquidacion, Neto Gravado de Ventas Netas), porque SSHH vence antes de que se presente IIBB en ARBA; la constancia de anticipo IIBB de ARBA (línea 477320) se usa para controlar. Usar cuando pidan "cargá la DDJJ de Seguridad e Higiene de [sociedad] [mes]", "SSHH Quilmes", "controlá Seguridad e Higiene de todas", "liquidación mensual de Tasa Comercial", "declarar ingresos gravados en ARQui", "TISEH" o "boleta de SEH". No paga nada.
---

# DDJJ mensual Seguridad e Higiene – ARQui Quilmes

Hablar en español rioplatense, respuestas breves. Usar el navegador con las herramientas de Claude in Chrome (no clicks a nivel sistema). No paga: nunca usar botón de pago ni pagos online.

## Por qué el monto sale del papel y no de ARBA
La DDJJ de SSHH del mes M vence alrededor del 16 de M+1 (09/2026: boletas emitidas el 12–13/10, vencimiento 16/10). La DJ de anticipo IIBB de ARBA del mismo mes se presenta después (09/2026: del 15 al 21/10). Entonces:
- **Cargar** con el monto del papel de trabajo.
- **Controlar** contra ARBA cuando la DJ ya está presentada.

## Datos de entrada
- Sociedad (CUIT) y período (mes/año). Si piden "todas", son las cuentas del bloque Seguridad e Higiene de la planilla (15 a 10/2026).
- **Nº de cuenta TISEH:** `0 - ESTUDIO - 00000000000/Tasas Municipales CLAVES.xlsx`, hoja **QUILMES**, bloque **SEGURIDAD E HIGIENE** (columnas A–C: Farmacia | CUIT | Nro de Cuenta). Buscar **por CUIT**: los nombres de la planilla no son la razón social ni el de la carpeta ("Factory" = Quilmes Factory, "Sanar  52" con doble espacio, "Nueva 50" = Soy 50 Roasio). No usar los bloques MARQUESINAS ni PUBLICIDAD Y PROPAGANDA: son otras tasas y a veces traen otro número. Las cuentas van con dígito verificador (ej. 500123456-7), salvo alguna cargada sin él.
- **Actividad: siempre 477320** (venta al por menor de cosméticos y perfumería). No declarar 477311 / 477312 (medicamentos) en esta cuenta. Comprobado en 09/2026: en 9 de las 10 sociedades con la DJ de ARBA presentada, lo declarado en SSHH es exactamente lo de 477320.
- **Monto para cargar:** papel de trabajo del mes, `[Sociedad]/01-Impuestos Mensuales/AAAAMM/AAAAMM-papeldetrabajo-<SOC>*.xlsx` (o `01-Impuestos Mensuales/AAAA/AAAAMM/` si la sociedad usa carpeta de año), hoja **liquidacion**, fila **"Ventas Netas (s/ Notas de Crédito)"**, columna **Neto Gravado**. Ese neto gravado es la venta de 477320 (los medicamentos van exentos de IVA).
  - Si no hay papel, no tiene la hoja `liquidacion` (ej. Sanar 1 en 09/2026) o la celda viene vacía (fórmulas sin calcular, ej. Quilmes Factory en 09/2026): no calcularlo por otro lado. Avisar y pedir el monto a quien liquida IIBB de la sociedad.
- **Monto ARBA (control):** DJ de anticipo IIBB del período, línea `477320 $ X`. Buscarla en la carpeta del mes, en sus subcarpetas `IIBB/` e `IVA/` y en la carpeta de año. Los nombres varían en mayúsculas y forma: `AAAAMM-iibb-dj-<soc>.pdf`, `AAAAMM-IIBB-DJ-<SOC>.pdf`, `AAAAMM-IIBB-DDJJ-<Soc>.pdf`. No confundir con el `acuse`. Si no está y la DJ ya se presentó, usar la skill constancias-anticipo-arba (de la cuenta de claude.ai, no está en este marketplace). Si todavía no se presentó, el control contra ARBA queda pendiente: decirlo, no suponerlo.

## Cómo escribir el monto
Solo la parte entera, sin separador de miles ni decimales, **truncando** (sin redondear): 12.345.678,92 → `12345678`; $ 12.345.678,00 → `12345678`. Así lo declara el equipo en ARQui y en ARBA. Cuidado: sacarle "puntos y comas" a `12.345.678,00` da 1234567800, cien veces más.

## Login
El usuario inicia sesión a mano en ARQui (Chrome completa la clave con PIN de Windows). Nunca manejar ni pedir credenciales. Arrancar desde la sesión abierta.

ARQui tiene un firewall que, con muchos pedidos seguidos, contesta **403 Forbidden** a todo durante unos minutos. Con varias sociedades, ir de a una y sin recargar de más. Si aparece el 403, parar, esperar 5–10 minutos y seguir con las que faltan.

## Pasos
1. En ARQui ir a Servicios > Tasa Comercial – Seguridad e Higiene > Régimen General > **Liquidación Mensual** (`/municipal/App/DDJJ/Sshh/SSHHInicio.aspx`).
2. **Nueva**: en "Seleccione una cuenta" elegir la cuenta **sin dígito verificador** (500123456-7 → 500123456; lista larga: seleccionar por texto), Año y Mes. Lleva a `ConsultarPeriodo.aspx` con las DDJJ ya presentadas del período.
3. **Control previo (caso normal):** si ya existe una DDJJ ORIGINAL del período, NO cargar otra. Informar cuenta, período, fecha de presentación e ingreso declarado, y compararlo con el monto del papel y, si ya está, con ARBA 477320. Si hay diferencia, avisarla y no tocar nada: rectificar lo decide la persona. Terminar con **Volver**.
4. Solo si no existe ninguna DDJJ del período: clic en **Nuevo**, aceptar el aviso "DECLARE EL TOTAL DE INGRESOS GRAVADOS DEL PERÍODO".
5. En "Ingresos devengados por Actividad": Código de Actividad 477320. La alícuota (0.013 en 09/2026) y el Tipo (Principal) se completan solos. Escribir el Monto como dice "Cómo escribir el monto" y **Aplicar**.
6. Verificar en la grilla: actividad 477320, Base Imp. Total = monto, y **Tasa Determinada = monto × alícuota** (ej. 12345678 × 0,013 = 160.493,81). Si no da, no seguir.
7. Mostrar el resumen al usuario y esperar confirmación antes de **Generar Comprobante**. Después **Volver** y confirmar en ConsultarPeriodo que figura la DDJJ.
8. Archivar la boleta que da ARQui en `[Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-SEH-<soc>.pdf` (o dentro de la carpeta de año), con el mismo `<soc>` que los otros archivos de la sociedad (ej. `202609-SEH-sanar65.pdf`, `202609-SEH-quilmesfactory.pdf`). Si ya hay una, no pisarla. Si no se puede guardar desde el navegador, decirle al usuario el nombre y la carpeta.

## Control rápido sin entrar a ARQui
Para "controlá SSHH de todas" alcanza con las boletas ya guardadas en TSH (`AAAAMM-SEH-*.pdf`, a veces `AAAAMM-seh-boleta-*.pdf`):
- Sumar la columna DEUDA de las líneas del período `AAAA.MM.M.001.xx`. La `.00` es el mínimo que emite el municipio (igual para todas: $56.786 en 09/2026) y la `.01` es la diferencia que generó la DDJJ.
- Base declarada = total ÷ alícuota (0,013), redondeada al peso. Tiene que dar igual a la parte entera del papel y a ARBA 477320.
- Si una sociedad no tiene boleta en TSH, o el total no cierra con la alícuota, verificar en ARQui (paso 3).

## Resultado
Una línea por sociedad: sociedad | cuenta | período | estado (ya existía original / cargada / sin cuenta / sin monto) | monto declarado | papel | ARBA 477320 | diferencia. Después, en una línea cada uno, lo que no se pudo y por qué, y lo que queda pendiente (control contra ARBA si la DJ todavía no se presentó, boletas que faltan en TSH).

## Ejemplo
Papel con Neto Gravado 12.345.678,54 → se declara `12345678`; tasa = 12345678 × 0,013 = 160.493,81 (en la boleta: 56.786 de la `.00` + 103.707,81 de la `.01`); ARBA 477320 = $ 12.345.678,00. Coincide.
