# ddjj-sshh-quilmes

DDJJ mensual de la Tasa de Seguridad e Higiene (TISEH, Régimen General) de las sociedades con cuenta en ARQui
Quilmes, siempre con la actividad 477320. Es solo una skill: **no tiene script**, Claude la hace en el navegador con
Claude in Chrome y la persona entra a ARQui a mano.

- **Monto para cargar:** papel de trabajo del mes, hoja `liquidacion`, Neto Gravado de "Ventas Netas (s/ Notas de
  Crédito)", parte entera sin redondear. SSHH vence (~16 del mes siguiente) antes de que se presente IIBB en ARBA.
- **Control:** contra la línea 477320 de la DJ de anticipo IIBB de ARBA cuando ya está presentada. Si ya hay una DDJJ
  original en ARQui, no carga otra: informa y marca diferencias. Hay un control rápido con las boletas de TSH, sin
  entrar a ARQui.
- **Cuenta:** `Tasas Municipales CLAVES.xlsx`, hoja QUILMES, bloque SEGURIDAD E HIGIENE, por CUIT.
- **Seguridad:** no maneja credenciales ni paga. Pide confirmación antes de **Generar Comprobante**.
- **Archivo:** la boleta va a `01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-SEH-<soc>.pdf`, sin pisar una existente.
- **Requisitos:** la extensión Claude in Chrome y acceso a la unidad compartida.
