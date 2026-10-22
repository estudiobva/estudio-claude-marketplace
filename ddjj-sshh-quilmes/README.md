# ddjj-sshh-quilmes

Controla en **ARQui** (Municipio de Quilmes) la **DDJJ mensual de la Tasa de Seguridad e Higiene** (TISEH, Régimen
General) de una, varias o todas las sociedades, y archiva la boleta donde el equipo la guarda:

```
[Sociedad]/01-Impuestos Mensuales/AAAAMM/TSH/AAAAMM-SEH-<soc>.pdf
```

El login es automático: usuario y clave del grupo MUNICIPAL (Quilmes) de `Claves_Organismos.xlsx`. La cuenta sale
del bloque SEGURIDAD E HIGIENE de `Tasas Municipales CLAVES.xlsx`, por CUIT. Solo consulta y descarga: no presenta
DDJJ ni paga.

## Instalacion

```bash
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

Usa el mismo `~/.fisco-ar/.env` que los demas plugins (`CLAVES_ORGANISMOS_PATH`, `CLAVES_ORGANISMOS_PASSWORD` y, en
Mac, `BVA_UNIDAD_PATH`). Ver `.env.example`.

## Uso

```bash
node scripts/ddjj-sshh-quilmes.js --nombre="<sociedad>" [--periodo=AAAAMM]
node scripts/ddjj-sshh-quilmes.js --nombres="<A>;<B>" [--periodo=AAAAMM]
node scripts/ddjj-sshh-quilmes.js --todas [--periodo=AAAAMM]
```

- **Período:** por defecto el mes anterior.
- **Qué controla:** que lo declarado en ARQui sea solo la actividad 477320 y que coincida con el papel de trabajo del
  mes (hoja `liquidacion`, Neto Gravado de "Ventas Netas (s/ Notas de Crédito)", parte entera) y con la línea 477320
  de la DJ de anticipo IIBB de ARBA, si ya está en la carpeta del mes. Menos de $1 de diferencia no cuenta.
- **Boleta:** baja el Recibo de Pago de la DDJJ y lo archiva en TSH sin pisar uno existente (`--reemplazar` para
  pisarlo, `--no-archivar` para dejarlo solo en la salida).
- **`--todas`:** recorre las cuentas de SSHH vinculadas al usuario de ARQui y suma la hoja **Control planilla**
  (cuentas que faltan en la planilla de tasas, CUIT distinto, cuentas de la planilla no vinculadas).
- **Salida:** `~/Documents/BVA-salidas/ddjj-sshh-quilmes/AAAA-MM-DD/` (`BVA_SALIDAS_PATH`): resumen
  `DDJJ SSHH QUILMES AAAAMM ....xlsx`, la Declaración y la boleta de cada cuenta, y JSON por stdout.
- **Bloqueos:** ARQui contesta 403 un rato si recibe muchos pedidos o logins seguidos. El script espera 5 minutos
  (`ARQUI_ESPERA_BLOQUEO_MS`) y reintenta una vez; si sigue, corta y avisa.
