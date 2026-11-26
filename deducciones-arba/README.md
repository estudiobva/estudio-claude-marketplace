# deducciones-arba

Deducciones de IIBB que ARBA no toma en la DJ de anticipo: percepciones bancarias sin número
(tipo O, sucursal 0, número 0) y retenciones SIRTAC con el número de liquidación repetido.
ARBA toma una por agente y descarta el resto como duplicadas.

Qué hace, por sociedad (login con la CIT de Claves_Organismos.xlsx, Playwright headless):
1. Baja el ZIP de deducciones: DFE, rol Contribuyente → Deducciones → Descarga para importar.
2. Busca la DJ mensual del período en Presentación → Dj Anticipo → Consulta DDJJ pendientes y baja
   el Archivo de deducciones (lo que ARBA tomó). El cartel de Chrome lo acepta solo.
3. Cruza y arma los lotes de **modificación (M)**: al certificado que ARBA tomó de cada banco o
   liquidación SIRTAC le pone la suma de todas. Es lo que ARBA aceptó (Argentina 10/2026).
   Darlas de alta con otro número da "Deducción duplicada".
4. Guarda en `[Sociedad]/01-Impuestos Mensuales/AAAA/AAAAMM/ARBA AAAAMM DEDUCCIONES`: ZIP RP,
   Deducciones_DJ, `CUIT-AAAAMMM-P-LOTE1_<MD5>.zip` / `-T-LOTE1_<MD5>.zip` y el Excel
   `ARBA AAAAMM <SOC> - percepciones no tomadas.xlsx`. Nunca pisa archivos.

**No sube los lotes ni presenta la DJ.** Eso lo hace la persona: Deducciones → Carga por lote,
un ZIP por vez, y después revisar "Archivo de errores" antes de borrarlo.

```
npm install
node scripts/deducciones-arba.js --nombre="Soy 60" --periodo=10/2026
node scripts/deducciones-arba.js --cuits="30714280224;30714564044" --periodo=202610
node scripts/deducciones-arba.js --cuit=30714280224 --periodo=202610 --iniciar   # inicia la DJ si falta (queda Pendiente)
node scripts/deducciones-arba.js --cuit=... --periodo=202610 --zip=RP-...zip --xls=Deducciones_DJ_...xls   # sin ARBA
node scripts/probar-cruce.js 30564252022 30711251363   # recalcula con lo archivado y compara contra los lotes manuales
```
`--ver` abre el navegador visible. `--no-archivar` deja todo en ~/Documents/BVA-salidas/deducciones-arba.

Controles del JSON: `control.cierra` = total del ZIP − deducciones de agentes de la DJ = faltante
de los lotes. `otras_diferencias` avisa si faltan retenciones, bancarias o percepciones con número
(no las arregla, hay que mirarlas a mano). Si el Deducciones_DJ se abrió y guardó en Excel deja de
ser el XML de ARBA: hay que bajarlo de nuevo.
