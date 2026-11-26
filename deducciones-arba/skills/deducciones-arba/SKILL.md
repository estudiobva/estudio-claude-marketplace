---
name: deducciones-arba
description: "Deducciones de Ingresos Brutos que ARBA NO toma en la DJ de anticipo mensual: percepciones bancarias informadas sin numero de certificado (tipo O, sucursal 0, numero 0) y retenciones SIRTAC con el numero de liquidacion repetido, que ARBA descarta como 'duplicadas' dejando una sola por agente. Baja el ZIP de deducciones (RP-CUIT-AAAAMMC.zip) y el Archivo de deducciones de la DJ (Deducciones_DJ_<nro>.xls), los cruza y arma los lotes de MODIFICACION (operacion M) CUIT-AAAAMMM-P-LOTE1_<MD5>.zip y -T-LOTE1_<MD5>.zip para la Carga por lote, mas el Excel 'percepciones no tomadas'. Login automatico con la CIT de la planilla de claves (Playwright). Usar SIEMPRE que pidan 'ARBA no toma todas las percepciones', 'certificados duplicados en ARBA', 'deducciones de la DJ de ARBA', 'armar el lote de deducciones', 'TXT de deducciones para cargar por lote', 'percepciones bancarias que no toma', 'Deduccion duplicada', 'lo mismo para [sociedad]' en este circuito. NO sube los lotes ni presenta la DJ."
---

# Deducciones que ARBA no toma (lotes M)

```bash
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deducciones-arba.js --nombre="<sociedad>" --periodo=10/2026
cd "${CLAUDE_PLUGIN_ROOT:-.}" && node scripts/deducciones-arba.js --cuits="<CUIT1>;<CUIT2>" --periodo=202610
```

| Argumento | Para que |
|---|---|
| `--nombre` / `--cuit`, `--nombres` / `--cuits` | Sociedad(es), como en los otros scripts. |
| `--periodo` | `MM/AAAA` o `AAAAMM`. Sin periodo: el mes anterior. |
| `--iniciar` | Si la DJ del mes no esta iniciada, la inicia (queda Pendiente, no se presenta). Sin el flag sale `sin_dj` y no toca nada. Preguntar antes de usarlo. |
| `--no-archivar` | Deja todo en `~/Documents/BVA-salidas/deducciones-arba/<CUIT>/<AAAAMM>/`. |
| `--zip` / `--xls` | Cruce con archivos ya bajados (una sociedad, no entra a ARBA). |
| `--ver` | Navegador visible. |

## El problema y la solucion que ARBA acepta

Los bancos informan percepciones sin numero (tipo O, sucursal 0, numero 0) y el SIRTAC a veces
repite el numero de liquidacion. ARBA toma UNA por agente y descarta el resto.

- Cargarlas de **alta** con otro numero da **"Deduccion duplicada"** (probado: cambiar numero, y
  sucursal + numero, tampoco).
- Lo que ARBA **acepta** es **modificar (M)** el certificado que tomo de cada agente y ponerle la
  suma de todas sus lineas (Argentina 10/2026: el total del contribuyente cerro exacto).

Diseno de registro vigente (`DISENIODEREGISTROiibbddjjweb.pdf`, el de 49 caracteres es viejo y da
"Longitud de linea invalida"): percepciones 53 car. (CUIT 99-99999999-9, fecha, tipo, letra,
sucursal 5, emision 8, importe 14 con coma, operacion); SIRTAC 70 car. (CUIT, AAAA, MM, fecha,
liquidacion 20, signo, monto 19, operacion). Nombre: `CUIT-AAAAMMM-<P|T>-LOTE1_<MD5 del zip>.zip`
con un TXT adentro llamado igual sin el MD5. El nombre del ZIP de descarga (`RP-...C.zip`) da
"no tiene la longitud requerida".

## Que hace, por sociedad

1. Entra con la CIT (columnas ARBA de `Claves_Organismos.xlsx`). Un solo intento: si ARBA
   rechaza, corta con `arba_login_rechazado` (reintentar bloquea la CIT). Si aparece
   "Seleccione un rol" elige **Contribuyente** (con rol Agente no esta el menu Deducciones).
2. DFE → Deducciones → Descarga para importar → baja el ZIP del periodo.
3. IBPresentaciones → Consulta DDJJ pendientes del mes → Ver detalle → Deducciones → Archivo de
   deducciones (acepta el cartel nativo del navegador). Si no hay DJ pendiente: `sin_dj`.
4. Cruza (TXT CP/CT contra hojas Percepciones / Retenciones Sirtac) y arma los lotes M.
5. Guarda en `[Sociedad]/01-Impuestos Mensuales/AAAA/AAAAMM/ARBA AAAAMM DEDUCCIONES`: ZIP RP,
   Deducciones_DJ, lotes y `ARBA AAAAMM <SOC> - percepciones no tomadas.xlsx` (hojas No tomadas,
   Bancarias sin nro (todas), SIRTAC no tomada, Resumen lotes). Nunca pisa archivos.

Controles en el JSON: `control.cierra` (total del ZIP - deducciones de agentes de la DJ = faltante
de los lotes) y `otras_diferencias` (retenciones o bancarias que faltan, o percepciones con numero
no tomadas: no se arreglan solas, avisar). Si el `Deducciones_DJ` fue abierto y guardado en Excel
ya no es el XML de ARBA: bajarlo de nuevo.

## Respuesta al usuario

Por sociedad: faltante total, tabla por banco (certificado que tomo ARBA, importe tomado, lineas,
nuevo importe) y SIRTAC si hay, los ZIP a subir y el total al que tiene que llegar la DJ
("Deducciones declaradas" agentes + contribuyente = total del ZIP). Recordar: subir cada ZIP por
**Carga por lote**, esperar unos minutos, y si aparece **Archivo de errores** bajarlo ANTES de tocar
"Borrar errores". El aviso "archivo recibido correctamente" no significa que lo haya tomado.
