#!/usr/bin/env python3
"""
facturacion_planilla.py — lee y marca la planilla mensual de facturación
(modelo: "Plantilla Facturacion ARCA.xlsx" de la skill facturacion-arca).

    python3 facturacion_planilla.py leer   <planilla.xlsx> [<hoja>]
    python3 facturacion_planilla.py marcar <planilla.xlsx> <hoja> <fila> <n_comprobante>

`leer` imprime un JSON con el emisor, los clientes y las filas de la hoja del
mes. No calcula nada: las fórmulas de la planilla (razón social, tipo, IVA) las
rehace facturacion-arca.js, porque un .xlsx guardado por openpyxl o por Google
no trae los valores calculados.

`marcar` completa "N° comprobante" y "¿Emitida?" = "Sí" en la fila.

Hojas: Instrucciones, Emisor, Clientes, Plantilla y una por mes. Si no se
indica la hoja y hay una sola de mes, se usa esa.
"""

import json
import sys
from datetime import date, datetime

import openpyxl

FIJAS = {"instrucciones", "emisor", "clientes", "plantilla"}

# Columnas de la hoja del mes (fila 4 = encabezados, datos desde la 5).
COLS = {
    "cuit": 1, "descripcion": 5, "neto": 6, "fecha": 9, "desde": 10, "hasta": 11,
    "vencimiento": 12, "condVenta": 13, "observaciones": 14, "nroComprobante": 15, "emitida": 16,
}
COL_NRO, COL_EMITIDA = 15, 16


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def fecha(v):
    """dd/mm/aaaa, o el texto tal cual para que el script lo valide."""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    return txt(v)


def numero(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return s  # lo reporta el script como importe inválido


def hojas_mes(wb):
    return [ws.title for ws in wb.worksheets if ws.title.strip().lower() not in FIJAS]


def leer(ruta, hoja=None):
    wb = openpyxl.load_workbook(ruta, data_only=False)
    faltan = [h for h in ("Emisor", "Clientes") if h not in wb.sheetnames]
    if faltan:
        raise SystemExit(json.dumps({"error": f"La planilla no tiene las hojas {faltan}. ¿Es la plantilla de facturación?"}))

    meses = hojas_mes(wb)
    if not hoja:
        if len(meses) != 1:
            raise SystemExit(json.dumps({
                "error": "Indicá la hoja del mes con --hoja." if meses else "La planilla no tiene hoja de mes (copiá 'Plantilla' y renombrala).",
                "hojas": meses}))
        hoja = meses[0]
    if hoja not in wb.sheetnames:
        raise SystemExit(json.dumps({"error": f"No existe la hoja '{hoja}'.", "hojas": meses}))

    em = wb["Emisor"]
    etiquetas = {
        "nombre (como figura en arca)": "nombre", "cuit": "cuit", "condición frente al iva": "condicionIva",
        "punto de venta": "puntoVenta", "actividad asociada (código)": "actividad",
        "alícuota de iva habitual": "alicuota", "carpeta donde guardar los pdf": "carpeta",
    }
    emisor = {}
    for fila in em.iter_rows(min_row=1, max_row=40):
        clave = etiquetas.get(txt(fila[0].value).lower())
        if clave and len(fila) > 1:
            v = fila[1].value
            emisor[clave] = v if (clave == "alicuota" and isinstance(v, (int, float))) else txt(v)

    clientes = []
    cl = wb["Clientes"]
    for r in range(5, cl.max_row + 1):
        cuit = txt(cl.cell(r, 1).value)
        if not cuit:
            continue
        clientes.append({
            "cuit": cuit, "razonSocial": txt(cl.cell(r, 2).value), "condicionIva": txt(cl.cell(r, 3).value),
            "descripcion": txt(cl.cell(r, 4).value), "condVenta": txt(cl.cell(r, 5).value),
            "activo": txt(cl.cell(r, 6).value),
        })

    ws = wb[hoja]
    filas = []
    for r in range(5, ws.max_row + 1):
        if txt(ws.cell(r, 5).value).upper() == "TOTAL":
            break
        cuit = txt(ws.cell(r, COLS["cuit"]).value)
        if not cuit:
            continue
        filas.append({
            "fila": r, "cuit": cuit,
            "descripcion": txt(ws.cell(r, COLS["descripcion"]).value),
            "neto": numero(ws.cell(r, COLS["neto"]).value),
            "fecha": fecha(ws.cell(r, COLS["fecha"]).value),
            "desde": fecha(ws.cell(r, COLS["desde"]).value),
            "hasta": fecha(ws.cell(r, COLS["hasta"]).value),
            "vencimiento": fecha(ws.cell(r, COLS["vencimiento"]).value),
            "condVenta": txt(ws.cell(r, COLS["condVenta"]).value),
            "observaciones": txt(ws.cell(r, COLS["observaciones"]).value),
            "nroComprobante": txt(ws.cell(r, COLS["nroComprobante"]).value),
            "emitida": txt(ws.cell(r, COLS["emitida"]).value),
        })

    return {"hoja": hoja, "hojas": meses, "emisor": emisor, "clientes": clientes, "filas": filas}


def marcar(ruta, hoja, fila, nro):
    wb = openpyxl.load_workbook(ruta)
    ws = wb[hoja]
    fila = int(fila)
    ws.cell(fila, COL_NRO).value = nro
    ws.cell(fila, COL_EMITIDA).value = "Sí"
    wb.save(ruta)
    return {"ok": True, "hoja": hoja, "fila": fila, "nroComprobante": nro}


if __name__ == "__main__":
    if len(sys.argv) < 3 or sys.argv[1] not in ("leer", "marcar"):
        raise SystemExit(__doc__)
    if sys.argv[1] == "leer":
        out = leer(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    else:
        if len(sys.argv) != 6:
            raise SystemExit(__doc__)
        out = marcar(*sys.argv[2:6])
    print(json.dumps(out, ensure_ascii=False))
