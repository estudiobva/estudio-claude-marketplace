#!/usr/bin/env python3
"""
Cruza la hoja "Ventas Zetti" contra "Ventas Portal IVA" (o "Ventas Portal iva")
de un papel de trabajo de IVA/IIBB, y vuelca las DIFERENCIAS en una hoja nueva
"cruze" agregada al final del mismo archivo (copia; el original no se toca).

Tres formas de usarlo:

    # 1) Carpeta IVA del mes (lo normal desde 10/2026): lee los dos archivos que
    #    dejan zetti-iva-ventas.js y portal-iva-descarga.js y deja el cruce al lado.
    python3 cruzar_ventas.py --cuit=<CUIT> --periodo=09/2026
        lee   [IVA del mes]/AAAAMM - ZETTI - VENTAS.xlsx        (hoja "Ventas Zetti")
              [IVA del mes]/AAAAMM - PORTAL IVA - VENTAS.xlsx   (hoja "Ventas Portal IVA")
        deja  [IVA del mes]/AAAAMM - CRUCE ZETTI VS PORTAL IVA.xlsx

    # 2) Dos archivos sueltos (mismos formatos, o cualquier libro con esas hojas)
    python3 cruzar_ventas.py --zetti=<zetti.xlsx> --portal=<portal.xlsx> [--salida=<cruce.xlsx>]

    # 3) Papel de trabajo armado a mano, con las dos hojas (modo original)
    python3 cruzar_ventas.py <entrada.xlsx> <salida.xlsx>

Opciones: --tolerancia 0.5 | --si-existe=abortar (default) o reemplazar, si el
cruce ya esta (en los modos 1 y 2) | --salida=<ruta> tambien en el modo 1, para
no escribir en la unidad | --incluir-nc-controlador (modos 1 y 2).

En los modos 1 y 2 el archivo de Zetti es el Subdiario 5.6.5 completo, que
incluye las NC B a Consumidor Final del controlador fiscal (PV 1 / PV 2) que
el papel de trabajo armado a mano no traia: se separan del cruce y se listan
en una hoja aparte (ver separar_nc_controlador).

La carpeta IVA se busca como rutas.carpetaIva de bva-playwright: carpeta de la
sociedad por CUIT en la unidad compartida (BVA_UNIDAD_PATH, o I:\\...), mes
suelto (AAAAMM) o dentro del año (AAAA/AAAAMM), y subcarpeta IVA o AAAAMM-IVA.

Imprime un JSON por stdout con el resumen (totales, cantidad de filas de
diferencia, chequeo de sanidad, y la lista completa de filas de diferencia)
para que quien use este script pueda razonar sobre el patron de las
diferencias y escribir las Observaciones antes de entregar el archivo. Ver
SKILL.md para el paso de interpretacion. El chequeo de sanidad compara la
diferencia que explica el cruce contra:
  - modo 3: las filas TOTAL ZETTI / TOTAL PORTAL IVA del papel de trabajo;
  - modos 1 y 2: el total de cada archivo, sumado de TODAS sus filas con
    importe (no solo de las que el cruce entendio), asi una fila que el
    parseo descarte aparece como diferencia en vez de perderse.

Por que existe este script en vez de escribirlo de cero cada vez: el mismo
cruce se hizo a mano, sociedad por sociedad, y cada archivo tiene un layout de
columnas levemente distinto en "Ventas Portal IVA" (a veces trae "Importe No
Gravado", a veces no; a veces viene con todas las alicuotas de IVA discriminadas
0%/2.5%/5%/10.5%/21%/27%, a veces solo 10.5/21/27; a veces trae percepciones y
fecha de vencimiento de pago intercaladas). Este script detecta las columnas
por NOMBRE de encabezado en vez de por posicion fija, para no romperse cuando
cambia el layout.
"""
import argparse
import glob
import os
import sys
import json
import re
import shutil
import time
import unicodedata
import zipfile
from collections import defaultdict
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def norm(s):
    """Normaliza un texto de encabezado: sin acentos, minusculas, sin espacios extra."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def r2(x):
    return round(x, 2) if x is not None else 0.0


def to_int(v):
    """Punto de venta / numero como entero ('0006' -> 6, 6.0 -> 6), o el valor tal cual."""
    if isinstance(v, bool) or v is None:
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return v


def find_sheet(wb, must_contain_all):
    """Busca la primer hoja cuyo nombre normalizado contenga todas las palabras dadas."""
    for name in wb.sheetnames:
        n = norm(name)
        if all(word in n for word in must_contain_all):
            return name
    return None


# ---------------------------------------------------------------------------
# Parseo de "Ventas Zetti"
# ---------------------------------------------------------------------------

NRO_COMBINADO = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")
CUIT_RE = re.compile(r"\b(\d{2})-?(\d{8})-?(\d)\b")


def parse_zetti(ws):
    """
    Devuelve (rows, totales_declarados, info) donde:
      rows: lista de dicts {fecha, tc, m, punto, numero, cliente, exen, grav, iva, total}
      totales_declarados: dict con TOTAL ZETTI / TOTAL PORTAL IVA / diferencia si el
        archivo los trae (para chequeo de sanidad), o None si no se encuentran.
      info: razon social y CUIT del encabezado (si estan), filas descartadas y
        la suma cruda de Exen/Grav/IVA/Total de todas las filas con importe.

    Conviven dos layouts de la misma hoja:
      - el que se pegaba a mano en el papel de trabajo: 'Nro. Comp.' trae SOLO
        el Punto de Venta (a pesar del nombre) y el Numero de Comprobante va en
        la columna SIN nombre que le sigue; a veces con 'P.IB' antes de Total.
      - el de zetti-iva-ventas.js (AAAAMM - ZETTI - VENTAS.xlsx): 'Nro. Comp.'
        trae los dos juntos, '0006-00000070', y la columna siguiente es 'Cliente'.
    Se decide fila por fila mirando el valor de 'Nro. Comp.'.
    """
    header_row = None
    col = {}
    for r in range(1, 15):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        normed = [norm(v) for v in vals]
        if "fecha" in normed and "tc" in normed and "exen" in normed:
            header_row = r
            for idx, h in enumerate(normed, start=1):
                if h:
                    col[h] = idx
            # la columna de Numero de Comprobante no tiene encabezado: es la
            # que sigue inmediatamente a 'Nro. Comp.' (si tiene encabezado,
            # es 'Cliente' y el numero viene junto con el PV)
            if "nro. comp." in col:
                siguiente = col["nro. comp."]  # posicion 0-based de la columna que sigue
                if siguiente >= len(normed) or normed[siguiente] == "":
                    col["__numero__"] = col["nro. comp."] + 1
            break
    if header_row is None:
        raise ValueError("No encontre la fila de encabezado en Ventas Zetti "
                          "(busco una fila con 'Fecha', 'TC' y 'Exen').")

    def g(row_vals, key, default=None):
        idx = col.get(key)
        if idx is None:
            return default
        v = row_vals[idx - 1] if idx - 1 < len(row_vals) else None
        return v

    info = {"razon_social": None, "cuit": None, "filas_descartadas": [],
            "suma_archivo": {"exen": 0.0, "grav": 0.0, "iva": 0.0, "total": 0.0}}
    for r in range(1, header_row):
        for v in (ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)):
            if not isinstance(v, str) or not v.strip():
                continue
            m = CUIT_RE.search(v)
            if m and info["cuit"] is None:
                info["cuit"] = "".join(m.groups())
                # titulo de una sola celda: "<razon social>  -  CUIT NN-...  -  Subdiario ..."
                antes = re.split(r"\s+-\s+CUIT\b", v, flags=re.I)[0].strip(" -")
                if info["razon_social"] is None and antes != v.strip() and len(antes) > 8:
                    info["razon_social"] = antes
            elif (info["razon_social"] is None and len(v.strip()) > 8 and ":" not in v
                  and norm(v) not in ("t&s web", "subdiario de iva ventas")):
                info["razon_social"] = v.strip()

    def num(v):
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0

    rows = []
    totales = {}
    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        fecha = g(vals, "fecha")
        tc = g(vals, "tc")
        # fila de totales / separador: buscar "TOTAL ZETTI" / "TOTAL PORTAL IVA"
        # en cualquier celda de texto de la fila
        texto_fila = " ".join(norm(v) for v in vals if isinstance(v, str))
        if not texto_fila.startswith("total") and isinstance(g(vals, "total"), (int, float)):
            for k in ("exen", "grav", "iva", "total"):
                info["suma_archivo"][k] += num(g(vals, k))
        if "total zetti" in texto_fila:
            totales["zetti"] = {"exen": g(vals, "exen"), "grav": g(vals, "grav"),
                                 "iva": g(vals, "iva"), "total": g(vals, "total")}
            continue
        if "total portal iva" in texto_fila:
            totales["portal"] = {"exen": g(vals, "exen"), "grav": g(vals, "grav"),
                                  "iva": g(vals, "iva"), "total": g(vals, "total")}
            continue
        if fecha is None and tc is None:
            continue
        if fecha is None or not isinstance(fecha, (datetime, date)):
            continue
        punto = g(vals, "nro. comp.")
        m = NRO_COMBINADO.match(punto) if isinstance(punto, str) else None
        if m:
            punto, numero = int(m.group(1)), int(m.group(2))
        else:
            punto, numero = to_int(punto), to_int(g(vals, "__numero__"))
            if punto is not None and numero is None and "__numero__" not in col:
                # comprobante sin numero en Zetti (ej. FV con PV 9999999999999999,
                # sin CAE): sigue en el cruce, distinguido por fila
                numero = f" s/n (fila {r})"
        if punto is None or numero is None:
            info["filas_descartadas"].append(r)
            continue
        d = fecha.date() if isinstance(fecha, datetime) else fecha
        rows.append({
            "fecha": d, "tc": tc, "m": g(vals, "m"), "punto": punto, "numero": numero,
            "cliente": (g(vals, "cliente") or "").strip() if isinstance(g(vals, "cliente"), str) else "",
            "resp": norm(g(vals, "resp")),
            "exen": g(vals, "exen") or 0, "grav": g(vals, "grav") or 0,
            "iva": g(vals, "iva") or 0, "total": g(vals, "total") or 0,
        })
    info["suma_archivo"] = {k: r2(v) for k, v in info["suma_archivo"].items()}
    return rows, (totales or None), info


# ---------------------------------------------------------------------------
# Parseo de "Ventas Portal IVA"
# ---------------------------------------------------------------------------

# Mapeo de nombre de encabezado normalizado -> clave interna. Se acumulan acá
# todas las variantes vistas hasta ahora; si aparece un archivo con un nombre
# de columna nuevo, agregar la variante acá en vez de tocar el resto del script.
PORTAL_HEADER_MAP = {
    "fecha de emision": "fecha",
    "tipo de comprobante": "tipo",
    "punto de venta": "punto",
    "numero de comprobante": "numero",
    "importe total": "total",
    "importe no gravado": "no_gravado",
    "importe exento": "exento",
    "total neto gravado": "grav_total",
    "total iva": "iva_total",
}
# columnas de gravado/iva por alicuota, usadas SOLO como fallback si el
# archivo no trae las columnas "Total Neto Gravado" / "Total IVA" ya sumadas
GRAV_PREFIX = "neto gravado iva"
IVA_PREFIX = "importe iva"


def parse_portal(ws):
    header_row = None
    col = {}
    grav_fallback_cols = []
    iva_fallback_cols = []
    for r in range(1, 6):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        normed = [norm(v) for v in vals]
        if "fecha de emision" in normed and "punto de venta" in normed:
            header_row = r
            for idx, h in enumerate(normed, start=1):
                if h in PORTAL_HEADER_MAP:
                    col[PORTAL_HEADER_MAP[h]] = idx
                elif h.startswith(GRAV_PREFIX):
                    grav_fallback_cols.append(idx)
                elif h.startswith(IVA_PREFIX) and "por" not in h:
                    iva_fallback_cols.append(idx)
            break
    if header_row is None:
        raise ValueError("No encontre la fila de encabezado en Ventas Portal IVA "
                          "(busco 'Fecha de Emision' y 'Punto de Venta').")
    if "punto" not in col or "numero" not in col or "total" not in col:
        raise ValueError(f"Faltan columnas clave en Ventas Portal IVA. Encontre: {col}")

    def g(row_vals, key, default=0):
        idx = col.get(key)
        if idx is None:
            return default
        v = row_vals[idx - 1] if idx - 1 < len(row_vals) else None
        return v if v is not None else default

    def importes(vals):
        exen = (g(vals, "no_gravado", 0) or 0) + (g(vals, "exento", 0) or 0)
        if "grav_total" in col:
            grav = g(vals, "grav_total", 0) or 0
        else:
            grav = sum((vals[i - 1] or 0) for i in grav_fallback_cols if i - 1 < len(vals))
        if "iva_total" in col:
            iva = g(vals, "iva_total", 0) or 0
        else:
            iva = sum((vals[i - 1] or 0) for i in iva_fallback_cols if i - 1 < len(vals))
        return exen, grav, iva, g(vals, "total", 0) or 0

    rows = []
    info = {"filas_descartadas": [],
            "suma_archivo": {"exen": 0.0, "grav": 0.0, "iva": 0.0, "total": 0.0}}
    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        fecha = g(vals, "fecha", None)
        # la fila TOTAL (SUM) trae numeros si el archivo se abrio y guardo en Excel
        es_total = any(isinstance(v, str) and norm(v).startswith("total") for v in vals[:3])
        es_dato = isinstance(g(vals, "total", None), (int, float)) and not es_total
        if es_dato:
            for k, v in zip(("exen", "grav", "iva", "total"), importes(vals)):
                info["suma_archivo"][k] += v
        if fecha is None or not isinstance(fecha, (datetime, date)):
            # la hoja Portal IVA siempre termina en un bloque de totales /
            # resumen por tipo de comprobante sin fecha: ahi termina la data
            if es_dato:
                info["filas_descartadas"].append(r)
            continue
        d = fecha.date() if isinstance(fecha, datetime) else fecha
        exen, grav, iva, total = importes(vals)
        rows.append({
            "fecha": d, "tipo": g(vals, "tipo", None), "punto": to_int(g(vals, "punto", None)),
            "numero": to_int(g(vals, "numero", None)), "exen": exen, "grav": grav, "iva": iva,
            "total": total,
        })
    info["suma_archivo"] = {k: r2(v) for k, v in info["suma_archivo"].items()}
    return rows, info


# ---------------------------------------------------------------------------
# Cruce
# ---------------------------------------------------------------------------

def separar_nc_controlador(zetti_rows, portal_rows):
    """
    Separa de Zetti las NC B a Consumidor Final del controlador fiscal (PV 1 /
    2 / 3 / 4 con numeracion de controlador, ej. 0001-00908007). El Subdiario
    5.6.5 las lista, pero ya estan netas dentro del cierre Z: en papeles de 07/2026 y
    09/2026 sumaban exactamente la diferencia del mes con los Z cerrando en 0.
    Por eso quien armaba el papel de trabajo las borraba a mano de "Ventas
    Zetti" (varias sociedades, 07 y 08/2026).

    Se separa una fila si: es NC letra B a consumidor final, su PV no es de
    tickets Z, el comprobante no esta en Portal IVA, y en ese PV el Portal no
    tiene NINGUNA NC B (tipo 8). Asi una NC B electronica que no llego al
    Portal (PV con otras NC B informadas) sigue en el cruce como solo_zetti,
    y tambien cualquier factura o NC a un no-consumidor-final del mismo PV
    (hay sociedades con una FV A electronica en el PV del controlador).
    Devuelve (filas para el cruce, filas separadas).
    """
    puntos_z = {r["punto"] for r in zetti_rows if r["tc"] == "Z"}
    claves_portal = {(r["punto"], r["numero"]) for r in portal_rows}
    pv_con_nc_b = {r["punto"] for r in portal_rows if to_int(r["tipo"]) == 8}

    def separar(r):
        return (r["tc"] == "NC" and norm(r["m"]) == "b"
                and (r["resp"].startswith("cons") or "consumidor final" in norm(r["cliente"]))
                and r["punto"] not in puntos_z and r["punto"] not in pv_con_nc_b
                and (r["punto"], r["numero"]) not in claves_portal)

    return ([r for r in zetti_rows if not separar(r)],
            [r for r in zetti_rows if separar(r)])


def cruzar(zetti_rows, portal_rows, tolerancia=0.5):
    z_puntos = sorted({r["punto"] for r in zetti_rows if r["tc"] == "Z"})

    diffs = []  # cada uno: dict con fecha, concepto, z_exen,p_exen,d_exen,... , tipo ('directo'|'zticket')

    # --- match directo: comprobantes individuales (todo lo que no es TC='Z') ---
    z_direct = [r for r in zetti_rows if r["punto"] not in z_puntos]
    p_direct = [r for r in portal_rows if r["punto"] not in z_puntos]
    p_dict = {}
    for r in p_direct:
        p_dict[(r["punto"], r["numero"])] = r

    matched_keys = set()
    for zr in z_direct:
        key = (zr["punto"], zr["numero"])
        matched_keys.add(key)
        pr = p_dict.get(key)
        if pr is None:
            diffs.append(_row(zr["fecha"], f"Comp. PV{key[0]} Nro{key[1]}"
                               + (f" ({zr['cliente']})" if zr["cliente"] else ""),
                               zr["exen"], 0, zr["grav"], 0, zr["iva"], 0, zr["total"], 0,
                               "directo", "solo_zetti"))
            continue
        d_exen, d_grav, d_iva, d_total = (r2(zr["exen"] - pr["exen"]), r2(zr["grav"] - pr["grav"]),
                                           r2(zr["iva"] - pr["iva"]), r2(zr["total"] - pr["total"]))
        if max(abs(d_exen), abs(d_grav), abs(d_iva), abs(d_total)) > tolerancia:
            diffs.append(_row(zr["fecha"], f"Comp. PV{key[0]} Nro{key[1]}"
                               + (f" ({zr['cliente']})" if zr["cliente"] else ""),
                               zr["exen"], pr["exen"], zr["grav"], pr["grav"],
                               zr["iva"], pr["iva"], zr["total"], pr["total"],
                               "directo", "diferencia_importe"))

    for pr in p_direct:
        key = (pr["punto"], pr["numero"])
        if key in matched_keys:
            continue
        diffs.append(_row(pr["fecha"], f"Comp. PV{key[0]} Nro{key[1]}",
                           0, pr["exen"], 0, pr["grav"], 0, pr["iva"], 0, pr["total"],
                           "directo", "solo_portal"))

    # --- match agregado de tickets Z: por (fecha, punto de venta) ---
    z_agg = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for r in zetti_rows:
        if r["punto"] in z_puntos:
            k = (r["fecha"], r["punto"])
            z_agg[k][0] += r["exen"]; z_agg[k][1] += r["grav"]
            z_agg[k][2] += r["iva"]; z_agg[k][3] += r["total"]
    p_agg = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for r in portal_rows:
        if r["punto"] in z_puntos:
            k = (r["fecha"], r["punto"])
            p_agg[k][0] += r["exen"]; p_agg[k][1] += r["grav"]
            p_agg[k][2] += r["iva"]; p_agg[k][3] += r["total"]

    for k in sorted(set(z_agg) | set(p_agg)):
        z = z_agg[k]; p = p_agg[k]
        d = [r2(z[i] - p[i]) for i in range(4)]
        if max(abs(x) for x in d) > tolerancia:
            fecha, punto = k
            diffs.append(_row(fecha, f"Tickets Z - Consumidor Final (PV {punto})",
                               z[0], p[0], z[1], p[1], z[2], p[2], z[3], p[3],
                               "zticket", None, punto=punto))

    diffs.sort(key=lambda x: (x["fecha"], x["concepto"]))

    # heuristica de compensacion: para las filas 'zticket', ver si existe otra
    # fecha del MISMO punto de venta cuya Dif.Total sea aprox. la opuesta
    for i, row in enumerate(diffs):
        if row["tipo"] != "zticket":
            continue
        if abs(row["d_total"]) <= tolerancia * 4:
            row["patron"] = "redondeo"
            continue
        candidatos = [d2 for j, d2 in enumerate(diffs)
                      if j != i and d2["tipo"] == "zticket" and d2.get("punto") == row.get("punto")
                      and d2["fecha"] != row["fecha"]]
        compensa = [d2 for d2 in candidatos if abs(d2["d_total"] + row["d_total"]) <= max(2.0, tolerancia * 4)]
        if compensa:
            row["patron"] = "corrimiento_fecha"
            row["compensa_con"] = sorted(d2["fecha"].isoformat() for d2 in compensa)
        else:
            row["patron"] = "sin_compensar"

    for row in diffs:
        if row["tipo"] == "directo":
            row["patron"] = row.pop("motivo")

    totales_diff = {
        "exen": r2(sum(d["d_exen"] for d in diffs)),
        "grav": r2(sum(d["d_grav"] for d in diffs)),
        "iva": r2(sum(d["d_iva"] for d in diffs)),
        "total": r2(sum(d["d_total"] for d in diffs)),
    }
    return diffs, z_puntos, totales_diff


def _row(fecha, concepto, z_exen, p_exen, z_grav, p_grav, z_iva, p_iva, z_total, p_total,
         tipo, motivo, punto=None):
    d = {
        "fecha": fecha, "concepto": concepto,
        "z_exen": r2(z_exen), "p_exen": r2(p_exen), "d_exen": r2(z_exen - p_exen),
        "z_grav": r2(z_grav), "p_grav": r2(p_grav), "d_grav": r2(z_grav - p_grav),
        "z_iva": r2(z_iva), "p_iva": r2(p_iva), "d_iva": r2(z_iva - p_iva),
        "z_total": r2(z_total), "p_total": r2(p_total), "d_total": r2(z_total - p_total),
        "tipo": tipo,
    }
    if motivo is not None:
        d["motivo"] = motivo
    if punto is not None:
        d["punto"] = punto
    return d


AUTO_OBS = {
    "solo_zetti": "Comprobante solo en Ventas Zetti",
    "solo_portal": "Comprobante solo en Portal IVA",
    "diferencia_importe": "Diferencia de importes entre ambos reportes",
    "redondeo": "Diferencia minima de redondeo, sin relevancia",
}


def auto_observacion(row):
    if row["patron"] in AUTO_OBS:
        return AUTO_OBS[row["patron"]]
    if row["patron"] == "corrimiento_fecha":
        otras = ", ".join(row["compensa_con"])
        return f"Diferencia de fecha entre sistemas - se compensaria con {otras} (revisar)"
    if row["patron"] == "sin_compensar":
        return "Diferencia real, no compensa con otra fecha del mismo punto de venta - revisar (posible origen de la diferencia total)"
    return ""


# ---------------------------------------------------------------------------
# Escritura de la hoja "cruze"
# ---------------------------------------------------------------------------

def escribir_hoja_cruze(wb, diffs, sociedad, periodo, nota=None):
    if "cruze" in wb.sheetnames:
        del wb["cruze"]
    ws = wb.create_sheet("cruze")

    headers = ["Fecha", "Concepto", "Exento Zetti", "Exento Portal", "Dif. Exento",
               "Gravado Zetti", "Gravado Portal", "Dif. Gravado",
               "IVA Zetti", "IVA Portal", "Dif. IVA",
               "Total Zetti", "Total Portal", "Dif. Total", "Observacion"]

    title_font = Font(bold=True, size=14, color="FFFFFF")
    title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    total_font = Font(bold=True)
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money_fmt = "#,##0.00"
    red_font = Font(color="C00000")

    ws.merge_cells("A1:O1")
    titulo = f"CRUCE VENTAS ZETTI vs VENTAS PORTAL IVA - {sociedad} - {periodo} (solo diferencias)"
    ws["A1"] = titulo
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24
    if nota:
        ws.merge_cells("A2:O2")
        ws["A2"] = nota
        ws["A2"].font = Font(italic=True, color="404040")

    header_row_idx = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[header_row_idx].height = 30

    r = header_row_idx + 1
    totals = [0.0] * 12
    for row in diffs:
        ws.cell(row=r, column=1, value=row["fecha"]).number_format = "DD/MM/YYYY"
        ws.cell(row=r, column=2, value=row["concepto"])
        vals = [row["z_exen"], row["p_exen"], row["d_exen"], row["z_grav"], row["p_grav"], row["d_grav"],
                row["z_iva"], row["p_iva"], row["d_iva"], row["z_total"], row["p_total"], row["d_total"]]
        for i, v in enumerate(vals):
            cell = ws.cell(row=r, column=3 + i, value=v)
            cell.number_format = money_fmt
            totals[i] += v
        ws.cell(row=r, column=15, value=row.get("observacion", auto_observacion(row)))
        for c in range(1, 16):
            ws.cell(row=r, column=c).border = border
        for col_idx in (5, 8, 11, 14):
            cell = ws.cell(row=r, column=col_idx)
            if abs(cell.value) > 0.5:
                cell.font = red_font
        r += 1

    ws.cell(row=r, column=1, value="TOTAL DIFERENCIAS")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for i, v in enumerate(totals):
        cell = ws.cell(row=r, column=3 + i, value=round(v, 2))
        cell.number_format = money_fmt
    for c in range(1, 16):
        cell = ws.cell(row=r, column=c)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border

    widths = [12, 40, 14, 14, 13, 14, 14, 13, 13, 13, 12, 14, 14, 13, 70]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{header_row_idx + 1}"
    return totals


HOJA_NC_SEPARADAS = "NC controlador (fuera)"


def escribir_hoja_nc_separadas(wb, rows):
    """Las NC que separar_nc_controlador saco del cruce, para que se vean."""
    ws = wb.create_sheet(HOJA_NC_SEPARADAS)
    ws.merge_cells("A1:H1")
    ws["A1"] = ("NC B a Consumidor Final de Zetti en puntos de venta sin comprobantes en "
                "Portal IVA: quedan fuera del cruce (ya estan dentro del cierre Z)")
    ws["A1"].font = Font(bold=True)
    headers = ["Fecha", "PV", "Numero", "Cliente", "Exento", "Gravado", "IVA", "Total"]
    header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    r = 4
    for row in rows:
        vals = [row["fecha"], row["punto"], row["numero"], row["cliente"],
                r2(row["exen"]), r2(row["grav"]), r2(row["iva"]), r2(row["total"])]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if c == 1:
                cell.number_format = "DD/MM/YYYY"
            elif c >= 5:
                cell.number_format = "#,##0.00"
        r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    for c, k in zip(range(5, 9), ("exen", "grav", "iva", "total")):
        cell = ws.cell(row=r, column=c, value=r2(sum(x[k] for x in rows)))
        cell.number_format = "#,##0.00"
        cell.font = Font(bold=True)
    for i, w in enumerate([12, 6, 12, 30, 14, 14, 13, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


# ---------------------------------------------------------------------------
# Carpeta IVA del periodo (misma logica que lib/rutas-bva.js de bva-playwright)
# ---------------------------------------------------------------------------

RAIZ_DEFECTO = "I:\\Unidades compartidas\\BVA - Sociedades Farmaceuticas"
NOMBRE_UNIDAD = "BVA - Sociedades Farmaceuticas"
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def fallar(msg, **extra):
    print(json.dumps({"error": msg, **extra}, ensure_ascii=False, indent=2, default=str))
    sys.exit(2)


def raiz_unidad():
    env = (os.environ.get("BVA_UNIDAD_PATH") or "").strip()
    if env:
        if not os.path.isdir(env):
            fallar(f'BVA_UNIDAD_PATH apunta a "{env}", que no existe. Verifica que Google Drive este montado.')
        return env
    candidatas = [RAIZ_DEFECTO]
    # el script suele vivir adentro de la misma unidad (0 - ESTUDIO/.../skills/...)
    p = os.path.dirname(os.path.abspath(__file__))
    while p != os.path.dirname(p):
        if os.path.basename(p) == NOMBRE_UNIDAD:
            candidatas.append(p)
            break
        p = os.path.dirname(p)
    for c in candidatas:
        if os.path.isdir(c):
            return c
    # Mac: Drive montado por cuenta. Puede haber mas de una cuenta con la
    # misma unidad compartida: en ese caso no se elige a ciegas.
    en_mac = glob.glob(os.path.expanduser(
        f"~/Library/CloudStorage/GoogleDrive-*/Unidades compartidas/{NOMBRE_UNIDAD}"))
    if len(en_mac) == 1:
        return en_mac[0]
    if en_mac:
        fallar("La unidad compartida aparece en mas de una cuenta de Drive: defini BVA_UNIDAD_PATH.",
               candidatas=en_mac)
    fallar(f'No encuentro la unidad compartida (probe {RAIZ_DEFECTO}). Defini BVA_UNIDAD_PATH.')


def parse_periodo(s):
    s = str(s).strip()
    m = re.fullmatch(r"(\d{1,2})[/-](\d{4})", s)
    if m and 1 <= int(m.group(1)) <= 12:
        return f"{m.group(2)}{int(m.group(1)):02d}"
    if re.fullmatch(r"\d{4}(0[1-9]|1[0-2])", s):
        return s
    fallar(f'Periodo "{s}" invalido: usar MM/AAAA o AAAAMM.')


def carpeta_iva(cuit, aaaamm):
    """(carpeta de la sociedad, carpeta IVA del periodo). No crea nada."""
    c = re.sub(r"\D", "", str(cuit))
    base = raiz_unidad()
    candidatas = [d for d in sorted(os.listdir(base))
                  if os.path.isdir(os.path.join(base, d)) and c in re.sub(r"\D", "", d)]
    if not candidatas:
        fallar(f"No hay ninguna carpeta en la unidad compartida cuyo nombre contenga el CUIT {c}.")
    if len(candidatas) > 1:
        fallar(f"El CUIT {c} matchea {len(candidatas)} carpetas: {' | '.join(candidatas)}.")
    soc = os.path.join(base, candidatas[0])
    mensuales = os.path.join(soc, "01-Impuestos Mensuales")
    suelto = os.path.join(mensuales, aaaamm)
    anio = os.path.join(mensuales, aaaamm[:4])
    mes = os.path.join(anio, aaaamm) if (not os.path.exists(suelto) and os.path.isdir(anio)) else suelto
    for variante in (os.path.join(mes, f"{aaaamm}-IVA"), os.path.join(mensuales, f"{aaaamm}-IVA")):
        if os.path.isdir(variante):
            return soc, variante
    destino = os.path.join(mes, "IVA")
    if not os.path.isdir(destino):
        fallar(f"No existe la carpeta IVA de {aaaamm}: {destino}")
    return soc, destino


def nombre_sociedad(carpeta):
    """'3 - Farmacia Ejemplo - 30712345678' -> 'Farmacia Ejemplo'."""
    n = re.sub(r"^\s*\d+\s*-\s*", "", os.path.basename(carpeta))
    return re.sub(r"\s*-\s*\d{11}\s*$", "", n).strip()


def leer_libro(ruta, intentos=4):
    """load_workbook con reintentos: Drive a veces corta la lectura de archivos
    que estan solo online (llega como BadZipFile / timeout)."""
    for i in range(intentos):
        try:
            return openpyxl.load_workbook(ruta, data_only=True)
        except (OSError, zipfile.BadZipFile):
            if i == intentos - 1:
                raise
            time.sleep(10)


def pesos(x):
    s = f"{x:,.2f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def lista_diferencias(diffs):
    return [
        {
            "fecha": d["fecha"].isoformat(), "concepto": d["concepto"], "tipo": d["tipo"],
            "patron": d["patron"],
            "dif_exento": d["d_exen"], "dif_gravado": d["d_grav"],
            "dif_iva": d["d_iva"], "dif_total": d["d_total"],
            "compensa_con": d.get("compensa_con"),
        }
        for d in diffs
    ]


def coinciden(referencia, calculado):
    return all(abs(referencia[k] - calculado[k]) <= max(5.0, abs(referencia[k]) * 0.001)
               for k in referencia)


def modo_papel(entrada, salida, tolerancia):
    """Modo original: un papel de trabajo con las hojas Ventas Zetti y Ventas Portal IVA."""
    shutil.copy(entrada, salida)
    wb_val = openpyxl.load_workbook(entrada, data_only=True)
    wb_out = openpyxl.load_workbook(salida)

    zetti_name = find_sheet(wb_val, ["zetti"])
    portal_name = find_sheet(wb_val, ["portal", "iva"]) or find_sheet(wb_val, ["portal"])
    if zetti_name is None or portal_name is None:
        print(json.dumps({
            "error": "No encontre las hojas Ventas Zetti / Ventas Portal IVA",
            "hojas_disponibles": wb_val.sheetnames,
        }, ensure_ascii=False, indent=2))
        sys.exit(2)

    zetti_rows, totales_declarados, _ = parse_zetti(wb_val[zetti_name])
    portal_rows, _ = parse_portal(wb_val[portal_name])

    diffs, z_puntos, totales_diff = cruzar(zetti_rows, portal_rows, tolerancia=tolerancia)

    # sociedad / periodo para el titulo: se intentan inferir, si no se
    # completan despues a mano (ver SKILL.md)
    fechas = [r["fecha"] for r in zetti_rows]
    periodo = f"{min(fechas).strftime('%m/%Y')}" if fechas else "periodo"
    sociedad = "SOCIEDAD"

    escribir_hoja_cruze(wb_out, diffs, sociedad, periodo)
    wb_out.save(salida)

    # chequeo de sanidad contra el total que el propio archivo ya trae
    chequeo = None
    if totales_declarados and "zetti" in totales_declarados and "portal" in totales_declarados:
        tz, tp = totales_declarados["zetti"], totales_declarados["portal"]
        declarado = {
            "exen": r2((tz["exen"] or 0) - (tp["exen"] or 0)),
            "grav": r2((tz["grav"] or 0) - (tp["grav"] or 0)),
            "iva": r2((tz["iva"] or 0) - (tp["iva"] or 0)),
            "total": r2((tz["total"] or 0) - (tp["total"] or 0)),
        }
        chequeo = {
            "declarado_en_archivo": declarado,
            "calculado_por_cruze": totales_diff,
            "coincide": coinciden(declarado, totales_diff),
        }

    resumen = {
        "modo": "papel_de_trabajo",
        "hoja_zetti": zetti_name, "hoja_portal": portal_name,
        "filas_zetti": len(zetti_rows), "filas_portal": len(portal_rows),
        "puntos_venta_tickets_z": z_puntos,
        "cantidad_filas_diferencia": len(diffs),
        "totales_diferencia_cruze": totales_diff,
        "chequeo_vs_total_declarado_en_archivo": chequeo,
        "diferencias": lista_diferencias(diffs),
        "archivo_salida": salida,
    }
    print(json.dumps(resumen, ensure_ascii=False, indent=2, default=str))


def modo_dos_archivos(a):
    """Modos 1 (--cuit + --periodo) y 2 (--zetti + --portal): un archivo por reporte."""
    sociedad = None
    cuit = re.sub(r"\D", "", a.cuit) if a.cuit else None
    aaaamm = parse_periodo(a.periodo) if a.periodo else None
    if cuit:
        if not aaaamm:
            fallar("Con --cuit hace falta --periodo=MM/AAAA.")
        soc_dir, iva = carpeta_iva(cuit, aaaamm)
        sociedad = nombre_sociedad(soc_dir)
        zetti_path = a.zetti or os.path.join(iva, f"{aaaamm} - ZETTI - VENTAS.xlsx")
        portal_path = a.portal or os.path.join(iva, f"{aaaamm} - PORTAL IVA - VENTAS.xlsx")
        salida = a.salida or os.path.join(iva, f"{aaaamm} - CRUCE ZETTI VS PORTAL IVA.xlsx")
    else:
        if not (a.zetti and a.portal):
            fallar("Hacen falta --zetti y --portal (o --cuit y --periodo).")
        zetti_path, portal_path, salida = a.zetti, a.portal, a.salida

    faltan = [p for p in (zetti_path, portal_path) if not os.path.isfile(p)]
    if faltan:
        mm_aaaa = f"{aaaamm[4:]}/{aaaamm[:4]}" if aaaamm else "MM/AAAA"
        quien = f"--cuit={cuit}" if cuit else '--nombre="<sociedad>"'
        fallar("Falta el archivo de entrada.", faltan=faltan, como_bajarlo={
            "zetti": f"node scripts/zetti-iva-ventas.js {quien} --periodo={mm_aaaa}",
            "portal": f"node scripts/portal-iva-descarga.js {quien} --periodo={mm_aaaa} --libros=ventas",
        })

    wb_z, wb_p = leer_libro(zetti_path), leer_libro(portal_path)
    zetti_name = find_sheet(wb_z, ["zetti"])
    portal_name = find_sheet(wb_p, ["portal", "iva"]) or find_sheet(wb_p, ["portal"])
    if zetti_name is None or portal_name is None:
        fallar("No encontre las hojas Ventas Zetti / Ventas Portal IVA",
               hojas_zetti=wb_z.sheetnames, hojas_portal=wb_p.sheetnames)

    zetti_rows, _, zinfo = parse_zetti(wb_z[zetti_name])
    portal_rows, pinfo = parse_portal(wb_p[portal_name])
    if cuit and zinfo["cuit"] and zinfo["cuit"] != cuit:
        fallar(f"El archivo de Zetti es del CUIT {zinfo['cuit']}, no del {cuit}.", archivo=zetti_path)

    fechas = [r["fecha"] for r in zetti_rows + portal_rows]
    if not aaaamm:
        aaaamm = min(fechas).strftime("%Y%m") if fechas else None
    fuera = {
        "zetti": sum(1 for r in zetti_rows if r["fecha"].strftime("%Y%m") != aaaamm),
        "portal": sum(1 for r in portal_rows if r["fecha"].strftime("%Y%m") != aaaamm),
    }
    if salida is None:
        salida = os.path.join(os.path.dirname(os.path.abspath(zetti_path)),
                              f"{aaaamm} - CRUCE ZETTI VS PORTAL IVA.xlsx")
    if os.path.exists(salida) and a.si_existe != "reemplazar":
        fallar("Ya existe el cruce: no lo piso (puede tener Observaciones escritas a mano). "
               "Para rehacerlo: --si-existe=reemplazar.", archivo=salida)

    separadas = []
    if not a.incluir_nc_controlador:
        zetti_rows, separadas = separar_nc_controlador(zetti_rows, portal_rows)
    diffs, z_puntos, totales_diff = cruzar(zetti_rows, portal_rows, tolerancia=a.tolerancia)

    tz, tp = zinfo["suma_archivo"], pinfo["suma_archivo"]
    ts = {k: r2(sum(r[k] for r in separadas)) for k in tz}
    dif_archivos = {k: r2(tz[k] - tp[k]) for k in tz}
    a_explicar = {k: r2(tz[k] - ts[k] - tp[k]) for k in tz}
    chequeo = {
        "total_zetti": tz, "total_portal": tp,
        "diferencia_entre_archivos": dif_archivos,
        "nc_controlador_fuera_del_cruce": ts,
        "diferencia_a_explicar": a_explicar,
        "calculado_por_cruze": totales_diff,
        "coincide": coinciden(a_explicar, totales_diff),
        "filas_descartadas": {"zetti": zinfo["filas_descartadas"],
                              "portal": pinfo["filas_descartadas"]},
    }
    nc_fuera = None
    if separadas:
        por_pv = defaultdict(lambda: [0, 0.0])
        for r in separadas:
            por_pv[r["punto"]][0] += 1
            por_pv[r["punto"]][1] += r["total"]
        nc_fuera = [{"punto": pv, "cantidad": n, "total": r2(t)} for pv, (n, t) in sorted(por_pv.items())]

    razon = zinfo["razon_social"]
    titulo_soc = f"{sociedad} ({razon})" if sociedad and razon else (sociedad or razon or "SOCIEDAD")
    periodo_lindo = f"{MESES[int(aaaamm[4:]) - 1]} {aaaamm[:4]}" if aaaamm else "periodo"
    nota = (f"Total Zetti $ {pesos(tz['total'])} · Total Portal IVA $ {pesos(tp['total'])} · "
            f"Diferencia $ {pesos(dif_archivos['total'])}")
    if separadas:
        pvs = ", ".join(str(x["punto"]) for x in nc_fuera)
        nota += (f" · Fuera del cruce: {len(separadas)} NC B a Cons. Final del PV {pvs} "
                 f"($ {pesos(ts['total'])}), ver hoja \"{HOJA_NC_SEPARADAS}\"")
    nota += f" · Fuentes: {os.path.basename(zetti_path)} y {os.path.basename(portal_path)}"

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    escribir_hoja_cruze(wb_out, diffs, titulo_soc, periodo_lindo, nota=nota)
    if separadas:
        escribir_hoja_nc_separadas(wb_out, separadas)
    wb_out.save(salida)

    resumen = {
        "modo": "carpeta_iva" if cuit else "dos_archivos",
        "sociedad": sociedad, "razon_social": razon, "cuit": cuit or zinfo["cuit"],
        "periodo": aaaamm,
        "archivo_zetti": zetti_path, "archivo_portal": portal_path,
        "hoja_zetti": zetti_name, "hoja_portal": portal_name,
        "filas_zetti": len(zetti_rows), "filas_portal": len(portal_rows),
        "filas_fuera_del_periodo": fuera,
        "nc_controlador_fuera_del_cruce": nc_fuera,
        "puntos_venta_tickets_z": z_puntos,
        "cantidad_filas_diferencia": len(diffs),
        "totales_diferencia_cruze": totales_diff,
        "chequeo_vs_totales_de_los_archivos": chequeo,
        "diferencias": lista_diferencias(diffs),
        "archivo_salida": salida,
    }
    print(json.dumps(resumen, ensure_ascii=False, indent=2, default=str))


def main():
    ap = argparse.ArgumentParser(
        description="Cruce Ventas Zetti vs Ventas Portal IVA (ver docstring del script).")
    ap.add_argument("archivos", nargs="*", help="modo papel de trabajo: <entrada.xlsx> <salida.xlsx>")
    ap.add_argument("--cuit")
    ap.add_argument("--periodo", help="MM/AAAA o AAAAMM")
    ap.add_argument("--zetti", help="AAAAMM - ZETTI - VENTAS.xlsx")
    ap.add_argument("--portal", help="AAAAMM - PORTAL IVA - VENTAS.xlsx")
    ap.add_argument("--salida")
    ap.add_argument("--tolerancia", type=float, default=0.5)
    ap.add_argument("--si-existe", dest="si_existe", choices=["abortar", "reemplazar"], default="abortar")
    ap.add_argument("--incluir-nc-controlador", dest="incluir_nc_controlador", action="store_true",
                    help="no separar las NC B a Cons. Final de PV sin comprobantes en Portal IVA")
    a = ap.parse_args()

    if a.cuit or a.zetti or a.portal:
        modo_dos_archivos(a)
    elif len(a.archivos) == 2:
        modo_papel(a.archivos[0], a.archivos[1], a.tolerancia)
    else:
        ap.print_usage(sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
