"""Lectura de las fuentes acordadas el 12/11/2026 para el Balance (cualquier sociedad: llamar a configurar(cuit) antes).

Todo se lee de la unidad compartida, sin escribir nada.
- Gastos: papel de trabajo MENSUAL (hoja 'compras portal iva', columna PROVEEDORES) cuando el mes
  está clasificado; si no, las compras imputadas con el Maestro (hoja del EERR 2026).
- Matrícula Colegio: Zetti T&S Gastos Varios (AAAAMM - ZETTI - GASTOS VARIOS.xlsx, hoja Resumen).
- FCI: valorización y rendimiento de 0 - ESTUDIO/DRIVE - FCI/DRIVE - FCI 2026.
"""
import collections
import glob
import os
import unicodedata

import openpyxl

import re as _re
import sys as _sys


def _env():
    """Variables de entorno y de ~/.fisco-ar/.env (o FISCO_ENV_FILE), como el resto de los scripts del estudio."""
    f = os.environ.get("FISCO_ENV_FILE") or os.path.join(os.path.expanduser("~/.fisco-ar"), ".env")
    out = {}
    if os.path.exists(f):
        for l in open(f, encoding="utf-8", errors="replace"):
            m = _re.match(r"^\s*(?:export\s+)?([A-Za-z_]\w*)\s*=\s*(.*?)\s*$", l)
            if m and not l.lstrip().startswith("#"):
                out[m[1]] = m[2].strip('"').strip("'")
    return lambda k: os.environ.get(k) or out.get(k) or ""


var = _env()
NOMBRE_UNIDAD = "BVA - Sociedades Farmaceuticas"


def raiz_unidad():
    r = var("BVA_UNIDAD_PATH").strip()
    if r:
        if not os.path.isdir(r):
            _sys.exit(f'BVA_UNIDAD_PATH apunta a "{r}", que no existe. Verificá que Google Drive esté abierto.')
        return r
    if os.name == "nt" and os.path.isdir(f"I:\\Unidades compartidas\\{NOMBRE_UNIDAD}"):
        return f"I:\\Unidades compartidas\\{NOMBRE_UNIDAD}"
    c = glob.glob(os.path.expanduser(f"~/Library/CloudStorage/GoogleDrive-*/Unidades compartidas/{NOMBRE_UNIDAD}"))
    if len(c) == 1:
        return c[0]
    _sys.exit("No encuentro la unidad compartida: definí BVA_UNIDAD_PATH en ~/.fisco-ar/.env.")


UNIDAD = raiz_unidad()
ESTUDIO = os.path.join(UNIDAD, "0 - ESTUDIO - 00000000000")
EERR_2026 = os.path.join(ESTUDIO, "EERR 2026")
# Configs por sociedad (datos del balance anterior): en la unidad, no en el plugin (el repo es público)
CONFIG_DIR = var("BALANCE_CONFIG_DIR") or os.path.join(EERR_2026, "Config Balance")
SOC_DIR = IMP = SUELDOS_DIR = None
CUIT = CARPETA = None

FCI_DIR = os.path.join(UNIDAD, "0 - ESTUDIO - 00000000000", "DRIVE - FCI", "DRIVE - FCI 2026")
CUIT_FMT = None
MESES_AAAAMM = ["202512"] + [f"2026{m:02d}" for m in range(1, 12)]


def configurar(cuit, meses=None):
    """Apunta el módulo a la carpeta de la sociedad (por CUIT) y a sus meses de ejercicio."""
    global SOC_DIR, IMP, SUELDOS_DIR, CUIT, CUIT_FMT, CARPETA, MESES_AAAAMM
    import re
    CUIT = re.sub(r"\D", "", cuit)
    CUIT_FMT = f"{CUIT[:2]}-{CUIT[2:10]}-{CUIT[10]}"
    CARPETA = next(d for d in os.listdir(UNIDAD) if d.endswith(" - " + CUIT))
    SOC_DIR = os.path.join(UNIDAD, CARPETA)
    IMP = os.path.join(SOC_DIR, "01-Impuestos Mensuales")
    SUELDOS_DIR = os.path.join(SOC_DIR, "02-Sueldos y F.931")
    if meses:
        MESES_AAAAMM = list(meses)

# clase del papel mensual → concepto del EERR (el mismo que usa el Maestro en 'Compras imputadas')
CLASE_A_EERR = {
    "CMV": "CMV", "MERCADERIA": "CMV",
    "ALQUILERES": "ALQUILERES", "HONORARIOS POR SERVICIOS": "HONORARIOS POR SERVICIOS",
    "SEGUROS": "SEGUROS", "GASTOS VARIOS": "GASTOS VARIOS", "GASTOS BANCARIOS": "Gastos Bancarios",
    "LUZ": "Edesur", "AGUA": "Agua", "GAS": "Metrogas",
    "INTERNET / TELEFONIA": "INTERNET/TELEFONIA", "GASTOS DE SISTEMAS": "GASTOS DE SISTEMAS",
    "SEGURIDAD": "SEGURIDAD", "RESIDUOS PATOGENICOS": "RESIDUOS PATOGENICOS",
    "GASTOS DE TARJETAS": "GASTOS DE TARJETAS", "IMPRENTA Y LIBRERIA": "IMPRENTA Y PAPELERIA",
    "GASTOS DE MANTENIMIENTO": "GASTOS DE MANTENIMIENTO", "MANTEN. RODADOS": "GASTOS DE MANTENIMIENTO",
    "GASTOS DE TRANSPORTE": "GASTOS DE TRANSPORTE", "INTERESES PAGADOS": "Intereses de Prestamos",
    "GASTOS FINANCIEROS": "SERVICIOS FINANCIEROS", "SERVICIOS DE EMERGENCIAS Y TRASLADOS": "SERVICIOS MEDICOS",
    "MATRICULA COLEGIO FARMACEUTICOS": "CAMARA",
}
# El Maestro imputa la cuota de la Cámara de Farmacias a "Matrícula Colegio": acá va a su propia línea
RUBRO_MAESTRO = {"Matricula Colegio Farmaceuticos": "CAMARA"}


def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _num(v):
    return v if isinstance(v, (int, float)) else 0


def _mes_dir(aaaamm):
    for d in (os.path.join(IMP, aaaamm), os.path.join(IMP, aaaamm[:4], aaaamm)):
        if os.path.isdir(d):
            return d
    return None


def _papeles_mes(aaaamm):
    """Papeles de trabajo mensuales del mes, en orden de preferencia: carpeta del mes y su subcarpeta IVA;
    nombres "…papel…" o "Pdetrabajo-AAAAMM…" (nunca "NO VA" ni el AUTOMATICO). Primero la versión más alta
    ("(v2)") y después el más nuevo."""
    d = _mes_dir(aaaamm)
    if not d:
        return []
    cands = []
    for base in (d, os.path.join(d, "IVA")):
        if not os.path.isdir(base):
            continue
        for f in os.listdir(base):
            fl = f.lower()
            if (fl.endswith(".xlsx") and not f.startswith("~$") and "no va" not in fl and "automatico" not in fl
                    and ("papel" in fl or fl.startswith("pdetrabajo"))):
                cands.append(os.path.join(base, f))
    def clave(p):
        m = _re.search(r"\(v(\d+)\)", os.path.basename(p).lower())
        return (int(m.group(1)) if m else 1, os.path.getmtime(p))
    return sorted(cands, key=clave, reverse=True)


def _papel_mes(aaaamm):
    ps = _papeles_mes(aaaamm)
    return ps[0] if ps else None


def leer_papel_compras(path):
    """Comprobantes clasificados del papel mensual: [(concepto EERR, costo, proveedor, clase)].
    None si el papel no tiene la columna PROVEEDORES. Costo = total − IVA − percepciones/impuestos."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    hojas = [n for n in wb.sheetnames if "compra" in n.lower() and "portal" in n.lower()]
    if not hojas:
        return None, "sin hoja de compras"
    hoja = next((n for n in hojas if n.strip().lower() == "compras portal iva"), hojas[0])
    rows = list(wb[hoja].iter_rows(values_only=True))
    # encabezados en la fila que tiene "Importe Total" (algunos papeles tienen un título arriba)
    hi = next((i for i, r in enumerate(rows[:6]) if any(str(v).strip().upper() == "IMPORTE TOTAL" for v in r if v)), None)
    if hi is None:
        return None, f"hoja '{hoja}' vacía o sin encabezados del Portal IVA"
    h = [str(v).strip().upper() if v else "" for v in rows[hi]]
    if "PROVEEDORES" not in h and "RUBRO EERR" in h:
        # formato "Pdetrabajo": cada comprobante trae Rubro EERR (Maestro) e Importe imputado
        ir, it, iva = h.index("RUBRO EERR"), h.index("IMPORTE TOTAL"), h.index("TOTAL IVA")
        ii = h.index("IMPORTE IMPUTADO") if "IMPORTE IMPUTADO" in h else None
        perc = [j for j, v in enumerate(h) if v.startswith(("IMPORTE DE PER", "IMPORTE DE IMP", "IMPORTE OTROS TRIB"))]
        out, avisos = [], []
        for r in rows[hi + 1:]:
            if len(r) <= ir or not hasattr(r[0], "year") or not r[6]:
                continue  # títulos, cuadros y totales
            prov, rubro = str(r[6]).strip(), str(r[ir] or "").strip()
            costo = _num(r[ii]) if ii is not None and r[ii] not in (None, "") else _num(r[it]) - _num(r[iva]) - sum(_num(r[j]) for j in perc)
            if "CAMARA DE FARMACIAS" in _sin_tildes(prov.upper()):
                concepto = "CAMARA"
            elif rubro:
                concepto = RUBRO_MAESTRO.get(rubro, rubro)
            else:
                concepto = "GASTOS VARIOS"
                avisos.append(f"{prov}: sin Rubro EERR → Gastos varios")
            out.append((concepto, costo, prov, rubro))
        return out, "; ".join(sorted(set(avisos)))
    if "PROVEEDORES" not in h:
        return None, f"hoja '{hoja}' sin clasificación por cuenta (solo marca servicios/mercadería)"
    pc, it, iva = h.index("PROVEEDORES"), h.index("IMPORTE TOTAL"), h.index("TOTAL IVA")
    perc = [j for j, v in enumerate(h) if v.startswith(("IMPORTE DE PER", "IMPORTE DE IMP", "IMPORTE OTROS TRIB"))]
    out, avisos = [], []
    for r in rows[hi + 1:]:
        if len(r) <= pc or not isinstance(r[1], (int, float)) or not r[6]:
            continue  # filas de totales / resumen
        clase = _sin_tildes(str(r[pc] or "").strip().upper())
        prov = str(r[6]).strip()
        costo = _num(r[it]) - _num(r[iva]) - sum(_num(r[j]) for j in perc)
        if "CAMARA DE FARMACIAS" in _sin_tildes(prov.upper()):
            concepto = "CAMARA"
        elif clase in CLASE_A_EERR:
            concepto = CLASE_A_EERR[clase]
        else:
            concepto = "GASTOS VARIOS"
            avisos.append(f"{prov}: clase '{clase}' sin mapeo → Gastos varios")
        out.append((concepto, costo, prov, clase))
    return out, "; ".join(sorted(set(avisos)))


def compras_por_concepto(eerr_path):
    """{concepto EERR: [12 meses]}, origen por mes y controles."""
    wb = openpyxl.load_workbook(eerr_path, data_only=True)
    ws = wb["Compras imputadas"]
    hdr = [c.value for c in ws[1]]
    col_mes = {}
    for j, v in enumerate(hdr):
        if isinstance(v, str) and "/" in v and len(v) == 7:
            mm, aaaa = v.split("/")
            col_mes[aaaa + mm] = j
    maestro = collections.defaultdict(lambda: [0.0] * 12)
    maestro_prov = collections.defaultdict(lambda: [0.0] * 12)
    excluidos = collections.defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        rubro = r[3]
        if not rubro:
            continue
        concepto = RUBRO_MAESTRO.get(rubro, rubro)
        for i, m in enumerate(MESES_AAAAMM):
            v = r[col_mes[m]] if m in col_mes else None
            if isinstance(v, (int, float)):
                maestro[concepto][i] += v
                maestro_prov[(concepto, str(r[1]))][i] += v

    papel = {}
    origen = {}
    for i, m in enumerate(MESES_AAAAMM):
        ps = _papeles_mes(m)
        if not ps:
            origen[m] = ("Maestro", "no hay papel de trabajo del mes en la carpeta")
            continue
        datos, nota, p, notas = None, "", None, []
        for cand in ps:  # el primero que tenga la hoja de compras clasificada
            d_, n_ = leer_papel_compras(cand)
            if d_ is not None:
                datos, nota, p = d_, n_, cand
                break
            notas.append(f"{os.path.basename(cand)}: {n_}")
        if datos is None:
            origen[m] = ("Maestro", "; ".join(notas))
            continue
        agg = collections.defaultdict(float)
        for concepto, costo, *_ in datos:
            agg[concepto] += costo
        papel[i] = (agg, datos)
        rel = os.path.relpath(p, _mes_dir(m))
        origen[m] = ("Papel", f"{rel}, hoja de compras clasificada" + (f" — {nota}" if nota else ""))

    conceptos = set(maestro) | {c for agg, _ in papel.values() for c in agg}
    valores = {}
    for c in conceptos:
        fila = []
        for i in range(12):
            if i in papel:
                fila.append(papel[i][0].get(c, 0.0))
            else:
                fila.append(maestro[c][i] if maestro.get(c) else 0.0)
        valores[c] = fila
    comparacion = []  # meses con papel: papel vs Maestro por concepto
    for i, (agg, _) in sorted(papel.items()):
        for c in sorted(set(agg) | {k for k in maestro if maestro[k][i]}):
            comparacion.append((MESES_AAAAMM[i], c, agg.get(c, 0.0), maestro[c][i] if c in maestro else 0.0))
    return valores, origen, comparacion


def ncr_papel(aaaamm):
    p = _papel_mes(aaaamm)
    if not p:
        return None
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    if "NCR" not in wb.sheetnames:
        return None
    ult = None
    for r in wb["NCR"].iter_rows(values_only=True):
        nums = [v for v in r if isinstance(v, (int, float))]
        if nums:
            ult = nums[-1]
    return ult


def matricula_tys():
    """Matrícula Colegio por mes desde el resumen de Gastos Varios de Zetti (T&S)."""
    out, notas = [None] * 12, []
    for i, m in enumerate(MESES_AAAAMM):
        base = _mes_dir(m)
        f = os.path.join(base, f"{m} - ZETTI - GASTOS VARIOS.xlsx") if base else ""
        if not os.path.exists(f):
            continue
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        tot = 0.0
        for r in wb["Resumen"].iter_rows(values_only=True):
            if len(r) > 3 and isinstance(r[1], str) and r[1].lower().startswith("matricula") and isinstance(r[3], (int, float)):
                tot += r[3]
                notas.append(f"{m}: {r[0]} {r[3]:,.2f}")
        out[i] = tot
    return out, notas


def sindicatos_tys():
    """Último mes con pagos a FATFA/OSPF/ADEF en T&S (referencia de deuda sindical mensual)."""
    for m in reversed(MESES_AAAAMM):
        base = _mes_dir(m)
        f = os.path.join(base, f"{m} - ZETTI - GASTOS VARIOS.xlsx") if base else ""
        if not f or not os.path.exists(f):
            continue
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        tot, det = 0.0, []
        for r in wb["Resumen"].iter_rows(values_only=True):
            if r and isinstance(r[0], str) and any(k in r[0].upper() for k in ("FATFA", "OSPF", "ADEF")) and isinstance(r[3], (int, float)):
                tot += r[3]
                det.append(r[0].strip())
        if tot:
            return tot, m, det
    return None, None, []


def fci():
    """Saldo (valorización) y rendimiento mensual de los FCI de la sociedad por banco."""
    nombres_mes = {"ENERO": "01", "FEBRERO": "02", "MARZO": "03", "ABRIL": "04", "MAYO": "05", "JUNIO": "06",
                   "JULIO": "07", "AGOSTO": "08", "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12"}
    saldos, rend = {}, collections.defaultdict(float)
    archivos = []
    for f in sorted(glob.glob(os.path.join(FCI_DIR, "FONDOS *.xlsx"))):
        partes = os.path.basename(f).upper().split()
        banco = partes[1]
        mes = next((nombres_mes[p] for p in partes if p in nombres_mes), None)
        if not mes:
            continue
        aaaamm = "2026" + mes
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                if not any(isinstance(v, str) and CUIT_FMT in v for v in r):
                    continue
                ult = [v for v in r if isinstance(v, (int, float))][-1]
                if "aloriz" in ws.title:
                    if banco not in saldos or saldos[banco][1] < aaaamm:
                        saldos[banco] = (ult, aaaamm, os.path.basename(f))
                elif "onsolid" in ws.title:
                    rend[aaaamm] += ult
                    archivos.append(os.path.basename(f))
    rend_meses = [rend.get(m) for m in MESES_AAAAMM]
    return saldos, rend_meses, sorted(set(archivos))


if __name__ == "__main__":
    import sys
    v, o, comp = compras_por_concepto(sys.argv[1])
    for m, (src, nota) in o.items():
        print(m, src, nota)
    print({k: round(sum(x)) for k, x in v.items()})
    print(matricula_tys())
    print(sindicatos_tys())
    print(fci())
    print(ncr_papel("202608"), ncr_papel("202609"))


PLANES = os.path.join(UNIDAD, "0 - ESTUDIO - 00000000000", "INFORME_PLANES_FACILIDADES_20260925-1_2.xlsx")


def planes():
    """Planes de facilidades ARCA: capital pendiente (sin intereses futuros) e intereses pagados por mes."""
    wb = openpyxl.load_workbook(PLANES, read_only=True, data_only=True)
    pend, det, planes_ids = 0.0, [], set()
    for r in wb["PLANES VIGENTES"].iter_rows(values_only=True):
        if r and str(r[1]) == CUIT:
            pend += r[13] or 0
            planes_ids.add(r[2])
            det.append(f"{r[2]} ({str(r[4])[:22]}…): {(r[13] or 0):,.2f}")
    inter = collections.defaultdict(float)
    for r in wb["DETALLE CUOTAS"].iter_rows(values_only=True):
        if r and r[1] in planes_ids and r[6] == "Pagada" and hasattr(r[7], "strftime"):
            inter[r[7].strftime("%Y%m")] += (r[4] or 0) - (r[3] or 0)
    return pend, det, [inter.get(m) for m in MESES_AAAAMM]




def f931(aaaamm):
    """Datos del F.931 archivado (sección VIII) o None (lector leer_f931.py, en esta misma carpeta)."""
    _sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from leer_f931 import leer
    d = os.path.join(SUELDOS_DIR, aaaamm[:4])
    if not os.path.isdir(d):
        return None
    for f in sorted(os.listdir(d)):
        if f.startswith(aaaamm) and "931" in f and f.lower().endswith(".pdf"):
            x = leer(os.path.join(d, f))
            x["archivo"] = f
            # mismo criterio que la hoja Sueldos del EERR: remuneración + contribuciones + ART + seguro de vida
            x["costo_eerr"] = (x["rem_art"] or 0) + (x["c351"] or 0) + (x["c352"] or 0) + (x["c312"] or 0) + (x["c028"] or 0)
            x["a_pagar"] = sum((x[k] or 0) for k in ("c301", "c302", "c351", "c352", "c312"))
            return x
    return None


def papel_saldos():
    """Saldos de IVA, IIBB y tasas del papel de liquidaciones del equipo, del último mes con crédito fiscal cargado."""
    cands = []
    for root in (IMP, os.path.join(IMP, "2026")):
        if os.path.isdir(root):
            cands += [os.path.join(root, f) for f in os.listdir(root)
                      if "liquidaciones" in f.lower() and "automatico" not in f.lower() and f.endswith(".xlsx") and not f.startswith("~$")]
    if not cands:
        return None
    arch = sorted(cands)[-1]
    wb = openpyxl.load_workbook(arch, read_only=True, data_only=True)
    for aaaamm in reversed(MESES_AAAAMM):
        hoja = f"{aaaamm[4:]}-{aaaamm[:4]}"
        if hoja not in wb.sheetnames:
            continue
        filas, bloque = {}, ""
        for i, r in enumerate(wb[hoja].iter_rows(values_only=True), 1):
            if not r or len(r) < 3 or not isinstance(r[1], str):
                continue
            lab = _sin_tildes(r[1].strip().upper())
            if "SEGURIDAD E HIGIENE" in lab:
                bloque = "SH"
            elif "PUBLICIDAD Y PROPAGANDA" in lab:
                bloque = "PYP"
            v = r[2] if isinstance(r[2], (int, float)) else None
            key = (bloque + "|" + lab) if lab.startswith("IMPORTE A PAGAR") else lab
            filas.setdefault(key, (v, i))
        if not filas.get("TOTAL CREDITO FISCAL", (None, 0))[0]:
            continue
        g = lambda k: next((filas[x] for x in filas if x.startswith(k)), (None, None))
        return dict(mes=aaaamm, hoja=hoja, archivo=os.path.basename(arch),
                    saldo_iva=g("SALDO DE IVA"), tecnico=g("SALDO TECNICO"),
                    libre=g("SALDO DE LIBRE DISPONIBILIDAD DEL PERIODO"), iibb=g("SALDO A FAVOR / (A PAGAR)"),
                    sh=filas.get("SH|IMPORTE A PAGAR", (None, None)), pyp=filas.get("PYP|IMPORTE A PAGAR", (None, None)))
    return None


def anticipos_ganancias():
    """Anticipos de Ganancias por VEP (sin los compensados), desde los comprobantes de 04-Balances y EECC/2026/Ganancias."""
    import re
    d = os.path.join(SOC_DIR, "04-Balances y EECC", "2026", "Ganancias")
    if not os.path.isdir(d):
        return None
    vep = comp = 0.0
    n_vep = n_comp = 0
    for f in os.listdir(d):
        fu = f.upper()
        if "ANTICIPO" not in fu or "INTERES" in fu or not fu.endswith(".PDF"):
            continue
        m = re.search(r"\$\s*([\d\.]+,\d{2})", f)
        if not m:
            continue
        v = float(m.group(1).replace(".", "").replace(",", "."))
        if "COMPENSAD" in fu:
            comp += v; n_comp += 1
        else:
            vep += v; n_vep += 1
    if not (n_vep or n_comp):
        return None
    return dict(vep=round(vep, 2), n_vep=n_vep, comp=round(comp, 2), n_comp=n_comp)


def ddjj_ganancias():
    d = os.path.join(SOC_DIR, "04-Balances y EECC", "2026", "Ganancias")
    if not os.path.isdir(d):
        return None
    return next((f for f in sorted(os.listdir(d)) if "DDJJ" in f.upper() and "GANANCIAS" in f.upper() and f.lower().endswith(".pdf")), None)


def bancos_carpeta():
    """Bancos con carpeta en 03-Bancos y Conciliaciones y los meses de extracto que tienen."""
    d = os.path.join(SOC_DIR, "03-Bancos y Conciliaciones")
    out = {}
    if os.path.isdir(d):
        for b in sorted(os.listdir(d)):
            p = os.path.join(d, b)
            if os.path.isdir(p) and not b.startswith(("_", ".")):
                out[b] = sorted({x for _, ds, _ in os.walk(p) for x in ds if x.isdigit() and len(x) == 6})
    return out


def legalizado():
    """Balance legalizado anterior subido (04-Balances y EECC/AAAA/EECC Legalizados)."""
    for y in ("2026", "2025"):
        d = os.path.join(SOC_DIR, "04-Balances y EECC", y, "EECC Legalizados")
        if os.path.isdir(d):
            fs = [f for f in os.listdir(d) if f.lower().endswith(".pdf")]
            if fs:  # primero los EECC_<cuit>_...; si no, el que diga Balance/Certificado; si no, cualquiera
                fs.sort(key=lambda f: (not f.upper().startswith("EECC"), "balance" not in f.lower() and "certificad" not in f.lower(), f))
                return os.path.join(y, "EECC Legalizados", fs[0])
    return None
