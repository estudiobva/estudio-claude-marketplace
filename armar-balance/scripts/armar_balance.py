#!/usr/bin/env python3
"""Libro de Balance (EECC) de una sociedad — prototipo (cierre del config "cierre", por defecto 30/11/2026).

Uso: armar_balance.py <EERR.xlsx> <salida.xlsx> <config.json>
La config (config/<sociedad>.json) trae los datos de la sociedad y el comparativo del balance legalizado anterior
(si ya está transcripto); sin comparativo, esa columna queda vacía y marcada como pendiente.
El EERR se lee con data_only (valores ya calculados por Excel).
"""
import sys
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

EERR, OUT, CFGP = sys.argv[1], sys.argv[2], sys.argv[3]
import os
import json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fuentes

CFG = json.load(open(CFGP, encoding="utf-8"))
# Cierre del ejercicio (config "cierre" AAAA-MM-DD; por defecto 30/11/2026) → meses, fechas y textos
_C = date.fromisoformat(CFG.get("cierre", "2026-11-30"))
_MES = [((_C.year * 12 + _C.month - 1 - k) // 12, (_C.year * 12 + _C.month - 1 - k) % 12 + 1) for k in range(11, -1, -1)]
fuentes.configurar(CFG["cuit"], meses=[f"{y}{m:02d}" for y, m in _MES])
SOC = CFG["razon_social"]
CORTO = CFG["nombre"]
CUIT = fuentes.CUIT_FMT
REG = CFG.get("registro") or "[a completar]"
CIERRE, CIERRE_ANT = _C.strftime("%d/%m/%Y"), _C.replace(year=_C.year - 1).strftime("%d/%m/%Y")
A_ACT, A_ANT = str(_C.year), str(_C.year - 1)
EJ_NRO = CFG.get("ejercicio_nro") or "[a completar]"
COMP = CFG.get("comparativo") or {}
SIN_COMP = "Pendiente: falta el balance legalizado anterior" if not COMP else ""
MESES = [f"{m:02d}-{y}" for y, m in _MES]

NUM = '#,##0.00;(#,##0.00);"-"'
AZUL = PatternFill("solid", fgColor="1F3864")
GRIS = PatternFill("solid", fgColor="D9D9D9")
CELESTE = PatternFill("solid", fgColor="DDEBF7")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
BLANCO_B = Font(bold=True, color="FFFFFF")
B = Font(bold=True)
TIT = Font(bold=True, size=13)
NOTA = Font(italic=True, size=9, color="595959")
fino = Side(style="thin", color="A6A6A6")
CAJA = Border(top=fino, bottom=fino, left=fino, right=fino)
ARRIBA = Border(top=Side(style="thin"))

wb = openpyxl.Workbook()
wb.remove(wb.active)


def hoja(nombre, titulo, anchos):
    ws = wb.create_sheet(nombre)
    ws["A1"] = f"DENOMINACIÓN DE LA SOCIEDAD: {SOC}   ·   CUIT {CUIT}"
    ws["A2"] = f"EJERCICIO FINALIZADO EL: {CIERRE}   ·   Nº de registro DPPJ: {REG}"
    ws["A3"] = titulo
    ws["A3"].font = TIT
    ws["A1"].font = ws["A2"].font = Font(size=9)
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[L(i)].width = w
    ws.sheet_view.showGridLines = False
    return ws


def encabezado(ws, fila, textos):
    for i, t in enumerate(textos, 1):
        c = ws.cell(fila, i, t)
        c.font, c.fill, c.border = BLANCO_B, AZUL, CAJA
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30


def num(c):
    c.number_format = NUM


def total(ws, fila, ncols):
    for i in range(1, ncols + 1):
        c = ws.cell(fila, i)
        c.font, c.fill, c.border = B, GRIS, CAJA


def entrada(c):
    c.fill = AMARILLO


# ---------------------------------------------------------------- datos EERR
src = openpyxl.load_workbook(EERR, data_only=True)
se = src["EERR"]
filas_eerr = {}
for r in range(8, 70):
    a = se.cell(r, 1).value
    if a:
        filas_eerr.setdefault(str(a).strip(), r)


_HOY = date.today(); _MA = _HOY.year * 12 + _HOY.month - 2  # mes anterior al actual
ULT_DATO = CFG.get("ultimo_mes_con_datos") or f"{_MA // 12}{_MA % 12 + 1:02d}"  # último mes con fuentes en la unidad
HASTA = int(CFG.get("meses_cargados", sum(1 for y, m in _MES if f"{y}{m:02d}" <= ULT_DATO)))


def meses_eerr(nombre, hasta=HASTA, quitar=()):
    r = filas_eerr[nombre]
    out = []
    for i in range(12):
        v = se.cell(r, 2 + i).value
        if i >= hasta or MESES[i] in quitar or not isinstance(v, (int, float)):
            v = None
        out.append(v)
    return out


COMPRAS_V, ORIGEN, COMPARACION = fuentes.compras_por_concepto(EERR)
MATR, MATR_NOTAS = fuentes.matricula_tys()
FCI_SALDOS, FCI_REND, FCI_ARCH = fuentes.fci()
SIND = fuentes.sindicatos_tys()
PLAN_PEND, PLAN_DET, PLAN_INT = fuentes.planes()
F931_CARGADOS = []
F931_ULT = next(((m, x) for m in reversed(fuentes.MESES_AAAAMM) for x in [fuentes.f931(m)] if x), (None, None))
NCR = {m: fuentes.ncr_papel(m) for m in fuentes.MESES_AAAAMM if ORIGEN.get(m, ("",))[0] == "Papel"}
MESES_PAPEL = [MESES[i] for i, m in enumerate(fuentes.MESES_AAAAMM) if ORIGEN[m][0] == "Papel"]


def valores_compras(e, hasta=HASTA):
    v = COMPRAS_V.get(e, [0.0] * 12)
    return [(v[i] if i < hasta else None) for i in range(12)]


# ============================================================ Resumen Resultados
PAPEL = f"Papel de liquidaciones del equipo (01-Impuestos Mensuales/202600-liquidaciones…{CORTO.upper()}.xlsx, hoja EERR), copiado del EERR 2026"
COMPRAS = ("Papel de trabajo mensual (AAAAMM- papeles de trabajo, hoja de compras, columna PROVEEDORES) en "
           + ", ".join(MESES_PAPEL) + "; los demás meses no están clasificados en el papel: libro de compras con el Maestro (ver 'Origen por mes')")
CMV_F = "Libro IVA Compras sin NCR conciliadas (EERR 2026, criterio del 05/11). El papel mensual resta las NCR: ver 'Origen por mes'"
F931 = "F.931 (02-Sueldos y F.931) → hoja 'Sueldos'; cargas sociales a revisar con SETI"
TYS = "Zetti T&S Gastos Varios (01-Impuestos Mensuales/AAAAMM/AAAAMM - ZETTI - GASTOS VARIOS.xlsx, hoja Resumen, concepto COLEGIO)"

# (concepto en el EERR, concepto acá, fuente, fase, observación, tipo)
LINEAS = [
    ("Ventas exentas y no gravadas", "Ventas exentas y no gravadas", PAPEL, 1, "", "venta"),
    ("Ventas gravadas", "Ventas gravadas", PAPEL, 1, "", "venta"),
    ("CMV", "Compras de mercadería", CMV_F, 1, "Compras del ejercicio; el CMV sale del Anexo 3 (con existencias)", "compra"),
    ("IIBB", "IIBB", PAPEL, 1, "", "gasto"),
    ("IVA Prorrateo", "IVA Prorrateo", PAPEL, 1, "", "gasto"),
    ("Sueldos y Cargas Sociales", "Sueldos y Cargas Sociales", F931, 1, "Faltan los F.931 de los últimos meses y la información de SETI", "gasto"),
    ("Matricula Colegio Farmaceuticos", "Matricula Colegio Farmaceuticos", TYS, 1, "Acuerdo 12/11: no sale del libro de compras", "matr"),
    ("CAMARA", "Cámara de Farmacias (cuota)", COMPRAS, 1, "Factura de la Cámara de Farmacias; antes el Maestro la imputaba a Matrícula", "gc"),
    ("Seguridad e Higiene", "Seguridad e Higiene", PAPEL, 1, "", "gasto"),
    ("Publicidad", "Publicidad", PAPEL + " (PyP) + compras imputadas a Publicidad", 1, "", "gasto"),
    ("Marquesinas", "Marquesinas", "Boleta de marquesina (01-Impuestos Mensuales/AAAAMM/TSH)", 2, "", "gasto"),
    ("Metrogas", "Metrogas", COMPRAS, 1, "", "gc"),
    ("Agua", "Agua", COMPRAS, 1, "", "gc"),
    ("Edesur", "Luz (Edesur)", COMPRAS, 1, "", "gc"),
    ("Gastos Bancarios", "Gastos Bancarios", "Extractos bancarios, conciliados con las facturas del banco (acuerdo 12/11)", 2, "PROVISORIO: facturas del banco en el libro de compras; antes de 07/2026 no hay facturas", "gc"),
    ("Intereses de Prestamos", "Intereses de Prestamos", "Cuadro de marcha de préstamos: solo intereses pagados, no la cuota (acuerdo 12/11)", 2, "PROVISORIO: facturas de intereses del libro de compras (Galicia)", "gc"),
    ("Comisiones de tarjetas", "Comisiones de tarjetas", "Reportes de comisiones de tarjetas (skill a desarrollar)", 2, "Pendiente: faltan las claves de las tarjetas (consulta al cliente). En T&S hay un concepto COMIS (09/2026: $5,6 M) a revisar", "gasto"),
    ("Gastos IVA Compras", "Gastos IVA Compras", COMPRAS, 1, "", "gc"),
    ("ALQUILERES", "Alquileres", COMPRAS, 1, "", "gc"),
    ("GASTOS DE COMERCIALIZACION", "Gastos de comercialización", COMPRAS, 1, "", "gc"),
    ("COMISIONES DE MERCADO LIBRE", "Comisiones de Mercado Libre", COMPRAS, 1, "", "gc"),
    ("COMISIONES DE MERCADO ENVIOS", "Comisiones de Mercado Envíos", COMPRAS, 1, "", "gc"),
    ("GASTOS DE MANTENIMIENTO", "Gastos de mantenimiento", COMPRAS, 1, "", "gc"),
    ("GASTOS DE MANTENIMIENTO BS . USO", "Gastos de mantenimiento bienes de uso", COMPRAS, 1, "", "gc"),
    ("GASTOS DE SISTEMAS", "Gastos de sistemas", COMPRAS, 1, "", "gc"),
    ("GASTOS DE TARJETAS", "Gastos de tarjetas", COMPRAS, 1, "", "gc"),
    ("GASTOS DE TRANSPORTE", "Gastos de transporte", COMPRAS, 1, "", "gc"),
    ("COMBUSTIBLES", "Combustibles", COMPRAS, 1, "", "gc"),
    ("GASTOS VARIOS", "Gastos varios", COMPRAS, 1, "", "gc"),
    ("HONORARIOS POR SERVICIOS", "Honorarios por servicios", COMPRAS, 1, "", "gc"),
    ("IMPRENTA Y PAPELERIA", "Imprenta y papelería", COMPRAS, 1, "", "gc"),
    ("GASTOS DE LIMPIEZA", "Gastos de limpieza", COMPRAS, 1, "", "gc"),
    ("INTERNET/TELEFONIA", "Internet / telefonía", COMPRAS, 1, "", "gc"),
    ("REFRIGERIO", "Refrigerio", COMPRAS, 1, "", "gc"),
    ("GASTOS DE REPRESENTACION", "Gastos de representación", COMPRAS, 1, "", "gc"),
    ("RESIDUOS PATOGENICOS", "Residuos patogénicos", COMPRAS, 1, "", "gc"),
    ("SEGURIDAD", "Seguridad", COMPRAS, 1, "", "gc"),
    ("SERVICIOS FINANCIEROS", "Servicios financieros", COMPRAS, 1, "", "gc"),
    ("SEGUROS", "Seguros", COMPRAS, 1, "", "gc"),
    ("SERVICIOS MEDICOS", "Servicios médicos", COMPRAS, 1, "", "gc"),
    ("UNIFORMES", "Uniformes", COMPRAS, 1, "", "gc"),
    ("Intereses por Facilidades de Pago", "Intereses por facilidades de pago", "Mis Facilidades ARCA: cuotas pagadas, total − capital (0 - ESTUDIO/INFORME_PLANES_FACILIDADES_20260925, hoja DETALLE CUOTAS)", 2, "PROVISORIO: informe al 25/09/2026; faltan cuotas de 10 y 11/2026 y planes de ARBA", "planes"),
    ("Indemnización por Despido", "Indemnización por despido", "Pedido al cliente (liquidación final)", 3, "", "gasto"),
    ("AMORTIZACION", "Amortización bienes de uso", "Anexo 1 (amortización del ejercicio)", 4, "Sale del Anexo 1", "amort"),
]

rr = hoja("Resumen Resultados", "RESUMEN DE RESULTADOS — mensual (hoja de trabajo)",
          [40] + [14] * 12 + [14, 16, 60, 6, 40])
rr["A4"] = ("Fuentes acordadas el 12/11: gastos del papel de trabajo mensual (donde está clasificado), Matrícula de T&S, "
            "FCI de DRIVE - FCI; ventas, IIBB, prorrateo y tasas del papel de liquidaciones; sueldos del F.931. "
            "Columna 'Ajustes' = suma de la hoja Ajustes. Toda línea debe tener Fuente.")
rr["A4"].font = NOTA
encabezado(rr, 5, ["Concepto"] + MESES + ["Ajustes", "Total ejercicio", "Fuente", "Fase", "Observaciones"])
RRF = {}
fila = 6


def rr_linea(nombre, valores, fuente, fase, obs, formula_total=None):
    global fila
    rr.cell(fila, 1, nombre)
    for i, v in enumerate(valores or [None] * 12):
        c = rr.cell(fila, 2 + i, v)
        num(c)
    c = rr.cell(fila, 14, f'=SUMIFS(Ajustes!$F:$F,Ajustes!$D:$D,"Resultados",Ajustes!$E:$E,$A{fila},Ajustes!$I:$I,"<>Descartado")')
    num(c)
    c = rr.cell(fila, 15, formula_total or f"=SUM(B{fila}:N{fila})")
    num(c)
    c.font = B
    rr.cell(fila, 16, fuente)
    rr.cell(fila, 17, fase).alignment = Alignment(horizontal="center")
    rr.cell(fila, 18, obs)
    for col in range(1, 19):
        rr.cell(fila, col).border = CAJA
    RRF[nombre] = fila
    fila += 1


def rr_titulo(t):
    global fila
    rr.cell(fila, 1, t).font = B
    for col in range(1, 19):
        rr.cell(fila, col).fill = CELESTE
    fila += 1


rr_titulo("INGRESOS")
for e, n, f, fa, o, t in LINEAS:
    if t == "venta":
        rr_linea(n, meses_eerr(e), f, fa, o)
r1, r2 = RRF["Ventas exentas y no gravadas"], RRF["Ventas gravadas"]
rr.cell(fila, 1, "Total ventas")
for col in range(2, 16):
    num(rr.cell(fila, col, f"=SUM({L(col)}{r1}:{L(col)}{r2})"))
total(rr, fila, 18)
RRF["Total ventas"] = fila
fila += 2

rr_titulo("COSTO")
for e, n, f, fa, o, t in LINEAS:
    if t == "compra":
        rr_linea(n, meses_eerr(e), f, fa, o)
fila += 1

rr_titulo("GASTOS")
g_ini = fila
for e, n, f, fa, o, t in LINEAS:
    if t == "gasto":
        vals = meses_eerr(e)
        if e == "Sueldos y Cargas Sociales":
            for i in range(HASTA):
                if vals[i] in (None, 0):
                    x = fuentes.f931(fuentes.MESES_AAAAMM[i])
                    vals[i] = x["costo_eerr"] if x else None
                    if x:
                        F931_CARGADOS.append(f"{MESES[i]} desde {x['archivo']}")
            o = ("F.931 cargado directo al Balance: " + "; ".join(F931_CARGADOS) + ". " if F931_CARGADOS else "") + f"Cargado hasta {MESES[HASTA - 1]}; falta la información de SETI"
        rr_linea(n, vals, f, fa, o)
    elif t == "gc":
        rr_linea(n, valores_compras(e), f, fa, o)
    elif t == "planes":
        rr_linea(n, PLAN_INT, f, fa, o)
    elif t == "matr":
        rr_linea(n, [(v if i < HASTA else None) for i, v in enumerate(MATR)], f, fa, o)
    elif t == "amort":
        rr_linea(n, None, f, fa, o, formula_total=None)  # total se completa abajo (Anexo 1)
        AMORT_ROW = fila - 1
g_fin = fila - 1
rr.cell(fila, 1, "Total gastos")
for col in range(2, 16):
    num(rr.cell(fila, col, f"=SUM({L(col)}{g_ini}:{L(col)}{g_fin})"))
total(rr, fila, 18)
RRF["Total gastos"] = fila
fila += 2

rr_titulo("RESULTADOS FINANCIEROS Y OTROS")
rr_linea("Resultados financieros y por tenencia (incl. RECPAM)", FCI_REND,
         "Rendimiento de los FCI: 0 - ESTUDIO/DRIVE - FCI/DRIVE - FCI 2026, hoja Consolidado ($ Rendimiento) + RECPAM (el socio a cargo)", 2,
         "Reportes de FCI de DRIVE - FCI (hasta 06/2026); RECPAM pendiente")
rr_linea("Otros ingresos", None, "Ajuste manual con respaldo", 4, "")
rr.freeze_panes = "B6"

# ============================================================ Resumen Patrimonial
from datetime import date
EXTR = f"Extracto bancario: saldo al {CIERRE} (03-Bancos y Conciliaciones/[BANCO])"
def d_(aaaamm):  # último día del mes
    import calendar
    a, m = int(aaaamm[:4]), int(aaaamm[4:])
    return date(a, m, calendar.monthrange(a, m)[1])

PS = fuentes.papel_saldos()
ANT_G = fuentes.anticipos_ganancias()
DDJJG = fuentes.ddjj_ganancias()
BANC = fuentes.bancos_carpeta()
LEG = fuentes.legalizado()
N2C = COMP.get("nota2", {})
ant = lambda code: N2C.get(code)
PAPEL_M = (f"Papel de liquidaciones ({PS['archivo']}), hoja del último mes liquidado" if PS else "Papel de liquidaciones del equipo")
pv = lambda k: (PS[k] if PS else (None, None))
CORTE_P = d_(PS["mes"]) if PS else None
iva_s, iva_f = pv("saldo_iva"); tec, tec_f = pv("tecnico"); lib, lib_f = pv("libre")
ib, ib_f = pv("iibb"); sh, sh_f = pv("sh"); pyp, pyp_f = pv("pyp")
fav = iva_s is not None and iva_s < 0
r2 = lambda v: None if v is None else round(v, 2)
D1 = r2((max(tec or 0, 0) if fav else 0) if iva_s is not None else None)
D2 = r2(((lib or 0) if fav else 0) if iva_s is not None else None)
I2 = r2((iva_s if iva_s > 0 else 0) if iva_s is not None else None)
D4 = r2((-ib if ib < 0 else 0) if ib is not None else None)
I1 = r2((ib if ib > 0 else 0) if ib is not None else None)
I4 = r2((sh or 0) + (pyp or 0)) if PS else None
hj = PS["hoja"] if PS else ""

_n = lambda t: fuentes._sin_tildes(t).lower()
# Bancos: los del balance anterior (config) + los que tienen carpeta en 03-Bancos y Conciliaciones
bancos = [list(b) for b in CFG.get("bancos", [])]
for nom, meses in BANC.items():
    if not any(_n(nom)[:5] in _n(b[0]) for b in bancos):
        bancos.append([f"Banco {nom.title()} Cta. Cte.", None, ""])
filas_a = [("a.1", "Caja en pesos", ant("a.1"), "Pedido a la farmacia (arqueo al cierre); lo cruza la auditoría", 3, None, None, "", "En el pedido al cliente (pedidos centralizados)")]
for k, (cta, v25, obs) in enumerate(bancos, 2):
    carpeta = next((nb for nb in BANC if _n(nb)[:5] in _n(cta)), None)
    donde = f"03-Bancos y Conciliaciones/{carpeta}" if carpeta else ""
    if not obs:
        obs = (f"Carpeta del banco en la unidad, sin extracto de {CIERRE[3:]}" if carpeta else "En la unidad no hay extractos de este banco: confirmar si sigue abierta")
    filas_a.append((f"a.{k}", cta, v25, EXTR, 1, None, None, donde, obs))

NOM_FCI = {"FRANCES": "FCI BBVA (Francés)", "COMAFI": "FCI Comafi", "GALICIA": "FCI Galicia", "HSBC": "FCI HSBC"}
fci25 = CFG.get("fci_2025", {})
filas_b = []
for k, banco in enumerate(sorted(set(FCI_SALDOS) | set(fci25)), 1):
    sal = FCI_SALDOS.get(banco)
    filas_b.append((f"b.{k}", NOM_FCI.get(banco, f"FCI {banco.title()}"), fci25.get(banco), "Extracto del FCI al cierre (acuerdo 12/11); lo cruza la auditoría", 1,
                    sal and r2(sal[0]), sal and d_(sal[1]), f"0 - ESTUDIO/DRIVE - FCI/DRIVE - FCI 2026/{sal[2]}, hoja Valorización Inversión" if sal else "",
                    "" if sal else "Sin reportes 2026 en DRIVE - FCI"))
if not filas_b:
    filas_b = [("b.1", "Fondos comunes de inversión", None, "Extracto del FCI al cierre (acuerdo 12/11)", 1, None, None, "", "Sin reportes 2026 en DRIVE - FCI para esta sociedad")]

G25 = CFG.get("ganancias_2025") or {}
gan_f = f"DDJJ Ganancias 2025 (04-Balances y EECC/2026/Ganancias/{DDJJG})" if DDJJG else "DDJJ Ganancias del ejercicio anterior (no está en 04-Balances y EECC/2026/Ganancias)"
ant_txt = (f"04-Balances y EECC/2026/Ganancias: {ANT_G['n_vep']} cuota(s) por VEP; {ANT_G['n_comp']} 'compensado'" if ANT_G else "")
ant_obs = (f"Se excluyen las cuotas compensadas (${ANT_G['comp']:,.2f}) para no duplicar el crédito. Revisar qué cuotas vencen después del {CIERRE[:5]}" if ANT_G else "No hay comprobantes de anticipos en 04-Balances y EECC/2026/Ganancias")
deudas_b = [list(x) for x in CFG.get("deudas_bancarias", [])]

# (código, cuenta, saldo 2025, fuente acordada, fase, saldo según fuente, corte, dónde está el dato, observación)
PAT = [
    ("a", "Caja y Bancos", filas_a),
    ("b", "Inversiones", filas_b),
    ("c", "Créditos por Ventas", [
        ("c.1", "Deudores por ventas", ant("c.1"), "Pedido a la farmacia / SETI: obras sociales y tarjetas a cobrar", 3, None, None, "", "En el pedido al cliente"),
    ]),
    ("d", "Créditos Fiscales", [
        ("d.1", "IVA saldo técnico a favor", ant("d.1"), PAPEL_M + ", fila 'Saldo Técnico'", 1, D1, CORTE_P, f"hoja {hj}, fila {tec_f}" if PS else "", ""),
        ("d.2", "IVA saldo de libre disponibilidad a favor", ant("d.2"), PAPEL_M + ", fila 'Saldo de Libre Disponibilidad del Período'", 1, D2, CORTE_P,
         f"hoja {hj}, fila {lib_f}" if PS else "", f"Saldo de IVA del mes: {iva_s:,.2f} (fila {iva_f}; negativo = a favor)" if iva_s is not None else ""),
        ("d.3", "Saldo a favor Impuesto a las Ganancias", ant("d.3"), gan_f, 1, G25.get("saldo_favor"), G25.get("fecha") and date.fromisoformat(G25["fecha"]),
         G25.get("nota", ""), "" if G25 else "Falta leer la DDJJ (PDF escaneado)"),
        ("d.4", "Saldo a favor Ingresos Brutos", ant("d.4"), PAPEL_M + ", fila 'SALDO A FAVOR / (A PAGAR)'", 1, D4, CORTE_P,
         f"hoja {hj}, fila {ib_f} (negativo = a favor)" if PS else "", ""),
        ("d.5", "Crédito por Impuesto a los Débitos y Créditos Bancarios", ant("d.5"), "Extractos bancarios: impuesto Ley 25.413 del ejercicio (100% computable si es PyME)", 2, None, None, "", ""),
        ("d.6", "Retenciones Impuesto a las Ganancias", ant("d.6"), "Mis Retenciones ARCA (fisco-ar arca-mis-retenciones)", 1, None, None, "", "Falta correr la consulta"),
        ("d.7", "Anticipos Impuesto a las Ganancias", ant("d.7"), "SCT ARCA: anticipos del ejercicio MENOS los compensados (acuerdo 12/11)", 1,
         ANT_G and ANT_G["vep"], ANT_G and date(2026, 11, 19), ant_txt, ant_obs),
    ]),
    ("e", "Bienes de Cambio", [
        ("e.1", "Mercadería de reventa", ant("e.1"), "Inventario al cierre (inventariador); costos validados con el cliente → 'Cuadro de stock' (hoja Ajustes)", 3, None, None, "", "Se completa desde el Cuadro de stock"),
    ]),
    ("f", "Créditos Fiscales no corrientes", [
        ("f.1", "Quebranto impositivo", ant("f.1"), "DDJJ Ganancias del ejercicio anterior (acumulado al inicio, usado, remanente)", 1,
         G25.get("quebranto"), G25.get("fecha_quebranto") and date.fromisoformat(G25["fecha_quebranto"]), G25.get("nota_quebranto", ""), G25.get("obs_quebranto", "")),
        ("f.2", "Impuesto diferido – ajuste por inflación impositivo", ant("f.2"), "Hoja Ganancias: sesión aparte con el socio a cargo", 4, None, None, "", ""),
    ]),
    ("g", "Deudas Comerciales", [
        ("g.1", "Proveedores", ant("g.1"), "Cuentas corrientes de proveedores: compras + pagos bancarios + NC comerciales", 2, None, None, "", "Módulo de cuentas corrientes a desarrollar"),
    ]),
    ("h", "Deudas Sociales", [
        ("h.1", "Sueldos a pagar", ant("h.1"), "SETI / planilla de sueldos del equipo (02-Sueldos y F.931/2026, fila 'Sueldos a Pagar')", 2, None, None, "", ""),
        ("h.2", "Cargas sociales a pagar (F.931 del último mes)", ant("h.2"), "F.931 de noviembre (aportes + contribuciones + ART)", 1,
         F931_ULT[1] and round(F931_ULT[1]["a_pagar"], 2), F931_ULT[0] and d_(F931_ULT[0]),
         f"02-Sueldos y F.931/2026/{F931_ULT[1]['archivo']}: 301 + 302 + 351 + 352 + 312" if F931_ULT[1] else "", f"Último F.931 en la unidad: {F931_ULT[0]}"),
        ("h.3", "Sindicatos y obras sociales a pagar", ant("h.3"), "SETI (FATFA, OSPF, ADEF del último mes)", 2, SIND[0], SIND[1] and d_(SIND[1]),
         (f"{SIND[1]} - ZETTI - GASTOS VARIOS.xlsx, hoja Resumen: " + ", ".join(SIND[2])) if SIND[1] else "", "Referencia: lo pagado en el último mes con datos en T&S"),
        ("h.4", "Planes de facilidades de pago ARCA", ant("h.4"), "Mis Facilidades ARCA: capital pendiente, sin intereses futuros", 1, PLAN_PEND if PLAN_DET else None,
         date(2026, 9, 25) if PLAN_DET else None, "0 - ESTUDIO/INFORME_PLANES_FACILIDADES_20260925-1_2.xlsx, hoja PLANES VIGENTES" if PLAN_DET else "",
         ("; ".join(PLAN_DET) + ". Revisar desafectación de pagos") if PLAN_DET else "Sin planes vigentes en el informe del 25/09"),
    ]),
    ("i", "Deudas Fiscales", [
        ("i.1", "IIBB a pagar", ant("i.1"), PAPEL_M + ", fila 'SALDO A FAVOR / (A PAGAR)'", 1, I1, CORTE_P, f"hoja {hj}, fila {ib_f}" if PS else "", ""),
        ("i.2", "IVA a pagar", ant("i.2"), PAPEL_M + ", fila 'SALDO DE IVA + A PAGAR / - A FAVOR'", 1, I2, CORTE_P, f"hoja {hj}, fila {iva_f}" if PS else "", ""),
        ("i.3", "Provisión Impuesto a las Ganancias", ant("i.3"), "Hoja Ganancias: impuesto determinado (fórmula)", 2, None, None, "", ""),
        ("i.4", "Tasas municipales a pagar (Seguridad e Higiene + PyP del último mes)", ant("i.4"), PAPEL_M + ", bloques TSH y PyP", 1, I4, CORTE_P,
         f"hoja {hj}: fila {sh_f} (S&H {sh or 0:,.2f}) + fila {pyp_f} (PyP {pyp or 0:,.2f})" if PS else "", ""),
    ]),
    ("j", "Deudas Financieras", [(f"j.{k}", cta, v25, EXTR, 1, None, None, "", obs) for k, (cta, v25, obs) in enumerate(deudas_b, 1)] + [
        (f"j.{len(deudas_b) + 1}", "Préstamos bancarios (capital, sin intereses no devengados)", ant("j.prestamos"),
         "Cuadro de marcha de préstamos (bancos-bva): valor nominal sin intereses futuros", 1, None, None, "", CFG.get("obs_prestamos", "")),
    ]),
    ("k", "Otras Deudas no corrientes", [
        ("k.1", "Impuesto diferido", ant("k.1"), "Hoja Ganancias: sesión aparte con el socio a cargo", 4, None, None, "", ""),
    ]),
]

SIN_DET = COMP.get("rubros_sin_detalle", {})
for letra, rubro, cuentas in PAT:
    if letra in SIN_DET:
        cuentas.append((f"{letra}.s", "Saldo del balance anterior sin detalle por cuenta", SIN_DET[letra],
                        "Balance legalizado anterior: no trae la Nota 2 de este rubro por cuenta", 4, 0, None, "", "Solo para el comparativo"))
# Cuentas que el balance anterior trae y el modelo no tiene (config "cuentas_extra": {letra: [[cuenta, saldo 2025, fuente]]})
for letra, rubro, cuentas in PAT:
    for k, (cta, v25, fte) in enumerate(COMP.get("cuentas_extra", {}).get(letra, []), 1):
        cuentas.append((f"{letra}.x{k}", cta, v25, fte or "A definir", 2, None, None, "", "Cuenta del balance anterior"))

rp = hoja("Resumen Patrimonial", "RESUMEN PATRIMONIAL — recopilación de saldos al cierre (hoja de trabajo)",
          [7, 24, 44, 17, 14, 18, 18, 56, 6, 12, 12, 58, 52])
rp["A4"] = ("'Saldo según fuente' = último dato disponible de la fuente acordada, con su fecha en 'Corte del dato'. "
            f"Si el corte es anterior al {CIERRE} queda Provisorio; vacío = Pendiente. 'Saldo al cierre' = fuente + ajustes.")
rp["A4"].font = NOTA
encabezado(rp, 5, ["Código", "Rubro (Nota 2)", "Cuenta", "Saldo según fuente", "Ajustes", f"Saldo al {CIERRE}",
                   f"Saldo al {CIERRE_ANT}", "Fuente acordada", "Fase", "Estado", "Corte del dato", "Dónde está el dato (archivo, hoja, fila)", "Observaciones"])
RPF = {}
fila = 6
for letra, rubro, cuentas in PAT:
    for cod, cuenta, ant, fuente, fase, valor, corte, donde, obs in cuentas:
        rp.cell(fila, 1, cod)
        rp.cell(fila, 2, f"{letra}) {rubro}")
        rp.cell(fila, 3, cuenta)
        d = rp.cell(fila, 4, valor)
        entrada(d)
        num(d)
        num(rp.cell(fila, 5, f'=SUMIFS(Ajustes!$F:$F,Ajustes!$D:$D,"Patrimonial",Ajustes!$E:$E,$A{fila},Ajustes!$I:$I,"<>Descartado")'))
        c = rp.cell(fila, 6, f"=N(D{fila})+E{fila}")
        num(c)
        c.font = B
        num(rp.cell(fila, 7, ant))
        rp.cell(fila, 8, fuente)
        rp.cell(fila, 9, fase).alignment = Alignment(horizontal="center")
        rp.cell(fila, 10, f'=IF(D{fila}="","Pendiente",IF(K{fila}="","Cargado",IF(K{fila}<DATE({_C.year},{_C.month},{_C.day}),"Provisorio","Cargado")))')
        k = rp.cell(fila, 11, corte)
        k.number_format = "dd/mm/yyyy"
        entrada(k)
        rp.cell(fila, 12, donde)
        rp.cell(fila, 13, obs)
        for col in range(1, 14):
            rp.cell(fila, col).border = CAJA
            if col in (8, 12, 13):
                rp.cell(fila, col).alignment = Alignment(wrap_text=True, vertical="top")
        RPF[cod] = fila
        fila += 1
rp.freeze_panes = "D6"
RP_FIN = fila - 1

# ============================================================ Ajustes
aj = hoja("Ajustes", "AJUSTES — cuadro de stock y ajustes manuales (fase 4)", [6, 12, 46, 14, 34, 18, 26, 34, 13, 16])
aj["A5"] = "CUADRO DE STOCK (existencia final)"
aj["A5"].font = B
stock = [("Existencia según inventario de la farmacia al cierre", True),
         ("Menos: mercadería vencida / destruida", True),
         ("Más / menos: diferencias de valuación (costo de reposición)", True)]
for i, (t, inp) in enumerate(stock):
    r = 6 + i
    aj.cell(r, 3, t)
    c = aj.cell(r, 6)
    entrada(c)
    num(c)
aj.cell(9, 3, "Existencia final ajustada → Resumen Patrimonial e.1").font = B
c = aj.cell(9, 6, '=IF(F6="","",F6+N(F7)+N(F8))')
num(c)
c.font = B
STOCK_CELL = "Ajustes!F9"

aj["A11"] = "AJUSTES MANUALES (suman en la hoja destino; 'Descartado' no suma)"
aj["A11"].font = B
encabezado(aj, 12, ["Nº", "Fecha", "Concepto", "Hoja destino", "Línea (concepto exacto o código)", "Importe (+ suma / − resta)",
                    "Contrapartida", "Respaldo (link o archivo)", "Estado", "Responsable"])
dv1 = DataValidation(type="list", formula1='"Resultados,Patrimonial"', allow_blank=True)
dv2 = DataValidation(type="list", formula1='"Propuesto,Aprobado,Descartado"', allow_blank=True)
aj.add_data_validation(dv1)
aj.add_data_validation(dv2)
for r in range(13, 53):
    aj.cell(r, 1, r - 12)
    for col in range(1, 11):
        aj.cell(r, col).border = CAJA
    for col in (2, 3, 4, 5, 6, 7, 8, 9, 10):
        entrada(aj.cell(r, col))
    num(aj.cell(r, 6))
    dv1.add(aj.cell(r, 4))
    dv2.add(aj.cell(r, 9))
aj.freeze_panes = "A13"

# e.1 con fórmula del cuadro de stock
rp.cell(RPF["e.1"], 4, f'=IF({STOCK_CELL}="","",{STOCK_CELL})')
rp.cell(RPF["e.1"], 4).fill = PatternFill()

# ============================================================ Anexo 1 Bienes de uso
a1 = hoja("Anexo 1", "ANEXO 1 — EVOLUCIÓN DE LOS BIENES DE USO", [26, 15, 13, 13, 15, 15, 13, 15, 15, 16, 16])
encabezado(a1, 5, ["Rubro", "Valor al inicio", "Altas", "Bajas", "Valor al cierre", "Amort. acum. inicio", "Bajas",
                   "Amort. del ejercicio", "Amort. acum. cierre", f"Valor residual {CIERRE}", f"Valor residual {CIERRE_ANT}"])
for i, rubro in enumerate(["Muebles y útiles", "Instalaciones", "Llave de negocio"]):
    r = 6 + i
    a1.cell(r, 1, rubro)
    for col in (2, 3, 4, 6, 7, 8):
        entrada(a1.cell(r, col))
    a1.cell(r, 5, f"=N(B{r})+N(C{r})-N(D{r})")
    a1.cell(r, 9, f"=N(F{r})-N(G{r})+N(H{r})")
    a1.cell(r, 10, f"=E{r}-I{r}")
    a1.cell(r, 11, 0)
    for col in range(2, 12):
        num(a1.cell(r, col))
        a1.cell(r, col).border = CAJA
a1.cell(9, 1, "TOTALES")
for col in range(2, 12):
    num(a1.cell(9, col, f"=SUM({L(col)}6:{L(col)}8)"))
total(a1, 9, 11)
a1["A11"] = "Fuente: balance anterior (valores de origen) + facturas de compra de bienes de uso del ejercicio (COMPRA DE EQUIPOS en el Maestro). Fase 4."
a1["A11"].font = NOTA
rr.cell(AMORT_ROW, 15, f"='Anexo 1'!H9+N{AMORT_ROW}")

# ============================================================ Clasificación gastos → Anexo 2
# Distribución por función del Anexo 2: la del balance anterior de la sociedad (config "anexo2");
# si todavía no está, % de un balance de referencia como punto de partida (se pueden pisar).
PCT_DEFECTO = {
    "Alquileres y Expensas": (1, 0, 0), "Gastos bancarios": (1, 0, 0), "Gastos de mantenimiento": (1, 0, 0),
    "Gastos de sistemas": (0, 1, 0), "Comisiones de tarjetas": (0, 1, 0), "Movilidad y viáticos": (0.35, 0.65, 0),
    "Honorarios profesionales": (0.2845, 0.7155, 0), "Imprenta y Papelería": (0.4, 0.6, 0),
    "Impuesto a los Ingresos Brutos": (0, 1, 0), "Impuestos, tasas y contribuciones": (1, 0, 0), "Intereses": (0, 0, 1),
    "IVA Crédito Fiscal no computable": (0, 1, 0), "Residuos patogénicos": (0, 1, 0),
    "Servicios de luz, gas, teléfono, internet": (0.3, 0.7, 0), "Servicios de vigilancia": (1, 0, 0),
    "Servicios varios": (0.7, 0.3, 0), "Sueldos y cargas sociales": (0, 1, 0), "Amortizaciones de bienes de uso": (1, 0, 0),
}
A2C = COMP.get("anexo2", {})
_cuentas = list(PCT_DEFECTO) + [c for c in A2C if c not in PCT_DEFECTO]
CUENTAS = [(c,) + tuple(A2C[c]) if c in A2C else (c, None, None, None, None) for c in _cuentas]
MAPA = {
    "IIBB": "Impuesto a los Ingresos Brutos",
    "IVA Prorrateo": "IVA Crédito Fiscal no computable",
    "Sueldos y Cargas Sociales": "Sueldos y cargas sociales",
    "Matricula Colegio Farmaceuticos": ("Impuestos, tasas y contribuciones", "A definir con el socio a cargo"),
    "Seguridad e Higiene": "Impuestos, tasas y contribuciones",
    "Publicidad": ("Impuestos, tasas y contribuciones", "PyP es tasa; publicidad comprada ¿va a Servicios varios?"),
    "Marquesinas": "Impuestos, tasas y contribuciones",
    "Metrogas": "Servicios de luz, gas, teléfono, internet",
    "Agua": "Servicios de luz, gas, teléfono, internet",
    "Luz (Edesur)": "Servicios de luz, gas, teléfono, internet",
    "Cámara de Farmacias (cuota)": ("Impuestos, tasas y contribuciones", "¿Cuota gremial con cuenta propia?"),
    "Internet / telefonía": "Servicios de luz, gas, teléfono, internet",
    "Gastos Bancarios": "Gastos bancarios",
    "Servicios financieros": "Gastos bancarios",
    "Intereses de Prestamos": "Intereses",
    "Intereses por facilidades de pago": "Intereses",
    "Comisiones de tarjetas": "Comisiones de tarjetas",
    "Gastos de tarjetas": "Comisiones de tarjetas",
    "Comisiones de Mercado Libre": "Comisiones de tarjetas",
    "Comisiones de Mercado Envíos": ("Comisiones de tarjetas", "¿Movilidad o comisiones?"),
    "Gastos IVA Compras": "Servicios varios",
    "Alquileres": "Alquileres y Expensas",
    "Gastos de comercialización": "Servicios varios",
    "Gastos de mantenimiento": "Gastos de mantenimiento",
    "Gastos de mantenimiento bienes de uso": "Gastos de mantenimiento",
    "Gastos de limpieza": "Gastos de mantenimiento",
    "Gastos de sistemas": "Gastos de sistemas",
    "Gastos de transporte": "Movilidad y viáticos",
    "Combustibles": "Movilidad y viáticos",
    "Gastos varios": "Servicios varios",
    "Refrigerio": "Servicios varios",
    "Gastos de representación": "Servicios varios",
    "Uniformes": "Servicios varios",
    "Seguros": ("Servicios varios", "¿Cuenta propia 'Seguros'?"),
    "Honorarios por servicios": "Honorarios profesionales",
    "Servicios médicos": "Honorarios profesionales",
    "Imprenta y papelería": "Imprenta y Papelería",
    "Residuos patogénicos": "Residuos patogénicos",
    "Seguridad": "Servicios de vigilancia",
    "Indemnización por despido": ("Sueldos y cargas sociales", "¿Exponer aparte?"),
    "Amortización bienes de uso": "Amortizaciones de bienes de uso",
}
MAPA.update(CFG.get("mapa_anexo2", {}))
cl = hoja("Clasificación gastos", "CLASIFICACIÓN DE GASTOS → ANEXO 2 (estandarizar con el socio a cargo)",
          [38, 38, 17, 34, 3, 38, 15, 15, 15, 16, 9, 9, 9])
encabezado(cl, 5, ["Línea del Resumen de Resultados", "Cuenta del Anexo 2", "Total ejercicio", "Observación", "",
                   "Cuenta del Anexo 2", f"Adm. {A_ANT}", f"Comerc. {A_ANT}", f"Financ. {A_ANT}", f"Total {A_ANT}", "% Adm.", "% Com.", "% Fin."])
r = 6
for e, n, f, fa, o, t in LINEAS:
    if t not in ("gasto", "gc", "matr", "planes", "amort"):
        continue
    m = MAPA.get(n, "")
    obs = ""
    if isinstance(m, tuple):
        m, obs = m
    cl.cell(r, 1, n)
    c = cl.cell(r, 2, m)
    entrada(c)
    num(cl.cell(r, 3, f"=INDEX('Resumen Resultados'!$O:$O,MATCH(A{r},'Resumen Resultados'!$A:$A,0))"))
    cl.cell(r, 4, obs)
    for col in range(1, 5):
        cl.cell(r, col).border = CAJA
    r += 1
CL_FIN = r - 1
cl.cell(r, 1, "Total")
num(cl.cell(r, 3, f"=SUM(C6:C{CL_FIN})"))
total(cl, r, 4)
CL_TOT = r
dvc = DataValidation(type="list", formula1=f"=$F$6:$F${5 + len(CUENTAS)}", allow_blank=True)
cl.add_data_validation(dvc)
dvc.add(f"B6:B{CL_FIN}")
for i, (cta, a, cm, fi, tt) in enumerate(CUENTAS):
    rr_ = 6 + i
    cl.cell(rr_, 6, cta)
    for col, v in zip((7, 8, 9, 10), (a, cm, fi, tt)):
        num(cl.cell(rr_, col, v))
    if tt:
        for col, s_ in zip((11, 12, 13), ("G", "H", "I")):
            cl.cell(rr_, col, f"=IF($J{rr_}=0,0,{s_}{rr_}/$J{rr_})").number_format = "0.0%"
    else:
        for col, v in zip((11, 12, 13), PCT_DEFECTO.get(cta, (1, 0, 0))):
            c = cl.cell(rr_, col, v)
            c.number_format = "0.0%"
            entrada(c)
    for col in range(6, 14):
        cl.cell(rr_, col).border = CAJA
cl.cell(5 + len(CUENTAS) + 2, 6, ("Los % por defecto repiten la distribución del Anexo 2 del balance anterior; se pueden pisar." if A2C
       else "Sin balance anterior: los % (amarillo) son los de un balance de referencia como punto de partida; revisar.")).font = NOTA

a2 = hoja("Anexo 2", "ANEXO 2 — CUADRO DE GASTOS (art. 64 ap. I inc. b, Ley 19.550)", [42, 18, 18, 18, 18, 18])
encabezado(a2, 5, ["Nombre de la cuenta", f"Total gastos al {CIERRE}", "Gastos de Administración", "Gastos de Comercialización",
                   "Gastos Financieros", f"Total gastos al {CIERRE_ANT}"])
for i, (cta, a, cm, fi, tt) in enumerate(CUENTAS):
    r = 6 + i
    cr = 6 + i  # misma fila en la tabla de criterio
    a2.cell(r, 1, cta)
    num(a2.cell(r, 2, f"=SUMIFS('Clasificación gastos'!$C:$C,'Clasificación gastos'!$B:$B,A{r})"))
    num(a2.cell(r, 3, f"=B{r}*'Clasificación gastos'!K{cr}"))
    num(a2.cell(r, 4, f"=B{r}*'Clasificación gastos'!L{cr}"))
    num(a2.cell(r, 5, f"=B{r}*'Clasificación gastos'!M{cr}"))
    num(a2.cell(r, 6, tt))
    for col in range(1, 7):
        a2.cell(r, col).border = CAJA
A2_TOT = 6 + len(CUENTAS)
a2.cell(A2_TOT, 1, "TOTALES")
for col in range(2, 7):
    num(a2.cell(A2_TOT, col, f"=SUM({L(col)}6:{L(col)}{A2_TOT - 1})"))
total(a2, A2_TOT, 6)

# ============================================================ Nota 2
n2 = hoja("Nota 2", "NOTA 2 — DETALLE DE LOS PRINCIPALES RUBROS", [52, 20, 20])
encabezado(n2, 5, ["Rubro / cuenta", CIERRE, CIERRE_ANT])
N2TOT = {}
r = 6
for letra, rubro, cuentas in PAT:
    n2.cell(r, 1, f"{letra}) {rubro}").font = B
    for col in range(1, 4):
        n2.cell(r, col).fill = CELESTE
    r += 1
    ini = r
    for cod, cuenta, *_ in cuentas:
        n2.cell(r, 1, "    " + cuenta)
        num(n2.cell(r, 2, f"='Resumen Patrimonial'!F{RPF[cod]}"))
        num(n2.cell(r, 3, f"='Resumen Patrimonial'!G{RPF[cod]}"))
        r += 1
    n2.cell(r, 1, "TOTAL")
    num(n2.cell(r, 2, f"=SUM(B{ini}:B{r - 1})"))
    num(n2.cell(r, 3, f"=SUM(C{ini}:C{r - 1})"))
    total(n2, r, 3)
    N2TOT[letra] = r
    r += 2

# ============================================================ Anexo 3 CMV
a3 = hoja("Anexo 3", "ANEXO 3 — DETERMINACIÓN DEL COSTO DE LAS MERCADERÍAS VENDIDAS", [48, 20, 20])
encabezado(a3, 5, ["", CIERRE, CIERRE_ANT])
filas_a3 = [
    ("EXISTENCIA INICIAL", f"='Resumen Patrimonial'!G{RPF['e.1']}", (COMP.get("anexo3") or [None] * 3)[0]),
    ("MÁS: COMPRAS DEL EJERCICIO (gravadas / no gravadas)", f"='Resumen Resultados'!O{RRF['Compras de mercadería']}", (COMP.get("anexo3") or [None] * 3)[1]),
    ("MENOS: EXISTENCIA FINAL", f"='Resumen Patrimonial'!F{RPF['e.1']}", (COMP.get("anexo3") or [None] * 3)[2]),
]
for i, (t, f, v) in enumerate(filas_a3):
    a3.cell(6 + i, 1, t).font = B
    num(a3.cell(6 + i, 2, f))
    num(a3.cell(6 + i, 3, v))
a3.cell(9, 1, "COSTO DE MERCADERÍAS VENDIDAS")
num(a3.cell(9, 2, "=B6+B7-B8"))
num(a3.cell(9, 3, "=C6+C7-C8" if COMP.get("anexo3") else None))
total(a3, 9, 3)
a3["A11"] = "Si la existencia final no está cargada, el CMV queda igual a las compras + existencia inicial (resultado no definitivo)."
a3["A11"].font = NOTA

# ============================================================ Ganancias
gn = hoja("Ganancias", "IMPUESTO A LAS GANANCIAS — determinación (a completar con el socio a cargo)", [52, 20, 20, 12, 20, 34])
gn["A5"] = "Resultado contable antes de Impuesto a las Ganancias"
gn["A6"] = "Más: gastos no deducibles"
gn["A7"] = "Más / menos: ajuste por inflación impositivo"
gn["A8"] = "Más / menos: otros ajustes impositivos"
gn["A9"] = "Ganancia neta"
gn["A10"] = "Menos: quebrantos de ejercicios anteriores " + (G25.get("nota_quebranto_corta") or "(DDJJ Ganancias del ejercicio anterior)")
gn["A11"] = "GANANCIA NETA IMPONIBLE"
for r in (6, 7, 8, 10):
    entrada(gn.cell(r, 2))
for r in range(5, 12):
    num(gn.cell(r, 2))
    gn.cell(r, 2).border = CAJA
GN_ER = None  # se completa al armar ER
gn["B9"] = "=B5+N(B6)+N(B7)+N(B8)"
gn["B11"] = "=MAX(0,B9-N(B10))"
gn["B10"] = "='Resumen Patrimonial'!F" + str(RPF["f.1"])
gn["A9"].font = gn["A11"].font = B
gn["A13"] = "Escala (Ley 27.630) — ACTUALIZAR a la vigente para el ejercicio"
gn["A13"].font = B
encabezado(gn, 14, ["Ganancia neta imponible acumulada: más de $", "A $", "Pagarán $", "Más el %", "Sobre el excedente de $", "Origen"])
escala = [(0, 34703523.08, 0, 0.25, 0), (34703523.08, 347035230.79, 8675880.77, 0.30, 34703523.08),
          (347035230.79, None, 102375393.08, 0.35, 347035230.79)]
for i, fila_e in enumerate(escala):
    r = 15 + i
    for j, v in enumerate(fila_e):
        c = gn.cell(r, 1 + j, v if v is not None else "En adelante")
        entrada(c)
        c.border = CAJA
        if j == 3:
            c.number_format = "0%"
        elif v is not None:
            num(c)
    gn.cell(r, 6, "Escala del balance 2025 (Nota 5 d)")
gn["A19"] = "¿Escala actualizada para este ejercicio? (SI / NO)"
gn["B19"] = "NO"
entrada(gn["B19"])
gn["A21"] = "IMPUESTO DETERMINADO"
gn["B21"] = "=IF(B11<=0,0,IF(B11<=B15,B11*D15,IF(B11<=B16,C16+(B11-E16)*D16,C17+(B11-E17)*D17)))"
gn["A22"] = "Menos: anticipos (Resumen Patrimonial d.7)"
gn["B22"] = f"='Resumen Patrimonial'!F{RPF['d.7']}"
gn["A23"] = "Menos: retenciones (Resumen Patrimonial d.6)"
gn["B23"] = f"='Resumen Patrimonial'!F{RPF['d.6']}"
gn["A24"] = "Saldo a pagar / (a favor)"
gn["B24"] = "=B21-B22-B23"
for r in range(21, 25):
    num(gn.cell(r, 2))
    gn.cell(r, 2).border = CAJA
gn["A21"].font = gn["A24"].font = B
gn["A26"] = "El impuesto determinado va al Estado de Resultados y a la Provisión (Resumen Patrimonial i.3). Impuesto diferido: pendiente (el socio a cargo)."
gn["A26"].font = NOTA
rp.cell(RPF["i.3"], 4, "=Ganancias!B21")
rp.cell(RPF["i.3"], 4).fill = PatternFill()
rp.cell(RPF["i.3"], 12, "Fórmula: hoja Ganancias, celda B21")

# ============================================================ Estado de Resultados
er = hoja("Estado de Resultados", "ESTADO DE RESULTADOS COMPARATIVO (expresado en pesos)", [58, 22, 22])
encabezado(er, 5, ["Ejercicio finalizado el:", CIERRE, CIERRE_ANT])
RRO = lambda n: f"'Resumen Resultados'!O{RRF[n]}"
ER25 = COMP.get("er") or [None] * 12
ER = [
    ("Ingresos por ventas de bienes", f"={RRO('Total ventas')}", ER25[0], False),
    ("Costo de las mercaderías vendidas (Anexo 3)", "=-'Anexo 3'!B9", ER25[1], False),
    ("Ganancia Bruta", "=B6+B7", ER25[2], True),
    ("Gastos de Administración (Anexo 2)", f"=-'Anexo 2'!C{A2_TOT}", ER25[3], False),
    ("Gastos de Comercialización (Anexo 2)", f"=-'Anexo 2'!D{A2_TOT}", ER25[4], False),
    ("Gastos Financieros (Anexo 2)", f"=-'Anexo 2'!E{A2_TOT}", ER25[5], False),
    ("Ganancia (Pérdida) Ordinaria", "=B8+B9+B10+B11", ER25[6], True),
    ("Resultado Financiero y por Tenencia (incl. RECPAM)", f"={RRO('Resultados financieros y por tenencia (incl. RECPAM)')}", ER25[7], False),
    ("Otros Ingresos", f"={RRO('Otros ingresos')}", ER25[8], False),
    ("Ganancia (Pérdida) del ejercicio antes de Impuesto a las Ganancias", "=B12+B13+B14", ER25[9], True),
    ("Impuesto a las Ganancias", "=-Ganancias!B21", ER25[10], False),
    ("RESULTADO DEL EJERCICIO", "=B15+B16", ER25[11], True),
]
for i, (t, f, v, tot) in enumerate(ER):
    r = 6 + i
    er.cell(r, 1, t)
    num(er.cell(r, 2, f))
    num(er.cell(r, 3, v))
    for col in range(1, 4):
        er.cell(r, col).border = CAJA
    if tot:
        total(er, r, 3)
ER_RES = 17
gn["B5"] = "='Estado de Resultados'!B15"
num(gn["B5"])

# ============================================================ Evolución del PN
pn = hoja("Evolución PN", "ESTADO DE EVOLUCIÓN DEL PATRIMONIO NETO", [30, 15, 16, 16, 14, 14, 18, 18, 19, 19])
encabezado(pn, 5, ["Rubros", "Capital social", "Ajuste de capital", "Total aportes", "Reserva legal", "Otras reservas",
                   "Resultados no asignados", "Total ganancias reservadas", f"Total PN {CIERRE}", f"Total PN {CIERRE_ANT}"])
pn.cell(6, 1, "Saldos al inicio")
PN0 = COMP.get("pn_inicio") or {}
for col, k in zip((2, 3, 5, 6, 7), ("capital", "ajuste", "reserva_legal", "otras_reservas", "rna")):
    c = pn.cell(6, col, PN0.get(k))
    num(c)
    if not PN0:
        entrada(c)
pn.cell(7, 1, "Distribución de utilidades / dividendos")
for col in (2, 3, 5, 6, 7):
    entrada(pn.cell(7, col))
    num(pn.cell(7, col))
pn.cell(8, 1, "Resultado del ejercicio")
num(pn.cell(8, 7, f"='Estado de Resultados'!B{ER_RES}"))
pn.cell(9, 1, "Saldos al cierre")
for r in (6, 7, 8, 9):
    if r == 9:
        for col in (2, 3, 5, 6, 7):
            num(pn.cell(9, col, f"=SUM({L(col)}6:{L(col)}8)"))
    num(pn.cell(r, 4, f"=N(B{r})+N(C{r})"))
    num(pn.cell(r, 8, f"=N(E{r})+N(F{r})+N(G{r})"))
    num(pn.cell(r, 9, f"=D{r}+H{r}"))
    for col in range(1, 11):
        pn.cell(r, col).border = CAJA
for r, v in zip((6, 7, 8, 9), COMP.get("pn_2025") or [None] * 4):
    num(pn.cell(r, 10, v))
total(pn, 9, 10)
pn["A11"] = ("Saldos al inicio = saldos al cierre del balance legalizado anterior." if PN0
             else "Pendiente: falta el balance legalizado anterior; cargar los saldos al inicio (amarillo).")
pn["A11"].font = NOTA

# ============================================================ Situación Patrimonial
sp = hoja("Situación Patrimonial", "ESTADO DE SITUACIÓN PATRIMONIAL COMPARATIVO (expresado en pesos)", [52, 22, 22])
encabezado(sp, 5, ["", CIERRE, CIERRE_ANT])
N2 = lambda letra, col: f"='Nota 2'!{col}{N2TOT[letra]}"
SP = [
    ("ACTIVO", None, None, "titulo"),
    ("ACTIVO CORRIENTE", None, None, "sub"),
    ("Caja y Bancos (Nota 2 - a)", N2("a", "B"), N2("a", "C"), ""),
    ("Inversiones (Nota 2 - b)", N2("b", "B"), N2("b", "C"), ""),
    ("Créditos por Ventas (Nota 2 - c)", N2("c", "B"), N2("c", "C"), ""),
    ("Créditos Fiscales (Nota 2 - d)", N2("d", "B"), N2("d", "C"), ""),
    ("Bienes de Cambio (Nota 2 - e)", N2("e", "B"), N2("e", "C"), ""),
    ("TOTAL DE ACTIVO CORRIENTE", "=SUM(B8:B12)", "=SUM(C8:C12)", "tot"),
    ("ACTIVO NO CORRIENTE", None, None, "sub"),
    ("Bienes de Uso (Anexo 1)", "='Anexo 1'!J9", "='Anexo 1'!K9", ""),
    ("Créditos Fiscales no corrientes (Nota 2 - f)", N2("f", "B"), N2("f", "C"), ""),
    ("TOTAL DE ACTIVO NO CORRIENTE", "=B15+B16", "=C15+C16", "tot"),
    ("TOTAL DE ACTIVO", "=B13+B17", "=C13+C17", "tot"),
    ("PASIVO", None, None, "titulo"),
    ("PASIVO CORRIENTE", None, None, "sub"),
    ("Deudas Comerciales (Nota 2 - g)", N2("g", "B"), N2("g", "C"), ""),
    ("Deudas Sociales (Nota 2 - h)", N2("h", "B"), N2("h", "C"), ""),
    ("Deudas Fiscales (Nota 2 - i)", N2("i", "B"), N2("i", "C"), ""),
    ("Deudas Financieras (Nota 2 - j)", N2("j", "B"), N2("j", "C"), ""),
    ("TOTAL DE PASIVO CORRIENTE", "=SUM(B21:B24)", "=SUM(C21:C24)", "tot"),
    ("PASIVO NO CORRIENTE", None, None, "sub"),
    ("Otras Deudas no corrientes (Nota 2 - k)", N2("k", "B"), N2("k", "C"), ""),
    ("TOTAL DE PASIVO NO CORRIENTE", "=B27", "=C27", "tot"),
    ("TOTAL DE PASIVO", "=B25+B28", "=C25+C28", "tot"),
    ("PATRIMONIO NETO (según estado correspondiente)", "='Evolución PN'!I9", "='Evolución PN'!J9", "tot"),
    ("TOTAL DE PASIVO Y PATRIMONIO NETO", "=B29+B30", "=C29+C30", "tot"),
    ("Diferencia Activo − (Pasivo + PN)  → debe dar 0", "=ROUND(B18-B31,2)", "=ROUND(C18-C31,2)" if COMP else None, "dif"),
]
for i, (t, fb, fc, k) in enumerate(SP):
    r = 6 + i
    sp.cell(r, 1, t)
    if fb:
        num(sp.cell(r, 2, fb))
        num(sp.cell(r, 3, fc))
    if k == "titulo":
        sp.cell(r, 1).font = Font(bold=True, underline="single")
    elif k == "sub":
        sp.cell(r, 1).font = B
    elif k == "tot":
        total(sp, r, 3)
    elif k == "dif":
        sp.cell(r, 1).font = NOTA
    if k in ("", "tot", "dif"):
        for col in range(1, 4):
            sp.cell(r, col).border = CAJA
assert sp["A18"].value == "TOTAL DE ACTIVO" and sp["A31"].value.startswith("TOTAL DE PASIVO Y")

# ============================================================ Flujo de efectivo
fe = hoja("Flujo de Efectivo", "ESTADO DE FLUJO DE EFECTIVO COMPARATIVO (método indirecto)", [62, 22, 22])
encabezado(fe, 5, ["", CIERRE, CIERRE_ANT])
EF = lambda col: f"='Situación Patrimonial'!{col}8+'Situación Patrimonial'!{col}9"
var_act = lambda f: f"=-('Situación Patrimonial'!B{f}-'Situación Patrimonial'!C{f})"
var_pas = lambda f: f"='Situación Patrimonial'!B{f}-'Situación Patrimonial'!C{f}"
FE25 = COMP.get("flujo") or [None] * 22
FE = [
    ("Variaciones del efectivo", None, None, "sub"),
    ("Efectivo al inicio del ejercicio", "=C8", FE25[0], ""),                       # 7
    ("Efectivo al cierre del ejercicio", EF("B"), FE25[1], ""),                     # 8
    ("Aumento (Disminución) neto de efectivo", "=B8-B7", FE25[2], "tot"),          # 9
    ("Causas de las variaciones del efectivo", None, None, "sub"),                      # 10
    ("Actividades operativas", None, None, "sub"),                                      # 11
    ("Resultado del ejercicio", f"='Estado de Resultados'!B{ER_RES}", FE25[3], ""),  # 12
    ("Depreciación de bienes de uso", "='Anexo 1'!H9", FE25[4], ""),                          # 13
    ("Créditos por ventas", var_act(10), FE25[5], ""),                             # 14
    ("Créditos fiscales", var_act(11), FE25[6], ""),                                # 15
    ("Bienes de cambio", var_act(12), FE25[7], ""),                                 # 16
    ("Créditos fiscales no corrientes", var_act(16), FE25[8], ""),                    # 17
    ("Deudas comerciales", var_pas(21), FE25[9], ""),                             # 18
    ("Deudas sociales", var_pas(22), FE25[10], ""),                                  # 19
    ("Deudas fiscales", var_pas(23), FE25[11], ""),                                  # 20
    ("Deudas financieras", var_pas(24), FE25[12], ""),                                   # 21
    ("Otras deudas no corrientes", var_pas(27), FE25[13], ""),                                 # 22
    ("Variación RECPAM / resultados financieros", "=-'Estado de Resultados'!B13", FE25[14], ""),  # 23
    ("Flujo neto de efectivo de las actividades operativas", "=SUM(B12:B23)", FE25[15], "tot"),  # 24
    ("Actividades de inversión", None, None, "sub"),                                    # 25
    ("Pago por compras de bienes de uso", "=-'Anexo 1'!C9", FE25[16], ""),                     # 26
    ("Flujo neto de efectivo de las actividades de inversión", "=B26", FE25[17], "tot"),       # 27
    ("Actividades de financiación", None, None, "sub"),                                 # 28
    ("Resultados financieros y por tenencia generados por efectivo", "='Estado de Resultados'!B13", FE25[18], ""),  # 29
    ("Distribución de utilidades", "='Evolución PN'!I7", FE25[19], ""),                        # 30
    ("Flujo neto de efectivo de las actividades de financiación", "=B29+B30", FE25[20], "tot"),  # 31
    ("Aumento (Disminución) neto de efectivo", "=B24+B27+B31", FE25[21], "tot"),   # 32
    ("Diferencia con la variación del efectivo → debe dar 0", "=ROUND(B32-B9,2)", "=ROUND(C32-C9,2)" if COMP.get("flujo") else None, "dif"),
]
for i, (t, fb, fc, k) in enumerate(FE):
    r = 6 + i
    fe.cell(r, 1, t)
    if fb is not None:
        num(fe.cell(r, 2, fb))
        num(fe.cell(r, 3, fc))
    if k == "sub":
        fe.cell(r, 1).font = B
    elif k == "tot":
        total(fe, r, 3)
    elif k == "dif":
        fe.cell(r, 1).font = NOTA
    if k != "sub":
        for col in range(1, 4):
            fe.cell(r, col).border = CAJA
assert fe["A24"].value.startswith("Flujo neto de efectivo de las actividades operativas")
assert fe["A32"].value.startswith("Aumento")
assert sp["A10"].value.startswith("Créditos por Ventas") and sp["A27"].value.startswith("Otras Deudas")
assert er["A17"].value == "RESULTADO DEL EJERCICIO" and er["A13"].value.startswith("Resultado Financiero")

# ============================================================ Respaldos (copiados del EERR 2026)
def copiar_valores(nombre_src, nombre_dst, nota):
    s = src[nombre_src]
    d = wb.create_sheet(nombre_dst)
    d["A1"] = nota
    d["A1"].font = NOTA
    for row in s.iter_rows():
        for c in row:
            if c.value is not None:
                n = d.cell(c.row + 1, c.column, c.value)
                if isinstance(c.value, (int, float)):
                    n.number_format = NUM
    for col, dim in s.column_dimensions.items():
        if dim.width:
            d.column_dimensions[col].width = dim.width
    d.sheet_view.showGridLines = False
    return d

copiar_valores("Compras imputadas", "Compras imputadas",
               "Respaldo: copia en valores de la hoja 'Compras imputadas' del EERR 2026 (libro de compras × Maestro de Proveedores).")
copiar_valores("Sueldos", "Sueldos", "Respaldo: copia en valores de la hoja 'Sueldos' del EERR 2026 (F.931).")

# ============================================================ Origen por mes
og = hoja("Origen por mes", "ORIGEN DE LOS DATOS POR MES (gastos, CMV, Matrícula y FCI)", [12, 16, 46, 18, 18, 18, 60])
r = 5
og.cell(r, 1, "1. Gastos del libro de compras: de dónde sale cada mes").font = B
r += 1
encabezado(og, r, ["Mes", "Origen", "Detalle", "", "", "", ""])
r += 1
for i, m in enumerate(fuentes.MESES_AAAAMM):
    src, nota = ORIGEN[m]
    og.cell(r, 1, MESES[i])
    c = og.cell(r, 2, "Papel de trabajo" if src == "Papel" else "Libro + Maestro")
    c.fill = PatternFill("solid", fgColor="C6EFCE" if src == "Papel" else "FFE699")
    og.cell(r, 3, nota)
    og.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    r += 1
og.cell(r, 1, "Acuerdo del 12/11: los gastos salen del papel de trabajo mensual cuando el mes está clasificado (columna PROVEEDORES); "
              "los meses sin clasificar siguen con el libro de compras y el Maestro.").font = NOTA
r += 2

og.cell(r, 1, "2. Meses con papel: papel de trabajo vs libro + Maestro, por concepto").font = B
r += 1
encabezado(og, r, ["Mes", "Concepto", "", "Papel de trabajo", "Libro + Maestro", "Diferencia", "Observación"])
r += 1
nombre_rr = {e: n for e, n, *_ in LINEAS}
nombre_rr["CAMARA"] = "Cámara de Farmacias (cuota)"
for m, conc, vp, vm in COMPARACION:
    og.cell(r, 1, MESES[fuentes.MESES_AAAAMM.index(m)])
    og.cell(r, 2, nombre_rr.get(conc, conc))
    num(og.cell(r, 4, vp)); num(og.cell(r, 5, vm)); num(og.cell(r, 6, f"=D{r}-E{r}"))
    if conc == "CMV":
        ncr = NCR.get(m)
        og.cell(r, 7, f"El CMV del Balance sigue con el libro sin NCR. NCR del papel (hoja NCR): {ncr:,.2f}" if ncr else "El CMV del Balance sigue con el libro sin NCR")
    elif abs(vp - vm) > 1:
        og.cell(r, 7, "Distinta clasificación entre el papel y el Maestro")
    for col in range(1, 8):
        og.cell(r, col).border = CAJA
    r += 1
og.cell(r, 1, "En el CMV la diferencia son las NCR: el papel de trabajo las resta de las compras y el EERR (criterio del 05/11) no. A definir con el socio a cargo.").font = NOTA
r += 2

og.cell(r, 1, "3. Otras fuentes").font = B
r += 1
og.cell(r, 1, "Matrícula Colegio (T&S): " + "; ".join(MATR_NOTAS))
r += 1
og.cell(r, 1, "FCI (rendimiento y saldos): " + ", ".join(FCI_ARCH))
r += 1
og.cell(r, 1, "Planes de facilidades ARCA: INFORME_PLANES_FACILIDADES_20260925-1_2.xlsx (hojas PLANES VIGENTES y DETALLE CUOTAS)")

# ============================================================ Controles
ct = hoja("Controles", "CONTROLES DEL BALANCE", [6, 70, 20, 14, 60])
encabezado(ct, 5, ["Nº", "Control", "Valor", "Resultado", "Qué hacer"])
RR_V = RRF["Ventas exentas y no gravadas"]
RR_G = RRF["Ventas gravadas"]
RR_S = RRF["Sueldos y Cargas Sociales"]
CONTROLES = [
    ("Situación Patrimonial: Activo − (Pasivo + PN) al cierre", "='Situación Patrimonial'!B32",
     '=IF(C{r}=0,"OK",IF(D7>0,"PENDIENTE","REVISAR"))', "No cierra mientras haya saldos patrimoniales pendientes (control 7)."),
    (f"Situación Patrimonial {A_ANT} (comparativo) cuadra", "='Situación Patrimonial'!C32",
     '=IF(ABS(C{r})<0.05,"OK","REVISAR")', "Transcripción del balance legalizado 2025."),
    ("Flujo de efectivo: variación explicada = variación de caja + inversiones", "='Flujo de Efectivo'!B33",
     '=IF(ABS(C{r})<0.05,"OK",IF(D7>0,"PENDIENTE","REVISAR"))', "Cierra solo cuando cierra la Situación Patrimonial."),
    ("Anexo 2 = total de gastos del Resumen de Resultados",
     f"=ROUND('Anexo 2'!B{A2_TOT}-'Resumen Resultados'!O{RRF['Total gastos']},2)",
     '=IF(ABS(C{r})<0.05,"OK","REVISAR")', "Alguna línea de gastos sin cuenta del Anexo 2 o con cuenta mal escrita."),
    ("Líneas de gastos sin cuenta del Anexo 2", f"=COUNTBLANK('Clasificación gastos'!B6:B{CL_FIN})",
     '=IF(C{r}=0,"OK","REVISAR")', "Completar la columna B de 'Clasificación gastos'."),
    ("Saldos patrimoniales pendientes de cargar", f"=COUNTIF('Resumen Patrimonial'!J6:J{RP_FIN},\"Pendiente\")",
     '=IF(C{r}=0,"OK","PENDIENTE")', "Ver columna Fase en 'Resumen Patrimonial' (1 automática … 4 manual)."),
    ("Saldos patrimoniales provisorios (dato anterior al cierre)", f"=COUNTIF('Resumen Patrimonial'!J6:J{RP_FIN},\"Provisorio\")",
     '=IF(C{r}=0,"OK","PENDIENTE")', "Actualizar con el dato al 30/11 cuando esté (ver columna 'Corte del dato')."),
    ("Meses de ventas sin cargar", f"=COUNTBLANK('Resumen Resultados'!B{RR_V}:M{RR_V})+COUNTBLANK('Resumen Resultados'!B{RR_G}:M{RR_G})",
     '=IF(C{r}=0,"OK","PENDIENTE")', "Cargar desde el papel de liquidaciones (armado-eerr)."),
    ("Meses de sueldos sin cargar (F.931)", f"=COUNTBLANK('Resumen Resultados'!B{RR_S}:M{RR_S})",
     '=IF(C{r}=0,"OK","PENDIENTE")', "Bajar F.931 y cargar."),
    ("Líneas de resultados sin fuente", f"=SUMPRODUCT(('Resumen Resultados'!A6:A{max(RRF.values())}<>\"\")*('Resumen Resultados'!P6:P{max(RRF.values())}=\"\")*(LEFT('Resumen Resultados'!A6:A{max(RRF.values())},5)<>\"Total\")*('Resumen Resultados'!Q6:Q{max(RRF.values())}<>\"\"))",
     '=IF(C{r}=0,"OK","REVISAR")', "Toda línea debe decir de dónde sale."),
    ("Ajustes propuestos sin aprobar", '=COUNTIF(Ajustes!I13:I52,"Propuesto")',
     '=IF(C{r}=0,"OK","PENDIENTE")', "Los propuestos ya suman; el socio a cargo los aprueba o descarta."),
    ("Escala de Ganancias actualizada para el ejercicio", "=Ganancias!B19",
     '=IF(C{r}="SI","OK","PENDIENTE")', "Hoy usa la escala del balance 2025."),
    ("Reexpresión en moneda constante (RT 6 / RECPAM)", "Sin implementar",
     '="PENDIENTE"', "El prototipo está en valores nominales. Definir método con el socio a cargo."),
]
for i, (t, v, res, que) in enumerate(CONTROLES):
    r = 6 + i
    ct.cell(r, 1, i + 1)
    ct.cell(r, 2, t)
    c = ct.cell(r, 3, v)
    if i < 4:
        num(c)
    ct.cell(r, 4, res.format(r=r)).alignment = Alignment(horizontal="center")
    ct.cell(r, 5, que)
    for col in range(1, 6):
        ct.cell(r, col).border = CAJA
# el control 1 y 3 miran la fila del control 6 (pendientes) → D7 apunta mal; corregir a C11
for r in (6, 8):
    ct.cell(r, 4).value = ct.cell(r, 4).value.replace("D7>0", "C11>0")
rojo = PatternFill("solid", fgColor="F8CBAD")
verde = PatternFill("solid", fgColor="C6EFCE")
naranja = PatternFill("solid", fgColor="FFE699")
fin_ct = 5 + len(CONTROLES)
ct.conditional_formatting.add(f"D6:D{fin_ct}", CellIsRule(operator="equal", formula=['"OK"'], fill=verde))
ct.conditional_formatting.add(f"D6:D{fin_ct}", CellIsRule(operator="equal", formula=['"REVISAR"'], fill=rojo))
ct.conditional_formatting.add(f"D6:D{fin_ct}", CellIsRule(operator="equal", formula=['"PENDIENTE"'], fill=naranja))
rp.conditional_formatting.add(f"J6:J{RP_FIN}", CellIsRule(operator="equal", formula=['"Pendiente"'], fill=naranja))
rp.conditional_formatting.add(f"J6:J{RP_FIN}", CellIsRule(operator="equal", formula=['"Cargado"'], fill=verde))

# ============================================================ Índice
ix = hoja("Índice", f"BALANCE — PROTOTIPO ({CORTO}, ejercicio Nº {EJ_NRO}: 01/{MESES[0].replace('-', '/')} – {CIERRE})", [30, 70, 60])
ix["A4"] = (f"Generado el {datetime.now():%d/%m/%Y %H:%M}. Las liquidaciones mensuales (IVA, IIBB, Seguridad e Higiene, Publicidad) NO están acá: "
            "viven en el papel de liquidaciones. Este libro toma de ahí los totales por mes y suma lo que el papel no tiene.")
ix["A4"].font = NOTA
ix["A4"].alignment = Alignment(wrap_text=True)
ix.merge_cells("A4:C4")
ix.row_dimensions[4].height = 30
encabezado(ix, 6, ["Hoja", "Qué es", "De dónde sale"])
INDICE = [
    ("A. ESTADOS PARA PRESENTAR", "", ""),
    ("Situación Patrimonial", "Estado de Situación Patrimonial comparativo", "Nota 2, Anexo 1 y Evolución PN"),
    ("Estado de Resultados", "Estado de Resultados comparativo", "Resumen de Resultados, Anexos 2 y 3, Ganancias"),
    ("Evolución PN", "Estado de Evolución del Patrimonio Neto", f"Saldos al cierre {A_ANT} + resultado del ejercicio"),
    ("Flujo de Efectivo", "Flujo de efectivo, método indirecto", "Variaciones de la Situación Patrimonial"),
    ("Nota 2", "Detalle de los principales rubros (a–k)", "Resumen Patrimonial"),
    ("Anexo 1", "Evolución de los bienes de uso", "Manual (balance anterior + altas del ejercicio)"),
    ("Anexo 2", "Cuadro de gastos (Adm. / Comerc. / Financ.)", "Clasificación gastos"),
    ("Anexo 3", "Costo de las mercaderías vendidas", "Existencias (Resumen Patrimonial) + compras"),
    ("B. HOJAS DE TRABAJO", "", ""),
    ("Resumen Resultados", "Resultados mes por mes, con Fuente y Fase en cada línea", "Papel de liquidaciones, libro de compras + Maestro, F.931"),
    ("Resumen Patrimonial", "Saldos al cierre de cada cuenta, con Fuente, Fase y Estado", "Extractos, papel, SCT, Mis Facilidades, F.931, pedidos"),
    ("Clasificación gastos", "Línea de gastos → cuenta del Anexo 2 y % por función", "Criterio a estandarizar con el socio a cargo"),
    ("Origen por mes", "De dónde sale cada mes de los gastos, comparación papel vs Maestro, CMV y NCR", "Papeles mensuales, EERR 2026, T&S, DRIVE - FCI"),
    ("Ajustes", "Cuadro de stock + ajustes manuales con estado", "Fase 4 (manual, con respaldo)"),
    ("Ganancias", "Determinación del impuesto", "A completar con el socio a cargo"),
    ("Controles", "Semáforo de cierre del balance", "Fórmulas sobre todo el libro"),
    ("C. RESPALDOS", "", ""),
    ("Compras imputadas", "Compras por proveedor y mes", "EERR 2026 (valores)"),
    ("Sueldos", "F.931 por mes", "EERR 2026 (valores)"),
]
for i, (h, q, d) in enumerate(INDICE):
    r = 7 + i
    c = ix.cell(r, 1, h)
    if not q:
        c.font = B
        for col in range(1, 4):
            ix.cell(r, col).fill = CELESTE
    else:
        c.hyperlink = f"#'{h}'!A1"
        c.font = Font(color="0563C1", underline="single")
    ix.cell(r, 2, q)
    ix.cell(r, 3, d)
r = 7 + len(INDICE) + 1
ix.cell(r, 1, "FASES").font = B
for j, t in enumerate(["1 — Automática: sale de un script o del papel sin intervención",
                       "2 — Parcial: hay fuente automática pero hay que completar o validar",
                       "3 — Pedido externo: SETI o la farmacia",
                       "4 — Ajuste manual con respaldo"]):
    ix.cell(r + 1 + j, 1, t)
ix.cell(r + 6, 1, "COLORES").font = B
ix.cell(r + 7, 1, "Celda amarilla = se carga a mano").fill = AMARILLO
ix.cell(r + 8, 1, "Sin color = fórmula o dato copiado")

# orden de hojas
orden = ["Índice", "Controles", "Situación Patrimonial", "Estado de Resultados", "Evolución PN", "Flujo de Efectivo",
         "Nota 2", "Anexo 1", "Anexo 2", "Anexo 3", "Resumen Resultados", "Resumen Patrimonial", "Clasificación gastos",
         "Origen por mes", "Ajustes", "Ganancias", "Compras imputadas", "Sueldos"]
wb._sheets = [wb[n] for n in orden]
for n in orden[2:10]:
    wb[n].sheet_properties.tabColor = "1F3864"
for n in orden[10:16]:
    wb[n].sheet_properties.tabColor = "70AD47"
for n in orden[16:]:
    wb[n].sheet_properties.tabColor = "A6A6A6"
wb.active = 0
for ws in wb:
    ws.page_setup.orientation = "portrait" if ws.title in orden[2:10] else "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
print("ok", OUT)
