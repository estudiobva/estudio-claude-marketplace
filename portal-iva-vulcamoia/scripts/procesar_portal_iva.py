#!/usr/bin/env python3
"""
Procesa las descargas del Portal IVA (ARCA) y completa los inputs de COMPRAS y VENTAS.

Uso:
  python procesar_portal_iva.py --zips <zip1> <zip2> ... \
      --compras <input_portal_COMPRAS.xlsx> --ventas <input_portal_VENTAS.xlsx> \
      --padron <Clasificacion_Proveedores.xlsx> --out <carpeta_salida>

Genera en --out:
  input_portal_COMPRAS.xlsx  (PORTAL IVA + Control con Portal IVA + Resumen por Clasificación + Proveedores sin clasificar)
  input_portal_VENTAS.xlsx   (PORTAL IVA completo + hoja Comprobantes Ventas)
  Clasificacion_Proveedores_*.xlsx (padrón con variantes de escritura unificadas)
  resumen.json               (conteos, sin clasificar, controles) para informar al usuario
Después hay que recalcular ambos inputs con el recalc.py del skill xlsx.
"""
import argparse, glob, json, os, re, shutil, zipfile, datetime as dt, unicodedata, copy
import xml.etree.ElementTree as ET
from collections import Counter
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as CL

NS = {'a': 'http://fisca.afip.gob.ar/xml'}
NUM = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
NC_TIPOS = (3, 8, 13)            # notas de crédito A, B, C
TIPO_C = (11, 13)                # factura y NC tipo C: no generan crédito fiscal
LAST = 4186                      # última fila con formato en la plantilla de COMPRAS
YELLOW = PatternFill('solid', fgColor='FFFF00')
ORANGE = PatternFill('solid', fgColor='FFC000')
HDR = PatternFill('solid', fgColor='BDD7EE')
THIN = Side(style='thin', color='A6A6A6')

# Variantes de escritura ya detectadas en el padrón -> forma canónica
VARIANTES = {'Mercaderias': 'Mercaderías', 'Mercaderia': 'Mercaderías',
             'Gasto de mantenimiento': 'Gastos de mantenimiento',
             'Gastos Varios': 'Gastos varios', 'Rodado': 'Rodados'}


def canon(v):
    if not isinstance(v, str):
        return v
    v = v.strip()
    return VARIANTES.get(v, v)


def clave(s):
    s = unicodedata.normalize('NFKD', str(s).strip().lower())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.rstrip('s').replace('gastos ', 'gasto ')


def leer_csv(path):
    return pd.read_csv(path, sep=';', encoding='latin1', decimal=',', dtype={'Número de Despacho': str})


def extraer(zips, work):
    periodos = {}
    for z in zips:
        m = re.search(r'(20\d{2})(\d{2})', os.path.basename(z))
        if not m:
            raise SystemExit(f'No encuentro el período (AAAAMM) en el nombre: {z}')
        p = m.group(1) + m.group(2)
        d = os.path.join(work, p)
        os.makedirs(d, exist_ok=True)
        with zipfile.ZipFile(z) as zf:
            zf.extractall(d)
        # por si el zip trae una subcarpeta
        for f in glob.glob(d + '/**/*', recursive=True):
            if os.path.isfile(f) and os.path.dirname(f) != d:
                shutil.move(f, os.path.join(d, os.path.basename(f)))
        periodos[p] = d
    return dict(sorted(periodos.items()))


def xml_consol(d, libro):
    r = ET.parse(os.path.join(d, 'libro_iva.xml')).getroot()
    return r.find(f'a:{libro}/a:informacion_consolidada', NS)


def xval(x, paths):
    vals = []
    for q in paths:
        t = x.findtext('a:' + q.replace('/', '/a:'), namespaces=NS) if x is not None else None
        vals.append(round(float(t or 0), 2))
    nz = [v for v in vals if v]
    if not nz:
        return 0
    return nz[0] if len(nz) == 1 else '=' + '+'.join(f'{v:.2f}' for v in nz)


def xnum(v):
    if isinstance(v, str):
        return sum(float(t) for t in v[1:].split('+'))
    return v


# ---------------------------------------------------------------- padrón
def cargar_padron(path, out_dir):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    cambios = 0
    for r in range(2, ws.max_row + 1):
        c = ws.cell(r, 3)
        nv = canon(c.value)
        if nv != c.value:
            c.value = nv
            cambios += 1
    out = os.path.join(out_dir, os.path.basename(path))
    wb.save(out)
    pad = pd.read_excel(out)
    pad['CLASIFICACION'] = pad['CLASIFICACION'].astype(str).str.strip()
    clasif, conflictos = {}, {}
    for cuit, g in pad.groupby('IDENTIFTRI'):
        cnt = Counter(g['CLASIFICACION'])
        clasif[int(cuit)] = cnt.most_common(1)[0][0]
        if len(cnt) > 1:
            conflictos[int(cuit)] = dict(cnt)
    # variantes nuevas no contempladas en VARIANTES
    grupos = {}
    for v in pad['CLASIFICACION'].unique():
        grupos.setdefault(clave(v), set()).add(v)
    otras = [sorted(s) for s in grupos.values() if len(s) > 1]
    return out, clasif, conflictos, cambios, otras


# ---------------------------------------------------------------- COMPRAS
def armar_compras(periodos, tmpl, clasif, out_dir):
    rows = []
    for p, d in periodos.items():
        for fn, impo in (('comprobantes_compras.csv', False), ('comprobantes_compras_impo_bienes.csv', True)):
            f = os.path.join(d, fn)
            if not os.path.exists(f):
                continue
            for _, r in leer_csv(f).iterrows():
                rows.append(dict(
                    fecha=pd.to_datetime(r['Fecha de Emisión']), tipo=int(r['Tipo de Comprobante']),
                    pv=None if impo else int(r['Punto de Venta']),
                    nro=str(r['Número de Despacho']) if impo else int(r['Número de Comprobante']),
                    cuit=int(r['Nro. Doc. Vendedor']), nom=r['Denominación Vendedor'],
                    neto=r['Total Neto Gravado'], ng=r['Importe No Gravado'], ex=r['Importe Exento'],
                    iva=r['Total IVA'], tot=r['Importe Total'], per=p))
    c = pd.DataFrame(rows)
    for k in ['neto', 'ng', 'ex', 'iva', 'tot']:
        c[k] = pd.to_numeric(c[k], errors='coerce').fillna(0).round(2)
    c['ns'] = c['nro'].astype(str).str.zfill(20)
    c = c.sort_values(['fecha', 'tipo', 'pv', 'ns'], na_position='last').reset_index(drop=True)
    c['clasif'] = [clasif.get(x) for x in c.cuit]   # None si no está en el padrón

    wb = openpyxl.load_workbook(tmpl)
    ws = wb['PORTAL IVA']
    for i, r in c.iterrows():
        n = 3 + i
        ws.cell(n, 1, r.fecha.to_pydatetime()).number_format = 'mm-dd-yy'
        ws.cell(n, 2, int(r.tipo))
        ws.cell(n, 3, None if pd.isna(r.pv) else int(r.pv))
        ws.cell(n, 4, r.nro)
        ws.cell(n, 5, int(r.cuit))
        ws.cell(n, 6, r.nom)
        for j, k in zip(range(7, 12), ['neto', 'ng', 'ex', 'iva', 'tot']):
            ws.cell(n, j, float(r[k])).number_format = NUM
        # tipo C: sin discriminar -> toma el importe total
        ws.cell(n, 12, f'=IF(OR(B{n}=11,B{n}=13),K{n},G{n}+H{n}+I{n})').number_format = NUM
        m = ws.cell(n, 13)
        if isinstance(r.clasif, str) and r.clasif:
            m.value = r.clasif
        else:
            m.value = 'SIN CLASIFICAR'
            m.fill = ORANGE
    ws['L2'].comment = Comment('Comprobantes tipo C (11 y 13), que no generan crédito fiscal: toma el Importe Total (col. K). Resto: G+H+I.', 'Claude')
    last = 2 + len(c)
    ws.auto_filter.ref = f'A2:M{last}'
    ws.freeze_panes = 'A3'

    anio = int(list(periodos)[0][:4])
    hoja_control_compras(wb, periodos, anio)
    hoja_resumen(wb, anio)
    sc = hoja_sin_clasificar(wb, c)
    out = os.path.join(out_dir, os.path.basename(tmpl))
    wb.save(out)
    return out, c, sc


def hoja_control_compras(wb, periodos, anio):
    nm = 'Control con Portal IVA'
    if nm in wb.sheetnames:
        del wb[nm]
    ws = wb.create_sheet(nm, 1)
    F = Font(name='Calibri', size=11); B = Font(name='Calibri', size=11, bold=True)
    blue = Font(name='Calibri', size=11, color='0000FF')
    ws['A1'] = 'Control con Portal IVA'
    for m in range(1, 13):
        ws.cell(1, 1 + m, m)
    ws.cell(1, 14, 'Total')
    for j in range(1, 15):
        c = ws.cell(1, j); c.font = B; c.fill = HDR; c.border = Border(bottom=Side(style='thin'))
        c.alignment = Alignment(horizontal='left' if j == 1 else 'center')
    labels = {3: ('Neto Gravado', 1), 5: ('Operaciones de Compras', 0), 6: ('Notas de Crédito de Compras', 0),
              8: ('Exento o No Gravado', 1), 10: ('Operaciones de Compras', 0), 11: ('Notas de Crédito de Compras', 0),
              12: ('Operaciones que no generan Crédito Fiscal', 0)}
    for r, (t, b) in labels.items():
        ws.cell(r, 1, t).font = B if b else F
    src = {5: ['totales_compras/total_neto_gravado', 'totales_bienes_usados/total_neto_gravado', 'totales_importacion_bienes/total_neto_gravado'],
           6: ['totales_compras_mlcf_restituir/total_neto_gravado'],
           10: ['total_exento_no_gravado'], 11: ['total_exento_no_gravado_nc'], 12: ['total_compras_no_fiscal']}
    for r, qs in src.items():
        ws.cell(r, 1).comment = Comment('libro_iva.xml > iva_compras: ' + ' + '.join(qs), 'Claude')
        for m in range(1, 13):
            c = ws.cell(r, 1 + m); c.number_format = NUM; c.font = blue
        ws.cell(r, 14, f'=SUM(B{r}:M{r})').number_format = NUM
    for p, d in periodos.items():
        x = xml_consol(d, 'iva_compras')
        for r, qs in src.items():
            ws.cell(r, 1 + int(p[4:])).value = xval(x, qs)
    for j in range(1, 15):
        L = CL(j)
        c = ws.cell(13, j, None if j == 1 else f'={L}5-{L}6+{L}10-{L}11+{L}12')
        c.number_format = NUM; c.font = B; c.border = Border(top=Side(style='thin'), bottom=Side(style='double'))
    P = "'PORTAL IVA'!"; A = f"{P}$A$3:$A${LAST}"; Lc = f"{P}$L$3:$L${LAST}"
    ws.cell(15, 1, 'Total según detalle (hoja PORTAL IVA)').font = F
    ws.cell(16, 1, 'Diferencia').font = B
    for m in range(1, 13):
        L = CL(1 + m)
        ws.cell(15, 1 + m, f'=SUMIFS({Lc},{A},">="&DATE({anio},{L}$1,1),{A},"<"&EDATE(DATE({anio},{L}$1,1),1))').number_format = NUM
    ws.cell(15, 14, '=SUM(B15:M15)').number_format = NUM
    for j in range(2, 15):
        L = CL(j); c = ws.cell(16, j, f'=ROUND({L}13-{L}15,2)'); c.number_format = NUM; c.font = B
    ws.conditional_formatting.add('B16:N16', FormulaRule(formula=['ABS(B16)>1'], fill=PatternFill('solid', fgColor='FFC7CE'), font=Font(color='9C0006', bold=True)))
    ws.cell(18, 1, 'Valores en azul: Libro IVA Digital. NC en positivo (restan en el total).').font = Font(name='Calibri', size=9, italic=True)
    ws.column_dimensions['A'].width = 44
    for j in range(2, 15):
        ws.column_dimensions[CL(j)].width = 16
    ws.freeze_panes = 'B2'


def hoja_resumen(wb, anio=None):
    """Reconstruye el resumen leyendo la hoja PORTAL IVA tal como está (sirve después de reclasificar)."""
    src = wb['PORTAL IVA']
    tot = {}
    for r in range(3, src.max_row + 1):
        f, tipo, m = src.cell(r, 1).value, src.cell(r, 2).value, src.cell(r, 13).value
        if f is None or not isinstance(m, str) or m == 'SIN CLASIFICAR':
            continue
        anio = anio or f.year
        g = lambda j: float(src.cell(r, j).value or 0)
        tot[m] = tot.get(m, 0) + (g(11) if tipo in TIPO_C else g(7) + g(8) + g(9))
    nm = 'Resumen por Clasificación'
    pos = wb.sheetnames.index(nm) if nm in wb.sheetnames else 2
    if nm in wb.sheetnames:
        del wb[nm]
    ws = wb.create_sheet(nm, pos)
    clases = list(tot)
    clases = sorted(clases, key=lambda k: -tot[k])
    P = "'PORTAL IVA'!"; L = f"{P}$L$3:$L${LAST}"; M = f"{P}$M$3:$M${LAST}"; A = f"{P}$A$3:$A${LAST}"
    F = Font(name='Calibri', size=11); B = Font(name='Calibri', size=11, bold=True)
    ws['A1'] = 'Compras por Clasificación y Mes - Neto Gravado + No Gravado + Exento'; ws['A1'].font = Font(name='Calibri', size=13, bold=True)
    ws['A2'] = 'Fuente: hoja PORTAL IVA (columna L). Las notas de crédito restan.'; ws['A2'].font = Font(name='Calibri', size=9, italic=True)
    ws.cell(4, 1, 'Clasificación')
    for m in range(12):
        ws.cell(4, 2 + m, dt.datetime(anio, m + 1, 1)).number_format = 'mmm-yy'
    ws.cell(4, 14, 'Total'); ws.cell(4, 15, '% s/Total')
    for j in range(1, 16):
        x = ws.cell(4, j); x.font = B; x.fill = HDR; x.alignment = Alignment(horizontal='center'); x.border = Border(top=THIN, bottom=THIN)
    r0 = 5; last = r0 + len(clases) - 1; T = last + 1
    for i, k in enumerate(clases):
        r = r0 + i
        ws.cell(r, 1, k).font = F
        for m in range(12):
            col = CL(2 + m)
            ws.cell(r, 2 + m, f'=SUMIFS({L},{M},$A{r},{A},">="&{col}$4,{A},"<"&EDATE({col}$4,1))').number_format = NUM
        ws.cell(r, 14, f'=SUM(B{r}:M{r})').number_format = NUM
        ws.cell(r, 15, f'=IF($N${T}=0,0,N{r}/$N${T})').number_format = '0.0%'
    ws.cell(T, 1, 'TOTAL')
    for j in range(2, 16):
        col = CL(j)
        ws.cell(T, j, f'=SUM({col}{r0}:{col}{last})').number_format = NUM if j < 15 else '0.0%'
    for j in range(1, 16):
        x = ws.cell(T, j); x.font = B; x.fill = PatternFill('solid', fgColor='DDEBF7'); x.border = Border(top=THIN, bottom=Side(style='double'))
    C = T + 2
    for k, (lab, f) in enumerate([('Control: total columna L PORTAL IVA', f'=SUM({L})'),
                                  ('Diferencia (debe ser 0)', f'=ROUND(N{T}-N{C},2)'),
                                  ('Sin clasificar ("SIN CLASIFICAR")', f'=SUMIFS({L},{M},"SIN CLASIFICAR")')]):
        ws.cell(C + k, 1, lab).font = Font(name='Calibri', size=10, italic=True)
        ws.cell(C + k, 14, f).number_format = NUM
    ws.column_dimensions['A'].width = 42
    for j in range(2, 15):
        ws.column_dimensions[CL(j)].width = 15
    ws.freeze_panes = 'B5'


def hoja_sin_clasificar(wb, c):
    nm = 'Proveedores sin clasificar'
    if nm in wb.sheetnames:
        del wb[nm]
    s = c[c.clasif.isna()].copy()
    if s.empty:
        return []
    s['imp'] = [r.tot if r.tipo in TIPO_C else r.neto + r.ng + r.ex for _, r in s.iterrows()]
    g = s.groupby('cuit').agg(nom=('nom', 'first'), n=('nom', 'size'), imp=('imp', 'sum')).sort_values('imp', ascending=False).reset_index()
    ws = wb.create_sheet(nm)
    H = ['CUIT', 'Denominación', 'Comprobantes', 'Neto + NG + Exento', 'Actividad encontrada', 'Clasificación sugerida', 'Confianza', 'Fuente', 'Estado']
    for j, h in enumerate(H, 1):
        x = ws.cell(1, j, h); x.font = Font(bold=True); x.fill = HDR; x.alignment = Alignment(wrap_text=True, vertical='center')
    for i, r in g.iterrows():
        vals = [int(r.cuit), r.nom, int(r.n), round(float(r.imp), 2), None, 'SIN CLASIFICAR', None, None, 'Pendiente']
        for j, v in enumerate(vals, 1):
            ws.cell(2 + i, j, v).alignment = Alignment(wrap_text=True, vertical='top')
        ws.cell(2 + i, 4).number_format = '#,##0.00'
        ws.cell(2 + i, 6).fill = ORANGE
    for col, w in zip('ABCDEFGHI', [14, 36, 13, 18, 55, 26, 11, 45, 20]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    return [dict(cuit=int(r.cuit), nombre=r.nom, comprobantes=int(r.n), importe=round(float(r.imp), 2)) for _, r in g.iterrows()]


# ---------------------------------------------------------------- VENTAS
def armar_ventas(periodos, tmpl, out_dir):
    rows = []
    for p, d in periodos.items():
        f = os.path.join(d, 'comprobantes_ventas.csv')
        if not os.path.exists(f):
            continue
        for _, r in leer_csv(f).iterrows():
            rows.append(dict(fecha=pd.to_datetime(r['Fecha de Emisión']), tipo=int(r['Tipo de Comprobante']),
                             pv=int(r['Punto de Venta']), nro=int(r['Número de Comprobante']),
                             cuit=int(r['Nro. Doc. Comprador']) if pd.notna(r['Nro. Doc. Comprador']) else None,
                             nom=r['Denominación Comprador'], neto=r['Total Neto Gravado'], ng=r['Importe No Gravado'],
                             ex=r['Importe Exento'], iva=r['Total IVA'], tot=r['Importe Total']))
    v = pd.DataFrame(rows)
    for k in ['neto', 'ng', 'ex', 'iva', 'tot']:
        v[k] = pd.to_numeric(v[k], errors='coerce').fillna(0).round(2)
    v = v.sort_values(['fecha', 'tipo', 'pv', 'nro']).reset_index(drop=True)

    wb = openpyxl.load_workbook(tmpl)
    ws = wb['PORTAL IVA']
    dsn = 'Comprobantes Ventas'
    if dsn in wb.sheetnames:
        del wb[dsn]
    ds = wb.create_sheet(dsn)
    ds['A1'] = 'Comprobantes de Ventas (Portal IVA)'; ds['A1'].font = Font(name='Arial', bold=True, size=11)
    hdr = ['Fecha', 'Tipo', 'Punto de Venta', 'Número Desde', 'Nro. Doc. Comprador', 'Denominación Comprador',
           'Neto Gravado', 'No Gravado', 'Exento', 'IVA', 'Total', 'Neto Gravado + No Gravado + Exento']
    for j, (h, w) in enumerate(zip(hdr, [13, 8, 15, 15, 22, 34, 16, 14, 14, 15, 16, 32]), 1):
        x = ds.cell(2, j, h); x.font = Font(name='Arial', bold=True); x.fill = HDR
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ds.column_dimensions[CL(j)].width = w
    for i, r in v.iterrows():
        n = 3 + i
        vals = [r.fecha.to_pydatetime(), r.tipo, r.pv, r.nro, r.cuit, r.nom] + [float(r[k]) for k in ['neto', 'ng', 'ex', 'iva', 'tot']]
        for j, val in enumerate(vals, 1):
            x = ds.cell(n, j, val); x.font = Font(name='Arial')
            if j == 1: x.number_format = 'dd/mm/yyyy'
            if j >= 7: x.number_format = NUM
        x = ds.cell(n, 12, f'=IF(OR(B{n}=11,B{n}=13),K{n},G{n}+H{n}+I{n})'); x.number_format = NUM; x.font = Font(name='Arial')
    last = 2 + len(v)
    ds.auto_filter.ref = f'A2:L{last}'; ds.freeze_panes = 'A3'

    Lr = f"'{dsn}'!$L$3:$L${last}"; Ar = f"'{dsn}'!$A$3:$A${last}"; Br = f"'{dsn}'!$B$3:$B${last}"
    notnc = ','.join(f'{Br},"<>{t}"' for t in NC_TIPOS)
    colmes = {}
    for j in range(2, 14):
        h = ws.cell(5, j).value
        col = CL(j)
        if isinstance(h, dt.datetime):
            colmes[(h.year, h.month)] = col
        ws[f'{col}6'] = f'=SUMIFS({Lr},{Ar},">="&{col}$5,{Ar},"<"&EDATE({col}$5,1),{notnc})'
        ws[f'{col}7'] = '=-(' + '+'.join(f'SUMIFS({Lr},{Ar},">="&{col}$5,{Ar},"<"&EDATE({col}$5,1),{Br},{t})' for t in NC_TIPOS) + ')'
        for rr in (6, 7):
            ws[f'{col}{rr}'].number_format = NUM
    ws['A6'].comment = Comment('Facturas y ND: Neto + No Gravado + Exento, desde la hoja Comprobantes Ventas.', 'Claude')
    ws['A7'].comment = Comment('Notas de Crédito (tipos 3, 8, 13) en positivo, desde la hoja Comprobantes Ventas.', 'Claude')
    blue = Font(name='Arial', color='0000FF')
    src = {14: ['totales_ventas_ri/total_neto_gravado', 'totales_ventas_cons_fin_exento_no_alcanzado/total_neto_gravado', 'totales_ventas_otros_conceptos_deb_fiscal/total_neto_gravado'],
           15: ['totales_ventas_ri_otros_rest_deb_fiscal/total_neto_gravado', 'totales_ventas_exentos_no_alcanzado_mono_cons_fin_rest_deb_fiscal/total_neto_gravado'],
           19: ['total_ventas_importes_exentos_no_gravados_excepto_exportaciones', 'total_ventas_exportacion'],
           20: ['total_ventas_importes_exentos_no_gravados_nc_excepto_exportaciones']}
    faltan = []
    for p, d in periodos.items():
        col = colmes.get((int(p[:4]), int(p[4:])))
        if not col:
            faltan.append(p); continue
        x = xml_consol(d, 'iva_ventas')
        for rr, qs in src.items():
            c = ws[f'{col}{rr}']; c.value = xval(x, qs); c.number_format = NUM; c.font = blue
    for rr, qs in src.items():
        ws[f'A{rr}'].comment = Comment('libro_iva.xml > iva_ventas: ' + ' + '.join(qs), 'Claude')
    out = os.path.join(out_dir, os.path.basename(tmpl))
    wb.save(out)
    return out, v, faltan


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--zips', nargs='+', required=True)
    ap.add_argument('--compras', required=True)
    ap.add_argument('--ventas', required=True)
    ap.add_argument('--padron', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--work', default='/home/claude/portal_iva_work')
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    periodos = extraer(a.zips, a.work)
    pad_out, clasif, conflictos, cambios_pad, otras = cargar_padron(a.padron, a.out)
    cout, c, sc = armar_compras(periodos, a.compras, clasif, a.out)
    vout, v, faltan = armar_ventas(periodos, a.ventas, a.out)
    usados = {k: m for k, m in conflictos.items() if k in set(c.cuit)}
    res = dict(periodos=list(periodos), compras_filas=len(c), ventas_filas=len(v),
               despachos_impo=int((c.pv.isna()).sum()),
               padron_variantes_unificadas=cambios_pad, padron_otras_variantes_posibles=otras,
               cuit_con_varias_clasificaciones_usados={str(k): m for k, m in usados.items()},
               sin_clasificar=sc, ventas_periodos_sin_columna=faltan,
               archivos=[cout, vout, pad_out])
    with open(os.path.join(a.work, 'resumen.json'), 'w', encoding='utf-8') as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
