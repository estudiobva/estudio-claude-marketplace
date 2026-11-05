"""Gastos Varios (GAVA) de T&S Web -> Excel con Detalle, Resumen y Cruce con el EERR.

    python3 gastos_varios_a_excel.py buscar-eerr --cuit=30... --eerr-dir=<EERR 2026> [--hoy=AAAA-MM-DD]
    python3 gastos_varios_a_excel.py armar <reporte.csv> <carpeta> --cuit=30...
            --desde=DD/MM/AAAA --hasta=DD/MM/AAAA [--eerr=<EERR.xlsx>] [--entidad=..] [--razon=..]

buscar-eerr: busca el EERR de la sociedad (CUIT en la celda E3 de la hoja EERR)
e imprime {"eerr", "desde", "hasta"} del ejercicio que contiene el mes anterior
(o el mas nuevo). Sin EERR imprime {"eerr": null}.

armar: deja en <carpeta> un AAAAMM.xlsx por cada mes con valores (Resumen,
Detalle y Cruce EERR del mes). Lee el CSV del reporte 4.5.2 Seguimiento de Valores (UTF-8, ";"):
    cabecera   "T&S Web", fecha, "Consulta de valores", "Fecha desde / hasta:"
               (col H y L), total general (col O), "Venc. desde / hasta:"
    concepto   ";<NOMBRE>;...;<$ total en col O>"
    titulos    ";Codificación;Tipo Valor;Ent. Sec.;;Monto;;;Fecha Em.;;Fecha Venc.;;Nodo;;;Estado;Transmision"
    detalle    "<cod>;;<tipo>;<ent sec>;;<$ monto>;;;<fecha em>;;<fecha venc>;;<nodo>;;;<estado>;"
    pie        "Página 1 de ;...;1;..."
Controles (sobre el reporte completo, antes de partirlo): suma del detalle de cada concepto = total del concepto, suma de los
totales = total general, fechas de emision dentro del rango pedido. Si algo no
cierra sale con 1 y {"error"}.

Reglas de la habilidad: IGNORADO fuera del total (se listan aparte), INGRESADO
adentro marcado como pendiente de pago, mes = mes de la Fecha Em. (aviso si el
vencimiento cae en otro mes). El cruce compara mes a mes los conceptos
mapeados contra los valores calculados del EERR (no escribe en el EERR).
"""
import glob
import json
import os
import re
import sys
import unicodedata
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TOL = 0.01
MONEDA = '"$" #,##0.00;[Red]-"$" #,##0.00'
AMARILLO = PatternFill('solid', fgColor='FFF2CC')
ROJO = PatternFill('solid', fgColor='F8CBAD')
GRIS = PatternFill('solid', fgColor='D9D9D9')
NEGRITA = Font(bold=True)

# Concepto de T&S (regex sobre el nombre, que Zetti corta a ~37 caracteres)
# -> fila del EERR (rotulo para mostrar, regex sobre la columna A del EERR,
# si el mapeo esta confirmado).
MAPEO = [
    (r'COLEGIO', 'Matricula Colegio Farmaceuticos', r'^MATRICULA COLEGIO', True),
    (r'CONSEJO PROFESIONAL', 'Certificacion Balances', r'CERTIFICACION', True),
    (r'TASA PUBLICIDAD', 'Publicidad', r'^PUBLICIDAD', True),
    (r'TASA SEGURIDAD', 'Seguridad e Higiene', r'^SEGURIDAD E HIGIENE', True),
    (r'^MUNICIPAL', 'Municipalidad (ABL)', r'MUNICIPALIDAD|\bABL\b', True),
    (r'ADEF', 'Sueldos y Cargas Sociales / ADEF', r'^SUELDOS Y CARGAS', False),
    (r'JUICIOS? LABORAL', 'Indemnizacion por Despido / juicios', r'^INDEMNIZACION', False),
    (r'METROGAS', 'Metrogas', r'^METROGAS', True),
    (r'EDESUR', 'Edesur', r'^EDESUR', True),
    (r'AYSA|\bAGUA\b', 'Agua', r'^AGUA', True),
    (r'ALQUILER', 'ALQUILERES', r'^ALQUILERES', True),
    (r'CAFAR', 'Cafar', r'CAFAR', True),
]


def fallar(msg, **extra):
    print(json.dumps({'error': msg, **extra}, ensure_ascii=False))
    sys.exit(1)


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().upper()


def importe(s):
    s = str(s or '').replace('$', '').replace('\xa0', '').strip()
    if not s:
        return None
    neg = s.startswith('-') or (s.startswith('(') and s.endswith(')'))
    s = s.strip('-() ').replace('.', '').replace(',', '.')
    v = float(s)
    return -v if neg else v


def fecha(s):
    s = str(s or '').strip()
    return datetime.strptime(s, '%d/%m/%Y').date() if re.match(r'^\d{2}/\d{2}/\d{4}$', s) else None


def args_kv(argv):
    kv, pos = {}, []
    for a in argv:
        if a.startswith('--'):
            k, _, v = a[2:].partition('=')
            kv[k] = v
        else:
            pos.append(a)
    return kv, pos


# ---------------------------------------------------------------- EERR

RE_EERR = re.compile(r'^EERR (\d{2})-(\d{4}) a (\d{2})-(\d{4}) - (.+)\.xlsx$')


def cuit_de_eerr(path):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb['EERR']
            for fila in ws.iter_rows(min_row=2, max_row=4, min_col=4, max_col=6, values_only=True):
                for v in fila:
                    d = re.sub(r'\D', '', str(v or ''))
                    if len(d) == 11:
                        return d
        finally:
            wb.close()
    except Exception:
        return None
    return None


def buscar_eerr(cuit, eerr_dir, hoy=None, nombre=None):
    hoy = hoy or date.today()
    ref = date(hoy.year, hoy.month, 1)
    ref = date(ref.year - 1, 12, 1) if ref.month == 1 else date(ref.year, ref.month - 1, 1)
    archivos = []
    for p in glob.glob(os.path.join(eerr_dir, 'EERR *.xlsx')):
        m = RE_EERR.match(unicodedata.normalize('NFC', os.path.basename(p)))
        if m:
            archivos.append((p, m))
    if not archivos:
        return None
    # Primero los que se llaman como la sociedad (ahorra abrir los 70 archivos).
    nom = norm(nombre).replace(' ', '') if nombre else ''
    primero = [a for a in archivos if nom and norm(a[1].group(5)).replace(' ', '') == nom]
    resto = [a for a in archivos if a not in primero]
    hits = [a for a in primero if cuit_de_eerr(a[0]) == cuit]
    if not hits:
        hits = [a for a in resto if cuit_de_eerr(a[0]) == cuit]
    if not hits:
        return None

    def rango(m):
        d = date(int(m.group(2)), int(m.group(1)), 1)
        h = date(int(m.group(4)), int(m.group(3)), 1)
        return d, h

    def ultimo(d):
        n = date(d.year + 1, 1, 1) if d.month == 12 else date(d.year, d.month + 1, 1)
        return date.fromordinal(n.toordinal() - 1)

    vigente = [a for a in hits if rango(a[1])[0] <= ref <= rango(a[1])[1]]
    elegido = vigente[0] if vigente else max(hits, key=lambda a: rango(a[1])[0])
    d, h = rango(elegido[1])
    return {'eerr': elegido[0], 'desde': d.strftime('%d/%m/%Y'), 'hasta': ultimo(h).strftime('%d/%m/%Y')}


def leer_eerr(path):
    """{rotulo_normalizado: (rotulo, {(anio, mes): valor})} con los valores calculados."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb['EERR']
        filas = list(ws.iter_rows(min_row=1, max_row=80, max_col=14, values_only=True))
    finally:
        wb.close()
    meses = {}
    for fila in filas:
        if fila and norm(fila[0]) == 'CONCEPTO':
            for j, v in enumerate(fila[1:13], start=1):
                if isinstance(v, datetime):
                    meses[j] = (v.year, v.month)
            break
    if not meses:
        raise ValueError('el EERR no tiene la fila "Concepto" con los meses')
    out = {}
    sin_calcular = False
    for fila in filas:
        if not fila or not fila[0] or norm(fila[0]) == 'CONCEPTO':
            continue
        vals = {}
        for j, ym in meses.items():
            v = fila[j] if j < len(fila) else None
            if isinstance(v, (int, float)):
                vals[ym] = float(v)
            elif isinstance(v, str) and v.startswith('='):
                sin_calcular = True
        out[norm(fila[0])] = (str(fila[0]).strip(), vals)
    return out, sorted(meses.values()), sin_calcular


# ---------------------------------------------------------------- CSV

def parsear(csv_path):
    with open(csv_path, encoding='utf-8-sig') as fh:
        lineas = [l.rstrip('\r\n') for l in fh]
    if any('CodificaciÃ' in l for l in lineas):
        fallar('El CSV se leyo con la codificacion equivocada (CodificaciÃ³n).')
    cab = {'desde': None, 'hasta': None, 'total_general': None}
    conceptos, actual, raras = [], None, []
    estado = 'cabecera'
    for n, l in enumerate(lineas, start=1):
        c = l.split(';')
        c += [''] * (17 - len(c))
        if not l.strip(';').strip():
            continue
        if c[4].startswith('Fecha desde'):
            cab['desde'], cab['hasta'] = c[7].strip(), c[11].strip()
            continue
        if c[4].startswith('Venc.'):
            continue
        if estado == 'cabecera':
            if c[0] in ('T&S Web', 'Consulta de valores') or c[4].startswith('Venc.') or re.search(r'\d{4}$', c[13]):
                continue
            if not c[0] and not c[1] and c[14].strip().startswith('$'):
                # Con totales de $ 1.000 M o mas T&S deja el total general en
                # blanco ("$ "): se controla igual concepto por concepto.
                cab['total_general'] = importe(c[14])
                estado = 'cuerpo'
                continue
        if c[0].startswith('Página') or c[0].startswith('Pagina'):
            continue
        if not c[0] and c[1].startswith('Codificaci'):
            continue
        if not c[0] and c[1].strip() and c[14].strip().startswith('$'):
            actual = {'concepto': c[1].strip(), 'total_reporte': importe(c[14]), 'items': []}
            conceptos.append(actual)
            continue
        # Detalle: codificacion (casi siempre numerica, a veces con letras:
        # "6547U65"), monto con $ y fecha de emision.
        if c[0].strip() and c[5].strip().startswith('$') and fecha(c[8]) and actual is not None:
            cod = c[0].strip()
            actual['items'].append({
                'codificacion': int(cod) if cod.isdigit() else cod, 'tipo_valor': c[2].strip(), 'ent_sec': c[3].strip(),
                'monto': importe(c[5]), 'fecha_em': fecha(c[8]), 'fecha_venc': fecha(c[10]),
                'nodo': c[12].strip(), 'estado': c[15].strip(), 'transmision': c[16].strip(),
            })
            continue
        raras.append({'linea': n, 'texto': l[:120]})
    if raras:
        fallar(f'{len(raras)} fila/s del reporte que no reconozco', filas=raras[:10])
    if estado != 'cuerpo' or not cab['desde']:
        fallar('El CSV no trae la cabecera esperada (fechas y total general)')
    return cab, conceptos


# ---------------------------------------------------------------- armado

def mapeo_de(concepto):
    for rx, rotulo, rx_eerr, confirmado in MAPEO:
        if re.search(rx, norm(concepto)):
            return rotulo, rx_eerr, confirmado
    return None


def autoancho(ws, minimo=8, maximo=60):
    anchos = {}
    for fila in ws.iter_rows():
        for c in fila:
            if c.value is None:
                continue
            v = c.value
            largo = len(f'{v:,.2f}') + 3 if isinstance(v, float) else len(str(v))
            anchos[c.column_letter] = max(anchos.get(c.column_letter, 0), largo)
    for k, v in anchos.items():
        ws.column_dimensions[k].width = max(minimo, min(maximo, v + 2))


def armar(csv_path, dir_salida, kv):
    """Controla el reporte completo y deja un libro por mes con valores en
    dir_salida/AAAAMM.xlsx. Imprime el JSON con el resumen y la lista de meses."""
    cab, conceptos = parsear(csv_path)
    desde_p, hasta_p = kv.get('desde'), kv.get('hasta')
    if desde_p and (cab['desde'] != desde_p or cab['hasta'] != hasta_p):
        fallar(f"El reporte salio del {cab['desde']} al {cab['hasta']} y se pidio del {desde_p} al {hasta_p}")
    d0 = datetime.strptime(cab['desde'], '%d/%m/%Y').date()
    d1 = datetime.strptime(cab['hasta'], '%d/%m/%Y').date()

    # Controles del reporte completo (antes de partirlo por mes).
    dif_conceptos = []
    for co in conceptos:
        s = round(sum(i['monto'] for i in co['items']), 2)
        co['suma_detalle'] = s
        if abs(s - co['total_reporte']) > TOL:
            dif_conceptos.append({'concepto': co['concepto'], 'total_reporte': co['total_reporte'], 'suma_detalle': s})
    suma_conceptos = round(sum(co['total_reporte'] for co in conceptos), 2)
    aviso_total = None
    if cab['total_general'] is None:
        aviso_total = f'T&S no imprimio el total general (pasa con mas de $ 1.000 M): se usa la suma de los conceptos, {suma_conceptos:,.2f}'
        cab['total_general'] = suma_conceptos
    elif abs(suma_conceptos - cab['total_general']) > TOL:
        fallar(f"La suma de los conceptos ({suma_conceptos:,.2f}) no da el total general del reporte ({cab['total_general']:,.2f})")
    if dif_conceptos:
        fallar('El detalle de algun concepto no suma su total', diferencias=dif_conceptos)
    fuera = [i for co in conceptos for i in co['items'] if not i['fecha_em'] or not (d0 <= i['fecha_em'] <= d1)]
    if fuera:
        fallar(f'{len(fuera)} valor/es con fecha de emision fuera del rango pedido',
               ejemplos=[{'cod': i['codificacion'], 'fecha_em': str(i['fecha_em'])} for i in fuera[:5]])

    filas = [{**i, 'concepto': co['concepto']} for co in conceptos for i in co['items']]
    for f in filas:
        f['ym'] = (f['fecha_em'].year, f['fecha_em'].month)
        f['ignorado'] = f['estado'].upper() == 'IGNORADO'
        f['ingresado'] = f['estado'].upper() == 'INGRESADO'
        f['venc_otro_mes'] = bool(f['fecha_venc']) and (f['fecha_venc'].year, f['fecha_venc'].month) != f['ym']

    meses = []
    y, m = d0.year, d0.month
    while (y, m) <= (d1.year, d1.month):
        meses.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)

    # EERR (una sola lectura para todos los meses).
    eerr_path, eerr, meses_eerr, aviso_eerr = kv.get('eerr') or None, {}, [], None
    if eerr_path:
        try:
            eerr, meses_eerr, sin_calcular = leer_eerr(eerr_path)
            if sin_calcular:
                aviso_eerr = 'El EERR tiene formulas sin valor calculado (abrirlo y guardarlo en Excel): esas celdas se toman como vacias.'
        except Exception as e:  # noqa: BLE001
            aviso_eerr = f'No pude leer el EERR: {e}'

    os.makedirs(dir_salida, exist_ok=True)
    salida_meses, cruce_total = [], []
    for ym in meses:
        del_mes = [f for f in filas if f['ym'] == ym]
        if not del_mes:
            continue
        archivo = os.path.join(dir_salida, f'{ym[0]}{ym[1]:02d}.xlsx')
        res = libro_mes(archivo, ym, del_mes, filas, cab, kv, eerr_path, eerr, meses_eerr, aviso_eerr, aviso_total)
        salida_meses.append(res)
        cruce_total += res.pop('_cruce')

    validos = [f for f in filas if not f['ignorado']]
    tot = lambda xs: round(sum(x['monto'] for x in xs), 2)  # noqa: E731
    mapeados = {}
    for f in validos:
        mp = mapeo_de(f['concepto'])
        if mp:
            k = mp[0] + ('' if mp[2] else ' (a confirmar)')
            mapeados[k] = round(mapeados.get(k, 0) + f['monto'], 2)
    print(json.dumps({
        'desde': cab['desde'], 'hasta': cab['hasta'], 'total_reporte': cab['total_general'],
        'total_sin_ignorados': tot(validos), **({'aviso_total': aviso_total} if aviso_total else {}),
        'valores': len(filas), 'conceptos': len(conceptos),
        'ignorados': {'cantidad': sum(f['ignorado'] for f in filas), 'total': tot([f for f in filas if f['ignorado']])},
        'ingresados': {'cantidad': sum(f['ingresado'] for f in filas), 'total': tot([f for f in filas if f['ingresado']])},
        'vencen_otro_mes': sum(f['venc_otro_mes'] for f in validos),
        'mapeados': mapeados,
        'no_mapeados': sorted({f['concepto'] for f in validos if not mapeo_de(f['concepto'])}),
        'meses': salida_meses,
        'meses_sin_valores': [f'{mm:02d}/{yy}' for yy, mm in meses if not any(f['ym'] == (yy, mm) for f in filas)],
        'eerr': eerr_path, 'aviso_eerr': aviso_eerr,
        'cruce_con_diferencia': [c for c in cruce_total if abs(c['dif']) > 1],
    }, ensure_ascii=False, default=str))


def libro_mes(archivo, ym, del_mes, todas, cab, kv, eerr_path, eerr, meses_eerr, aviso_eerr, aviso_total):
    etiqueta = f'{ym[1]:02d}/{ym[0]}'
    validos = [f for f in del_mes if not f['ignorado']]
    ignorados = [f for f in del_mes if f['ignorado']]
    wb = Workbook()
    titulo = f"{kv.get('razon') or ''}  -  CUIT {kv.get('cuit', '')}  -  T&S 4.5.2 GAVA Gastos Varios {etiqueta}".strip(' -')

    # --- Resumen del mes por concepto
    por_concepto = {}
    for f in del_mes:
        r = por_concepto.setdefault(f['concepto'], {'total': 0.0, 'ign': 0.0, 'ing': 0.0, 'n': 0})
        r['n'] += 1
        if f['ignorado']:
            r['ign'] += f['monto']
        else:
            r['total'] += f['monto']
        if f['ingresado']:
            r['ing'] += f['monto']
    ws = wb.active
    ws.title = 'Resumen'
    ws['A1'] = titulo
    ws['A1'].font = NEGRITA
    ws['A2'] = (f"Entidad T&S: {kv.get('entidad', '')}. Fecha de emision en {etiqueta}. Total sin IGNORADO; "
                'INGRESADO incluido (pendiente de pago). En amarillo, los conceptos que van al EERR.')
    ws['A3'] = (f"Reporte bajado del {cab['desde']} al {cab['hasta']}: detalle = total de cada concepto"
                + (' (T&S no imprimio el total general).' if aviso_total else ' y total general.'))
    ws.append([])
    ws.append(['Concepto', 'Fila EERR', 'Valores', 'Total', 'INGRESADO (pendiente)', 'IGNORADO (fuera)'])
    for c in ws[ws.max_row]:
        c.font = NEGRITA
        c.fill = GRIS
    ini = ws.max_row + 1
    for co in sorted(por_concepto):
        r = por_concepto[co]
        mp = mapeo_de(co)
        destino = (mp[0] + ('' if mp[2] else ' (a confirmar)')) if mp else 'no mapeado'
        ws.append([co, destino, r['n'], round(r['total'], 2), round(r['ing'], 2) or None, round(r['ign'], 2) or None])
        if mp:
            for c in ws[ws.max_row]:
                c.fill = AMARILLO
    fin = ws.max_row
    ws.append(['TOTAL', None, f'=SUM(C{ini}:C{fin})', f'=SUM(D{ini}:D{fin})', f'=SUM(E{ini}:E{fin})', f'=SUM(F{ini}:F{fin})'])
    for c in ws[ws.max_row]:
        c.font = NEGRITA
    for fila in ws.iter_rows(min_row=ini, min_col=4, max_col=6):
        for c in fila:
            c.number_format = MONEDA
    autoancho(ws)
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 36

    # --- Detalle
    wd = wb.create_sheet('Detalle')
    wd.append(['Concepto', 'Codificacion', 'Tipo Valor', 'Ent. Sec.', 'Monto', 'Fecha Em.', 'Fecha Venc.', 'Nodo', 'Estado', 'Observacion'])
    for c in wd[1]:
        c.font = NEGRITA
        c.fill = GRIS
    for f in del_mes:
        obs = []
        if f['ignorado']:
            obs.append('IGNORADO: fuera del total')
            par = [g['codificacion'] for g in todas if not g['ignorado'] and g['concepto'] == f['concepto'] and abs(g['monto'] - f['monto']) <= TOL]
            if par:
                obs.append('mismo monto PAGADO: ' + ', '.join(map(str, par)))
        if f['ingresado']:
            obs.append('pendiente de pago')
        if f['venc_otro_mes']:
            obs.append('vence en otro mes')
        wd.append([f['concepto'], f['codificacion'], f['tipo_valor'], f['ent_sec'], f['monto'], f['fecha_em'], f['fecha_venc'],
                   f['nodo'], f['estado'], '; '.join(obs)])
        if f['ignorado']:
            for c in wd[wd.max_row]:
                c.font = Font(color='808080', strike=True)
    for fila in wd.iter_rows(min_row=2):
        fila[4].number_format = MONEDA
        fila[5].number_format = 'DD/MM/YYYY'
        fila[6].number_format = 'DD/MM/YYYY'
    wd.auto_filter.ref = f'A1:J{wd.max_row}'
    wd.freeze_panes = 'A2'
    autoancho(wd)

    # --- Cruce con el EERR del mes
    cruce = []
    if eerr_path:
        wx = wb.create_sheet('Cruce EERR')
        wx['A1'] = f'Cruce T&S GAVA {etiqueta} vs EERR: {os.path.basename(eerr_path)}'
        wx['A1'].font = NEGRITA
        wx['A2'] = 'T&S = suma del mes sin IGNORADO. EERR = valor calculado de la fila. Solo lectura: no se escribio nada en el EERR.'
        if aviso_eerr:
            wx['A3'] = aviso_eerr
            wx['A3'].font = Font(color='C00000', bold=True)
        wx.append([])
        wx.append(['Fila EERR', 'Conceptos T&S', 'T&S', 'EERR', 'Diferencia (T&S - EERR)', 'Observacion'])
        for c in wx[wx.max_row]:
            c.font = NEGRITA
            c.fill = GRIS
        grupos = {}
        for f in validos:
            mp = mapeo_de(f['concepto'])
            if mp:
                g = grupos.setdefault(mp[0], {'rx': mp[1], 'confirmado': mp[2], 'conceptos': set(), 'ts': 0.0})
                g['conceptos'].add(f['concepto'])
                g['ts'] += f['monto']
        en_eerr = ym in set(meses_eerr)
        ant = (ym[0] - 1, 12) if ym[1] == 1 else (ym[0], ym[1] - 1)
        for rotulo, g in grupos.items():
            fila_eerr = next((v for k, v in eerr.items() if re.search(g['rx'], k)), None) if g['confirmado'] else None
            obs = []
            if not g['confirmado']:
                obs.append('mapeo a confirmar: no se compara')
            elif not fila_eerr:
                obs.append('la fila no existe en el EERR')
            elif not en_eerr:
                obs.append('el mes no esta en el EERR')
            ts = round(g['ts'], 2)
            ee = fila_eerr[1].get(ym) if (fila_eerr and en_eerr) else None
            dif = round(ts - (ee or 0), 2) if (g['confirmado'] and en_eerr) else None
            if fila_eerr and dif and abs(dif) > 1 and abs(ts - (fila_eerr[1].get(ant) or 0)) <= 1:
                obs.append(f'T&S = EERR de {ant[1]:02d}/{ant[0]} (se paga el mes siguiente)')
            nombre = fila_eerr[0] if fila_eerr else rotulo
            wx.append([nombre, ', '.join(sorted(g['conceptos'])), ts, ee, dif, '; '.join(obs)])
            if dif is not None and abs(dif) > 1:
                wx[wx.max_row][4].fill = ROJO
                cruce.append({'fila': nombre, 'mes': etiqueta, 'tys': ts, 'eerr': ee, 'dif': dif, **({'obs': '; '.join(obs)} if obs else {})})
        for fila in wx.iter_rows(min_row=6, min_col=3, max_col=5):
            for c in fila:
                c.number_format = MONEDA
        autoancho(wx)
        wx.column_dimensions['A'].width = 38
        wx.column_dimensions['B'].width = 45

    wb.save(archivo)
    return {
        'periodo': f'{ym[0]}{ym[1]:02d}', 'archivo_tmp': archivo, 'valores': len(del_mes),
        'total': round(sum(f['monto'] for f in validos), 2),
        'ingresados': sum(f['ingresado'] for f in del_mes), 'ignorados': len(ignorados),
        '_cruce': cruce,
    }


def main():
    if len(sys.argv) < 2:
        fallar(__doc__)
    modo = sys.argv[1]
    kv, pos = args_kv(sys.argv[2:])
    if modo == 'buscar-eerr':
        hoy = datetime.strptime(kv['hoy'], '%Y-%m-%d').date() if kv.get('hoy') else None
        res = buscar_eerr(re.sub(r'\D', '', kv.get('cuit', '')), kv['eerr-dir'], hoy, kv.get('nombre'))
        print(json.dumps(res or {'eerr': None}, ensure_ascii=False))
    elif modo == 'armar':
        if len(pos) != 2:
            fallar('armar <reporte.csv> <carpeta de salida> ...')
        armar(pos[0], pos[1], kv)
    else:
        fallar(f'modo desconocido: {modo}')


if __name__ == '__main__':
    main()
