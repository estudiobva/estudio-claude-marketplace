#!/usr/bin/env python3
# conciliar-ncr.py — concilia las compras del Portal IVA de una sociedad contra
# las Notas de Credito de Recupero (NCR) de Suizo y Monroe que dejo
# descargar-ncr.js en la carpeta ncr/ del periodo, y deja la salida que despues
# lee eliminar-ncr.js ([AAAAMM]-COMPRAS-SIN-NCR-[SOCIEDAD].xlsx).
#
#   python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=08/2026
#   python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=08/2026 --solo-revisar
#   python scripts/conciliar-ncr.py --cuit=<CUIT> --periodo=08/2026 --si-existe=reemplazar
#
# Argumentos:
#   --cuit          CUIT de la sociedad (busca su carpeta en la unidad compartida).
#   --periodo       MM/AAAA o AAAAMM.
#   --portal        Archivo de compras del Portal IVA (.csv, .zip o .xlsx). Default:
#                   lo busca en ncr/, despues en la carpeta IVA y despues en la del mes.
#   --clientes      Valores de CLIENTE de Suizo, separados por coma. Default: salen
#                   de los nombres de los archivos de Monroe (+ lib/ncr-alias.json).
#   --sin-monroe    La sociedad no tiene archivos de Monroe este mes (exige --clientes).
#   --sociedad      Nombre para el archivo de salida. Default: el de la carpeta.
#   --si-existe     error (default) | reemplazar (el anterior va a ncr/_anteriores/) | saltear
#   --solo-revisar  Concilia y reporta, pero no escribe nada.
#
# Salida: JSON por stdout (logs por stderr). Si algo no cierra imprime
# {"error": ...} y termina con codigo 1, igual que los scripts de Node.
#
# Reglas: las de la skill conciliacion-ncr (validadas contra las sociedades
# conciliadas en 08/2026). Criterio conservador: se elimina solo lo que se pudo PROBAR. Una NCR
# que matchea pero con otro importe o fecha, o contra mas de una fila del
# Portal, va a la hoja "Diferencias" y se queda en compras.
#
# Solo usa openpyxl (sin pandas) para que corra en cualquier maquina del estudio.

import csv, datetime, io, json, os, re, shutil, sys, tempfile, unicodedata, zipfile
from calendar import monthrange

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print(json.dumps({'error': 'Falta openpyxl: pip install openpyxl'}))
    sys.exit(1)

CUIT_MONROE, CUIT_SUIZO, TIPO_NC_A = '30517059095', '30516968431', 3
DROGUERIAS = {CUIT_MONROE: 'MONROE', CUIT_SUIZO: 'SUIZO'}
RAIZ_DEFECTO = r'I:\Unidades compartidas\BVA - Sociedades Farmaceuticas'
ALIAS_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'lib', 'ncr-alias.json')

# Columnas del Portal IVA que NO son numericas (el resto si).
TEXTO = {'nro. doc. vendedor', 'denominacion vendedor', 'moneda original'}
ENTEROS = {'tipo de comprobante', 'punto de venta', 'numero de comprobante', 'tipo doc. vendedor'}
FECHA = 'fecha de emision'


def log(m):
    sys.stderr.write(f'[conciliar-ncr] {m}\n')


def fail(msg, **extra):
    print(json.dumps({'error': msg, **extra}, ensure_ascii=False, default=str))
    sys.exit(1)


def args():
    out = {}
    for a in sys.argv[1:]:
        k, _, v = a.lstrip('-').partition('=')
        out[k] = v
    return out


def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().upper()


def digitos(s):
    return re.sub(r'\D', '', str(s or ''))


def periodo_aaaamm(p):
    s = str(p or '').strip()
    m = re.match(r'^(\d{1,2})[/\-.](\d{4})$', s)
    if m:
        return m.group(2) + m.group(1).zfill(2)
    if re.match(r'^\d{6}$', s):
        return s
    return None


# ── carpetas (misma logica que lib/rutas-bva.js) ────────────────────────────
def raiz():
    r = (os.environ.get('BVA_UNIDAD_PATH') or '').strip() or RAIZ_DEFECTO
    if not os.path.isdir(r):
        fail(f'No encuentro la unidad compartida en "{r}". Defini BVA_UNIDAD_PATH.')
    return r


def carpeta_sociedad(cuit):
    base = raiz()
    hits = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d)) and cuit in digitos(d)]
    if not hits:
        fail(f'No hay ninguna carpeta en la unidad compartida cuyo nombre contenga el CUIT {cuit}.')
    if len(hits) > 1:
        fail(f'El CUIT {cuit} matchea {len(hits)} carpetas: {" | ".join(hits)}. Resolvelo a mano.')
    return os.path.join(base, hits[0])


def carpeta_ncr(soc, per):
    """Misma busqueda que carpetaNcr() de lib/rutas-bva.js."""
    mens = os.path.join(soc, '01-Impuestos Mensuales')
    cand = [os.path.join(mens, per, f'{per}-IVA', 'ncr'), os.path.join(mens, f'{per}-IVA', 'ncr'),
            os.path.join(mens, per, 'IVA', 'ncr'), os.path.join(mens, per, 'ncr')]
    hay = [d for d in cand if os.path.isdir(d)]
    if len(hay) > 1:
        fail(f'Hay {len(hay)} carpetas ncr/ para {per}: {" | ".join(hay)}. Deja una sola.')
    if not hay:
        fail(f'No hay carpeta ncr/ del periodo {per} en {mens}. Primero hay que correr descargar-ncr.')
    return hay[0]


def carpetas_compras(soc, per, ncr):
    """Donde buscar las compras del Portal IVA, de la mas especifica a la mas general."""
    mens = os.path.join(soc, '01-Impuestos Mensuales')
    out = [ncr, os.path.dirname(ncr)]
    for d in (os.path.join(mens, per, f'{per}-IVA'), os.path.join(mens, f'{per}-IVA'),
              os.path.join(mens, per, 'IVA'), os.path.join(mens, per)):
        if os.path.isdir(d) and d not in out:
            out.append(d)
    return out


def nombre_sociedad(soc):
    n = os.path.basename(soc)
    n = re.sub(r'\d{11}', '', n)            # CUIT
    n = re.sub(r'^\s*\d+\s*-\s*', '', n)    # "9 - "
    return norm(n.strip(' -'))


# ── Portal IVA ──────────────────────────────────────────────────────────────
def numero_ar(s):
    """'1.234,56' -> 1234.56 (CSV de ARCA). '' -> None."""
    s = str(s).strip()
    if s == '':
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    return float(s)


def fecha(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v or '').strip()
    for f in ('%Y-%m-%d', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%m-%d-%y', '%d-%m-%Y'):
        try:
            return datetime.datetime.strptime(s[:19] if ' ' in s else s, f).date()
        except ValueError:
            pass
    return None


def tipar_portal(cab, filas):
    """Devuelve filas como dicts con tipos reales, en el orden de columnas del archivo."""
    claves = [norm(c).lower() for c in cab]
    falt = [c for c in (FECHA, 'tipo de comprobante', 'punto de venta', 'numero de comprobante',
                        'nro. doc. vendedor', 'importe total') if c not in claves]
    if falt:
        fail(f'El archivo de compras no tiene las columnas {falt}: no parece un export del Portal IVA.')
    out = []
    for n, f in enumerate(filas, start=2):
        if not any(v not in (None, '') for v in f):
            continue
        r = {}
        for c, k, v in zip(cab, claves, f):
            try:
                if k == FECHA:
                    r[c] = fecha(v)
                elif k in TEXTO:
                    r[c] = '' if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip())
                elif isinstance(v, (int, float)):
                    r[c] = int(v) if k in ENTEROS else float(v)
                else:
                    x = numero_ar(v) if v is not None else None
                    r[c] = (int(x) if k in ENTEROS and x is not None else x)
            except ValueError:
                fail(f'Compras, fila {n}: "{c}" no es numerico ({v!r}).')
        out.append(r)
    return out


def leer_csv(texto):
    filas = list(csv.reader(io.StringIO(texto), delimiter=';'))
    return filas[0], filas[1:]


def decodificar(b):
    for enc in ('utf-8-sig', 'ISO-8859-1'):     # ARCA manda ISO-8859-1; se prueba UTF-8 antes
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass


def leer_portal(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == '.csv':
        cab, filas = leer_csv(decodificar(open(ruta, 'rb').read()))
    elif ext == '.zip':
        with zipfile.ZipFile(ruta) as z:
            csvs = [n for n in z.namelist() if n.lower().endswith('.csv')]
            if len(csvs) != 1:
                fail(f'El zip {os.path.basename(ruta)} tiene {len(csvs)} CSV (se espera uno).')
            cab, filas = leer_csv(decodificar(z.read(csvs[0])))
    elif ext == '.xlsx':
        wb = load_workbook(ruta, read_only=True, data_only=True)
        for ws in wb.worksheets:
            filas = list(ws.iter_rows(values_only=True))
            if filas and FECHA in [norm(c).lower() for c in filas[0] if c is not None]:
                cab, filas = [str(c).strip() if c is not None else '' for c in filas[0]], filas[1:]
                break
        else:
            fail(f'{os.path.basename(ruta)} no tiene ninguna hoja con el formato del Portal IVA.')
    else:
        fail(f'No se leer compras en formato {ext}.')
    return cab, tipar_portal(cab, filas)


def es_compras(nombre, per):
    n = norm(nombre)
    if nombre.startswith('~$') or 'SIN-NCR' in n or 'SIN NCR' in n or 'ELIMINACION' in n:
        return False
    ext = os.path.splitext(nombre)[1].lower()
    if ext in ('.csv', '.zip') and re.search(rf'COMPROBANTES_PERIODO_{per}_COMPRAS', n):
        return True
    return ext in ('.csv', '.zip', '.xlsx') and 'COMPRAS' in n and 'PORTAL' in n.replace(' ', '')


def ubicar_portal(soc, per, ncr):
    for d in carpetas_compras(soc, per, ncr):
        hits = sorted(f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f)) and es_compras(f, per))
        if len(hits) == 1:
            return os.path.join(d, hits[0])
        if len(hits) > 1:
            fail(f'Hay {len(hits)} archivos de compras en {d}: {" | ".join(hits)}. Indica cual con --portal.')
    fail(f'No encontre las compras del Portal IVA de {per} (busque en ncr/, en la carpeta IVA y en la del mes). '
         'Bajalas con portal-iva-compras.js o pasa --portal.')


# ── NCR ─────────────────────────────────────────────────────────────────────
def hoja_con(wb, columna):
    for ws in wb.worksheets:
        filas = list(ws.iter_rows(values_only=True))
        if filas and columna in [str(c).strip() for c in filas[0] if c is not None]:
            yield ws.title, [str(c).strip() if c is not None else '' for c in filas[0]], filas[1:]


def leer_monroe(ruta, cuit_soc):
    base = norm(os.path.splitext(os.path.basename(ruta))[0])
    hojas = list(hoja_con(load_workbook(ruta, read_only=True, data_only=True), 'Tipo Linea'))
    if not hojas:
        return None
    _, cab, filas = hojas[0]
    filas = [dict(zip(cab, f)) for f in filas]
    declarado = None
    for f in filas:
        m = re.search(r'Cantidad de Registros:\s*(\d+)', str(f.get('Tipo Linea') or ''))
        if m:
            declarado = int(m.group(1))
    det = [f for f in filas if str(f.get('Tipo Linea') or '').strip() == 'Detalle']
    cuits = {digitos(f.get('Cuit')) for f in det if f.get('Cuit')}
    if cuits and cuits != {cuit_soc}:
        fail(f'{os.path.basename(ruta)} es de otro CUIT ({", ".join(sorted(cuits))}), no de {cuit_soc}. '
             'Sacalo de ncr/ o volve a correr descargar-ncr.')
    ncr = []
    for f in det:
        pv, _, nro = str(f.get('Numero Formateado') or '').partition('-')
        imp = f.get('Importe Total')
        ncr.append({'_drog': 'MONROE', '_farmacia': base, '_cuit': CUIT_MONROE,
                    '_pv': int(pv), '_nro': int(nro),
                    # Monroe: TEXTO con punto decimal y sin miles (-79047.4). float() directo.
                    '_imp': round(float(imp), 2), '_fecha': fecha(f.get('Fecha')),
                    'concepto': str(f.get('Descripcion') or '').strip()})
    pie = {'archivo': base, 'leidas': len(det), 'declaradas': declarado,
           'ok': declarado is None or declarado == len(det)}
    return base, ncr, pie


def leer_suizo(ruta):
    hojas = list(hoja_con(load_workbook(ruta, read_only=True, data_only=True), 'CLIENTE'))
    if len(hojas) != 1:
        fail(f'{os.path.basename(ruta)} tiene {len(hojas)} hojas con columna CLIENTE (se espera una).')
    _, cab, filas = hojas[0]
    out = []
    for f in filas:
        r = dict(zip(cab, f))
        if r.get('terminal') in (None, '') or r.get('numero') in (None, ''):
            continue
        out.append({'_drog': 'SUIZO', '_farmacia': norm(r.get('CLIENTE')), '_cuit': CUIT_SUIZO,
                    '_pv': int(float(r['terminal'])), '_nro': int(float(r['numero'])),
                    '_imp': round(float(r['total']), 2),
                    '_fecha': fecha(r.get('fecha')),       # Suizo: texto DD/MM/YYYY
                    'concepto': str(r.get('concepto') or '').strip()})
    return out


def resolver_clientes(farmacias, vocab):
    try:
        alias = {norm(k): norm(v) for k, v in json.load(open(ALIAS_JSON, encoding='utf-8'))['alias'].items()}
    except (OSError, KeyError, ValueError):
        alias = {}
    res, sin = [], []
    for f in farmacias:
        c = alias.get(f, f)
        (res if c in vocab else sin).append({'farmacia': f, 'cliente': c, 'via': 'alias' if f in alias else 'igual'})
    return res, sin


# ── Excel de salida ─────────────────────────────────────────────────────────
FUENTE, NEGRITA = Font(name='Arial', size=10), Font(name='Arial', size=10, bold=True)


def volcar(ws, cab, filas):
    ws.append(cab)
    for r in filas:
        ws.append([r.get(c) for c in cab])
    for fila in ws.iter_rows():
        for cel in fila:
            cel.font = NEGRITA if cel.row == 1 else FUENTE
            if isinstance(cel.value, datetime.date):
                cel.number_format = 'DD/MM/YYYY'
            elif isinstance(cel.value, float):
                cel.number_format = '#,##0.00'
    for i, c in enumerate(cab, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(c)) + 2, 12), 40)
    ws.freeze_panes = 'A2'
    if filas:
        ws.auto_filter.ref = ws.dimensions


def escribir(destino, cab_portal, resto, elim, no_enc, sin_resp, difs, control):
    wb = Workbook()
    wb.remove(wb.active)
    volcar(wb.create_sheet('Compras S.NCR'), cab_portal, resto)
    cols_ncr = ['_drog', '_farmacia', '_cuit', '_pv', '_nro', '_imp', '_fecha', 'concepto']
    volcar(wb.create_sheet('Eliminados'), cols_ncr + ['portal_importe', 'portal_fecha'], elim)
    volcar(wb.create_sheet('NCR no encontradas'), cols_ncr + ['fuera_del_periodo'], no_enc)
    volcar(wb.create_sheet('NC sin respaldo'), cab_portal + ['drogueria', 'en_suizo_bajo_cliente'], sin_resp)
    if difs:
        volcar(wb.create_sheet('Diferencias'), cols_ncr + ['portal_importe', 'portal_fecha', 'motivo'], difs)
    ws = wb.create_sheet('Control')
    ws.append(['Concepto', 'Comprobantes', 'Importe'])
    for fila in control:
        ws.append(fila)
    ws.append([])
    ws.append(['Recalculado en Excel (Compras S.NCR)', "=COUNTA('Compras S.NCR'!A:A)-1",
               f"=SUM('Compras S.NCR'!{get_column_letter(cab_portal.index('Importe Total') + 1)}:"
               f"{get_column_letter(cab_portal.index('Importe Total') + 1)})"])
    for fila in ws.iter_rows():
        for cel in fila:
            cel.font = NEGRITA if cel.row == 1 else FUENTE
            if cel.column == 3:
                cel.number_format = '#,##0.00'
    ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 48, 14, 20
    wb.save(destino)


# ── principal ───────────────────────────────────────────────────────────────
def main():
    a = args()
    cuit = digitos(a.get('cuit'))
    per = periodo_aaaamm(a.get('periodo'))
    if len(cuit) != 11:
        fail('Falta --cuit (11 digitos).')
    if not per:
        fail('Falta --periodo (MM/AAAA o AAAAMM).')
    si_existe = a.get('si-existe') or 'error'
    if si_existe not in ('error', 'reemplazar', 'saltear'):
        fail('--si-existe tiene que ser error, reemplazar o saltear.')
    solo_revisar = 'solo-revisar' in a

    soc = carpeta_sociedad(cuit)
    ncr_dir = carpeta_ncr(soc, per)
    sociedad = norm(a.get('sociedad')) if a.get('sociedad') else nombre_sociedad(soc)
    salida_nombre = f'{per}-COMPRAS-SIN-NCR-{sociedad}.xlsx'

    re_salida = re.compile(rf'^{per}-COMPRAS-SIN-NCR-.*\.xlsx$', re.I)
    previas = [f for f in os.listdir(ncr_dir) if re_salida.match(f) and not f.startswith('~$')]
    if previas and si_existe == 'saltear':
        print(json.dumps({'ok': True, 'salteada': True, 'cuit': cuit, 'periodo': per,
                          'salida': os.path.join(ncr_dir, previas[0]),
                          'motivo': 'ya habia una conciliacion (--si-existe=saltear)'}, ensure_ascii=False))
        return
    if previas and si_existe == 'error' and not solo_revisar:
        fail(f'Ya hay una conciliacion en ncr/: {" | ".join(previas)}. '
             'Usa --si-existe=reemplazar (el anterior va a ncr/_anteriores/) o --si-existe=saltear.')

    # Fuentes
    portal_ruta = os.path.abspath(a['portal']) if a.get('portal') else ubicar_portal(soc, per, ncr_dir)
    cab, p = leer_portal(portal_ruta)
    log(f'compras: {os.path.basename(portal_ruta)} · {len(p)} comprobantes')
    k = {norm(c).lower(): c for c in cab}
    C_FECHA, C_TIPO, C_PV, C_NRO = k[FECHA], k['tipo de comprobante'], k['punto de venta'], k['numero de comprobante']
    C_CUIT, C_IMP = k['nro. doc. vendedor'], k['importe total']

    archivos = sorted(f for f in os.listdir(ncr_dir) if os.path.isfile(os.path.join(ncr_dir, f))
                      and f.lower().endswith('.xlsx') and not f.startswith('~$') and not re_salida.match(f))
    suizos = [f for f in archivos if 'SUIZO' in norm(f)]
    if len(suizos) != 1:
        fail(f'Se espera un archivo de Suizo en ncr/ y hay {len(suizos)}: {suizos}. Corre descargar-ncr.')
    suizo_todo = leer_suizo(os.path.join(ncr_dir, suizos[0]))

    mon, pies, farmacias = [], [], []
    for f in archivos:
        if f in suizos:
            continue
        r = leer_monroe(os.path.join(ncr_dir, f), cuit)
        if r is None:
            log(f'{f}: no es una exportacion de Monroe, lo ignoro')
            continue
        farmacias.append(r[0]); mon += r[1]; pies.append(r[2])
    if not farmacias and 'sin-monroe' not in a:
        fail('No hay archivos de Monroe en ncr/. Si la sociedad no le compro a Monroe este mes, '
             'pasa --sin-monroe y --clientes=<CLIENTE de Suizo>.')

    vocab = {x['_farmacia'] for x in suizo_todo}
    if a.get('clientes'):
        clientes = [norm(c) for c in a['clientes'].split(',') if c.strip()]
        resolucion, sin_resolver = [{'cliente': c, 'via': '--clientes'} for c in clientes], []
    else:
        resolucion, sin_resolver = resolver_clientes(farmacias, vocab)
        if sin_resolver:
            fail('Hay farmacias de Monroe sin CLIENTE en Suizo. Preguntar a cual corresponde y agregarlo a '
                 'lib/ncr-alias.json (o pasar --clientes).', sin_resolver_en_suizo=sin_resolver)
        clientes = sorted({x['cliente'] for x in resolucion})
    suizo = [x for x in suizo_todo if x['_farmacia'] in clientes]
    ncr = mon + suizo
    log(f'NCR: Monroe {len(mon)} ({len(farmacias)} archivos) + Suizo {len(suizo)} ({", ".join(clientes)})')

    # Matcheo: tipo 3 + CUIT emisor + PV + numero (enteros)
    idx = {}
    for i, r in enumerate(p):
        if r[C_TIPO] == TIPO_NC_A:
            idx.setdefault((digitos(r[C_CUIT]), r[C_PV], r[C_NRO]), []).append(i)

    elim_i, elim, no_enc, difs, vistas = set(), [], [], [], set()
    for x in ncr:
        key = (x['_cuit'], x['_pv'], x['_nro'])
        hit = idx.get(key)
        if key in vistas:
            difs.append({**x, 'motivo': 'NCR repetida en los archivos'}); continue
        vistas.add(key)
        if not hit:
            fuera = not (x['_fecha'] and x['_fecha'].strftime('%Y%m') == per)
            no_enc.append({**x, 'fuera_del_periodo': 'SI' if fuera else 'NO'}); continue
        if len(hit) > 1:
            difs.append({**x, 'motivo': f'{len(hit)} filas del Portal con la misma clave'}); continue
        i = hit[0]
        reg = {**x, 'portal_importe': p[i][C_IMP], 'portal_fecha': p[i][C_FECHA]}
        motivos = []
        if abs(round((p[i][C_IMP] or 0) - x['_imp'], 2)) > 0.01:
            motivos.append('importe distinto')
        if p[i][C_FECHA] != x['_fecha']:
            motivos.append('fecha distinta')
        if motivos:
            difs.append({**reg, 'motivo': ' y '.join(motivos)}); continue
        elim_i.add(i); elim.append(reg)

    # NC de las droguerias sin respaldo: se informan, NO se eliminan
    suizo_por_clave = {}
    for x in suizo_todo:
        suizo_por_clave.setdefault((x['_pv'], x['_nro']), set()).add(x['_farmacia'])
    sin_resp = []
    for i, r in enumerate(p):
        c = digitos(r[C_CUIT])
        if r[C_TIPO] == TIPO_NC_A and c in DROGUERIAS and i not in elim_i:
            otro = sorted(suizo_por_clave.get((r[C_PV], r[C_NRO]), set())) if c == CUIT_SUIZO else []
            sin_resp.append({**r, 'drogueria': DROGUERIAS[c], 'en_suizo_bajo_cliente': ', '.join(otro)})

    resto = [r for i, r in enumerate(p) if i not in elim_i]
    suma = lambda filas, col: round(sum((r[col] or 0) for r in filas), 2)
    total_portal, total_elim, total_resto = suma(p, C_IMP), suma([p[i] for i in elim_i], C_IMP), suma(resto, C_IMP)
    verif = round(total_portal - total_elim - total_resto, 2)

    fechas = [r[C_FECHA] for r in p if r[C_FECHA]]
    ult = monthrange(int(per[:4]), int(per[4:]))[1]
    avisos = []
    if fechas and max(fechas).strftime('%Y%m') == per and max(fechas).day < ult - 5:
        avisos.append(f'El ultimo comprobante de compras es del {max(fechas):%d/%m}: la descarga puede estar cortada.')
    if any(not pz['ok'] for pz in pies):
        avisos.append('Hay exportaciones de Monroe truncadas: rehacerlas y volver a conciliar.')

    control = [
        ['Compras Portal IVA', len(p), total_portal],
        ['NCR conciliadas (eliminadas)', len(elim_i), total_elim],
        ['Compras S.NCR', len(resto), total_resto],
        ['Verificacion aritmetica (tiene que dar 0)', None, verif],
        [],
        ['NCR leidas Monroe', len(mon), round(sum(x['_imp'] for x in mon), 2)],
        ['NCR leidas Suizo', len(suizo), round(sum(x['_imp'] for x in suizo), 2)],
        ['NCR no encontradas en el Portal', len(no_enc), round(sum(x['_imp'] for x in no_enc), 2)],
        ['Diferencias (no eliminadas)', len(difs), round(sum(x['_imp'] for x in difs), 2)],
        ['NC de Suizo/Monroe sin respaldo (no eliminadas)', len(sin_resp), suma(sin_resp, C_IMP)],
        [],
        ['Fuente compras', os.path.basename(portal_ruta)],
        ['Fuente Suizo', suizos[0], ', '.join(clientes)],
    ] + [[f'Monroe {pz["archivo"]}', f'{pz["leidas"]} de {pz["declaradas"]}', 'OK' if pz['ok'] else 'TRUNCADA'] for pz in pies] + [
        ['Generado por conciliar-ncr.py', datetime.datetime.now().strftime('%d/%m/%Y %H:%M')],
    ]

    resumen = {
        'ok': verif == 0, 'cuit': cuit, 'sociedad': sociedad, 'periodo': per,
        'compras': portal_ruta,
        'compras_fechas': [min(fechas).isoformat(), max(fechas).isoformat()] if fechas else None,
        'portal_comprobantes': len(p), 'portal_total': total_portal,
        'ncr_leidas': len(ncr), 'ncr_monroe': len(mon), 'ncr_suizo': len(suizo),
        'clientes_suizo': clientes, 'resolucion_suizo': resolucion,
        'eliminados': len(elim_i), 'eliminados_total': total_elim,
        'resultantes': len(resto), 'resultantes_total': total_resto, 'verificacion': verif,
        'ncr_no_encontradas': len(no_enc),
        'ncr_no_encontradas_detalle': [{k2: x[k2] for k2 in ('_drog', '_pv', '_nro', '_imp', '_fecha', 'fuera_del_periodo')} for x in no_enc],
        'diferencias': len(difs),
        'diferencias_detalle': [{k2: x.get(k2) for k2 in ('_drog', '_pv', '_nro', '_imp', '_fecha', 'motivo')} for x in difs],
        'nc_sin_respaldo': len(sin_resp), 'nc_sin_respaldo_total': suma(sin_resp, C_IMP),
        'nc_sin_respaldo_por_drogueria': {
            d: {'comprobantes': len([r for r in sin_resp if r['drogueria'] == d]),
                'total': suma([r for r in sin_resp if r['drogueria'] == d], C_IMP)} for d in ('MONROE', 'SUIZO')},
        'suizo_sin_respaldo_bajo_otro_cliente': [
            {'pv': r[C_PV], 'nro': r[C_NRO], 'cliente': r['en_suizo_bajo_cliente']} for r in sin_resp if r['en_suizo_bajo_cliente']],
        'pies_archivos_monroe': pies, 'avisos': avisos,
    }

    if verif != 0:
        fail(f'La verificacion aritmetica no cierra ({verif}). No escribo nada.', resumen=resumen)
    if solo_revisar:
        resumen.update({'modo': 'solo_revisar', 'salida': None, 'ya_habia': previas})
        print(json.dumps(resumen, ensure_ascii=False, indent=2, default=str))
        return

    # Escribe en temporal, relee y controla, y recien ahi lo deja en ncr/
    tmp = os.path.join(tempfile.mkdtemp(prefix='conciliar-ncr-'), salida_nombre)
    escribir(tmp, cab, resto, elim, no_enc, sin_resp, difs, control)
    wb = load_workbook(tmp, read_only=True, data_only=True)
    ws = wb['Compras S.NCR']
    rows = list(ws.iter_rows(values_only=True))
    ci = list(rows[0]).index(C_IMP)
    releido = round(sum((r[ci] or 0) for r in rows[1:]), 2)
    n_elim = sum(1 for r in wb['Eliminados'].iter_rows(min_row=2, values_only=True) if any(r))
    if len(rows) - 1 != len(resto) or releido != total_resto or n_elim != len(elim):
        fail('El Excel escrito no cierra contra lo calculado. No lo dejo en ncr/.',
             releido={'comprobantes': len(rows) - 1, 'total': releido, 'eliminados': n_elim})

    if previas:
        ant = os.path.join(ncr_dir, '_anteriores')
        os.makedirs(ant, exist_ok=True)
        sello = datetime.datetime.now().strftime('%Y%m%d-%H%M')
        for f in previas:
            shutil.move(os.path.join(ncr_dir, f), os.path.join(ant, f.replace('.xlsx', f' ({sello}).xlsx')))
            log(f'{f} -> _anteriores/')
    destino = os.path.join(ncr_dir, salida_nombre)
    shutil.move(tmp, destino)
    resumen.update({'modo': 'escrito', 'salida': destino, 'reemplazados': previas})
    print(json.dumps(resumen, ensure_ascii=False, indent=2, default=str))


if __name__ == '__main__':
    main()
