# Arma y actualiza los Estados de Resultados (EERR) mensuales de las sociedades, desde las fuentes que ya están en la
# unidad compartida. Es el armado con el que se hicieron los EERR 2026 (01-08/10/2026), guardado como herramienta
# (pedido del equipo en la reunión del 04/11/2026).
#
# Dos modos:
#   actualizar (default): completa en el EERR existente los meses del ejercicio hasta --hasta. Cada concepto de cada mes
#                         se carga solo si está vacío: nunca pisa lo que ya tiene (lo que cargó o corrigió el equipo),
#                         salvo los meses que se pidan con --pisar.
#   --crear:              arma el EERR del ejercicio siguiente al último que tiene la sociedad (o el que empieza en
#                         --desde) desde la plantilla limpia, y lo completa hasta --hasta.
#
# De dónde sale cada fila (criterios acordados con el usuario, 01-08/10/2026):
#   Ventas exentas y no gravadas / gravadas  libro de VENTAS del Portal IVA (exento + no gravado + neto 0% / neto gravado),
#                                            por tipo de comprobante en la hoja "Ventas Portal IVA"
#   CMV y gastos de compras ................ libro de COMPRAS del Portal IVA imputado por proveedor con el Maestro de
#                                            Proveedores (manda el Maestro), hoja "Compras imputadas" + SUMIFS.
#                                            Importe: no gravado + exento + netos; sin IVA discriminado (B, C, tiques),
#                                            total - percepciones. Bloqueados y bienes de uso quedan fuera del resultado;
#                                            estado Revisar o sin Maestro -> GASTOS VARIOS (y se avisa).
#   IIBB ................................... impuesto determinado de la DJ de anticipo de ARBA (PDF del mes)
#   IVA Prorrateo .......................... IVA de las compras que no son CMV x proporción de ventas exentas y no
#                                            gravadas del mes (mismo método que el papel de trabajo), hoja "IVA prorrateo"
#   Sueldos y Cargas Sociales .............. F.931 del mes (sección VIII; rectificativa con VIII en cero: I, II y VI)
#   Seguridad e Higiene .................... ventas gravadas x alícuota de la tasa + conceptos fijos, tomados del papel de
#                                            trabajo del equipo (hoja MM-AAAA del "...liquidaciones..."; si el mes no
#                                            está cargado, la alícuota del último mes que la tenga)
#   Publicidad ............................. PyP del papel de trabajo del mes + compras imputadas a Publicidad
# Lo que no tiene fuente queda vacío y figura en la hoja "Armado" (registro de cada corrida: mes, concepto, acción,
# valor y archivo de origen).
#
#   python3 armar_eerr.py --nombre="<sociedad>" [--hasta=AAAAMM] [--desde=AAAAMM] [--pisar=AAAAMM,AAAAMM] [--simular]
#   python3 armar_eerr.py --todas [--hasta=AAAAMM] [--simular]
#   python3 armar_eerr.py --nombre="<sociedad>" --crear [--desde=AAAAMM] [--hasta=AAAAMM] [--simular]
#   python3 armar_eerr.py --todas --crear [--hasta=AAAAMM] [--simular]   (las sociedades cuyo ejercicio ya cerró)
# --hasta: último mes a cargar (default: último mes con la DDJJ de IVA vencida, como el auditor).
# --desde: (actualizar) primer mes que se revisa; por defecto los 3 últimos meses hasta --hasta. (crear) primer mes del ejercicio.
# --pisar: meses que se rehacen aunque tengan datos (los valores anteriores quedan en un comentario de la celda).
# --simular: no toca la unidad; deja el resultado en <BVA_SALIDAS_PATH>/armado-eerr/simulacion/.
# Excel recalcula cada archivo antes de devolverlo a la unidad (Windows: PowerShell + Excel; Mac: AppleScript + Excel).
# Si el archivo cambió en la unidad mientras se armaba, no se pisa: queda en armado-eerr/no-guardados/.
import os, re, sys, glob, json, shutil, datetime, subprocess, unicodedata, collections, time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL

WIN, MAC = os.name == 'nt', sys.platform == 'darwin'
FISCO = os.path.expanduser('~/.fisco-ar')
SP = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.path.join(SP, '..', 'assets', 'Plantilla EERR.xlsx')


def _env():
    f = os.environ.get('FISCO_ENV_FILE') or os.path.join(FISCO, '.env'); out = {}
    if os.path.exists(f):
        for l in open(f, encoding='utf-8', errors='replace'):
            m = re.match(r'^\s*(?:export\s+)?([A-Za-z_]\w*)\s*=\s*(.*?)\s*$', l)
            if m and not l.lstrip().startswith('#'): out[m[1]] = m[2].strip('"').strip("'")
    return lambda k: os.environ.get(k) or out.get(k) or ''
var = _env()

NOMBRE_UNIDAD = 'BVA - Sociedades Farmaceuticas'
def raiz_unidad():
    r = var('BVA_UNIDAD_PATH').strip()
    if r:
        if not os.path.isdir(r): sys.exit(f'BVA_UNIDAD_PATH apunta a "{r}", que no existe. Verificá que Google Drive esté abierto.')
        return r
    if WIN and os.path.isdir(f'I:\\Unidades compartidas\\{NOMBRE_UNIDAD}'): return f'I:\\Unidades compartidas\\{NOMBRE_UNIDAD}'
    c = glob.glob(os.path.expanduser(f'~/Library/CloudStorage/GoogleDrive-*/Unidades compartidas/{NOMBRE_UNIDAD}'))
    if len(c) == 1: return c[0]
    sys.exit('No encuentro la unidad compartida: definí BVA_UNIDAD_PATH en ~/.fisco-ar/.env.')

U = raiz_unidad()
ESTUDIO = os.path.join(U, '0 - ESTUDIO - 00000000000')
EERR_DIR = os.environ.get('ARMADO_EERR_DIR') or os.path.join(ESTUDIO, 'EERR 2026')
MAESTRO = os.path.join(ESTUDIO, 'Maestro de Proveedores.xlsx')
SALIDAS = var('BVA_SALIDAS_PATH') or os.path.expanduser('~/Documents/BVA-salidas')
SALIDA = os.path.join(SALIDAS, 'armado-eerr')
ARBA_LOCAL = os.path.join(SALIDAS, 'arba-ddjj')          # DJ de ARBA que baja constancias-anticipo-arba, si están
F931_LOCAL = os.path.join(SALIDAS, 'f931')               # F.931 que baja f931-descarga-arca, si están

# Cuenta del Maestro -> fila del EERR (etiqueta de la columna A). Acordado el 01/10/2026.
RUBRO = {
 'CMV': 'CMV', 'AGUA': 'Agua', 'GAS': 'Metrogas', 'LUZ': 'Edesur',
 'GASTOS BANCARIOS': 'Gastos Bancarios', 'INTERESES PAGADOS': 'Intereses de Prestamos',
 'TARJETA DE CRÉDITO': 'GASTOS DE TARJETAS', 'GASTOS DE TARJETAS': 'GASTOS DE TARJETAS', 'ALQUILERES': 'ALQUILERES',
 'GASTOS DE COMERCIALIZACION': 'GASTOS DE COMERCIALIZACION',
 'COMISION MERCADO LIBRE': 'COMISIONES DE MERCADO LIBRE', 'VENTA ON LINE': 'COMISIONES DE MERCADO LIBRE',
 'MANTEN. ESTRUCTURA': 'GASTOS DE MANTENIMIENTO', 'MANTEN. RODADOS': 'GASTOS DE MANTENIMIENTO',
 'GASTOS DE MANTENIMIENTO': 'GASTOS DE MANTENIMIENTO',
 'MANTEN. EQ. DE COMPUTACION': 'GASTOS DE MANTENIMIENTO BS . USO',
 'GASTOS DE SISTEMAS': 'GASTOS DE SISTEMAS', 'SERVICIOS DE INFORMATICA': 'GASTOS DE SISTEMAS',
 'GASTOS DE TRANSPORTE': 'GASTOS DE TRANSPORTE', 'GASTOS DE MENSAJERIA, MOVILIDAD, VIATICOS': 'GASTOS DE TRANSPORTE',
 'GASTOS VARIOS': 'GASTOS VARIOS', 'ARCHIVO': 'GASTOS VARIOS', 'GASTOS DE ADMINISTRACION': 'GASTOS VARIOS', 'A DEFINIR': 'GASTOS VARIOS',
 'HONORARIOS POR SERVICIOS': 'HONORARIOS POR SERVICIOS',
 'IMPRENTA Y LIBRERÍA': 'IMPRENTA Y PAPELERIA', 'PAPELERA': 'IMPRENTA Y PAPELERIA',
 'LIMPIEZA': 'GASTOS DE LIMPIEZA', 'INTERNET / TELEFONIA': 'INTERNET/TELEFONIA', 'REFRIGERIO': 'REFRIGERIO',
 'GASTOS DE REPRESENTACION - VIAJE': 'GASTOS DE REPRESENTACION', 'GASTOS DE REPRESENTACION - EVENTOS': 'GASTOS DE REPRESENTACION',
 'GASTOS DE REPRESENTACION - ALOJAMIENTO': 'GASTOS DE REPRESENTACION', 'REGALOS EMPRESARIALES': 'GASTOS DE REPRESENTACION',
 'RESIDUOS PATOGENICOS': 'RESIDUOS PATOGENICOS', 'SEGURIDAD': 'SEGURIDAD', 'GASTOS FINANCIEROS': 'SERVICIOS FINANCIEROS',
 'SEGUROS': 'SEGUROS', 'GASTOS MEDICOS': 'SERVICIOS MEDICOS', 'SERVICIOS DE EMERGENCIAS Y TRASLADOS': 'SERVICIOS MEDICOS',
 'UNIFORMES': 'UNIFORMES', 'MATRICULA COLEGIO FARMACEUTICOS': 'Matricula Colegio Farmaceuticos',
 'PUBLICIDAD - CARTELERIA': 'Publicidad', 'PUBLICIDAD - FOLLETOS': 'Publicidad',
}
EXCL_BLOQ, EXCL_BU = 'EXCLUIDO - PROVEEDOR BLOQUEADO', 'EXCLUIDO - BIEN DE USO'
ALICS = ['2,5%', '5%', '10,5%', '21%', '27%']
PERC = ['Importe de Per. o Pagos a Cta. de Otros Imp. Nac.', 'Importe de Percepciones de Ingresos Brutos', 'Importe de Impuestos Municipales',
        'Importe de Percepciones o Pagos a Cuenta de IVA', 'Importe de Impuestos Internos', 'Importe Otros Tributos']
TIPOS = {1: 'Factura A', 2: 'Nota de Débito A', 3: 'Nota de Crédito A', 6: 'Factura B', 7: 'Nota de Débito B', 8: 'Nota de Crédito B',
         11: 'Factura C', 12: 'Nota de Débito C', 13: 'Nota de Crédito C', 51: 'Factura M', 52: 'Nota de Débito M', 53: 'Nota de Crédito M',
         81: 'Tique Factura A', 82: 'Tique Factura B', 83: 'Tique', 110: 'Tique Nota de Crédito', 111: 'Tique Factura C',
         112: 'Tique Nota de Crédito A', 113: 'Tique Nota de Crédito B', 114: 'Tique Nota de Crédito C', 115: 'Tique Nota de Débito A',
         116: 'Tique Nota de Débito B', 201: 'Factura de Crédito Electrónica A', 202: 'Nota de Débito FCE A', 203: 'Nota de Crédito FCE A',
         206: 'Factura de Crédito Electrónica B', 207: 'Nota de Débito FCE B', 208: 'Nota de Crédito FCE B'}
FMT = '#,##0.00;[Red]-#,##0.00'
VERDE = PatternFill('solid', fgColor='8BC34A')
AMARILLO = PatternFill('solid', fgColor='FFF2CC')


def num(v): return isinstance(v, (int, float)) and not isinstance(v, bool)
def a_num(v):
    # número de una celda: los libros que vienen de CSV (p. ej. el "sin NCR" de la conciliación) traen texto "1.234,56"
    if num(v): return float(v)
    if isinstance(v, str) and re.fullmatch(r'\s*-?[\d\.]*,?\d*\s*', v) and re.search(r'\d', v):
        t = v.strip()
        t = t.replace('.', '').replace(',', '.') if ',' in t else t
        try: return float(t)
        except ValueError: return 0.0
    return 0.0
def ndoc(v): return str(int(v)) if num(v) else re.sub(r'\D', '', str(v or ''))
def nfc(s): return unicodedata.normalize('NFC', str(s or '')).strip()
def norm(s): return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', nfc(s).lower()) if unicodedata.category(c) != 'Mn')).strip()
def mm(per): return f"{per[4:]}/{per[:4]}"
def hoja_mes(per): return f"{per[4:]}-{per[:4]}"
def pesos(v): return f"$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
def mas_mes(per, n=1):
    y, m = int(per[:4]), int(per[4:]) + n
    while m > 12: y, m = y + 1, m - 12
    while m < 1: y, m = y - 1, m + 12
    return f"{y}{m:02d}"
def fecha(per): return datetime.datetime(int(per[:4]), int(per[4:]), 1)


def abrir(f, **kw):
    for i in range(4):
        try: return openpyxl.load_workbook(f, **kw)
        except Exception:
            if i == 3: raise
            time.sleep(5 * (i + 1))


def hasta_default(hoy=None):
    hoy = hoy or datetime.date.today()
    y, m = hoy.year, hoy.month - (1 if hoy.day >= 26 else 2)
    while m < 1: y, m = y - 1, m + 12
    return f"{y}{m:02d}"


carpetas, cortos = {}, {}
for d in os.listdir(U):
    x = re.match(r'^\d+ - (.+) - (\d{11})$', nfc(d))
    if x: carpetas[x.group(2)] = d; cortos[x.group(2)] = x.group(1).strip()


# ------------------------------------------------------------------ fuentes
class Fuentes:
    """Archivos de una sociedad en la unidad (01-Impuestos Mensuales y 02-Sueldos y F.931) y lectura de cada fuente."""
    def __init__(self, cuit):
        self.cuit = cuit
        self.carpeta = os.path.join(U, carpetas[cuit]) if cuit in carpetas else None
        self.idx = []
        if self.carpeta:
            for sub in ('01-Impuestos Mensuales', '02-Sueldos y F.931'):
                for raiz, _, fs in os.walk(os.path.join(self.carpeta, sub)):
                    for x in fs:
                        if not x.startswith(('~$', '.')): self.idx.append((nfc(x), os.path.join(raiz, x)))
        self._papel = None
        self.tasas_sh = {}        # {AAAAMM: S&H / ventas gravadas} de los EERR de la sociedad (respaldo sin alícuota en el papel)

    def buscar(self, patron, per=None):
        rx = re.compile(patron, re.I)
        out = [p for n, p in self.idx if rx.search(n)]
        if per: out = [p for p in out if per in nfc(os.path.relpath(os.path.dirname(p), self.carpeta)).split(os.sep) or nfc(os.path.basename(p)).startswith(per)]
        return sorted(out, key=lambda p: (len(p), p))

    # ---- libros del Portal IVA
    def libro(self, per, cual):
        # orden: libro de la DDJJ presentada ("(DDJJ presentada)" en el nombre) > el más nuevo entre el libro del Portal IVA
        # (carpeta IVA primero) y, en compras, el libro sin NCR de la conciliación: un libro bajado después de conciliar
        # ya viene sin NCR y con lo que el equipo agregó (Edesur, AySA...), uno anterior trae las NCR. Nunca cruces ni CSV.
        C = cual.upper()
        cand = [p for p in self.buscar(rf'^{per}\s*-\s*PORTAL IVA\s*-\s*{C}.*\.xlsx$', per) if 'cruce' not in norm(os.path.basename(p))]
        pres = [p for p in cand if 'presentada' in norm(os.path.basename(p))]
        if pres: return pres[0], 'libro de la DDJJ presentada'
        iva = [p for p in cand if os.path.basename(os.path.dirname(p)).upper() == 'IVA' and
               re.match(rf'^{per} - PORTAL IVA - {C}\.xlsx$', nfc(os.path.basename(p)), re.I)]
        libro = (iva or cand or [None])[0]
        que = 'libro del Portal IVA' if iva else 'libro del Portal IVA (archivo del equipo)'
        # un libro bajado después del acuse de la DDJJ de IVA del mes es el presentado (el borrador ya no se puede abrir)
        acuses = [os.path.getmtime(p) for p in self.buscar(rf'^{per}-IVA-(ACUSE|DDJJ)', per)]
        if libro and acuses and os.path.getmtime(libro) > min(acuses):
            return libro, 'libro de la DDJJ presentada (bajado después del acuse)'
        if C == 'COMPRAS':
            sin = self.buscar(rf'^{per}-COMPRAS-SIN-NCR.*\.xlsx$', per)
            if sin and (libro is None or os.path.getmtime(sin[0]) > os.path.getmtime(libro)):
                return sin[0], 'libro de compras sin NCR (conciliación NCR)'
        return (libro, que) if libro else (None, None)

    # ---- DJ de ARBA
    def dj_arba(self, per):
        for f in self.buscar(rf'^{per}-IIBB-(DJ|DDJJ)\b.*\.pdf$', per) + [os.path.join(ARBA_LOCAL, self.cuit, f'{per}.pdf')]:
            if os.path.exists(f):
                try: v = dj_arba(f)
                except Exception: v = None
                if v is not None: return v, f
        return None, None

    # ---- F.931
    def f931(self, per):
        cands = self.buscar(rf'^{per}-F\.?931.*\.pdf$') + [os.path.join(F931_LOCAL, self.cuit, f'{per}.pdf')]
        for f in cands:
            if os.path.exists(f):
                try: d = leer_f931(f)
                except Exception: d = None
                if d and d.get('rem_art') is not None: return d, f
        return None, None

    # ---- papel de trabajo del equipo (hojas MM-AAAA de "AAAA00-liquidaciones-...")
    def papel(self):
        if self._papel is not None: return self._papel
        self._papel = {}
        # el papel del equipo manda; el "- AUTOMATICO" (papel-trabajo-automatico) cubre los meses que el del equipo no tiene
        fs = [p for n, p in self.idx if 'liquidaciones' in norm(n) and n.lower().endswith('.xlsx') and not re.search(r'no va|no usar', norm(n))]
        for f in sorted(fs, key=lambda p: ('automatico' not in norm(os.path.basename(p)), os.path.basename(p))):
            try: wb = abrir(f, read_only=True, data_only=True)
            except Exception: continue
            for h in wb.sheetnames:
                m = re.match(r'^(\d{2})-(\d{4})$', h.strip())
                if not m: continue
                d = tasas_papel(list(wb[h].iter_rows(max_col=12, values_only=True)))
                if not d: continue
                per = f"{m[2]}{m[1]}"; d['f'] = f; d['hoja'] = h
                ant = self._papel.get(per)
                if ant and not num(d.get('alic')): d['alic'], d['fijos'] = ant['alic'], ant['fijos']
                if ant and d.get('pyp') is None: d['pyp'] = ant.get('pyp')
                self._papel[per] = d
        return self._papel


def filas_libro(f):
    wb = abrir(f, read_only=True, data_only=True)
    ws = wb['Compras S.NCR'] if 'Compras S.NCR' in wb.sheetnames else wb.worksheets[0]
    it = ws.iter_rows(values_only=True); h = {nfc(k): i for i, k in enumerate(next(it)) if k}
    for r in it:
        if r[0] is None or r[1] is None: break
        g = lambda k: a_num(r[h[k]]) if k in h else 0.0
        tipo = re.match(r'\s*(\d+)', str(r[h['Tipo de Comprobante']] if 'Tipo de Comprobante' in h else ''))
        doc = 'Nro. Doc. Vendedor' if 'Nro. Doc. Vendedor' in h else 'Nro. Doc. Comprador'
        den = 'Denominación Vendedor' if 'Denominación Vendedor' in h else 'Denominación Comprador'
        ivas = {a: g(f'Importe IVA {a}') for a in ALICS}
        yield dict(cuit=ndoc(r[h[doc]]) if doc in h else '', den=nfc(r[h[den]]) if den in h else '',
                   tipo=int(tipo[1]) if tipo else None, nograv=g('Importe No Gravado'), exento=g('Importe Exento'),
                   neto=sum(g(f'Neto Gravado IVA {a}') for a in ALICS), neto0=g('Neto Gravado IVA 0%'),
                   tot_neto=g('Total Neto Gravado'), iva=g('Total IVA') or sum(ivas.values()), ivas=ivas,
                   total=g('Importe Total'), perc=sum(g(k) for k in PERC))


def gasto(x):
    # con IVA discriminado: no gravado + exento + netos; sin IVA discriminado (B, C, tiques, recibos C): total - percepciones
    monto = x['nograv'] + x['exento'] + x['neto0'] + (x['tot_neto'] or x['neto'])
    if abs(monto) > 0.004 or abs(x['iva']) > 0.004: return monto
    return x['total'] - x['perc']


def pdf_texto(f):
    import pypdf
    return '\n'.join(p.extract_text() or '' for p in pypdf.PdfReader(f).pages)


ACT = re.compile(r'^(\d{6}) \$ ?(-?[\d\.]+,\d{2}) (?:\* )?(\d+(?:,\d+)?) \$ ?(-?[\d\.]+,\d{2})$')
def dj_arba(f):
    # impuesto determinado de la DJ de anticipo (total después de las actividades; si no se lee, base x alícuota)
    n = lambda s: float(s.replace('$', '').replace('.', '').replace(',', '.').strip())
    L = [l.strip() for l in pdf_texto(f).splitlines()]
    acts, last = [], None
    for i, l in enumerate(L):
        m = ACT.match(l)
        if m: acts.append((n(m[2]), n(m[3]))); last = i
    if last is None: return None
    if last + 1 < len(L) and L[last + 1].startswith('$'): return n(L[last + 1])
    return round(sum(b * a / 100 for b, a in acts), 2)


def leer_f931(ruta):
    # sección VIII "Montos que se ingresan"; rectificativa con la VIII en cero: secciones I, II y VI (skill f931-sueldos)
    t = pdf_texto(ruta)
    N = r'(-?[\d\.]+,\d{2})'
    g = lambda pat: (lambda m: float(m.group(1).replace('.', '').replace(',', '.')) if m else None)(re.search(pat, t))
    rect = int((re.search(r'Rect\. \(1/9\):\s*(\d)', t) or [0, 0])[1])
    d = dict(rect=rect, empleados=int((re.search(r'Empleados en n[oó]mina:\s*(\d+)', t) or [0, 0])[1]),
             rem_art=g(r'Remun\. con ART\s+' + N),
             c301=g(r'301 - Aportes de Seguridad Social\s+' + N), c302=g(r'302 - Aportes de Obra Social\s+' + N),
             c351=g(r'351 - Contribuciones de Seguridad Social\s+' + N), c352=g(r'352 - Contribuciones de Obra Social\s+' + N),
             c312=g(r'312 - L\.R\.T\.\s+' + N), c028=g(r'028 - Seguro Colectivo de Vida Obligatorio\s+' + N), fuente='sección VIII')
    if rect > 0 and all((d[k] or 0) == 0 for k in ('c301', 'c302', 'c351', 'c352', 'c312', 'c028')):
        d.update(c301=g(r'a3 - Aportes S\.S\. a pagar\s+' + N), c302=g(r'a3 - Aportes O\.S\. a pagar\s+' + N),
                 c351=g(r'Contribuciones S\.S\. a pagar\s+' + N), c352=g(r'Contribuciones O\.S\. a pagar\s+' + N),
                 c312=g(r'L\.R\.T\. total a pagar\s+' + N), c028=g(r'S\.C\.V\.O\. a Pagar:\s*' + N),
                 fuente='secciones I, II y VI (rectificativa con la VIII en cero)')
    return d


def tasas_papel(rows):
    # Seguridad e Higiene (alícuota, conceptos fijos e importe) y PyP de una hoja MM-AAAA del papel de liquidaciones.
    # Etiquetas en la columna B, valores en C (mismo modelo en todas las sociedades).
    B = lambda r: norm(r[1]) if len(r) > 1 else ''
    C = lambda r: float(r[2]) if len(r) > 2 and num(r[2]) else None
    i_base = next((i for i, r in enumerate(rows) if B(r) == 'base imponible (gravado)'), None)
    i_alic = next((i for i, r in enumerate(rows) if B(r) == 'alicuota' and (i_base is None or i > i_base)), None)
    imp = [i for i, r in enumerate(rows) if B(r) == 'importe a pagar']
    if i_alic is None or not imp: return None
    i_sh = next((i for i in imp if i > i_alic), None)
    if i_sh is None: return None
    fijos = [(nfc(rows[i][1]), C(rows[i]) or 0.0) for i in range(i_alic + 1, i_sh) if B(rows[i])]
    i_pp = next((i for i in imp if i > i_sh), None)
    i_pyp = next((i for i, r in enumerate(rows) if B(r) == 'pyp'), None)
    i_int = next((i for i, r in enumerate(rows) if B(r) == 'intereses' and i_pyp is not None and i > i_pyp), None)
    pyp = None
    if i_pyp is not None:
        pyp = (C(rows[i_pyp]) or 0.0) + ((C(rows[i_int]) or 0.0) if i_int is not None else 0.0)
    base = C(rows[i_base]) if i_base is not None else None
    return dict(alic=C(rows[i_alic]), fijos=fijos, sh=C(rows[i_sh]), base=base, pyp=pyp, pp=C(rows[i_pp]) if i_pp else None)


_maestro = None
def maestro():
    global _maestro
    if _maestro is None:
        ws = abrir(MAESTRO, read_only=True, data_only=True)['MAESTRO']
        _maestro = {str(r[0]).strip(): r for r in ws.iter_rows(min_row=3, values_only=True) if r[0]}
    return _maestro


def imputacion(cuit, den):
    r = maestro().get(cuit)
    if not r: return ('SIN MAESTRO', 'GASTOS VARIOS', 'Sin maestro', den)
    cuenta, estado, razon = r[4], r[5], r[1] or den
    if estado == 'Bloqueado': return (cuenta, EXCL_BLOQ, estado, razon)
    if estado == 'Revisar': return (cuenta, 'GASTOS VARIOS', estado, razon)
    if cuenta and str(cuenta).startswith('COMPRA DE EQUIPOS'): return (cuenta, EXCL_BU, estado, razon)
    return (cuenta, RUBRO.get(cuenta, 'GASTOS VARIOS'), estado if cuenta in RUBRO else f"{estado} (cuenta sin fila en el EERR)", razon)


def datos_mes(F, per):
    """Todo lo que el EERR necesita de un mes, con la fuente de cada dato. None = sin fuente."""
    d = dict(per=per, fuentes={}, avisos=[])
    # ventas
    f, que = F.libro(per, 'ventas')
    if f:
        tipos = collections.defaultdict(lambda: [0.0] * 5 + [0])
        for x in filas_libro(f):
            t = tipos[x['tipo']]
            t[0] += x['tot_neto'] or x['neto']; t[1] += x['nograv']; t[2] += x['exento'] + x['neto0']; t[3] += x['iva']; t[4] += x['total']; t[5] += 1
        d['ventas'] = dict(tipos)
        d['gr'] = sum(v[0] for v in tipos.values()); d['ex'] = sum(v[1] + v[2] for v in tipos.values())
        d['fuentes']['ventas'] = (f, que)
    # compras
    f, que = F.libro(per, 'compras')
    if f:
        acc = collections.defaultdict(float); info = {}; iva_serv = collections.defaultdict(float); iva_merc = iva_fuera = 0.0
        n = 0
        for x in filas_libro(f):
            if not x['cuit']: continue
            cta, rubro, est, razon = imputacion(x['cuit'], x['den'])
            k = (x['cuit'], cta, rubro, est); info[k] = razon; acc[k] += gasto(x); n += 1
            if rubro.startswith('EXCLUIDO'): iva_fuera += x['iva']
            elif rubro == 'CMV': iva_merc += x['iva']
            else:
                for a in ALICS: iva_serv[a] += x['ivas'][a]
                iva_serv['otros'] += x['iva'] - sum(x['ivas'].values())
        d['compras'] = dict(acc); d['compras_info'] = info; d['comprobantes'] = n
        d['iva_serv'] = dict(iva_serv); d['iva_merc'] = iva_merc; d['iva_fuera'] = iva_fuera
        d['fuentes']['compras'] = (f, que)
        if 'sin NCR' not in que and F.buscar(rf'^{per}-COMPRAS-SIN-NCR', per) == [] and que != 'libro de la DDJJ presentada':
            pass
    # IIBB
    v, f = F.dj_arba(per)
    if v is not None: d['iibb'] = v; d['fuentes']['iibb'] = (f, 'DJ de anticipo de ARBA')
    # F.931
    v, f = F.f931(per)
    if v: d['f931'] = v; d['fuentes']['f931'] = (f, f"F.931 {'original' if not v['rect'] else 'rectificativa ' + str(v['rect'])}, {v['fuente']}")
    # S&H y PyP del papel de trabajo
    P = F.papel()
    if 'gr' in d:
        ref = P.get(per) if P.get(per) and num(P[per].get('alic')) else None
        if ref is None:
            ant = [p for p in sorted(P) if p <= per and num(P[p].get('alic'))]
            ref = P[ant[-1]] if ant else None
        if ref:
            fijos = sum(v for _, v in ref['fijos'])
            d['sh'] = round(d['gr'] * ref['alic'] + fijos, 2)
            d['sh_det'] = f"ventas gravadas {pesos(d['gr'])} x {ref['alic'] * 100:.4g}%" + (f" + fijos {pesos(fijos)}" if fijos else '') + \
                          f" (alícuota de la hoja {ref['hoja']} del papel)"
            d['fuentes']['sh'] = (ref['f'], f"papel de trabajo, hoja {ref['hoja']}")
        else:
            ant = [p for p in sorted(F.tasas_sh) if p < per]
            if ant:
                t = F.tasas_sh[ant[-1]]
                d['sh'] = round(d['gr'] * t, 2)
                d['sh_det'] = (f"ESTIMADO: ventas gravadas {pesos(d['gr'])} x {t * 100:.4g}% (tasa efectiva de {mm(ant[-1])} en el EERR: "
                               f"el papel de trabajo no tiene la alícuota de la tasa)")
                d['fuentes']['sh'] = ('', f"estimado con la tasa efectiva de {mm(ant[-1])} del EERR")
    if P.get(per) and P[per].get('pyp') is not None:
        d['pyp'] = round(P[per]['pyp'], 2); d['fuentes']['pyp'] = (P[per]['f'], f"papel de trabajo, hoja {P[per]['hoja']}")
    # prorrateo
    if 'compras' in d and 'gr' in d:
        tot = d['gr'] + d['ex']
        d['pct_ex'] = d['ex'] / tot if tot else 0.0
        d['prorrateo'] = round(sum(d['iva_serv'].values()) * d['pct_ex'], 2)
    return d


# ------------------------------------------------------------------ EERR
def filas_eerr(e):
    hdr = next(r for r in range(1, 15) if norm(e.cell(r, 1).value) == 'concepto')
    col = {}
    for c in range(2, 30):
        v = e.cell(hdr, c).value
        if hasattr(v, 'year'): col[f"{v.year}{v.month:02d}"] = c
    filas = {}
    costo = gastos = None
    for r in range(hdr + 1, e.max_row + 1):
        l = norm(e.cell(r, 1).value)
        if not l: continue
        if l == 'costo': costo = r
        if l == 'gastos': gastos = r
        if l == 'cmv' and 'cmv' in filas: continue          # la segunda "CMV" es la de existencias
        filas.setdefault(l, r)
    return hdr, col, filas


def tasas_sh_de(f):
    # S&H / ventas gravadas de cada mes del EERR (valores calculados), para estimar la tasa cuando el papel no tiene alícuota
    out = {}
    try: e = abrir(f, data_only=True)['EERR']
    except Exception: return out
    _, col, fl = filas_eerr(e)
    rs, rg = fl.get('seguridad e higiene'), fl.get('ventas gravadas')
    if not rs or not rg: return out
    for p, c in col.items():
        sh, gr = e.cell(rs, c).value, e.cell(rg, c).value
        if num(sh) and num(gr) and sh > 0 and gr > 0: out[p] = sh / gr
    return out


def vacia(c):
    # vacía = sin dato o con una fórmula que apunta a las hojas mensuales de liquidación (las dejó la plantilla)
    v = c.value
    if v is None or v == '': return True
    return isinstance(v, str) and v.startswith('=') and re.fullmatch(r"=\+?'\d{2}-\d{4}'![A-Z]+\d+", v.replace(' ', '')) is not None


def fin_cuerpo(ws, primera=2):
    # última fila con datos antes del renglón en blanco + fila de totales
    r = primera
    while r <= ws.max_row and any(ws.cell(r, c).value not in (None, '') for c in range(1, 4)):
        r += 1
    return r - 1


def reescribir_total(ws, cuerpo_fin, etiqueta, col_lab, cols, max_col):
    # borra lo que haya debajo del cuerpo y pone renglón en blanco + totales
    for r in range(cuerpo_fin + 1, ws.max_row + 1):
        for c in range(1, max_col + 1): ws.cell(r, c).value = None
    t = cuerpo_fin + 2
    ws.cell(t, col_lab).value = etiqueta
    for c in cols: ws.cell(t, c).value = f"=SUM({CL(c)}2:{CL(c)}{cuerpo_fin})"; ws.cell(t, c).number_format = FMT
    for c in range(1, max_col + 1): ws.cell(t, c).font = Font(bold=True)


def hoja_compras(wb, meses):
    if 'Compras imputadas' in wb.sheetnames: return wb['Compras imputadas']
    ws = wb.create_sheet('Compras imputadas')
    hdr = ['CUIT', 'Proveedor', 'Cuenta Maestro', 'Rubro EERR', 'Estado Maestro'] + [mm(p) for p in meses] + ['Total']
    ws.append(hdr)
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = VERDE; c.alignment = Alignment(horizontal='center', wrap_text=True)
    for col, w in zip('ABCDE', (13, 42, 30, 30, 17)): ws.column_dimensions[col].width = w
    for j in range(6, 19): ws.column_dimensions[CL(j)].width = 15
    ws.freeze_panes = 'F2'
    reescribir_total(ws, 1, 'TOTAL COMPRAS PORTAL IVA', 2, range(6, 19), 18)
    return ws


def cargar_compras(ws, per, d):
    cols = {nfc(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
    c = cols.get(mm(per))
    if c is None: return 'el mes no tiene columna en la hoja'
    fin = fin_cuerpo(ws)
    filas = {}
    for r in range(2, fin + 1):
        filas[(ndoc(ws.cell(r, 1).value), nfc(ws.cell(r, 4).value))] = r
    if any(num(ws.cell(r, c).value) for r in range(2, fin + 1)): return 'ya tenía datos'
    for r in range(2, fin + 1): ws.cell(r, c).value = None
    for r in range(fin + 1, ws.max_row + 1):          # renglón en blanco y totales: se rehacen al final
        for j in range(1, ws.max_column + 1): ws.cell(r, j).value = None; ws.cell(r, j).font = Font()
    tot_col = cols.get('Total')
    for k, v in sorted(d['compras'].items(), key=lambda kv: (kv[0][2].startswith('EXCLUIDO'), kv[0][2], -kv[1])):
        if abs(v) <= 0.004: continue
        r = filas.get((k[0], k[2]))
        if r is None:
            fin += 1; r = fin; filas[(k[0], k[2])] = r
            for j, val in enumerate([k[0], d['compras_info'][k], k[1], k[2], k[3]], 1): ws.cell(r, j).value = val
            if tot_col: ws.cell(r, tot_col).value = f"=SUM(F{r}:{CL(tot_col - 1)}{r})"; ws.cell(r, tot_col).number_format = FMT
            for j in range(6, (tot_col or 18)): ws.cell(r, j).number_format = FMT
        ws.cell(r, c).value = round((ws.cell(r, c).value or 0) + v, 2); ws.cell(r, c).number_format = FMT
    reescribir_total(ws, fin, 'TOTAL COMPRAS PORTAL IVA', 2, range(6, (tot_col or 18) + 1), tot_col or 18)
    ws.auto_filter.ref = f"A1:{CL(tot_col or 18)}{fin}"
    return 'cargado'


def hoja_ventas(wb):
    if 'Ventas Portal IVA' in wb.sheetnames: return wb['Ventas Portal IVA']
    ws = wb.create_sheet('Ventas Portal IVA')
    ws.append(['Mes', 'Tipo de comprobante', 'Comprobantes', 'Neto gravado', 'No gravado', 'Exento', 'IVA', 'Total'])
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = VERDE; c.alignment = Alignment(horizontal='center', wrap_text=True)
    for col, w in zip('ABCDEFGH', (10, 34, 13, 18, 18, 18, 18, 18)): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    reescribir_total(ws, 1, 'TOTAL VENTAS PORTAL IVA', 2, range(3, 9), 8)
    return ws


def cargar_ventas(ws, per, d, pisar):
    fin = fin_cuerpo(ws)
    tiene = [r for r in range(2, fin + 1) if ws.cell(r, 1).value == mm(per)]
    if tiene and not pisar: return 'ya tenía datos'
    filas = [[ws.cell(r, c).value for c in range(1, 9)] for r in range(2, fin + 1) if ws.cell(r, 1).value != mm(per)]
    for t in sorted(d['ventas'], key=lambda t: (t is None, t or 0)):
        v = d['ventas'][t]
        filas.append([mm(per), f"{t} - {TIPOS.get(t, 'Otro')}", int(v[5])] + [round(x, 2) for x in v[:5]])
    filas.sort(key=lambda f: (str(f[0])[3:] + str(f[0])[:2]))
    for r in range(2, ws.max_row + 1):
        for c in range(1, 9): ws.cell(r, c).value = None; ws.cell(r, c).font = Font()
    for i, f in enumerate(filas, 2):
        for j, v in enumerate(f, 1):
            ws.cell(i, j).value = v
            if j >= 4: ws.cell(i, j).number_format = FMT
    fin = len(filas) + 1
    reescribir_total(ws, fin, 'TOTAL VENTAS PORTAL IVA', 2, range(3, 9), 8)
    ws.auto_filter.ref = f"A1:H{fin}"
    return 'cargado'


IP_HDR = ['Mes', 'IVA servicios 2,5%', 'IVA servicios 5%', 'IVA servicios 10,5%', 'IVA servicios 21%', 'IVA servicios 27%', 'IVA servicios otros',
          'Total IVA servicios', 'Ventas gravadas', 'Ventas exentas y no gravadas', '% venta exenta', 'IVA no computable',
          'Control: papel de trabajo', 'Diferencia', 'IVA mercadería (computable)', 'IVA bienes de uso / bloqueados (fuera)', 'Observación']
def hoja_prorrateo(wb):
    if 'IVA prorrateo' in wb.sheetnames: return wb['IVA prorrateo']
    ws = wb.create_sheet('IVA prorrateo')
    ws['A1'] = 'IVA no computable por prorrateo'; ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = ('IVA de compras que no son mercadería (rubro distinto de CMV en "Compras imputadas"; sin bienes de uso ni bloqueados) x '
                'proporción de ventas exentas y no gravadas del mes (Portal IVA). Mismo método que el papel de trabajo mensual.')
    for j, h in enumerate(IP_HDR, 1):
        c = ws.cell(4, j, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = VERDE; c.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.cell(5, 1, 'Total')
    ws.column_dimensions['A'].width = 10
    for j in range(2, 18): ws.column_dimensions[CL(j)].width = 16
    return ws


def cargar_prorrateo(ws, per, d, pisar):
    # devuelve la fila del mes en la hoja (el EERR la referencia)
    r_tot = next((r for r in range(5, ws.max_row + 1) if norm(ws.cell(r, 1).value) == 'total'), None)
    if r_tot is None: r_tot = ws.max_row + 1; ws.cell(r_tot, 1, 'Total')
    r = next((r for r in range(5, r_tot) if ws.cell(r, 1).value == mm(per)), None)
    if r is not None and not pisar: return r, 'ya tenía datos'
    if r is None:
        r = r_tot; r_tot += 1
        for c in range(1, 18): ws.cell(r_tot, c).value = ws.cell(r, c).value; ws.cell(r_tot, c).font = Font(bold=True)
        for c in range(1, 18): ws.cell(r, c).value = None; ws.cell(r, c).font = Font()
    s = d['iva_serv']
    vals = [mm(per)] + [round(s.get(a, 0.0), 2) for a in ALICS] + [round(s.get('otros', 0.0), 2), f'=SUM(B{r}:G{r})',
            round(d['gr'], 2), round(d['ex'], 2), f'=IF(I{r}+J{r}=0,0,J{r}/(I{r}+J{r}))', f'=H{r}*K{r}', None, None,
            round(d['iva_merc'], 2), round(d['iva_fuera'], 2), None]
    for j, v in enumerate(vals, 1):
        ws.cell(r, j).value = v
        if j > 1: ws.cell(r, j).number_format = '0.00%' if j == 11 else FMT
    for j in (2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16):
        ws.cell(r_tot, j).value = f'=SUM({CL(j)}5:{CL(j)}{r_tot - 1})'; ws.cell(r_tot, j).number_format = FMT
    ws.cell(r_tot, 1).value = 'Total'
    return r, 'cargado'


def falta(e, r, c):
    return r is not None and vacia(e.cell(r, c))


def hay_sueldos(wb, per):
    if 'Sueldos' not in wb.sheetnames: return True
    su = wb['Sueldos']
    for r in range(9, 21):
        v = su.cell(r, 2).value
        if hasattr(v, 'year') and f"{v.year}{v.month:02d}" == per: return num(su.cell(r, 3).value)
    return True


def poner(e, r, c, valor, fuente, log, per, concepto, pisar, comentario=None, mostrar=None):
    # escribe la celda solo si está vacía (o si el mes se pidió con --pisar); lo que ya tenía no se registra
    if r is None: log.append((per, concepto, 'sin fila en el EERR', None, '')); return
    cel = e.cell(r, c)
    if not vacia(cel) and not pisar: return
    anterior = cel.value
    cel.value = valor; cel.number_format = FMT
    if comentario or (pisar and anterior not in (None, '') and anterior != valor):
        txt = (comentario or '') + (f"\nValor anterior: {anterior}" if pisar and anterior not in (None, '') and anterior != valor else '')
        cel.comment = Comment(txt.strip(), 'Armado EERR')
    log.append((per, concepto, 'cargado', mostrar if mostrar is not None else (valor if num(valor) else None), fuente))


def armar_mes(wb, F, per, pisar, log):
    e = wb['EERR']; hdr, col, fl = filas_eerr(e)
    if per not in col: log.append((per, '-', 'el mes no es del ejercicio de este EERR', None, '')); return
    c = col[per]
    d = datos_mes(F, per)
    nombre = lambda k: os.path.basename(d['fuentes'][k][0]) + ' — ' + d['fuentes'][k][1] if k in d['fuentes'] else ''
    # ventas
    if 'ventas' in d:
        vs = hoja_ventas(wb); est = cargar_ventas(vs, per, d, pisar)
        V = "'Ventas Portal IVA'!"; m = f'"{mm(per)}"'
        poner(e, fl.get('ventas exentas y no gravadas'), c, f"=SUMIFS({V}E:E,{V}$A:$A,{m})+SUMIFS({V}F:F,{V}$A:$A,{m})", nombre('ventas'), log, per,
              'Ventas exentas y no gravadas', pisar, mostrar=round(d['ex'], 2))
        poner(e, fl.get('ventas gravadas'), c, f"=SUMIFS({V}D:D,{V}$A:$A,{m})", nombre('ventas'), log, per, 'Ventas gravadas', pisar,
              mostrar=round(d['gr'], 2))
    else:
        if falta(e, fl.get('ventas gravadas'), c): log.append((per, 'Ventas', 'sin fuente: falta el libro de ventas del Portal IVA', None, ''))
    # compras
    if 'compras' in d:
        ci = wb['Compras imputadas'] if 'Compras imputadas' in wb.sheetnames else hoja_compras(wb, sorted(col, key=lambda p: col[p]))
        est = cargar_compras(ci, per, d) if not pisar else None
        if pisar:
            # rehacer la columna del mes: se vacía y se vuelve a cargar
            cols = {nfc(ci.cell(1, j).value): j for j in range(1, ci.max_column + 1) if ci.cell(1, j).value}
            j = cols.get(mm(per))
            if j:
                for r in range(2, fin_cuerpo(ci) + 1): ci.cell(r, j).value = None
            est = cargar_compras(ci, per, d)
        cc = {nfc(ci.cell(1, j).value): j for j in range(1, ci.max_column + 1) if ci.cell(1, j).value}
        letra = CL(cc[mm(per)]) if mm(per) in cc else None
        if letra:
            CI = "'Compras imputadas'!"
            rubros = sorted(set(RUBRO.values()) | {'GASTOS VARIOS'})
            for rub in rubros:
                if rub == 'Publicidad': continue
                r = fl.get('cmv') if rub == 'CMV' else fl.get(norm(rub))
                if r is None:
                    if any(k[2] == rub for k in d['compras']): log.append((per, rub, 'sin fila en el EERR para este rubro (queda solo en Compras imputadas)', None, ''))
                    continue
                crit = '"CMV"' if rub == 'CMV' else f'$A${r}'
                cel = e.cell(r, c)
                if vacia(cel) or pisar:
                    cel.value = f"=SUMIFS({CI}{letra}:{letra},{CI}$D:$D,{crit})"; cel.number_format = FMT
            if est == 'cargado':
                log.append((per, 'Compras (CMV y gastos)', 'cargado',
                            round(sum(v for k, v in d['compras'].items() if not k[2].startswith('EXCLUIDO')), 2), nombre('compras')))
                if not d['fuentes']['compras'][1].startswith('libro de la DDJJ presentada'):
                    log.append((per, 'Libro de compras', 'aviso: no es el libro de la DDJJ presentada; puede faltar lo que el equipo carga a mano '
                                '(Edesur, AySA, Metrogas). Bajarlo con portal-iva-descarga --presentada y correr con --pisar', None, nombre('compras')))
                sin = [(k, v) for k, v in d['compras'].items() if k[3] in ('Sin maestro', 'Revisar') or 'sin fila' in k[3]]
                for k, v in sorted(sin, key=lambda kv: -abs(kv[1]))[:15]:
                    log.append((per, 'Proveedor a revisar', f"{k[3]}: {d['compras_info'][k]} (CUIT {k[0]}) va a {k[2]}", round(v, 2), ''))
            # publicidad = PyP del papel + compras imputadas a Publicidad
            r = fl.get('publicidad')
            if r:
                pyp = d.get('pyp')
                f = f"={pyp or 0}+SUMIFS({CI}{letra}:{letra},{CI}$D:$D,\"Publicidad\")"
                poner(e, r, c, f, nombre('pyp') or 'compras imputadas a Publicidad (sin PyP en el papel del mes)', log, per, 'Publicidad', pisar,
                      mostrar=pyp)
        # IVA prorrateo
        if 'prorrateo' in d:
            ip = hoja_prorrateo(wb); rr, est = cargar_prorrateo(ip, per, d, pisar)
            poner(e, fl.get('iva prorrateo'), c, f"='IVA prorrateo'!L{rr}", 'hoja IVA prorrateo (compras + ventas del Portal IVA)', log, per,
                  'IVA Prorrateo', pisar, mostrar=d['prorrateo'])
    else:
        if falta(e, fl.get('cmv'), c): log.append((per, 'Compras', 'sin fuente: falta el libro de compras del Portal IVA', None, ''))
    # IIBB
    if 'iibb' in d:
        poner(e, fl.get('iibb'), c, round(d['iibb'], 2), nombre('iibb'), log, per, 'IIBB', pisar,
              comentario=f"Impuesto determinado de la DJ de ARBA: {os.path.basename(d['fuentes']['iibb'][0])}")
    else:
        if falta(e, fl.get('iibb'), c): log.append((per, 'IIBB', 'sin fuente: falta la DJ de anticipo de ARBA (constancias-anticipo-arba; AGIP y Convenio Multilateral se cargan a mano)', None, ''))
    # Seguridad e Higiene
    if 'sh' in d:
        poner(e, fl.get('seguridad e higiene'), c, d['sh'], nombre('sh'), log, per, 'Seguridad e Higiene', pisar, comentario=d['sh_det'])
    else:
        # sociedades sin tasa (CABA) o que nunca la tuvieron cargada: no se avisa
        tuvo_tasa = bool(F.tasas_sh) or any(num(x.get('alic')) for x in F.papel().values())
        if falta(e, fl.get('seguridad e higiene'), c) and 'gr' in d and tuvo_tasa: log.append((per, 'Seguridad e Higiene', 'sin fuente: no hay alícuota de la tasa en el papel de trabajo ni meses anteriores con la tasa', None, ''))
    # Sueldos (hoja Sueldos, la fila del EERR ya la referencia)
    if 'f931' in d and 'Sueldos' in wb.sheetnames:
        su = wb['Sueldos']; f9 = d['f931']
        r = next((r for r in range(9, 21) if hasattr(su.cell(r, 2).value, 'year') and f"{su.cell(r, 2).value.year}{su.cell(r, 2).value.month:02d}" == per), None)
        if r is None: log.append((per, 'Sueldos', 'el mes no está en la hoja Sueldos', None, ''))
        elif num(su.cell(r, 3).value) and not pisar: pass
        else:
            for colx, k in zip('CDEFGIJK', ('c301', 'c302', 'c351', 'c352', 'c312', 'empleados', 'rem_art', 'c028')):
                su[f'{colx}{r}'].value = f9[k]
                if colx != 'I': su[f'{colx}{r}'].number_format = '#,##0.00'
            j = r - 9; cl = CL(3 + j)
            if su[f'{cl}23'].value in (None, ''): su[f'{cl}23'].value = f'=+J{r}'
            if su[f'{cl}24'].value in (None, ''): su[f'{cl}24'].value = f'=+E{r}+F{r}+G{r}+K{r}'
            log.append((per, 'Sueldos', 'cargado', round((f9['rem_art'] or 0) + sum(f9[k] or 0 for k in ('c351', 'c352', 'c312', 'c028')), 2), nombre('f931')))
        rs = fl.get('sueldos y cargas sociales')
        if rs and r and vacia(e.cell(rs, c)):
            j = r - 9; e.cell(rs, c).value = f"=+Sueldos!{CL(3 + j)}23+Sueldos!{CL(3 + j)}24"; e.cell(rs, c).number_format = FMT
    else:
        rs = fl.get('sueldos y cargas sociales')
        if not hay_sueldos(wb, per): log.append((per, 'Sueldos', 'sin fuente: falta el F.931 del mes (f931-descarga-arca)', None, ''))
    return d


def registrar(wb, log, hasta, modo):
    ws = wb['Armado'] if 'Armado' in wb.sheetnames else wb.create_sheet('Armado')
    if ws.max_row <= 1 and ws['A1'].value is None:
        ws.append(['Fecha', 'Modo', 'Hasta', 'Mes', 'Concepto', 'Acción', 'Valor', 'Fuente'])
        for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = VERDE
        for col, w in zip('ABCDEFGH', (17, 11, 9, 9, 28, 60, 18, 80)): ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A2'
    hoy = datetime.datetime.now().replace(microsecond=0, second=0)
    for per, conc, acc, val, fuente in log:
        ws.append([hoy, modo, mm(hasta), mm(per) if per and per != '-' else '-', conc, acc, val, fuente])
        r = ws.max_row; ws.cell(r, 1).number_format = 'dd/mm/yyyy hh:mm'; ws.cell(r, 7).number_format = FMT
        if acc.startswith('sin fuente') or 'revisar' in conc.lower() or acc.startswith('sin fila'):
            for c in range(1, 9): ws.cell(r, c).fill = AMARILLO


# ------------------------------------------------------------------ crear
def ejercicio_desde(cuit, desde=None):
    if desde: return desde
    prev = eerr_de(cuit)
    if not prev: return None
    m = re.search(r'EERR (\d{2})-(\d{4}) a (\d{2})-(\d{4})', os.path.basename(prev[-1]))
    return mas_mes(f"{m[4]}{m[3]}") if m else None


def eerr_de(cuit):
    out = []
    for f in glob.glob(os.path.join(EERR_DIR, 'EERR *.xlsx')):
        if os.path.basename(f).startswith('~$'): continue
        try: e = abrir(f, read_only=True)['EERR']
        except Exception: continue
        for row in e.iter_rows(min_row=1, max_row=5, max_col=10, values_only=True):
            if any(len(re.sub(r'\D', '', str(v or ''))) == 11 and re.sub(r'\D', '', str(v)) == cuit for v in row):
                out.append(f); break
    def clave(f):
        m = re.search(r'EERR (\d{2})-(\d{4}) a', os.path.basename(f)); return f"{m[2]}{m[1]}" if m else ''
    return sorted(out, key=clave)


def crear(cuit, desde, wbprev=None):
    meses = [mas_mes(desde, i) for i in range(12)]
    wb = abrir(PLANTILLA)
    cuit_fmt = f"{cuit[:2]}-{cuit[2:10]}-{cuit[10]}"
    razon = None
    if wbprev is not None:
        razon = wbprev['EERR']['E2'].value
    razon = razon or cortos.get(cuit, cuit)
    e = wb['EERR']; e['E2'] = razon; e['E3'] = cuit_fmt
    hdr, _, _ = filas_eerr(e)
    for i, p in enumerate(meses): e.cell(hdr, 2 + i).value = fecha(p)
    su = wb['Sueldos']; su['F2'] = razon; su['F3'] = cuit_fmt
    for i, p in enumerate(meses): su.cell(9 + i, 2).value = fecha(p)
    hoja_compras(wb, meses)
    nombre = f"EERR {hoja_mes(meses[0])} a {hoja_mes(meses[-1])} - {cortos.get(cuit, cuit)}.xlsx"
    return wb, nombre, meses


# ------------------------------------------------------------------ Excel
PS = r"""
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($env:ARMADO_XLSX, 0)
  $xl.CalculateFull(); $wb.Save(); $wb.Close($false)
} finally { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
"""
OSA = """with timeout of 600 seconds
tell application "Microsoft Excel"
  set display alerts to false
  set wb to open workbook workbook file name (POSIX file "{f}" as text) update links do not update links
  calculate full
  save wb
  close wb saving no
end tell
end timeout"""


def recalcular(f):
    f = os.path.abspath(f)
    err = ''
    for i in range(3):
        try:
            if WIN:
                r = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', PS], capture_output=True, text=True,
                                   timeout=900, env={**os.environ, 'ARMADO_XLSX': f})
            elif MAC:
                r = subprocess.run(['osascript', '-e', OSA.format(f=f.replace('"', '\\"'))], capture_output=True, text=True, timeout=900)
            else:
                print('  ! sin Excel en este sistema: el archivo queda sin valores calculados'); return False
            if r.returncode == 0: return True
            err = (r.stderr or r.stdout).strip()
        except Exception as ex: err = str(ex)
        time.sleep(10)
    print('  ! Excel no pudo recalcular', os.path.basename(f), err[:200]); return False


# ------------------------------------------------------------------ main
def cuit_de_eerr(f):
    e = abrir(f, read_only=True)['EERR']
    for row in e.iter_rows(min_row=1, max_row=5, max_col=10, values_only=True):
        for v in row:
            d = re.sub(r'\D', '', str(v or ''))
            if len(d) == 11: return d
    return None


def elegir(a):
    if 'cuit' in a: return [re.sub(r'\D', '', a['cuit'])]
    if 'nombre' in a:
        q = norm(a['nombre'])
        c = [k for k, v in cortos.items() if norm(v) == q] or [k for k, v in cortos.items() if q in norm(v)]
        if len(c) > 1: sys.exit('Hay varias sociedades con ese nombre: ' + '; '.join(cortos[k] for k in c) + '. Usá --cuit.')
        if not c: sys.exit(f'No encuentro la sociedad "{a["nombre"]}" en la unidad.')
        return c
    if 'todas' in a:
        return sorted({cuit_de_eerr(f) for f in glob.glob(os.path.join(EERR_DIR, 'EERR *.xlsx')) if not os.path.basename(f).startswith('~$')} - {None})
    sys.exit('Uso: armar_eerr.py --nombre="Sociedad" | --cuit=CUIT | --todas  [--crear [--desde=AAAAMM]] [--hasta=AAAAMM] [--pisar=AAAAMM,...] [--simular]')


def main():
    try: sys.stdout.reconfigure(encoding='utf-8')
    except Exception: pass
    a = {k: v for k, _, v in (s.lstrip('-').partition('=') for s in sys.argv[1:])}
    hasta = a.get('hasta') or hasta_default()
    simular = 'simular' in a
    pisar = set(x.strip() for x in a.get('pisar', '').split(',') if x.strip())
    modo = 'crear' if 'crear' in a else 'actualizar'
    if not os.path.exists(MAESTRO): sys.exit(f'No encuentro el Maestro de Proveedores en {MAESTRO}')
    trabajo = os.path.join(SALIDA, 'trabajo'); os.makedirs(trabajo, exist_ok=True)
    hoy = datetime.datetime.now()
    resp = os.path.join(EERR_DIR, '_respaldos', f"{hoy:%Y%m%d-%H%M}")
    print(f"Armado de EERR ({modo}), meses hasta {mm(hasta)}{' — simulación' if simular else ''}")
    resumen = []
    for cuit in elegir(a):
        if cuit not in carpetas: print(f"  {cuit}: sin carpeta en la unidad, se saltea"); continue
        soc = cortos[cuit]; F = Fuentes(cuit); log = []
        previos = eerr_de(cuit)
        for f in previos[-2:]: F.tasas_sh.update(tasas_sh_de(f))
        if modo == 'crear':
            desde = ejercicio_desde(cuit, a.get('desde'))
            if not desde: print(f"  {soc}: no tiene EERR anterior; indicá el primer mes con --desde=AAAAMM"); continue
            if desde > hasta and 'desde' not in a:
                if 'todas' in a: continue
                print(f"  {soc}: el ejercicio actual ({mm(mas_mes(desde, -12))} a {mm(mas_mes(desde, -1))}) todavía no cerró; se actualiza con el modo normal"); continue
            wbprev = abrir(previos[-1], read_only=True) if previos else None
            wb, nombre, meses = crear(cuit, desde, wbprev)
            destino = os.path.join(EERR_DIR, nombre)
            if os.path.exists(destino): print(f"  {soc}: ya existe {nombre}; usá el modo actualizar"); continue
            mt = None
        else:
            if not previos: print(f"  {soc}: no tiene EERR; usá --crear --desde=AAAAMM"); continue
            destino = previos[-1]; nombre = os.path.basename(destino); mt = os.path.getmtime(destino)
            w = os.path.join(trabajo, nombre); shutil.copy2(destino, w)
            wb = abrir(w)
            _, col, _ = filas_eerr(wb['EERR']); meses = sorted(col)
            if hasta > meses[-1]:
                log.append(('-', 'Ejercicio', f"el ejercicio cerró en {mm(meses[-1])}: para los meses siguientes, correr con --crear", None, ''))
        if modo == 'crear': aplicar = [p for p in meses if p <= hasta]
        else:
            # por defecto los últimos 3 meses hasta --hasta (lo que llega tarde: F.931, DJ, libro presentado); --desde amplía
            desde_v = a.get('desde') or mas_mes(hasta, -2)
            aplicar = [p for p in meses if (desde_v <= p <= hasta) or p in pisar]
        for p in aplicar:
            try: armar_mes(wb, F, p, p in pisar, log)
            except Exception as ex:
                log.append((p, '-', f'error: {type(ex).__name__}: {ex}', None, ''))
        registrar(wb, log, hasta, modo)
        w = os.path.join(trabajo, nombre)
        wb.save(w); recalcular(w)
        if simular:
            os.makedirs(os.path.join(SALIDA, 'simulacion'), exist_ok=True); shutil.copy2(w, os.path.join(SALIDA, 'simulacion', nombre)); estado = 'simulación'
        elif mt is not None and os.path.getmtime(destino) != mt:
            os.makedirs(os.path.join(SALIDA, 'no-guardados'), exist_ok=True); shutil.copy2(w, os.path.join(SALIDA, 'no-guardados', nombre))
            estado = 'NO GUARDADO: el EERR cambió en la unidad mientras se armaba'
        else:
            if mt is not None:
                os.makedirs(resp, exist_ok=True); shutil.copy2(destino, os.path.join(resp, nombre))
            shutil.copy2(w, destino); estado = 'guardado' if mt is not None else 'creado'
        cargados = sum(1 for x in log if x[2] == 'cargado')
        faltan = [x for x in log if x[2].startswith('sin fuente')]
        revisar = [x for x in log if x[1] == 'Proveedor a revisar']
        errores = [x for x in log if x[2].startswith('error')]
        print(f"  {soc:28} {nombre}: {cargados} conceptos cargados, {len(faltan)} sin fuente, {len(revisar)} proveedores a revisar — {estado}")
        for x in faltan + errores: print(f"     {mm(x[0]) if x[0] != '-' else '-'} {x[1]}: {x[2]}")
        for x in log:
            if x[1] == 'Ejercicio': print('     ' + x[2])
        resumen.append(dict(sociedad=soc, archivo=nombre, estado=estado, cargados=cargados,
                            sin_fuente=[f"{mm(x[0])} {x[1]}" for x in faltan], revisar=len(revisar), errores=[x[2] for x in errores]))
    os.makedirs(SALIDA, exist_ok=True)
    json.dump(resumen, open(os.path.join(SALIDA, 'ultima-ejecucion.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
