#!/usr/bin/env python3
"""
Portal IVA -> inputs COMPRAS / VENTAS con clasificacion de proveedores, sobre la unidad compartida.

Envuelve los scripts de la skill original (procesar_portal_iva.py y agregar_clasificaciones.py) para
correrlos en Claude Code: busca la carpeta PORTAL IVA de la sociedad, reconoce los zips, las plantillas
y el padron, calcula los controles en Python (sin LibreOffice) y deja los resultados en la misma carpeta
sin pisar nada.

  procesar   python portal-iva-vulcamoia.py procesar [--sociedad=Vulcamoia] [--anio=2026] [--solo-revisar]
                                                       [--carpeta=<ruta>] [--compras=..] [--ventas=..] [--padron=..]
                                                       [--si-existe=error|reemplazar]
  clasificar python portal-iva-vulcamoia.py clasificar --json=<sugerencias.json> --modo=sugerida|confirmada
                                                       [--sociedad / --carpeta / --padron como arriba]

Salida de procesar (en la carpeta PORTAL IVA):
  input_portal_COMPRAS Completo.xlsx  PORTAL IVA + Control con Portal IVA + Resumen por Clasificacion
                                      + Proveedores sin clasificar
  input_portal_VENTAS Completo.xlsx   PORTAL IVA completo + Comprobantes Ventas
El padron NO se toca al procesar (las variantes de escritura se unifican en memoria). Solo
`clasificar --modo=confirmada` le agrega filas, despues de guardar una copia en _anteriores/.
Por stdout: JSON con conteos, controles por mes, sin clasificar y avisos.
"""
import argparse, datetime as dt, glob, json, os, re, shutil, sys, tempfile, unicodedata

AQUI = os.path.dirname(os.path.abspath(__file__))
PLUGIN = os.path.dirname(AQUI)
sys.path.insert(0, AQUI)

RAIZ_DEFECTO = r'I:\Unidades compartidas\BVA - Sociedades Farmaceuticas'
SOCIEDAD_DEFECTO = 'Vulcamoia'
SALIDA_COMPRAS = 'input_portal_COMPRAS Completo.xlsx'
SALIDA_VENTAS = 'input_portal_VENTAS Completo.xlsx'
TOL = 1.0   # diferencias de hasta $1 son redondeo de ARCA


def log(m):
    print(m, file=sys.stderr, flush=True)


def fail(msg, **extra):
    print(json.dumps(dict(ok=False, error=msg, **extra), ensure_ascii=False, indent=1))
    sys.exit(1)


def cargar_env():
    for ruta in (os.environ.get('FISCO_ENV_FILE'), os.path.join(PLUGIN, '.env'),
                 os.path.join(os.path.expanduser('~'), '.fisco-ar', '.env')):
        if not ruta or not os.path.isfile(ruta):
            continue
        for linea in open(ruta, encoding='utf-8', errors='replace'):
            m = re.match(r'^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$', linea)
            if not m or linea.lstrip().startswith('#'):
                continue
            v = m.group(2)
            if len(v) >= 2 and v[0] == v[-1] and v[0] in '"\'':
                v = v[1:-1]
            if v and not os.environ.get(m.group(1)):
                os.environ[m.group(1)] = v


def norm(s):
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


# ── carpeta y archivos ──────────────────────────────────────────────────────
def carpeta_portal(a):
    if a.carpeta:
        if not os.path.isdir(a.carpeta):
            fail(f'No existe la carpeta "{a.carpeta}".')
        return a.carpeta
    raiz = (os.environ.get('BVA_UNIDAD_PATH') or '').strip() or RAIZ_DEFECTO
    if not os.path.isdir(raiz):
        fail(f'No encuentro la unidad compartida en "{raiz}". Defini BVA_UNIDAD_PATH en el .env o pasa --carpeta.')
    q = norm(a.sociedad or SOCIEDAD_DEFECTO)
    socs = [d for d in os.listdir(raiz) if os.path.isdir(os.path.join(raiz, d))
            and (q in norm(d) or re.sub(r'\D', '', q) and re.sub(r'\D', '', q) in re.sub(r'\D', '', d))]
    if len(socs) != 1:
        fail(f'La sociedad "{a.sociedad or SOCIEDAD_DEFECTO}" coincide con {len(socs)} carpetas de la unidad.',
             carpetas=sorted(socs))
    base = os.path.join(raiz, socs[0], '04-Balances y EECC')
    if a.anio:
        anios = [str(a.anio)]
    else:
        anios = sorted((d for d in (os.listdir(base) if os.path.isdir(base) else []) if re.fullmatch(r'20\d\d', d)
                        and os.path.isdir(os.path.join(base, d, 'PORTAL IVA'))), reverse=True)
    for an in anios:
        c = os.path.join(base, an, 'PORTAL IVA')
        if os.path.isdir(c):
            return c
    fail(f'No encuentro la carpeta "04-Balances y EECC/<año>/PORTAL IVA" de {socs[0]}.'
         + (f' (año {a.anio})' if a.anio else ''))


def es_resultado(nombre):
    n = norm(nombre)
    return 'completo' in n or n.startswith('copia')


def plantilla(carpeta, libro, hoja_agregada, explicita):
    import openpyxl
    if explicita:
        return explicita
    cands = [f for f in glob.glob(os.path.join(carpeta, '*.xlsx'))
             if re.search(r'input.?portal', os.path.basename(f), re.I) and libro in os.path.basename(f).upper()
             and not es_resultado(os.path.basename(f)) and not os.path.basename(f).startswith('~$')]
    vacias = []
    for f in cands:
        wb = openpyxl.load_workbook(f, read_only=True)
        if 'PORTAL IVA' in wb.sheetnames and hoja_agregada not in wb.sheetnames:
            vacias.append(f)
        wb.close()
    if len(vacias) != 1:
        fail(f'No encuentro una única plantilla de {libro} (input_portal {libro}.xlsx sin procesar) en la carpeta. '
             f'Pasala con --{libro.lower()}=<ruta>.', candidatas=[os.path.basename(f) for f in cands])
    return vacias[0]


def padron(carpeta, explicito, avisos):
    if explicito:
        return explicito
    cands = [f for f in glob.glob(os.path.join(carpeta, '*.xlsx'))
             if 'clasificacion' in norm(os.path.basename(f)) and 'proveedor' in norm(os.path.basename(f))
             and not os.path.basename(f).startswith('~$')]
    if not cands:
        fail('No encuentro el padrón (Clasificacion Proveedores ... .xlsx) en la carpeta. Pasalo con --padron.')
    cands.sort(key=os.path.getmtime, reverse=True)
    if len(cands) > 1:
        avisos.append(f'Hay {len(cands)} padrones en la carpeta; se usó el más reciente: {os.path.basename(cands[0])} '
                      f'(los otros: {", ".join(os.path.basename(f) for f in cands[1:])}).')
    return cands[0]


def zips_por_periodo(carpeta, avisos):
    por = {}
    for f in glob.glob(os.path.join(carpeta, '*.zip')):
        m = re.search(r'libro_iva_periodo_(20\d{2})(\d{2})', os.path.basename(f), re.I)
        if m:
            por.setdefault(m.group(1) + m.group(2), []).append(f)
    if not por:
        fail('No hay zips libro_iva_periodo_AAAAMM*.zip en la carpeta. Bajalos del Portal IVA (Libro IVA Digital).')
    elegidos = {}
    for p, fs in sorted(por.items()):
        fs.sort(key=os.path.getmtime, reverse=True)
        elegidos[p] = fs[0]
        if len(fs) > 1:
            avisos.append(f'{p}: {len(fs)} zips del mismo período; se usó el más reciente ({os.path.basename(fs[0])}).')
    return elegidos


def apartar(ruta):
    """Mueve un archivo existente a _anteriores/ con fecha y hora."""
    d = os.path.join(os.path.dirname(ruta), '_anteriores')
    os.makedirs(d, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(ruta))
    dest = os.path.join(d, f'{base} {dt.datetime.now():%Y%m%d-%H%M%S}{ext}')
    shutil.move(ruta, dest)
    return dest


# ── controles (los mismos que las fórmulas del libro, calculados en Python) ─
def importe_l(r):
    import procesar_portal_iva as P
    return float(r.tot) if int(r.tipo) in P.TIPO_C else float(r.neto) + float(r.ng) + float(r.ex)


def controles(periodos, c, v):
    import procesar_portal_iva as P
    from openpyxl.utils import get_column_letter  # noqa: F401  (asegura openpyxl)
    c = c.copy(); v = v.copy()
    c['L'] = [importe_l(r) for _, r in c.iterrows()]
    v['L'] = [importe_l(r) for _, r in v.iterrows()]
    c['p'] = c.fecha.dt.strftime('%Y%m'); v['p'] = v.fecha.dt.strftime('%Y%m')
    filas = []
    for p, d in periodos.items():
        xc = P.xml_consol(d, 'iva_compras'); xv = P.xml_consol(d, 'iva_ventas')
        n = lambda x, qs: P.xnum(P.xval(x, qs))
        libro_c = (n(xc, ['totales_compras/total_neto_gravado', 'totales_bienes_usados/total_neto_gravado',
                          'totales_importacion_bienes/total_neto_gravado'])
                   - n(xc, ['totales_compras_mlcf_restituir/total_neto_gravado'])
                   + n(xc, ['total_exento_no_gravado']) - n(xc, ['total_exento_no_gravado_nc'])
                   + n(xc, ['total_compras_no_fiscal']))
        det_c = c.loc[c.p == p, 'L'].sum()
        libro_v = (n(xv, ['totales_ventas_ri/total_neto_gravado', 'totales_ventas_cons_fin_exento_no_alcanzado/total_neto_gravado',
                          'totales_ventas_otros_conceptos_deb_fiscal/total_neto_gravado'])
                   - n(xv, ['totales_ventas_ri_otros_rest_deb_fiscal/total_neto_gravado',
                            'totales_ventas_exentos_no_alcanzado_mono_cons_fin_rest_deb_fiscal/total_neto_gravado'])
                   + n(xv, ['total_ventas_importes_exentos_no_gravados_excepto_exportaciones', 'total_ventas_exportacion'])
                   - n(xv, ['total_ventas_importes_exentos_no_gravados_nc_excepto_exportaciones']))
        det_v = v.loc[v.p == p, 'L'].sum()
        filas.append(dict(periodo=p,
                          compras_libro=round(libro_c, 2), compras_detalle=round(det_c, 2),
                          compras_dif=round(libro_c - det_c, 2),
                          ventas_libro=round(libro_v, 2), ventas_detalle=round(det_v, 2),
                          ventas_dif=round(det_v - libro_v, 2)))
    sin = round(float(c.loc[c.clasif.isna(), 'L'].sum()), 2)
    return filas, sin


# ── subcomandos ─────────────────────────────────────────────────────────────
def procesar(a):
    import procesar_portal_iva as P
    avisos = []
    carpeta = carpeta_portal(a)
    log(f'Carpeta: {carpeta}')
    compras = plantilla(carpeta, 'COMPRAS', 'Control con Portal IVA', a.compras)
    ventas = plantilla(carpeta, 'VENTAS', 'Comprobantes Ventas', a.ventas)
    pad = padron(carpeta, a.padron, avisos)
    zips = zips_por_periodo(carpeta, avisos)
    anios = sorted({p[:4] for p in zips})
    if len(anios) > 1:
        avisos.append(f'Los zips son de más de un año ({", ".join(anios)}): la hoja Control de COMPRAS usa {anios[0]} '
                      f'y el Resumen por Clasificación también; los meses del otro año no entran ahí.')
    destinos = {SALIDA_COMPRAS: os.path.join(carpeta, SALIDA_COMPRAS), SALIDA_VENTAS: os.path.join(carpeta, SALIDA_VENTAS)}
    existentes = [n for n, r in destinos.items() if os.path.exists(r)]
    if existentes and not a.solo_revisar and a.si_existe == 'error':
        fail('Ya hay resultados en la carpeta. Volvé a correr con --si-existe=reemplazar '
             '(los anteriores van a _anteriores/) o con --solo-revisar.', existentes=existentes)

    tmp = tempfile.mkdtemp(prefix='portal-iva-vulcamoia-')
    work, out = os.path.join(tmp, 'work'), os.path.join(tmp, 'out')
    os.makedirs(out)
    log(f'Procesando {len(zips)} períodos ({", ".join(zips)})...')
    periodos = P.extraer(list(zips.values()), work)
    _, clasif, conflictos, cambios_pad, otras = P.cargar_padron(pad, out)   # copia en tmp; el padrón real no se toca
    cout, c, sc = P.armar_compras(periodos, compras, clasif, out)
    vout, v, faltan = P.armar_ventas(periodos, ventas, out)
    filas, sin_imp = controles(periodos, c, v)

    malos_c = [f['periodo'] for f in filas if abs(f['compras_dif']) > TOL]
    malos_v = [f['periodo'] for f in filas if abs(f['ventas_dif']) > TOL]
    if faltan:
        avisos.append(f'VENTAS: la plantilla no tiene columna para {", ".join(faltan)} (fila 5, B:M): esos meses no entran en el control.')

    guardados, apartados = [], []
    if not a.solo_revisar:
        for nombre, origen in ((SALIDA_COMPRAS, cout), (SALIDA_VENTAS, vout)):
            dest = destinos[nombre]
            if os.path.exists(dest):
                apartados.append(os.path.relpath(apartar(dest), carpeta))
            shutil.copyfile(origen, dest)
            guardados.append(nombre)

    usados = {str(k): m for k, m in conflictos.items() if k in set(c.cuit)}
    res = dict(ok=True, carpeta=carpeta, solo_revisar=bool(a.solo_revisar),
               entradas=dict(compras=os.path.basename(compras), ventas=os.path.basename(ventas),
                             padron=os.path.basename(pad), zips={p: os.path.basename(f) for p, f in zips.items()}),
               compras_filas=len(c), ventas_filas=len(v), despachos_impo=int(c.pv.isna().sum()),
               controles=dict(
                   compras_cierran=not malos_c, compras_meses_con_diferencia=malos_c,
                   ventas_cierran=not malos_v, ventas_meses_con_diferencia=malos_v,
                   resumen_diferencia=round(-sin_imp, 2),
                   resumen_nota='La "Diferencia (debe ser 0)" del Resumen por Clasificación es menos lo que quedó '
                                'SIN CLASIFICAR: da 0 cuando se clasifican todos.',
                   por_mes=filas),
               sin_clasificar_importe=sin_imp, sin_clasificar=sc,
               padron_variantes_unificadas=cambios_pad, padron_otras_variantes_posibles=otras,
               cuit_con_varias_clasificaciones_usados=usados,
               guardados=guardados, anteriores=apartados,
               carpeta_trabajo=tmp, avisos=avisos)
    print(json.dumps(res, ensure_ascii=False, indent=1, default=str))


def clasificar(a):
    import subprocess
    carpeta = carpeta_portal(a)
    avisos = []
    compras = a.compras or os.path.join(carpeta, SALIDA_COMPRAS)
    if not os.path.exists(compras):
        fail(f'No está "{os.path.basename(compras)}" en la carpeta: corré primero procesar.')
    pad = padron(carpeta, a.padron, avisos)
    if not os.path.isfile(a.json):
        fail(f'No existe el JSON de clasificaciones "{a.json}".')
    copia_pad = None
    if a.modo == 'confirmada':
        d = os.path.join(os.path.dirname(pad), '_anteriores')
        os.makedirs(d, exist_ok=True)
        b, e = os.path.splitext(os.path.basename(pad))
        copia_pad = os.path.join(d, f'{b} {dt.datetime.now():%Y%m%d-%H%M%S}{e}')
        shutil.copy2(pad, copia_pad)
    r = subprocess.run([sys.executable, os.path.join(AQUI, 'agregar_clasificaciones.py'), '--compras', compras,
                        '--padron', pad, '--json', a.json, '--modo', a.modo], capture_output=True, text=True)
    if r.returncode:
        fail('agregar_clasificaciones.py terminó con error.', detalle=(r.stderr or r.stdout)[-2000:])
    print(json.dumps(dict(ok=True, modo=a.modo, compras=os.path.basename(compras), padron=os.path.basename(pad),
                          copia_padron=os.path.relpath(copia_pad, carpeta) if copia_pad else None,
                          salida=r.stdout.strip().splitlines(), avisos=avisos), ensure_ascii=False, indent=1))


def main():
    cargar_env()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    comunes = argparse.ArgumentParser(add_help=False)
    comunes.add_argument('--sociedad', help=f'nombre o CUIT de la carpeta de la sociedad (default {SOCIEDAD_DEFECTO})')
    comunes.add_argument('--anio', help='año de "04-Balances y EECC/<año>/PORTAL IVA" (default el más nuevo)')
    comunes.add_argument('--carpeta', help='carpeta PORTAL IVA explícita (en vez de --sociedad/--anio)')
    comunes.add_argument('--compras', help='plantilla de COMPRAS (procesar) o COMPRAS Completo (clasificar)')
    comunes.add_argument('--padron', help='padrón Clasificacion Proveedores (default el más reciente de la carpeta)')
    p = sub.add_parser('procesar', parents=[comunes])
    p.add_argument('--ventas', help='plantilla de VENTAS')
    p.add_argument('--si-existe', choices=['error', 'reemplazar'], default='error')
    p.add_argument('--solo-revisar', action='store_true', help='procesa y muestra los controles sin guardar en la carpeta')
    q = sub.add_parser('clasificar', parents=[comunes])
    q.add_argument('--json', required=True)
    q.add_argument('--modo', choices=['sugerida', 'confirmada'], required=True)
    a = ap.parse_args()
    try:
        import pandas, openpyxl  # noqa: F401
    except ImportError as e:
        fail(f'Falta {e.name}: python -m pip install --user -r requirements.txt (o correr setup-bva).')
    procesar(a) if a.cmd == 'procesar' else clasificar(a)


if __name__ == '__main__':
    main()
