#!/usr/bin/env python3
# saldos-iva-iibb.py — arma los saldos de IVA e IIBB de un mes (columnas C a I de la
# planilla FARMACIAS SALDO) leyendo las DDJJ de cada sociedad en la unidad compartida.
#
#   python scripts/saldos-iva-iibb.py --periodo=09/2026
#   python scripts/saldos-iva-iibb.py --periodo=202609 --cuits="<CUIT>;<CUIT>"
#   python scripts/saldos-iva-iibb.py --periodo=09/2026 --valores=<valores.json>
#   python scripts/saldos-iva-iibb.py --periodo=09/2026 --valores=<valores.json> --preparar-subida
#
# Argumentos:
#   --periodo            MM/AAAA o AAAAMM (el de las DDJJ, no el de presentacion).
#   --cuit / --cuits     Una sociedad o varias separadas por ";" (default: todas).
#   --nombre / --nombres Nombre (o parte) de la carpeta de la sociedad (o varios con ";").
#   --valores            JSON con lo leido a mano de las DJ que son imagen (ver abajo).
#   --preparar-subida    Copia las DJ usadas a <salida>/DDJJ IVA-IIBB MM-AAAA/ para subirlas
#                        al SharePoint. No sube nada.
#   --salida             Carpeta de salida (default <BVA_SALIDAS_PATH>/saldos-iva-iibb).
#
# Que lee (por sociedad, en 01-Impuestos Mensuales/AAAAMM o .../AAAA/AAAAMM y subcarpetas,
# sin ACUSE, VEP, SICORE, SIRE, TSH, AGENTES ni CONDOMINIO):
#   - IVA: F.2051 (nombres "iva-dj", "dj-iva", "IVA-DDJJ"...).
#   - IIBB: R-606M de ARBA, SICOL de AGIP o CM03 de Convenio Multilateral ("iibb-dj",
#     "IIBB-DDJJ", "IIBB AGIP-DDJJ", "IIBB CM-DDJJ", "IIBBCM-DJ"...).
# El F.2051 y el R-606M con texto se leen solos. Las DJ que son imagen, y todas las de AGIP
# y CM03 (el texto sale desordenado), se pasan a PNG en <salida>/imagenes/AAAAMM/ y quedan
# como "pendientes" en el JSON de salida, con los campos que hay que leer. Claude (o una
# persona) los lee de la imagen, los escribe en valores.json y vuelve a correr con --valores:
#
#   {"<CUIT>": {"iva":  {"st_arca": 0, "st_contribuyente": 0, "sld_anterior": 0,
#                         "retenciones": 0, "sld": 0, "a_pagar": 0},
#               "iibb": {"tipo": "ARBA", "saldo_dj": 0, "saldo_anterior": 0, "acumulado": 0}}}
#   AGIP: {"tipo": "AGIP", "anticipo": 0, "retenciones": 0, "percepciones": 0,
#          "pagos_a_cuenta": 0, "saldos_a_favor": 0, "otros_creditos": 0,
#          "favor_dgr": 0, "favor_contribuyente": 0, "total_a_pagar": 0}
#   CM03: {"tipo": "CM", "ba_favor": 0, "ba_pagar": 0, "caba_favor": 0, "caba_pagar": 0,
#          "total_a_pagar": 0, "otras": "Cordoba a favor 1.319.642,68; ..."}
#   Importes como numeros (punto decimal), en positivo salvo los saldos de ARBA, que van
#   con el signo del PDF (negativo = a favor del contribuyente).
#
# Columnas (criterio del SKILL.md):
#   C saldo tecnico a favor del contribuyente; D SLD del periodo; E SLD anterior usado =
#   min(SLD anterior, max(0, saldo tecnico ARCA - retenciones)); F saldo a favor de ARCA.
#   ARBA: acumulado < 0 -> G = -acumulado; > 0 -> I = acumulado. AGIP: H = a favor del
#   contribuyente, I = total a pagar. CM03: G = Buenos Aires a favor, H = CABA a favor,
#   I = total a pagar. Jurisdiccion donde no presenta -> 0; DJ que falta -> vacio.
# Controles: IVA  F - D = ST ARCA - SLD anterior - retenciones;
#            ARBA acumulado = saldo DJ - saldo anterior (ARBA redondea a 10 centavos);
#            AGIP anticipo - deducciones = a favor DGR - a favor contribuyente.
#
# Salida: <salida>/SALDOS IVA IIBB MM-AAAA.xlsx y un JSON por stdout (logs por stderr).
# Si algo no cierra imprime {"error": ...} y termina con 1. Solo lee la unidad.

import datetime, json, os, re, shutil, sys, unicodedata

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    import pypdf
except ImportError:
    print(json.dumps({'error': 'Faltan dependencias: pip install -r requirements.txt'}))
    sys.exit(1)

RAIZ_DEFECTO = r'I:\Unidades compartidas\BVA - Sociedades Farmaceuticas'
PLUGIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
EXCLUIR = re.compile(r'acuse|vep|sicore|sire|tsh|agentes|condominio|retenciones', re.I)
PAT_IVA = re.compile(r'(iva[-_ ]*d?dj|d?dj[-_ ]*iva)', re.I)
PAT_IIBB = re.compile(r'(iibb[-_ ]*(cm|agip)?[-_ ]*d?dj|d?dj[-_ ]*iibb)', re.I)
CAMPOS_IVA = ['st_arca', 'st_contribuyente', 'sld_anterior', 'retenciones', 'sld', 'a_pagar']
CAMPOS = {
    'ARBA': ['saldo_dj', 'saldo_anterior', 'acumulado'],
    'AGIP': ['anticipo', 'retenciones', 'percepciones', 'pagos_a_cuenta', 'saldos_a_favor',
             'otros_creditos', 'favor_dgr', 'favor_contribuyente', 'total_a_pagar'],
    'CM': ['ba_favor', 'ba_pagar', 'caba_favor', 'caba_pagar', 'total_a_pagar', 'otras'],
}


class Falla(Exception):
    pass


def fail(msg):
    print(json.dumps({'error': msg}, ensure_ascii=False))
    sys.exit(1)


def log(*a):
    print(*a, file=sys.stderr)


def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', s).strip()


def num(s):
    return float(s.replace('$', '').replace(' ', '').replace('.', '').replace(',', '.'))


# ── .env (mismo orden que lib/env.js de los plugins de Node) ────────────────
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


def raiz():
    r = (os.environ.get('BVA_UNIDAD_PATH') or '').strip() or RAIZ_DEFECTO
    if not os.path.isdir(r):
        fail(f'No encuentro la unidad compartida en "{r}". Defini BVA_UNIDAD_PATH en el .env.')
    return r


def sociedades():
    """{cuit: (nombre, carpeta)} de las carpetas '<n> - <Nombre> - <CUIT>' (solo sociedades 30/33)."""
    base = raiz()
    out = {}
    for d in sorted(os.listdir(base)):
        m = re.match(r'^\s*\d+\s*-\s*(.+?)\s*-\s*(\d{11})\s*$', unicodedata.normalize('NFC', d))
        if m and m.group(2)[:2] in ('30', '33') and os.path.isdir(os.path.join(base, d)):
            out[m.group(2)] = (m.group(1), os.path.join(base, d))
    return out


def elegir(socs, a):
    if a.get('cuit') or a.get('cuits'):
        cuits = [c.strip().replace('-', '') for c in (a.get('cuit') or a.get('cuits')).split(';') if c.strip()]
        faltan = [c for c in cuits if c not in socs]
        if faltan:
            raise Falla(f'No hay carpeta de sociedad con el CUIT {", ".join(faltan)}.')
        return cuits
    if a.get('nombre') or a.get('nombres'):
        out = []
        for q in (a.get('nombre') or a.get('nombres')).split(';'):
            q = norm(q)
            hits = [c for c, (n, _) in socs.items() if norm(n) == q] or \
                   [c for c, (n, _) in socs.items() if re.search(rf'(^|\W){re.escape(q)}(\W|$)', norm(n))]
            if len(hits) != 1:
                raise Falla(f'"{q}" coincide con {len(hits)} sociedades: '
                            f'{" | ".join(socs[c][0] for c in hits) or "ninguna"}. Usa --cuit.')
            out.append(hits[0])
        return out
    return sorted(socs, key=lambda c: norm(socs[c][0]))


def carpeta_mes(soc, per):
    mens = os.path.join(soc, '01-Impuestos Mensuales')
    for d in (os.path.join(mens, per), os.path.join(mens, per[:4], per)):
        if os.path.isdir(d):
            return d
    return None


def pdfs_dj(soc, per):
    base = carpeta_mes(soc, per)
    if not base:
        return [], []
    todos = [os.path.join(d, f) for d, _, fs in os.walk(base) for f in fs
             if f.lower().endswith('.pdf') and not EXCLUIR.search(os.path.relpath(os.path.join(d, f), base))]
    iva = sorted(p for p in todos if PAT_IVA.search(os.path.basename(p)))
    iibb = sorted(p for p in todos if PAT_IIBB.search(os.path.basename(p)))
    return iva, iibb


def texto(pdf):
    try:
        return '\n'.join(p.extract_text() or '' for p in pypdf.PdfReader(pdf).pages)
    except Exception as e:
        log(f'  no pude leer {os.path.basename(pdf)}: {e}')
        return ''


def tipo_iibb(pdf, txt):
    if re.search(r'R-606M', txt):
        return 'ARBA'
    if re.search(r'SICOL|CONTRIBUYENTES LOCALES', txt, re.I):
        return 'AGIP'
    if re.search(r'CM03|CONVENIO MULTILATERAL', txt, re.I):
        return 'CM'
    n = os.path.basename(pdf)
    if re.search(r'agip', n, re.I):
        return 'AGIP'
    if re.search(r'iibb[-_ ]*cm|cm[-_ ]*d?dj', n, re.I):
        return 'CM'
    return None  # imagen sin pista en el nombre: lo dice quien la lea


# ── lectura del texto ───────────────────────────────────────────────────────
def leer_iva(txt):
    def g(lab):
        m = re.search(re.escape(lab) + r'\s*\$\s*([\d\.,]+)', txt)
        return num(m.group(1)) if m else 0.0
    cu = re.search(r'CUIT\s*\n?\s*(\d{2}-\d{8}-\d)', txt)
    pe = re.search(r'Per[ií]odo\s*\n?\s*(\d{6})', txt)
    sec = re.search(r'Secuencia\s*\n?\s*(Original|Rectificativa \d+)', txt)
    v = {'st_arca': g('Saldo técnico a favor de ARCA'),
         'st_contribuyente': 0.0,
         'sld_anterior': g('Saldo a favor de libre disponibilidad del período anterior neto de usos'),
         'retenciones': g('Total de retenciones, percepciones y pagos a cuenta neto de restituciones'),
         'sld': g('Saldo de libre disponibilidad a favor del contribuyente del período'),
         'a_pagar': g('Saldo de impuesto a favor de ARCA')}
    m = re.search(r'Saldo técnico a favor del contribuyente\s*\$\s*([\d\.,]+)', txt)
    if m:
        v['st_contribuyente'] = num(m.group(1))
    return v, {'cuit': cu and cu.group(1).replace('-', ''), 'periodo': pe and pe.group(1),
               'secuencia': sec and sec.group(1)}


def leer_arba(txt):
    """Saldo de la DJ, saldo a favor anterior y acumulado: tres importes seguidos (renglones
    de un solo importe) con a - b = c, seguidos del saldo excedente cooperativa ($ 0,00)."""
    v = [num(x) for x in re.findall(r'^\$\s?(-?[\d\.]+,\d\d)$', txt, re.M)]
    cand = [i for i in range(len(v) - 3)
            if (v[i] or v[i + 1]) and abs(v[i] - v[i + 1] - v[i + 2]) < 1 and v[i + 3] == 0]
    cu = re.search(r'R-606M\s*\n(\d{2}-\d{8}-\d)', txt)
    pe = re.search(r'\n(\d{1,2})/(\d{4})\n', txt)
    rect = re.search(r'\n(Original|Rectificativa[^\n]*)\nCerrada', txt)
    meta = {'cuit': cu and cu.group(1).replace('-', ''),
            'periodo': pe and f'{pe.group(2)}{int(pe.group(1)):02d}', 'secuencia': rect and rect.group(1)}
    if len(cand) != 1:
        return None, meta
    a, b, c = v[cand[0]:cand[0] + 3]
    return {'saldo_dj': a, 'saldo_anterior': b, 'acumulado': c}, meta


# ── columnas y controles ────────────────────────────────────────────────────
def columnas_iva(v):
    E = min(v['sld_anterior'], max(0.0, v['st_arca'] - v['retenciones']))
    ok = abs((v['a_pagar'] - v['sld']) - (v['st_arca'] - v['sld_anterior'] - v['retenciones'])) < 0.02
    return {'C': v['st_contribuyente'], 'D': v['sld'], 'E': round(E, 2), 'F': v['a_pagar']}, ok


def columnas_iibb(tipo, v):
    if tipo == 'ARBA':
        acum = v['acumulado']
        ok = abs(v['saldo_dj'] - v['saldo_anterior'] - acum) < 1
        return {'G': -acum if acum < 0 else 0.0, 'H': 0.0, 'I': acum if acum > 0 else 0.0}, ok
    if tipo == 'AGIP':
        ded = sum(v[k] for k in ('retenciones', 'percepciones', 'pagos_a_cuenta', 'saldos_a_favor', 'otros_creditos'))
        ok = abs((v['anticipo'] - ded) - (v['favor_dgr'] - v['favor_contribuyente'])) < 0.02
        return {'G': 0.0, 'H': v['favor_contribuyente'], 'I': v['total_a_pagar']}, ok
    if tipo == 'CM':
        return {'G': v['ba_favor'], 'H': v['caba_favor'], 'I': v['total_a_pagar']}, True
    raise Falla(f'Tipo de IIBB desconocido: {tipo}')


def imagen(pdf, destino):
    """PNG de cada pagina para leer a mano. Devuelve las rutas."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return []
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    doc, out = pdfium.PdfDocument(pdf), []
    for i in range(len(doc)):
        ruta = destino if len(doc) == 1 else destino.replace('.png', f'-p{i + 1}.png')
        doc[i].render(scale=1.6).to_pil().convert('RGB').save(ruta)
        out.append(ruta)
    return out


def args(argv):
    a = {}
    for x in argv:
        m = re.match(r'^--([a-z-]+)(?:=(.*))?$', x)
        if not m:
            fail(f'Argumento no reconocido: {x}')
        a[m.group(1)] = m.group(2) if m.group(2) is not None else True
    return a


def periodo(s):
    s = (s or '').strip()
    m = re.match(r'^(\d{1,2})/(\d{4})$', s) or re.match(r'^(\d{4})(\d{2})$', s)
    if not m:
        fail('Falta --periodo=MM/AAAA (o AAAAMM).')
    mm, aa = (m.group(1), m.group(2)) if '/' in s else (m.group(2), m.group(1))
    return f'{aa}{int(mm):02d}'


def main():
    cargar_env()
    a = args(sys.argv[1:])
    per = periodo(a.get('periodo'))
    mmaaaa = f'{per[4:]}-{per[:4]}'
    previo = f'{int(per[:4]) - 1}12' if per[4:] == '01' else f'{per[:4]}{int(per[4:]) - 1:02d}'
    salida = a.get('salida') or os.path.join(os.environ.get('BVA_SALIDAS_PATH') or
                                             os.path.expanduser('~/Documents/BVA-salidas'), 'saldos-iva-iibb')
    manual = {}
    if a.get('valores'):
        try:
            manual = json.load(open(a['valores'], encoding='utf-8'))
        except Exception as e:
            fail(f'No pude leer {a["valores"]}: {e}')
    socs = sociedades()
    try:
        cuits = elegir(socs, a)
    except Falla as e:
        fail(str(e))

    filas, pendientes, subir = [], [], []
    for cu in cuits:
        nombre, soc = socs[cu]
        iva_f, iibb_f = pdfs_dj(soc, per)
        f = {'sociedad': nombre, 'cuit': cu, 'tipo_iibb': None, 'avisos': [],
             'dj_iva': os.path.basename(iva_f[0]) if iva_f else None,
             'dj_iibb': os.path.basename(iibb_f[0]) if iibb_f else None}
        f.update({k: None for k in 'CDEFGHI'})
        sin_previo = not iva_f and not iibb_f and not any(pdfs_dj(soc, previo))
        if len(iva_f) > 1 or len(iibb_f) > 1:
            f['avisos'].append('varias DJ: ' + '; '.join(os.path.basename(p) for p in iva_f[1:] + iibb_f[1:]) +
                               ' (se usa la primera: revisar si hay rectificativa)')
        man = manual.get(cu, {})

        # IVA
        if iva_f:
            subir.append(iva_f[0])
            v = None
            if man.get('iva'):
                v = {k: float(man['iva'].get(k) or 0) for k in CAMPOS_IVA}
                f['avisos'].append('IVA leida de imagen')
            else:
                txt = texto(iva_f[0])
                if len(txt.strip()) > 50:
                    v, meta = leer_iva(txt)
                    if meta['cuit'] != cu:
                        f['avisos'].append(f'CUIT de la DJ IVA {meta["cuit"]}')
                    if meta['periodo'] != per:
                        f['avisos'].append(f'periodo de la DJ IVA {meta["periodo"]}')
                    if meta['secuencia'] and meta['secuencia'] != 'Original':
                        f['avisos'].append(f'IVA {meta["secuencia"]}')
                else:
                    pendientes.append({'cuit': cu, 'sociedad': nombre, 'impuesto': 'iva', 'campos': CAMPOS_IVA,
                                       'pdf': iva_f[0], 'imagenes': imagen(iva_f[0], os.path.join(
                                           salida, 'imagenes', per, f'{cu}-iva.png'))})
            if v:
                cols, ok = columnas_iva(v)
                f.update(cols)
                if not ok:
                    f['avisos'].append('CONTROL IVA NO CIERRA')
        # IIBB
        if iibb_f:
            subir.append(iibb_f[0])
            txt = texto(iibb_f[0])
            tipo = (man.get('iibb') or {}).get('tipo') or tipo_iibb(iibb_f[0], txt)
            f['tipo_iibb'] = tipo
            v = None
            if man.get('iibb'):
                v = {k: (man['iibb'].get(k) if k == 'otras' else float(man['iibb'].get(k) or 0))
                     for k in CAMPOS.get(tipo, [])}
                f['avisos'].append(f'IIBB {tipo} leida de imagen')
            elif tipo == 'ARBA' and len(txt.strip()) > 50:
                v, meta = leer_arba(txt)
                if meta['cuit'] != cu:
                    f['avisos'].append(f'CUIT de la DJ IIBB {meta["cuit"]}')
                if meta['periodo'] != per:
                    f['avisos'].append(f'periodo de la DJ IIBB {meta["periodo"]}')
                if meta['secuencia'] and meta['secuencia'] != 'Original':
                    f['avisos'].append(f'IIBB {meta["secuencia"]}')
            if v is None:
                pendientes.append({'cuit': cu, 'sociedad': nombre, 'impuesto': 'iibb', 'tipo': tipo,
                                   'campos': CAMPOS.get(tipo) or 'tipo + campos de ARBA, AGIP o CM',
                                   'pdf': iibb_f[0], 'imagenes': imagen(iibb_f[0], os.path.join(
                                       salida, 'imagenes', per, f'{cu}-iibb.png'))})
            else:
                cols, ok = columnas_iibb(tipo, v)
                f.update(cols)
                if not ok:
                    f['avisos'].append(f'CONTROL {tipo} NO CIERRA')
                if tipo == 'AGIP':
                    f['avisos'].append('G=0: solo DJ de AGIP (confirmar que no presenta ARBA)')
                if tipo == 'CM' and v.get('otras'):
                    f['avisos'].append(f'CM03 otras jurisdicciones: {v["otras"]}')
        else:
            prev = pdfs_dj(soc, previo)[1]
            f['tipo_iibb'] = tipo_iibb(prev[0], texto(prev[0])) if prev else None
        faltan = [n for n, ok in (('IVA', iva_f), ('IIBB', iibb_f)) if not ok]
        f['estado'] = 'falta DJ ' + ' y '.join(faltan) if faltan else \
                      ('pendiente de leer imagen' if any(p['cuit'] == cu for p in pendientes) else 'completa')
        if sin_previo:
            f['avisos'].append('tampoco hay DJ del mes anterior en la unidad: ver si va en la planilla')
        filas.append(f)
        log(f'{nombre[:24]:24} {cu} {f["estado"]}')

    os.makedirs(salida, exist_ok=True)
    xlsx = os.path.join(salida, f'SALDOS IVA IIBB {mmaaaa}.xlsx')
    escribir_excel(xlsx, filas, mmaaaa)
    carpeta_subida = None
    if a.get('preparar-subida'):
        carpeta_subida = os.path.join(salida, f'DDJJ IVA-IIBB {mmaaaa}')
        os.makedirs(carpeta_subida, exist_ok=True)
        for p in subir:
            shutil.copy2(p, os.path.join(carpeta_subida, os.path.basename(p)))
    print(json.dumps({
        'periodo': per, 'excel': xlsx, 'sociedades': len(filas),
        'completas': sum(f['estado'] == 'completa' for f in filas),
        'faltan_dj': [{'sociedad': f['sociedad'], 'cuit': f['cuit'], 'estado': f['estado']}
                      for f in filas if f['estado'].startswith('falta')],
        'controles_que_no_cierran': [{'sociedad': f['sociedad'], 'avisos': f['avisos']}
                                     for f in filas if any('NO CIERRA' in x for x in f['avisos'])],
        'pendientes': pendientes,
        'carpeta_para_sharepoint': carpeta_subida, 'dj_para_subir': len(subir),
    }, ensure_ascii=False, indent=1))


def escribir_excel(ruta, filas, mmaaaa):
    wb = Workbook()
    ws = wb.active
    ws.title = mmaaaa
    hdr = ['Sociedad', 'CUIT', 'C SALDO TECNICO', 'D SLD', 'E SLD Utilizado Comp.', 'F Pagos IVA',
           'G IIBB PROV. (Sdo a Favor)', 'H IIBB CABA (Sdo a favor)', 'I Pagos IIBB', 'Tipo IIBB',
           'Estado', 'Avisos', 'DJ IVA', 'DJ IIBB']
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    rojo, amarillo = PatternFill('solid', fgColor='F8CBAD'), PatternFill('solid', fgColor='FFF2CC')
    for f in filas:
        ws.append([f['sociedad'], f['cuit']] + [f[k] for k in 'CDEFGHI'] +
                  [f['tipo_iibb'], f['estado'], ' | '.join(f['avisos']), f['dj_iva'], f['dj_iibb']])
        r = ws.max_row
        for col in range(3, 10):
            cel = ws.cell(r, col)
            cel.number_format = '#,##0.00;-#,##0.00;"-"'
            if cel.value is None:
                cel.fill = rojo
        if any('NO CIERRA' in x or 'imagen' in x for x in f['avisos']):
            ws.cell(r, 12).fill = amarillo
    for i, w in enumerate([24, 13, 15, 15, 15, 15, 17, 17, 15, 9, 22, 60, 38, 38], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C2'
    n = wb.create_sheet('Notas')
    for linea in [
        f'Saldos {mmaaaa} para la planilla FARMACIAS SALDO {mmaaaa} (SharePoint). Pegar por CUIT.',
        f'Generado {datetime.datetime.now():%d/%m/%Y %H:%M} con saldos-iva-iibb.py.',
        'Vacio en rojo = la DJ no esta en la unidad (no poner 0). Amarillo = leida de imagen o control que no cierra.',
        'Control IVA: F - D = saldo tecnico ARCA - SLD anterior - retenciones. ARBA: acumulado = saldo DJ - saldo anterior.',
    ]:
        n.append([linea])
    n.column_dimensions['A'].width = 120
    wb.save(ruta)


if __name__ == '__main__':
    main()
