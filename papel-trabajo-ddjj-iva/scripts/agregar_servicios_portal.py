"""Etapa 2 (DDJJ de IVA), paso 5: agrega a la solapa "Compras portal iva" las facturas de AySA,
EDESUR y METROGAS de "Compras Zetti" (las del paso 4) y vuelve a totalizar.

    python agregar_servicios_portal.py --papel=<papel de trabajo.xlsx>

Por cada factura de servicios de Zetti:
  * Si tiene algun importe negativo sin ser NC: NO se agrega; en "Compras Zetti", al costado de
    su fila, queda la aclaracion "No lo tomamos porque esta mal cargado".
  * Si ya esta en el Portal (CUIT + PV + numero) no se vuelve a agregar.
  * Alicuota de IVA = IVA / Grav de Zetti (2,5 / 5 / 10,5 / 21 / 27 %): el neto y el IVA van a las
    columnas de esa alicuota. Si no da ninguna, corta sin grabar.
  * Columnas (como ARCA las carga en la DDJJ presentada): Exen -> Importe Exento, P/R IIBB -> Perc.
    IIBB, P/R IVA -> Perc. IVA, P/R IG -> Per. Otros Imp. Nac.; Total Neto Gravado y Total IVA.

Las filas se agregan debajo del ultimo comprobante con el formato de esa fila, sin nada a la
derecha (la columna AG queda libre para la usuaria, pedido 2026-11-02); despues van la fila TOTAL (=SUM hasta la ultima
fila nueva) y la nota con la cantidad de comprobantes. Imprime un JSON con lo hecho.
"""
import json
import re
import sys
import unicodedata
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from filtrar_servicios import servicio

HOJA_Z = 'Compras Zetti'
HOJA_P = 'Compras portal iva'
MAL_CARGADO = 'No lo tomamos porque esta mal cargado'
DENOMINACION = {'30709565075': 'AGUA Y SANEAMIENTOS ARGENTINOS S A', '30655116512': 'EDESUR S A',
                '30657863676': 'METROGAS S A'}
TIPO_ARCA = {('FC', 'A'): 1, ('ND', 'A'): 2, ('NC', 'A'): 3, ('FC', 'B'): 6, ('ND', 'B'): 7, ('NC', 'B'): 8,
             ('FC', 'C'): 11, ('ND', 'C'): 12, ('NC', 'C'): 13}
ALICUOTAS = {0.025: '2,5%', 0.05: '5%', 0.105: '10,5%', 0.21: '21%', 0.27: '27%'}


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return ' '.join(s.lower().split())


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def fallar(msg, **extra):
    sys.exit(json.dumps({'error': msg, **extra}, ensure_ascii=False))


def alicuota(grav, iva):
    if abs(grav) < 0.005:
        return None if abs(iva) < 0.005 else 'sin_neto'
    tasa = iva / grav
    return next((a for a in ALICUOTAS if abs(tasa - a) <= 0.002), 'no_reconocida')


CONTROL_TITULO = 'Control alicuota IVA servicios (AySA / EDESUR / METROGAS)'


def control_alicuotas(wz, cz, controles, no_tomados):
    """Debajo del TOTAL de Compras Zetti: un renglon por servicio con la alicuota con que se cargo.
    En la columna IVA, =+<Grav de la fila>*alicuota (como lo controla la usuaria a mano), al lado
    del texto con proveedor, comprobante y alicuota. Si se vuelve a correr, reemplaza el bloque."""
    fila_total = next((r for r in range(3, wz.max_row + 1)
                       if any(str(wz.cell(r, c).value or '').strip().upper() == 'TOTAL' for c in range(1, cz['exen']))), None)
    if fila_total is None:
        fallar('No encontre la fila TOTAL en la solapa Compras Zetti')
    for r in range(fila_total + 1, wz.max_row + 1):  # bloque anterior
        if wz.cell(r, cz['proveedor']).value == CONTROL_TITULO:
            for rr in range(r, wz.max_row + 1):
                for c in range(1, wz.max_column + 1):
                    wz.cell(rr, c).value = None
            break
    if not controles and not no_tomados:
        return
    letra_grav = get_column_letter(cz['grav'])
    fuente = dict(name='Aptos Narrow', size=11)
    r = fila_total + 2
    wz.cell(r, cz['proveedor'], CONTROL_TITULO).font = Font(bold=True, **fuente)
    for sv in sorted(controles, key=lambda x: x['fila']):
        r += 1
        a = sv['alicuota']
        nombre = f"{sv['proveedor']} {sv['tc']} {sv['m']} {sv['pv']}-{sv['numero']} (fila {sv['fila']})"
        if a is None:
            texto, formula = f'{nombre}: todo exento / no gravado, sin IVA', 0
        else:
            texto = f'{nombre}: IVA al {ALICUOTAS[a]}'
            formula = f'=+{letra_grav}{sv["fila"]}*{a}'
        wz.cell(r, cz['proveedor'], texto).font = Font(**fuente)
        c = wz.cell(r, cz['iva'], formula)
        c.number_format, c.font = '#,##0.00', Font(**fuente)
    for nt in no_tomados:
        r += 1
        wz.cell(r, cz['proveedor'], f"{nt['comprobante']} (fila {nt['fila']}): {MAL_CARGADO}").font =             Font(bold=True, color='C00000', **fuente)


def main():
    opts = dict((a[2:].split('=', 1) + [''])[:2] for a in sys.argv[1:] if a.startswith('--'))
    papel = opts.get('papel')
    if not papel:
        sys.exit(__doc__)
    wb = load_workbook(papel)
    for h in (HOJA_Z, HOJA_P):
        if h not in wb.sheetnames:
            fallar(f'Falta la solapa "{h}"', solapas=wb.sheetnames)
    wz, wp = wb[HOJA_Z], wb[HOJA_P]

    # --- Compras Zetti: facturas de servicios
    tz = [norm(c.value) for c in wz[2]]
    cz = {k: tz.index(k) + 1 for k in ('tc', 'm', 'nro. comp.', 'proveedor', 'cuit', 'exen', 'grav', 'iva',
                                       'p/r iibb', 'p/r iva', 'p/r ig', 'total')}
    col_aclaracion = wz.max_column + 1
    for r in range(3, wz.max_row + 1):  # aclaraciones viejas (re-ejecucion)
        if wz.cell(r, col_aclaracion - 1).value == MAL_CARGADO:
            col_aclaracion -= 1
            break
    servicios = []
    for r in range(3, wz.max_row + 1):
        fecha = wz.cell(r, 1).value
        if not isinstance(fecha, datetime):
            continue
        g = lambda k: wz.cell(r, cz[k]).value  # noqa: E731
        s = servicio(g('cuit'), g('proveedor'))
        if not s:
            continue
        imp = {k: num(g(k)) for k in ('exen', 'grav', 'iva', 'p/r iibb', 'p/r iva', 'p/r ig', 'total')}
        servicios.append({'fila': r, 'servicio': s, 'fecha': fecha, 'tc': str(g('tc') or '').strip().upper(),
                          'm': str(g('m') or '').strip().upper(), 'pv': g('nro. comp.'),
                          'numero': wz.cell(r, cz['nro. comp.'] + 1).value, 'proveedor': g('proveedor'),
                          'cuit': re.sub(r'\D', '', str(g('cuit') or '')), 'imp': imp})

    # --- Compras portal iva: columnas, ultima fila de datos, TOTAL y nota
    tp = [norm(c.value) for c in wp[1]]
    cp = lambda t: tp.index(t) + 1  # noqa: E731
    ultima = max(r for r in range(2, wp.max_row + 1) if isinstance(wp.cell(r, 1).value, datetime))
    existentes = {(re.sub(r'\D', '', str(wp.cell(r, cp('nro. doc. vendedor')).value or '')),
                   wp.cell(r, cp('punto de venta')).value, wp.cell(r, cp('numero de comprobante')).value)
                  for r in range(2, ultima + 1)}
    fila_total = next((r for r in range(ultima + 1, wp.max_row + 1) if wp.cell(r, 1).value == 'TOTAL'), None)
    cols_total = [c for c in range(1, wp.max_column + 1)
                  if fila_total and str(wp.cell(fila_total, c).value or '').startswith('=SUM(')]

    agregados, no_tomados, ya_estaban, controles = [], [], [], []
    for sv in servicios:
        imp = sv['imp']
        desc = f"{sv['tc']} {sv['m']} {sv['pv']}-{sv['numero']} {sv['proveedor']} ({sv['fecha']:%d/%m/%Y})"
        if sv['tc'] != 'NC' and any(v < 0 for v in imp.values()):
            c = wz.cell(sv['fila'], col_aclaracion, MAL_CARGADO)
            c.font = Font(name='Aptos Narrow', size=11, bold=True, color='C00000')
            no_tomados.append({**{k: sv[k] for k in ('servicio', 'fila')}, 'comprobante': desc,
                               'negativos': {k: v for k, v in imp.items() if v < 0}})
            continue
        a = alicuota(imp['grav'], imp['iva'])
        if a in ('sin_neto', 'no_reconocida'):
            fallar(f'No reconozco la alicuota de IVA de {desc}: IVA {imp["iva"]} / Grav {imp["grav"]}'
                   + (f' = {imp["iva"] / imp["grav"]:.4%}' if imp['grav'] else ''))
        controles.append({**sv, 'alicuota': a})
        if (sv['cuit'], sv['pv'], sv['numero']) in existentes:
            ya_estaban.append(desc)
            continue
        tipo = TIPO_ARCA.get((sv['tc'], sv['m']))
        if tipo is None:
            fallar(f'No se que tipo de ARCA es {sv["tc"]} {sv["m"]}: {desc}')
        agregados.append({**sv, 'tipo': tipo, 'alicuota': a, 'desc': desc})

    control_alicuotas(wz, cz, controles, no_tomados)

    if agregados:
        if not cols_total:
            fallar('No encontre la fila TOTAL con =SUM en la solapa Compras portal iva')
        nota = next((wp.cell(r, 1).value for r in range(fila_total + 1, wp.max_row + 1)
                     if isinstance(wp.cell(r, 1).value, str)), '')
        estilos = {c: copy(wp.cell(ultima, c)._style) for c in range(1, wp.max_column + 1)}
        estilo_total = {c: copy(wp.cell(fila_total, c)._style) for c in range(1, wp.max_column + 1)}
        alto_nota = fila_total + 1
        # se limpia desde la primera fila libre hasta la nota (se vuelven a escribir mas abajo)
        for r in range(ultima + 1, max(wp.max_row, alto_nota) + 1):
            for c in range(1, wp.max_column + 2):
                wp.cell(r, c).value = None
        r = ultima
        for sv in agregados:
            r += 1
            imp = sv['imp']
            for c, st in estilos.items():
                wp.cell(r, c)._style = copy(st)
            pct = ALICUOTAS.get(sv['alicuota'])
            valores = {
                'fecha de emision': sv['fecha'], 'tipo de comprobante': sv['tipo'], 'punto de venta': sv['pv'],
                'numero de comprobante': sv['numero'], 'tipo doc. vendedor': 80,
                'nro. doc. vendedor': int(sv['cuit']), 'denominacion vendedor': DENOMINACION.get(sv['cuit'], sv['proveedor']),
                'importe total': imp['total'], 'moneda original': 'PES', 'tipo de cambio': 1,
                'importe no gravado': 0, 'importe exento': imp['exen'], 'credito fiscal computable': 0,
                'importe de per. o pagos a cta. de otros imp. nac.': imp['p/r ig'],
                'importe de percepciones de ingresos brutos': imp['p/r iibb'],
                'importe de impuestos municipales': 0,
                'importe de percepciones o pagos a cuenta de iva': imp['p/r iva'],
                'importe de impuestos internos': 0, 'importe otros tributos': 0,
                'total neto gravado': imp['grav'], 'total iva': imp['iva'],
            }
            if pct:
                valores[norm(f'Neto Gravado IVA {pct}')] = imp['grav']
                valores[norm(f'Importe IVA {pct}')] = imp['iva']
            for t, v in valores.items():
                if t in tp:
                    wp.cell(r, cp(t), v)
        nueva_ult = r
        fila_total = nueva_ult + 2
        for c, st in estilo_total.items():
            wp.cell(fila_total, c)._style = copy(st)
        wp.cell(fila_total, 1, 'TOTAL')
        for c in cols_total:
            L = get_column_letter(c)
            wp.cell(fila_total, c, f'=SUM({L}2:{L}{nueva_ult})')
        n = nueva_ult - 1
        nota = re.sub(r'^\d+ comprobantes', f'{n} comprobantes', nota or '')
        nota += f' · {len(agregados)} agregados de Compras Zetti (servicios).'
        c = wp.cell(fila_total + 1, 1, nota)
        c.font = Font(name='Arial', size=10)
        if wp.auto_filter.ref:
            wp.auto_filter.ref = re.sub(r'\d+$', str(nueva_ult), wp.auto_filter.ref)
    if agregados or no_tomados or controles:
        wb.save(papel)

    print(json.dumps({
        'papel': papel,
        'agregados': [{'comprobante': a['desc'], 'alicuota_iva': ALICUOTAS.get(a['alicuota'], 'sin IVA'),
                       'neto': a['imp']['grav'], 'iva': a['imp']['iva'], 'total': a['imp']['total']} for a in agregados],
        'no_tomados_mal_cargados': no_tomados,
        'ya_estaban_en_el_portal': ya_estaban,
        'sin_comprobantes_en_zetti': sorted({'AySA (agua)', 'EDESUR (luz)', 'METROGAS (gas)'}
                                            - {s['servicio'] for s in servicios}),
    }, ensure_ascii=False, indent=1, default=str))


if __name__ == '__main__':
    main()
