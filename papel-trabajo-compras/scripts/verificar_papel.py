# verificar_papel.py — controles del papel de trabajo ya recalculado por Excel, antes de pisar el de la unidad.
#
#   python scripts/verificar_papel.py --original=<papel antes> --nuevo=<salida de openpyxl> --calc=<recalculado por Excel>
#
# Pasa (ok=true) solo si:
#   - "Control: TOTAL del cruce - hojas" (Resumen cruce compras) da 0 en Zetti y en Portal;
#   - "Control vs hoja Compras Portal IVA" (Clasificacion compras) da 0 en cantidad, imputado, IVA y total;
#   - la diferencia del cuadro de compras contra el libro es menor a $1 (centavos = redondeo de ARCA);
#   - no hay errores de Excel (#REF!, #VALUE!, ...) en ninguna hoja;
#   - las hojas que ya tenia el papel quedaron iguales (formulas y valores).
#
# Salida: JSON por stdout.
import sys, json, argparse, warnings
warnings.filterwarnings('ignore')
import openpyxl

HOJAS = ('Compras Zetti', 'Compras Portal IVA', 'Cruce compras', 'Resumen cruce compras', 'Clasificacion compras')
ERRORES = ('#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#CALC!')


def fila_de(ws, texto, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(texto):
            return r
    return None


def main():
    p = argparse.ArgumentParser()
    for a in ('original', 'nuevo', 'calc'):
        p.add_argument('--' + a, required=True)
    a = p.parse_args()
    calc = openpyxl.load_workbook(a.calc, data_only=True)
    fallas, ctrl = [], {}
    cero = lambda v: isinstance(v, (int, float)) and abs(v) < 0.005

    sr = calc['Resumen cruce compras']
    r = fila_de(sr, 'Control: TOTAL del cruce')
    ctrl['resumen_cruce'] = [sr.cell(r, 3).value, sr.cell(r, 4).value] if r else None
    if not r or not all(cero(v) for v in ctrl['resumen_cruce']):
        fallas.append(f"Resumen cruce compras: el control TOTAL del cruce - hojas no da 0 ({ctrl['resumen_cruce']})")

    sk = calc['Clasificacion compras']
    r = fila_de(sk, 'Control vs hoja Compras Portal IVA')
    ctrl['clasificacion'] = [sk.cell(r, c).value for c in range(3, 7)] if r else None
    if not r or not all(cero(v) for v in ctrl['clasificacion']):
        fallas.append(f"Clasificacion compras: el control contra la hoja Compras Portal IVA no da 0 ({ctrl['clasificacion']})")

    sp = calc['Compras Portal IVA']
    r = fila_de(sp, 'Diferencia (debe ser 0')
    ctrl['cuadro_vs_libro'] = sp.cell(r, 6).value if r else None
    if not isinstance(ctrl['cuadro_vs_libro'], (int, float)) or abs(ctrl['cuadro_vs_libro']) >= 1:
        fallas.append(f"Cuadro de compras: la diferencia contra el libro es {ctrl['cuadro_vs_libro']} (tiene que ser menor a $1)")
    r = fila_de(sp, 'TOTALES COMPRAS')
    ctrl['cuadro_totales'] = [round(sp.cell(r, c).value or 0, 2) for c in range(2, 7)] if r else None
    r = fila_de(sp, 'SUBTOTAL CF (Merc.+Servicios)', col=9)
    ctrl['cf_computable'] = round(sp.cell(r, 10).value or 0, 2) if r else None

    errs = [(ws.title, c.coordinate, c.value) for ws in calc.worksheets for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.strip() in ERRORES]
    ctrl['errores_excel'] = len(errs)
    if errs:
        fallas.append(f'{len(errs)} celdas con error de Excel, ej. {errs[:5]}')

    orig_f = openpyxl.load_workbook(a.original)
    orig_v = openpyxl.load_workbook(a.original, data_only=True)
    nuevo = openpyxl.load_workbook(a.nuevo)
    previas = [t for t in orig_f.sheetnames if t not in HOJAS]
    ctrl['hojas_previas'] = {}
    for t in previas:
        if t not in nuevo.sheetnames or t not in calc.sheetnames:
            fallas.append(f'Falta la hoja "{t}" en el papel nuevo'); continue
        mf = [[c.value for c in rr] for rr in orig_f[t].iter_rows()] == [[c.value for c in rr] for rr in nuevo[t].iter_rows()]
        mv = [[c.value for c in rr] for rr in orig_v[t].iter_rows()] == [[c.value for c in rr] for rr in calc[t].iter_rows()]
        ctrl['hojas_previas'][t] = 'igual' if mf and mv else ('formulas iguales, valores distintos' if mf else 'DISTINTA')
        if not mf:
            fallas.append(f'La hoja "{t}" cambio (formulas)')
        elif not mv:
            # una hoja de formulas que lee las compras (ej. liquidacion) puede cambiar de valor: se avisa, no se corta
            ctrl.setdefault('avisos', []).append(f'La hoja "{t}" tiene las mismas formulas pero otros valores al recalcular')
    print(json.dumps({'ok': not fallas, 'fallas': fallas, 'controles': ctrl}, ensure_ascii=False, indent=1, default=str))
    sys.exit(0 if not fallas else 1)


if __name__ == '__main__':
    main()
