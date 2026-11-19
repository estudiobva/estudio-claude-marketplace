"""Copia los EERR de <origen> a <destino> con las fórmulas de la hoja EERR reemplazadas por valores (pycel), para
generar Balances cuando Excel no puede recalcular/guardar (ej. una simulación de armado-eerr).
Uso: python calcular_eerr.py <carpeta origen> <carpeta destino>   (requiere: pip install pycel)"""
import sys, glob, os, openpyxl
from pycel import ExcelCompiler
src, dst = sys.argv[1], sys.argv[2]
os.makedirs(dst, exist_ok=True)
for f in sorted(glob.glob(os.path.join(src, "EERR *.xlsx"))):
    xl = ExcelCompiler(filename=f)
    wb = openpyxl.load_workbook(f)
    ws = wb["EERR"]; n = err = 0
    for row in ws.iter_rows(min_row=8, max_row=70, max_col=15):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                try:
                    v = xl.evaluate(f"EERR!{c.coordinate}")
                    c.value = v if isinstance(v, (int, float)) else None; n += 1
                except Exception:
                    c.value = None; err += 1
    wb.save(os.path.join(dst, os.path.basename(f)))
    print(os.path.basename(f), "fórmulas calculadas:", n, "errores:", err)
