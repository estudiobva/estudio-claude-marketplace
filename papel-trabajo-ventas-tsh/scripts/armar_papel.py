"""Proceso TSH / IIBB: pasos 2 a 4 y 6 del papel de ventas en una sola pasada.

    python armar_papel.py <papel.xlsx> <portal.xlsx> <mis_comprobantes.xlsx> [--excepcion=...]

Corre, en este orden y sobre el MISMO libro en memoria:
    copiar_solapa (Ventas portal iva), copiar_solapa (Ventas MCR), totales_mcr,
    depurar_pv y cuadro_resumen.
Cada script hace exactamente lo mismo que corrido suelto. La diferencia es que el papel
se abre una sola vez y se guarda una sola vez al final, en lugar de una vez por paso.
Con sociedades grandes (~77.000 comprobantes por mes) cada apertura y
guardado tarda varios minutos.

Si un paso falla, no se guarda nada y el papel queda como estaba.
"""
import os
import sys
import time

import openpyxl

import copiar_solapa
import cuadro_resumen
import depurar_pv
import totales_mcr

_cargar_original = openpyxl.load_workbook


def main():
    pos = [a for a in sys.argv[1:] if not a.startswith('--')]
    extra = [a for a in sys.argv[1:] if a.startswith('--')]
    if len(pos) != 3:
        sys.exit(__doc__)
    papel, portal, mc = pos
    clave = os.path.normcase(os.path.abspath(papel))

    t0 = time.time()
    wb = _cargar_original(papel)
    print(f'[armar_papel] papel abierto en {time.time() - t0:.0f} s', file=sys.stderr, flush=True)

    def cargar(ruta, *a, **k):
        if os.path.normcase(os.path.abspath(ruta)) == clave:
            return wb
        return _cargar_original(ruta, *a, **k)

    guardar_original = wb.save
    wb.save = lambda *a, **k: None  # los pasos "guardan" en memoria; se guarda una vez al final
    for mod in (copiar_solapa, totales_mcr, depurar_pv, cuadro_resumen):
        mod.load_workbook = cargar

    pasos = [
        (copiar_solapa, [portal, papel, 'Ventas portal iva']),
        (copiar_solapa, [mc, papel, 'Ventas MCR']),
        (totales_mcr, [papel]),
        (depurar_pv, [papel] + extra),
        (cuadro_resumen, [papel]),
    ]
    for mod, args in pasos:
        t = time.time()
        sys.argv = [mod.__file__] + args
        try:
            mod.main()
        except SystemExit as e:
            if e.code not in (None, 0):
                print(f'[armar_papel] fallo {mod.__name__}: no se guarda el papel', file=sys.stderr)
                raise
        sys.stdout.flush()
        print(f'[armar_papel] {mod.__name__} en {time.time() - t:.0f} s', file=sys.stderr, flush=True)

    t = time.time()
    del wb.save
    guardar_original(papel)
    print(f'[armar_papel] guardado en {time.time() - t:.0f} s; total {time.time() - t0:.0f} s', file=sys.stderr)


if __name__ == '__main__':
    main()
