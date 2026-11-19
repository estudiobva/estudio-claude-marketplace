"""Imprime en JSON las solapas de un Excel (sin cargar las celdas).

    python solapas.py <archivo.xlsx>
"""
import json
import re
import sys
import zipfile
from xml.sax.saxutils import unescape


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    with zipfile.ZipFile(sys.argv[1]) as z:
        xml = z.read('xl/workbook.xml').decode('utf-8')
    nombres = [unescape(n, {'&quot;': '"', '&apos;': "'"}) for n in re.findall(r'<(?:\w+:)?sheet\b[^>]*?\bname="([^"]*)"', xml)]
    print(json.dumps({'solapas': nombres}, ensure_ascii=False))


if __name__ == '__main__':
    main()
