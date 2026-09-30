"""Reaplica roundedCorners=1 nos gráficos de um .xlsx (o recálculo pelo LibreOffice grava 0)."""
import sys, zipfile, shutil, os
f = sys.argv[1]; tmp = f + '.tmp'
with zipfile.ZipFile(f) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
    n = 0
    for it in zi.infolist():
        b = zi.read(it.filename)
        if it.filename.startswith('xl/charts/chart'):
            b2 = b.replace(b'<c:roundedCorners val="0"/>', b'<c:roundedCorners val="1"/>')
            if b'roundedCorners' not in b2: b2 = b2.replace(b'<c:chart>', b'<c:roundedCorners val="1"/><c:chart>', 1)
            n += b2 != b; b = b2
        zo.writestr(it, b)
os.replace(tmp, f); print('cantos arredondados em', n, 'gráficos')
