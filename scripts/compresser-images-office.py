"""Allège un fichier Word ou PowerPoint en recompressant ses images, sans toucher au contenu.

Usage : python3 scripts/compresser-images-office.py <entrée.docx|pptx> <sortie> [largeur_max=1600] [couleurs=128]

Chaque image PNG ou JPEG du dossier media est redimensionnée (largeur maximale) et réenregistrée
(PNG en palette optimisée, JPEG qualité 82). Les noms de fichiers sont conservés : aucune relation
ni aucun type de contenu n'est modifié. Une image n'est remplacée que si la nouvelle version est plus légère.
"""
import io
import sys
import zipfile

from PIL import Image

src, dst = sys.argv[1], sys.argv[2]
maxw = int(sys.argv[3]) if len(sys.argv) > 3 else 1600
ncol = int(sys.argv[4]) if len(sys.argv) > 4 else 128

zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
gain = 0
for info in zin.infolist():
    data = zin.read(info.filename)
    low = info.filename.lower()
    if '/media/' in low and low.endswith(('.png', '.jpg', '.jpeg')):
        try:
            im = Image.open(io.BytesIO(data))
            im.load()
            if im.width > maxw:
                im = im.resize((maxw, round(im.height * maxw / im.width)), Image.LANCZOS)
            buf = io.BytesIO()
            if low.endswith('.png'):
                has_alpha = im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info)
                if has_alpha:
                    im = im.convert('RGBA').quantize(colors=ncol, method=Image.FASTOCTREE)
                else:
                    im = im.convert('RGB').quantize(colors=ncol, method=Image.MEDIANCUT)
                im.save(buf, 'PNG', optimize=True)
            else:
                im.convert('RGB').save(buf, 'JPEG', quality=82, optimize=True)
            new = buf.getvalue()
            if len(new) < len(data):
                gain += len(data) - len(new)
                data = new
        except Exception as e:  # image illisible : conservée telle quelle
            print('conservée :', info.filename, e)
    zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    zi.compress_type = zipfile.ZIP_STORED if info.filename == 'mimetype' else zipfile.ZIP_DEFLATED
    zi.external_attr = info.external_attr
    zout.writestr(zi, data)
zout.close()
import os
print(f'{os.path.basename(dst)} : {os.path.getsize(src) / 1e6:.2f} Mo -> {os.path.getsize(dst) / 1e6:.2f} Mo (gain {gain / 1e6:.2f} Mo)')
