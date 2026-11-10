"""Visuels de la présentation « guide utilisateur » : messages d'erreur Excel simulés posés sur les captures du canevas.

Usage : python3 scripts/build-ppt-visuels.py <dossier_captures> <dossier_sortie>
Les messages reproduisent le texte paramétré dans le canevas (validation des données) ; la boîte de dialogue est une illustration.
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFilter, ImageFont

CAPS, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
FB, FR = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'


def wrap(d, text, font, width):
    words, lines, cur = text.split(), [], ''
    for w in words:
        t = (cur + ' ' + w).strip()
        if d.textlength(t, font=font) <= width:
            cur = t
        else:
            lines.append(cur); cur = w
    if cur: lines.append(cur)
    return lines


def dialog(base, out, title, message, buttons, icon='x', crop=None, scale=1.0):
    im = Image.open(os.path.join(CAPS, base + '.png')).convert('RGB')
    if crop:
        im = im.crop(crop)
    W, H = im.size
    dim = Image.new('RGB', im.size, (0, 0, 0))
    im = Image.blend(im, dim, 0.18)
    dw = int(min(W * 0.8, H * 1.55) * scale)
    ft, fm, fb = ImageFont.truetype(FR, int(dw * 0.032)), ImageFont.truetype(FR, int(dw * 0.034)), ImageFont.truetype(FR, int(dw * 0.031))
    tmp = ImageDraw.Draw(im)
    lines = wrap(tmp, message, fm, dw * 0.70)
    lh = int(fm.size * 1.45)
    th = int(dw * 0.075)
    dh = th + int(dw * 0.07) + max(len(lines) * lh, int(dw * 0.11)) + int(dw * 0.16)
    x0, y0 = (W - dw) // 2, (H - dh) // 2
    sh = Image.new('RGBA', im.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rectangle([x0 + 14, y0 + 18, x0 + dw + 14, y0 + dh + 18], fill=(0, 0, 0, 110))
    sh = sh.filter(ImageFilter.GaussianBlur(14))
    im = Image.alpha_composite(im.convert('RGBA'), sh)
    d = ImageDraw.Draw(im)
    d.rectangle([x0, y0, x0 + dw, y0 + dh], fill=(255, 255, 255), outline=(150, 150, 150), width=2)
    d.rectangle([x0, y0, x0 + dw, y0 + th], fill=(243, 243, 243))
    d.text((x0 + int(dw * 0.03), y0 + (th - ft.size) // 2 - 2), title, font=ft, fill=(30, 30, 30))
    d.text((x0 + dw - int(dw * 0.055), y0 + (th - ft.size) // 2 - 4), '✕', font=ft, fill=(60, 60, 60))
    # icône
    r = int(dw * 0.045); cx, cy = x0 + int(dw * 0.09), y0 + th + int(dw * 0.07) + r
    if icon == 'x':
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(196, 43, 28))
        k = int(r * 0.45); d.line([cx - k, cy - k, cx + k, cy + k], fill='white', width=max(4, r // 5)); d.line([cx - k, cy + k, cx + k, cy - k], fill='white', width=max(4, r // 5))
    else:
        d.polygon([(cx, cy - r), (cx + r, cy + r * 0.85), (cx - r, cy + r * 0.85)], fill=(232, 168, 0))
        fi = ImageFont.truetype(FB, int(r * 1.1)); d.text((cx - d.textlength('!', font=fi) / 2, cy - r * 0.45), '!', font=fi, fill='black')
    ty = y0 + th + int(dw * 0.06)
    for i, ln in enumerate(lines):
        d.text((x0 + int(dw * 0.18), ty + i * lh), ln, font=fm, fill=(20, 20, 20))
    bw, bh = int(dw * 0.25), int(dw * 0.065)
    bx = x0 + dw - int(dw * 0.03) - len(buttons) * (bw + int(dw * 0.02)) + int(dw * 0.02)
    by = y0 + dh - int(dw * 0.03) - bh
    for i, b in enumerate(buttons):
        xx = bx + i * (bw + int(dw * 0.02))
        d.rectangle([xx, by, xx + bw, by + bh], fill=(253, 253, 253), outline=(0, 120, 215) if i == 0 else (170, 170, 170), width=3 if i == 0 else 2)
        d.text((xx + (bw - d.textlength(b, font=fb)) / 2, by + (bh - fb.size) / 2 - 2), b, font=fb, fill=(20, 20, 20))
    im.convert('RGB').save(os.path.join(OUT, out + '.png'), optimize=True)
    print(out, im.size)


dialog('12_liste_deroulante', 'v_liste_refusee', 'Valeur non autorisée', 'Choisissez une valeur dans la liste déroulante.', ['Recommencer', 'Annuler', 'Aide'], crop=(0, 0, 1100, 398))
dialog('09_electricite_zone_saisie', 'v_nombre_attendu', 'Nombre attendu', 'Saisissez un nombre positif, sans unité ni espace (ex. 2850 ou 12.5).', ['Recommencer', 'Annuler', 'Aide'], crop=(0, 0, 1000, 328), scale=0.95)
dialog('08_electricite_cas_possibles', 'v_feuille_protegee', 'Microsoft Excel', 'La cellule ou le graphique que vous essayez de modifier se trouve sur une feuille protégée. Pour apporter une modification, ôtez la protection de la feuille.', ['OK'], icon='warn', crop=(0, 0, 1400, 567))
