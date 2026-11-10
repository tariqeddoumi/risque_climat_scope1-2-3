"""Produit les captures d'écran du canevas de collecte pour le guide utilisateur.

Chaque capture est rendue par LibreOffice à partir d'une copie du canevas réel
(zone d'impression limitée, autres onglets masqués), puis rognée et annotée
(repères numérotés, liste déroulante simulée). Deux classeurs de démonstration
sont aussi produits : l'un avec des erreurs volontaires, l'autre corrigé.

Usage : python3 scripts/build-guide-captures.py <canevas.xlsx> <dossier_sortie>
"""
import math
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from PIL import Image, ImageDraw, ImageFont, ImageChops

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, 'assets', 'groupe-bcp-logo.png')
ORANGE, BRUN, ROUGE = (227, 123, 5), (70, 32, 22), (179, 38, 30)
FONT_B = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FONT_R = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
SOFFICE = None
for cand in ('soffice', 'libreoffice'):
    if shutil.which(cand):
        SOFFICE = cand
        break
SKILL_SOFFICE = '/root/.claude/skills/synced/9df9f9fd-c5eb-4a25-a54f-3634d6250208_e294bec3-e9cc-4010-ae1a-3d08cd0d3cfd/docx/scripts/office/soffice.py'

FR = Side(style='thin', color='B8A890')


def zones(ws):
    z = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        for k in ('COMMENT REMPLIR', 'CAS POSSIBLES', 'ERREURS', 'VOTRE SAISIE'):
            if v.startswith(k):
                z[k] = r
    z['header'] = z['VOTRE SAISIE'] + 1
    z['first'] = z['VOTRE SAISIE'] + 3
    return z


# ---------------------------------------------------------------- démonstrations
def fill_demo(path_in, path_out, corrected):
    wb = openpyxl.load_workbook(path_in)
    idw = wb['Identification']
    vals = [2025, 'BP Tanger-Tétouan', 'Agence', '4896', 'Agence Al Boughaz', 'Tanger', 'Ce site uniquement', 180, 9,
            'Locataire', 0, 410, 'Chef d\'agence', '15/01/2026']
    for i, v in enumerate(vals):
        idw.cell(5 + i, 2, v)
    ws = wb['1-Électricité']
    f = zones(ws)['first']
    H = headers(ws)
    K = ['Numéro de compteur ou de police', 'Usage du compteur', 'Mois', 'Quantité (kWh)', 'Montant HT (MAD)', 'Type de contrat', 'Statut', 'Méthode', 'Pièce justificative', 'Commentaire']
    RS = 'Réseau standard (ONEE ou distributeur)'
    rows = [
        ['123456', 'Site entier', 'Janvier', 2850, 3420, RS, 'Valeur', 'Mesuré (facture, compteur, pesée)', 'Facture Amendis 01-2025', ''],
        ['123456', 'Site entier', 'Février', 2610, 3132, RS, 'Valeur', 'Mesuré (facture, compteur, pesée)', 'Facture Amendis 02-2025', ''],
        ['123456', 'Site entier', 'Mars', 2700, 3240, RS, '', 'Mesuré (facture, compteur, pesée)', 'Facture Amendis 03-2025', ''],
        ['123456', 'Site entier', 'Avril', '2 480 kWh', '', RS, 'Valeur', 'Mesuré (facture, compteur, pesée)', 'Facture Amendis 04-2025', ''],
        ['Onduleur PV', 'Production photovoltaïque', 'Année entière', '', '', 'Autoproduction (loi 82-21)', 'Valeur', '', '', '619 panneaux'],
        ['', 'Recharge véhicules électriques', 'Année entière', 0, '', 'Sans objet', 'Néant (aucune activité)', '', '', 'Pas de véhicule électrique'],
    ]
    if corrected:
        rows[2][6] = 'Valeur'
        rows[3][3] = 2480
        rows[4] = ['Onduleur PV', 'Production photovoltaïque', 'Année entière', '', '', 'Sans objet', 'Néant (aucune activité)', '', '', 'Pas de panneaux solaires sur le site']
        rows[5][3] = ''
    for i, row in enumerate(rows):
        for k, v in zip(K, row):
            if v != '':
                ws.cell(f + i, H[k], v)
    ws2 = wb['2-Carburants et gaz']
    f2 = zones(ws2)['first']; H2 = headers(ws2)
    for k, v in [('Usage', 'Groupe électrogène'), ('Statut', 'Néant (aucune activité)'), ('Commentaire', 'Pas de groupe électrogène')]:
        ws2.cell(f2, H2[k], v)
    wb.save(path_out)


def fill_examples(path_in, path_out):
    """Recopie les lignes d'exemple de chaque onglet dans sa zone de saisie (démonstration de l'export IFC)."""
    wb = openpyxl.load_workbook(path_in)
    fill_ident = [2025, 'BP Tanger-Tétouan', 'Agence', '4896', 'Agence Al Boughaz', 'Tanger', 'Ce site uniquement', 400, 100, 'Locataire', 0, 1200, 'Chef d\'agence', '15/01/2026']
    for i, v in enumerate(fill_ident):
        wb['Identification'].cell(5 + i, 2, v)
    for ws in wb.worksheets:
        if not ws.title[0].isdigit():
            continue
        ex = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == 'Ex.']
        f = zones(ws)['first']
        for k, r in enumerate(ex):
            for c in range(2, ws.max_column):
                ws.cell(f + k, c).value = ws.cell(r, c).value
    wb.save(path_out)


def headers(ws):
    z = zones(ws)
    return {ws.cell(z['header'], c).value: c for c in range(1, ws.max_column + 1)}


# ---------------------------------------------------------------- rendu
def est_row_height(ws, r, c1, c2):
    h = ws.row_dimensions[r].height
    if h:
        return h
    lines = 1
    for c in range(c1, c2 + 1):
        v = ws.cell(r, c).value
        if v is None or (isinstance(v, str) and v.startswith('=')):
            continue
        w = ws.column_dimensions[L(c)].width or 8.43
        size = (ws.cell(r, c).font.sz or 10)
        cap = max(1.0, w * 1.15 * 10 / size)
        lines = max(lines, math.ceil(len(str(v)) / cap))
    return max(15, 12.5 * lines + 3)


def render(path, sheet, rng, name, logo_anchor=None, scale=2.2):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    c1s, c2s = rng.split(':')
    import re
    m1 = re.match(r'([A-Z]+)(\d+)', c1s); m2 = re.match(r'([A-Z]+)(\d+)', c2s)
    c1, r1, c2, r2 = CI(m1.group(1)), int(m1.group(2)), CI(m2.group(1)), int(m2.group(2))
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = est_row_height(ws, r, c1, c2)
    # cadre extérieur pour un rognage exact
    for c in range(c1, c2 + 1):
        for r, side in ((r1, 'top'), (r2, 'bottom')):
            cell = ws.cell(r, c); b = cell.border
            cell.border = Border(left=b.left, right=b.right, top=FR if side == 'top' else b.top, bottom=FR if side == 'bottom' else b.bottom)
    for r in range(r1, r2 + 1):
        for c, side in ((c1, 'left'), (c2, 'right')):
            cell = ws.cell(r, c); b = cell.border
            cell.border = Border(top=b.top, bottom=b.bottom, left=FR if side == 'left' else b.left, right=FR if side == 'right' else b.right)
    if logo_anchor:
        img = XLImage(LOGO); img.width, img.height = 92, 48; ws.add_image(img, logo_anchor)
    for other in wb.worksheets:
        if other.title != sheet:
            other.sheet_state = 'hidden'
    wb.active = wb.worksheets.index(ws)
    ws.print_area = rng
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = ws.page_margins.bottom = 0.2
    ws.oddHeader.center.text = ''; ws.oddFooter.center.text = ''
    tmp = tempfile.mkdtemp()
    x = os.path.join(tmp, name + '.xlsx'); wb.save(x)
    subprocess.run([sys.executable, SKILL_SOFFICE, '--headless', '--convert-to', 'pdf', '--outdir', tmp, x], capture_output=True)
    pdf = os.path.join(tmp, name + '.pdf')
    subprocess.run(['pdftoppm', '-png', '-r', '220', '-singlefile', pdf, os.path.join(tmp, 'p')], capture_output=True)
    im = Image.open(os.path.join(tmp, 'p.png')).convert('RGB')
    bg = Image.new('RGB', im.size, (255, 255, 255))
    bbox = ImageChops.difference(im, bg).getbbox()
    im = im.crop(bbox)
    # géométrie pour les annotations
    colw = [(ws.column_dimensions[L(c)].width or 8.43) * 7 + 5 for c in range(c1, c2 + 1)]
    rowh = [ws.row_dimensions[r].height for r in range(r1, r2 + 1)]
    geo = dict(c1=c1, r1=r1, colw=colw, rowh=rowh, W=im.size[0], H=im.size[1])
    shutil.rmtree(tmp, ignore_errors=True)
    return im, geo


def col_x(geo, col, where=0.5):
    i = CI(col) - geo['c1']
    tot = sum(geo['colw'])
    return (sum(geo['colw'][:i]) + geo['colw'][i] * where) / tot * geo['W']


def row_y(geo, row, where=0.5):
    i = row - geo['r1']
    tot = sum(geo['rowh'])
    return (sum(geo['rowh'][:i]) + geo['rowh'][i] * where) / tot * geo['H']


def font(sz, bold=True):
    return ImageFont.truetype(FONT_B if bold else FONT_R, sz)


def badge(d, x, y, n, r=22, col=ORANGE):
    d.ellipse([x - r, y - r, x + r, y + r], fill=col, outline=(255, 255, 255), width=3)
    f = font(int(r * 1.15))
    tw = d.textlength(str(n), font=f)
    d.text((x - tw / 2, y - r * 0.68), str(n), font=f, fill=(255, 255, 255))


def annotate(im, geo, zones_left=(), cols_top=(), cells=()):
    """zones_left: [(n, row_start, row_end)] ; cols_top: [(n, col)] ; cells: [(n, col, row)]"""
    ml = 90 if zones_left else 0
    mt = 80 if cols_top else 0
    canvas = Image.new('RGB', (im.size[0] + ml, im.size[1] + mt), (255, 255, 255))
    canvas.paste(im, (ml, mt))
    d = ImageDraw.Draw(canvas)
    for n, ra, rb in zones_left:
        y1 = mt + row_y(geo, ra, 0.05); y2 = mt + row_y(geo, rb, 0.95)
        x = ml - 22
        d.line([(x, y1), (x, y2)], fill=ORANGE, width=6)
        d.line([(x, y1), (x + 14, y1)], fill=ORANGE, width=6)
        d.line([(x, y2), (x + 14, y2)], fill=ORANGE, width=6)
        badge(d, 38, (y1 + y2) / 2, n, 28)
    for n, col in cols_top:
        x = ml + col_x(geo, col)
        d.line([(x, 52), (x, mt + 4)], fill=ORANGE, width=4)
        badge(d, x, 30, n, 24)
    for n, col, row in cells:
        x = ml + col_x(geo, col, 0.95); y = mt + row_y(geo, row)
        badge(d, x, y, n, 22, ROUGE)
    return canvas


def dropdown(im, geo, col, row, options, chosen):
    d = ImageDraw.Draw(im)
    x0 = col_x(geo, col, 0); x1 = col_x(geo, col, 1)
    y0 = row_y(geo, row, 0); y1 = row_y(geo, row, 1)
    # bouton de liste
    bw = (y1 - y0)
    d.rectangle([x1 - bw, y0, x1, y1], fill=(235, 235, 235), outline=(150, 150, 150), width=2)
    cx, cy = x1 - bw / 2, (y0 + y1) / 2
    d.polygon([(cx - bw * 0.22, cy - bw * 0.1), (cx + bw * 0.22, cy - bw * 0.1), (cx, cy + bw * 0.18)], fill=(60, 60, 60))
    f = font(30, False)
    ih = 48
    wbox = max(x1 - x0, max(d.textlength(o, font=f) for o in options) + 40)
    top = y1 + 2
    d.rectangle([x0 + 6, top + 6, x0 + wbox + 6, top + ih * len(options) + 6], fill=(200, 200, 200))
    d.rectangle([x0, top, x0 + wbox, top + ih * len(options)], fill=(255, 255, 255), outline=(120, 120, 120), width=2)
    for i, o in enumerate(options):
        yy = top + i * ih
        if o == chosen:
            d.rectangle([x0 + 2, yy + 2, x0 + wbox - 2, yy + ih - 2], fill=(253, 235, 208))
            d.rectangle([x0 + 2, yy + 2, x0 + 10, yy + ih - 2], fill=ORANGE)
        d.text((x0 + 20, yy + 8), o, font=f, fill=(40, 40, 40))
    return im


def save(im, name, maxw=2400):
    if im.size[0] > maxw:
        im = im.resize((maxw, int(im.size[1] * maxw / im.size[0])), Image.LANCZOS)
    im = im.convert('P', palette=Image.ADAPTIVE, colors=256)  # captures plates : palette sans perte visible, fichier 3 à 4 fois plus léger
    im.save(os.path.join(OUT, name + '.png'), optimize=True)
    print(name, im.size)


# ---------------------------------------------------------------- captures
tmpd = tempfile.mkdtemp()
DEMO = os.path.join(OUT, 'demo_avec_erreurs.xlsx')
DEMO_OK = os.path.join(OUT, 'demo_corrige.xlsx')
fill_demo(SRC, DEMO, False)
fill_demo(SRC, DEMO_OK, True)

wb0 = openpyxl.load_workbook(SRC)
rd = wb0['Lisez-moi']
secs = {}
for r in range(1, rd.max_row + 1):
    v = rd.cell(r, 1).value
    if isinstance(v, str) and v[:2] in ('1.', '2.', '3.', '4.', '5.', '6.', '7.'):
        secs[v[:1]] = r
last_rd = rd.max_row

im, g = render(SRC, 'Lisez-moi', f'A1:E{secs["3"] - 2}', 'lisez1', 'E1')
save(annotate(im, g, zones_left=[(1, 1, 2), (2, secs['1'], secs['2'] - 2), (3, secs['2'], secs['3'] - 2)]), '01_lisez_moi_etapes_regles')
im, g = render(SRC, 'Lisez-moi', f'A{secs["3"]}:D{secs["5"] - 2}', 'lisez2')
save(im, '02_lisez_moi_statuts_methodes')
im, g = render(SRC, 'Lisez-moi', f'A{secs["5"]}:D{secs["7"] - 2}', 'lisez3')
save(im, '03_lisez_moi_onglets_couleurs')
im, g = render(SRC, 'Lisez-moi', f'A{secs["7"]}:D{last_rd - 2}', 'lisez4')
save(im, '03b_lisez_moi_nouveautes')

IDR = wb0['Identification'].max_row
im, g = render(SRC, 'Identification', f'A1:D{IDR}', 'id_vide', 'D1')
save(annotate(im, g, cols_top=[(1, 'A'), (2, 'B'), (3, 'C')]), '04_identification_vide')
im, g = render(DEMO, 'Identification', f'A1:D{IDR}', 'id_rempli', 'D1')
save(im, '05_identification_remplie')

ws = wb0['1-Électricité']; z = zones(ws); nc = ws.max_column; LC = L(nc)
im, g = render(SRC, '1-Électricité', f'A1:{LC}{z["first"] + 5}', 'anat')
save(annotate(im, g, zones_left=[(1, 1, 2), (2, z['COMMENT REMPLIR'], z['CAS POSSIBLES'] - 2), (3, z['CAS POSSIBLES'], z['ERREURS'] - 2),
                                  (4, z['ERREURS'], z['VOTRE SAISIE'] - 2), (5, z['VOTRE SAISIE'], z['first'] + 5)]), '06_anatomie_onglet')
im, g = render(SRC, '1-Électricité', f'A1:{L(nc + 1)}{z["CAS POSSIBLES"] - 2}', 'haut', f'{L(nc + 1)}1')
save(im, '07_electricite_bandeau_consignes')
im, g = render(SRC, '1-Électricité', f'A{z["CAS POSSIBLES"]}:{LC}{z["ERREURS"] - 2}', 'cas')
save(annotate(im, g, zones_left=[(1, z['CAS POSSIBLES'] + 1, z['CAS POSSIBLES'] + 1), (2, z['CAS POSSIBLES'] + 2, z['CAS POSSIBLES'] + 2), (3, z['CAS POSSIBLES'] + 3, z['CAS POSSIBLES'] + 4)]), '08_electricite_cas_possibles')
im, g = render(SRC, '1-Électricité', f'A{z["VOTRE SAISIE"]}:{LC}{z["first"] + 3}', 'saisie_vide')
cols = [(i + 1, L(c)) for i, c in enumerate(range(2, nc + 1))]
save(annotate(im, g, cols_top=cols), '09_electricite_zone_saisie')
im, g = render(DEMO, '1-Électricité', f'A{z["header"]}:{LC}{z["first"] + 6}', 'saisie_demo')
f = z['first']
save(annotate(im, g, cells=[(1, LC, f + 2), (2, LC, f + 3), (3, LC, f + 4), (4, LC, f + 5)]), '10_electricite_saisie_alertes')
im, g = render(DEMO_OK, '1-Électricité', f'A{z["header"]}:{LC}{z["first"] + 6}', 'saisie_ok')
save(im, '11_electricite_saisie_corrigee')
im, g = render(DEMO, '1-Électricité', f'A{z["header"]}:{LC}{z["first"] + 9}', 'dd')
im = dropdown(im, g, L(headers(wb0['1-Électricité'])['Usage du compteur']), f + 1, ['Site entier', 'Partie du site', 'Salle informatique (sous-compteur)', 'Partagé avec un tiers', 'Recharge véhicules électriques', 'Production photovoltaïque'], 'Site entier')
save(im, '12_liste_deroulante')
im, g = render(SRC, '1-Électricité', f'A{z["ERREURS"]}:{LC}{z["VOTRE SAISIE"] - 2}', 'err')
save(im, '13_electricite_erreurs')

for sheet, code in [('2-Carburants et gaz', '20'), ('3-Climatisation', '21'), ('4-Achats', '22'), ('5-Équipements', '23'), ('6-Déchets', '24'),
                    ('7-Déplacements pro', '25'), ('8-Domicile et visiteurs', '26'), ('9-Transport fonds-courrier', '27')]:
    ws = wb0[sheet]; z = zones(ws); LC = L(ws.max_column)
    im, g = render(SRC, sheet, f'A{z["COMMENT REMPLIR"]}:{LC}{z["VOTRE SAISIE"] - 2}', 'p' + code)
    save(im, f'{code}_{sheet.split("-", 1)[1].lower().replace(" ", "_").replace("é", "e").replace("è", "e")}')
    im, g = render(SRC, sheet, f'A{z["header"]}:{LC}{z["header"] + 1}', 'h' + code)
    save(im, f'{code}b_colonnes')

ctw = wb0['Contrôle']
coh = [r for r in range(1, ctw.max_row + 1) if isinstance(ctw.cell(r, 1).value, str) and ctw.cell(r, 1).value.startswith('CONTRÔLES DE COHÉRENCE')][0]
im, g = render(DEMO, 'Contrôle', f'A1:F{coh - 2}', 'ctl', 'F1')
save(im, '30_controle_alertes')
im, g = render(DEMO_OK, 'Contrôle', f'A1:F{coh - 2}', 'ctlok', 'F1')
save(im, '31_controle_pret')
im, g = render(DEMO_OK, 'Contrôle', f'A{coh}:E{ctw.max_row}', 'coh')
save(im, '32_controle_coherence')
DEMO_EX = os.path.join(OUT, 'demo_exemples.xlsx')
fill_examples(SRC, DEMO_EX)
im, g = render(DEMO_EX, 'Export IFC', 'A1:I44', 'exp', 'I1')
save(im, '33_export_ifc')
shutil.rmtree(tmpd, ignore_errors=True)
