"""Produit les captures d'écran du canevas de collecte pour le guide utilisateur.

Usage : python3 scripts/build-guide-captures.py <canevas.xlsx> <dossier_sortie>
"""
import os
import shutil
import sys
import tempfile

import openpyxl
from openpyxl.utils import get_column_letter as L

import capture_lib
from capture_lib import zones, fill_demo, fill_examples, headers, render, annotate, dropdown, save

SRC, OUT = sys.argv[1], sys.argv[2]
capture_lib.set_out(OUT)
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
