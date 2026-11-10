"""Captures d'écran de l'outil émissions financées pour sa note d'utilisation.

Usage : python3 scripts/build-pcaf-captures.py <outil.xlsx> <dossier_sortie>
"""
import sys

import openpyxl
from openpyxl.utils import get_column_letter as L

import capture_lib
from capture_lib import render, annotate, save

SRC, OUT = sys.argv[1], sys.argv[2]
capture_lib.set_out(OUT)
wb = openpyxl.load_workbook(SRC)


def find(ws, text, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(text):
            return r
    raise KeyError(text)


rd = wb['Lisez-moi']
im, g = render(SRC, 'Lisez-moi', f'A1:E{find(rd, "3.") - 2}', 'p_lisez', 'E1')
save(im, 'p01_lisez_moi')

pm = wb['Paramètres']
sect = find(pm, 'FACTEURS SECTORIELS', 2)
im, g = render(SRC, 'Paramètres', f'A1:I{sect - 2}', 'p_param1', 'I1')
save(im, 'p02_parametres_general')
im, g = render(SRC, 'Paramètres', f'A{sect}:H{sect + 12}', 'p_param2')
save(im, 'p03_parametres_secteurs')

pf = wb['Portefeuille']
hdr = {pf.cell(5, c).value: c for c in range(1, pf.max_column + 1) if pf.cell(5, c).value}
last_input = hdr['Commentaire']
calc0 = hdr['Facteur d\'attribution']
ctl = hdr['Contrôle']
ent0, ent1 = hdr['Société cotée ?'], hdr['Énergie consommée (MWh)']
imm0, veh1 = hdr['Valeur du bien à l\'octroi (MAD)'], hdr['Commentaire']
im, g = render(SRC, 'Portefeuille', f'A1:{L(hdr["Encours au bilan (MAD)"])}20', 'p_pf1')
save(annotate(im, g, zones_left=[(1, 1, 2), (2, 4, 6), (3, 7, 19)]), 'p04_portefeuille_identification')
hide = [L(c) for c in range(4, ent0)]
im, g = render(SRC, 'Portefeuille', f'A4:{L(ent1)}20', 'p_pf2', hide=hide)
save(im, 'p05_portefeuille_entreprises')
hide = [L(c) for c in range(4, imm0)]
im, g = render(SRC, 'Portefeuille', f'A4:{L(veh1)}20', 'p_pf3', hide=hide)
save(im, 'p06_portefeuille_immobilier_vehicules')
hide = [L(c) for c in range(4, calc0)]
im, g = render(SRC, 'Portefeuille', f'A4:{L(ctl)}20', 'p_pf4', hide=hide)
save(im, 'p07_portefeuille_resultats')

sy = wb['Synthèse']
cls_end = find(sy, 'PAR SECTEUR', 2) - 2
im, g = render(SRC, 'Synthèse', f'A1:I{cls_end}', 'p_syn1', 'I1')
save(im, 'p08_synthese_indicateurs_classes')
top = find(sy, 'LES 10 PREMIÈRES', 2)
im, g = render(SRC, 'Synthèse', f'A{top}:H{top + 11}', 'p_syn2')
save(im, 'p09_synthese_top10')
im, g = render(SRC, 'Contrôle', 'A1:D16', 'p_ctl', 'D1')
save(im, 'p10_controle')
