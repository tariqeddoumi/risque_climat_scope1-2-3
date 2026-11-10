"""Outil de calcul des émissions financées (scope 3, catégorie 15) — standard PCAF, approche IFC.

Usage : python3 scripts/build-outil-pcaf-xlsx.py <sortie.xlsx>

Classeur autonome, distinct du canevas de collecte de l'empreinte propre. Charte Groupe BCP.
Les facteurs d'émission fournis sont des VALEURS DE DÉMONSTRATION : ils doivent être remplacés
par des facteurs sourcés (base PCAF, inventaire national, ADEME, EXIOBASE…) avant tout usage.
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1]
ORANGE, BRUN, GRIS, TINT, TINT2, WHITE = 'E37B05', '462016', '3D3D3B', 'F2E8DB', 'FBF6F0', 'FFFFFF'
INPUT, INPUT_FONT, RED, GREEN, GREY, ROSE = 'FDEBD0', '1F3FBF', 'B3261E', '146C4C', 'EDEDED', 'FBE3E1'
VERSION = '1.0'
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'groupe-bcp-logo.png')
thin = Side(style='thin', color='D9C7B0')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')
WRAPC = Alignment(wrap_text=True, vertical='center')
UNLOCK = Protection(locked=False)


def fill(c): return PatternFill('solid', start_color=c, end_color=c)
def font(**k): return Font(name='Arial', **{'size': 10, 'color': GRIS, **k})


def add_logo(ws, anchor):
    img = XLImage(LOGO); img.width, img.height = 92, 48; ws.add_image(img, anchor)


def protect(ws):
    p = ws.protection
    p.sheet = True
    p.insertRows = p.deleteRows = p.formatColumns = p.formatRows = p.formatCells = p.sort = p.autoFilter = False


def merged(ws, r, c1, c2, text, f=None, fl=None, h=None, al=WRAP):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1, text); cell.font = f or font(); cell.alignment = al
    if fl:
        for c in range(c1, c2 + 1): ws.cell(r, c).fill = fill(fl)
    if h: ws.row_dimensions[r].height = h
    return cell


def banner(ws, title, sub, width):
    merged(ws, 1, 1, width, title, font(bold=True, size=16, color=WHITE), ORANGE, 40, WRAPC)
    merged(ws, 2, 1, width, sub, font(color=BRUN), TINT, 44)
    add_logo(ws, f'{L(width + 1)}1'); ws.column_dimensions[L(width + 1)].width = 15
    ws.sheet_view.showGridLines = False


def head(ws, r, c, labels, fl=ORANGE, h=32):
    for i, t in enumerate(labels):
        x = ws.cell(r, c + i, t); x.font = font(bold=True, color=WHITE, size=9); x.fill = fill(fl); x.alignment = WRAPC; x.border = BOX
    ws.row_dimensions[r].height = h


def inp(cell, fmt=None):
    cell.fill = fill(INPUT); cell.font = font(bold=True, color=INPUT_FONT, size=9); cell.border = BOX; cell.protection = UNLOCK
    if fmt: cell.number_format = fmt


wb = Workbook()

# ------------------------------------------------------------------ listes
CLASSES = ['Actions cotées et obligations d\'entreprises', 'Prêts aux entreprises et actions non cotées', 'Financement de projets',
           'Immobilier commercial', 'Crédits immobiliers (habitat)', 'Crédits automobiles', 'Dette souveraine']
LISTS = {
    'classe': CLASSES,
    'ouinon': ['Oui', 'Non'],
    'source_em': ['Rapport ESG ou extra-financier', 'Étude d\'impact (EIES) du projet', 'Outil IFC CEET (projets)', 'Questionnaire client', 'Base de données externe', 'Autre (préciser)'],
    'type_bien': ['Logement', 'Bureaux', 'Commerce', 'Hôtel', 'Entrepôt / logistique', 'Autre'],
    'type_vehicule': ['Voiture essence', 'Voiture diesel', 'Voiture hybride', 'Voiture électrique', 'Utilitaire léger', 'Camion', 'Deux-roues', 'Inconnu'],
    'pays': ['Maroc', 'Côte d\'Ivoire', 'Sénégal', 'Cameroun', 'Mali', 'Guinée', 'Niger', 'Burkina Faso', 'Bénin', 'Togo', 'République centrafricaine', 'Congo', 'Rwanda', 'Madagascar', 'Maurice', 'France', 'Autre'],
}
SECTEURS = [
    # secteur, FE S1+S2 (tCO2e / M MAD de CA), FE S3, FE actifs (tCO2e / M MAD d'actif total)
    ('Agriculture, élevage, pêche', 160, 120, 90),
    ('Industries extractives (mines, phosphates)', 220, 140, 70),
    ('Industrie agroalimentaire', 70, 160, 50),
    ('Textile, cuir, habillement', 45, 110, 35),
    ('Chimie, parachimie, engrais', 300, 200, 120),
    ('Ciment, chaux, matériaux de construction', 650, 120, 260),
    ('Sidérurgie, métallurgie', 380, 160, 150),
    ('Production d\'électricité d\'origine fossile', 700, 90, 140),
    ('Énergies renouvelables', 10, 40, 4),
    ('Raffinage et distribution de carburants', 60, 900, 40),
    ('Eau, assainissement, déchets', 200, 60, 50),
    ('Bâtiment et travaux publics', 55, 180, 40),
    ('Promotion et activités immobilières', 15, 90, 4),
    ('Commerce de gros et de détail', 12, 140, 15),
    ('Transport routier et logistique', 150, 60, 80),
    ('Transport aérien et maritime', 320, 80, 90),
    ('Tourisme, hôtellerie, restauration', 40, 70, 15),
    ('Télécommunications et informatique', 10, 45, 5),
    ('Activités financières et assurance', 2, 10, 0.3),
    ('Santé, éducation, services sociaux', 18, 50, 10),
    ('Administration publique', 20, 60, 8),
    ('Autres services', 10, 50, 8),
]
LISTS['secteur'] = [s[0] for s in SECTEURS]

ls = wb.active; ls.title = 'Listes'
RANGES = {}
for j, (k, vals) in enumerate(LISTS.items(), start=1):
    c = ls.cell(1, j, k); c.font = font(bold=True, color=WHITE); c.fill = fill(ORANGE)
    for i, v in enumerate(vals, start=2):
        ls.cell(i, j, v).font = font()
    ls.column_dimensions[L(j)].width = 40
    RANGES[k] = f'Listes!${L(j)}$2:${L(j)}${len(vals) + 1}'


def dv_list(ws, key, rng):
    dv = DataValidation(type='list', formula1='=' + RANGES[key], allow_blank=True, showErrorMessage=True, errorTitle='Valeur non autorisée', error='Choisissez une valeur dans la liste déroulante.')
    ws.add_data_validation(dv); dv.add(rng)


def dv_num(ws, rng, lo=0, hi=None, msg='Saisissez un nombre positif, sans unité ni espace.'):
    if hi is None:
        dv = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1=str(lo), allow_blank=True, showErrorMessage=True, errorTitle='Nombre attendu', error=msg)
    else:
        dv = DataValidation(type='decimal', operator='between', formula1=str(lo), formula2=str(hi), allow_blank=True, showErrorMessage=True, errorTitle='Nombre attendu', error=msg)
    ws.add_data_validation(dv); dv.add(rng)


# ------------------------------------------------------------------ Paramètres
pm = wb.create_sheet('Paramètres')
for col, w in zip('ABCDEFGH', [4, 46, 18, 18, 18, 18, 36, 12]): pm.column_dimensions[col].width = w
banner(pm, 'Paramètres et facteurs d\'émission', 'À renseigner par l\'équipe Climat avant le calcul. Les valeurs en rouge sont des VALEURS DE DÉMONSTRATION, non sourcées : remplacez-les par des facteurs documentés (base PCAF si la banque y adhère, inventaire national, ADEME, EXIOBASE…), puis passez « Facteurs validés » à Oui.', 8)
r = 4
P = {}
def param(label, key, value, fmt='#,##0.00', helpt='', demo=False, listkey=None):
    global r
    a = pm.cell(r, 2, label); a.font = font(bold=True); a.border = BOX; a.alignment = WRAP
    c = pm.cell(r, 3, value); inp(c, fmt)
    if demo: c.font = font(bold=True, italic=True, color=RED, size=9)
    if listkey: dv_list(pm, listkey, f'C{r}')
    h = pm.cell(r, 4, helpt); h.font = font(italic=True, size=8, color=BRUN); h.alignment = WRAP
    pm.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    pm.row_dimensions[r].height = 30
    P[key] = f'Paramètres!$C${r}'
    r += 1
merged(pm, r, 2, 7, 'GÉNÉRAL', font(bold=True, color=WHITE), BRUN, 20); r += 1
param('Date d\'arrêté du portefeuille', 'date', '31/12/2025', '@', 'Date des encours saisis dans l\'onglet Portefeuille.')
param('Facteurs validés', 'valide', 'Non', '@', 'Passer à « Oui » quand toutes les valeurs de démonstration ont été remplacées par des facteurs sourcés.', listkey='ouinon')
merged(pm, r, 2, 7, 'ÉNERGIE', font(bold=True, color=WHITE), BRUN, 20); r += 1
param('Facteur d\'émission de l\'énergie consommée par les entreprises (tCO2e / MWh)', 'fe_energie', 0.6, '0.000', 'Mix moyen électricité et combustibles. Démonstration.', True)
param('Facteur d\'émission de l\'énergie des bâtiments (tCO2e / MWh)', 'fe_bat', 0.65, '0.000', 'Pour l\'immobilier : électricité du réseau marocain et part de gaz. Démonstration.', True)
merged(pm, r, 2, 7, 'DETTE SOUVERAINE (PCAF : émissions territoriales et PIB en parité de pouvoir d\'achat)', font(bold=True, color=WHITE), BRUN, 20); r += 1
param('Émissions territoriales du Maroc (tCO2e, hors UTCATF)', 'sov_em', 90000000, '#,##0', 'Source à retenir : dernier inventaire national soumis à la CCNUCC. Démonstration.', True)
param('PIB du Maroc en parité de pouvoir d\'achat (MAD)', 'sov_pib', 4000000000000, '#,##0', 'Source : Banque mondiale ou FMI, converti en MAD. Démonstration.', True)
r += 1
merged(pm, r, 2, 7, 'FACTEURS SECTORIELS (estimation quand l\'entreprise ne publie pas ses émissions)', font(bold=True, color=WHITE), BRUN, 20); r += 1
head(pm, r, 2, ['Secteur', 'Scopes 1 et 2 (tCO2e / M MAD de chiffre d\'affaires)', 'Scope 3 (tCO2e / M MAD de chiffre d\'affaires)', 'Scopes 1 et 2 (tCO2e / M MAD d\'actif total)', 'Source et année', 'Statut', ''][:6], BRUN, 44)
r += 1
S0 = r
for name, a, b, c in SECTEURS:
    pm.cell(r, 2, name).font = font(); pm.cell(r, 2).border = BOX
    for j, v in enumerate([a, b, c], start=3):
        x = pm.cell(r, j, v); inp(x, '#,##0.0'); x.font = font(bold=True, italic=True, color=RED, size=9)
    s = pm.cell(r, 6, 'Valeur de démonstration'); inp(s, '@'); s.font = font(italic=True, color=RED, size=9)
    st = pm.cell(r, 7, f'=IF(F{r}="Valeur de démonstration","⚠ à remplacer","✓ sourcé")'); st.font = font(size=9, bold=True, color=BRUN); st.border = BOX
    r += 1
S1 = r - 1
SECT = f'Paramètres!$B${S0}:$B${S1}'
FE_CA, FE_S3, FE_ACT = (f'Paramètres!${col}${S0}:${col}${S1}' for col in 'CDE')
r += 1
merged(pm, r, 2, 7, 'IMMOBILIER (intensité énergétique et surface par défaut, par type de bien)', font(bold=True, color=WHITE), BRUN, 20); r += 1
head(pm, r, 2, ['Type de bien', 'Intensité énergétique (kWh / m² / an)', 'Surface moyenne (m²) si inconnue'], BRUN, 44); r += 1
B0 = r
for t, k, m2 in [('Logement', 110, 90), ('Bureaux', 180, 800), ('Commerce', 250, 300), ('Hôtel', 280, 3000), ('Entrepôt / logistique', 80, 2000), ('Autre', 150, 500)]:
    pm.cell(r, 2, t).font = font(); pm.cell(r, 2).border = BOX
    for j, v in enumerate([k, m2], start=3):
        x = pm.cell(r, j, v); inp(x, '#,##0'); x.font = font(bold=True, italic=True, color=RED, size=9)
    r += 1
B1 = r - 1
BIEN, BIEN_KWH, BIEN_M2 = f'Paramètres!$B${B0}:$B${B1}', f'Paramètres!$C${B0}:$C${B1}', f'Paramètres!$D${B0}:$D${B1}'
r += 1
merged(pm, r, 2, 7, 'VÉHICULES (facteur d\'émission et kilométrage annuel par défaut)', font(bold=True, color=WHITE), BRUN, 20); r += 1
head(pm, r, 2, ['Type de véhicule', 'Facteur d\'émission (kgCO2e / km)', 'Kilométrage annuel par défaut (km)'], BRUN, 44); r += 1
V0 = r
for t, fe, km in [('Voiture essence', 0.19, 13000), ('Voiture diesel', 0.17, 17000), ('Voiture hybride', 0.12, 15000), ('Voiture électrique', 0.09, 15000),
                  ('Utilitaire léger', 0.25, 25000), ('Camion', 0.90, 60000), ('Deux-roues', 0.08, 8000), ('Inconnu', 0.19, 15000)]:
    pm.cell(r, 2, t).font = font(); pm.cell(r, 2).border = BOX
    for j, (v, f_) in enumerate([(fe, '0.000'), (km, '#,##0')], start=3):
        x = pm.cell(r, j, v); inp(x, f_); x.font = font(bold=True, italic=True, color=RED, size=9)
    r += 1
V1 = r - 1
VEH, VEH_FE, VEH_KM = f'Paramètres!$B${V0}:$B${V1}', f'Paramètres!$C${V0}:$C${V1}', f'Paramètres!$D${V0}:$D${V1}'
protect(pm)

# ------------------------------------------------------------------ Portefeuille
pf = wb.create_sheet('Portefeuille')
COLS = [
    # clé, en-tête, largeur, aide, liste, num, groupe
    ('n', 'N°', 5, 'Numéro de ligne.', None, None, 'id'),
    ('id', 'Identifiant contrepartie', 14, 'Code client ou numéro de dossier.', None, None, 'id'),
    ('nom', 'Raison sociale ou libellé', 26, 'Nom de la contrepartie ou du projet.', None, None, 'id'),
    ('grp', 'Groupe d\'affaires', 16, 'Facultatif : groupe auquel appartient la contrepartie.', None, None, 'id'),
    ('cls', 'Classe d\'actifs PCAF', 28, 'Choisir dans la liste : elle détermine la formule de calcul.', 'classe', None, 'id'),
    ('sec', 'Secteur d\'activité', 28, 'Obligatoire pour les entreprises et les projets.', 'secteur', None, 'id'),
    ('pays', 'Pays', 12, 'Pays de la contrepartie ou du projet.', 'pays', None, 'id'),
    ('enc', 'Encours au bilan (MAD)', 15, 'Montant utilisé au bilan à la date d\'arrêté (hors engagements non tirés).', None, 'dec', 'id'),
    ('cote', 'Société cotée ?', 9, 'Oui : l\'EVIC est utilisée.', 'ouinon', None, 'ent'),
    ('evic', 'EVIC (MAD)', 15, 'Sociétés cotées : capitalisation boursière + dettes + intérêts minoritaires + actions préférentielles (fin d\'exercice).', None, 'dec', 'ent'),
    ('fpd', 'Fonds propres + dettes (MAD)', 15, 'Sociétés non cotées et projets : total des fonds propres et des dettes financières (pour un projet : coût total financé).', None, 'dec', 'ent'),
    ('ca', 'Chiffre d\'affaires (MAD)', 15, 'Pour l\'estimation sectorielle (score 4).', None, 'dec', 'ent'),
    ('act', 'Total actif (MAD)', 15, 'Pour l\'estimation par les actifs (score 5) si le chiffre d\'affaires est inconnu.', None, 'dec', 'ent'),
    ('s1', 'Émissions scope 1 déclarées (tCO2e)', 13, 'Publiées par l\'entreprise, ou émissions du projet (EIES, outil IFC CEET).', None, 'dec', 'ent'),
    ('s2', 'Émissions scope 2 déclarées (tCO2e)', 13, 'Publiées par l\'entreprise ou le projet.', None, 'dec', 'ent'),
    ('s3', 'Émissions scope 3 déclarées (tCO2e)', 13, 'Facultatif.', None, 'dec', 'ent'),
    ('verif', 'Émissions vérifiées ?', 10, 'Oui si un tiers indépendant a vérifié les émissions (score 1).', 'ouinon', None, 'ent'),
    ('src', 'Source des émissions', 22, 'Rapport ESG, étude d\'impact, outil CEET…', 'source_em', None, 'ent'),
    ('mwh', 'Énergie consommée (MWh)', 13, 'Si les émissions ne sont pas publiées mais la consommation d\'énergie est connue (score 3).', None, 'dec', 'ent'),
    ('val', 'Valeur du bien à l\'octroi (MAD)', 15, 'Immobilier : valeur du bien au moment de l\'octroi du crédit.', None, 'dec', 'imm'),
    ('typb', 'Type de bien', 14, 'Immobilier : type de bien.', 'type_bien', None, 'imm'),
    ('m2', 'Surface (m²)', 10, 'Immobilier : surface du bien.', None, 'dec', 'imm'),
    ('kwh', 'Consommation d\'énergie réelle (kWh / an)', 13, 'Immobilier : si connue (factures, diagnostic).', None, 'dec', 'imm'),
    ('nb', 'Nombre de biens ou de véhicules', 11, 'Portefeuille agrégé (habitat, automobile) : nombre de dossiers regroupés sur la ligne. 1 si vide.', None, 'dec', 'agr'),
    ('valv', 'Valeur du véhicule à l\'octroi (MAD)', 15, 'Crédit automobile : prix du véhicule.', None, 'dec', 'veh'),
    ('typv', 'Type de véhicule', 14, 'Crédit automobile.', 'type_vehicule', None, 'veh'),
    ('km', 'Kilométrage annuel (km)', 12, 'Crédit automobile : si connu.', None, 'dec', 'veh'),
    ('com', 'Commentaire', 26, 'Hypothèses, source des données.', None, None, 'veh'),
    # colonnes calculées
    ('fa', 'Facteur d\'attribution', 10, 'Part des émissions de la contrepartie attribuée à la banque (plafonnée à 100 %).', None, None, 'calc'),
    ('emc', 'Émissions scopes 1+2 de la contrepartie (tCO2e)', 14, 'Émissions retenues, selon la meilleure donnée disponible.', None, None, 'calc'),
    ('meth', 'Méthode retenue', 42, 'Option PCAF utilisée.', None, None, 'calc'),
    ('dq', 'Score de qualité PCAF (1 à 5)', 9, '1 = meilleure qualité (émissions vérifiées) ; 5 = estimation la plus grossière.', None, None, 'calc'),
    ('ef', 'Émissions financées scopes 1+2 (tCO2e)', 14, 'Facteur d\'attribution × émissions de la contrepartie.', None, None, 'calc'),
    ('ef3', 'Émissions financées scope 3 (tCO2e)', 13, 'Entreprises et projets uniquement.', None, None, 'calc'),
    ('int', 'Intensité (tCO2e / M MAD d\'encours)', 12, 'Émissions financées scopes 1+2 par million de dirhams d\'encours.', None, None, 'calc'),
    ('ctl', 'Contrôle', 30, 'Ne pas remplir. Corriger la ligne tant que ce n\'est pas « OK ».', None, None, 'calc'),
]
K = {c[0]: L(i) for i, c in enumerate(COLS, start=1)}
NC = len(COLS)
for i, c in enumerate(COLS, start=1):
    pf.column_dimensions[L(i)].width = c[2]
banner(pf, 'Portefeuille : une ligne par exposition', 'Saisissez dans les cellules orangées (une ligne par crédit, titre ou projet). Les colonnes grises se calculent seules. Les lignes fournies sont des EXEMPLES FICTIFS : effacez-les (sélectionner les cellules orangées › Suppr) avant de saisir votre portefeuille.', NC)
GROUPS = [('id', 'IDENTIFICATION ET ENCOURS', BRUN), ('ent', 'ENTREPRISES, PROJETS (et données d\'émissions)', '7A4A2A'), ('imm', 'IMMOBILIER', '8A5A3B'), ('agr', 'AGRÉGAT', '6A3A1A'), ('veh', 'VÉHICULES', '9A6A4B'), ('calc', 'RÉSULTATS (calculés, ne pas modifier)', '595959')]
for g, lab, col in GROUPS:
    idxs = [i for i, c in enumerate(COLS, start=1) if c[6] == g]
    merged(pf, 4, min(idxs), max(idxs), lab, font(bold=True, color=WHITE, size=9), col, 20, WRAPC)
HDR = 5
head(pf, HDR, 1, [c[1] for c in COLS], ORANGE, 54)
for i, c in enumerate(COLS, start=1):
    if c[6] == 'calc':
        pf.cell(HDR, i).fill = fill('595959')
    pf.cell(HDR, i).comment = Comment(c[3], 'Aide')
for i, c in enumerate(COLS, start=1):
    x = pf.cell(HDR + 1, i, c[3]); x.font = font(italic=True, size=8, color=BRUN); x.fill = fill(TINT); x.alignment = WRAP; x.border = BOX
pf.row_dimensions[HDR + 1].height = 84
FIRST = HDR + 2
N = 500
LAST = FIRST + N - 1


def c(k, r): return f'{K[k]}{r}'


def formulas(r):
    cls, enc = c('cls', r), c('enc', r)
    fct = (f'IF({cls}="","",IF({enc}="","",MIN(1,'
           f'IF({cls}="{CLASSES[0]}",{enc}/{c("evic", r)},'
           f'IF({cls}="{CLASSES[1]}",IF(AND({c("cote", r)}="Oui",{c("evic", r)}>0),{enc}/{c("evic", r)},{enc}/{c("fpd", r)}),'
           f'IF({cls}="{CLASSES[2]}",{enc}/{c("fpd", r)},'
           f'IF(OR({cls}="{CLASSES[3]}",{cls}="{CLASSES[4]}"),{enc}/{c("val", r)},'
           f'IF({cls}="{CLASSES[5]}",{enc}/{c("valv", r)},'
           f'IF({cls}="{CLASSES[6]}",{enc}/{P["sov_pib"]},0)))))))))')
    fa = f'=IFERROR({fct},"⚠")'
    corp = f'OR({cls}="{CLASSES[0]}",{cls}="{CLASSES[1]}",{cls}="{CLASSES[2]}")'
    imm = f'OR({cls}="{CLASSES[3]}",{cls}="{CLASSES[4]}")'
    veh = f'{cls}="{CLASSES[5]}"'
    feca = f'INDEX({FE_CA},MATCH({c("sec", r)},{SECT},0))'
    feact = f'INDEX({FE_ACT},MATCH({c("sec", r)},{SECT},0))'
    kwhm2 = f'INDEX({BIEN_KWH},MATCH({c("typb", r)},{BIEN},0))'
    m2def = f'INDEX({BIEN_M2},MATCH({c("typb", r)},{BIEN},0))'
    vfe = f'INDEX({VEH_FE},MATCH(IF({c("typv", r)}="","Inconnu",{c("typv", r)}),{VEH},0))'
    vkm = f'INDEX({VEH_KM},MATCH(IF({c("typv", r)}="","Inconnu",{c("typv", r)}),{VEH},0))'
    rep = f'(N({c("s1", r)})+N({c("s2", r)}))'
    nbk = f'MAX(1,N({c("nb", r)}))'
    # émissions de la contrepartie
    emc = (f'=IFERROR(IF({cls}="","",'
           f'IF({corp},IF({rep}>0,{rep},IF(N({c("mwh", r)})>0,{c("mwh", r)}*{P["fe_energie"]},IF(N({c("ca", r)})>0,{c("ca", r)}/1000000*{feca},IF(N({c("act", r)})>0,{c("act", r)}/1000000*{feact},IF(N({c("fpd", r)})>0,{c("fpd", r)}/1000000*{feact},""))))),'
           f'IF({imm},IF(N({c("kwh", r)})>0,{c("kwh", r)}/1000*{P["fe_bat"]},IF(N({c("m2", r)})>0,{c("m2", r)}*{kwhm2}/1000*{P["fe_bat"]},{nbk}*{m2def}*{kwhm2}/1000*{P["fe_bat"]})),'
           f'IF({veh},{nbk}*IF(N({c("km", r)})>0,{c("km", r)},{vkm})*{vfe}/1000,'
           f'IF({cls}="{CLASSES[6]}",{P["sov_em"]},""))))),"⚠")')
    meth = (f'=IF({cls}="","",'
            f'IF({corp},IF({rep}>0,IF({c("verif", r)}="Oui","1 — émissions déclarées et vérifiées","2 — émissions déclarées non vérifiées"),IF(N({c("mwh", r)})>0,"3 — énergie consommée × facteur",IF(N({c("ca", r)})>0,"4 — chiffre d\'affaires × facteur sectoriel",IF(OR(N({c("act", r)})>0,N({c("fpd", r)})>0),"5 — actifs × facteur sectoriel","⚠ données insuffisantes")))),'
            f'IF({imm},IF(N({c("kwh", r)})>0,"2 — consommation d\'énergie réelle",IF(N({c("m2", r)})>0,"4 — surface × intensité par défaut","5 — surface moyenne × intensité par défaut")),'
            f'IF({veh},IF(N({c("km", r)})>0,"3 — kilométrage réel × facteur du véhicule",IF(OR({c("typv", r)}="",{c("typv", r)}="Inconnu"),"5 — véhicule moyen, kilométrage par défaut","4 — type de véhicule, kilométrage par défaut")),'
            f'IF({cls}="{CLASSES[6]}","2 — émissions territoriales et PIB (PPA)","")))))')
    dq = f'=IF(OR({c("meth", r)}="",LEFT({c("meth", r)},1)="⚠"),"",VALUE(LEFT({c("meth", r)},1)))'
    ef = f'=IFERROR(IF(OR({c("fa", r)}="",{c("emc", r)}=""),"",{c("fa", r)}*{c("emc", r)}),"")'
    fes3 = f'INDEX({FE_S3},MATCH({c("sec", r)},{SECT},0))'
    ef3 = f'=IFERROR(IF({corp},{c("fa", r)}*IF(N({c("s3", r)})>0,{c("s3", r)},IF(N({c("ca", r)})>0,{c("ca", r)}/1000000*{fes3},"")),""),"")'
    it = f'=IFERROR(IF({c("ef", r)}="","",{c("ef", r)}/({enc}/1000000)),"")'
    need = (f'IF(AND({corp},{c("sec", r)}=""),"⚠ Secteur manquant",'
            f'IF(AND({cls}="{CLASSES[0]}",N({c("evic", r)})=0),"⚠ EVIC manquante",'
            f'IF(AND(OR({cls}="{CLASSES[1]}",{cls}="{CLASSES[2]}"),N({c("fpd", r)})=0,N({c("evic", r)})=0),"⚠ Fonds propres + dettes manquants",'
            f'IF(AND({imm},N({c("val", r)})=0),"⚠ Valeur du bien manquante",'
            f'IF(AND({imm},{c("typb", r)}=""),"⚠ Type de bien manquant",'
            f'IF(AND({veh},N({c("valv", r)})=0),"⚠ Valeur du véhicule manquante",'
            f'IF(LEFT({c("meth", r)},1)="⚠","⚠ Données d\'émissions insuffisantes",'
            f'IF(OR({c("fa", r)}="⚠",{c("emc", r)}="⚠"),"⚠ Vérifier les montants",'
            f'IF(AND({c("fa", r)}<>"",{c("fa", r)}>=1),"⚠ Attribution plafonnée à 100 % : vérifier le dénominateur","OK")))))))))')
    ctl = (f'=IF(COUNTA({K["id"]}{r}:{K["com"]}{r})=0,"",IF({cls}="","⚠ Choisir la classe d\'actifs",IF(N({enc})=0,"⚠ Encours manquant",{need})))')
    return {'fa': fa, 'emc': emc, 'meth': meth, 'dq': dq, 'ef': ef, 'ef3': ef3, 'int': it, 'ctl': ctl}


for i in range(N):
    rr = FIRST + i
    pf.cell(rr, 1, i + 1).font = font(size=9, color='808080'); pf.cell(rr, 1).border = BOX
    for j, col in enumerate(COLS, start=1):
        if col[0] == 'n':
            continue
        x = pf.cell(rr, j)
        if col[6] == 'calc':
            x.value = formulas(rr)[col[0]]; x.fill = fill(GREY if col[0] != 'ctl' else WHITE); x.border = BOX
            x.font = font(size=9, bold=col[0] in ('ef', 'ctl'), color=RED if col[0] == 'ctl' else BRUN)
            x.number_format = {'fa': '0.00%', 'emc': '#,##0.0', 'ef': '#,##0.0', 'ef3': '#,##0.0', 'int': '#,##0.0'}.get(col[0], 'General')
        else:
            inp(x, '#,##0' if col[5] == 'dec' else None)
            x.alignment = Alignment(vertical='top', wrap_text=True)
    pf.row_dimensions[rr].height = 18
for j, col in enumerate(COLS, start=1):
    rng = f'{L(j)}{FIRST}:{L(j)}{LAST}'
    if col[4]: dv_list(pf, col[4], rng)
    if col[5] == 'dec': dv_num(pf, rng)
pf.freeze_panes = f'D{FIRST}'
pf.conditional_formatting.add(f'{K["ctl"]}{FIRST}:{K["ctl"]}{LAST}', FormulaRule(formula=[f'{K["ctl"]}{FIRST}="OK"'], font=Font(name='Arial', bold=True, color=GREEN)))

# exemples fictifs
EX = [
    dict(id='EX-001', nom='Cimenterie A (exemple fictif)', cls=CLASSES[1], sec='Ciment, chaux, matériaux de construction', pays='Maroc', enc=450e6, cote='Non', fpd=6.5e9, ca=3.2e9, s1=1150000, s2=85000, verif='Oui', src='Rapport ESG ou extra-financier'),
    dict(id='EX-002', nom='Groupe agroalimentaire B (exemple fictif)', cls=CLASSES[1], sec='Industrie agroalimentaire', pays='Maroc', enc=220e6, cote='Non', fpd=2.1e9, ca=1.8e9, s1=42000, s2=31000, verif='Non', src='Questionnaire client'),
    dict(id='EX-003', nom='Société de transport C (exemple fictif)', cls=CLASSES[1], sec='Transport routier et logistique', pays='Maroc', enc=95e6, cote='Non', fpd=600e6, ca=420e6, mwh=38000),
    dict(id='EX-004', nom='PME textile D (exemple fictif)', cls=CLASSES[1], sec='Textile, cuir, habillement', pays='Maroc', enc=18e6, cote='Non', fpd=75e6, ca=60e6),
    dict(id='EX-005', nom='Distributeur E (exemple fictif)', cls=CLASSES[1], sec='Commerce de gros et de détail', pays='Côte d\'Ivoire', enc=40e6, cote='Non', fpd=210e6, act=260e6),
    dict(id='EX-006', nom='Groupe industriel coté F (exemple fictif)', cls=CLASSES[0], sec='Chimie, parachimie, engrais', pays='Maroc', enc=150e6, cote='Oui', evic=95e9, ca=40e9, s1=2900000, s2=600000, s3=4100000, verif='Oui', src='Rapport ESG ou extra-financier'),
    dict(id='EX-007', nom='Parc éolien G (exemple fictif)', cls=CLASSES[2], sec='Énergies renouvelables', pays='Maroc', enc=600e6, fpd=2.4e9, s1=900, s2=300, verif='Non', src='Étude d\'impact (EIES) du projet'),
    dict(id='EX-008', nom='Centrale thermique H (exemple fictif)', cls=CLASSES[2], sec='Production d\'électricité d\'origine fossile', pays='Maroc', enc=800e6, fpd=9e9, s1=3100000, s2=15000, verif='Oui', src='Outil IFC CEET (projets)'),
    dict(id='EX-009', nom='Immeuble de bureaux I (exemple fictif)', cls=CLASSES[3], sec='Promotion et activités immobilières', pays='Maroc', enc=70e6, val=120e6, typb='Bureaux', m2=6500),
    dict(id='EX-010', nom='Hôtel J (exemple fictif)', cls=CLASSES[3], sec='Tourisme, hôtellerie, restauration', pays='Maroc', enc=130e6, val=210e6, typb='Hôtel', m2=14000, kwh=4100000),
    dict(id='EX-011', nom='Crédits habitat — portefeuille agrégé (exemple fictif)', cls=CLASSES[4], pays='Maroc', enc=2.5e9, val=4.6e9, typb='Logement', m2=520000, nb=5200, com='Agrégation de 5 200 dossiers ; surface totale'),
    dict(id='EX-012', nom='Crédits automobiles — portefeuille agrégé (exemple fictif)', cls=CLASSES[5], pays='Maroc', enc=380e6, valv=560e6, typv='Voiture diesel', nb=2800, com='Agrégation de 2 800 dossiers ; à affiner par type'),
    dict(id='EX-013', nom='Bons du Trésor marocain (exemple fictif)', cls=CLASSES[6], pays='Maroc', enc=12e9, com='Encours de bons du Trésor'),
]
for k_, d in enumerate(EX):
    rr = FIRST + k_
    for key, v in d.items():
        pf[f'{K[key]}{rr}'] = v
protect(pf)

# ------------------------------------------------------------------ Synthèse
sy = wb.create_sheet('Synthèse')
for col, w in zip('ABCDEFGH', [4, 44, 18, 12, 20, 18, 16, 14]): sy.column_dimensions[col].width = w
banner(sy, 'Synthèse des émissions financées (PCAF)', 'Résultats calculés à partir de l\'onglet Portefeuille. Tableau de publication conforme à la structure PCAF : émissions absolues, couverture, intensité et score de qualité, par classe d\'actifs et par secteur.', 8)
R = lambda k: f'Portefeuille!${K[k]}${FIRST}:${K[k]}${LAST}'
sy['B4'] = f'=IF({P["valide"]}="Oui","✓ Facteurs d\'émission validés","⚠ RÉSULTATS INDICATIFS : les facteurs d\'émission sont des valeurs de démonstration (onglet Paramètres)")'
sy['B4'].font = font(bold=True, size=12, color=RED)
sy.merge_cells('B4:H4')
sy['B5'] = '="Date d\'arrêté : "&' + P['date']; sy['B5'].font = font(italic=True, color=BRUN)
rows_kpi = [
    ('Encours total saisi (MAD)', f'=SUM({R("enc")})', '#,##0'),
    ('Encours couvert par un calcul (MAD)', f'=SUMIFS({R("enc")},{R("ef")},">=0")', '#,##0'),
    ('Taux de couverture', '=IFERROR(C9/C8,0)', '0.0%'),
    ('Émissions financées scopes 1+2 (tCO2e)', f'=SUM({R("ef")})', '#,##0'),
    ('Émissions financées scope 3 (tCO2e)', f'=SUM({R("ef3")})', '#,##0'),
    ('Intensité économique (tCO2e / M MAD d\'encours couvert)', '=IFERROR(C11/(C9/1000000),0)', '#,##0.0'),
    ('Score de qualité moyen pondéré par l\'encours (1 = meilleur)', f'=IFERROR(SUMPRODUCT({R("dq")},{R("enc")})/SUMIFS({R("enc")},{R("dq")},">0"),"")', '0.00'),
    ('WACI entreprises (tCO2e / M MAD de chiffre d\'affaires, pondéré par l\'encours)', f'=IFERROR(SUM(Portefeuille!$ZX${FIRST}:$ZX${LAST})/SUM(Portefeuille!$ZY${FIRST}:$ZY${LAST}),"")', '#,##0.0'),
    ('Lignes avec alerte (onglet Portefeuille)', f'=COUNTIF({R("ctl")},"⚠*")', '0'),
]
head(sy, 7, 2, ['Indicateur', 'Valeur'], ORANGE, 22)
for i, (lab, f_, fmt) in enumerate(rows_kpi, start=8):
    a = sy.cell(i, 2, lab); a.font = font(bold=True, color=BRUN); a.border = BOX; a.fill = fill(TINT if i % 2 else TINT2); a.alignment = WRAP
    b = sy.cell(i, 3, f_); b.font = font(bold=True, color=ORANGE, size=11); b.border = BOX; b.number_format = fmt; b.fill = fill(TINT if i % 2 else TINT2)
    sy.row_dimensions[i].height = 28
r0 = 8 + len(rows_kpi) + 1
merged(sy, r0, 2, 8, 'PAR CLASSE D\'ACTIFS', font(bold=True, color=WHITE), BRUN, 20)
head(sy, r0 + 1, 2, ['Classe d\'actifs PCAF', 'Encours (MAD)', 'Part', 'Émissions financées 1+2 (tCO2e)', 'Émissions financées 3 (tCO2e)', 'Intensité (tCO2e / M MAD)', 'Score de qualité'], ORANGE, 32)
def block(r_start, keys, keycol):
    for i, kv in enumerate(keys):
        rr = r_start + i
        sy.cell(rr, 2, kv)
        sy.cell(rr, 3, f'=SUMIFS({R("enc")},{R(keycol)},B{rr})')
        sy.cell(rr, 4, f'=IFERROR(C{rr}/$C$8,0)')
        sy.cell(rr, 5, f'=SUMIFS({R("ef")},{R(keycol)},B{rr})')
        sy.cell(rr, 6, f'=SUMIFS({R("ef3")},{R(keycol)},B{rr})')
        sy.cell(rr, 7, f'=IFERROR(E{rr}/(SUMIFS({R("enc")},{R(keycol)},B{rr},{R("ef")},">=0")/1000000),"")')
        sy.cell(rr, 8, f'=IFERROR(SUMPRODUCT(--({R(keycol)}=B{rr}),{R("dq")},{R("enc")})/SUMIFS({R("enc")},{R(keycol)},B{rr},{R("dq")},">0"),"")')
        for j, fmt in zip(range(2, 9), ['@', '#,##0', '0.0%', '#,##0', '#,##0', '#,##0.0', '0.0']):
            x = sy.cell(rr, j); x.border = BOX; x.font = font(size=9, bold=(j == 2), color=BRUN if j == 2 else GRIS); x.number_format = fmt; x.fill = fill(TINT if i % 2 else TINT2)
    rt = r_start + len(keys)
    sy.cell(rt, 2, 'Total')
    for j, col in zip(range(3, 7), 'CDEF'):
        sy.cell(rt, j, f'=SUM({col}{r_start}:{col}{rt - 1})')
    sy.cell(rt, 7, f'=IFERROR(E{rt}/(SUMPRODUCT(SUMIFS({R("enc")},{R(keycol)},B{r_start}:B{rt - 1},{R("ef")},">=0"))/1000000),"")')
    sy.cell(rt, 8, f'=IFERROR(SUMPRODUCT(H{r_start}:H{rt - 1},SUMIFS({R("enc")},{R(keycol)},B{r_start}:B{rt - 1},{R("dq")},">0"))/SUMPRODUCT(SUMIFS({R("enc")},{R(keycol)},B{r_start}:B{rt - 1},{R("dq")},">0")),"")')
    for j, fmt in zip(range(2, 9), ['@', '#,##0', '0.0%', '#,##0', '#,##0', '#,##0.0', '0.0']):
        x = sy.cell(rt, j); x.border = BOX; x.font = font(bold=True, color=BRUN); x.number_format = fmt; x.fill = fill(TINT)
    return rt
rt = block(r0 + 2, CLASSES, 'cls')
r1 = rt + 2
merged(sy, r1, 2, 8, 'PAR SECTEUR D\'ACTIVITÉ (entreprises, projets et immobilier commercial)', font(bold=True, color=WHITE), BRUN, 20)
head(sy, r1 + 1, 2, ['Secteur', 'Encours (MAD)', 'Part', 'Émissions financées 1+2 (tCO2e)', 'Émissions financées 3 (tCO2e)', 'Intensité (tCO2e / M MAD)', 'Score de qualité'], ORANGE, 32)
rt2 = block(r1 + 2, LISTS['secteur'], 'sec')
r2 = rt2 + 2
merged(sy, r2, 2, 8, 'LES 10 PREMIÈRES EXPOSITIONS PAR ÉMISSIONS FINANCÉES', font(bold=True, color=WHITE), BRUN, 20)
head(sy, r2 + 1, 2, ['Contrepartie', 'Encours (MAD)', 'Rang', 'Émissions financées 1+2 (tCO2e)', 'Part du total', 'Classe d\'actifs', 'Score'], ORANGE, 32)
# colonne d'aide pour le classement (départage des ex aequo)
HELP = 'ZZ'
for hc in (HELP, 'ZX', 'ZY'): pf.column_dimensions[hc].hidden = True
for i in range(N):
    rr = FIRST + i
    pf[f'{HELP}{rr}'] = f'=IF(ISNUMBER({c("ef", rr)}),{c("ef", rr)}+ROW()/1000000000,"")'
    pf[f'ZX{rr}'] = f'=IF(AND(ISNUMBER({c("emc", rr)}),N({c("ca", rr)})>0,N({c("enc", rr)})>0),{c("enc", rr)}*{c("emc", rr)}/({c("ca", rr)}/1000000),"")'
    pf[f'ZY{rr}'] = f'=IF(ISNUMBER(ZX{rr}),{c("enc", rr)},"")'
HR = f'Portefeuille!${HELP}${FIRST}:${HELP}${LAST}'
for k_ in range(10):
    rr = r2 + 2 + k_
    mt = f'MATCH(LARGE({HR},{k_ + 1}),{HR},0)'
    vals = [f'=IFERROR(INDEX({R("nom")},{mt}),"")', f'=IFERROR(INDEX({R("enc")},{mt}),"")', k_ + 1, f'=IFERROR(INDEX({R("ef")},{mt}),"")',
            f'=IFERROR(E{rr}/$C$11,"")', f'=IFERROR(INDEX({R("cls")},{mt}),"")', f'=IFERROR(INDEX({R("dq")},{mt}),"")']
    for j, (v, fmt) in enumerate(zip(vals, ['@', '#,##0', '0', '#,##0', '0.0%', '@', '0']), start=2):
        x = sy.cell(rr, j, v); x.border = BOX; x.font = font(size=9); x.number_format = fmt; x.fill = fill(TINT if k_ % 2 else TINT2); x.alignment = WRAP
sy.freeze_panes = 'A7'
protect(sy)

# ------------------------------------------------------------------ Lisez-moi
rd = wb.create_sheet('Lisez-moi', 0)
for col, w in zip('ABCD', [4, 34, 62, 52]): rd.column_dimensions[col].width = w
banner(rd, f'Outil émissions financées — version {VERSION}', 'Calcul des émissions de gaz à effet de serre financées par la banque (scope 3, catégorie 15 du GHG Protocol) selon le standard PCAF, que l\'IFC promeut auprès des institutions financières, avec l\'approche IFC pour les financements de projets. Outil distinct du canevas de l\'empreinte propre.', 4)
r = 4
def section(t):
    global r
    merged(rd, r, 1, 4, t, font(bold=True, size=13, color=BRUN)); r += 1
def table(headers, rows):
    global r
    head(rd, r, 2, headers, ORANGE, 22); r += 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row, start=2):
            x = rd.cell(r, j, v); x.font = font(size=9); x.border = BOX; x.alignment = WRAP; x.fill = fill(TINT if i % 2 else TINT2)
        rd.row_dimensions[r].height = 14 * max(1, -(-max(len(str(v)) for v in row) // 62)) + 4
        r += 1
    r += 1
section('1. En 5 étapes')
table(['Étape', 'Ce que vous faites', 'Où'], [
    ['1. Paramétrer', 'Remplacer les valeurs de démonstration (en rouge) par des facteurs sourcés ; passer « Facteurs validés » à Oui.', 'Paramètres'],
    ['2. Extraire le portefeuille', 'Extraire du système d\'information les encours au bilan à la date d\'arrêté, avec la classe d\'actifs et le secteur.', 'Système crédit, salle des marchés'],
    ['3. Saisir ou coller', 'Une ligne par exposition (crédit, titre, projet), ou par portefeuille agrégé pour l\'habitat et l\'automobile. Effacer d\'abord les exemples fictifs.', 'Portefeuille (cellules orangées)'],
    ['4. Compléter les données d\'émissions', 'Émissions publiées par les clients, émissions des projets (EIES, outil IFC CEET), surfaces et types de biens, types de véhicules.', 'Portefeuille'],
    ['5. Contrôler et publier', 'Corriger les lignes dont la colonne Contrôle n\'affiche pas « OK », puis lire la synthèse.', 'Portefeuille, Synthèse'],
])
section('2. Les classes d\'actifs et les formules (standard PCAF)')
table(['Classe d\'actifs', 'Facteur d\'attribution', 'Émissions de la contrepartie'], [
    [CLASSES[0], 'Encours ÷ EVIC', 'Émissions déclarées, sinon énergie, chiffre d\'affaires ou actifs × facteur sectoriel'],
    [CLASSES[1], 'Encours ÷ (fonds propres + dettes) ; EVIC si la société est cotée', 'Idem'],
    [CLASSES[2], 'Encours ÷ (fonds propres + dettes du projet)', 'Émissions du projet (étude d\'impact, outil IFC CEET), sinon facteur sectoriel'],
    [CLASSES[3], 'Encours ÷ valeur du bien à l\'octroi', 'Consommation d\'énergie réelle, sinon surface × intensité par défaut'],
    [CLASSES[4], 'Encours ÷ valeur du bien à l\'octroi', 'Idem (portefeuille agrégé possible)'],
    [CLASSES[5], 'Encours ÷ valeur du véhicule à l\'octroi', 'Kilométrage × facteur du véhicule (par défaut selon le type) × nombre de véhicules si la ligne agrège plusieurs dossiers'],
    [CLASSES[6], 'Encours ÷ PIB en parité de pouvoir d\'achat', 'Émissions territoriales du pays'],
])
section('3. Le score de qualité des données (1 = meilleur, 5 = plus faible)')
table(['Score', 'Signification', 'Exemple'], [
    ['1', 'Émissions déclarées et vérifiées par un tiers', 'Rapport ESG audité d\'un groupe coté'],
    ['2', 'Émissions déclarées non vérifiées ; consommation d\'énergie réelle d\'un bâtiment', 'Questionnaire client ; factures d\'énergie d\'un hôtel'],
    ['3', 'Données physiques : énergie consommée ou kilométrage réel', 'MWh consommés par un transporteur'],
    ['4', 'Estimation par le chiffre d\'affaires et le secteur ; surface connue', 'PME sans bilan carbone'],
    ['5', 'Estimation par les actifs ou des valeurs moyennes', 'Distributeur dont seul le total actif est connu'],
])
section('4. Code couleur')
for txt, fl in [('Cellule orangée : à remplir', INPUT), ('Colonne grise : calculée, ne pas modifier', GREY), ('Texte rouge en italique : valeur de démonstration à remplacer', TINT2)]:
    x = rd.cell(r, 2, txt); x.fill = fill(fl); x.font = font(italic='rouge' in txt, color=RED if 'rouge' in txt else GRIS); x.border = BOX; r += 1
r += 1
merged(rd, r, 1, 4, 'Avertissement : les facteurs d\'émission fournis sont des valeurs de démonstration destinées à montrer le fonctionnement de l\'outil. Ils ne doivent pas être utilisés pour une publication. Les exemples du portefeuille sont fictifs.', font(bold=True, color=RED), ROSE, 36)
protect(rd)

# ------------------------------------------------------------------ Contrôle
ct = wb.create_sheet('Contrôle')
for col, w in zip('ABC', [4, 60, 16]): ct.column_dimensions[col].width = w
banner(ct, 'Contrôle du portefeuille', 'Nombre de lignes par type d\'alerte. Une alerte se corrige dans l\'onglet Portefeuille (colonne Contrôle).', 3)
head(ct, 4, 2, ['Alerte', 'Lignes'], ORANGE, 22)
alerts = ['OK', '⚠ Choisir la classe d\'actifs', '⚠ Encours manquant', '⚠ Secteur manquant', '⚠ EVIC manquante', '⚠ Fonds propres + dettes manquants', '⚠ Valeur du bien manquante',
          '⚠ Type de bien manquant', '⚠ Valeur du véhicule manquante', '⚠ Données d\'émissions insuffisantes', '⚠ Vérifier les montants', '⚠ Attribution plafonnée à 100 % : vérifier le dénominateur']
for i, a in enumerate(alerts, start=5):
    x = ct.cell(i, 2, a); x.border = BOX; x.font = font(bold=(i == 5), color=GREEN if i == 5 else RED)
    y = ct.cell(i, 3, f'=COUNTIF({R("ctl")},B{i})'); y.border = BOX; y.font = font(bold=True)
protect(ct)

wb.move_sheet('Listes', offset=len(wb.sheetnames) - 1 - wb.sheetnames.index('Listes'))
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'Lisez-moi': BRUN, 'Paramètres': BRUN, 'Synthèse': GREEN, 'Contrôle': RED, 'Listes': '8A7A6A'}.get(ws.title, ORANGE)
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
protect(ls)
wb.active = 0
wb.save(OUT)
print('ok', wb.sheetnames)
