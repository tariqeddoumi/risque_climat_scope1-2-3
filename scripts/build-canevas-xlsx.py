import sys, os
from openpyxl.drawing.image import Image as XLImage
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

OUT = sys.argv[1]
F = 'Arial'
# Charte Groupe BCP (identique aux classeurs de l'outil de notation)
NAVY, GREEN, LBLUE, GREY, YEL, ORANGE, WHITE = 'E37B05', '462016', 'F2E8DB', 'EDEDED', 'FDEBD0', 'FBE3E1', 'FFFFFF'
GRIS, MUTED, RED = '3D3D3B', '8A7A6A', 'B3261E'
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'groupe-bcp-logo.png')
def add_logo(ws, anchor):
    img = XLImage(LOGO); img.width, img.height = 92, 48; ws.add_image(img, anchor)
thin = Side(style='thin', color='D9C7B0')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def fill(c): return PatternFill('solid', start_color=c, end_color=c)
def font(**k): return Font(name=F, **{'size': 10, 'color': GRIS, **k})
WRAP = Alignment(wrap_text=True, vertical='top')
WRAPC = Alignment(wrap_text=True, vertical='center')

wb = Workbook()

# ---------------- Listes ----------------
LISTS = {
 'statut': ['Valeur', 'Néant (aucune activité)', 'Non disponible', 'Non applicable'],
 'methode': ['Mesuré (facture, compteur, pesée)', 'Extrait d\'un système (ERP, inventaire)', 'Estimé à partir d\'un montant', 'Extrapolé (ratio, échantillon)', 'Estimation d\'expert'],
 'mois': ['Année entière', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
 'type_site': ['Siège central', 'Siège régional (BPR)', 'Data center', 'Succursale', 'Agence', 'Centre d\'affaires', 'Centre d\'estivage', 'Centre de formation', 'Archives / entrepôt', 'Autre'],
 'entite': ['BCP - Siège', 'BCP - Réseau Casablanca', 'BP Centre Sud', 'BP Fès-Meknès', 'BP Laâyoune', 'BP Marrakech-Béni Mellal', 'BP Nador-Al Hoceima', 'BP Oujda', 'BP Rabat-Kénitra', 'BP Tanger-Tétouan', 'Autre (préciser)'],
 'occupation': ['Propriétaire', 'Locataire', 'Mis à disposition'],
 'usage_elec': ['Site entier', 'Partie du site', 'Salle informatique (sous-compteur)', 'Équipements informatiques (sous-compteur IT)', 'Partagé avec un tiers', 'Recharge véhicules électriques', 'Production photovoltaïque', 'Chaleur ou froid acheté à un réseau'],
 'contrat_elec': ['Réseau standard (ONEE ou distributeur)', 'Électricité renouvelable sous contrat (loi 13-09)', 'Autoproduction (loi 82-21)', 'Garantie d\'origine ou certificat', 'Sans objet'],
 'usage_comb': ['Véhicules de service', 'Groupe électrogène', 'Cuisson / chauffage / eau chaude', 'Autre (préciser)'],
 'combustible': ['Gasoil', 'Essence', 'Butane', 'Propane', 'Bois', 'Électricité (véhicule électrique)', 'Autre (préciser)'],
 'unite_comb': ['Litre', 'kg', 'Bouteille 12 kg', 'Bouteille 3 kg', 'm³', 'kWh'],
 'type_clim': ['Split / climatiseur mural', 'Armoire de climatisation', 'Groupe froid / centrale (chiller)', 'Climatisation salle informatique', 'Climatiseur de véhicule', 'Réfrigérateur / chambre froide'],
 'fluide': ['R22', 'R410A', 'R407C', 'R32', 'R134a', 'R404A', 'R290 (propane)', 'Inconnu'],
 'famille': ['Papier vierge', 'Papier recyclé', 'Consommables d\'impression (toners, cartouches)', 'Fournitures de bureau', 'Produits d\'hygiène et d\'entretien', 'Restauration (repas, café)', 'Maintenance des équipements et bâtiments', 'Gardiennage', 'Nettoyage', 'Télécoms et abonnements internet', 'Impressions externalisées et publicité', 'Location de véhicules', 'Loyers et charges locatives', 'Travaux et matériaux de construction', 'Logiciels et services informatiques', 'Conseil et autres services', 'Autres achats (préciser)'],
 'materiau': ['Acier', 'Aluminium', 'Béton / ciment', 'Verre', 'Bois', 'Plâtre', 'Cuivre', 'Plastique / PVC', 'Autre (préciser)'],
 'unite_achat': ['Rame de 500 feuilles', 'Tonne', 'kg', 'Repas', 'Tasse / dose', 'Unité', 'Mois de prestation', 'Sans objet'],
 'categorie_immo': ['PC fixe (unité centrale)', 'Écran', 'PC portable', 'Imprimante individuelle', 'Multifonction / photocopieur', 'Scanner', 'Scanner de chèques', 'Serveur', 'Switch / routeur / firewall', 'Onduleur', 'Téléphone fixe', 'Smartphone / téléphone portable', 'Tablette', 'Vidéoprojecteur / écran d\'affichage', 'GAB / guichet automatique', 'Bureau (mobilier)', 'Chaise / fauteuil', 'Armoire / rangement', 'Véhicule de service', 'Bâtiment', 'Autre équipement (préciser)'],
 'flux': ['Ordures ménagères (mélange)', 'Papier et carton', 'Archives détruites', 'Plastique', 'Verre', 'Métaux', 'Déchets électroniques (DEEE)', 'Cartouches et toners', 'Piles et lampes', 'Déchets alimentaires', 'Autres (préciser)'],
 'unite_dechet': ['kg', 'Tonne', 'Sac 100 L', 'Bac 240 L', 'Bac 660 L', 'Benne'],
 'periode_dechet': ['Année entière', 'Un mois', 'Une semaine de pesée'],
 'traitement': ['Décharge / enfouissement', 'Recyclage', 'Incinération', 'Compostage', 'Réemploi / don', 'Inconnu'],
 'mode_dep': ['Avion', 'Train', 'Voiture personnelle (km remboursés)', 'Voiture de service', 'Voiture de location', 'Taxi', 'Autocar', 'Hôtel (nuitées)'],
 'trajet': ['Aller simple', 'Aller-retour', 'Sans objet'],
 'classe': ['Économique', 'Affaires', 'Première', 'Sans objet'],
 'ville': ['Agadir (AGA)', 'Al Hoceima (AHU)', 'Casablanca (CMN)', 'Dakhla (VIL)', 'Errachidia (ERH)', 'Essaouira (ESU)', 'Fès (FEZ)', 'Guelmim (GLN)', 'Laâyoune (EUN)', 'Marrakech (RAK)', 'Nador (NDR)', 'Ouarzazate (OZZ)', 'Oujda (OUD)', 'Rabat (RBA)', 'Tanger (TNG)', 'Tan-Tan (TTA)', 'Tétouan (TTU)', 'Kénitra (gare)', 'El Jadida (gare)', 'Meknès (gare)',
          'Abidjan (ABJ)', 'Abou Dhabi (AUH)', 'Amsterdam (AMS)', 'Antananarivo (TNR)', 'Bamako (BKO)', 'Barcelone (BCN)', 'Brazzaville (BZV)', 'Bruxelles (BRU)', 'Charm el-Cheikh (SSH)', 'Conakry (CKY)', 'Dakar (DSS)', 'Douala (DLA)', 'Doha (DOH)', 'Dubaï (DXB)', 'Francfort (FRA)', 'Genève (GVA)', 'Istanbul (IST)', 'Le Caire (CAI)', 'Libreville (LBV)', 'Lisbonne (LIS)', 'Londres (LHR)', 'Madrid (MAD)', 'Maurice (MRU)', 'Milan (MXP)', 'Montpellier (MPL)', 'Montréal (YUL)', 'New York (JFK)', 'Niamey (NIM)', 'Nouakchott (NKC)', 'Paris (CDG / ORY)', 'Séoul (ICN)', 'Tunis (TUN)', 'Vienne (VIE)', 'Washington (IAD)', 'Autre (préciser en commentaire)'],
 'zone': ['Maroc', 'International'],
 'carburant_vp': ['Diesel', 'Essence', 'Hybride', 'Électrique', 'Sans objet'],
 'categorie_dt': ['Employés', 'Visiteurs et clients'],
 'mode_dt': ['Voiture personnelle diesel', 'Voiture personnelle essence', 'Voiture hybride', 'Voiture électrique', 'Covoiturage', 'Moto / scooter', 'Bus', 'Tramway', 'Train', 'Grand taxi', 'Petit taxi', 'Transport du personnel (navette)', 'Moto > 50 cm³', 'Moto < 50 cm³', 'Avion', 'À pied / vélo'],
 'type_transport': ['Transport de fonds', 'Livraisons fournisseurs (fret amont)', 'Courrier envoyé aux clients', 'Navette courrier interne', 'Colis / messagerie', 'Autre (préciser)'],
 'type_vehicule': ['Fourgon blindé', 'Utilitaire léger', 'Camion porteur', 'Voiture', 'Deux-roues', 'Avion (fret express)', 'Inconnu'],
}
ls = wb.active; ls.title = 'Listes'
RANGES = {}
for j, (k, vals) in enumerate(LISTS.items(), start=1):
    c = ls.cell(row=1, column=j, value=k); c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY)
    for i, v in enumerate(vals, start=2):
        ls.cell(row=i, column=j, value=v).font = font()
    ls.column_dimensions[L(j)].width = 34
    RANGES[k] = f"Listes!${L(j)}$2:${L(j)}${len(vals)+1}"
ls.freeze_panes = 'A2'

def dv_list(ws, key, rng):
    dv = DataValidation(type='list', formula1='=' + RANGES[key], allow_blank=True, showErrorMessage=True,
                        errorTitle='Valeur non autorisée', error='Choisissez une valeur dans la liste déroulante.')
    ws.add_data_validation(dv); dv.add(rng)

def dv_num(ws, rng, kind='decimal', lo=0, hi=None, msg='Saisissez un nombre positif, sans unité ni espace (ex. 2850 ou 12.5).'):
    if hi is None:
        dv = DataValidation(type=kind, operator='greaterThanOrEqual', formula1=str(lo), allow_blank=True, showErrorMessage=True, errorTitle='Nombre attendu', error=msg)
    else:
        dv = DataValidation(type=kind, operator='between', formula1=str(lo), formula2=str(hi), allow_blank=True, showErrorMessage=True, errorTitle='Nombre attendu', error=msg)
    ws.add_data_validation(dv); dv.add(rng)

def merged(ws, r, c1, c2, text, f=None, fl=None, h=None, al=WRAP):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(row=r, column=c1, value=text)
    cell.font = f or font(); cell.alignment = al
    if fl:
        for c in range(c1, c2 + 1): ws.cell(row=r, column=c).fill = fill(fl)
    if h: ws.row_dimensions[r].height = h
    return cell

COMMON_TAIL = [
    ('Statut', 22, 'Obligatoire. « Valeur » si vous donnez un chiffre ; sinon Néant, Non disponible ou Non applicable.', 'statut', None),
    ('Méthode', 34, 'Comment avez-vous obtenu le chiffre ?', 'methode', None),
    ('Pièce justificative', 24, 'Nom du fichier joint ou référence (ex. Facture Lydec 01/2025).', None, None),
    ('Commentaire', 34, 'Hypothèses, explication d\'un écart, nom de la direction qui détient l\'info…', None, None),
]
N_INPUT = 100
CONTROL_SHEETS = []
REG = {}  # onglet -> colonnes et plage de saisie, pour l'export IFC et les contrôles
VERSION = '2.0'
UNLOCK = Protection(locked=False)
def protect(ws):
    pr = ws.protection
    pr.sheet = True
    pr.insertRows = pr.deleteRows = pr.formatColumns = pr.formatRows = pr.formatCells = pr.sort = pr.autoFilter = False

def poste_sheet(name, title, scope, owner, where, steps, cols, cases, errors, qty_key, amt_key=None, unit_key=None):
    """cols: list of (header, width, help, listkey|None, numkind|None). numkind: 'dec' or 'year' or 'int'."""
    ws = wb.create_sheet(name)
    cols = [('N°', 5, 'Numéro de ligne.', None, None), ('Code site (si plusieurs sites)', 13, 'Seulement si le fichier couvre plusieurs sites : code du site concerné par la ligne.', None, None)] + cols + COMMON_TAIL
    cols.append(('Contrôle automatique', 30, 'Ne pas remplir : se calcule tout seul. Corrigez la ligne tant que ce n\'est pas « OK ».', None, None))
    nc = len(cols)
    for j, c in enumerate(cols, start=1):
        ws.column_dimensions[L(j)].width = c[1]
    idx = {c[0]: j for j, c in enumerate(cols, start=1)}
    ctrl_col = nc
    def ctrl_formula(r):
        s = f'{L(idx["Statut"])}{r}'
        q = f'{L(idx[qty_key])}{r}'
        alts = [amt_key] if isinstance(amt_key, str) else (amt_key or [])
        cond_val = 'AND(' + ','.join([f'{q}=""'] + [f'{L(idx[a])}{r}=""' for a in alts]) + ')'
        msg_val = '"⚠ Quantité ou montant manquant"' if alts else '"⚠ Quantité manquante"'
        unit = f'IF(AND({q}<>"",{L(idx[unit_key])}{r}=""),"⚠ Choisir l\'unité",' if unit_key else ''
        close = ')' if unit_key else ''
        return (f'=IF(COUNTA({L(2)}{r}:{L(nc-1)}{r})=0,"",IF({s}="","⚠ Choisir un statut",'
                f'IF(AND({s}="Valeur",{cond_val}),{msg_val},'
                f'IF(AND({s}<>"Valeur",{q}<>""),"⚠ Statut incohérent avec le chiffre saisi",'
                f'IF(AND({q}<>"",NOT(ISNUMBER({q}))),"⚠ Nombre attendu, sans unité",'
                f'{unit}"OK"{close})))))')

    r = 1
    merged(ws, r, 1, nc, title, font(bold=True, size=16, color=WHITE), NAVY, 40, WRAPC); r += 1
    add_logo(ws, f'{L(nc + 1)}1'); ws.sheet_view.showGridLines = False
    merged(ws, r, 1, nc, f'{scope}     |     Qui remplit : {owner}     |     Où trouver l\'information : {where}', font(size=10, color=NAVY), LBLUE, 32); r += 2
    merged(ws, r, 1, nc, 'COMMENT REMPLIR', font(bold=True, color=GREEN, size=11)); r += 1
    for s in steps:
        merged(ws, r, 1, nc, s, font(), None, 16 if len(s) < 170 else 30); r += 1
    r += 1
    merged(ws, r, 1, nc, 'CAS POSSIBLES ET EXEMPLES  (lignes grises : exemples à lire, ne pas les modifier ni les recopier)', font(bold=True, color=GREEN, size=11)); r += 1
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=j, value=c[0]); cell.font = font(bold=True, color=WHITE, size=9); cell.fill = fill(GREEN); cell.alignment = WRAPC; cell.border = BOX
    ws.row_dimensions[r].height = 30; r += 1
    for cas_title, cas_text, rows in cases:
        merged(ws, r, 1, nc, f'{cas_title} : {cas_text}', font(bold=True, color=NAVY, size=10), LBLUE, 30 if len(cas_text) > 120 else 18); r += 1
        for ex in rows:
            ws.cell(row=r, column=1, value='Ex.')
            for k, v in ex.items():
                ws.cell(row=r, column=idx[k], value=v)
            ws.cell(row=r, column=ctrl_col, value=ctrl_formula(r))
            for j in range(1, nc + 1):
                cell = ws.cell(row=r, column=j); cell.fill = fill(GREY); cell.font = font(italic=True, color='404040', size=9); cell.border = BOX; cell.alignment = WRAP
            r += 1
    r += 1
    merged(ws, r, 1, nc, 'ERREURS À ÉVITER', font(bold=True, color=RED, size=11)); r += 1
    for e in errors:
        merged(ws, r, 1, nc, '✗  ' + e, font(color=RED), ORANGE, 16 if len(e) < 170 else 30); r += 1
    r += 1
    merged(ws, r, 1, nc, 'VOTRE SAISIE  (remplir uniquement les cellules orangées, une ligne par cas)', font(bold=True, color=WHITE, size=12), GREEN, 22, WRAPC); r += 1
    hdr = r
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=j, value=c[0]); cell.font = font(bold=True, color=WHITE, size=9); cell.fill = fill(NAVY); cell.alignment = WRAPC; cell.border = BOX
        cell.comment = Comment(c[2], 'Aide'); cell.comment.width = 260; cell.comment.height = 90
    ws.row_dimensions[r].height = 32; r += 1
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=r, column=j, value=c[2]); cell.font = font(italic=True, size=8, color=GREEN); cell.fill = fill(LBLUE); cell.alignment = WRAP; cell.border = BOX
    ws.row_dimensions[r].height = 64; r += 1
    first = r; last = r + N_INPUT - 1
    for i in range(N_INPUT):
        rr = first + i
        ws.cell(row=rr, column=1, value=i + 1).font = font(size=9, color='808080')
        for j in range(2, nc):
            cell = ws.cell(row=rr, column=j); cell.fill = fill(YEL); cell.border = BOX; cell.font = font(size=9, bold=True, color='1F3FBF'); cell.alignment = Alignment(vertical='top', wrap_text=True); cell.protection = UNLOCK
        cc = ws.cell(row=rr, column=ctrl_col, value=ctrl_formula(rr)); cc.font = font(bold=True, size=9, color=RED); cc.border = BOX
        ws.cell(row=rr, column=1).border = BOX
    for j, c in enumerate(cols, start=1):
        rng = f'{L(j)}{first}:{L(j)}{last}'
        if c[3]: dv_list(ws, c[3], rng)
        if c[4] == 'dec':
            dv_num(ws, rng); [setattr(ws.cell(row=x, column=j), 'number_format', '#,##0.00') for x in range(first, last + 1)]
        elif c[4] == 'int':
            dv_num(ws, rng, 'whole', 0, msg='Saisissez un nombre entier positif.'); [setattr(ws.cell(row=x, column=j), 'number_format', '#,##0') for x in range(first, last + 1)]
        elif c[4] == 'year':
            dv_num(ws, rng, 'whole', 1950, 2035, msg='Saisissez une année sur 4 chiffres (ex. 2021).')
        elif c[4] == 'days':
            dv_num(ws, rng, 'decimal', 0, 7, msg='Nombre de jours par semaine, entre 0 et 7.')
    ws.freeze_panes = None
    ws.column_dimensions[L(nc + 1)].width = 15
    for rr in range(first, last + 1): ws.row_dimensions[rr].height = 18
    ws.sheet_view.zoomScale = 90
    ws.print_title_rows = f'{hdr}:{hdr}'
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    CONTROL_SHEETS.append((name, title, f"'{name}'!${L(ctrl_col)}${first}:${L(ctrl_col)}${last}"))
    REG[name] = {'idx': idx, 'first': first, 'last': last}
    protect(ws)
    return ws

# ---------------- Lisez-moi ----------------
rd = wb.create_sheet('Lisez-moi', 0)
for col, w in zip('ABCD', [4, 34, 62, 52]): rd.column_dimensions[col].width = w
r = 1
merged(rd, r, 1, 4, f'Canevas de collecte : empreinte carbone du Groupe BCP — version {VERSION}', font(bold=True, size=18, color=WHITE), NAVY, 40, WRAPC); r += 1
add_logo(rd, 'E1'); rd.column_dimensions['E'].width = 15; rd.sheet_view.showGridLines = False
merged(rd, r, 1, 4, 'Ce fichier sert à recueillir les données nécessaires au calcul des émissions de gaz à effet de serre de votre site ou de votre direction. Vous n\'avez aucun calcul à faire : donnez simplement les chiffres demandés, dans l\'unité demandée.', font(size=11), LBLUE, 36); r += 2

def section(t):
    global r
    merged(rd, r, 1, 4, t, font(bold=True, size=13, color=GREEN)); r += 1
def table(headers, rows, hfill=NAVY):
    global r
    for j, h in enumerate(headers, start=2):
        c = rd.cell(row=r, column=j, value=h); c.font = font(bold=True, color=WHITE); c.fill = fill(hfill); c.border = BOX; c.alignment = WRAPC
    r += 1
    for row in rows:
        for j, v in enumerate(row, start=2):
            c = rd.cell(row=r, column=j, value=v); c.font = font(); c.border = BOX; c.alignment = WRAP
        longest = max(len(str(v)) for v in row)
        rd.row_dimensions[r].height = 15 * max(1, -(-longest // 60))
        r += 1
    r += 1

section('1. En 5 étapes')
table(['Étape', 'Ce que vous faites', 'Exemple'], [
    ['1. Identifiez-vous', 'Remplissez l\'onglet « Identification » : le site ou la direction, l\'année des données, votre fonction.', 'Agence Al Boughaz, exercice 2025, chef d\'agence'],
    ['2. Ouvrez l\'onglet de chaque poste', 'Les onglets 1 à 9 correspondent chacun à une source d\'émission. Lisez le cadre « Comment remplir » et les cas grisés.', 'Onglet « 1-Électricité »'],
    ['3. Saisissez dans les cellules orangées', 'Une ligne par compteur, par véhicule, par famille d\'achat… Utilisez les listes déroulantes.', 'Compteur 123456, Janvier, 2 850 kWh'],
    ['4. Donnez un statut à chaque ligne', 'Même si vous n\'avez pas l\'information : le statut dit pourquoi.', '« Non disponible » + « géré par la Direction du Patrimoine »'],
    ['5. Vérifiez puis envoyez', 'La colonne « Contrôle automatique » doit afficher « OK » partout. L\'onglet « Contrôle » fait le bilan.', '0 alerte → fichier prêt à envoyer avec les pièces'],
])
section('2. Les 6 règles d\'or')
table(['Règle', 'Pourquoi', 'Bon / mauvais exemple'], [
    ['Un chiffre seul dans la cellule, sans unité', 'L\'unité est choisie dans la colonne « Unité ». Un texte ne peut pas être additionné.', '✓ 35555   ✗ « 35555 kWh »   ✗ « 35 555 kw »'],
    ['La quantité avant le montant', 'Les émissions dépendent des litres, kWh ou kg, pas des dirhams. Le montant n\'est utile que si la quantité est inconnue.', '✓ 2 140 litres de gasoil   ✗ seulement 29 960 Dhs'],
    ['Montants toujours hors taxes (HT)', 'Tout le monde doit utiliser la même base. Si vous n\'avez que le TTC, divisez par 1,2 (TVA de 20 %) et indiquez-le en commentaire.', '✓ 10 000 HT   ✗ 12 000 TTC sans le dire'],
    ['Ne laissez jamais une ligne sans statut', 'Une case vide ne dit pas si l\'activité n\'existe pas ou si l\'information manque.', '✓ « Néant » : pas de groupe électrogène   ✗ case vide ou 0'],
    ['Ne comptez qu\'une fois', 'Si une donnée est déjà fournie par une direction centrale (ex. gardiennage payé par le siège), ne la ressaisissez pas pour votre site.', '✓ « Non applicable : payé et déclaré par le siège »'],
    ['Pas de calcul ni d\'extrapolation dans les cellules', 'Les hypothèses doivent rester visibles. Donnez la donnée brute et expliquez en commentaire.', '✓ 1 agence = 955 Dhs   ✗ « =955*113 »'],
])
section('3. Quel statut choisir ?')
table(['Statut', 'Quand l\'utiliser', 'Exemple'], [
    ['Valeur', 'Vous donnez un chiffre (quantité ou montant). Le chiffre 0 est admis s\'il est vrai et justifié.', 'Recharge de fluide = 0 kg : le rapport du mainteneur n\'indique aucune recharge'],
    ['Néant (aucune activité)', 'L\'activité n\'existe pas sur votre site.', 'Pas de groupe électrogène ; pas de panneau solaire'],
    ['Non disponible', 'L\'activité existe mais vous n\'avez pas le chiffre. Indiquez en commentaire qui pourrait l\'avoir.', 'Électricité incluse dans les charges du bailleur'],
    ['Non applicable', 'La donnée est fournie ailleurs, ou la ligne ne concerne pas votre type de site.', 'Achats de toners déclarés par la Direction des Achats pour tout le réseau'],
])
section('4. Quelle méthode choisir ?')
table(['Méthode', 'Signification', 'Exemple'], [
    ['Mesuré (facture, compteur, pesée)', 'Le chiffre vient directement d\'un document.', 'kWh lus sur la facture ONEE / Lydec / Redal / Amendis'],
    ['Extrait d\'un système (ERP, inventaire)', 'Le chiffre vient d\'une extraction informatique.', 'Extraction des achats HT par famille dans l\'ERP'],
    ['Estimé à partir d\'un montant', 'Vous connaissez le montant et le prix unitaire.', '29 960 Dhs ÷ 14 Dhs/L = 2 140 L de gasoil'],
    ['Extrapolé (ratio, échantillon)', 'Vous avez appliqué un ratio ou extrapolé un échantillon.', 'Déchets pesés pendant 1 semaine'],
    ['Estimation d\'expert', 'Évaluation raisonnée faute de document.', 'Année d\'achat estimée des chaises'],
])
section('5. Les onglets')
table(['Onglet', 'Contenu', 'Qui remplit en principe'], [
    ['Identification', 'Le site ou la direction, l\'année des données, le contributeur', 'Tous'],
    ['1-Électricité', 'kWh consommés par compteur, production solaire', 'Maintenance Multitechnique / gestionnaire de site'],
    ['2-Carburants et gaz', 'Gasoil, essence, butane, propane, bois', 'Gestion du Patrimoine / gestionnaire de site'],
    ['3-Climatisation', 'Climatiseurs et fluides frigorigènes rechargés', 'Maintenance Multitechnique'],
    ['4-Achats', 'Dépenses HT par famille d\'achat', 'Achats, Marchés & Budgets ; Développement social ; directions concernées'],
    ['5-Équipements', 'Matériel informatique, GAB, mobilier, véhicules, bâtiments, avec l\'année d\'achat', 'Support SI, Data Center, Télécom, PCA, Patrimoine'],
    ['6-Déchets', 'Quantités de déchets par type et leur traitement', 'Gestion du Patrimoine / gestionnaire de site'],
    ['7-Déplacements pro', 'Vols, trains, voitures, hôtels pour le travail', 'Pilotage RH / gestionnaire de site'],
    ['8-Domicile et visiteurs', 'Comment les employés viennent au travail ; venue des clients et visiteurs', 'RH / gestionnaire de site'],
    ['9-Transport fonds-courrier', 'Transport de fonds, courrier, navettes', 'Achats, Sécurité, Patrimoine'],
    ['Contrôle', 'Bilan automatique des lignes saisies, des alertes et des contrôles de cohérence', 'Ne pas remplir'],
    ['Export IFC', 'Vos données recomposées automatiquement au format du canevas IFC', 'Ne pas remplir (sauf le nombre de jours ouvrés)'],
    ['Listes', 'Valeurs des listes déroulantes', 'Ne pas modifier'],
])
section('6. Code couleur')
for txt, fl in [('Cellule orangée : à remplir', YEL), ('Ligne grise : exemple à lire, ne pas modifier', GREY), ('Bandeau beige : explication ou cas', LBLUE), ('Bandeau rose : erreur à éviter', ORANGE)]:
    c = rd.cell(row=r, column=2, value=txt); c.fill = fill(fl); c.font = font(); c.border = BOX; r += 1
r += 1
section(f'7. Nouveautés de la version {VERSION}')
table(['Nouveauté', 'Où', 'Pourquoi'], [
    ['Colonne « Code site » sur chaque ligne', 'Onglets 1 à 9', 'Une direction centrale peut répondre pour plusieurs sites dans un seul fichier'],
    ['Visiteurs et clients', 'Onglet 8', 'Poste demandé par la méthodologie IFC'],
    ['Loyers, papier recyclé, matériaux de travaux', 'Onglet 4', 'Rubriques du canevas IFC'],
    ['Véhicules électriques et kilométrage de la flotte', 'Onglet 2', 'Méthode IFC au kilomètre'],
    ['Type de contrat d\'électricité, puissance solaire (kWc), sous-compteur informatique, réseau de chaleur', 'Onglet 1', 'Scope 2 selon les contrats, efficacité des data centers, rubriques IFC'],
    ['Fluide récupéré, modèle du véhicule', 'Onglet 3', 'Bilan massique du GHG Protocol ; estimation IFC de la charge'],
    ['Zone Maroc / International, carburant du véhicule', 'Onglet 7', 'Distinction IFC des frais et des vols'],
    ['Fret amont, type de véhicule, consommation', 'Onglet 9', 'Rubriques IFC du transport amont'],
    ['Surface louée à des tiers, consommation d\'eau', 'Identification', 'Pratiques de reporting ESG des banques'],
    ['Contrôles de cohérence, onglet Export IFC', 'Contrôle, Export IFC', 'Qualité des données et continuité avec la mesure 2022'],
    ['100 lignes par onglet ; feuilles protégées (cellules orangées seules modifiables)', 'Tous', 'Éviter d\'effacer une formule ou une liste par erreur'],
])
merged(rd, r, 1, 4, 'Une question ? Écrivez à l\'équipe Empreinte carbone du Groupe BCP (Direction des Risques) en indiquant l\'onglet et le numéro de ligne.', font(italic=True, color=NAVY), LBLUE, 22)

# ---------------- Identification ----------------
idw = wb.create_sheet('Identification', 1)
for col, w in zip('ABC', [36, 40, 70]): idw.column_dimensions[col].width = w
merged(idw, 1, 1, 3, 'Identification du site ou de la direction', font(bold=True, size=16, color=WHITE), NAVY, 40, WRAPC)
add_logo(idw, 'D1'); idw.column_dimensions['D'].width = 15; idw.sheet_view.showGridLines = False
merged(idw, 2, 1, 3, f'Canevas version {VERSION}. Remplissez une fois, en haut du fichier. Un fichier = un site (ou une direction centrale qui répond pour plusieurs sites : indiquez-le et remplissez la colonne « Code site » de chaque ligne).', font(), LBLUE, 30)
for j, h in enumerate(['Information', 'Votre réponse', 'Aide et exemple'], start=1):
    c = idw.cell(row=4, column=j, value=h); c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY); c.border = BOX
ident = [
    ('Exercice (année des données)', 'Année sur 4 chiffres. Ex. 2025 = consommations du 1er janvier au 31 décembre 2025.', 'year'),
    ('Entité / BPR', 'Choisir dans la liste. Ex. BP Tanger-Tétouan.', 'entite'),
    ('Type de site', 'Choisir dans la liste. Ex. Agence.', 'type_site'),
    ('Code du site', 'Code interne de l\'agence ou du site (code entité). Ex. 4896.', None),
    ('Nom du site', 'Ex. Agence Al Boughaz.', None),
    ('Ville', 'Ex. Tanger.', None),
    ('Le fichier couvre', 'Ex. « Ce site uniquement » ou « Toutes les agences de la BPR » (dans ce cas, une ligne par agence dans chaque onglet, avec le code site en commentaire).', None),
    ('Surface du site (m²)', 'Surface de plancher occupée. Sert à comparer les sites entre eux.', 'dec'),
    ('Effectif présent (ETP)', 'Nombre d\'équivalents temps plein travaillant sur le site au 31/12.', 'dec'),
    ('Occupation des locaux', 'Propriétaire, locataire ou mis à disposition.', 'occupation'),
    ('Surface louée à des tiers (m²)', 'Facultatif : partie du bâtiment louée à une autre entreprise (0 si aucune).', 'dec'),
    ('Consommation d\'eau de l\'exercice (m³)', 'Facultatif : total des factures d\'eau (indicateur ESG, hors calcul carbone).', 'dec'),
    ('Fonction du contributeur', 'Ex. Chef d\'agence, Responsable logistique BPR. (pas d\'adresse électronique)', None),
    ('Date d\'envoi', 'Format JJ/MM/AAAA.', None),
]
for i, (lab, helpt, kind) in enumerate(ident, start=5):
    a = idw.cell(row=i, column=1, value=lab); a.font = font(bold=True); a.border = BOX
    b = idw.cell(row=i, column=2); b.fill = fill(YEL); b.border = BOX; b.font = font(); b.protection = UNLOCK
    c = idw.cell(row=i, column=3, value=helpt); c.font = font(italic=True, size=9, color=GREEN); c.alignment = WRAP; c.border = BOX
    idw.row_dimensions[i].height = 30
    if kind == 'year': dv_num(idw, f'B{i}', 'whole', 2015, 2035, 'Année sur 4 chiffres (ex. 2025).')
    elif kind == 'dec': dv_num(idw, f'B{i}')
    elif kind: dv_list(idw, kind, f'B{i}')

# ---------------- Onglets postes ----------------
poste_sheet('1-Électricité', '1. Électricité (scope 2)', 'Scope 2 : électricité achetée',
    'Maintenance Multitechnique ou gestionnaire de site', 'factures ONEE, Lydec, Redal, Amendis ou régie ; portail client du distributeur ; supervision des panneaux solaires',
    ['① Une ligne par compteur et par mois (12 lignes par compteur). Si vous n\'avez que le total de l\'année, une seule ligne avec « Année entière ».',
     '② Recopiez les kWh consommés figurant sur la facture (ligne « consommation » ou « énergie active »), pas la puissance souscrite en kW.',
     '③ Si vous n\'avez que le montant, laissez la quantité vide, indiquez le montant HT et la méthode « Estimé à partir d\'un montant ».',
     '④ Indiquez l\'usage du compteur : c\'est ce qui évite de compter deux fois la salle informatique ou les locaux d\'un tiers.',
     '⑤ Indiquez le type de contrat (réseau standard, électricité renouvelable sous contrat, autoproduction) : il permet de publier le scope 2 selon les contrats.'],
    [('Numéro de compteur ou de police', 22, 'Tel qu\'il figure sur la facture.', None, None),
     ('Usage du compteur', 26, 'Ce que le compteur alimente.', 'usage_elec', None),
     ('Mois', 14, 'Mois de la facture, ou « Année entière ».', 'mois', None),
     ('Quantité (kWh)', 14, 'kWh consommés (ou produits pour le solaire). Nombre seul.', None, 'dec'),
     ('Montant HT (MAD)', 14, 'Seulement si la quantité est inconnue, ou en complément.', None, 'dec'),
     ('Puissance installée (kWc)', 11, 'Photovoltaïque uniquement : puissance des panneaux.', None, 'dec'),
     ('Type de contrat', 36, 'Contrat d\'achat de l\'électricité (sert au scope 2 « market-based »).', 'contrat_elec', None)],
    [('Cas 1', 'J\'ai les factures mensuelles → une ligne par mois.', [
        {'Numéro de compteur ou de police': '123456', 'Usage du compteur': 'Site entier', 'Mois': 'Janvier', 'Quantité (kWh)': 2850, 'Montant HT (MAD)': 3420, 'Type de contrat': 'Réseau standard (ONEE ou distributeur)', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Facture Lydec 01-2025.pdf'},
        {'Numéro de compteur ou de police': '123456', 'Usage du compteur': 'Site entier', 'Mois': 'Février', 'Quantité (kWh)': 2610, 'Montant HT (MAD)': 3132, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Facture Lydec 02-2025.pdf', 'Commentaire': '… et ainsi de suite jusqu\'à décembre'}]),
     ('Cas 2', 'Je n\'ai que le total annuel en kWh → une ligne « Année entière ».', [
        {'Numéro de compteur ou de police': '123456', 'Usage du compteur': 'Site entier', 'Mois': 'Année entière', 'Quantité (kWh)': 35555, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Récapitulatif annuel distributeur'}]),
     ('Cas 3', 'Je n\'ai que le montant payé → quantité vide, montant HT, méthode « Estimé à partir d\'un montant ».', [
        {'Numéro de compteur ou de police': '123456', 'Usage du compteur': 'Site entier', 'Mois': 'Année entière', 'Montant HT (MAD)': 40000, 'Statut': 'Valeur', 'Méthode': 'Estimé à partir d\'un montant', 'Commentaire': '48 000 Dh TTC ÷ 1,2'}]),
     ('Cas 4', 'Un sous-compteur alimente la salle informatique (ou un locataire) → une ligne pour le compteur principal, une ligne pour le sous-compteur.', [
        {'Numéro de compteur ou de police': '778899', 'Usage du compteur': 'Site entier', 'Mois': 'Année entière', 'Quantité (kWh)': 5613674, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Inclut la salle informatique ligne suivante'},
        {'Numéro de compteur ou de police': 'SC-01', 'Usage du compteur': 'Salle informatique (sous-compteur)', 'Mois': 'Année entière', 'Quantité (kWh)': 505321, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Déjà compris dans le compteur 778899'}]),
     ('Cas 5', 'L\'électricité est payée par le bailleur ou un tiers → statut « Non disponible » et explication.', [
        {'Usage du compteur': 'Site entier', 'Mois': 'Année entière', 'Statut': 'Non disponible', 'Commentaire': 'Incluse dans les charges locatives ; surface 180 m²'}]),
     ('Cas 6', 'Le site a des panneaux solaires → kWh produits (pas le nombre de panneaux). S\'il n\'y en a pas : statut « Néant ».', [
        {'Numéro de compteur ou de police': 'Onduleur PV', 'Usage du compteur': 'Production photovoltaïque', 'Mois': 'Année entière', 'Quantité (kWh)': 92000, 'Puissance installée (kWc)': 150, 'Type de contrat': 'Autoproduction (loi 82-21)', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': '619 panneaux'}]),
     ('Cas 7', 'Data center : déclarez aussi le sous-compteur des équipements informatiques (il sert à calculer l\'efficacité énergétique, PUE).', [
        {'Numéro de compteur ou de police': 'DC-IT', 'Usage du compteur': 'Équipements informatiques (sous-compteur IT)', 'Mois': 'Année entière', 'Quantité (kWh)': 1850000, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Déjà compris dans le compteur principal du data center'}]),
     ('Cas 8', 'Pas de réseau de chaleur ou de froid (cas général) → une ligne « Néant ». La climatisation électrique est déjà dans la facture.', [
        {'Usage du compteur': 'Chaleur ou froid acheté à un réseau', 'Mois': 'Année entière', 'Statut': 'Néant (aucune activité)'}])],
    ['Écrire l\'unité dans la cellule (« 35555 kw ») : la cellule devient du texte et n\'est pas comptée.',
     'Confondre kW (puissance souscrite) et kWh (énergie consommée).',
     'Déclarer « chaleur » ou « froid » à part : la climatisation électrique est déjà dans la facture d\'électricité.',
     'Additionner plusieurs sites sur une même ligne sans le signaler en commentaire.'],
    'Quantité (kWh)', 'Montant HT (MAD)')

poste_sheet('2-Carburants et gaz', '2. Carburants et gaz (scope 1)', 'Scope 1 : combustion directe',
    'Gestion du Patrimoine ou gestionnaire de site', 'relevés des cartes carburant, bons d\'achat, bons de livraison de gaz, suivi des groupes électrogènes',
    ['① Une ligne par usage et par combustible (ex. gasoil des véhicules, gasoil du groupe électrogène, butane de la cuisine).',
     '② Donnez de préférence la quantité (litres, kg ou nombre de bouteilles) et choisissez l\'unité.',
     '③ Si vous n\'avez que le montant : montant HT + prix unitaire moyen de l\'année (ex. 14 Dh/L). Le calcul des litres sera fait par l\'équipe.',
     '④ Le carburant des véhicules loués avec carburant inclus va dans l\'onglet 4-Achats (location de véhicules), pas ici.'],
    [('Usage', 24, 'À quoi sert le combustible.', 'usage_comb', None),
     ('Combustible', 14, 'Type de combustible.', 'combustible', None),
     ('Quantité', 12, 'Nombre seul.', None, 'dec'),
     ('Unité', 15, 'Litre pour gasoil/essence ; bouteille ou kg pour le gaz.', 'unite_comb', None),
     ('Montant HT (MAD)', 14, 'Si la quantité est inconnue.', None, 'dec'),
     ('Prix unitaire HT (MAD par unité)', 16, 'Obligatoire si vous ne donnez que le montant.', None, 'dec'),
     ('Kilométrage annuel (km)', 13, 'Véhicules : km parcourus dans l\'année. Obligatoire pour un véhicule électrique.', None, 'dec'),
     ('Consommation aux 100 km', 13, 'L/100 km (thermique) ou kWh/100 km (électrique).', None, 'dec')],
    [('Cas 1', 'J\'ai les litres (relevé de cartes carburant).', [
        {'Usage': 'Véhicules de service', 'Combustible': 'Gasoil', 'Quantité': 2140, 'Unité': 'Litre', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Relevé cartes Afriquia 2025'}]),
     ('Cas 2', 'Je n\'ai que le montant dépensé → montant HT et prix moyen du litre.', [
        {'Usage': 'Véhicules de service', 'Combustible': 'Essence', 'Montant HT (MAD)': 15000, 'Prix unitaire HT (MAD par unité)': 13.5, 'Statut': 'Valeur', 'Méthode': 'Estimé à partir d\'un montant', 'Commentaire': 'Prix moyen 2025 relevé sur les factures'}]),
     ('Cas 3', 'Gaz butane en bouteilles → nombre de bouteilles et format.', [
        {'Usage': 'Cuisson / chauffage / eau chaude', 'Combustible': 'Butane', 'Quantité': 10, 'Unité': 'Bouteille 12 kg', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Bons de livraison'}]),
     ('Cas 4', 'Groupe électrogène utilisé pendant les coupures.', [
        {'Usage': 'Groupe électrogène', 'Combustible': 'Gasoil', 'Quantité': 1622, 'Unité': 'Litre', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Pleins du réservoir 2025'}]),
     ('Cas 5', 'Le site n\'a ni véhicule, ni groupe électrogène, ni gaz → une ligne par usage avec « Néant ».', [
        {'Usage': 'Groupe électrogène', 'Statut': 'Néant (aucune activité)'}]),
     ('Cas 6', 'Les véhicules sont gérés par une direction centrale → « Non applicable » et qui la déclare.', [
        {'Usage': 'Véhicules de service', 'Combustible': 'Gasoil', 'Statut': 'Non applicable', 'Commentaire': 'Déclaré par la Direction du Patrimoine pour toute la BPR'}]),
     ('Cas 7', 'Véhicule électrique → combustible « Électricité », kilométrage et consommation aux 100 km (kWh si connus).', [
        {'Usage': 'Véhicules de service', 'Combustible': 'Électricité (véhicule électrique)', 'Kilométrage annuel (km)': 18000, 'Consommation aux 100 km': 17, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Commentaire': 'Compteur kilométrique ; consommation constructeur'}]),
     ('Cas 8', 'Flotte : litres connus et kilométrage en complément.', [
        {'Usage': 'Véhicules de service', 'Combustible': 'Gasoil', 'Quantité': 5200, 'Unité': 'Litre', 'Kilométrage annuel (km)': 78000, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Commentaire': '4 véhicules'}])],
    ['Donner seulement « 29 960 Dhs » sans prix unitaire : impossible d\'en déduire les litres.',
     'Utiliser des bouteilles de 13 kg : au Maroc, les formats sont 12 kg et 3 kg.',
     'Inclure péages, entretien ou lavage dans le montant carburant.'],
    'Quantité', ['Montant HT (MAD)', 'Kilométrage annuel (km)'], 'Unité')

poste_sheet('3-Climatisation', '3. Climatisation et fluides frigorigènes (scope 1)', 'Scope 1 : fuites de gaz frigorigènes',
    'Maintenance Multitechnique', 'rapports d\'intervention du mainteneur, registre des fluides, plaque signalétique de l\'appareil',
    ['① Une ligne par appareil, ou par groupe d\'appareils identiques (même type, même fluide, même année) en indiquant leur nombre.',
     '② Ce qui compte pour les émissions, c\'est la quantité de gaz RECHARGÉE dans l\'année (elle compense ce qui a fui). Elle figure sur le rapport du mainteneur.',
     '③ La charge nominale (sur la plaque de l\'appareil) est demandée en complément ; ce n\'est PAS une fuite.',
     '④ Pas de recharge dans l\'année ? Saisissez 0 avec le statut « Valeur » : c\'est une vraie information.'],
    [('Emplacement / identifiant', 22, 'Ex. Bureau 2e étage, salle serveurs, véhicule 12345-A-6.', None, None),
     ('Type d\'appareil', 24, 'Type d\'équipement.', 'type_clim', None),
     ('Nombre d\'appareils identiques', 12, 'Nombre entier. 1 par défaut.', None, 'int'),
     ('Fluide', 12, 'Inscrit sur la plaque (R410A, R22…). « Inconnu » si illisible.', 'fluide', None),
     ('Charge nominale par appareil (kg)', 14, 'Sur la plaque signalétique.', None, 'dec'),
     ('Quantité rechargée dans l\'année (kg)', 16, 'Total rechargé pour la ligne. Sur le rapport d\'intervention.', None, 'dec'),
     ('Année de mise en service', 12, 'Année sur 4 chiffres.', None, 'year'),
     ('Fluide récupéré (kg)', 12, 'Facultatif : fluide récupéré lors d\'un démantèlement ou d\'une vidange (bilan massique).', None, 'dec'),
     ('Marque et modèle du véhicule', 18, 'Climatiseur de véhicule : sert à estimer la charge si elle est inconnue.', None, None)],
    [('Cas 1', 'Le mainteneur a rechargé l\'appareil.', [
        {'Emplacement / identifiant': 'Salle serveurs', 'Type d\'appareil': 'Climatisation salle informatique', 'Nombre d\'appareils identiques': 2, 'Fluide': 'R410A', 'Charge nominale par appareil (kg)': 12, 'Quantité rechargée dans l\'année (kg)': 3.5, 'Année de mise en service': 2018, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Rapport intervention 14/06/2025'}]),
     ('Cas 2', 'Aucune recharge dans l\'année → 0 kg, statut « Valeur ».', [
        {'Emplacement / identifiant': 'Bureaux 1er étage', 'Type d\'appareil': 'Split / climatiseur mural', 'Nombre d\'appareils identiques': 25, 'Fluide': 'R410A', 'Charge nominale par appareil (kg)': 1.2, 'Quantité rechargée dans l\'année (kg)': 0, 'Année de mise en service': 2020, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Contrat de maintenance : aucune recharge'}]),
     ('Cas 3', 'Le fluide n\'est pas lisible → « Inconnu » et photo de la plaque en pièce jointe.', [
        {'Emplacement / identifiant': 'Hall agence', 'Type d\'appareil': 'Armoire de climatisation', 'Nombre d\'appareils identiques': 1, 'Fluide': 'Inconnu', 'Quantité rechargée dans l\'année (kg)': 2, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Photo plaque.jpg'}]),
     ('Cas 4', 'Climatisation d\'un véhicule rechargée au garage.', [
        {'Emplacement / identifiant': 'Véhicule 12345-A-6', 'Marque et modèle du véhicule': 'Dacia Logan 2019', 'Type d\'appareil': 'Climatiseur de véhicule', 'Nombre d\'appareils identiques': 1, 'Fluide': 'R134a', 'Charge nominale par appareil (kg)': 0.5, 'Quantité rechargée dans l\'année (kg)': 0.4, 'Année de mise en service': 2019, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Facture garage'}]),
     ('Cas 5', 'Vous ne connaissez pas les recharges (info chez le mainteneur) → « Non disponible ».', [
        {'Emplacement / identifiant': 'Tout le site', 'Type d\'appareil': 'Split / climatiseur mural', 'Nombre d\'appareils identiques': 8, 'Fluide': 'R410A', 'Statut': 'Non disponible', 'Commentaire': 'Rapports chez la Maintenance Multitechnique'}]),
     ('Cas 6', 'Appareil démantelé dans l\'année : indiquez le fluide récupéré par le prestataire (il n\'est pas émis).', [
        {'Emplacement / identifiant': 'Ancienne armoire salle réunion', 'Type d\'appareil': 'Armoire de climatisation', 'Nombre d\'appareils identiques': 1, 'Fluide': 'R22', 'Charge nominale par appareil (kg)': 6, 'Quantité rechargée dans l\'année (kg)': 0, 'Année de mise en service': 2006, 'Fluide récupéré (kg)': 5.2, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Bordereau de récupération des fluides'}])],
    ['Reporter la charge totale (ex. 46,4 kg de R22) comme si tout avait fui dans l\'année.',
     'Écrire « 0 kg/an » ou « 4 ans » : un nombre seul dans chaque cellule.',
     'Oublier le statut lorsque la recharge est nulle.'],
    'Quantité rechargée dans l\'année (kg)')

poste_sheet('4-Achats', '4. Achats de biens et services (scope 3)', 'Scope 3 : biens et services achetés',
    'Achats, Marchés & Budgets ; Développement social (restauration) ; directions concernées', 'extraction de l\'ERP achats ou de la comptabilité par compte de charge, bons de commande',
    ['① Une ligne par famille d\'achat (liste déroulante), avec le montant TOTAL HT de l\'année. Inutile de détailler article par article (stylos, gommes…).',
     '② Pour le papier, séparez papier vierge et papier recyclé, et donnez aussi la quantité en rames (ou en tonnes). Pour les travaux, une ligne par matériau.',
     '③ Pour la restauration, donnez le nombre de repas ou de cafés si vous l\'avez.',
     '④ Si l\'achat est payé et déclaré par le siège pour tout le réseau, ne le ressaisissez pas pour votre site (« Non applicable »).'],
    [('Famille d\'achat', 30, 'Choisir dans la liste.', 'famille', None),
     ('Précision (facultatif)', 22, 'Ex. toners HP, contrat de gardiennage G4S.', None, None),
     ('Montant HT (MAD)', 14, 'Total de l\'année, hors taxes.', None, 'dec'),
     ('Quantité (facultatif)', 12, 'Ex. nombre de rames, de repas.', None, 'dec'),
     ('Unité', 18, 'Unité de la quantité.', 'unite_achat', None),
     ('Matériau (travaux)', 16, 'Famille « Travaux et matériaux » : nature du matériau ; quantité en kg ou tonnes.', 'materiau', None)],
    [('Cas 1', 'J\'ai une extraction de l\'ERP → une ligne par famille, total HT.', [
        {'Famille d\'achat': 'Consommables d\'impression (toners, cartouches)', 'Montant HT (MAD)': 185400, 'Unité': 'Sans objet', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Extraction ERP 2025.xlsx'}]),
     ('Cas 2', 'Papier → montant HT et nombre de rames.', [
        {'Famille d\'achat': 'Papier vierge', 'Précision (facultatif)': 'A4 80 g', 'Montant HT (MAD)': 352711, 'Quantité (facultatif)': 9534, 'Unité': 'Rame de 500 feuilles', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)'}]),
     ('Cas 3', 'Je n\'ai que des montants TTC → je divise par 1,2 et je le signale.', [
        {'Famille d\'achat': 'Gardiennage', 'Montant HT (MAD)': 3896649, 'Unité': 'Sans objet', 'Statut': 'Valeur', 'Méthode': 'Estimé à partir d\'un montant', 'Commentaire': '4 675 979 TTC ÷ 1,2'}]),
     ('Cas 4', 'Restauration → nombre de repas et montant.', [
        {'Famille d\'achat': 'Restauration (repas, café)', 'Précision (facultatif)': 'Repas cantine', 'Montant HT (MAD)': 2522360, 'Quantité (facultatif)': 126118, 'Unité': 'Repas', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)'}]),
     ('Cas 5', 'Achat payé par le siège pour tout le réseau → pas de ressaisie.', [
        {'Famille d\'achat': 'Gardiennage', 'Statut': 'Non applicable', 'Commentaire': 'Contrat central déclaré par la Direction des Achats'}]),
     ('Cas 6', 'Loyers des agences louées → famille « Loyers et charges locatives », montant HT annuel.', [
        {'Famille d\'achat': 'Loyers et charges locatives', 'Précision (facultatif)': 'Loyer agence + charges', 'Montant HT (MAD)': 216000, 'Unité': 'Sans objet', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Contrat de bail'}]),
     ('Cas 7', 'Travaux : une ligne par matériau acheté, quantité en kg ou tonnes.', [
        {'Famille d\'achat': 'Travaux et matériaux de construction', 'Précision (facultatif)': 'Réaménagement agence', 'Montant HT (MAD)': 85000, 'Quantité (facultatif)': 2.4, 'Unité': 'Tonne', 'Matériau (travaux)': 'Acier', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Bon de livraison'}]),
     ('Cas 8', 'Papier recyclé : famille distincte (facteur d\'émission plus faible).', [
        {'Famille d\'achat': 'Papier recyclé', 'Montant HT (MAD)': 18400, 'Quantité (facultatif)': 480, 'Unité': 'Rame de 500 feuilles', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)'}])],
    ['Détailler 200 articles ligne par ligne au lieu de totaliser par famille.',
     'Mélanger HT, TTC et « montant comptabilisé » selon les lignes.',
     'Saisir une estimation pour tout le réseau en multipliant un site par le nombre d\'agences (ex. « =955*113 »).'],
    'Montant HT (MAD)', None, None)

poste_sheet('5-Équipements', '5. Équipements et bâtiments (scope 3)', 'Scope 3 : immobilisations (fabrication des équipements)',
    'Support et Performance SI, Data Center, Télécom, PCA (GAB), Gestion du Patrimoine', 'inventaire informatique, registre des immobilisations, parc GAB, patrimoine immobilier',
    ['① Regroupez les équipements par catégorie ET par année d\'achat : une ligne = « 35 PC portables achetés en 2022 ».',
     '② Donnez l\'ANNÉE d\'achat ou d\'installation (ex. 2021), pas une tranche (« plus de 4 ans »). Si vous ne connaissez que l\'âge : année = exercice − âge (ex. 2025 − 3 = 2022).',
     '③ Ne comptez que le matériel en service au 31/12 sur le site. Le matériel sorti ou en stock central est déclaré par le SI.',
     '④ Bâtiment : une ligne par bâtiment, quantité 1, surface en m² et année de construction (ou de dernière grosse rénovation).'],
    [('Catégorie', 26, 'Choisir dans la liste.', 'categorie_immo', None),
     ('Modèle / référence (facultatif)', 22, 'Ex. CINEO C2560, HP ProBook.', None, None),
     ('Quantité', 10, 'Nombre d\'équipements (entier).', None, 'int'),
     ('Année d\'achat ou d\'installation', 14, 'Année sur 4 chiffres.', None, 'year'),
     ('Surface (m²) – bâtiments', 12, 'Uniquement pour les bâtiments.', None, 'dec')],
    [('Cas 1', 'J\'ai l\'inventaire avec les dates → je regroupe par catégorie et par année.', [
        {'Catégorie': 'PC portable', 'Quantité': 35, 'Année d\'achat ou d\'installation': 2022, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Extraction inventaire SI'},
        {'Catégorie': 'PC portable', 'Quantité': 12, 'Année d\'achat ou d\'installation': 2024, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Extraction inventaire SI'}]),
     ('Cas 2', 'Je ne connais que l\'âge → je calcule l\'année.', [
        {'Catégorie': 'GAB / guichet automatique', 'Modèle / référence (facultatif)': 'CINEO C2560', 'Quantité': 26, 'Année d\'achat ou d\'installation': 2019, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Commentaire': 'Âge 6 ans en 2025'}]),
     ('Cas 3', 'Je ne connais ni la date ni l\'âge exact → estimation avec méthode « Estimation d\'expert ».', [
        {'Catégorie': 'Chaise / fauteuil', 'Quantité': 40, 'Année d\'achat ou d\'installation': 2015, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert', 'Commentaire': 'Mobilier d\'origine, agence ouverte en 2015'}]),
     ('Cas 4', 'Bâtiment.', [
        {'Catégorie': 'Bâtiment', 'Quantité': 1, 'Année d\'achat ou d\'installation': 2008, 'Surface (m²) – bâtiments': 400, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Commentaire': 'Propriété ; rénovation légère en 2019'}]),
     ('Cas 5', 'L\'inventaire est tenu par une direction centrale → « Non applicable ».', [
        {'Catégorie': 'Serveur', 'Statut': 'Non applicable', 'Commentaire': 'Déclaré par la Direction Gestion du Data Center'}]),
     ('Cas 6', 'Équipement absent de la liste → « Autre équipement (préciser) » et nom dans Modèle / référence.', [
        {'Catégorie': 'Autre équipement (préciser)', 'Modèle / référence (facultatif)': 'Compteuse de billets', 'Quantité': 3, 'Année d\'achat ou d\'installation': 2021, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)'}])],
    ['Écrire « Age > 4 ans » ou « Antérieure à 2018 » : donnez une année.',
     'Laisser l\'exemple « Âge = 3 ans » sur toutes les lignes.',
     'Compter le même parc deux fois (par le site et par la direction SI).'],
    'Quantité')

poste_sheet('6-Déchets', '6. Déchets (scope 3)', 'Scope 3 : traitement des déchets',
    'Gestion du Patrimoine ou gestionnaire de site ; Data Center pour les déchets électroniques', 'bordereaux et factures du prestataire de collecte, certificats de destruction d\'archives et de DEEE, pesées',
    ['① Une ligne par type de déchet (liste déroulante) avec la quantité et ce qu\'il devient (recyclage, décharge…).',
     '② Si vous connaissez le poids (kg ou tonnes), donnez-le. Sinon, donnez le nombre de sacs ou de bacs collectés sur l\'année.',
     '③ Si vous avez pesé pendant une semaine ou un mois, indiquez la période : l\'équipe fera l\'extrapolation à l\'année.',
     '④ Si la collecte est faite par la commune sans aucun suivi : « Non disponible » (l\'effectif du site servira à estimer).'],
    [('Type de déchet', 26, 'Choisir dans la liste.', 'flux', None),
     ('Quantité', 10, 'Nombre seul.', None, 'dec'),
     ('Unité', 12, 'kg, tonne, ou nombre de sacs / bacs.', 'unite_dechet', None),
     ('Période couverte', 18, 'Sur quelle durée porte la quantité.', 'periode_dechet', None),
     ('Traitement', 20, 'Ce que devient le déchet.', 'traitement', None),
     ('Prestataire', 16, 'Facultatif.', None, None)],
    [('Cas 1', 'Bordereaux annuels du prestataire.', [
        {'Type de déchet': 'Papier et carton', 'Quantité': 3.2, 'Unité': 'Tonne', 'Période couverte': 'Année entière', 'Traitement': 'Recyclage', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Bordereaux 2025'}]),
     ('Cas 2', 'Je connais seulement le nombre de bacs collectés dans l\'année.', [
        {'Type de déchet': 'Ordures ménagères (mélange)', 'Quantité': 104, 'Unité': 'Bac 240 L', 'Période couverte': 'Année entière', 'Traitement': 'Décharge / enfouissement', 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert', 'Commentaire': '2 bacs par semaine × 52'}]),
     ('Cas 3', 'Pesée sur une semaine.', [
        {'Type de déchet': 'Plastique', 'Quantité': 4.5, 'Unité': 'kg', 'Période couverte': 'Une semaine de pesée', 'Traitement': 'Inconnu', 'Statut': 'Valeur', 'Méthode': 'Extrapolé (ratio, échantillon)', 'Commentaire': 'Semaine du 10 au 14 mars'}]),
     ('Cas 4', 'Destruction d\'archives avec certificat.', [
        {'Type de déchet': 'Archives détruites', 'Quantité': 1.8, 'Unité': 'Tonne', 'Période couverte': 'Année entière', 'Traitement': 'Recyclage', 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'Certificat de destruction'}]),
     ('Cas 5', 'Collecte communale sans aucun suivi.', [
        {'Type de déchet': 'Ordures ménagères (mélange)', 'Statut': 'Non disponible', 'Commentaire': 'Collecte communale ; 19 ETP sur le site'}])],
    ['Donner une quantité « par jour » sans préciser la période.',
     'Mettre 0 alors que des déchets sont produits mais non pesés : utilisez « Non disponible ».'],
    'Quantité', None, 'Unité')

poste_sheet('7-Déplacements pro', '7. Déplacements professionnels (scope 3)', 'Scope 3 : voyages d\'affaires',
    'Pilotage RH et Rémunération ; gestionnaire de site pour les frais locaux', 'extraction de l\'agence de voyage, ordres de mission, notes de frais',
    ['① Avion : une ligne par trajet, avec la ville de départ et d\'arrivée choisies dans la liste (code aéroport), le type de trajet et le nombre de voyages.',
     '② Un aller-retour = UNE ligne « Aller-retour ». Ne créez pas une deuxième ligne pour le retour.',
     '③ Vol avec escale (ex. Casablanca – Paris – Montréal) : une ligne par tronçon.',
     '④ Train, voiture, taxi : indiquez la distance totale en km. Hôtel : le nombre de nuitées. À défaut, le montant HT.',
     '⑤ Indiquez toujours la zone (Maroc ou International) : l\'IFC distingue les deux pour l\'avion et les frais.'],
    [('Mode', 22, 'Moyen de transport, ou hôtel.', 'mode_dep', None),
     ('Ville de départ', 20, 'Choisir dans la liste.', 'ville', None),
     ('Ville d\'arrivée', 20, 'Choisir dans la liste.', 'ville', None),
     ('Zone', 12, 'Maroc ou International (obligatoire pour l\'avion et les frais).', 'zone', None),
     ('Type de trajet', 13, 'Aller simple ou aller-retour.', 'trajet', None),
     ('Classe', 12, 'Pour l\'avion.', 'classe', None),
     ('Nombre de voyages', 10, 'Ou nombre de nuitées pour l\'hôtel.', None, 'dec'),
     ('Distance totale (km)', 12, 'Pour train, voiture, taxi, autocar.', None, 'dec'),
     ('Montant HT (MAD)', 13, 'Si le détail est inconnu.', None, 'dec'),
     ('Carburant du véhicule', 13, 'Voiture personnelle ou de location : diesel, essence…', 'carburant_vp', None)],
    [('Cas 1', 'Vols aller-retour : une ligne « Aller-retour ».', [
        {'Mode': 'Avion', 'Ville de départ': 'Casablanca (CMN)', 'Ville d\'arrivée': 'Paris (CDG / ORY)', 'Zone': 'International', 'Type de trajet': 'Aller-retour', 'Classe': 'Économique', 'Nombre de voyages': 23, 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Extraction agence de voyage'}]),
     ('Cas 2', 'Aller simple (le retour s\'est fait autrement).', [
        {'Mode': 'Avion', 'Ville de départ': 'Oujda (OUD)', 'Ville d\'arrivée': 'Casablanca (CMN)', 'Zone': 'Maroc', 'Type de trajet': 'Aller simple', 'Classe': 'Économique', 'Nombre de voyages': 1, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)'}]),
     ('Cas 3', 'Vol avec escale : une ligne par tronçon.', [
        {'Mode': 'Avion', 'Ville de départ': 'Casablanca (CMN)', 'Ville d\'arrivée': 'Paris (CDG / ORY)', 'Zone': 'International', 'Type de trajet': 'Aller-retour', 'Classe': 'Affaires', 'Nombre de voyages': 1, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Tronçon 1'},
        {'Mode': 'Avion', 'Ville de départ': 'Paris (CDG / ORY)', 'Ville d\'arrivée': 'Montréal (YUL)', 'Zone': 'International', 'Type de trajet': 'Aller-retour', 'Classe': 'Affaires', 'Nombre de voyages': 1, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': 'Tronçon 2'}]),
     ('Cas 4', 'Train, voiture personnelle, hôtel.', [
        {'Mode': 'Train', 'Ville de départ': 'Casablanca (CMN)', 'Ville d\'arrivée': 'Rabat (RBA)', 'Zone': 'Maroc', 'Type de trajet': 'Aller-retour', 'Classe': 'Sans objet', 'Nombre de voyages': 12, 'Distance totale (km)': 2100, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': '12 AR × 175 km'},
        {'Mode': 'Voiture personnelle (km remboursés)', 'Zone': 'Maroc', 'Type de trajet': 'Sans objet', 'Classe': 'Sans objet', 'Distance totale (km)': 8400, 'Carburant du véhicule': 'Diesel', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)', 'Pièce justificative': 'Indemnités kilométriques'},
        {'Mode': 'Hôtel (nuitées)', 'Ville d\'arrivée': 'Dakar (DSS)', 'Zone': 'International', 'Type de trajet': 'Sans objet', 'Classe': 'Sans objet', 'Nombre de voyages': 14, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Commentaire': '14 nuitées'}]),
     ('Cas 5', 'Je n\'ai que le montant des frais.', [
        {'Mode': 'Taxi', 'Zone': 'Maroc', 'Type de trajet': 'Sans objet', 'Classe': 'Sans objet', 'Montant HT (MAD)': 27117, 'Statut': 'Valeur', 'Méthode': 'Estimé à partir d\'un montant', 'Commentaire': 'Frais de déplacement Maroc hors avion'}])],
    ['Écrire les destinations en texte libre (« MAROC-Dubai abu dhabi », « Amesterdam ») : utilisez la liste.',
     'Saisir l\'aller et le retour sur deux lignes ET choisir « Aller-retour ».',
     'Donner seulement le nombre de billets sans les trajets.'],
    'Nombre de voyages', ['Distance totale (km)', 'Montant HT (MAD)'], None)

# distance column also acceptable: tweak handled in instructions (qty = Nombre de voyages or montant)

poste_sheet('8-Domicile et visiteurs', '8. Trajets domicile-travail et déplacements des visiteurs (scope 3)', 'Scope 3 : déplacements quotidiens des employés et venue des clients et visiteurs',
    'RH ou gestionnaire de site', 'enquête mobilité (questionnaire rapide), registre des visiteurs, fréquentation de l\'agence, connaissance du site',
    ['① Une ligne par mode de transport utilisé par les employés du site.',
     '② Indiquez combien d\'employés l\'utilisent, la distance ALLER moyenne (domicile → bureau) et le nombre de jours de présence par semaine.',
     '③ Télétravail : comptez seulement les jours de présence (ex. 3 jours si 2 jours de télétravail).',
     '④ La somme des employés doit correspondre à l\'effectif du site (onglet Identification).',
     '⑤ Visiteurs et clients (demandé par l\'IFC) : catégorie « Visiteurs et clients », nombre de visites dans l\'année par mode et distance aller moyenne ; une estimation suffit.'],
    [('Catégorie', 16, 'Employés ou Visiteurs et clients.', 'categorie_dt', None),
     ('Mode principal', 26, 'Choisir dans la liste.', 'mode_dt', None),
     ('Nombre de personnes', 12, 'Employés : nombre d\'employés. Visiteurs : nombre de visites dans l\'année.', None, 'int'),
     ('Distance aller moyenne (km)', 13, 'Distance d\'un aller simple.', None, 'dec'),
     ('Jours de présence par semaine', 13, 'Employés seulement, entre 0 et 7.', None, 'days')],
    [('Cas 1', 'Résultats d\'une enquête.', [
        {'Catégorie': 'Employés', 'Mode principal': 'Voiture personnelle diesel', 'Nombre de personnes': 40, 'Distance aller moyenne (km)': 12, 'Jours de présence par semaine': 5, 'Statut': 'Valeur', 'Méthode': 'Extrapolé (ratio, échantillon)', 'Pièce justificative': 'Enquête mobilité 2025', 'Commentaire': '62 % de réponses'},
        {'Catégorie': 'Employés', 'Mode principal': 'Tramway', 'Nombre de personnes': 20, 'Distance aller moyenne (km)': 8, 'Jours de présence par semaine': 5, 'Statut': 'Valeur', 'Méthode': 'Extrapolé (ratio, échantillon)', 'Pièce justificative': 'Enquête mobilité 2025'}]),
     ('Cas 2', 'Télétravail deux jours par semaine.', [
        {'Catégorie': 'Employés', 'Mode principal': 'Train', 'Nombre de personnes': 6, 'Distance aller moyenne (km)': 90, 'Jours de présence par semaine': 3, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert'}]),
     ('Cas 3', 'Navette affrétée par la banque (dans ce cas, déclarer aussi le contrat dans 4-Achats).', [
        {'Catégorie': 'Employés', 'Mode principal': 'Transport du personnel (navette)', 'Nombre de personnes': 30, 'Distance aller moyenne (km)': 15, 'Jours de présence par semaine': 5, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert'}]),
     ('Cas 4', 'Pas d\'enquête possible cette année.', [
        {'Catégorie': 'Employés', 'Mode principal': 'Voiture personnelle essence', 'Statut': 'Non disponible', 'Commentaire': 'Enquête prévue l\'an prochain'}]),
     ('Cas 5', 'Visiteurs et clients : estimation à partir de la fréquentation (une ligne par mode).', [
        {'Catégorie': 'Visiteurs et clients', 'Mode principal': 'Voiture personnelle essence', 'Nombre de personnes': 9000, 'Distance aller moyenne (km)': 6, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert', 'Commentaire': '60 visites par jour ouvré, 60 % en voiture'},
        {'Catégorie': 'Visiteurs et clients', 'Mode principal': 'Petit taxi', 'Nombre de personnes': 4500, 'Distance aller moyenne (km)': 5, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert'}])],
    ['Donner la distance aller-retour dans la colonne « aller ».',
     'Oublier le nombre de jours de présence.'],
    'Nombre de personnes')

poste_sheet('9-Transport fonds-courrier', '9. Transport de fonds et courrier (scope 3)', 'Scope 3 : transport de marchandises amont et aval',
    'Achats, Sécurité, Gestion du Patrimoine, directions qui envoient du courrier aux clients', 'états mensuels du transporteur de fonds (ex. Brinks, G4S), factures postales, planning des navettes',
    ['① Transport de fonds : une ligne par site desservi avec le nombre de passages dans l\'année et la distance aller-retour depuis la base du transporteur.',
     '② Courrier : nombre d\'envois et poids total si connu ; sinon le montant HT de la facture.',
     '③ Navette interne : nombre de tournées et distance aller-retour d\'une tournée.',
     '④ Un état mensuel du transporteur peut être joint tel quel : indiquez le total de l\'année.'],
    [('Type', 22, 'Choisir dans la liste.', 'type_transport', None),
     ('Site desservi ou trajet', 24, 'Ex. BP Agadir 11 Janvier ; Casa → Rabat.', None, None),
     ('Nombre de passages ou d\'envois par an', 14, 'Total de l\'année.', None, 'dec'),
     ('Distance aller-retour (km)', 13, 'Par passage ou par tournée.', None, 'dec'),
     ('Poids total (kg)', 11, 'Pour le courrier, si connu.', None, 'dec'),
     ('Prestataire', 14, 'Facultatif.', None, None),
     ('Montant HT (MAD)', 13, 'Si le détail est inconnu.', None, 'dec'),
     ('Type de véhicule', 16, 'Véhicule utilisé (fourgon blindé, utilitaire…).', 'type_vehicule', None),
     ('Consommation aux 100 km (L)', 12, 'Facultatif : consommation moyenne du véhicule.', None, 'dec')],
    [('Cas 1', 'État des passages du transporteur de fonds.', [
        {'Type': 'Transport de fonds', 'Site desservi ou trajet': 'BP Agadir 11 Janvier', 'Nombre de passages ou d\'envois par an': 224, 'Distance aller-retour (km)': 14, 'Prestataire': 'Brinks', 'Type de véhicule': 'Fourgon blindé', 'Consommation aux 100 km (L)': 16, 'Statut': 'Valeur', 'Méthode': 'Mesuré (facture, compteur, pesée)', 'Pièce justificative': 'État récapitulatif Brinks 2025', 'Commentaire': 'Somme des 12 mois'}]),
     ('Cas 2', 'Courrier envoyé aux clients.', [
        {'Type': 'Courrier envoyé aux clients', 'Site desservi ou trajet': 'Relevés et mises en demeure', 'Nombre de passages ou d\'envois par an': 116030, 'Poids total (kg)': 939, 'Prestataire': 'Barid Al-Maghrib', 'Statut': 'Valeur', 'Méthode': 'Extrait d\'un système (ERP, inventaire)'}]),
     ('Cas 3', 'Navette courrier interne.', [
        {'Type': 'Navette courrier interne', 'Site desservi ou trajet': 'Siège BPR → agences', 'Nombre de passages ou d\'envois par an': 250, 'Distance aller-retour (km)': 85, 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert', 'Commentaire': '1 tournée par jour ouvré'}]),
     ('Cas 4', 'Seulement la facture.', [
        {'Type': 'Colis / messagerie', 'Montant HT (MAD)': 12500, 'Statut': 'Valeur', 'Méthode': 'Estimé à partir d\'un montant'}]),
     ('Cas 5', 'Livraisons des fournisseurs (fret amont), quand le fournisseur la facture à part : nombre de livraisons, distance, type de véhicule.', [
        {'Type': 'Livraisons fournisseurs (fret amont)', 'Site desservi ou trajet': 'Fournitures depuis Casablanca', 'Nombre de passages ou d\'envois par an': 24, 'Distance aller-retour (km)': 680, 'Type de véhicule': 'Utilitaire léger', 'Statut': 'Valeur', 'Méthode': 'Estimation d\'expert', 'Commentaire': '2 livraisons par mois'}])],
    ['Écrire un calcul dans la cellule (« =(225+714)/1000 ») : donnez le résultat et expliquez en commentaire.',
     'Additionner tous les sites sans détail quand l\'état par site existe.'],
    'Nombre de passages ou d\'envois par an', 'Montant HT (MAD)', None)

# ---------------- Contrôle ----------------
ct = wb.create_sheet('Contrôle')
for col, w in zip('ABCDE', [30, 44, 16, 16, 16]): ct.column_dimensions[col].width = w
merged(ct, 1, 1, 5, 'Contrôle avant envoi (se remplit automatiquement)', font(bold=True, size=16, color=WHITE), NAVY, 40, WRAPC)
add_logo(ct, 'F1'); ct.column_dimensions['F'].width = 15; ct.sheet_view.showGridLines = False
merged(ct, 2, 1, 5, 'Le fichier est prêt quand la colonne « Alertes » vaut 0 sur toutes les lignes. Une alerte se corrige dans l\'onglet concerné (voir la colonne « Contrôle automatique »).', font(), LBLUE, 30)
for j, h in enumerate(['Onglet', 'Poste', 'Lignes saisies', 'Lignes OK', 'Alertes'], start=1):
    c = ct.cell(row=4, column=j, value=h); c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY); c.border = BOX
for i, (nm, tt, rng) in enumerate(CONTROL_SHEETS, start=5):
    ct.cell(row=i, column=1, value=nm); ct.cell(row=i, column=2, value=tt)
    ct.cell(row=i, column=4, value=f'=COUNTIF({rng},"OK")')
    ct.cell(row=i, column=5, value=f'=COUNTIF({rng},"⚠*")')
    ct.cell(row=i, column=3, value=f'=D{i}+E{i}')
    for j in range(1, 6):
        c = ct.cell(row=i, column=j); c.border = BOX; c.font = font()
last = 4 + len(CONTROL_SHEETS)
t = last + 1
ct.cell(row=t, column=2, value='Total')
for col in 'CDE':
    ct[f'{col}{t}'] = f'=SUM({col}5:{col}{last})'
for j in range(1, 6):
    c = ct.cell(row=t, column=j); c.font = font(bold=True); c.border = BOX; c.fill = fill(LBLUE)
ct.cell(row=t + 2, column=1, value='État du fichier').font = font(bold=True)
ct[f'B{t+2}'] = f'=IF(C{t}=0,"Aucune donnée saisie",IF(E{t}=0,"✓ Prêt à envoyer","✗ "&E{t}&" alerte(s) à corriger"))'
ct[f'B{t+2}'].font = font(bold=True, size=12, color=GREEN)
ct.cell(row=t + 3, column=1, value='Identification complète').font = font(bold=True)
ct[f'B{t+3}'] = "=IF(COUNTA(Identification!B5:B14)+COUNTA(Identification!B17:B18)>=11,\"✓ Oui\",\"✗ Compléter l'onglet Identification\")"
ct[f'B{t+3}'].font = font(bold=True, color=NAVY)

# ---------------- Contrôles de cohérence ----------------
def R(sheet, header):
    i = REG[sheet]; col = L(i['idx'][header])
    return f"'{sheet}'!${col}${i['first']}:${col}${i['last']}"
def SI(sheet, sumh, conds=(), valeur=True):
    parts = [R(sheet, sumh)]
    for h, crit in conds:
        parts += [R(sheet, h), crit]
    if valeur:
        parts += [R(sheet, 'Statut'), '"Valeur"']
    return 'SUMIFS(' + ','.join(parts) + ')'
def SP(sheet, prods, conds=(), extra=''):
    terms = [f'({R(sheet, h)}{op}{crit})' for h, op, crit in conds] + [f'({R(sheet, "Statut")}="Valeur")'] + [R(sheet, h) for h in prods]
    return 'SUMPRODUCT(' + '*'.join(terms) + (('*' + extra) if extra else '') + ')'
E, CB, CL, A, EQ, DC, DP, DT, TR = '1-Électricité', '2-Carburants et gaz', '3-Climatisation', '4-Achats', '5-Équipements', '6-Déchets', '7-Déplacements pro', '8-Domicile et visiteurs', '9-Transport fonds-courrier'
EX = 'Identification!$B$5'
RESEAU = '+'.join(SI(E, 'Quantité (kWh)', [('Usage du compteur', f'"{u}"')]) for u in ['Site entier', 'Partie du site', 'Partagé avec un tiers'])
tc = t + 6
ct.cell(row=tc, column=1, value='CONTRÔLES DE COHÉRENCE (à vérifier, non bloquants)').font = font(bold=True, size=12, color=GREEN)
for j, h in enumerate(['Contrôle', 'Résultat'], start=1):
    c = ct.cell(row=tc + 1, column=j, value=h); c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY); c.border = BOX
nm = f'COUNTIFS({R(E, "Mois")},"<>Année entière",{R(E, "Mois")},"<>",{R(E, "Usage du compteur")},"Site entier",{R(E, "Statut")},"Valeur")'
emp = SI(DT, 'Nombre de personnes', [('Catégorie', '"Employés"')])
checks = [
    ('Électricité : 12 mois par compteur', f'=IF({nm}=0,"—",IF(MOD({nm},12)=0,"✓ "&{nm}&" mois saisis","⚠ "&{nm}&" mois saisis : vérifier les mois manquants (12 par compteur)"))'),
    ('Électricité : consommation par m²', f'=IF(OR(Identification!$B$12="",Identification!$B$12=0,({RESEAU})=0),"—",IF(OR(({RESEAU})/Identification!$B$12<20,({RESEAU})/Identification!$B$12>500),"⚠ "&ROUND(({RESEAU})/Identification!$B$12,0)&" kWh/m² : valeur inhabituelle, à vérifier","✓ "&ROUND(({RESEAU})/Identification!$B$12,0)&" kWh/m²"))'),
    ('Domicile-travail : employés déclarés / effectif', f'=IF(OR(Identification!$B$13="",Identification!$B$13=0,({emp})=0),"—",IF(ABS(({emp})-Identification!$B$13)/Identification!$B$13>0.1,"⚠ "&({emp})&" employés déclarés pour "&Identification!$B$13&" ETP","✓ cohérent avec l\'effectif"))'),
    ('Déplacements : vols sans zone', f'=IF(COUNTIFS({R(DP, "Mode")},"Avion",{R(DP, "Zone")},"")=0,"✓ OK","⚠ "&COUNTIFS({R(DP, "Mode")},"Avion",{R(DP, "Zone")},"")&" ligne(s) d\'avion sans zone Maroc / International")'),
    ('Carburants : montant sans prix unitaire', f'=IF(COUNTIFS({R(CB, "Montant HT (MAD)")},">0",{R(CB, "Quantité")},"",{R(CB, "Prix unitaire HT (MAD par unité)")},"")=0,"✓ OK","⚠ "&COUNTIFS({R(CB, "Montant HT (MAD)")},">0",{R(CB, "Quantité")},"",{R(CB, "Prix unitaire HT (MAD par unité)")},"")&" ligne(s) avec un montant mais sans prix unitaire")'),
    ('Électricité : contrat renseigné', f'=IF(COUNTIFS({R(E, "Statut")},"Valeur",{R(E, "Type de contrat")},"")=0,"✓ OK","⚠ "&COUNTIFS({R(E, "Statut")},"Valeur",{R(E, "Type de contrat")},"")&" ligne(s) sans type de contrat")'),
]
for k, (lab, f) in enumerate(checks):
    rr = tc + 2 + k
    a = ct.cell(row=rr, column=1, value=lab); a.font = font(); a.border = BOX
    b = ct.cell(row=rr, column=2, value=f); b.font = font(bold=True, color=GREEN); b.border = BOX
    ct.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=5)

# ---------------- Export IFC ----------------
xw = wb.create_sheet('Export IFC')
for col, w in zip('ABCDEFGH', [7, 30, 40, 14, 12, 16, 52, 22]): xw.column_dimensions[col].width = w
merged(xw, 1, 1, 8, 'Export au format du canevas IFC (se remplit automatiquement)', font(bold=True, size=16, color=WHITE), NAVY, 40, WRAPC)
add_logo(xw, 'I1'); xw.column_dimensions['I'].width = 15; xw.sheet_view.showGridLines = False
merged(xw, 2, 1, 8, 'Cet onglet recompose vos saisies selon les rubriques du canevas IFC « Carbon Footprint Data Request », pour assurer la continuité avec la mesure 2022. Seules les lignes au statut « Valeur » sont prises en compte. Ne rien saisir, sauf le nombre de jours ouvrés.', font(), LBLUE, 44)
xw['B4'] = 'Exercice'; xw['D4'] = f'={EX}'
xw['B5'] = 'Jours ouvrés par an (conversion des déchets en kg par jour)'; xw['D5'] = 250
for a in ('B4', 'B5'): xw[a].font = font(bold=True)
xw['D5'].fill = fill(YEL); xw['D5'].font = font(bold=True, color='1F3FBF'); xw['D5'].protection = UNLOCK; xw['D5'].border = BOX
xw['D4'].border = BOX
JO = '$D$5'
HDR = 7
for j, h in enumerate(['Scope', 'Poste (canevas IFC)', 'Source d\'émission (canevas IFC)', 'Quantité', 'Unité', 'Montant HT (MAD)', 'Complément (âge, charge, distance…)', 'Onglet source'], start=1):
    c = xw.cell(row=HDR, column=j, value=h); c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY); c.border = BOX; c.alignment = WRAPC
xw.row_dimensions[HDR].height = 30
ROWS = []
def sec(t): ROWS.append(('sec', t))
def row(scope, poste, src, q=None, unit='', m=None, comp=None, ong=''):
    ROWS.append(('row', scope, poste, src, q, unit, m, comp, ong))
def plus(*xs): return '+'.join(x for x in xs if x)
def guard(x, txt='⚠ vérifier les saisies'): return f'IFERROR({x},"{txt}")'

sec('Scope 1 — Émissions directes des sources fixes')
for comb in ['Butane', 'Propane']:
    q = plus(SI(CB, 'Quantité', [('Combustible', f'"{comb}"'), ('Unité', '"kg"')]), '12*' + SI(CB, 'Quantité', [('Combustible', f'"{comb}"'), ('Unité', '"Bouteille 12 kg"')]), '3*' + SI(CB, 'Quantité', [('Combustible', f'"{comb}"'), ('Unité', '"Bouteille 3 kg"')]))
    row('1', 'Cuisson, chauffage, eau sanitaire', comb, q, 'kg', SI(CB, 'Montant HT (MAD)', [('Combustible', f'"{comb}"')]), None, CB)
row('1', 'Cuisson, chauffage, eau sanitaire', 'Bois', SI(CB, 'Quantité', [('Combustible', '"Bois"'), ('Unité', '"kg"')]), 'kg', SI(CB, 'Montant HT (MAD)', [('Combustible', '"Bois"')]), None, CB)
row('1', 'Autres sources fixes', 'Groupe électrogène (gasoil)', plus(SI(CB, 'Quantité', [('Usage', '"Groupe électrogène"'), ('Unité', '"Litre"')]), '1000*' + SI(CB, 'Quantité', [('Usage', '"Groupe électrogène"'), ('Unité', '"m³"')])), 'litres', SI(CB, 'Montant HT (MAD)', [('Usage', '"Groupe électrogène"')]), None, CB)
sec('Scope 1 — Émissions directes des sources mobiles')
for comb, lab in [('Gasoil', 'Parc automobile fonctionnant au gasoil'), ('Essence', 'Parc automobile fonctionnant à l\'essence')]:
    row('1', 'Sources mobiles', lab, SI(CB, 'Quantité', [('Usage', '"Véhicules de service"'), ('Combustible', f'"{comb}"'), ('Unité', '"Litre"')]), 'litres',
        SI(CB, 'Montant HT (MAD)', [('Usage', '"Véhicules de service"'), ('Combustible', f'"{comb}"')]),
        f'"Kilométrage déclaré : "&{SI(CB, "Kilométrage annuel (km)", [("Usage", chr(34) + "Véhicules de service" + chr(34)), ("Combustible", chr(34) + comb + chr(34))])}&" km"', CB)
row('1', 'Sources mobiles', 'Autre (préciser)', SI(CB, 'Quantité', [('Combustible', '"Autre (préciser)"')]), 'voir onglet 2', SI(CB, 'Montant HT (MAD)', [('Combustible', '"Autre (préciser)"')]), None, CB)
sec('Scope 1 — Émissions directes fugitives (nature du gaz, quantité, âge)')
for veh, lab in [(False, 'Climatiseurs (bureaux et locaux)'), (True, 'Climatiseurs des véhicules')]:
    op = '=' if veh else '<>'
    for fl in LISTS['fluide']:
        cond = [('Fluide', '=', f'"{fl}"'), ("Type d\'appareil", op, '"Climatiseur de véhicule"')]
        rech = SP(CL, ["Quantité rechargée dans l\'année (kg)"], cond)
        charge = SP(CL, ['Charge nominale par appareil (kg)', "Nombre d\'appareils identiques"], cond)
        nb = SP(CL, ["Nombre d\'appareils identiques"], cond)
        yr = SP(CL, ['Année de mise en service', "Nombre d\'appareils identiques"], cond)
        den = SP(CL, ["Nombre d\'appareils identiques"], cond, f"({R(CL, 'Année de mise en service')}>0)")
        recup = SP(CL, ['Fluide récupéré (kg)'], cond)
        comp = guard(f'IF({nb}=0,"","{"Appareils" if not veh else "Véhicules"} : "&{nb}&" ; charge totale : "&ROUND({charge},1)&" kg"&IF({den}>0," ; âge moyen : "&ROUND({EX}-{yr}/{den},0)&" ans","")&IF({recup}>0," ; récupéré : "&ROUND({recup},1)&" kg",""))')
        row('1', lab, f'Gaz {fl} — quantité rechargée', guard(rech), 'kg', None, comp, CL)
sec('Scope 2 — Émissions indirectes liées à l\'électricité')
row('2', 'Électricité', 'Électricité du réseau (compteurs principaux)', RESEAU, 'kWh', plus(*[SI(E, 'Montant HT (MAD)', [('Usage du compteur', f'"{u}"')]) for u in ['Site entier', 'Partie du site', 'Partagé avec un tiers']]),
    f'"Dont électricité renouvelable sous contrat : "&{SI(E, "Quantité (kWh)", [("Type de contrat", chr(34) + "Électricité renouvelable sous contrat (loi 13-09)" + chr(34))])}&" kWh"', E)
row('2', 'Électricité', 'Électricité photovoltaïque (produite)', SI(E, 'Quantité (kWh)', [('Usage du compteur', '"Production photovoltaïque"')]), 'kWh', None,
    f'"Puissance installée : "&{SI(E, "Puissance installée (kWc)", [("Usage du compteur", chr(34) + "Production photovoltaïque" + chr(34))])}&" kWc"', E)
vek = SI(CB, 'Kilométrage annuel (km)', [('Combustible', '"Électricité (véhicule électrique)"')])
vec = SP(CB, ['Kilométrage annuel (km)', 'Consommation aux 100 km'], [('Combustible', '=', '"Électricité (véhicule électrique)"')])
row('2', 'Électricité', 'Véhicules électriques', plus(SI(E, 'Quantité (kWh)', [('Usage du compteur', '"Recharge véhicules électriques"')]), SI(CB, 'Quantité', [('Combustible', '"Électricité (véhicule électrique)"'), ('Unité', '"kWh"')]), guard(f'{vec}/100', '0')), 'kWh',
    None, guard(f'"Kilométrage : "&{vek}&" km"&IF({vek}>0," ; consommation moyenne : "&ROUND({vec}/{vek},1)&" kWh/100 km","")'), f'{E}, {CB}')
row('2', 'Vapeur, chaleur ou froid', 'Chaleur ou froid acheté à un réseau', SI(E, 'Quantité (kWh)', [('Usage du compteur', '"Chaleur ou froid acheté à un réseau"')]), 'kWh', None, '"La climatisation électrique est comprise dans l\'électricité du réseau"', E)
it = SI(E, 'Quantité (kWh)', [('Usage du compteur', '"Équipements informatiques (sous-compteur IT)"')])
row('2', 'Électricité (information)', 'Data center : équipements informatiques', it, 'kWh', None, guard(f'IF({it}>0,"PUE indicatif : "&ROUND(({RESEAU})/{it},2),"")', ''), E)
sec('Scope 3 — Achats')
fam = lambda f_: SI(A, 'Montant HT (MAD)', [("Famille d\'achat", f'"{f_}"')])
qte = lambda f_, u: SI(A, 'Quantité (facultatif)', [("Famille d\'achat", f'"{f_}"'), ('Unité', f'"{u}"')])
row('3', 'Achats', 'Petites fournitures (papier, consommables, fournitures, hygiène)', None, '', plus(*[fam(x) for x in ['Papier vierge', 'Papier recyclé', 'Consommables d\'impression (toners, cartouches)', 'Fournitures de bureau', 'Produits d\'hygiène et d\'entretien']]), None, A)
for pap in ['Papier vierge', 'Papier recyclé']:
    row('3', 'Achats', pap, plus(qte(pap, 'Rame de 500 feuilles') + '*2.5', qte(pap, 'Tonne') + '*1000', qte(pap, 'kg')), 'kg', fam(pap), f'"Rame A4 80 g = 2,5 kg ; rames déclarées : "&{qte(pap, "Rame de 500 feuilles")}', A)
for f_, lab in [('Télécoms et abonnements internet', 'Abonnements internet'), ('Impressions externalisées et publicité', 'Impressions publicité / autres'), ('Maintenance des équipements et bâtiments', 'Maintenance'), ('Gardiennage', 'Gardiennage'), ('Location de véhicules', 'Locations de véhicules'), ('Loyers et charges locatives', 'Location des agences (loyers)')]:
    row('3', 'Achats', lab, None, '', fam(f_), None, A)
row('3', 'Achats', 'Repas, café', SI(A, 'Quantité (facultatif)', [("Famille d\'achat", '"Restauration (repas, café)"'), ('Unité', '"Repas"')]), 'repas', fam('Restauration (repas, café)'),
    f'"Tasses / doses : "&{qte("Restauration (repas, café)", "Tasse / dose")}', A)
for mat in LISTS['materiau']:
    row('3', 'Achats de matériaux pour travaux', mat, plus(SI(A, 'Quantité (facultatif)', [("Famille d\'achat", '"Travaux et matériaux de construction"'), ('Matériau (travaux)', f'"{mat}"'), ('Unité', '"kg"')]), '1000*' + SI(A, 'Quantité (facultatif)', [("Famille d\'achat", '"Travaux et matériaux de construction"'), ('Matériau (travaux)', f'"{mat}"'), ('Unité', '"Tonne"')])), 'kg',
        SI(A, 'Montant HT (MAD)', [("Famille d\'achat", '"Travaux et matériaux de construction"'), ('Matériau (travaux)', f'"{mat}"')]), None, A)
row('3', 'Achats', 'Autres achats (nettoyage, logiciels, conseil, autres)', None, '', plus(*[fam(x) for x in ['Nettoyage', 'Logiciels et services informatiques', 'Conseil et autres services', 'Autres achats (préciser)']]), None, A)
sec('Scope 3 — Immobilisations (liste, âge, superficie)')
for cat in LISTS['categorie_immo']:
    cond = [('Catégorie', '=', f'"{cat}"')]
    nb = SP(EQ, ['Quantité'], cond); yr = SP(EQ, ["Année d\'achat ou d\'installation", 'Quantité'], cond)
    den = SP(EQ, ['Quantité'], cond, f"({R(EQ, "Année d'achat ou d'installation")}>0)")
    comp = guard(f'IF({den}=0,"","Âge moyen : "&ROUND({EX}-{yr}/{den},1)&" ans")' + ("&\" ; superficie : \"&" + SI(EQ, 'Surface (m²) – bâtiments', [('Catégorie', f'"{cat}"')]) + "&\" m²\"" if cat == 'Bâtiment' else ''))
    row('3', 'Immobilisations', cat, guard(nb), 'unités', None, comp, EQ)
sec('Scope 3 — Déchets (quantité journalière et devenir)')
for fl in LISTS['flux']:
    terms = []
    for u, fu in [('kg', ''), ('Tonne', '1000*')]:
        for per, fp in [('Année entière', ''), ('Un mois', '12*'), ('Une semaine de pesée', f'{JO}/5*')]:
            terms.append(fu + fp + SI(DC, 'Quantité', [('Type de déchet', f'"{fl}"'), ('Unité', f'"{u}"'), ('Période couverte', f'"{per}"')]))
    an = '(' + '+'.join(terms) + ')'
    cont = '+'.join(fp + SI(DC, 'Quantité', [('Type de déchet', f'"{fl}"'), ('Unité', f'"{u}"'), ('Période couverte', f'"{per}"')]) for u in ['Sac 100 L', 'Bac 240 L', 'Bac 660 L', 'Benne'] for per, fp in [('Année entière', ''), ('Un mois', '12*'), ('Une semaine de pesée', f'{JO}/5*')])
    rec = SI(DC, 'Quantité', [('Type de déchet', f'"{fl}"'), ('Traitement', '"Recyclage"')])
    comp = guard(f'"Total annuel : "&ROUND({an},0)&" kg"&IF(({cont})>0," ; + "&ROUND({cont},0)&" sacs, bacs ou bennes par an à convertir","")&IF(COUNTIFS({R(DC, "Type de déchet")},"{fl}",{R(DC, "Traitement")},"Recyclage")>0," ; recyclage déclaré","")')
    row('3', 'Déchets', fl, guard(f'{an}/{JO}'), 'kg / jour', None, comp, DC)
sec('Scope 3 — Transport de marchandises amont')
for typ, lab in [('Livraisons fournisseurs (fret amont)', 'Fret amont matière première'), ('Transport de fonds', 'Fret interne transport de fonds')]:
    km = SP(TR, ["Nombre de passages ou d\'envois par an", 'Distance aller-retour (km)'], [('Type', '=', f'"{typ}"')])
    conso = f'AVERAGEIFS({R(TR, "Consommation aux 100 km (L)")},{R(TR, "Type")},"{typ}",{R(TR, "Consommation aux 100 km (L)")},">0")'
    row('3', 'Transport amont', lab, guard(km), 'km par an', SI(TR, 'Montant HT (MAD)', [('Type', f'"{typ}"')]), f'IFERROR("Consommation moyenne : "&ROUND({conso},1)&" L/100 km","")', TR)
sec('Scope 3 — Déplacements domicile-travail (nombre, distance moyenne aller-retour)')
GROUPS_EMP = [('Véhicule particulier gasoil', ['Voiture personnelle diesel']), ('Véhicule particulier essence', ['Voiture personnelle essence', 'Voiture hybride']),
              ('Moto', ['Moto / scooter', 'Moto > 50 cm³', 'Moto < 50 cm³']), ('Bus', ['Bus', 'Transport du personnel (navette)']), ('Taxi', ['Grand taxi', 'Petit taxi']),
              ('Tramway', ['Tramway']), ('Train', ['Train']), ('Autres modalités', ['Voiture électrique', 'Covoiturage', 'Avion', 'À pied / vélo'])]
GROUPS_VIS = [('Voiture', ['Voiture personnelle diesel', 'Voiture personnelle essence', 'Voiture hybride', 'Voiture électrique', 'Covoiturage']), ('Bus', ['Bus', 'Transport du personnel (navette)']),
              ('Petit taxi', ['Petit taxi', 'Grand taxi']), ('Chemin de fer et tram', ['Train', 'Tramway']), ('Avion', ['Avion']), ('Moto > 50 cm³', ['Moto > 50 cm³', 'Moto / scooter']), ('Moto < 50 cm³', ['Moto < 50 cm³'])]
def dt_rows(cat, groups, poste, unit):
    for lab, modes in groups:
        n = '+'.join(SI(DT, 'Nombre de personnes', [('Catégorie', f'"{cat}"'), ('Mode principal', f'"{m_}"')]) for m_ in modes)
        nd = '+'.join(SP(DT, ['Nombre de personnes', 'Distance aller moyenne (km)'], [('Catégorie', '=', f'"{cat}"'), ('Mode principal', '=', f'"{m_}"')]) for m_ in modes)
        jd = '+'.join(SP(DT, ['Nombre de personnes', 'Jours de présence par semaine'], [('Catégorie', '=', f'"{cat}"'), ('Mode principal', '=', f'"{m_}"')]) for m_ in modes)
        extra = f'&" ; jours de présence moyens : "&ROUND(({jd})/({n}),1)' if cat == 'Employés' else ''
        row('3', poste, lab, guard(n), unit, None, guard(f'IF(({n})=0,"","Distance moyenne aller-retour : "&ROUND(2*({nd})/({n}),1)&" km"{extra})'), DT)
dt_rows('Employés', GROUPS_EMP, 'Déplacements domicile-travail', 'employés')
sec('Scope 3 — Déplacements professionnels')
for zone in ['Maroc', 'International']:
    row('3', 'Déplacements professionnels', f'Déplacements du personnel {"au Maroc" if zone == "Maroc" else "à l" + chr(39) + "international"}, par avion', SI(DP, 'Nombre de voyages', [('Mode', '"Avion"'), ('Zone', f'"{zone}"')]), 'voyages', None,
        f'"Dont aller-retour : "&{SI(DP, "Nombre de voyages", [("Mode", chr(34) + "Avion" + chr(34)), ("Zone", chr(34) + zone + chr(34)), ("Type de trajet", chr(34) + "Aller-retour" + chr(34))])}&" ; détail des trajets dans l\'onglet 7"', DP)
for zone in ['Maroc', 'International']:
    m_ = f'{SI(DP, "Montant HT (MAD)", [("Zone", chr(34) + zone + chr(34))])}-{SI(DP, "Montant HT (MAD)", [("Zone", chr(34) + zone + chr(34)), ("Mode", chr(34) + "Avion" + chr(34))])}'
    row('3', 'Déplacements professionnels', f'Frais de déplacement {"au Maroc" if zone == "Maroc" else "à l" + chr(39) + "international"} hors billet d\'avion', None, '', m_, None, DP)
for mode, lab, unit, col in [('Voiture personnelle (km remboursés)', 'Véhicules personnels pour déplacements professionnels', 'km', 'Distance totale (km)'), ('Train', 'Train', 'km', 'Distance totale (km)'), ('Taxi', 'Taxi', 'km', 'Distance totale (km)'),
                             ('Voiture de location', 'Voiture de location', 'km', 'Distance totale (km)'), ('Autocar', 'Autocar', 'km', 'Distance totale (km)'), ('Hôtel (nuitées)', 'Hôtel', 'nuitées', 'Nombre de voyages')]:
    row('3', 'Déplacements professionnels', lab, SI(DP, col, [('Mode', f'"{mode}"')]), unit, None, '"Détail ; les montants sont dans les frais de déplacement ci-dessus"', DP)
sec('Scope 3 — Flotte de véhicules')
nbv = SI(EQ, 'Quantité', [('Catégorie', '"Véhicule de service"')])
row('3', 'Flotte de véhicules', 'Flotte de véhicules de la banque', nbv, 'véhicules', None,
    f'"Kilométrage annuel déclaré : "&{SI(CB, "Kilométrage annuel (km)", [("Usage", chr(34) + "Véhicules de service" + chr(34))])}&" km"', f'{EQ}, {CB}')
sec('Scope 3 — Déplacements clients et visiteurs (nombre, distance moyenne aller-retour)')
dt_rows('Visiteurs et clients', GROUPS_VIS, 'Déplacements clients et visiteurs', 'visites')
sec('Scope 3 — Transport de marchandises et distribution aval')
for typ, lab in [('Courrier envoyé aux clients', 'Courriers et marchandises liés à la clientèle'), ('Navette courrier interne', 'Courriers en interne'), ('Colis / messagerie', 'Factures, appels d\'offres, colis')]:
    km = SP(TR, ["Nombre de passages ou d\'envois par an", 'Distance aller-retour (km)'], [('Type', '=', f'"{typ}"')])
    row('3', 'Transport aval', lab, f'{SI(TR, "Poids total (kg)", [("Type", chr(34) + typ + chr(34))])}/1000', 'tonnes', SI(TR, 'Montant HT (MAD)', [('Type', f'"{typ}"')]),
        guard(f'"Kilométrage annuel : "&ROUND({km},0)&" km ; envois : "&{SI(TR, "Nombre de passages ou d" + chr(39) + "envois par an", [("Type", chr(34) + typ + chr(34))])}'), TR)

rr = HDR + 1
for item in ROWS:
    if item[0] == 'sec':
        merged(xw, rr, 1, 8, item[1], font(bold=True, color=WHITE), GREEN, 20, WRAPC)
    else:
        _, scope, poste, src, q, unit, m_, comp, ong = item
        vals = [scope, poste, src, ('=' + q) if q else None, unit, ('=' + m_) if m_ else None, ('=' + comp) if comp else None, ong]
        for j, v in enumerate(vals, start=1):
            c = xw.cell(row=rr, column=j, value=v); c.font = font(size=9); c.border = BOX; c.alignment = WRAP
            c.fill = fill(LBLUE if (rr % 2) else 'FBF6F0')
        xw.cell(row=rr, column=4).number_format = '#,##0.00'; xw.cell(row=rr, column=6).number_format = '#,##0'
        xw.cell(row=rr, column=4).font = font(size=9, bold=True, color=GREEN)
    rr += 1
xw.freeze_panes = f'A{HDR + 1}'
xw.page_setup.orientation = 'landscape'; xw.page_setup.fitToWidth = 1; xw.page_setup.fitToHeight = 0; xw.sheet_properties.pageSetUpPr.fitToPage = True
xw.print_title_rows = f'{HDR}:{HDR}'
for w_ in (xw, ct, rd, idw):
    protect(w_)

# ordre des onglets : Lisez-moi, Identification, postes, Contrôle, Export IFC, Listes
wb.move_sheet('Listes', offset=len(wb.sheetnames) - 1 - wb.sheetnames.index('Listes'))
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = {'Lisez-moi': GREEN, 'Identification': GREEN, 'Contrôle': RED, 'Export IFC': GREEN, 'Listes': MUTED}.get(ws.title, NAVY)
for ws in wb.worksheets:
    if ws.title in ('Lisez-moi', 'Identification', 'Contrôle', 'Export IFC'):
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
protect(ls)
wb.active = 0
wb.save(OUT)
print('ok', wb.sheetnames)
