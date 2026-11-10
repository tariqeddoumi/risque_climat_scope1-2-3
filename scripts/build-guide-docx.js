/**
 * Guide utilisateur du canevas de collecte de l'empreinte carbone (charte Groupe BCP).
 *
 * Usage : node scripts/build-guide-docx.js <canevas.json> <dossier_captures> <sortie.docx>
 *   canevas.json : produit par scripts/extract-canevas-json.py (une seule source : le canevas)
 *   captures     : produites par scripts/build-guide-captures.py
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, LevelFormat, TableOfContents, Header, Footer, PageNumber, BorderStyle, PageBreak,
  VerticalAlign, ImageRun, TabStopType, TabStopPosition, PageOrientation,
} = require('docx');

const [, , JSON_IN, CAPS, OUT] = process.argv;
const CANEVAS = JSON.parse(fs.readFileSync(JSON_IN, 'utf8'));
const FONT = 'Arial';
const C = { orange: 'E37B05', brun: '462016', text: '3D3D3B', light: 'F2E8DB', light2: 'FBF6F0', bord: 'D9C7B0', muted: '8A7A6A', rose: 'FBE3E1', rouge: 'B3261E' };
const ASSETS = path.join(__dirname, 'assets');
const LOGO = fs.readFileSync(path.join(ASSETS, 'groupe-bcp-logo.png'));
const BAND = fs.readFileSync(path.join(ASSETS, 'groupe-bcp-bandeau.png'));
let W = 9600; // largeur utile courante (portrait 9600, paysage 14538)
let FIGW = 640;
const PORTRAIT = { w: 11906, h: 16838, W: 9600, FIGW: 640 };
const LANDSCAPE = { w: 16838, h: 11906, W: 14538, FIGW: 960 };
const VERSION = 'Version 1.0 — novembre 2026';

// ------------------------------------------------------------------ helpers
function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(s => s.startsWith('**')
    ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...o })
    : new TextRun({ text: s, font: FONT, ...o }));
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), spacing: { after: 120, line: 276 }, alignment: o.align || AlignmentType.JUSTIFIED, keepNext: o.keepNext });
const H = (lvl, t, br = false) => new Paragraph({ heading: lvl, pageBreakBefore: br, children: [new TextRun({ text: t, font: FONT })] });
const H1 = t => H(HeadingLevel.HEADING_1, t, true);
const H2 = t => H(HeadingLevel.HEADING_2, t);
const H3 = t => H(HeadingLevel.HEADING_3, t);
const B = (t, lvl = 0) => new Paragraph({ numbering: { reference: 'bul', level: lvl }, children: runs(t), spacing: { after: 60, line: 264 } });
let numRef = 0;
const newList = () => `n${++numRef}`;
const N = (t, ref) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 70, line: 264 } });
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });
const border = { style: BorderStyle.SINGLE, size: 4, color: C.bord };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, o = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    shading: { fill: o.fill || C.light2, type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: VerticalAlign.TOP,
    children: String(text ?? '').split('\n').map(l => new Paragraph({ children: runs(l, { size: o.size || 18, bold: o.bold, color: o.color || C.text }), spacing: { after: 30, line: 240 } })),
  });
}
function T(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map(r => Math.floor(W * r / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  return [new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, color: 'FFFFFF', fill: C.orange })) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => cell(v, widths[i], { fill: ri % 2 ? C.light2 : C.light, bold: o.firstBold && i === 0, color: o.firstBold && i === 0 ? C.brun : C.text })) })),
    ],
  }), SP()];
}
function box(title, lines, kind = 'info') {
  const k = { info: [C.light, C.orange, C.brun], attention: [C.rose, C.rouge, C.rouge], astuce: [C.light2, C.brun, C.brun] }[kind];
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { fill: k[0], type: ShadingType.CLEAR, color: 'auto' },
    borders: { top: border, bottom: border, right: border, left: { style: BorderStyle.SINGLE, size: 36, color: k[1] } },
    margins: { top: 110, bottom: 110, left: 180, right: 160 },
    children: [new Paragraph({ children: runs(title, { bold: true, size: 20, color: k[2] }), spacing: { after: 70 } }),
      ...lines.map(l => new Paragraph({ children: runs(l, { size: 19 }), spacing: { after: 50, line: 252 } }))],
  })] })] }), SP()];
}
function pngSize(buf) { return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; }
let figNo = 0;
function fig(file, caption, maxW = 640) {
  const data = fs.readFileSync(path.join(CAPS, file + '.png'));
  const { w, h } = pngSize(data);
  const maxH = FIGW === 960 ? 470 : 820;
  let width = Math.min(maxW === 640 ? FIGW : maxW, FIGW);
  if (h * width / w > maxH) width = Math.round(maxH * w / h);
  const height = Math.round(h * width / w);
  figNo += 1;
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 80, after: 60 },
      children: [new ImageRun({ type: 'png', data, transformation: { width, height }, altText: { title: caption, description: caption, name: file } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: `Figure ${figNo} — ${caption}`, font: FONT, italics: true, size: 18, color: C.brun })] }),
  ];
}
const SECTIONS = [{ geo: PORTRAIT, children: [] }];
let content = SECTIONS[0].children;
const add = (...xs) => xs.flat().forEach(x => content.push(x));
function newSection(geo) { W = geo.W; FIGW = geo.FIGW; const sct = { geo, children: [] }; SECTIONS.push(sct); content = sct.children; }

// ------------------------------------------------------------------ page de garde
const centered = (text, run, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing, children: [new TextRun({ text, font: FONT, ...run })] });
add(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 500 },
    children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 230, height: 120 }, altText: { title: 'Groupe BCP', description: 'Logo Groupe BCP', name: 'logo' } })] }),
  centered('GUIDE UTILISATEUR', { bold: true, size: 44, color: C.orange }, { after: 200 }),
  centered('Canevas de collecte des données de l\'empreinte carbone', { size: 30, color: C.brun }, { before: 60, after: 160 }),
  centered('Remplir le canevas pas à pas, avec des copies d\'écran et des exemples pour chaque cas', { italics: true, size: 22, color: C.text }, { after: 300 }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [new ImageRun({ type: 'png', data: BAND, transformation: { width: 620, height: 45 }, altText: { title: 'Bandeau', description: 'Bandeau décoratif Groupe BCP', name: 'bandeau' } })] }),
  centered('Destinataires : contributeurs des sites, correspondants BPR et référents des directions centrales', { size: 20, color: C.text }, { after: 40 }),
  centered(VERSION, { size: 20, color: C.text }, { after: 40 }),
  centered('Classification : usage interne', { size: 20, italics: true, color: C.muted }, { after: 360 }),
  new Paragraph({ indent: { left: 360, right: 360 }, shading: { type: ShadingType.CLEAR, fill: C.light }, spacing: { after: 160 },
    children: runs('Ce guide accompagne le fichier « Canevas_collecte_empreinte_carbone_commente.xlsx ». Les copies d\'écran sont tirées du canevas lui-même ; les valeurs qui y figurent sont des exemples.', { italics: true, size: 19, color: C.brun }) }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: [new TextRun({ text: 'Sommaire', font: FONT, size: 34, bold: true, color: C.orange })], spacing: { after: 220 } }),
  new TableOfContents('Sommaire', { hyperlink: true, headingStyleRange: '1-2' }),
);

// ------------------------------------------------------------------ 1. Présentation
add(H1('1. À quoi sert ce canevas'),
  P('Le Groupe mesure chaque année son **empreinte carbone propre**, c\'est-à-dire les gaz à effet de serre émis par son activité : électricité des agences, carburant des véhicules, climatisation, achats, équipements, déchets, déplacements. Pour cela, il a besoin de chiffres que seuls les sites et les directions détiennent.'),
  P('Le canevas est le fichier Excel qui recueille ces chiffres. **Vous n\'avez aucun calcul d\'émissions à faire** : vous donnez des quantités (kWh, litres, kg, nombre d\'équipements…) dans l\'unité demandée, et l\'équipe Empreinte carbone applique ensuite les facteurs d\'émission.'),
  H2('1.1 Qui remplit quoi'),
  ...T(['Vous êtes', 'Vous remplissez', 'Exemple'], [
    ['Contributeur d\'un site (chef d\'agence, gestionnaire de site, responsable logistique)', 'L\'onglet Identification et les onglets dont les données sont gérées localement', 'Électricité de l\'agence, déchets, gaz butane'],
    ['Correspondant BPR', 'Vous coordonnez les sites de la BPR, vous vérifiez les fichiers et vous les transmettez', 'Vérification de l\'onglet Contrôle de chaque agence'],
    ['Référent d\'une direction centrale', 'Les onglets de votre domaine, pour tout le périmètre que vous gérez', 'Achats (Direction des Achats), Équipements (Support SI), Déplacements (RH)'],
  ], [1.6, 2.2, 1.6], { firstBold: true }),
  H2('1.2 Ce qui change par rapport à l\'ancien canevas'),
  P('L\'ancien canevas mélangeait sur les mêmes lignes les consignes, les exemples et les cellules à remplir. Beaucoup de contributeurs ne savaient pas où écrire, ni comment traiter les cas particuliers. Le nouveau canevas corrige ces difficultés :'),
  ...T(['Avant', 'Maintenant'], [
    ['Exemple et saisie sur la même ligne (« Âge = 3 ans » répété partout)', 'Exemples en gris dans un bloc séparé ; saisie uniquement dans les cellules orangées'],
    ['Unités écrites à la main (« 35555 kw », « 0 kg/an »)', 'Nombre seul dans la cellule ; unité choisie dans une liste'],
    ['Texte libre (« Amesterdam », « MAROC-Dubai abu dhabi »)', 'Listes déroulantes (villes avec code aéroport, familles d\'achat, fluides…)'],
    ['Case vide ou 0 sans explication', 'Statut obligatoire : Valeur, Néant, Non disponible ou Non applicable'],
    ['Aucune vérification avant envoi', 'Contrôle automatique sur chaque ligne et bilan dans l\'onglet Contrôle'],
    ['Un seul exemple, sans les cas particuliers', '4 à 6 cas possibles par onglet, chacun illustré par une ligne d\'exemple'],
  ], [1, 1.3]),
);

// ------------------------------------------------------------------ 2. Avant de commencer
const lPrep = newList();
add(H1('2. Avant de commencer'),
  H2('2.1 Ouvrir le fichier correctement'),
  N('Enregistrez le fichier reçu sur votre poste **avant** de le remplir (ne travaillez pas depuis la pièce jointe du courriel).', lPrep),
  N('À l\'ouverture, si Excel affiche un bandeau jaune « Mode protégé », cliquez sur **Activer la modification**. Sinon les listes déroulantes et les contrôles ne fonctionnent pas.', lPrep),
  N('Renommez le fichier selon la règle : **EC_<exercice>_<code site>_<nom du site>.xlsx**. Exemple : EC_2025_4896_Agence_Al_Boughaz.xlsx.', lPrep),
  N('Ne renommez pas, ne supprimez pas et ne déplacez pas les onglets : l\'onglet Contrôle s\'appuie sur eux.', lPrep),
  ...box('Copier-coller : utilisez « Coller les valeurs »', [
    'Si vous recopiez des chiffres depuis un autre fichier (extraction, tableau de suivi), utilisez **Collage spécial › Valeurs** (raccourci Ctrl+Alt+V puis V). Un collage normal remplace les listes déroulantes et les contrôles de la cellule, et peut coller du texte à la place d\'un nombre.',
  ], 'attention'),
  H2('2.2 Rassembler les documents'),
  P('Gagnez du temps en réunissant les pièces avant d\'ouvrir le canevas. Ces documents serviront aussi de **pièces justificatives** à joindre à votre envoi.'),
  ...T(['Onglet', 'Documents à rassembler', 'Où les trouver'], [
    ['1-Électricité', 'Les 12 factures d\'électricité de l\'exercice, par compteur', 'ONEE, Lydec, Redal, Amendis ou régie ; portail client ; comptabilité'],
    ['2-Carburants et gaz', 'Relevés des cartes carburant ; bons de livraison de gaz ; suivi du groupe électrogène', 'Patrimoine, logistique du site'],
    ['3-Climatisation', 'Liste des climatiseurs ; rapports d\'intervention du mainteneur', 'Maintenance Multitechnique, mainteneur'],
    ['4-Achats', 'Extraction des dépenses HT de l\'exercice par famille', 'ERP Achats, comptabilité'],
    ['5-Équipements', 'Inventaire du matériel avec date d\'installation ; liste des GAB', 'Support SI, PCA, registre des immobilisations'],
    ['6-Déchets', 'Bordereaux du prestataire de collecte ; certificats de destruction', 'Patrimoine, prestataire'],
    ['7-Déplacements pro', 'Extraction de l\'agence de voyage ; notes de frais ; indemnités kilométriques', 'RH, agence de voyage'],
    ['8-Domicile-travail', 'Résultats de l\'enquête mobilité', 'RH'],
    ['9-Transport fonds-courrier', 'États mensuels du transporteur de fonds ; factures postales', 'Achats, Sécurité, Patrimoine'],
  ], [1.2, 2.2, 1.8], { firstBold: true }),
  ...box('Vous n\'avez pas tout ? Ce n\'est pas bloquant', [
    'Remplissez ce que vous avez. Pour une donnée que vous n\'avez pas, choisissez le statut **Non disponible** et indiquez en commentaire qui pourrait la détenir. Pour une activité qui n\'existe pas sur votre site, choisissez **Néant**.',
  ], 'astuce'),
);

// ------------------------------------------------------------------ 3. Découvrir
add(H1('3. Découvrir le canevas'),
  H2('3.1 Les onglets et le code couleur'),
  P('Le classeur compte 13 onglets. Les onglets verts d\'accueil (Lisez-moi, Identification), les neuf onglets de saisie numérotés de 1 à 9, l\'onglet Contrôle (rouge) et l\'onglet Listes (gris). Le bas de l\'onglet Lisez-moi les récapitule, avec le code couleur des cellules :'),
  ...fig('03_lisez_moi_onglets_couleurs', 'Onglet Lisez-moi : liste des onglets et code couleur'),
  ...T(['Couleur', 'Signification', 'Que faire'], [
    ['Orangé clair (texte bleu)', 'Cellule à remplir', 'Saisir ou choisir dans la liste'],
    ['Gris', 'Ligne d\'exemple', 'Lire, ne pas modifier, ne pas recopier'],
    ['Beige', 'Explication, cas possible, ligne d\'aide', 'Lire'],
    ['Rose', 'Erreur à éviter', 'Lire'],
    ['Blanc, texte rouge (dernière colonne)', 'Contrôle automatique', 'Ne pas remplir ; corriger la ligne tant que ce n\'est pas « OK »'],
  ], [1.3, 1.6, 1.8], { firstBold: true }),
  H2('3.2 L\'onglet Lisez-moi'),
  P('C\'est le mode d\'emploi intégré au fichier. Il se lit en deux minutes et résume l\'essentiel de ce guide.'),
  ...fig('01_lisez_moi_etapes_regles', 'Onglet Lisez-moi : les 5 étapes et les 6 règles d\'or'),
  ...T(['Repère', 'Zone', 'Ce qu\'il faut retenir'], [
    ['1', 'Bandeau', 'Le canevas recueille des chiffres : aucun calcul d\'émissions n\'est demandé.'],
    ['2', 'Les 5 étapes', 'S\'identifier, ouvrir chaque onglet, saisir dans les cellules orangées, donner un statut, vérifier puis envoyer.'],
    ['3', 'Les 6 règles d\'or', 'Un nombre seul sans unité ; la quantité avant le montant ; les montants HT ; un statut sur chaque ligne ; ne compter qu\'une fois ; aucun calcul dans les cellules.'],
  ], [0.5, 1.2, 4]),
  ...fig('02_lisez_moi_statuts_methodes', 'Onglet Lisez-moi : choisir le statut et la méthode'),
);

// ------------------------------------------------------------------ 4. Identification
add(H1('4. L\'onglet Identification'),
  P('Il se remplit **une seule fois**, en premier. Il dit à quel site, à quelle entité et à quelle année se rapportent toutes les données du fichier.'),
  ...fig('04_identification_vide', 'Onglet Identification, avant saisie'),
  ...T(['Repère', 'Colonne', 'Rôle'], [
    ['1', 'Information', 'Ce qui est demandé'],
    ['2', 'Votre réponse', 'La seule colonne à remplir (cellules orangées). Certaines cellules ont une liste déroulante.'],
    ['3', 'Aide et exemple', 'Comment répondre, avec un exemple'],
  ], [0.5, 1.2, 4]),
  H2('4.1 Les champs un par un'),
  ...T(['Champ', 'Comment le remplir', 'Piège à éviter'], [
    ['Exercice', 'Année des données, sur 4 chiffres. Exemple : 2025 pour les consommations du 1er janvier au 31 décembre 2025.', 'Mettre l\'année d\'envoi (2026) au lieu de l\'année des données (2025).'],
    ['Entité / BPR', 'Liste déroulante.', 'Écrire un libellé à la main : il sera refusé.'],
    ['Type de site', 'Liste déroulante (agence, succursale, siège régional…).', '—'],
    ['Code du site', 'Code interne de l\'agence ou du site (code entité).', 'Laisser vide : le site ne pourra pas être rapproché du référentiel.'],
    ['Nom du site, Ville', 'En clair.', '—'],
    ['Le fichier couvre', '« Ce site uniquement », ou le périmètre couvert si une direction répond pour plusieurs sites.', 'Regrouper plusieurs agences sans le dire.'],
    ['Surface (m²)', 'Surface de plancher occupée.', 'Donner la surface du terrain.'],
    ['Effectif (ETP)', 'Équivalents temps plein présents au 31/12.', '—'],
    ['Occupation des locaux', 'Propriétaire, locataire ou mis à disposition.', '—'],
    ['Fonction du contributeur', 'Votre fonction, sans nom ni adresse électronique.', '—'],
    ['Date d\'envoi', 'JJ/MM/AAAA.', '—'],
  ], [1.2, 2.6, 1.8], { firstBold: true }),
  ...fig('05_identification_remplie', 'Onglet Identification rempli (exemple de l\'agence Al Boughaz)'),
);

// ------------------------------------------------------------------ 5. Anatomie
newSection(LANDSCAPE);
add(H(HeadingLevel.HEADING_1, '5. Comprendre un onglet de saisie'),
  P('Les neuf onglets de saisie ont **tous la même organisation**. Une fois que vous avez compris l\'onglet 1-Électricité, vous savez remplir tous les autres. Chaque onglet se lit de haut en bas, en cinq zones :'),
  ...fig('06_anatomie_onglet', 'Les cinq zones d\'un onglet de saisie (exemple : 1-Électricité)'),
  ...T(['Repère', 'Zone', 'À quoi elle sert', 'Que faire'], [
    ['1', 'Bandeau', 'Nom du poste et scope ; qui remplit ; où trouver l\'information', 'Vérifier que l\'onglet vous concerne'],
    ['2', 'Comment remplir', 'La règle de saisie de l\'onglet, en quatre phrases', 'Lire avant de saisir'],
    ['3', 'Cas possibles et exemples', 'Les situations les plus fréquentes, chacune avec une ligne d\'exemple grise', 'Trouver le cas qui ressemble au vôtre et imiter sa ligne'],
    ['4', 'Erreurs à éviter', 'Les erreurs constatées lors de la collecte précédente', 'Lire'],
    ['5', 'Votre saisie', 'En-têtes, ligne d\'aide bleue sous chaque colonne, puis 40 lignes à remplir', 'Saisir une ligne par cas ; vérifier la dernière colonne'],
  ], [0.5, 1.1, 2.4, 1.8]),
  H2('5.1 Le bandeau et les consignes'),
  ...fig('07_electricite_bandeau_consignes', 'Bandeau et consignes de l\'onglet 1-Électricité'),
  P('Le bandeau indique **qui remplit** l\'onglet. Si ce n\'est pas vous et que la donnée est déclarée par une direction centrale, vous n\'avez rien à ressaisir : une ligne avec le statut **Non applicable** et un commentaire suffit.'),
  H2('5.2 Les cas possibles'),
  ...fig('08_electricite_cas_possibles', 'Cas possibles de l\'onglet 1-Électricité'),
  P('Chaque cas commence par un bandeau beige qui décrit la situation (repère 1) et dit quoi faire. La ou les lignes grises qui suivent (repères 2 et 3) montrent **exactement** ce qu\'il faut saisir, colonne par colonne. Leur colonne « Contrôle automatique » affiche « OK » : c\'est le résultat attendu sur vos propres lignes.'),
  ...box('Bon réflexe', ['Avant de saisir, cherchez dans les cas possibles celui qui ressemble le plus à votre situation, puis reproduisez sa ligne dans la zone orangée avec vos propres chiffres.'], 'astuce'),
  H2('5.3 Les erreurs à éviter'),
  ...fig('13_electricite_erreurs', 'Erreurs à éviter dans l\'onglet 1-Électricité'),
);

// ------------------------------------------------------------------ 6. Saisir
const lDD = newList(); const lAdd = newList();
add(H1('6. Saisir une ligne pas à pas'),
  H2('6.1 Les colonnes de la zone de saisie'),
  ...fig('09_electricite_zone_saisie', 'Zone de saisie de l\'onglet 1-Électricité : colonnes numérotées et ligne d\'aide'),
  P('Sous chaque en-tête, la **ligne d\'aide** (fond beige, texte en italique) rappelle ce qu\'il faut écrire. Si vous passez la souris sur un en-tête, une bulle d\'aide apparaît aussi. Pour l\'onglet Électricité :'),
  ...T(['Repère', 'Colonne', 'Ce qu\'il faut saisir'], CANEVAS['1-Électricité'].colonnes.map((c, i) => [String(i + 1), c.nom, c.aide]), [0.5, 1.5, 4]),
  P('Les quatre dernières colonnes de saisie (Statut, Méthode, Pièce justificative, Commentaire) et la colonne de contrôle sont **les mêmes dans tous les onglets**. Elles sont détaillées ci-dessous.'),
  H2('6.2 Choisir dans une liste déroulante'),
  ...fig('12_liste_deroulante', 'Choisir une valeur dans une liste déroulante (colonne « Usage du compteur »)'),
  N('Cliquez dans la cellule orangée.', lDD),
  N('Une petite flèche apparaît à droite de la cellule : cliquez dessus.', lDD),
  N('Cliquez sur la valeur voulue. Vous pouvez aussi taper la valeur exacte au clavier.', lDD),
  P('Une valeur qui n\'est pas dans la liste est refusée : Excel affiche le message « **Valeur non autorisée** — Choisissez une valeur dans la liste déroulante ». Cliquez sur **Annuler** puis choisissez dans la liste. Si aucune valeur ne convient, choisissez « Autre (préciser) » quand elle existe et expliquez en commentaire.'),
  H2('6.3 Saisir un nombre'),
  B('Tapez **le nombre seul** : 2850, et non « 2 850 kWh » ni « 2850 kw ».'),
  B('Pour les décimales, utilisez le séparateur de votre Excel (la **virgule** dans un Excel en français : 12,5).'),
  B('Pas d\'espace, pas de symbole monétaire, pas de formule (« =955*113 »).'),
  B('Un texte ou un nombre négatif est refusé (message « **Nombre attendu** »). Une année doit avoir 4 chiffres (message « Saisissez une année sur 4 chiffres »).'),
  H2('6.4 Choisir le statut : la question à se poser'),
  P('Le statut est **obligatoire sur chaque ligne**. Il dit si le chiffre existe et, sinon, pourquoi. Posez-vous les questions dans l\'ordre :'),
  ...T(['Question', 'Si oui', 'Exemple'], [
    ['1. Ai-je un chiffre (quantité ou montant), même égal à 0 ?', '**Valeur**', '2 850 kWh ; recharge de fluide = 0 kg selon le rapport du mainteneur'],
    ['2. L\'activité n\'existe-t-elle pas sur mon site ?', '**Néant (aucune activité)**', 'Pas de groupe électrogène, pas de panneaux solaires'],
    ['3. La donnée est-elle déclarée par quelqu\'un d\'autre, ou la ligne ne concerne-t-elle pas mon type de site ?', '**Non applicable**', 'Gardiennage payé et déclaré par la Direction des Achats'],
    ['4. L\'activité existe, mais je n\'ai pas le chiffre ?', '**Non disponible** + commentaire', 'Électricité incluse dans les charges du bailleur'],
  ], [2.2, 1.3, 2.4]),
  ...box('Attention au zéro', [
    'Avec le statut **Néant**, **Non disponible** ou **Non applicable**, laissez la quantité **vide**. Un 0 avec ces statuts déclenche l\'alerte « Statut incohérent avec le chiffre saisi ».',
    'Le 0 est réservé au statut **Valeur**, quand il est vrai et justifié.',
  ], 'attention'),
  H2('6.5 Choisir la méthode'),
  ...T(['Méthode', 'Quand la choisir', 'Exemple'], [
    ['Mesuré (facture, compteur, pesée)', 'Le chiffre est lu sur un document', 'kWh de la facture'],
    ['Extrait d\'un système (ERP, inventaire)', 'Le chiffre vient d\'une extraction informatique', 'Dépenses HT par famille dans l\'ERP'],
    ['Estimé à partir d\'un montant', 'Vous ne connaissez que le montant', 'Montant HT et prix moyen du litre'],
    ['Extrapolé (ratio, échantillon)', 'Vous avez appliqué un ratio ou une pesée partielle', 'Déchets pesés une semaine'],
    ['Estimation d\'expert', 'Évaluation raisonnée faute de document', 'Année d\'achat estimée du mobilier'],
  ], [1.7, 2, 2]),
  P('La méthode permet à l\'équipe d\'attribuer un **niveau de qualité** à chaque donnée. Une estimation n\'est pas une faute : elle est simplement signalée comme telle.'),
  H2('6.6 Pièce justificative et commentaire'),
  B('**Pièce justificative** : le nom du fichier joint ou la référence du document (« Facture Amendis 01-2025 »). Joignez les pièces à votre envoi.'),
  B('**Commentaire** : tout ce qui aide à comprendre le chiffre : hypothèse (« 48 000 TTC ÷ 1,2 »), écart avec l\'an dernier, période couverte, direction qui détient la donnée.'),
  H2('6.7 Besoin de plus de 40 lignes ?'),
  N('Faites un clic droit sur le numéro d\'une ligne orangée **au milieu** de la zone (par exemple la ligne 20 de la zone), puis **Insérer**.', lAdd),
  N('Sélectionnez la ligne du dessus, copiez-la, puis collez-la sur la nouvelle ligne : les listes et le contrôle sont recopiés. Effacez ensuite les valeurs copiées.', lAdd),
  N('N\'insérez pas de ligne **après** la dernière ligne orangée : elle ne serait pas comptée dans l\'onglet Contrôle.', lAdd),
);

// ------------------------------------------------------------------ 7. Contrôle automatique
add(H1('7. Le contrôle automatique'),
  P('La dernière colonne de chaque onglet, **Contrôle automatique**, vérifie chaque ligne dès que vous la commencez. Elle affiche « OK » quand la ligne est complète et cohérente, sinon un message d\'alerte qui indique quoi corriger.'),
  ...fig('10_electricite_saisie_alertes', 'Lignes saisies avec quatre erreurs signalées par le contrôle automatique'),
  ...T(['Repère', 'Message', 'Ce qui s\'est passé', 'Correction'], [
    ['1', '⚠ Choisir un statut', 'La ligne (mars) est remplie mais le statut est vide', 'Choisir « Valeur » dans la colonne Statut'],
    ['2', '⚠ Nombre attendu, sans unité', 'La quantité a été tapée « 2 480 kWh » : c\'est du texte', 'Retaper 2480'],
    ['3', '⚠ Quantité ou montant manquant', 'Statut « Valeur » sans aucun chiffre (panneaux solaires : seul le nombre de panneaux est en commentaire)', 'Saisir les kWh produits, ou choisir « Néant » si le site n\'a pas de panneaux'],
    ['4', '⚠ Statut incohérent avec le chiffre saisi', 'Un 0 a été saisi avec le statut « Néant »', 'Effacer le 0 (laisser la quantité vide)'],
  ], [0.5, 1.5, 2.2, 1.8]),
  P('Une fois les corrections faites, toutes les lignes affichent « OK » :'),
  ...fig('11_electricite_saisie_corrigee', 'Les mêmes lignes après correction'),
  H2('7.1 Tous les messages possibles'),
  ...T(['Message', 'Cause', 'Correction'], [
    ['OK', 'La ligne est complète et cohérente', 'Rien à faire'],
    ['⚠ Choisir un statut', 'Ligne commencée sans statut', 'Choisir un statut'],
    ['⚠ Quantité manquante / Quantité ou montant manquant', 'Statut « Valeur » mais aucun chiffre', 'Saisir le chiffre, ou changer le statut'],
    ['⚠ Statut incohérent avec le chiffre saisi', 'Un chiffre (même 0) avec Néant, Non disponible ou Non applicable', 'Effacer le chiffre, ou choisir « Valeur »'],
    ['⚠ Nombre attendu, sans unité', 'Texte dans une cellule de quantité (unité, espace, lettre)', 'Retaper le nombre seul'],
    ['⚠ Choisir l\'unité', 'Quantité sans unité (onglets Carburants et gaz, Déchets)', 'Choisir l\'unité dans la liste'],
    ['(vide)', 'Ligne non commencée', 'Rien à faire'],
  ], [1.8, 2, 1.8]),
  ...box('Le contrôle ne juge pas la justesse du chiffre', [
    'Il vérifie la forme : présence du statut, d\'un nombre, d\'une unité. Il ne peut pas savoir si 2 850 kWh est le bon chiffre. Les vérifications de cohérence (écart avec l\'an dernier, comparaison entre sites) sont faites ensuite par l\'équipe Empreinte carbone.',
  ], 'info'),
);

// ------------------------------------------------------------------ 8. Onglet par onglet
const COMMUNES = new Set(['Statut', 'Méthode', 'Pièce justificative', 'Commentaire', 'Contrôle automatique']);
const EXTRA = {
  '1-Électricité': {
    pourquoi: 'L\'électricité achetée au réseau est comptée en scope 2. C\'est souvent l\'un des premiers postes d\'émission d\'une banque : agences, data centers et climatisation en consomment beaucoup.',
    savoir: [
      'Sur la facture, relevez la ligne « consommation » ou « énergie active » en **kWh**. La « puissance souscrite » (en kW ou kVA) n\'est pas une consommation.',
      'Si la facture distingue plusieurs tranches ou heures pleines et heures creuses, **additionnez** les kWh.',
      'Facture bimestrielle : saisissez-la sur le mois de fin de période et indiquez la période en commentaire (« facture janvier-février »).',
      'Un compteur qui alimente aussi un locataire ou la salle informatique est signalé dans « Usage du compteur » : cela évite de compter deux fois la même électricité.',
    ],
  },
  '2-Carburants et gaz': {
    pourquoi: 'Le carburant brûlé dans les véhicules et les groupes électrogènes, le gaz et le bois brûlés sur place émettent directement du CO2 : c\'est le scope 1.',
    savoir: [
      'Une ligne par **usage** et par **combustible** : le gasoil des véhicules et celui du groupe électrogène font deux lignes.',
      'Le relevé des cartes carburant donne directement les litres : c\'est la meilleure source.',
      'Si vous ne connaissez que le montant, donnez aussi le **prix unitaire moyen** : sans lui, les litres ne peuvent pas être recalculés.',
      'Gaz butane : choisissez l\'unité « Bouteille 12 kg » ou « Bouteille 3 kg » et donnez le nombre de bouteilles.',
      'Les véhicules loués avec carburant inclus se déclarent dans 4-Achats (location de véhicules), pas ici.',
    ],
  },
  '3-Climatisation': {
    pourquoi: 'Les climatiseurs contiennent des gaz frigorigènes (R410A, R22…) qui réchauffent le climat des milliers de fois plus que le CO2 quand ils s\'échappent. La quantité **rechargée** par le mainteneur correspond à ce qui a fui.',
    savoir: [
      'La donnée clé est la colonne « Quantité rechargée dans l\'année (kg) ». Elle figure sur le rapport d\'intervention du mainteneur.',
      'La charge nominale (sur la plaque de l\'appareil) est demandée en complément. **Ce n\'est pas une fuite** : ne la reportez pas comme quantité rechargée.',
      'Aucune recharge cette année : saisissez 0 avec le statut « Valeur ». C\'est une information utile.',
      'Plusieurs appareils identiques (même type, fluide et année) : une seule ligne, avec leur nombre.',
      'Fluide illisible : choisissez « Inconnu » et joignez une photo de la plaque.',
    ],
  },
  '4-Achats': {
    pourquoi: 'Fabriquer et livrer ce que la banque achète (papier, toners, repas, prestations…) émet des gaz à effet de serre chez les fournisseurs. Ces émissions sont estimées à partir des dépenses et, pour le papier, des quantités.',
    savoir: [
      'Une ligne par **famille d\'achat**, avec le total HT de l\'année. Inutile de détailler les articles (stylos, gommes…).',
      'Montants **hors taxes**. Si vous n\'avez que le TTC, divisez par 1,2 (TVA à 20 %) et indiquez-le en commentaire ; si le taux de TVA est différent, précisez-le.',
      'Papier : donnez aussi le nombre de rames (ou les tonnes). Restauration : le nombre de repas si vous l\'avez.',
      'Un contrat payé par le siège pour tout le réseau (gardiennage, nettoyage…) est déclaré une seule fois, par la direction qui le paie. Les sites mettent « Non applicable ».',
    ],
  },
  '5-Équipements': {
    pourquoi: 'Fabriquer un ordinateur, un GAB ou un bâtiment émet beaucoup de gaz à effet de serre. Ces émissions sont réparties sur la durée d\'utilisation de l\'équipement : c\'est pourquoi l\'**année d\'achat** est indispensable.',
    savoir: [
      'Regroupez par **catégorie et année d\'achat** : « 35 PC portables achetés en 2022 » fait une ligne.',
      'Si vous ne connaissez que l\'âge : année = exercice − âge (2025 − 3 = 2022).',
      'Pas de tranche (« plus de 4 ans ») : donnez une année, quitte à l\'estimer avec la méthode « Estimation d\'expert ».',
      'Ne comptez que le matériel **en service au 31/12** sur le site. Le matériel sorti ou stocké au central est déclaré par le SI.',
      'Bâtiment : quantité 1, surface en m², année de construction ou de dernière grosse rénovation.',
    ],
  },
  '6-Déchets': {
    pourquoi: 'Les déchets émettent des gaz à effet de serre lors de leur traitement, surtout en décharge. Le recyclage en émet beaucoup moins : c\'est pourquoi le **traitement** est demandé.',
    savoir: [
      'Le poids (kg ou tonnes) figure sur les bordereaux ou les certificats du prestataire.',
      'À défaut, donnez le **nombre de sacs ou de bacs** collectés sur l\'année : l\'équipe fera la conversion en poids.',
      'Une pesée sur une semaine ou un mois est utile : indiquez la période, l\'extrapolation à l\'année est faite par l\'équipe.',
      'Collecte communale sans aucun suivi : « Non disponible ». L\'effectif du site servira à estimer.',
    ],
  },
  '7-Déplacements pro': {
    pourquoi: 'Les voyages professionnels, surtout en avion, pèsent lourd dans l\'empreinte. La distance est calculée par l\'équipe à partir des villes de départ et d\'arrivée.',
    savoir: [
      'Avion : une ligne par trajet, villes choisies dans la liste (avec le code aéroport).',
      'Un **aller-retour = une seule ligne** « Aller-retour ». Ne créez pas une deuxième ligne pour le retour.',
      'Vol avec escale : une ligne par tronçon (Casablanca – Paris, puis Paris – Montréal).',
      'Ville absente de la liste : choisissez « Autre (préciser en commentaire) » et écrivez la ville en commentaire.',
      'Train, voiture, taxi : distance totale en km. Hôtel : nombre de nuitées dans « Nombre de voyages ».',
    ],
  },
  '8-Domicile-travail': {
    pourquoi: 'Les trajets quotidiens des employés entre leur domicile et leur lieu de travail représentent un poste important, que la banque peut réduire (navettes, covoiturage, télétravail).',
    savoir: [
      'Une ligne par **mode de transport**.',
      'La distance demandée est celle d\'un **aller simple** (domicile → bureau) ; le retour est compté par l\'équipe.',
      'Jours de présence : comptez seulement les jours sur site (3 jours si 2 jours de télétravail).',
      'Le total des employés doit correspondre à l\'effectif indiqué dans l\'onglet Identification.',
    ],
  },
  '9-Transport fonds-courrier': {
    pourquoi: 'Le transport de fonds, les navettes de courrier et l\'envoi de courrier aux clients consomment du carburant chez les prestataires.',
    savoir: [
      'Transport de fonds : additionnez les passages des 12 états mensuels du transporteur et donnez la distance aller-retour depuis sa base.',
      'Courrier : nombre d\'envois et poids total si vous les avez ; sinon le montant HT de la facture.',
      'Navette interne : nombre de tournées et distance aller-retour d\'une tournée.',
    ],
  },
};
const CAP = { '1-Électricité': null, '2-Carburants et gaz': '20_carburants_et_gaz', '3-Climatisation': '21_climatisation', '4-Achats': '22_achats',
  '5-Équipements': '23_equipements', '6-Déchets': '24_dechets', '7-Déplacements pro': '25_deplacements_pro', '8-Domicile-travail': '26_domicile-travail',
  '9-Transport fonds-courrier': '27_transport_fonds-courrier' };
const fmt = v => (typeof v === 'number' ? v.toLocaleString('fr-FR') : String(v));
add(H1('8. Les onglets un par un'),
  P('Cette section reprend chaque onglet de saisie : à quoi il sert, qui le remplit, les consignes, chaque colonne avec un exemple, les cas possibles et les pièges. Les colonnes communes (Statut, Méthode, Pièce justificative, Commentaire, Contrôle) sont expliquées en section 6 et ne sont pas répétées.'),
);
Object.entries(CANEVAS).forEach(([sheet, d], idx) => {
  const x = EXTRA[sheet];
  const parts = d.bandeau.split('|').map(s => s.trim());
  add(idx === 0 ? H2(d.titre) : new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: true, children: [new TextRun({ text: d.titre, font: FONT })] }));
  add(...T(['Rubrique', 'Détail'], [
    ['Pourquoi ce poste', x.pourquoi],
    ['Périmètre', parts[0]],
    ['Qui remplit', parts[1].replace(/^Qui remplit : /, '')],
    ['Où trouver l\'information', parts[2].replace(/^Où trouver l'information : /, '')],
  ], [1.3, 4.3], { firstBold: true }));
  add(H3('Comment remplir'));
  d.consignes.forEach(s => add(P(s, { align: AlignmentType.LEFT })));
  if (CAP[sheet]) {
    add(...fig(CAP[sheet], `Consignes, cas possibles et erreurs à éviter de l'onglet ${sheet}`));
  } else {
    add(P('Les copies d\'écran de cet onglet figurent en sections 5 à 7.'));
  }
  // colonnes
  const firstEx = {};
  d.cas.forEach(c => c.exemples.forEach(e => Object.entries(e).forEach(([k, v]) => { if (!(k in firstEx)) firstEx[k] = v; })));
  add(H3('Les colonnes'));
  if (CAP[sheet]) add(...fig(CAP[sheet].slice(0, 2) + 'b_colonnes', `En-têtes et ligne d'aide de la zone de saisie (${sheet})`));
  add(...T(['Colonne', 'Ce qu\'il faut saisir', 'Exemple'], d.colonnes.filter(c => !COMMUNES.has(c.nom)).map(c => [c.nom, c.aide, firstEx[c.nom] !== undefined ? fmt(firstEx[c.nom]) : '—']), [1.5, 2.6, 1.6], { firstBold: true }));
  // cas
  add(H3('Les cas possibles expliqués'));
  d.cas.forEach(c => {
    add(new Paragraph({ children: runs(c.titre, { bold: true, color: C.orange, size: 20 }), spacing: { before: 120, after: 60 }, keepNext: true }));
    c.exemples.forEach((e, i) => {
      const txt = Object.entries(e).filter(([k]) => k !== 'Contrôle automatique').map(([k, v]) => `${k} = ${fmt(v)}`).join(' ; ');
      add(B(`${c.exemples.length > 1 ? `Ligne ${i + 1} : ` : 'À saisir : '}${txt}`));
    });
  });
  add(H3('Bon à savoir'));
  x.savoir.forEach(s => add(B(s)));
  add(...box('Erreurs à éviter', d.erreurs, 'attention'));
});

// ------------------------------------------------------------------ 9. Vérifier et envoyer
const lSend = newList();
newSection(PORTRAIT);
add(H(HeadingLevel.HEADING_1, '9. Vérifier et envoyer le fichier'),
  H2('9.1 L\'onglet Contrôle'),
  P('L\'onglet Contrôle fait le bilan de tout le fichier, sans aucune saisie de votre part. Pour chaque onglet, il compte les lignes saisies, les lignes « OK » et les alertes. En bas, il indique l\'état du fichier et vérifie que l\'onglet Identification est complet.'),
  ...fig('30_controle_alertes', 'Onglet Contrôle : 4 alertes restent à corriger dans l\'onglet Électricité'),
  P('Ici, l\'onglet 1-Électricité compte 6 lignes saisies, dont 2 « OK » et 4 alertes. Retournez dans cet onglet, repérez les lignes dont la dernière colonne n\'affiche pas « OK » et corrigez-les comme expliqué en section 7.'),
  ...fig('31_controle_pret', 'Onglet Contrôle : aucune alerte, le fichier est prêt à être envoyé'),
  H2('9.2 Liste de vérification avant envoi'),
  ...T(['Point à vérifier', 'Où'], [
    ['L\'onglet Identification est complet (« Identification complète : Oui »)', 'Onglet Contrôle'],
    ['L\'état du fichier indique « Prêt à envoyer » (0 alerte)', 'Onglet Contrôle'],
    ['Chaque onglet qui vous concerne a au moins une ligne, même avec le statut Néant ou Non applicable', 'Onglets 1 à 9'],
    ['Les montants sont hors taxes', 'Colonnes de montant'],
    ['Aucune donnée d\'un autre site n\'est incluse sans être signalée', 'Commentaires'],
    ['Un écart de plus de 20 % avec l\'an dernier est expliqué', 'Commentaires'],
    ['Les pièces justificatives citées sont jointes', 'Colonne Pièce justificative'],
  ], [3.5, 1.5]),
  H2('9.3 Envoyer'),
  N('Vérifiez le nom du fichier : **EC_<exercice>_<code site>_<nom du site>.xlsx**.', lSend),
  N('Regroupez les pièces justificatives dans un dossier compressé portant le même nom (EC_2025_4896_Agence_Al_Boughaz.zip).', lSend),
  N('Envoyez le fichier et le dossier à votre **correspondant BPR** (sites des BPR) ou à l\'**équipe Empreinte carbone** (directions centrales), avant la date indiquée dans le courriel de lancement.', lSend),
  N('Restez disponible : l\'équipe peut revenir vers vous pour une précision. Ses questions porteront le nom de l\'onglet et le numéro de ligne.', lSend),
);

// ------------------------------------------------------------------ 10. FAQ
add(H1('10. Questions fréquentes'),
  ...[
    ['Je n\'ai pas l\'information demandée. Que dois-je faire ?', 'Choisissez le statut « Non disponible » et indiquez en commentaire qui pourrait la détenir (« factures chez la Direction du Patrimoine »). Ne laissez pas la ligne vide et ne mettez pas 0.'],
    ['L\'activité n\'existe pas chez nous (pas de groupe électrogène, pas de véhicule).', 'Une ligne avec le statut « Néant (aucune activité) », sans quantité. C\'est une information à part entière.'],
    ['La donnée est déjà fournie par le siège. Dois-je la ressaisir ?', 'Non. Mettez « Non applicable » et précisez en commentaire la direction qui la déclare. Sinon, elle serait comptée deux fois.'],
    ['Je n\'ai que le montant, pas la quantité.', 'Saisissez le montant HT, laissez la quantité vide et choisissez la méthode « Estimé à partir d\'un montant ». Pour les carburants, ajoutez le prix unitaire moyen.'],
    ['Je n\'ai que des montants TTC.', 'Divisez par 1,2 (TVA à 20 %) et écrivez en commentaire « montant TTC ÷ 1,2 ». Si le taux de TVA est différent, précisez-le.'],
    ['La valeur que je cherche n\'est pas dans la liste déroulante.', 'Choisissez « Autre (préciser) » ou « Autre (préciser en commentaire) » quand elle existe, et précisez en commentaire. Sinon, prenez la valeur la plus proche et expliquez.'],
    ['Excel refuse ma saisie avec un message d\'erreur.', 'Le message indique la règle : choisir dans la liste, saisir un nombre positif sans unité, ou une année sur 4 chiffres. Cliquez sur « Annuler » et corrigez.'],
    ['Puis-je modifier les lignes grises d\'exemple ?', 'Non. Elles servent de modèle. Saisissez uniquement dans les cellules orangées de la zone « Votre saisie ».'],
    ['Le contrôle affiche « OK » : mon chiffre est-il juste ?', 'Le contrôle vérifie la forme de la ligne, pas l\'exactitude du chiffre. Relisez vos chiffres et joignez les pièces.'],
    ['Je réponds pour plusieurs agences. Un fichier ou plusieurs ?', 'De préférence un fichier par site. Si vous répondez pour plusieurs sites dans un même fichier, indiquez-le dans l\'onglet Identification et précisez le code du site en commentaire de chaque ligne.'],
    ['Ma facture d\'électricité est bimestrielle.', 'Saisissez chaque facture sur le mois de fin de période et indiquez la période couverte en commentaire.'],
    ['À qui poser une question ?', 'À votre correspondant BPR, ou à l\'équipe Empreinte carbone (Direction des Risques), en précisant l\'onglet et le numéro de ligne.'],
  ].flatMap(([q, a]) => [new Paragraph({ children: runs(q, { bold: true, color: C.brun, size: 21 }), spacing: { before: 160, after: 60 }, keepNext: true }), P(a)]),
);

// ------------------------------------------------------------------ Annexe
add(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun({ text: 'Annexe. Glossaire', font: FONT })] }),
  ...T(['Terme', 'Définition'], [
    ['Empreinte carbone', 'Total des gaz à effet de serre émis par l\'activité de la banque sur une année, exprimé en tonnes d\'équivalent CO2'],
    ['Scope 1', 'Émissions directes : carburant, gaz et fluides frigorigènes des sites et véhicules de la banque'],
    ['Scope 2', 'Émissions liées à l\'électricité achetée'],
    ['Scope 3', 'Autres émissions indirectes : achats, équipements, déchets, déplacements, transport'],
    ['kWh', 'Kilowattheure : unité d\'énergie consommée, lue sur la facture d\'électricité'],
    ['kW', 'Kilowatt : unité de puissance (puissance souscrite). Ce n\'est pas une consommation'],
    ['HT / TTC', 'Hors taxes / toutes taxes comprises. Les montants sont demandés HT'],
    ['Fluide frigorigène', 'Gaz contenu dans les climatiseurs (R410A, R22, R32…)'],
    ['Charge nominale', 'Quantité de fluide contenue dans l\'appareil, indiquée sur sa plaque'],
    ['Recharge', 'Quantité de fluide ajoutée par le mainteneur pour compenser une fuite'],
    ['DEEE', 'Déchets d\'équipements électriques et électroniques'],
    ['ETP', 'Équivalent temps plein'],
    ['Code IATA', 'Code à trois lettres d\'un aéroport (CMN pour Casablanca Mohammed V)'],
    ['Pièce justificative', 'Document qui prouve le chiffre : facture, extraction, bordereau, rapport d\'intervention'],
  ], [1.3, 4], { firstBold: true }),
);

// ------------------------------------------------------------------ document
const numbering = { config: [
  { reference: 'bul', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } }, run: { color: C.orange } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 260 } } } }] },
  ...Array.from({ length: numRef }, (_, i) => ({ reference: `n${i + 1}`, levels: [
    { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } }, run: { bold: true, color: C.orange } } }] })),
] };
const hs = (id, name, size, color, before, after, lvl) => ({ id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true,
  run: { font: FONT, size, bold: true, color }, paragraph: { spacing: { before, after }, keepNext: true, outlineLevel: lvl } });
const doc = new Document({
  creator: 'Groupe BCP — Direction des Risques',
  title: 'Guide utilisateur du canevas de collecte de l\'empreinte carbone',
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: C.text } } },
    paragraphStyles: [hs('Heading1', 'Heading 1', 34, C.orange, 360, 220, 0), hs('Heading2', 'Heading 2', 26, C.brun, 280, 140, 1), hs('Heading3', 'Heading 3', 22, C.orange, 200, 100, 2)],
  },
  numbering,
  features: { updateFields: true },
  sections: SECTIONS.map((sct, i) => ({
    properties: { page: { size: { width: Math.min(sct.geo.w, sct.geo.h), height: Math.max(sct.geo.w, sct.geo.h), orientation: sct.geo.w > sct.geo.h ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
      margin: { top: 1300, bottom: 1134, left: 1150, right: 1150, header: 600, footer: 500 } }, titlePage: i === 0 },
    headers: {
      default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: sct.geo.W }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.orange, space: 4 } },
        children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 69, height: 36 }, altText: { title: 'Groupe BCP', description: 'Logo', name: 'logo-entete' } }),
          new TextRun({ text: '\tGuide utilisateur — canevas de collecte de l\'empreinte carbone · Version 1.0', size: 16, color: C.brun, font: FONT })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: sct.geo.W }],
        children: [new TextRun({ text: 'Groupe BCP — document interne', size: 16, color: C.muted, font: FONT }),
          new TextRun({ text: '\tPage ', size: 16, color: C.muted, font: FONT }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.muted, font: FONT }),
          new TextRun({ text: ' / ', size: 16, color: C.muted, font: FONT }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.muted, font: FONT })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: sct.children,
  })),
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('written', OUT, (b.length / 1048576).toFixed(1), 'Mo', figNo, 'figures'); });
