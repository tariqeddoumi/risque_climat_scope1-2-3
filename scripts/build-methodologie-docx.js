const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, LevelFormat, TableOfContents, Header, Footer,
  PageNumber, BorderStyle, PageBreak, VerticalAlign, ImageRun, TabStopType, TabStopPosition,
} = require('docx');
const path = require('path');

const OUT = process.argv[2];
const FONT = 'Arial';
// Charte Groupe BCP : mêmes couleurs, police et logo que les documents de l'outil de notation
const C = { navy: '462016', green: 'E37B05', light: 'F2E8DB', grey: 'FBF6F0', bord: 'D9C7B0', muted: '8A7A6A', text: '3D3D3B' };
const ASSETS = path.join(__dirname, 'assets');
const LOGO = fs.readFileSync(path.join(ASSETS, 'groupe-bcp-logo.png'));
const BAND = fs.readFileSync(path.join(ASSETS, 'groupe-bcp-bandeau.png'));
const W = 9600; // largeur utile A4 (marges de la charte)

// ---------- helpers ----------
function runs(text, opts = {}) {
  // **gras** inline
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map(s => s.startsWith('**')
    ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...opts })
    : new TextRun({ text: s, font: FONT, ...opts }));
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), spacing: { after: 120, line: 276 }, alignment: o.align || AlignmentType.JUSTIFIED, ...(o.p || {}) });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })], pageBreakBefore: true });
const H1n = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const H3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })] });
const B = (t, lvl = 0) => new Paragraph({ numbering: { reference: 'bul', level: lvl }, children: runs(t), spacing: { after: 60, line: 264 } });
const N = (t, ref = 'num1') => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 60, line: 264 } });
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });

const border = { style: BorderStyle.SINGLE, size: 4, color: C.bord };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, o = {}) {
  const paras = String(text).split('\n').map(line => new Paragraph({
    children: runs(line, { size: o.size || 18, bold: o.bold, color: o.color }),
    spacing: { after: 40, line: 240 }, alignment: o.align || AlignmentType.LEFT,
  }));
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 }, verticalAlign: VerticalAlign.TOP,
    children: paras,
  });
}
function T(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map(r => Math.floor(W * r / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  const head = new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, color: 'FFFFFF', fill: C.green, size: 18 })) });
  const body = rows.map((r, ri) => new TableRow({ cantSplit: true,
    children: r.map((v, i) => cell(v, widths[i], { fill: (o.firstColFill && i === 0) ? C.light : (ri % 2 ? C.grey : C.light), bold: o.firstColBold && i === 0, color: (o.firstColBold && i === 0) ? C.navy : C.text })),
  }));
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] }), SP()];
}
function box(title, lines, fill = C.light) {
  const children = [new Paragraph({ children: runs(title, { bold: true, size: 20, color: C.navy }), spacing: { after: 80 } }),
    ...lines.map(l => new Paragraph({ children: runs(l, { size: 19 }), spacing: { after: 50, line: 252 } }))];
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: 'auto' },
    borders: { top: { style: BorderStyle.SINGLE, size: 12, color: C.green }, bottom: border, left: border, right: border },
    margins: { top: 120, bottom: 120, left: 160, right: 160 }, children })] })] }), SP()];
}

// ---------- contenu ----------
const content = [];
const add = (...xs) => xs.flat().forEach(x => content.push(x));

// Page de garde (charte Groupe BCP)
const centered = (text, run, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing, children: [new TextRun({ text, font: FONT, ...run })] });
add(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 500 },
    children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 230, height: 120 }, altText: { title: 'Groupe BCP', description: 'Logo Groupe BCP', name: 'logo' } })] }),
  centered('EMPREINTE CARBONE DE LA BANQUE', { bold: true, size: 44, color: C.green }, { after: 200 }),
  centered('Scopes 1, 2 et 3 : méthodologie, règles d\'alimentation, guide de collecte et cahier des charges de l\'outil web', { size: 28, color: C.navy }, { before: 60, after: 160 }),
  centered('Mesure de l\'empreinte carbone propre du Groupe : de la collecte Excel à un outil de collecte et de suivi', { italics: true, size: 22, color: C.text }, { after: 300 }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [new ImageRun({ type: 'png', data: BAND, transformation: { width: 620, height: 45 }, altText: { title: 'Bandeau', description: 'Bandeau décoratif Groupe BCP', name: 'bandeau' } })] }),
  centered('Destinataires : Risques climatiques, RSE, Logistique et Patrimoine, SI, Achats, RH, BPR', { size: 20, color: C.text }, { after: 40 }),
  centered('Version 1.1 — novembre 2026', { size: 20, color: C.text }, { after: 40 }),
  centered('Classification : usage interne', { size: 20, italics: true, color: C.muted }, { after: 360 }),
  new Paragraph({ indent: { left: 360, right: 360 }, shading: { type: ShadingType.CLEAR, fill: C.light }, spacing: { after: 160 },
    children: runs('Base documentaire : canevas IFC « Carbon Footprint Data Request » (versions 231015, 231027, 231108), Canevas EC consolidé V1, fichiers de suivi de collecte, Tableau Carbon 2024 et 16 fichiers de données brutes. Exercice de référence : 2022.', { italics: true, size: 19, color: C.navy }) }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: [new TextRun({ text: 'Sommaire', font: FONT, size: 34, bold: true, color: C.green })], spacing: { after: 220 } }),
  new TableOfContents('Sommaire', { hyperlink: true, headingStyleRange: '1-2' }),
);

// 1. Synthèse
add(H1('1. Synthèse'),
  P('La banque a mesuré une première fois son empreinte carbone propre, sur l\'exercice 2022, avec l\'appui de l\'IFC. La collecte s\'est faite par des canevas Excel envoyés aux directions centrales du siège et aux Banques Populaires Régionales (BPR). Elle a été consolidée à la main dans un « Canevas EC consolidé » et suivie dans plusieurs tableaux d\'avancement.'),
  P('L\'analyse des fichiers montre que la démarche est solide sur le fond : référentiel GHG Protocol, couverture des trois scopes, matrice des responsables par poste. En revanche, les **règles d\'alimentation** des canevas produisent des données hétérogènes, difficiles à contrôler et à recalculer d\'une année sur l\'autre. Les principaux constats :'),
  B('**Unités et natures mélangées** : montants en dirhams saisis à la place de quantités physiques (propane, gasoil), « kw » au lieu de « kWh », valeurs saisies en texte (« 35555 kw », « 0 kg/an », « 1,00 »).'),
  B('**Risques de double comptage** : chaleur et froid demandés « en consommation électrique des appareils », donc déjà inclus dans la facture d\'électricité ; salle informatique incluse dans l\'électricité du siège et dans celle du data center ; gasoil d\'un centre d\'estivage additionné au siège.'),
  B('**Fluides frigorigènes** : charge totale des climatiseurs reportée (46,4 kg de R22, 60 kg de R410A) au lieu des quantités rechargées, ce qui surestime fortement les fuites.'),
  B('**Données non normalisées** : destinations aériennes en texte libre (« Amesterdam » et « Amsterdam », « DAKAR » et « DAKKAR »), allers et retours comptés séparément, tranches d\'âge ouvertes (« Age > 4 ans »).'),
  B('**Sources contradictoires** non arbitrées : deux montants de gardiennage pour les centres d\'estivage, deux nombres de vols pour les BPR Oujda et Centre Sud.'),
  B('**Extrapolations implicites** : montant d\'une agence multiplié par 113 directement dans une cellule, sans hypothèse documentée.'),
  B('**Suivi peu objectivé** : un taux d\'avancement saisi à la main, un statut binaire (reçu, non reçu), des formats de suivi différents selon les BPR, ni échéance ni historique de relance.'),
  P('Ce document propose :'),
  N('une description de la **méthodologie actuelle** (section 2) ;'),
  N('un **diagnostic** des règles d\'alimentation et leur version corrigée (sections 3 et 4) ;'),
  N('un **guide de collecte** utilisable tel quel par les contributeurs (section 5) ;'),
  N('la **méthode de calcul cible** (section 6) ;'),
  N('le **cahier des charges de l\'outil web** de collecte et de suivi (section 7). L\'outil serait construit sur le même socle technique que les outils existants (SGES, outil de notation interne des entreprises).'),
  ...box('Recommandation principale', [
    'Remplacer le canevas Excel unique par une collecte **par poste et par site**, dans un outil web, avec des **unités imposées**, une **pièce justificative** par donnée, un **score de qualité** et un **circuit de validation**. Les facteurs d\'émission seraient versionnés et appliqués par un moteur de calcul déterministe.',
    'L\'Excel reste possible, mais uniquement comme **gabarit d\'import** contrôlé par l\'outil, et non plus comme support de calcul.',
  ]),
);

// 2. Méthodologie actuelle
add(H1('2. Méthodologie actuelle'),
  H2('2.1 Cadre et objectif'),
  P('Selon l\'onglet « Instructions » des canevas, l\'IFC a été mandatée par la banque dans le cadre du renforcement de son dispositif de gestion des risques climatiques. L\'objectif était de mesurer l\'empreinte carbone de l\'institution sur un périmètre défini, à partir d\'une demande de données structurée.'),
  P('La méthode suit le **GHG Protocol Corporate Standard**. Les émissions y sont classées en trois scopes :'),
  B('**Scope 1** : émissions directes (combustibles des sites et des véhicules, fuites de fluides frigorigènes) ;'),
  B('**Scope 2** : émissions indirectes liées à l\'énergie achetée (électricité, chaleur, froid) ;'),
  B('**Scope 3** : autres émissions indirectes (achats, immobilisations, déchets, transports, déplacements).'),
  P('Le calcul repose sur le principe : **émissions (tCO2e) = donnée d\'activité × facteur d\'émission**. Les facteurs d\'émission ne figurent pas dans les canevas ; ils ont été appliqués par l\'IFC lors du calcul.'),
  H2('2.2 Périmètre organisationnel'),
  P('Le périmètre couvre la Banque Centrale Populaire (siège et réseau) et les BPR. Les canevas sont découpés en onglets, un par type de site :'),
  ...T(['Type de site', 'Onglet du canevas', 'Sites cités dans les fichiers reçus'], [
    ['Siège central', 'Siège central n', 'Zerktouni, EPA, Avignon, Diouri, rue d\'Oran, CNRD 1 et 2, Centre de formation'],
    ['Data center', 'Data center', 'Diouri, Berrechid, centre de repli'],
    ['Sièges régionaux (BPR)', 'Siège régional n', 'Marrakech-Béni Mellal, Rabat, Oujda, Centre Sud, Tanger-Tétouan, Laâyoune, Fès-Meknès, Nador-Al Hoceima'],
    ['Succursales', 'Succursale n', 'Par BPR (par exemple Khouribga, Fkih Ben Salah, Kelâa, Béni Mellal, Guéliz, Safi, Tétouan ville)'],
    ['Agences', 'Agence n', 'Réseau BCP Casablanca et agences des BPR (par exemple Al Boughaz)'],
    ['Centres d\'affaires', 'Centres d\'affaires', 'Casablanca et BPR'],
    ['Centres d\'estivage', 'Centre d\'estivage n', 'Centres d\'estivage, NADI, CNF'],
  ], [1.2, 1.2, 3]),
  P('**Point d\'attention** : l\'approche de consolidation n\'est écrite nulle part. Il faut choisir entre contrôle opérationnel et contrôle financier, puis préciser le traitement des filiales et des locaux loués. Le traitement des sites hors exploitation (locaux et terrains, qui apparaissent dans les dépenses de gardiennage) n\'est pas défini non plus.'),
  H2('2.3 Périmètre opérationnel : postes collectés'),
  P('Le tableau suivant reprend les postes du canevas « Sources données ». Il indique pour chacun la donnée demandée et la direction responsable désignée.'),
  ...T(['Scope', 'Poste', 'Donnée demandée dans le canevas', 'Direction responsable'], [
    ['1', 'Sources fixes', 'Butane, propane, bois, gasoil des groupes électrogènes (kg, litres ou dirhams)', 'Gestion du Patrimoine'],
    ['1', 'Sources mobiles', 'Gasoil et essence du parc automobile (litres ou dirhams)', 'Gestion du Patrimoine'],
    ['1', 'Fugitives', 'Nature du gaz, quantité (kg), âge des climatiseurs des bureaux et des véhicules', 'Maintenance Multitechnique ; Gestion du Patrimoine'],
    ['2', 'Électricité', 'kWh du réseau (12 factures), photovoltaïque, véhicules électriques', 'Maintenance Multitechnique ; Gestion du Patrimoine'],
    ['2', 'Chaleur et froid', 'Consommation électrique des appareils chauffants et refroidissants', 'Maintenance Multitechnique'],
    ['3', 'Achats', 'Petites fournitures, papier, repas, café, abonnements internet, impressions, maintenance, gardiennage, location de véhicules, matériaux', 'Achats, Marchés & Budgets ; Développement social ; Télécom & Sécurité SI ; Communication ; Maintenance ; Infrastructures surveillance'],
    ['3', 'Immobilisations', 'Serveurs, switchs, routeurs, firewalls, onduleurs, PC, écrans, téléphones, imprimantes, scanners, photocopieurs, mobilier, bâtiments (m²), GAB : nombre et âge', 'Data Center ; Support et Performance SI ; Télécom ; Patrimoine ; PCA (GAB)'],
    ['3', 'Déchets', 'kg par jour par flux (ménagers, papier, carton, verre, métaux, plastique, DEEE) et devenir', 'Gestion du Patrimoine ; Data Center ; Maintenance'],
    ['3', 'Transport amont', 'Fret matières premières, transport de fonds (km, type de véhicule, consommation)', 'Achats, Marchés & Budgets'],
    ['3', 'Déplacements professionnels', 'Vols (trajet, nombre), frais de déplacement Maroc et international, flotte, véhicules personnels', 'Pilotage RH et Rémunération ; Patrimoine'],
    ['3', 'Domicile-travail', 'Nombre d\'employés par mode et distance moyenne aller-retour', 'RH'],
    ['3', 'Visiteurs et clients', 'Nombre de personnes et km par mode', 'Achats (à estimer)'],
    ['3', 'Transport aval', 'Courriers clientèle et courriers internes (tonnes, km)', 'Recouvrement, Affaires litigieuses, Marché Particuliers, Patrimoine, Achats'],
  ], [0.55, 1.1, 2.5, 1.85]),
  H2('2.4 Dispositif de collecte'),
  N('**Envoi des canevas** : un classeur Excel (« IFC_BCP CR_Carbon Footprint Data Request », plusieurs versions entre octobre et novembre 2023) adressé aux référents. Il comprend un onglet par type de site, et chaque ligne réunit le poste, la source, la description, un exemple et les cellules de saisie.', 'num2'),
  N('**Affectation** : l\'onglet « Sources données » associe chaque poste à une fonction et à un responsable nommé.', 'num2'),
  N('**Retours** : seize fichiers ont été reçus, dans des formats variables : canevas complétés, extractions brutes (fournitures, inventaire informatique, âge des GAB, état des passages Brinks, déplacements par BPR, gardiennage).', 'num2'),
  N('**Consolidation** : les retours ont été recopiés à la main dans le « Canevas EC consolidé V1 », qui compte 17 onglets, du siège central aux CA du Centre Sud. Le « Tableau Carbon 2024 » reprend la même structure pour la BPR Tanger-Tétouan.', 'num2'),
  N('**Suivi** : tableaux d\'avancement par feuille et par BPR, avec le poste reçu ou non reçu, un taux d\'avancement, le responsable, le traitement réalisé et un commentaire.', 'num2'),
  H2('2.5 État d\'avancement observé'),
  P('D\'après le dernier fichier de suivi, la collecte est restée partielle :'),
  ...T(['Périmètre', 'Avancement déclaré', 'Principaux manques'], [
    ['Siège central Casablanca', '40 %', 'Sources mobiles, climatisation des véhicules, véhicules électriques'],
    ['Agences et succursales Casablanca', '10 à 35 %', 'Scopes 1 et 2 des succursales, une partie des achats'],
    ['Data center', '50 %', 'Âge des climatiseurs, déchets, transport'],
    ['Centres d\'estivage', '80 %', 'Fluides, déchets, matériaux ; écart sur le gardiennage'],
    ['BPR Oujda, Centre Sud', '85 à 95 %', 'Transports amont et aval ; écarts sur les vols'],
    ['BPR Marrakech-BM, Rabat, Tanger-Tétouan', '30 à 80 %', 'Déchets, transports, flotte, visiteurs'],
    ['BPR Laâyoune, Fès-Meknès, Nador-Al Hoceima', '0 %', 'Aucune donnée reçue'],
  ], [1.6, 1, 2.6]),
);

// 3. Diagnostic
add(H1('3. Diagnostic des règles d\'alimentation'),
  P('Chaque anomalie ci-dessous a été relevée dans les fichiers reçus ou consolidés. La dernière colonne renvoie à la règle corrigée de la section 4.'),
  H2('3.1 Anomalies de structure des canevas'),
  ...T(['N°', 'Anomalie', 'Exemple observé', 'Conséquence', 'Règle'], [
    ['A1', 'En-têtes calculés par formule renvoyant à d\'autres cellules', 'Colonne « =D4 », « =H13 », « =I14 » en tête de tableau', 'Libellés faux dès qu\'une ligne est insérée', 'R-01'],
    ['A2', 'Libellés incrémentés par recopie', '« kilométrage en 2022, 2023, 2024, 2025 » pour les véhicules électriques 1 à 4', 'Années de référence erronées', 'R-01'],
    ['A3', 'Année figée', '« Données de l\'année 2022 » dans le Tableau Carbon 2024', 'Données affectées au mauvais exercice', 'R-02'],
    ['A4', 'Description, exemple et saisie sur la même ligne', 'Exemple « Âge = 3 ans » répété sur des centaines de lignes', 'Confusion entre exemple et réponse ; fichiers lourds', 'R-01'],
    ['A5', 'Colonnes différentes d\'un onglet à l\'autre', 'Scope 2 : « Nature de l\'agent frigorigène » comme en-tête de la colonne électricité', 'Consolidation manuelle source d\'erreurs', 'R-01'],
    ['A6', 'Calculs en dur dans les cellules', '« =(225+714)/1000 », « =I32*113 », « =80+70 »', 'Hypothèses invisibles, non auditables', 'R-09'],
  ], [0.35, 1.5, 1.8, 1.5, 0.45]),
  H2('3.2 Anomalies de contenu'),
  ...T(['N°', 'Anomalie', 'Exemple observé', 'Conséquence', 'Règle'], [
    ['B1', 'Montant saisi à la place d\'une quantité physique', 'Propane : 54 781 Dhs (prix 17,92 Dh/kg en commentaire) ; gasoil : 29 960 Dhs', 'Conversion manuelle nécessaire, prix non tracé', 'R-04'],
    ['B2', 'Unité saisie dans la valeur ou unité fausse', '« 35555 kw », « 0 kg/an », « 4 ans », « 619 panneaux » au lieu des kWh produits', 'Valeur non numérique ; confusion entre kW (puissance) et kWh (énergie)', 'R-03'],
    ['B3', 'Séparateur décimal en texte', '« 1,00 », « 1 283,80 » importés comme texte', 'Totaux faux ou nuls', 'R-03'],
    ['B4', 'Zéro, vide, « Non applicable » et « information disponible auprès de… » utilisés indifféremment', 'Agence Al Boughaz : 0 partout en scope 1', 'Impossible de distinguer « pas d\'émission » et « donnée manquante »', 'R-05'],
    ['B5', 'Charge de fluide au lieu de la recharge', 'Siège : R22 46,4 kg, R410 60 kg, R407 45,2 kg', 'Fuites surestimées (la charge totale ne fuit pas chaque année)', 'R-12'],
    ['B6', 'Chaleur et froid mesurés en électricité des appareils', 'Consigne « Reporter la consommation électrique des appareils chauffants »', 'Double comptage avec l\'électricité du réseau', 'R-11'],
    ['B7', 'Sous-compteurs non isolés', 'Électricité du siège incluant la tour CFC et la salle informatique', 'Double comptage possible avec le data center ou un tiers', 'R-07'],
    ['B8', 'Montant TTC, HT ou comptabilisé selon les fichiers', 'Fournitures : trois colonnes HT, TTC et comptabilisé ; le consolidé reprend le TTC', 'Ratios monétaires biaisés d\'environ 20 %', 'R-04'],
    ['B9', 'Achats ligne à ligne sans famille', '230 références de fournitures (stylos, toners, chemises)', 'Pas de facteur d\'émission par référence ; volume ingérable', 'R-14'],
    ['B10', 'Destinations aériennes en texte libre', '« Amesterdam » et « Amsterdam », « DAKAR » et « DAKKAR », « ABU DABI » et « abou dabi »', 'Distances incalculables sans retraitement', 'R-16'],
    ['B11', 'Aller simple et aller-retour confondus', '« CASA_OUJDA » 26 et « OUJDA_CASA » 1, alors que la consigne demandait des AR', 'Distances doublées ou divisées par deux', 'R-16'],
    ['B12', 'Âges en tranches ouvertes', '« Age > 4 ans », « Antérieure à 2018 »', 'Amortissement impossible à appliquer', 'R-13'],
    ['B13', 'Domicile-travail sans nombre de jours', '« Nombre = 40 et distance moyenne = 10 km »', 'Calcul annuel impossible (jours travaillés, télétravail)', 'R-17'],
    ['B14', 'Déchets en kg par jour', 'Consigne « Quantité journalière en kg »', 'Donnée rarement disponible, extrapolation arbitraire', 'R-15'],
    ['B15', 'Bouteille de butane à 13 kg', 'Exemple « 10 bouteilles de 13 kg »', 'Les bouteilles au Maroc font 12 kg ou 3 kg', 'R-10'],
  ], [0.35, 1.5, 1.8, 1.5, 0.45]),
  H2('3.3 Anomalies de processus'),
  ...T(['N°', 'Anomalie', 'Exemple observé', 'Conséquence', 'Règle'], [
    ['C1', 'Sources contradictoires non arbitrées', 'Gardiennage des centres d\'estivage : environ 3,7 MDh selon le Développement social et 2,16 MDh selon la Surveillance ; vols Oujda : 39 ou 29', 'Résultat dépendant du fichier retenu', 'R-06'],
    ['C2', 'Agrégation de sites à la consolidation', 'Gasoil du centre d\'estivage ajouté au siège Centre Sud', 'Ventilation par site perdue', 'R-07'],
    ['C3', 'Extrapolation non documentée', 'Fournitures d\'une agence × 113 agences', 'Biais non mesuré', 'R-09'],
    ['C4', 'Avancement déclaratif', 'Taux de 0,05 à 0,95 saisis à la main', 'Avancement non vérifiable', 'R-19'],
    ['C5', 'Formats de suivi hétérogènes', 'Le suivi BPRTT a 3 colonnes, celui de Rabat 4, celui du BPCS 2', 'Pas de vision consolidée fiable', 'R-19'],
    ['C6', 'Pas de pièce justificative ni de score de qualité', 'Aucune facture ni extraction jointe à la donnée', 'Audit et vérification externe impossibles', 'R-08'],
    ['C7', 'Données personnelles dans les canevas', 'Courriels nominatifs dans « Sources données »', 'Diffusion large de données personnelles', 'R-20'],
  ], [0.35, 1.5, 1.8, 1.5, 0.45]),
);

// 4. Règles corrigées
add(H1('4. Règles d\'alimentation corrigées'),
  H2('4.1 Règles générales'),
  ...T(['Code', 'Règle', 'Contrôle automatique dans l\'outil'], [
    ['R-01', 'Un formulaire par poste et par site. L\'exemple et la consigne s\'affichent en aide contextuelle et ne figurent jamais dans les lignes de saisie. Libellés fixes, sans formule.', 'Formulaire généré depuis le référentiel des postes'],
    ['R-02', 'Chaque donnée est rattachée à une campagne (exercice N) et à une période (année, ou mois si la donnée vient de factures).', 'Exercice hérité de la campagne, non modifiable'],
    ['R-03', 'Valeur numérique pure, unité choisie dans une liste fermée par poste. Séparateur décimal géré par l\'outil.', 'Refus du texte ; conversion vers l\'unité de calcul'],
    ['R-04', 'La quantité physique (kWh, litres, kg, km) prime sur le montant. À défaut : montant **HT** en MAD, avec le prix unitaire moyen et sa source pour la conversion. Le montant reste un dernier recours.', 'Montant seul : qualité « estimée » ; prix unitaire obligatoire'],
    ['R-05', 'Statut explicite pour chaque ligne : « Valeur », « Néant (pas d\'activité) », « Non disponible », « Non applicable ». Le zéro n\'est admis qu\'avec le statut « Valeur » et une justification.', 'Pas de soumission avec des lignes sans statut'],
    ['R-06', 'Un seul propriétaire de la donnée par poste et par périmètre (source de vérité). Une autre source sert au contrôle et non à l\'alimentation.', 'Écart de plus de 10 % entre deux sources : alerte et arbitrage tracé'],
    ['R-07', 'Saisie au niveau du site, identifié par un code dans le référentiel. Les sous-compteurs (salle informatique, tour louée) sont déclarés séparément et affectés à un seul site.', 'Somme des sous-compteurs inférieure ou égale au compteur principal'],
    ['R-08', 'Pièce justificative jointe (facture, extraction ERP, inventaire, bon de livraison) ou référence documentaire. Score de qualité de 1 à 5 (voir 6.4).', 'Qualité supérieure à 2 exigée sans pièce'],
    ['R-09', 'Aucune extrapolation dans la saisie. Les extrapolations (par m², par ETP, par agence type) sont paramétrées dans l\'outil, avec l\'hypothèse et l\'échantillon.', 'Données extrapolées signalées dans les restitutions'],
    ['R-20', 'Les responsables sont rattachés à une fonction et à un compte utilisateur. Les canevas ne contiennent plus de courriels nominatifs.', 'Gestion des droits par rôle'],
  ], [0.5, 3.3, 1.7]),
  H2('4.2 Règles par poste'),
  ...T(['Code', 'Poste', 'Donnée attendue (unité imposée)', 'Source de preuve', 'Méthode de repli'], [
    ['R-10', 'Combustibles fixes', 'Butane et propane en kg (bouteilles de 12 kg ou 3 kg converties automatiquement) ; gasoil des groupes électrogènes en litres ; bois en kg', 'Bons de livraison, factures du fournisseur de gaz, cartes carburant', 'Montant HT ÷ prix moyen annuel'],
    ['R-11', 'Électricité, chaleur et froid', 'Électricité du réseau en kWh par compteur (12 factures ONEE, Lydec, Redal, Amendis, régies) ; photovoltaïque en kWh produits et autoconsommés ; chaleur et froid **achetés à un réseau** seulement, sinon « Non applicable »', 'Factures ou portail du distributeur, compteur de l\'onduleur PV', 'Ratio kWh/m² par type de site, calculé sur les sites complets'],
    ['R-12', 'Fluides frigorigènes', 'Par équipement : type de fluide (liste : R22, R410A, R407C, R32, R134a…), charge nominale (kg), **quantité rechargée dans l\'année (kg)**, date de mise en service', 'Fiches d\'intervention du mainteneur, registre des fluides', 'Charge × taux de fuite annuel par type d\'équipement'],
    ['R-13', 'Immobilisations', 'Par catégorie et par site : quantité et **année d\'acquisition** (ou date d\'installation), ou extraction de l\'inventaire (code matériel, catégorie, date). Bâtiments : m² SHON, année de construction ou de rénovation lourde, structure', 'Inventaire informatique (GLPI ou équivalent), registre des immobilisations, fichier GAB', 'Âge médian de la catégorie'],
    ['R-14', 'Achats de biens et services', 'Montant HT par famille d\'achat (papier, consommables d\'impression, fournitures, restauration, maintenance, gardiennage, nettoyage, télécoms, impressions, location de véhicules, travaux). Papier également en tonnes (une rame A4 de 80 g pèse 2,5 kg)', 'Extraction ERP Achats ou comptabilité par compte ou famille', 'Dépense par ETP du site type'],
    ['R-15', 'Déchets', 'Tonnes par an et par flux, avec le traitement (enfouissement, recyclage, incinération, valorisation)', 'Bordereaux du prestataire, pesées, contrats de collecte', 'Ratio kg/ETP/an par flux, campagne de pesée sur un échantillon'],
    ['R-16', 'Déplacements professionnels', 'Par voyage : origine et destination (codes IATA à partir d\'une liste), aller simple ou aller-retour, classe, nombre de voyageurs. Train, voiture et taxi en km. Hôtel en nuitées', 'Extraction de l\'agence de voyage, notes de frais', 'Montant HT ÷ prix moyen par km'],
    ['R-17', 'Domicile-travail', 'Enquête mobilité : mode, distance aller simple, jours de présence par semaine, télétravail. Extrapolation à l\'effectif ETP', 'Enquête RH annuelle (taux de réponse indiqué)', 'Distance médiane par ville × jours ouvrés'],
    ['R-18', 'Transport de marchandises et de fonds', 'Transport de fonds : nombre de passages par site × distance aller-retour depuis la base du transporteur, type de véhicule. Courrier : tonnes et km, ou nombre de plis par catégorie', 'États mensuels de passages du transporteur, facturation postale', 'Montant HT × facteur monétaire'],
  ], [0.45, 0.95, 2.2, 1.35, 1.05]),
  ...box('Règle de priorité des méthodes (hiérarchie de la donnée)', [
    '1. Donnée mesurée et justifiée (compteur, facture, pesée) : qualité 5.',
    '2. Donnée physique tirée d\'un système (ERP, inventaire, outil de voyage) : qualité 4.',
    '3. Donnée physique estimée à partir d\'un montant et d\'un prix unitaire documenté : qualité 3.',
    '4. Donnée extrapolée par ratio (m², ETP, site type) : qualité 2.',
    '5. Montant converti par un facteur monétaire, ou estimation d\'expert : qualité 1.',
  ]),
);

// 5. Guide de collecte
add(H1('5. Guide de collecte à l\'usage des contributeurs'),
  P('Cette section peut être diffusée seule. Elle explique à chaque contributeur ce qu\'il doit fournir, où trouver l\'information et comment la saisir, d\'abord avec le gabarit Excel transitoire, puis dans l\'outil web.'),
  H2('5.1 Rôles'),
  ...T(['Rôle', 'Qui', 'Responsabilités'], [
    ['Pilote de la campagne', 'Direction Risques climatiques et RSE', 'Ouvre la campagne, fixe le périmètre et le calendrier, arbitre les écarts, valide le résultat'],
    ['Référent central', 'Une personne par direction propriétaire (Patrimoine, Maintenance, Achats, SI, RH, Data Center, PCA, Communication, Développement social, Surveillance)', 'Fournit les données groupe de son domaine, contrôle les données des sites'],
    ['Correspondant BPR', 'Un ou deux par BPR', 'Coordonne la collecte de son siège régional, de ses succursales, agences et CA ; valide en premier niveau'],
    ['Contributeur site', 'Gestionnaire de site, chef d\'agence ou de succursale', 'Saisit les données propres au site lorsqu\'elles ne sont pas centralisées'],
    ['Calculateur', 'Équipe empreinte carbone (avec l\'appui d\'un consultant si besoin)', 'Gère les facteurs d\'émission, lance les calculs, produit les restitutions'],
    ['Contrôleur', 'Contrôle permanent ou audit interne', 'Revue par échantillonnage, piste d\'audit'],
  ], [1.1, 1.9, 2.6]),
  H2('5.2 Calendrier type d\'une campagne annuelle'),
  ...T(['Étape', 'Période (exercice N)', 'Livrable'], [
    ['Préparation : mise à jour du référentiel des sites, des responsables et des facteurs d\'émission', 'Décembre N', 'Campagne paramétrée'],
    ['Lancement : réunion, envoi des accès et du guide', 'Mi-janvier N+1', 'Courriel de lancement'],
    ['Collecte des données centrales (énergie, achats, RH, SI)', 'Janvier à février N+1', 'Postes centraux soumis'],
    ['Collecte des données des sites et des BPR', 'Janvier à mars N+1', 'Postes des sites soumis'],
    ['Relances automatiques à J-7, J et J+7, puis escalade au directeur', 'Continu', 'Journal des relances'],
    ['Contrôles de cohérence et arbitrages', 'Mars N+1', 'Registre des arbitrages'],
    ['Calcul, revue et validation', 'Avril N+1', 'Bilan validé'],
    ['Publication : rapport annuel, reporting BAM, ESRS E1, plan d\'actions', 'Mai à juin N+1', 'Rapport'],
  ], [3, 1.3, 1.5]),
  H2('5.3 Fiches de collecte par poste'),
  P('Chaque fiche indique la donnée à fournir, l\'endroit où la trouver, la façon de la saisir et les erreurs à éviter.'),
);
const fiches = [
  ['Électricité du réseau', 'Maintenance Multitechnique ou gestionnaire de site', [
    '**À fournir** : kWh consommés par compteur, pour chacun des 12 mois de l\'exercice.',
    '**Où trouver** : factures ONEE, Lydec, Redal, Amendis ou régie ; portail client du distributeur ; fichier de suivi des factures de la comptabilité.',
    '**Comment saisir** : une ligne par compteur et par mois, en kWh, avec la facture jointe. Préciser si le compteur alimente une salle informatique ou un tiers (sous-location).',
    '**À éviter** : saisir des kW (puissance souscrite) ou le montant TTC seul ; additionner deux sites sur un même compteur sans le signaler.',
  ]],
  ['Photovoltaïque', 'Maintenance Multitechnique', [
    '**À fournir** : puissance installée (kWc), kWh produits et kWh autoconsommés sur l\'année.',
    '**Où trouver** : supervision de l\'onduleur, rapport du mainteneur.',
    '**À éviter** : donner un nombre de panneaux à la place des kWh.',
  ]],
  ['Carburants du parc et des groupes électrogènes', 'Gestion du Patrimoine', [
    '**À fournir** : litres de gasoil et d\'essence par véhicule ou par carte carburant ; litres pour les groupes électrogènes.',
    '**Où trouver** : relevés des cartes carburant (Afriquia, Total, Shell…), bons d\'achat, suivi des groupes électrogènes.',
    '**Comment saisir** : en litres. Si seul le montant est connu, saisir le montant HT et le prix moyen du litre sur l\'exercice : l\'outil fait la conversion.',
    '**À éviter** : regrouper carburant et péages ou entretien.',
  ]],
  ['Gaz butane et propane', 'Patrimoine ou gestionnaire de site', [
    '**À fournir** : nombre de bouteilles par format (12 kg ou 3 kg) ou kg livrés (citerne propane).',
    '**Où trouver** : bons de livraison, factures.',
    '**À éviter** : le format de 13 kg, qui n\'existe pas au Maroc.',
  ]],
  ['Climatisation et fluides frigorigènes', 'Maintenance Multitechnique', [
    '**À fournir** : liste des équipements (type, fluide, charge nominale, date de mise en service) et quantités **rechargées** dans l\'année.',
    '**Où trouver** : contrat et rapports d\'intervention du mainteneur, registre des fluides.',
    '**À éviter** : déclarer la charge totale comme une fuite.',
  ]],
  ['Achats de biens et services', 'Achats, Marchés & Budgets', [
    '**À fournir** : extraction des dépenses **HT** de l\'exercice par famille d\'achat et par entité ou site, plus les tonnes de papier.',
    '**Où trouver** : ERP Achats ou comptabilité générale, par compte de charge.',
    '**Comment saisir** : importer l\'extraction dans le gabarit « Achats ». La correspondance entre comptes et familles est paramétrée une fois pour toutes.',
    '**À éviter** : le détail article par article (stylo, gomme…) ; les montants TTC.',
  ]],
  ['Immobilisations (informatique, GAB, mobilier, bâtiments)', 'Support et Performance SI, Data Center, Télécom, PCA, Patrimoine', [
    '**À fournir** : extraction de l\'inventaire avec catégorie, modèle, site, date d\'installation ou d\'achat et statut (en service ou sorti).',
    '**Où trouver** : inventaire informatique (fichier « Informatique siège » : code matériel, date d\'installation, statut) ; parc GAB (référence et âge) ; registre des immobilisations ; patrimoine immobilier (m², année de construction).',
    '**À éviter** : les tranches ouvertes (« plus de 4 ans ») : donner l\'année exacte.',
  ]],
  ['Déchets', 'Patrimoine, Data Center (DEEE)', [
    '**À fournir** : tonnes par flux et traitement.',
    '**Où trouver** : bordereaux et factures du prestataire de collecte, certificats de destruction des archives et des DEEE.',
    '**À défaut** : participer à une campagne de pesée sur deux semaines dans les sites de l\'échantillon.',
  ]],
  ['Déplacements professionnels', 'Pilotage RH et Rémunération', [
    '**À fournir** : extraction de l\'agence de voyage (origine, destination, AR ou aller simple, classe) ; nuitées d\'hôtel ; km remboursés en véhicule personnel ; frais de déplacement HT.',
    '**Comment saisir** : choisir les villes dans la liste (codes IATA). L\'outil calcule les distances orthodromiques.',
    '**À éviter** : le texte libre (« MAROC-Dubai abu dhabi »).',
  ]],
  ['Domicile-travail', 'RH', [
    '**À fournir** : résultats de l\'enquête mobilité annuelle (mode principal, distance aller, jours de présence, télétravail) et effectif ETP par site.',
    '**Comment** : questionnaire intégré à l\'outil, anonyme et agrégé par site.',
  ]],
  ['Transport de fonds et courrier', 'Achats, Sécurité, Patrimoine', [
    '**À fournir** : états mensuels de passages du transporteur de fonds par site (exemple : état Brinks par agence) ; tonnage ou nombre de plis du courrier par catégorie.',
    '**Comment** : importer l\'état du transporteur. L\'outil calcule les distances depuis la base logistique.',
  ]],
];
fiches.forEach(([t, who, lines]) => add(H3(t), P(`**Propriétaire** : ${who}`), ...lines.map(l => B(l))));
add(
  H2('5.4 Liste de contrôle avant soumission'),
  B('Toutes les lignes ont un statut (Valeur, Néant, Non disponible, Non applicable).'),
  B('Les valeurs sont des nombres purs, dans l\'unité proposée.'),
  B('Les quantités physiques sont fournies quand elles existent ; à défaut, le montant HT et le prix unitaire.'),
  B('Une pièce justificative est jointe ou référencée.'),
  B('L\'écart avec l\'exercice précédent est expliqué s\'il dépasse 20 %.'),
  B('Aucune donnée d\'un autre site n\'est incluse sans être signalée.'),
  H2('5.5 Gabarit Excel transitoire'),
  P('En attendant l\'outil web, le gabarit Excel doit respecter le format suivant. Il deviendra ensuite le format d\'import de l\'outil.'),
  ...T(['Colonne', 'Contenu', 'Format'], [
    ['code_site', 'Code du site dans le référentiel', 'Liste déroulante'],
    ['exercice', 'Année de l\'exercice', 'AAAA'],
    ['periode', 'Mois ou année', 'AAAA-MM ou AAAA'],
    ['code_poste', 'Code du poste (par exemple S2-ELEC-RESEAU)', 'Liste déroulante'],
    ['sous_categorie', 'Fluide, carburant, famille d\'achat, flux de déchet…', 'Liste dépendante du poste'],
    ['statut', 'Valeur, Néant, Non disponible, Non applicable', 'Liste'],
    ['quantite', 'Valeur numérique', 'Nombre (pas de texte)'],
    ['unite', 'Unité autorisée pour le poste', 'Liste dépendante'],
    ['montant_ht_mad', 'Montant HT si la quantité est inconnue', 'Nombre'],
    ['prix_unitaire / source_prix', 'Pour la conversion d\'un montant', 'Nombre et texte'],
    ['methode', 'Mesuré, Système, Estimé, Extrapolé, Monétaire', 'Liste'],
    ['piece_ref', 'Référence ou nom de fichier de la pièce', 'Texte'],
    ['commentaire', 'Hypothèses, explication des écarts', 'Texte'],
  ], [1.2, 2.6, 1.6]),
  P('Le gabarit est protégé : seules les colonnes de saisie sont modifiables, et les listes et validations de données sont actives. Il ne contient aucune formule de calcul.'),
);

// 6. Méthode de calcul cible
add(H1('6. Méthode de calcul cible'),
  H2('6.1 Principes'),
  B('**Référentiels** : GHG Protocol Corporate Standard et Scope 3 Standard, ISO 14064-1. Les sorties sont compatibles avec l\'ESRS E1 (CSRD) et avec les attentes de Bank Al-Maghrib sur les risques climatiques.'),
  B('**Consolidation** : contrôle opérationnel (proposé), à valider par le comité.'),
  B('**Gaz** : CO2, CH4, N2O et HFC, exprimés en tCO2e avec les PRG à 100 ans du GIEC (AR6). La version des PRG est enregistrée avec chaque calcul.'),
  B('**Scope 2** : publication en approche « location-based » (facteur du réseau marocain de l\'exercice) et « market-based » (contrats, garanties d\'origine, photovoltaïque).'),
  B('**Année de référence** : 2022. Elle est recalculée si un changement de périmètre ou de méthode modifie le total de plus de 5 %.'),
  H2('6.2 Formules par poste'),
  ...T(['Poste', 'Formule'], [
    ['Combustibles', 'Quantité (L ou kg) × FE combustion (kgCO2e/unité). Amont des combustibles en scope 3, catégorie 3'],
    ['Fluides', 'Recharge (kg) × PRG du fluide. Repli : charge × taux de fuite annuel × PRG'],
    ['Électricité', 'kWh × FE du réseau (location-based) ; kWh × FE contractuel (market-based)'],
    ['Achats', 'Quantité physique × FE (papier en t) ou montant HT × FE monétaire (kgCO2e/MAD, corrigé de l\'inflation et du change)'],
    ['Immobilisations', 'Quantité × FE de fabrication ÷ durée d\'amortissement, si l\'âge est inférieur à cette durée (par exemple 4 ans pour un PC, 7 à 10 ans pour un GAB, 50 ans pour un bâtiment)'],
    ['Déchets', 'Tonnes par flux × FE du traitement'],
    ['Avion', 'Distance orthodromique × facteur de détour × FE par tranche de distance et par classe, avec ou sans forçage radiatif (option publiée)'],
    ['Domicile-travail', 'Répondants extrapolés à l\'ETP × distance AR × jours de présence × FE du mode'],
    ['Transport de fonds', 'Passages × distance AR × FE du véhicule blindé (par km)'],
  ], [1.2, 4]),
  H2('6.3 Facteurs d\'émission'),
  P('Les facteurs sont stockés dans une bibliothèque versionnée. Chaque facteur porte sa source, son année, son unité, son périmètre géographique et sa date de validité. Sources proposées, par ordre de préférence :'),
  N('facteurs nationaux (facteur du réseau électrique marocain publié par l\'ONEE ou le ministère de la Transition énergétique, ou à défaut celui de l\'AIE) ;', 'num3'),
  N('Base Empreinte de l\'ADEME, et facteurs du DESNZ (Royaume-Uni) pour les transports ;', 'num3'),
  N('GIEC (PRG AR6) ;', 'num3'),
  N('facteurs monétaires de type EEIO, à défaut seulement.', 'num3'),
  P('Une fois un calcul validé, le facteur utilisé n\'est plus modifié pour ce calcul. Toute nouvelle version donne lieu à un recalcul tracé.'),
  H2('6.4 Qualité et incertitude'),
  P('Chaque donnée reçoit un score de qualité de 1 à 5 (voir la hiérarchie en 4.2). Une incertitude est associée à chaque niveau, par exemple ±5 % pour une donnée mesurée et ±50 % pour une donnée monétaire. Le bilan publie le taux de couverture par donnée mesurée et l\'incertitude globale, obtenue par propagation quadratique.'),
  H2('6.5 Extrapolation au réseau'),
  P('Lorsque tous les sites ne répondent pas, l\'extrapolation se fait par **site type** (agence, succursale, CA). On applique la moyenne des sites complets, en kWh/m², en MAD/ETP ou en kg/ETP, à la surface ou à l\'effectif des sites manquants. L\'outil affiche l\'échantillon, la dispersion et la part extrapolée du total. Une extrapolation ne doit pas dépasser 30 % d\'un poste matériel sans validation du pilote.'),
);

// 7. Cahier des charges
add(H1('7. Cahier des charges de l\'outil web de collecte et de suivi'),
  H2('7.1 Vision'),
  P('L\'outil remplace les canevas Excel. Il industrialise une collecte annuelle, multi-sites et multi-entités : chaque donnée y est saisie une seule fois, contrôlée, justifiée, validée, calculée et restituée, avec une piste d\'audit complète. Il suit les mêmes principes que l\'outil SGES et que l\'outil de notation interne des entreprises : moteur déterministe et explicable, paramétrage à l\'écran par versions datées, workflow de validation, cloisonnement par entité, interface de programmation et assistance de l\'IA toujours soumise à une validation humaine.'),
  H2('7.2 Socle technique (identique aux outils existants)'),
  ...T(['Couche', 'Choix', 'Justification'], [
    ['Interface et serveur', 'Next.js 15 (App Router, actions serveur), React 19, TypeScript', 'Même socle que l\'outil de notation : compétences et composants réutilisés'],
    ['Style', 'Tailwind CSS 4', 'Charte commune'],
    ['Validation', 'zod (schémas partagés entre écrans, API et import)', 'Une seule définition des règles R-01 à R-18'],
    ['Données', 'Prisma 6 et PostgreSQL (Supabase, schéma dédié « carbon », jamais « public »)', 'Multi-dialecte, migrations, renforcement SQL déjà en place'],
    ['Fichiers', 'Stockage objet (Supabase Storage ou S3 interne) pour les pièces justificatives', 'Pièces chiffrées, liens signés'],
    ['IA', 'Fournisseur paramétrable (SDK Anthropic déjà intégré)', 'Lecture des factures et des extractions, propositions jamais appliquées sans accord'],
    ['Tests', 'Vitest (moteur de calcul, règles, imports)', 'Non-régression des calculs'],
    ['Déploiement', 'Docker Compose (sur site) ou Vercel et Supabase', 'Neutre vis-à-vis du cloud, comme les outils existants'],
    ['API', 'REST v1 documentée en OpenAPI, événements sortants signés', 'Connecteurs ERP, RH et énergie ; alimentation de l\'outil de risque climatique'],
  ], [1.1, 2.3, 2.2]),
  H2('7.3 Modules fonctionnels'),
);
const modules = [
  ['M1. Référentiels', [
    '**Organisation** : arbre groupe › banque › BPR › type de site › site, avec code, adresse, ville, région, m², ETP, statut (en exploitation ou non), propriété ou location et date d\'ouverture. Import et export Excel.',
    '**Postes d\'émission** : scope, catégorie GHG, poste, sous-catégories, unités autorisées, méthodes admises, propriétaire par défaut, aide contextuelle et exemple. Versionné.',
    '**Unités et conversions** : table de conversion (bouteille de 12 kg, rame, litre vers kg, m³ de gasoil…).',
    '**Facteurs d\'émission** : bibliothèque versionnée (source, année, unité, zone, gaz, PRG) et import des bases ADEME et DESNZ.',
    '**Listes** : aéroports (IATA, coordonnées), fluides et PRG, familles d\'achat et correspondance avec les comptes, flux et traitements de déchets, modes de transport.',
  ]],
  ['M2. Campagnes de collecte', [
    'Création d\'une campagne : exercice, périmètre, postes actifs, échéances par poste et par niveau.',
    'Affectation automatique des tâches (poste × site ou périmètre) aux responsables selon la matrice, avec délégation et suppléance.',
    'Reprise des données N-1 en préremplissage, pour comparaison.',
    'Clôture et gel de la campagne ; réouverture tracée.',
  ]],
  ['M3. Saisie et import', [
    'Formulaires générés depuis le référentiel des postes, avec aide contextuelle, unités en liste et statut obligatoire (R-05).',
    'Saisie mensuelle ou annuelle selon le poste ; saisie en masse sous forme de grille pour les inventaires.',
    'Import des gabarits Excel et CSV avec rapport d\'anomalies ligne à ligne. Mapping enregistré par source (ERP Achats, inventaire SI, agence de voyage, états Brinks, parc GAB).',
    'Pièces justificatives : glisser-déposer, rattachement à une ou plusieurs lignes, prévisualisation.',
    'Assistance de l\'IA : lecture d\'une facture d\'électricité ou d\'un bordereau, proposition de valeurs avec la source citée, acceptation explicite par le contributeur.',
    'Enquête mobilité domicile-travail intégrée (lien anonyme, agrégation par site).',
  ]],
  ['M4. Contrôles qualité automatiques', [
    'Contrôles bloquants : type, unité, statut, pièce exigée selon la méthode, période dans l\'exercice.',
    'Contrôles d\'alerte : écart N-1 supérieur à 20 %, ratio kWh/m² ou MAD/ETP hors de la plage du site type, valeur aberrante (écart interquartile), somme des sous-compteurs supérieure au compteur principal, mois manquants.',
    'Rapprochement de sources (R-06) : comparaison automatique de deux sources, avec registre des arbitrages motivés.',
    'Score de qualité calculé par ligne, par poste, par site et pour le bilan.',
  ]],
  ['M5. Workflow de validation', [
    'Statuts : brouillon › soumis › validé N1 (correspondant BPR ou référent) › validé N2 (référent central ou pilote) › verrouillé. Rejet avec motif obligatoire, puis reprise.',
    'Circuit paramétrable par poste et par type d\'entité ; séparation des fonctions (un saisisseur ne valide pas sa propre saisie).',
    'Commentaires et fils de discussion par ligne ; historique complet.',
  ]],
  ['M6. Suivi de collecte (tableau de bord)', [
    'Avancement **calculé** et non déclaré : part des tâches validées, pondérée par le poids du poste dans le bilan N-1.',
    'Vue matricielle site × poste (couleurs : non commencé, en cours, soumis, validé, en retard) ; filtres par BPR, type de site, scope et responsable.',
    'Relances automatiques par courriel à J-7, J et J+7, escalade au N+1, journal des relances.',
    'Liste « À traiter » par utilisateur ; export du suivi.',
  ]],
  ['M7. Moteur de calcul', [
    'Calcul déterministe : quantité convertie dans l\'unité du facteur × facteur, avec ventilation par gaz.',
    'Chaque résultat garde la donnée source, la conversion, le facteur (et sa version), la méthode et la qualité.',
    'Extrapolations paramétrées (site type, ratio, échantillon) et signalées.',
    'Simulation : changement de facteur ou d\'hypothèse sans écraser le calcul officiel ; recalcul de l\'année de référence.',
  ]],
  ['M8. Restitution et reporting', [
    'Tableaux de bord : total en tCO2e par scope, catégorie, poste, BPR et site ; évolution par rapport à N-1 et à l\'année de référence ; intensités (tCO2e par ETP, par m², par million de MAD de PNB).',
    'Analyse de Pareto des postes, carte des sites, qualité et incertitude.',
    'Exports : rapport GHG Protocol, tableaux ESRS E1-6, reporting Bank Al-Maghrib, format IFC, Excel détaillé et PDF.',
    'Trajectoire et plan d\'actions : actions de réduction par poste (porteur, échéance, gain estimé), suivi du réalisé.',
  ]],
  ['M9. Administration et sécurité', [
    'Comptes rattachés à une entité ; rôles : administrateur, pilote, référent, correspondant BPR, contributeur, lecteur, auditeur.',
    'Cloisonnement des données par entité dans la couche de données (même mécanisme que l\'outil de notation).',
    'Journal d\'audit immuable de chaque création, modification, validation et calcul.',
    'Authentification par l\'annuaire de la banque (SSO ou LDAP) ou par comptes locaux avec mot de passe fort ; sessions expirantes.',
  ]],
];
modules.forEach(([t, lines]) => add(H3(t), ...lines.map(l => B(l))));
add(
  H2('7.4 Modèle de données principal'),
  ...T(['Table', 'Contenu'], [
    ['Entity', 'Arbre organisationnel (groupe, banque, BPR) et habillage'],
    ['Site', 'Code, type, entité, m², ETP, statut, géolocalisation'],
    ['EmissionPost', 'Scope, catégorie GHG, unités, méthodes, aide (versionné)'],
    ['EmissionFactor', 'Valeur, unité, source, année, zone, gaz, PRG, validité (versionné)'],
    ['Campaign', 'Exercice, périmètre, échéances, statut'],
    ['Task', 'Campagne × poste × site ou périmètre, responsable, échéance, statut'],
    ['ActivityData', 'Tâche, période, sous-catégorie, statut, quantité, unité, montant HT, prix unitaire, méthode, qualité, commentaire'],
    ['Evidence', 'Pièce justificative (fichier, empreinte de hachage, lignes liées)'],
    ['Validation', 'Étape, décision, motif, auteur, date'],
    ['Arbitration', 'Sources comparées, écart, décision, motif'],
    ['CalculationRun', 'Version du moteur et des facteurs, résultats par ligne et agrégats, statut (simulation ou officiel)'],
    ['Action', 'Plan de réduction : poste, porteur, échéance, gain estimé et gain réalisé'],
    ['AuditLog', 'Journal de toutes les opérations'],
  ], [1.3, 4]),
  H2('7.5 Écrans principaux'),
  N('Accueil : mes tâches, échéances, avancement de mon périmètre.', 'num4'),
  N('Campagne : paramétrage, affectations, matrice de suivi site × poste.', 'num4'),
  N('Fiche de saisie d\'un poste : formulaire, pièces, contrôles, commentaires, historique.', 'num4'),
  N('Import : choix du gabarit, mapping, rapport d\'anomalies, intégration.', 'num4'),
  N('Validation : file des éléments à valider, comparaison avec N-1, décision.', 'num4'),
  N('Résultats : tableaux de bord, exports, simulation.', 'num4'),
  N('Référentiels : organisation, postes, facteurs, listes, utilisateurs.', 'num4'),
  H2('7.6 Exigences non fonctionnelles'),
  ...T(['Domaine', 'Exigence'], [
    ['Volumétrie', 'Environ 1 500 sites, 60 postes, 12 périodes : quelques millions de lignes sur 10 ans'],
    ['Performance', 'Écrans affichés en moins de 2 s ; import de 50 000 lignes en moins d\'une minute ; calcul complet en moins de 5 minutes'],
    ['Disponibilité', 'Heures ouvrées, sauvegarde quotidienne, restauration testée'],
    ['Sécurité', 'Chiffrement en transit et au repos, cloisonnement, journal d\'audit, conformité à la loi 09-08 (protection des données personnelles)'],
    ['Traçabilité', 'Tout chiffre publié est relié à sa donnée source, à son facteur et à sa pièce'],
    ['Ergonomie', 'Interface en français, usage possible sur tablette, aide contextuelle sur chaque champ'],
    ['Réversibilité', 'Export complet de la base (Excel, CSV, SQL), code source propriété de la banque'],
  ], [1.2, 4]),
  H2('7.7 Feuille de route proposée'),
  ...T(['Lot', 'Contenu', 'Durée indicative'], [
    ['Lot 0 : cadrage', 'Validation de la méthode, du périmètre, de la matrice des responsables et des facteurs ; reprise des données 2022', '3 à 4 semaines'],
    ['Lot 1 : socle et collecte', 'M1 Référentiels, M2 Campagnes, M3 Saisie et import, M5 Workflow, M9 Administration', '8 à 10 semaines'],
    ['Lot 2 : contrôle et calcul', 'M4 Contrôles qualité, M7 Moteur de calcul, reprise et recalcul 2022, collecte pilote sur une BPR', '6 à 8 semaines'],
    ['Lot 3 : suivi et reporting', 'M6 Tableau de bord, M8 Restitutions et exports (ESRS E1, BAM, IFC), plan d\'actions', '6 semaines'],
    ['Lot 4 : connecteurs et IA', 'Connecteurs ERP Achats, RH, inventaire SI et agence de voyage ; lecture des factures par l\'IA', '6 semaines'],
    ['Extension', 'Scope 3 catégorie 15 (émissions financées, méthode PCAF), en lien avec l\'outil de notation et l\'outil de risque climatique', 'À cadrer'],
  ], [1.3, 3.2, 1]),
  ...box('Critères de réussite', [
    'Au moins 90 % des tâches validées à la date de clôture de la campagne.',
    'Au moins 70 % des émissions couvertes par des données de qualité 4 ou 5.',
    'Aucune donnée publiée sans source, facteur et validation.',
    'Recalcul complet d\'un exercice possible en un clic, à résultat identique.',
  ]),
);

// Annexes
add(H1('Annexe A. Codification proposée des postes'),
  ...T(['Code', 'Scope / catégorie GHG', 'Poste', 'Unités autorisées'], [
    ['S1-COMB-FIXE', '1', 'Combustibles fixes (butane, propane, gasoil des groupes électrogènes, bois)', 'kg, L, bouteille de 12 kg ou 3 kg'],
    ['S1-COMB-MOB', '1', 'Carburants du parc', 'L'],
    ['S1-FUGI', '1', 'Fluides frigorigènes (bâtiments, véhicules)', 'kg'],
    ['S2-ELEC-RESEAU', '2', 'Électricité du réseau', 'kWh'],
    ['S2-ELEC-PV', '2 (information)', 'Photovoltaïque autoconsommé', 'kWh'],
    ['S2-CHAL-FROID', '2', 'Chaleur et froid achetés à un réseau', 'kWh'],
    ['S3-C1-ACHATS', '3.1', 'Biens et services achetés', 'MAD HT, t (papier)'],
    ['S3-C2-IMMO', '3.2', 'Immobilisations (informatique, GAB, mobilier, bâtiments, véhicules)', 'unité et année, m²'],
    ['S3-C3-ENERGIE', '3.3', 'Amont de l\'énergie (calculé)', 'calculé'],
    ['S3-C4-TRANSP-AMONT', '3.4', 'Transport de fonds, fret entrant', 'passages et km, t·km'],
    ['S3-C5-DECHETS', '3.5', 'Déchets', 't'],
    ['S3-C6-DEPLAC', '3.6', 'Déplacements professionnels (avion, train, voiture, hôtel)', 'voyage IATA, km, nuitées'],
    ['S3-C7-DOMTRAV', '3.7', 'Domicile-travail', 'enquête, km'],
    ['S3-C9-TRANSP-AVAL', '3.9', 'Courrier envoyé aux clients', 't·km, plis'],
    ['S3-VISITEURS', '3 (optionnel)', 'Déplacements des clients et visiteurs', 'personnes et km'],
    ['S3-C15-FINANCE', '3.15', 'Émissions financées (extension PCAF)', 'à définir'],
  ], [1.4, 0.9, 2.2, 1.4]),
  H1n('Annexe B. Correspondance avec l\'ancien canevas'),
  ...T(['Ligne de l\'ancien canevas', 'Nouveau poste', 'Changement'], [
    ['Cuisson, chauffage (butane, propane, bois)', 'S1-COMB-FIXE', 'Bouteilles de 12 kg ou 3 kg ; plus de montant seul'],
    ['Autre : groupe électrogène', 'S1-COMB-FIXE', 'Sous-catégorie dédiée'],
    ['Climatiseur 1 à n (bureaux), climatiseur du véhicule', 'S1-FUGI', 'Recharge annuelle et non plus charge totale'],
    ['Chaleur et froid (électricité des appareils)', 'Supprimé ou S2-CHAL-FROID', 'Réseau de chaleur uniquement, pour éviter le double comptage'],
    ['Véhicule électrique 1 à n', 'S2-ELEC-RESEAU (sous-compteur) ou S1-COMB-MOB', 'kWh de recharge ; libellés d\'année corrigés'],
    ['Petites fournitures ligne à ligne', 'S3-C1-ACHATS', 'Par famille, montant HT, papier en tonnes'],
    ['Repas, cafés', 'S3-C1-ACHATS (restauration)', 'Nombre de repas par type ou montant HT'],
    ['Liste des équipements avec « Âge = 3 ans »', 'S3-C2-IMMO', 'Import de l\'inventaire avec année d\'acquisition'],
    ['Déchets en kg par jour', 'S3-C5-DECHETS', 'Tonnes par an et traitement'],
    ['Trajet AR en texte libre', 'S3-C6-DEPLAC', 'Codes IATA, AR ou aller simple, classe'],
    ['Déplacements clients et visiteurs', 'S3-VISITEURS', 'Optionnel, estimé par enquête'],
    ['Courriers clientèle et internes', 'S3-C9-TRANSP-AVAL / S3-C4', 'Séparation entre aval (clients) et amont ou interne'],
  ], [2, 1.5, 2.2]),
  H1n('Annexe C. Glossaire'),
  ...T(['Terme', 'Définition'], [
    ['tCO2e', 'Tonne d\'équivalent CO2, unité commune à tous les gaz à effet de serre'],
    ['FE', 'Facteur d\'émission : quantité de GES émise par unité d\'activité'],
    ['PRG', 'Pouvoir de réchauffement global d\'un gaz rapporté au CO2, sur 100 ans'],
    ['Location-based / market-based', 'Deux méthodes de calcul du scope 2 : facteur moyen du réseau ou facteur des contrats d\'achat d\'énergie'],
    ['BPR', 'Banque Populaire Régionale'],
    ['ETP', 'Équivalent temps plein'],
    ['PCAF', 'Partnership for Carbon Accounting Financials : méthode de calcul des émissions financées'],
    ['ESRS E1', 'Norme européenne de reporting de durabilité sur le changement climatique'],
    ['SGES', 'Système de gestion environnementale et sociale'],
  ], [1.3, 4]),
);

const doc = new Document({
  creator: 'Groupe BCP — Direction des Risques',
  title: 'Empreinte carbone : méthodologie, guide de collecte et cahier des charges',
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: C.text } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 34, bold: true, color: C.green }, paragraph: { spacing: { before: 360, after: 220 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: C.navy }, paragraph: { spacing: { before: 280, after: 140 }, keepNext: true, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: C.green }, paragraph: { spacing: { before: 200, after: 100 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } } } },
      { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 260 } } } } ] },
    ...['num1', 'num2', 'num3', 'num4'].map(r => ({ reference: r, levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] })),
  ] },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1134, left: 1150, right: 1150, header: 600, footer: 500 } }, titlePage: true },
    headers: {
      default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.green, space: 4 } },
        children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 69, height: 36 }, altText: { title: 'Groupe BCP', description: 'Logo', name: 'logo-entete' } }),
          new TextRun({ text: '\tEmpreinte carbone — méthodologie, collecte et outil web · Version 1.1', size: 16, color: C.navy, font: FONT })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [new TextRun({ text: 'Groupe BCP — document interne', size: 16, color: C.muted, font: FONT }),
          new TextRun({ text: '\tPage ', size: 16, color: C.muted, font: FONT }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.muted, font: FONT }),
          new TextRun({ text: ' / ', size: 16, color: C.muted, font: FONT }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.muted, font: FONT })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: content,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('written', OUT, b.length); });
