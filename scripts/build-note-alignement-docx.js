/**
 * Note d'analyse : alignement du canevas de collecte avec la méthodologie IFC et les pratiques bancaires.
 * Charte Groupe BCP. Usage : node scripts/build-note-alignement-docx.js <sortie.docx>
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, LevelFormat, TableOfContents, Header, Footer, PageNumber, BorderStyle, PageBreak,
  VerticalAlign, ImageRun, TabStopType, PageOrientation,
} = require('docx');

const OUT = process.argv[2];
const FONT = 'Arial';
const C = { orange: 'E37B05', brun: '462016', text: '3D3D3B', light: 'F2E8DB', light2: 'FBF6F0', bord: 'D9C7B0', muted: '8A7A6A', rose: 'FBE3E1', rouge: 'B3261E', vert: 'E3F1E8', vertT: '146C4C', ambre: 'FFF1D6' };
const LOGO = fs.readFileSync(path.join(__dirname, 'assets', 'groupe-bcp-logo.png'));
const BAND = fs.readFileSync(path.join(__dirname, 'assets', 'groupe-bcp-bandeau.png'));
const PORTRAIT = { W: 9600 }, LANDSCAPE = { W: 14538 };
let W = PORTRAIT.W;

function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(s => s.startsWith('**')
    ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...o }) : new TextRun({ text: s, font: FONT, ...o }));
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run || {}), spacing: { after: 120, line: 276 }, alignment: AlignmentType.JUSTIFIED });
const H1 = (t, br = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: br, children: [new TextRun({ text: t, font: FONT })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const B = t => new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: runs(t), spacing: { after: 60, line: 264 } });
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });
const border = { style: BorderStyle.SINGLE, size: 4, color: C.bord };
const borders = { top: border, bottom: border, left: border, right: border };
const STATUT_FILL = { 'Identique': C.vert, 'Amélioré': C.vert, 'Équivalent': C.light, 'Écart mineur': C.ambre, 'Écart': C.rose, 'Absent': C.rose, 'Écart assumé': C.light2 };

function cell(text, width, o = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
    shading: { fill: o.fill || C.light2, type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: String(text ?? '').split('\n').map(l => new Paragraph({ children: runs(l, { size: o.size || 17, bold: o.bold, color: o.color || C.text }), spacing: { after: 20, line: 236 } })),
  });
}
function T(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map(r => Math.floor(W * r / sum));
  widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, color: 'FFFFFF', fill: C.orange })) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => {
      const st = o.statusCol === i ? STATUT_FILL[v] : null;
      return cell(v, widths[i], { fill: st || (ri % 2 ? C.light2 : C.light), bold: (o.firstBold && i === 0) || o.statusCol === i, color: (o.firstBold && i === 0) ? C.brun : C.text });
    }) })),
  ] }), SP()];
}
function box(title, lines, kind = 'info') {
  const k = { info: [C.light, C.orange, C.brun], attention: [C.rose, C.rouge, C.rouge], ok: [C.vert, C.vertT, C.vertT] }[kind];
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { fill: k[0], type: ShadingType.CLEAR, color: 'auto' },
    borders: { top: border, bottom: border, right: border, left: { style: BorderStyle.SINGLE, size: 36, color: k[1] } },
    margins: { top: 110, bottom: 110, left: 180, right: 160 },
    children: [new Paragraph({ children: runs(title, { bold: true, size: 20, color: k[2] }), spacing: { after: 70 } }),
      ...lines.map(l => new Paragraph({ children: runs(l, { size: 19 }), spacing: { after: 50, line: 252 } }))],
  })] })] }), SP()];
}
const SECTIONS = [{ geo: PORTRAIT, children: [] }];
let content = SECTIONS[0].children;
const add = (...xs) => xs.flat().forEach(x => content.push(x));
function newSection(geo) { W = geo.W; const s = { geo, children: [] }; SECTIONS.push(s); content = s.children; }

// ------------------------------------------------------------------ page de garde
const centered = (text, run, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing, children: [new TextRun({ text, font: FONT, ...run })] });
add(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 500 }, children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 230, height: 120 }, altText: { title: 'Groupe BCP', description: 'Logo Groupe BCP', name: 'logo' } })] }),
  centered('NOTE D\'ANALYSE', { bold: true, size: 44, color: C.orange }, { after: 200 }),
  centered('Alignement du canevas de collecte de l\'empreinte carbone avec la méthodologie IFC et les pratiques bancaires', { size: 28, color: C.brun }, { before: 60, after: 160 }),
  centered('Écarts constatés, différences assumées et propositions pour une version 2 du canevas', { italics: true, size: 22, color: C.text }, { after: 300 }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new ImageRun({ type: 'png', data: BAND, transformation: { width: 620, height: 45 }, altText: { title: 'Bandeau', description: 'Bandeau décoratif Groupe BCP', name: 'bandeau' } })] }),
  centered('Destinataires : Direction des Risques, RSE, équipe Empreinte carbone', { size: 20, color: C.text }, { after: 40 }),
  centered('Version 1.1 — novembre 2026 (améliorations mises en œuvre dans le canevas 2.0)', { size: 20, color: C.text }, { after: 40 }),
  centered('Classification : usage interne', { size: 20, italics: true, color: C.muted }, { after: 360 }),
  new Paragraph({ indent: { left: 360, right: 360 }, shading: { type: ShadingType.CLEAR, fill: C.light }, spacing: { after: 160 },
    children: runs('Base de comparaison : canevas IFC « Carbon Footprint Data Request » (versions 231015, 231027 et 231108, onglets « Périmètre-fonction », « Sources données » et onglets par type de site) et canevas de collecte commenté (version novembre 2026).', { italics: true, size: 19, color: C.brun }) }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: [new TextRun({ text: 'Sommaire', font: FONT, size: 34, bold: true, color: C.orange })], spacing: { after: 220 } }),
  new TableOfContents('Sommaire', { hyperlink: true, headingStyleRange: '1-2' }),
);

// ------------------------------------------------------------------ données de correspondance
const MAP = [
  // [scope, poste IFC, donnée IFC, canevas, statut, commentaire]
  ['1', 'Sources fixes : butane', 'kg, nombre de bouteilles ou montant (dh)', '2-Carburants et gaz : quantité en kg ou bouteilles de 12 / 3 kg ; montant HT et prix unitaire', 'Amélioré', 'Mêmes options que l\'IFC ; le prix unitaire rend la conversion du montant traçable. L\'exemple IFC « bouteille de 13 kg » est corrigé (12 kg ou 3 kg au Maroc).'],
  ['1', 'Sources fixes : propane', 'kg ou montant', '2-Carburants et gaz : kg, m³ ou montant HT et prix', 'Amélioré', '—'],
  ['1', 'Sources fixes : bois', 'kg', '2-Carburants et gaz : combustible « Bois », kg', 'Identique', '—'],
  ['1', 'Autres sources fixes (groupe électrogène)', 'Litres ou m³ de gasoil', '2-Carburants et gaz : usage « Groupe électrogène », litres ou m³', 'Identique', '—'],
  ['1', 'Sources mobiles : gasoil, essence', 'Litres ou montant (dh)', '2-Carburants et gaz : usage « Véhicules de service », litres ou montant HT et prix', 'Amélioré', '—'],
  ['1', 'Fugitives : climatiseurs des bureaux', 'Nature du gaz, quantité (kg), âge de l\'appareil', '3-Climatisation : fluide, charge nominale, quantité rechargée, année de mise en service, nombre d\'appareils', 'Amélioré', 'Le canevas contient toutes les données IFC (charge et âge) et ajoute la recharge, qui permet la méthode du bilan massique du GHG Protocol.'],
  ['1', 'Fugitives : climatiseurs des véhicules', 'Nature du gaz, quantité, marque et modèle du véhicule', '3-Climatisation : type « Climatiseur de véhicule », fluide, charge, recharge', 'Écart mineur', 'La marque et le modèle du véhicule ne sont pas demandés explicitement ; ils permettent à l\'IFC d\'estimer la charge quand elle est inconnue.'],
  ['2', 'Électricité du réseau', 'kWh des 12 factures', '1-Électricité : kWh par compteur et par mois, montant HT, usage du compteur', 'Amélioré', 'Détail mensuel et par compteur ; les sous-compteurs évitent le double comptage.'],
  ['2', 'Électricité photovoltaïque', 'kWh produits et puissance installée (kW)', '1-Électricité : usage « Production photovoltaïque », kWh produits', 'Écart mineur', 'La puissance installée (kWc) n\'a pas de colonne : elle n\'apparaît qu\'en commentaire dans l\'exemple.'],
  ['2', 'Véhicules électriques', 'Kilométrage annuel et consommation aux 100 km', '1-Électricité : usage « Recharge véhicules électriques », kWh', 'Écart', 'La méthode IFC (km × consommation) n\'est pas possible : seul le kWh de recharge est demandé, rarement mesuré séparément.'],
  ['2', 'Chaleur et froid', 'Consommation électrique des appareils chauffants et refroidissants', 'Aucune ligne', 'Écart assumé', 'Retiré volontairement : cette électricité est déjà dans la facture (double comptage). Le GHG Protocol ne compte en scope 2 que la chaleur ou le froid **achetés** à un réseau. Une ligne « réseau de chaleur ou de froid » avec statut Néant reste à prévoir pour garder la trace du poste IFC.'],
  ['3', 'Achats : petites fournitures', 'Nature, quantité et montant, ligne à ligne', '4-Achats : familles (papier, consommables d\'impression, fournitures, hygiène), montant HT, quantité facultative', 'Équivalent', 'Agrégation par famille : l\'IFC applique de toute façon des facteurs monétaires par catégorie. Le détail des 230 références n\'apportait pas de précision.'],
  ['3', 'Achats : abonnements internet', 'Montant', '4-Achats : « Télécoms et abonnements internet »', 'Identique', '—'],
  ['3', 'Achats : impressions, publicité', 'Montant', '4-Achats : « Impressions externalisées et publicité »', 'Identique', '—'],
  ['3', 'Achats : maintenance', 'Montant', '4-Achats : « Maintenance des équipements et bâtiments »', 'Identique', '—'],
  ['3', 'Achats : gardiennage', 'Montant', '4-Achats : « Gardiennage »', 'Identique', '—'],
  ['3', 'Achats : location de véhicules', 'Montant', '4-Achats : « Location de véhicules »', 'Identique', '—'],
  ['3', 'Achats : matériaux de travaux', 'Nature du matériau et quantité (kg)', '4-Achats : « Travaux et matériaux », précision, quantité en kg ou tonnes', 'Équivalent', 'La nature du matériau est saisie en texte libre (colonne Précision), sans liste (acier, aluminium, verre, bois, béton).'],
  ['3', 'Achats : location des agences (loyers)', 'Montant annuel du loyer', 'Aucune famille « Loyers »', 'Absent', 'Poste de l\'onglet IFC « Périmètre-fonction ». À ajouter comme famille d\'achat.'],
  ['3', 'Achats : papier vierge et papier recyclé', 'Deux lignes distinctes (onglet « Sources données »)', '4-Achats : une seule famille « Papier »', 'Écart mineur', 'Le facteur d\'émission du papier recyclé est nettement plus faible : la distinction est utile.'],
  ['3', 'Achats : repas, café', 'Nombre ou montant', '4-Achats : « Restauration (repas, café) », nombre de repas ou de doses', 'Identique', '—'],
  ['3', 'Immobilisations : informatique, téléphonie, impression', 'Liste des équipements et âge', '5-Équipements : catégorie, quantité, année d\'achat', 'Amélioré', 'L\'année d\'achat remplace l\'âge (tranches ouvertes impossibles) ; regroupement par catégorie et par année.'],
  ['3', 'Immobilisations : mobilier, véhicules de service', 'Liste et âge', '5-Équipements : bureaux, chaises, armoires, véhicules', 'Identique', '—'],
  ['3', 'Immobilisations : bâtiments', 'Superficie totale construite', '5-Équipements : bâtiment, m², année de construction', 'Amélioré', '—'],
  ['3', 'Immobilisations : GAB', 'Liste et âge', '5-Équipements : GAB, modèle, quantité, année', 'Identique', '—'],
  ['3', 'Immobilisations : autres', '« Autres immobilisations (préciser) »', 'Pas de catégorie « Autre » dans la liste', 'Écart mineur', 'Un équipement hors liste (coffre-fort, compteuse de billets, groupe électrogène) ne peut pas être saisi.'],
  ['3', 'Déchets (7 flux)', 'kg par jour et devenir ; à défaut, nombre d\'employés', '6-Déchets : 11 flux, quantité (kg, t, sacs, bacs), période, traitement ; effectif dans Identification', 'Amélioré', 'Les mêmes flux, plus archives, cartouches, piles, déchets alimentaires ; le repli « nombre d\'employés » est conservé via l\'effectif.'],
  ['3', 'Transport amont : fret des fournisseurs', 'km annuel, type de véhicule, consommation aux 100 km', '9-Transport : seulement « Colis / messagerie » ou « Autre »', 'Écart', 'Pas de type « livraisons fournisseurs », ni de colonnes type de véhicule et consommation.'],
  ['3', 'Transport amont : transport de fonds', 'km annuel, type de véhicule, consommation aux 100 km', '9-Transport : passages par an × distance aller-retour', 'Écart mineur', 'Le kilométrage est reconstitué (plus fiable), mais le type de véhicule (blindé) et la consommation ne sont pas demandés.'],
  ['3', 'Domicile-travail', 'Nombre d\'employés par mode et distance moyenne **aller-retour**', '8-Domicile-travail : nombre par mode, distance **aller**, jours de présence par semaine', 'Équivalent', 'Distance aller (plus simple à estimer, l\'équipe double) et jours de présence en plus, indispensables au calcul annuel.'],
  ['3', 'Déplacements pro : avion Maroc et international', 'Trajet aller-retour et nombre de voyages', '7-Déplacements pro : villes avec code IATA, aller simple ou aller-retour, classe, nombre', 'Amélioré', 'Supprime le texte libre ; la classe affine le facteur d\'émission.'],
  ['3', 'Déplacements pro : frais Maroc et international', 'Montant, séparé Maroc / international', '7-Déplacements pro : montant HT par mode', 'Écart mineur', 'Sur une ligne « montant seul », la distinction Maroc / international n\'existe pas.'],
  ['3', 'Déplacements pro : véhicules personnels', 'km, type de véhicule, consommation', '7-Déplacements pro : km remboursés', 'Écart mineur', 'Le carburant du véhicule (diesel, essence) n\'est pas demandé.'],
  ['3', 'Flotte de véhicules', 'Quantité, kilométrage annuel, gaz de climatisation, âge', 'Litres (onglet 2), véhicules et année (onglet 5), climatisation (onglet 3)', 'Écart mineur', 'Le kilométrage annuel, repli quand les litres manquent, n\'a pas de colonne.'],
  ['3', 'Déplacements des clients et visiteurs', 'Nombre de personnes et km moyen par mode (voiture, bus, taxi, train, avion, moto)', 'Aucun onglet', 'Absent', 'Poste présent dans tous les onglets IFC. Il est optionnel au sens du GHG Protocol mais demandé par l\'IFC.'],
  ['3', 'Transport aval : courrier clients', 'Tonnes et km', '9-Transport : nombre d\'envois, poids, distance', 'Identique', '—'],
  ['3', 'Transport aval : courrier interne', 'Tonnes et km', '9-Transport : « Navette courrier interne », tournées et km', 'Identique', '—'],
  ['—', 'Précision dans la donnée', 'Colonne texte libre', 'Méthode (liste), pièce justificative, commentaire, statut', 'Amélioré', 'Permet un score de qualité par donnée.'],
  ['—', 'Périmètre par type de site', 'Un onglet par type de site', 'Onglet Identification (type de site, code, entité)', 'Équivalent', 'Un fichier par site, rattaché au référentiel des sites.'],
  ['—', 'Année des données', 'Cellule « Données de l\'année »', 'Exercice (Identification)', 'Identique', '—'],
];
const count = s => MAP.filter(r => r[4] === s).length;
const nOK = count('Identique') + count('Amélioré'), nEq = count('Équivalent'), nMin = count('Écart mineur'), nEc = count('Écart') + count('Absent'), nAs = count('Écart assumé');

// ------------------------------------------------------------------ 1. Synthèse
add(H1('1. Synthèse'),
  P(`Le canevas a été comparé ligne par ligne aux ${MAP.length} rubriques du canevas IFC, puis aux référentiels internationaux (GHG Protocol, ISO 14064-1, ESRS E1) et aux pratiques de reporting des banques marocaines et internationales.`),
  ...box('Verdict', [
    `**Le canevas est aligné avec la méthodologie IFC.** Il reprend les trois scopes et les mêmes postes d'émission, dans le même découpage. Sur ${MAP.length} rubriques IFC : **${nOK}** sont reprises à l'identique ou améliorées, **${nEq}** sont couvertes de manière équivalente, **${nMin}** présentent un écart mineur, **${nEc}** un écart réel et **${nAs}** un écart assumé.`,
    'Les données nécessaires au calcul IFC sont donc présentes, sauf pour **deux postes absents** (déplacements des clients et visiteurs, loyers des agences) et **deux méthodes de repli** non prévues (véhicules électriques au kilomètre, fret amont avec type de véhicule).',
    'Les différences de forme (recharge de fluide, distance aller, tonnes par an, familles d\'achat) sont **voulues** : elles corrigent les défauts de la collecte 2022 et restent convertibles vers le format IFC.',
  ], 'ok'),
  H2('1.1 Écarts à corriger pour être complet par rapport à l\'IFC'),
  ...T(['Écart', 'Onglet', 'Correction proposée'], [
    ['Déplacements des clients et visiteurs absents', '8', 'Ajouter un type de ligne « Visiteurs et clients » (nombre de personnes, km moyen, mode)'],
    ['Loyers des agences absents', '4', 'Ajouter la famille « Loyers et charges locatives »'],
    ['Véhicules électriques : pas de méthode au kilomètre', '1 ou 2', 'Colonnes facultatives « kilométrage annuel » et « consommation aux 100 km »'],
    ['Fret amont et transport de fonds sans type de véhicule ni consommation', '9', 'Type « Livraisons fournisseurs » ; colonnes « type de véhicule » et « consommation aux 100 km »'],
    ['Réseau de chaleur ou de froid sans ligne', '1', 'Ajouter l\'usage « Chaleur ou froid acheté à un réseau » (statut Néant le plus souvent)'],
    ['Écarts mineurs (kWc, papier recyclé, autres immobilisations, modèle de véhicule, flotte en km, zone des frais)', '1, 3, 4, 5, 7', 'Colonnes ou valeurs de liste complémentaires (section 5)'],
  ], [2.3, 0.6, 2.8]),
  H2('1.2 Ce que les pratiques bancaires ajoutent'),
  B('**Scope 2 « market-based »** : les banques soumises à l\'ESRS E1 publient aussi le scope 2 selon leurs contrats d\'électricité renouvelable. Le canevas ne collecte pas le type de contrat.'),
  B('**Data centers** : les banques suivent l\'indicateur PUE (efficacité énergétique). Il faut pour cela l\'énergie consommée par les seuls équipements informatiques.'),
  B('**Fluides** : le bilan massique du GHG Protocol tient compte du fluide récupéré lors du démantèlement des appareils.'),
  B('**Indicateurs environnementaux voisins** : eau et papier figurent dans les rapports ESG des banques cotées au Maroc. Ils peuvent être collectés par le même canevas.'),
  B('**Émissions financées (scope 3, catégorie 15)** : c\'est de loin le premier poste d\'une banque. Il relève d\'un outil distinct (méthode PCAF), pas de ce canevas.'),
  ...box('Recommandation', [
    'Produire une **version 2 du canevas** intégrant les corrections de la section 1.1 et une sélection des pratiques de la section 1.2, et y ajouter un **onglet « Export IFC »** qui recompose automatiquement les données au format du canevas IFC. Cela garantit la continuité avec la mesure 2022 réalisée avec l\'IFC et la comparabilité d\'une année à l\'autre.',
  ]),
);

// ------------------------------------------------------------------ 2. Référentiels
add(H1('2. Référentiels de comparaison'),
  ...T(['Référentiel', 'Portée', 'Usage dans cette note'], [
    ['Canevas IFC « Carbon Footprint Data Request » (2023)', 'Demande de données de la mesure 2022, conduite avec l\'IFC', 'Référence principale : chaque rubrique est comparée (section 3)'],
    ['GHG Protocol Corporate Standard et Scope 3 Standard', 'Norme internationale de comptabilité carbone ; base de la méthode IFC', 'Définition des scopes, des catégories et des méthodes (bilan massique, scope 2 location / market)'],
    ['GHG Protocol Scope 2 Guidance', 'Double publication du scope 2', 'Données de contrat d\'électricité'],
    ['ISO 14064-1:2018', 'Norme de quantification et de vérification', 'Exigences de traçabilité et de pièces justificatives'],
    ['Bilan Carbone (ADEME, Base Empreinte)', 'Méthode et facteurs d\'émission très utilisés par les entreprises marocaines et francophones', 'Postes « visiteurs », « immobilisations amorties », facteurs monétaires'],
    ['ESRS E1 (CSRD)', 'Norme européenne ; attendue par les partenaires et investisseurs européens', 'Scopes 1, 2 (location et market) et 3 ; intensité'],
    ['PCAF', 'Méthode des émissions financées des banques', 'Hors canevas ; rappel du périmètre (catégorie 15)'],
    ['Bank Al-Maghrib, directive n° 5/W/2021', 'Gestion des risques financiers liés au climat et à l\'environnement', 'Contexte réglementaire national'],
    ['AMMC, circulaire n° 03/19', 'Rapport ESG des émetteurs faisant appel public à l\'épargne', 'Indicateurs environnementaux publiés par les banques cotées'],
  ], [1.7, 2, 2.1], { firstBold: true }),
  P('Les pratiques des banques ont été appréciées à partir de ce que publient habituellement les grandes banques marocaines cotées dans leurs rapports ESG (énergie, flotte, fluides, papier, informatique, déplacements, déchets, eau) et les banques européennes soumises à la CSRD (scope 2 market-based, data centers, voyages par classe, télétravail, émissions financées). Elles sont citées comme pratiques courantes, sans référence à un établissement en particulier.'),
);

// ------------------------------------------------------------------ 3. Correspondance (paysage)
newSection(LANDSCAPE);
add(H1('3. Correspondance ligne à ligne avec le canevas IFC', false),
  P('Légende de la colonne « Statut » : **Identique** — même donnée ; **Amélioré** — même donnée, mieux structurée ou plus précise ; **Équivalent** — donnée différente mais suffisante pour le même calcul ; **Écart mineur** — information secondaire manquante ; **Écart** ou **Absent** — donnée IFC impossible à obtenir avec le canevas ; **Écart assumé** — différence voulue et justifiée.'),
  ...T(['Scope', 'Rubrique IFC', 'Donnée demandée par l\'IFC', 'Dans le canevas', 'Statut', 'Commentaire'], MAP, [0.45, 1.6, 1.9, 2.6, 0.9, 3.3], { statusCol: 4 }),
);

// ------------------------------------------------------------------ 4. Différences assumées
add(H1('4. Différences de méthode assumées'),
  P('Ces différences sont volontaires. Elles corrigent des difficultés constatées en 2022 et chaque donnée du canevas reste convertible vers la forme attendue par l\'IFC.'),
  ...T(['Sujet', 'Canevas IFC', 'Canevas actuel', 'Pourquoi', 'Conversion vers le format IFC'], [
    ['Fluides frigorigènes', 'Nature, quantité (charge) et âge', 'Charge, quantité rechargée, année, nombre d\'appareils', 'La charge totale déclarée comme fuite surestimait les émissions (46,4 kg de R22 au siège). La recharge est la mesure directe de la fuite (bilan massique du GHG Protocol).', 'Charge et âge sont toujours disponibles : l\'IFC peut appliquer sa méthode par taux de fuite.'],
    ['Chaleur et froid', 'Électricité des appareils chauffants et refroidissants', 'Supprimé (sauf réseau de chaleur)', 'Cette électricité est déjà dans la facture : double comptage. Le scope 2 ne compte que la chaleur achetée.', 'Ligne IFC à zéro, avec renvoi vers l\'électricité.'],
    ['Déchets', 'kg par jour', 'Quantité sur une période (an, mois, semaine), kg, t, sacs, bacs', 'Le kg par jour est rarement connu ; les bordereaux sont annuels.', 'Total annuel ÷ jours ouvrés.'],
    ['Domicile-travail', 'Distance aller-retour', 'Distance aller et jours de présence', 'Plus simple à estimer ; les jours de présence (télétravail) sont indispensables au calcul annuel.', 'Distance × 2.'],
    ['Achats', 'Ligne à ligne (nature, quantité, montant)', 'Montant HT par famille, quantité facultative', 'Le détail article par article (230 lignes) était ingérable et sans gain de précision.', 'Famille → catégorie IFC ; facteurs monétaires identiques.'],
    ['Montants', 'Non précisé (HT ou TTC)', 'Toujours HT', 'Base homogène ; les facteurs monétaires s\'appliquent hors taxes.', 'Aucune.'],
    ['Âge des équipements', 'Âge en années', 'Année d\'achat', 'Évite les tranches ouvertes (« plus de 4 ans ») et reste juste d\'un exercice à l\'autre.', 'Âge = exercice − année.'],
    ['Avion', 'Trajet AR en texte libre', 'Villes de départ et d\'arrivée (IATA), AR ou aller simple, classe', 'Distances calculables sans retraitement.', 'Regroupement par trajet ; Maroc / international déduit des villes.'],
    ['Qualité', 'Colonne « Précision dans la donnée »', 'Statut, méthode, pièce justificative, contrôle automatique', 'Distingue l\'absence d\'activité de la donnée manquante ; permet un score de qualité.', 'Méthode → texte de la colonne « Précision ».'],
  ], [1.1, 1.6, 1.8, 3.2, 2.2], { firstBold: true }),
);

// ------------------------------------------------------------------ 5. Pratiques bancaires
add(H1('5. Comparaison avec les pratiques des banques'),
  ...T(['Pratique', 'Référence', 'Pratique des banques', 'Dans le canevas', 'Statut', 'Recommandation'], [
    ['Scope 2 en double approche (location-based et market-based)', 'GHG Protocol Scope 2 Guidance ; ESRS E1-6', 'Banques européennes : systématique. Banques marocaines : se développe avec les contrats d\'électricité renouvelable (lois 13-09 et 82-21).', 'Non collecté', 'Écart', 'Colonne « Type de contrat » : réseau standard, électricité renouvelable sous contrat, autoproduction'],
    ['Fluides : bilan massique', 'GHG Protocol', 'Recharge moins fluide récupéré', 'Recharge seulement', 'Écart mineur', 'Colonne facultative « fluide récupéré (kg) » lors des démantèlements'],
    ['Data center : efficacité énergétique (PUE)', 'Pratique courante ; EU Code of Conduct for Data Centres', 'Indicateur suivi par les banques disposant de data centers', 'Non collecté', 'Écart mineur', 'Pour le type de site « Data center » : kWh des équipements informatiques'],
    ['Voyages : classe, hôtel, forçage radiatif', 'DESNZ, ADEME', 'Classe et nuitées collectées ; forçage radiatif appliqué au calcul', 'Classe et nuitées collectées', 'Identique', 'Aucune ; le forçage radiatif est une option de calcul'],
    ['Télétravail', 'GHG Protocol catégorie 7 (optionnel)', 'Quelques banques européennes l\'estiment', 'Jours de présence collectés', 'Équivalent', 'Aucune ; les jours de télétravail se déduisent des jours de présence'],
    ['Locaux loués à des tiers', 'GHG Protocol catégorie 13', 'Surfaces louées à des tiers déclarées à part', 'Seulement « compteur partagé avec un tiers »', 'Écart mineur', 'Identification : surface louée à des tiers'],
    ['Eau et papier', 'GRI 301 et 303 ; rapports ESG (AMMC)', 'Publiés par les banques cotées au Maroc', 'Papier collecté ; eau absente', 'Écart mineur', 'Onglet ou ligne « Eau (m³) » hors calcul GES'],
    ['Plusieurs sites dans un fichier', 'Pratique de collecte', 'Les directions centrales répondent pour tout le réseau', 'Code du site en commentaire', 'Écart mineur', 'Colonne « Code site » sur chaque ligne'],
    ['Qualité des données', 'ISO 14064-1 ; logique des scores PCAF', 'Score de qualité par donnée', 'Méthode, statut et pièce', 'Identique', 'Aucune'],
    ['Émissions financées', 'PCAF ; ESRS E1 ; attentes de Bank Al-Maghrib', 'Premier poste d\'une banque ; outil dédié', 'Hors périmètre', 'Écart assumé', 'Outil distinct, relié à l\'outil de notation et à l\'outil de risque climatique'],
  ], [1.6, 1.5, 2.2, 1.5, 0.9, 2.3], { statusCol: 4 }),
);

// ------------------------------------------------------------------ 6. Version 2
add(H1('6. Propositions d\'amélioration : canevas version 2'),
  ...box('Mise en œuvre', ['Les **20 améliorations** ci-dessous sont intégrées dans le **canevas version 2.0** (novembre 2026), avec l\'onglet « Export IFC ». Le guide utilisateur a été mis à jour en conséquence.'], 'ok'),
  P('Les améliorations ci-dessous restent dans le cadre de la méthodologie IFC : elles ajoutent des données ou des contrôles sans retirer aucune rubrique IFC. Elles sont classées par priorité.'),
  ...T(['N°', 'Amélioration', 'Onglet', 'Origine', 'Priorité', 'Effort'], [
    ['1', 'Lignes « Visiteurs et clients » : mode, nombre de personnes, km moyen par personne', '8 (renommé « Domicile-travail et visiteurs »)', 'IFC', 'Haute', 'Faible'],
    ['2', 'Famille d\'achat « Loyers et charges locatives »', '4', 'IFC', 'Haute', 'Très faible'],
    ['3', 'Véhicules électriques : colonnes facultatives kilométrage annuel et consommation aux 100 km', '2 (combustible « Électricité ») ou 1', 'IFC', 'Haute', 'Faible'],
    ['4', 'Transport : type « Livraisons fournisseurs (fret amont) » ; colonnes type de véhicule et consommation aux 100 km', '9', 'IFC', 'Haute', 'Faible'],
    ['5', 'Usage « Chaleur ou froid acheté à un réseau »', '1', 'IFC / GHG Protocol', 'Haute', 'Très faible'],
    ['6', 'Onglet « Export IFC » : recomposition automatique au format du canevas IFC (formules), pour la continuité avec la mesure 2022', 'Nouveau', 'IFC', 'Haute', 'Moyen'],
    ['7', 'Colonne « Code site » sur chaque ligne, pour les fichiers multi-sites des directions centrales', 'Tous', 'Pratique', 'Haute', 'Faible'],
    ['8', 'Colonne « Puissance installée (kWc) » pour le photovoltaïque', '1', 'IFC', 'Moyenne', 'Très faible'],
    ['9', 'Familles « Papier recyclé » et « Papier vierge »', '4', 'IFC', 'Moyenne', 'Très faible'],
    ['10', 'Catégorie « Autre équipement (préciser) »', '5', 'IFC', 'Moyenne', 'Très faible'],
    ['11', 'Climatisation des véhicules : colonne « Marque et modèle »', '3', 'IFC', 'Moyenne', 'Très faible'],
    ['12', 'Colonne facultative « Kilométrage annuel » pour la flotte', '2', 'IFC', 'Moyenne', 'Très faible'],
    ['13', 'Liste « Zone » (Maroc / international) pour les frais de déplacement ; carburant du véhicule personnel', '7', 'IFC', 'Moyenne', 'Très faible'],
    ['14', 'Liste des matériaux de travaux (acier, aluminium, béton, verre, bois)', '4', 'IFC', 'Moyenne', 'Très faible'],
    ['15', 'Colonne « Type de contrat d\'électricité » (scope 2 market-based)', '1', 'Pratique internationale', 'Moyenne', 'Faible'],
    ['16', 'Colonne facultative « Fluide récupéré (kg) »', '3', 'GHG Protocol', 'Basse', 'Très faible'],
    ['17', 'kWh des équipements informatiques pour les data centers (PUE)', '1', 'Pratique bancaire', 'Basse', 'Faible'],
    ['18', 'Surface louée à des tiers ; consommation d\'eau (m³)', 'Identification, 1', 'Pratique ESG', 'Basse', 'Faible'],
    ['19', 'Contrôles de cohérence supplémentaires : 12 mois d\'électricité, ratio kWh/m² hors plage, total des employés ≠ effectif', '1, 8, Contrôle', 'Qualité', 'Moyenne', 'Moyen'],
    ['20', 'Protection de la structure (seules les cellules orangées modifiables) et numéro de version du canevas', 'Tous', 'Qualité', 'Moyenne', 'Faible'],
  ], [0.4, 4.2, 1.6, 1.2, 0.8, 0.8]),
  ...box('Ce que la version 2 ne change pas', [
    'Les neuf onglets, l\'ordre des colonnes, les règles de saisie (nombre seul, montant HT, statut obligatoire) et le contrôle automatique restent identiques : les contributeurs déjà formés ne sont pas perturbés, et le guide utilisateur est mis à jour, pas réécrit.',
  ]),
);

// ------------------------------------------------------------------ 7. Points à valider
newSection(PORTRAIT);
add(H1('7. Points à valider avec l\'IFC et la Direction', false),
  ...T(['Point', 'Question', 'Proposition'], [
    ['Visiteurs et clients', 'Le poste est optionnel au sens du GHG Protocol mais demandé par l\'IFC. Le conserve-t-on dans le périmètre publié ?', 'Le collecter (estimation par enquête), le publier séparément'],
    ['Chaleur et froid', 'L\'IFC accepte-t-elle de ne plus collecter l\'électricité des appareils, déjà comprise dans la facture ?', 'Oui, avec une ligne réseau de chaleur à Néant'],
    ['Domicile-travail', 'Distance aller au lieu d\'aller-retour : conversion acceptée ?', 'Oui, conversion automatique dans l\'export IFC'],
    ['Scope 2 market-based', 'Publier dès le prochain exercice ?', 'Collecter le type de contrat dès la version 2'],
    ['Consolidation', 'Contrôle opérationnel ou financier ; traitement des filiales', 'Contrôle opérationnel, cohérent avec la méthodologie'],
    ['Émissions financées', 'Calendrier de l\'outil PCAF', 'Projet distinct'],
  ], [1.2, 2.6, 1.8], { firstBold: true }),
  H2('Prochaines étapes'),
  B('Validation de cette note et des améliorations retenues.'),
  B('Production du canevas version 2 et de l\'onglet « Export IFC ».'),
  B('Mise à jour du guide utilisateur et de ses copies d\'écran (scripts déjà en place).'),
  B('Test sur une BPR pilote avant la diffusion de la campagne.'),
);

// ------------------------------------------------------------------ document
const hs = (id, name, size, color, before, after, lvl) => ({ id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size, bold: true, color }, paragraph: { spacing: { before, after }, keepNext: true, outlineLevel: lvl } });
const doc = new Document({
  creator: 'Groupe BCP — Direction des Risques',
  title: 'Alignement du canevas de collecte avec la méthodologie IFC',
  styles: { default: { document: { run: { font: FONT, size: 21, color: C.text } } },
    paragraphStyles: [hs('Heading1', 'Heading 1', 34, C.orange, 360, 220, 0), hs('Heading2', 'Heading 2', 26, C.brun, 280, 140, 1)] },
  numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } }, run: { color: C.orange } } }] }] },
  features: { updateFields: true },
  sections: SECTIONS.map((s, i) => ({
    properties: { page: { size: { width: 11906, height: 16838, orientation: s.geo === LANDSCAPE ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
      margin: { top: 1300, bottom: 1134, left: 1150, right: 1150, header: 600, footer: 500 } }, titlePage: i === 0 },
    headers: {
      default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: s.geo.W }], border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.orange, space: 4 } },
        children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 69, height: 36 }, altText: { title: 'Groupe BCP', description: 'Logo', name: 'logo-entete' } }),
          new TextRun({ text: '\tAlignement du canevas avec la méthodologie IFC · Version 1.0', size: 16, color: C.brun, font: FONT })] })] }),
      first: new Header({ children: [new Paragraph({ children: [] })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: s.geo.W }],
        children: [new TextRun({ text: 'Groupe BCP — document interne', size: 16, color: C.muted, font: FONT }), new TextRun({ text: '\tPage ', size: 16, color: C.muted, font: FONT }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.muted, font: FONT }), new TextRun({ text: ' / ', size: 16, color: C.muted, font: FONT }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.muted, font: FONT })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
    },
    children: s.children,
  })),
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('written', OUT, MAP.length, 'rubriques', { nOK, nEq, nMin, nEc, nAs }); });
