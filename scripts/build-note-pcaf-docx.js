/**
 * Note méthodologique et guide d'utilisation de l'outil émissions financées (PCAF, approche IFC). Charte Groupe BCP.
 * Usage : node scripts/build-note-pcaf-docx.js <dossier_captures> <sortie.docx>
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, LevelFormat, TableOfContents, Header, Footer, PageNumber, BorderStyle, PageBreak,
  VerticalAlign, ImageRun, TabStopType, PageOrientation,
} = require('docx');

const [, , CAPS, OUT] = process.argv;
const FONT = 'Arial';
const C = { orange: 'E37B05', brun: '462016', text: '3D3D3B', light: 'F2E8DB', light2: 'FBF6F0', bord: 'D9C7B0', muted: '8A7A6A', rose: 'FBE3E1', rouge: 'B3261E', vert: 'E3F1E8', vertT: '146C4C' };
const LOGO = fs.readFileSync(path.join(__dirname, 'assets', 'groupe-bcp-logo.png'));
const BAND = fs.readFileSync(path.join(__dirname, 'assets', 'groupe-bcp-bandeau.png'));
const PORTRAIT = { W: 9600, FIG: 640 }, LANDSCAPE = { W: 14538, FIG: 960 };
let W = PORTRAIT.W, FIGW = PORTRAIT.FIG;

function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(s => s.startsWith('**')
    ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...o }) : new TextRun({ text: s, font: FONT, ...o }));
}
const P = t => new Paragraph({ children: runs(t), spacing: { after: 120, line: 276 }, alignment: AlignmentType.JUSTIFIED });
const H1 = (t, br = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: br, children: [new TextRun({ text: t, font: FONT })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const B = t => new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: runs(t), spacing: { after: 60, line: 264 } });
let nref = 0; const newList = () => `n${++nref}`;
const N = (t, ref) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 70, line: 264 } });
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });
const border = { style: BorderStyle.SINGLE, size: 4, color: C.bord };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, width, o = {}) {
  return new TableCell({ width: { size: width, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.TOP,
    shading: { fill: o.fill || C.light2, type: ShadingType.CLEAR, color: 'auto' }, margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: String(text ?? '').split('\n').map(l => new Paragraph({ children: runs(l, { size: 17, bold: o.bold, color: o.color || C.text }), spacing: { after: 20, line: 236 } })) });
}
function T(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const widths = ratios.map(r => Math.floor(W * r / sum)); widths[widths.length - 1] += W - widths.reduce((a, b) => a + b, 0);
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: [
    new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, color: 'FFFFFF', fill: C.orange })) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => cell(v, widths[i], { fill: ri % 2 ? C.light2 : C.light, bold: o.firstBold && i === 0, color: o.firstBold && i === 0 ? C.brun : C.text })) })),
  ] }), SP()];
}
function box(title, lines, kind = 'info') {
  const k = { info: [C.light, C.orange, C.brun], attention: [C.rose, C.rouge, C.rouge], ok: [C.vert, C.vertT, C.vertT] }[kind];
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { fill: k[0], type: ShadingType.CLEAR, color: 'auto' },
    borders: { top: border, bottom: border, right: border, left: { style: BorderStyle.SINGLE, size: 36, color: k[1] } }, margins: { top: 110, bottom: 110, left: 180, right: 160 },
    children: [new Paragraph({ children: runs(title, { bold: true, size: 20, color: k[2] }), spacing: { after: 70 } }), ...lines.map(l => new Paragraph({ children: runs(l, { size: 19 }), spacing: { after: 50, line: 252 } }))],
  })] })] }), SP()];
}
let figNo = 0;
function fig(file, caption) {
  const data = fs.readFileSync(path.join(CAPS, file + '.png'));
  const w = data.readUInt32BE(16), h = data.readUInt32BE(20);
  const maxH = FIGW === 960 ? 470 : 760;
  let width = FIGW; if (h * width / w > maxH) width = Math.round(maxH * w / h);
  figNo += 1;
  return [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 80, after: 60 }, children: [new ImageRun({ type: 'png', data, transformation: { width, height: Math.round(h * width / w) }, altText: { title: caption, description: caption, name: file } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: `Figure ${figNo} — ${caption}`, font: FONT, italics: true, size: 18, color: C.brun })] })];
}
const SECTIONS = [{ geo: PORTRAIT, children: [] }];
let content = SECTIONS[0].children;
const add = (...xs) => xs.flat().forEach(x => content.push(x));
function newSection(geo) { W = geo.W; FIGW = geo.FIG; const s = { geo, children: [] }; SECTIONS.push(s); content = s.children; }

// ------------------------------------------------------------------ page de garde
const centered = (text, run, spacing) => new Paragraph({ alignment: AlignmentType.CENTER, spacing, children: [new TextRun({ text, font: FONT, ...run })] });
add(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 500 }, children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 230, height: 120 }, altText: { title: 'Groupe BCP', description: 'Logo Groupe BCP', name: 'logo' } })] }),
  centered('ÉMISSIONS FINANCÉES', { bold: true, size: 44, color: C.orange }, { after: 200 }),
  centered('Note méthodologique et guide d\'utilisation de l\'outil de calcul (standard PCAF, approche IFC)', { size: 28, color: C.brun }, { before: 60, after: 160 }),
  centered('Scope 3, catégorie 15 du GHG Protocol : les émissions des clients et des projets financés par la banque', { italics: true, size: 22, color: C.text }, { after: 300 }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new ImageRun({ type: 'png', data: BAND, transformation: { width: 620, height: 45 }, altText: { title: 'Bandeau', description: 'Bandeau décoratif Groupe BCP', name: 'bandeau' } })] }),
  centered('Destinataires : Direction des Risques, équipe Climat, Engagements, Finance, RSE', { size: 20, color: C.text }, { after: 40 }),
  centered('Version 1.0 — novembre 2026 (outil version 1.0)', { size: 20, color: C.text }, { after: 40 }),
  centered('Classification : usage interne', { size: 20, italics: true, color: C.muted }, { after: 360 }),
  new Paragraph({ indent: { left: 360, right: 360 }, shading: { type: ShadingType.CLEAR, fill: C.light }, spacing: { after: 160 },
    children: runs('Ce document accompagne le classeur « Outil_emissions_financees_PCAF.xlsx ». Les facteurs d\'émission livrés avec l\'outil et le portefeuille d\'exemple sont fictifs : ils servent à montrer le fonctionnement et ne doivent pas être publiés.', { italics: true, size: 19, color: C.brun }) }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ children: [new TextRun({ text: 'Sommaire', font: FONT, size: 34, bold: true, color: C.orange })], spacing: { after: 220 } }),
  new TableOfContents('Sommaire', { hyperlink: true, headingStyleRange: '1-2' }),
);

// ------------------------------------------------------------------ 1
add(H1('1. Pourquoi un outil distinct'),
  P('Le canevas de collecte mesure l\'**empreinte propre** de la banque : électricité, carburants, achats, déplacements… Pour une banque, ces émissions sont faibles au regard des **émissions financées**, c\'est-à-dire celles des entreprises, des projets, des logements et des véhicules qu\'elle finance. Dans les banques qui publient les deux, les émissions financées représentent généralement l\'essentiel du total.'),
  P('Le GHG Protocol les classe en **scope 3, catégorie 15 (investissements)**. Leur calcul repose sur des données différentes de celles du canevas : les encours de crédit et de titres, la valeur des entreprises et des biens financés, et les émissions des clients. Il relève donc d\'un outil séparé, alimenté par le système d\'information de crédit et non par les sites.'),
  ...T(['', 'Canevas de collecte (empreinte propre)', 'Outil émissions financées'], [
    ['Périmètre', 'Scopes 1, 2 et 3 (catégories 1 à 9) des activités de la banque', 'Scope 3, catégorie 15 : clients, projets, biens et titres financés'],
    ['Méthode', 'GHG Protocol, méthodologie IFC de la mesure 2022', 'Standard PCAF (Partie A), approche IFC pour les projets'],
    ['Données', 'Factures, inventaires, extractions des sites et des directions', 'Encours, valeurs des entreprises et des biens, émissions des clients'],
    ['Fournisseurs de données', 'Sites, BPR, directions supports', 'Engagements, Finance, salle des marchés, clients, études d\'impact'],
    ['Fréquence', 'Annuelle', 'Annuelle (arrêté au 31/12), avec suivi possible en cours d\'année'],
  ], [1.2, 2.4, 2.4], { firstBold: true }),
);

// ------------------------------------------------------------------ 2
add(H1('2. Cadre méthodologique'),
  H2('2.1 Les référentiels'),
  ...T(['Référentiel', 'Rôle dans l\'outil'], [
    ['**PCAF — Global GHG Accounting and Reporting Standard for the Financial Industry, Partie A** (2e édition, 2022)', 'Méthode de référence : classes d\'actifs, facteur d\'attribution, options de calcul, score de qualité des données, règles de publication'],
    ['**GHG Protocol, Scope 3 Standard, catégorie 15**', 'Définition du périmètre des émissions financées'],
    ['**IFC** — approche des institutions financières et Performance Standard 3', 'Méthode PCAF promue par l\'IFC auprès des banques qu\'elle accompagne ; pour les projets, émissions estimées selon le Performance Standard 3 (quantification exigée au-delà de 25 000 tCO2e par an) et l\'outil IFC CEET'],
    ['**ESRS E1** (CSRD)', 'Publication des émissions financées dans le scope 3, attendue des partenaires et investisseurs européens'],
    ['**Bank Al-Maghrib**, directive n° 5/W/2021', 'Gestion des risques financiers liés au climat : les émissions financées éclairent le risque de transition du portefeuille'],
  ], [2.4, 3.2]),
  ...box('Alignement avec l\'IFC', ['L\'outil applique la méthode PCAF, que l\'IFC recommande aux institutions financières, et l\'approche de l\'IFC pour les financements de projets (émissions issues de l\'étude d\'impact ou de l\'outil CEET). Les choix de paramétrage (facteurs, périmètre, seuils) sont à **valider avec l\'IFC** dans le prolongement de la mesure 2022.'], 'info'),
  H2('2.2 Le principe de calcul'),
  P('Pour chaque exposition, la banque se voit attribuer une part des émissions de la contrepartie, proportionnelle à sa part dans le financement :'),
  ...box('Formule générale', ['**Émissions financées = facteur d\'attribution × émissions de la contrepartie**', '**Facteur d\'attribution = encours de la banque ÷ valeur totale de la contrepartie** (EVIC, fonds propres + dettes, valeur du bien ou du véhicule, ou PIB selon la classe d\'actifs), plafonné à 100 %.'], 'ok'),
  P('L\'encours retenu est le montant **au bilan** (montant tiré) à la date d\'arrêté. Les engagements hors bilan non tirés ne sont pas comptés, conformément au PCAF.'),
);

// ------------------------------------------------------------------ 3 (paysage)
newSection(LANDSCAPE);
add(H1('3. Les classes d\'actifs et leurs formules', false),
  ...T(['Classe d\'actifs PCAF', 'Facteur d\'attribution', 'Émissions de la contrepartie (de la meilleure à la moins bonne donnée)', 'Données à collecter'], [
    ['Actions cotées et obligations d\'entreprises', 'Encours ÷ EVIC (capitalisation boursière + dettes + intérêts minoritaires + actions préférentielles, en fin d\'exercice)', '1. Émissions publiées et vérifiées ; 2. publiées non vérifiées ; 3. énergie consommée × facteur ; 4. chiffre d\'affaires × facteur sectoriel ; 5. actifs × facteur sectoriel', 'Encours, EVIC, secteur, émissions publiées, chiffre d\'affaires'],
    ['Prêts aux entreprises et actions non cotées', 'Encours ÷ (fonds propres + dettes) ; EVIC si l\'entreprise est cotée', 'Idem', 'Encours, total des fonds propres et des dettes (bilan), secteur, chiffre d\'affaires, émissions si disponibles'],
    ['Financement de projets', 'Encours ÷ (fonds propres + dettes du projet)', 'Émissions du projet (étude d\'impact, outil IFC CEET) ; à défaut, facteur sectoriel', 'Encours, coût total financé du projet, émissions estimées du projet'],
    ['Immobilier commercial', 'Encours ÷ valeur du bien à l\'octroi', '2. Consommation d\'énergie réelle × facteur ; 4. surface × intensité par défaut ; 5. surface moyenne × intensité', 'Encours, valeur à l\'octroi, type de bien, surface, consommation d\'énergie'],
    ['Crédits immobiliers (habitat)', 'Encours ÷ valeur du bien à l\'octroi', 'Idem ; une ligne peut agréger un portefeuille (surface totale, nombre de dossiers)', 'Encours, valeur à l\'octroi, surface, nombre de dossiers'],
    ['Crédits automobiles', 'Encours ÷ valeur du véhicule à l\'octroi', '3. Kilométrage réel × facteur du véhicule ; 4. type connu, kilométrage par défaut ; 5. véhicule moyen', 'Encours, valeur du véhicule, type de motorisation, nombre de dossiers'],
    ['Dette souveraine', 'Encours ÷ PIB en parité de pouvoir d\'achat', 'Émissions territoriales du pays (inventaire national)', 'Encours de titres souverains, inventaire national, PIB (PPA)'],
  ], [1.8, 2.6, 4, 2.6], { firstBold: true }),
  H2('3.1 Le score de qualité des données'),
  P('Le PCAF note chaque donnée de 1 (meilleure) à 5 (plus faible). L\'outil attribue le score automatiquement selon la donnée utilisée, puis calcule un **score moyen pondéré par l\'encours**. L\'objectif est de faire baisser ce score d\'année en année, en obtenant des clients les plus émetteurs leurs émissions réelles.'),
  ...T(['Score', 'Entreprises et projets', 'Immobilier', 'Véhicules', 'Souverain'], [
    ['1', 'Émissions publiées et vérifiées par un tiers', '—', '—', '—'],
    ['2', 'Émissions publiées non vérifiées (ou communiquées par le client)', 'Consommation d\'énergie réelle', '—', 'Inventaire national et PIB (PPA)'],
    ['3', 'Énergie consommée × facteur d\'émission', '—', 'Kilométrage réel et facteur du véhicule', '—'],
    ['4', 'Chiffre d\'affaires × facteur sectoriel', 'Surface connue × intensité par défaut', 'Type de véhicule, kilométrage par défaut', '—'],
    ['5', 'Actifs ou fonds propres + dettes × facteur sectoriel', 'Surface moyenne × intensité par défaut', 'Véhicule moyen', '—'],
  ], [0.5, 3, 2.4, 2.4, 1.6]),
  ...box('Simplification à connaître', ['Le PCAF distingue plus finement certains cas (score 1 pour un bâtiment dont l\'énergie est fournie par contrat spécifique, scores 2 et 3 selon la nature des données physiques). L\'outil retient le score le plus courant pour chaque situation ; un affinage est possible au cas par cas en commentaire.'], 'info'),
);

// ------------------------------------------------------------------ 4 utiliser l'outil
add(H1('4. Utiliser l\'outil'),
  H2('4.1 Le classeur'),
  ...T(['Onglet', 'Contenu', 'Qui intervient'], [
    ['Lisez-moi', 'Mode d\'emploi, formules, scores de qualité, avertissement', 'Tous'],
    ['Paramètres', 'Date d\'arrêté, facteurs d\'émission (énergie, secteurs, immobilier, véhicules), données souveraines, validation des facteurs', 'Équipe Climat'],
    ['Portefeuille', 'Une ligne par exposition (500 lignes) ; colonnes de saisie orangées, colonnes de résultats grises', 'Engagements, Finance, salle des marchés'],
    ['Synthèse', 'Indicateurs clés, tableaux par classe d\'actifs et par secteur, 10 premières expositions', 'Direction des Risques, RSE'],
    ['Contrôle', 'Nombre de lignes par type d\'alerte', 'Tous'],
  ], [1.2, 3.4, 1.8], { firstBold: true }),
  ...fig('p01_lisez_moi', 'Onglet Lisez-moi : étapes et formules par classe d\'actifs'),
  H2('4.2 Étape 1 : paramétrer'),
  ...fig('p02_parametres_general', 'Onglet Paramètres : paramètres généraux, énergie et dette souveraine'),
  P('Les valeurs en **rouge italique** sont des valeurs de démonstration. Remplacez chacune par un facteur sourcé, en indiquant la source et l\'année dans le tableau des secteurs. Le statut passe alors de « ⚠ à remplacer » à « ✓ sourcé ». Une fois tous les facteurs remplacés, passez **« Facteurs validés » à Oui** : l\'avertissement de la synthèse disparaît.'),
  ...fig('p03_parametres_secteurs', 'Facteurs sectoriels (valeurs de démonstration à remplacer)'),
  ...T(['Paramètre', 'Source recommandée'], [
    ['Facteurs sectoriels par chiffre d\'affaires et par actifs', 'Base de facteurs du PCAF (accessible aux établissements signataires), sinon base EXIOBASE ou facteurs monétaires de l\'ADEME, convertis en dirhams et corrigés de l\'inflation'],
    ['Facteur d\'émission de l\'énergie et des bâtiments', 'Facteur du réseau électrique marocain publié (ONEE, ministère de la Transition énergétique, ou AIE), mix avec les combustibles'],
    ['Intensités énergétiques par type de bien', 'Études nationales (AMEE), diagnostics énergétiques, données des promoteurs'],
    ['Facteurs et kilométrages des véhicules', 'ADEME ou DESNZ pour les facteurs ; enquêtes nationales pour les kilométrages'],
    ['Émissions territoriales et PIB (PPA)', 'Dernier inventaire national soumis à la CCNUCC ; PIB en PPA de la Banque mondiale ou du FMI'],
  ], [2.2, 3.4]),
  H2('4.3 Étape 2 : extraire et saisir le portefeuille'),
  P('Extrayez du système d\'information les encours au bilan à la date d\'arrêté. Collez-les dans les cellules orangées de l\'onglet Portefeuille (Collage spécial › Valeurs), une ligne par exposition. Pour l\'habitat et l\'automobile, une ligne peut regrouper un portefeuille : indiquez alors le nombre de dossiers, la surface totale ou la valeur totale.'),
  ...fig('p04_portefeuille_identification', 'Onglet Portefeuille : bandeau (1), groupes de colonnes et ligne d\'aide (2), lignes d\'exemple fictives (3)'),
  ...T(['Colonne', 'Ce qu\'il faut saisir'], [
    ['Classe d\'actifs PCAF', 'Choisie dans la liste : elle détermine la formule. Correspondance à établir une fois avec les codes produits du système d\'information.'],
    ['Secteur d\'activité', 'Obligatoire pour les entreprises et les projets. Correspondance à établir avec la nomenclature d\'activité de la banque.'],
    ['Encours au bilan (MAD)', 'Montant tiré à la date d\'arrêté, net des engagements non tirés.'],
    ['EVIC, fonds propres + dettes', 'EVIC pour les sociétés cotées ; fonds propres + dettes financières du dernier bilan pour les autres ; coût total financé pour un projet.'],
    ['Émissions déclarées, vérification, source', 'Émissions scopes 1 et 2 publiées par le client (rapport ESG, questionnaire) ou estimées pour le projet (étude d\'impact, outil CEET).'],
    ['Chiffre d\'affaires, total actif, énergie', 'Données de repli quand les émissions ne sont pas connues (scores 3 à 5).'],
    ['Valeur du bien, type, surface, consommation', 'Immobilier : valeur à l\'octroi, type de bien, surface, consommation réelle si connue.'],
    ['Nombre, valeur du véhicule, type, kilométrage', 'Automobile : nombre de dossiers agrégés, valeur à l\'octroi, motorisation, kilométrage si connu.'],
  ], [1.8, 4]),
  ...fig('p05_portefeuille_entreprises', 'Colonnes Entreprises et projets (identification à gauche, données d\'émissions à droite)'),
  ...fig('p06_portefeuille_immobilier_vehicules', 'Colonnes Immobilier, agrégat et véhicules'),
  H2('4.4 Étape 3 : lire les résultats et corriger'),
  P('Les colonnes grises se calculent seules : facteur d\'attribution, émissions de la contrepartie, méthode retenue, score de qualité, émissions financées (scopes 1+2 et scope 3) et intensité. La colonne **Contrôle** affiche « OK » ou l\'alerte à corriger.'),
  ...fig('p07_portefeuille_resultats', 'Colonnes de résultats calculées pour les 13 exemples fictifs'),
  ...T(['Alerte', 'Correction'], [
    ['⚠ Choisir la classe d\'actifs / ⚠ Encours manquant', 'Compléter la classe et l\'encours'],
    ['⚠ Secteur manquant', 'Choisir le secteur (entreprises et projets)'],
    ['⚠ EVIC manquante / ⚠ Fonds propres + dettes manquants', 'Renseigner le dénominateur du facteur d\'attribution'],
    ['⚠ Valeur du bien / du véhicule manquante, ⚠ Type de bien manquant', 'Compléter les données de l\'immobilier ou de l\'automobile'],
    ['⚠ Données d\'émissions insuffisantes', 'Saisir au moins une donnée : émissions, énergie, chiffre d\'affaires ou total actif'],
    ['⚠ Attribution plafonnée à 100 %', 'Le dénominateur est inférieur à l\'encours : vérifier l\'EVIC, les fonds propres + dettes ou la valeur du bien'],
  ], [2.6, 3.2]),
  ...fig('p10_controle', 'Onglet Contrôle : nombre de lignes par type d\'alerte'),
  H2('4.5 Étape 4 : la synthèse'),
  ...fig('p08_synthese_indicateurs_classes', 'Onglet Synthèse : indicateurs clés et tableau par classe d\'actifs (exemples fictifs)'),
  ...T(['Indicateur', 'Définition'], [
    ['Émissions financées absolues', 'Somme des émissions attribuées, scopes 1+2 et scope 3 séparément (exigence de publication PCAF)'],
    ['Taux de couverture', 'Part de l\'encours pour laquelle un calcul a pu être fait'],
    ['Intensité économique', 'Émissions financées par million de dirhams d\'encours couvert'],
    ['Score de qualité moyen', 'Moyenne des scores pondérée par l\'encours (1 = meilleur)'],
    ['WACI', 'Intensité carbone moyenne des entreprises financées (tCO2e par million de dirhams de chiffre d\'affaires), pondérée par l\'encours'],
  ], [1.8, 4]),
  ...fig('p09_synthese_top10', 'Les 10 premières expositions par émissions financées'),
);

// ------------------------------------------------------------------ 5 (portrait)
newSection(PORTRAIT);
const lMise = newList();
add(H1('5. Mise en œuvre et limites', false),
  H2('5.1 Démarche recommandée'),
  N('**Valider le paramétrage avec l\'IFC** : périmètre des classes d\'actifs, sources des facteurs, traitement de la dette souveraine et des filiales africaines.', lMise),
  N('**Établir les correspondances** une fois pour toutes : codes produits → classes d\'actifs PCAF ; nomenclature d\'activité → secteurs de l\'outil.', lMise),
  N('**Commencer par les secteurs les plus émetteurs** : production d\'électricité fossile, ciment, sidérurgie, chimie et engrais, mines, transport, agriculture, immobilier. Ils concentrent l\'essentiel des émissions financées.', lMise),
  N('**Collecter les émissions réelles** des principaux clients : rapports ESG des sociétés cotées, questionnaire carbone joint au dossier de crédit, études d\'impact des projets (Performance Standard 3 de l\'IFC).', lMise),
  N('**Calculer l\'année de référence**, puis suivre chaque année l\'évolution des émissions, de l\'intensité et du score de qualité.', lMise),
  H2('5.2 Limites à connaître'),
  B('Les **facteurs livrés sont fictifs** : aucun résultat ne doit être publié avant leur remplacement et la validation des paramètres.'),
  B('Les estimations par le chiffre d\'affaires (score 4) et par les actifs (score 5) donnent un **ordre de grandeur** : elles sont utiles pour repérer les secteurs à enjeu, pas pour piloter un client.'),
  B('Le scope 3 des contreparties n\'est calculé que pour les entreprises et les projets ; le PCAF l\'exige progressivement selon les secteurs.'),
  B('La dette souveraine est publiée **séparément** des autres classes, comme le recommande le PCAF, car ses émissions ne sont pas additives avec celles des entreprises du même pays.'),
  B('Les émissions évitées (projets d\'énergie renouvelable) et les émissions facilitées (marchés de capitaux, PCAF Partie B) ne sont pas calculées.'),
  H2('5.3 Évolutions possibles'),
  B('Alimentation automatique depuis le système d\'information de crédit et l\'outil de notation interne (secteur, chiffre d\'affaires, total bilan déjà saisis).'),
  B('Intégration dans la plateforme web, sur le même socle technique que les outils existants, avec historique, workflow de validation et lien avec l\'outil de risque climatique.'),
  B('Trajectoires sectorielles de décarbonation et objectifs de réduction alignés avec l\'Accord de Paris.'),
);
add(H1('Annexe. Glossaire'),
  ...T(['Terme', 'Définition'], [
    ['PCAF', 'Partnership for Carbon Accounting Financials : initiative qui a établi le standard de calcul des émissions financées'],
    ['Émissions financées', 'Émissions des clients et projets attribuées à la banque au prorata de son financement (scope 3, catégorie 15)'],
    ['EVIC', 'Enterprise value including cash : capitalisation boursière + dettes + intérêts minoritaires + actions préférentielles'],
    ['Facteur d\'attribution', 'Part des émissions de la contrepartie attribuée à la banque : encours ÷ valeur totale de la contrepartie'],
    ['Score de qualité', 'Note de 1 à 5 de la qualité de la donnée d\'émissions (1 = émissions vérifiées)'],
    ['WACI', 'Weighted average carbon intensity : intensité carbone moyenne pondérée par l\'encours'],
    ['PPA', 'Parité de pouvoir d\'achat'],
    ['CEET', 'Carbon Emissions Estimator Tool : outil de l\'IFC d\'estimation des émissions des projets'],
    ['Performance Standard 3', 'Norme de performance de l\'IFC sur l\'utilisation efficace des ressources et la prévention de la pollution'],
  ], [1.4, 4.2], { firstBold: true }),
);

// ------------------------------------------------------------------ document
const hs = (id, name, size, color, before, after, lvl) => ({ id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size, bold: true, color }, paragraph: { spacing: { before, after }, keepNext: true, outlineLevel: lvl } });
const doc = new Document({
  creator: 'Groupe BCP — Direction des Risques', title: 'Émissions financées : note méthodologique et guide de l\'outil',
  styles: { default: { document: { run: { font: FONT, size: 21, color: C.text } } }, paragraphStyles: [hs('Heading1', 'Heading 1', 34, C.orange, 360, 220, 0), hs('Heading2', 'Heading 2', 26, C.brun, 280, 140, 1)] },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 260 } }, run: { color: C.orange } } }] },
    ...Array.from({ length: nref }, (_, i) => ({ reference: `n${i + 1}`, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } }, run: { bold: true, color: C.orange } } }] })),
  ] },
  features: { updateFields: true },
  sections: SECTIONS.map((s, i) => ({
    properties: { page: { size: { width: 11906, height: 16838, orientation: s.geo === LANDSCAPE ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT }, margin: { top: 1300, bottom: 1134, left: 1150, right: 1150, header: 600, footer: 500 } }, titlePage: i === 0 },
    headers: {
      default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: s.geo.W }], border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.orange, space: 4 } },
        children: [new ImageRun({ type: 'png', data: LOGO, transformation: { width: 69, height: 36 }, altText: { title: 'Groupe BCP', description: 'Logo', name: 'logo-entete' } }),
          new TextRun({ text: '\tÉmissions financées — méthodologie et outil · Version 1.0', size: 16, color: C.brun, font: FONT })] })] }),
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
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('written', OUT, (b.length / 1048576).toFixed(1), 'Mo', figNo, 'figures'); });
