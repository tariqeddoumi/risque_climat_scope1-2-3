/**
 * Présentation Direction générale et directeurs : empreinte carbone, scopes 1-2-3, avec et hors émissions financées.
 * Charte Groupe BCP.
 *
 * Usage : NODE_PATH=<node_modules : pptxgenjs, react, react-dom, react-icons, sharp> node scripts/build-presentation-dg-pptx.js <sortie.pptx> [apply_theme.js]
 */
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const md = require('react-icons/md');

const [, , OUT, APPLY] = process.argv;
const A = p => path.join(__dirname, 'assets', p);
const THEME = {
  name: 'Groupe BCP', headFontFace: 'Arial', bodyFontFace: 'Arial',
  colors: { dk1: '3D3D3B', lt1: 'FFFFFF', dk2: '462016', lt2: 'F2E8DB', accent1: 'E37B05', accent2: '462016', accent3: '8C5B52', accent4: 'B3261E', accent5: '146C4C', accent6: 'D9C7B0', hlink: 'E37B05', folHlink: '8C5B52' },
};
const HEX = { orange: 'E37B05', brun: '462016', brun2: '8C5B52', tint: 'F2E8DB', tint2: 'FBF6F0', bord: 'D9C7B0', rouge: 'B3261E', vert: '146C4C', gris: '3D3D3B' };

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(md[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = 'Empreinte carbone du Groupe — scopes 1, 2 et 3, avec et hors émissions financées';
  pres.author = 'Groupe BCP — Direction des Risques';
  pres.company = 'Groupe BCP';
  pres.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial' };
  const C = pres.SchemeColor;
  const FOOT = 'Empreinte carbone du Groupe — présentation à la Direction générale — novembre 2026';

  pres.defineSlideMaster({
    title: 'TITRE', background: { path: A('groupe-bcp-fond-orange.png') },
    objects: [
      { image: { path: A('groupe-bcp-logo-blanc.png'), x: 0.7, y: 0.5, w: 1.6, h: 0.835 } },
      { image: { path: A('groupe-bcp-motif.png'), x: 10.35, y: 1.6, w: 2.6, h: 3.84 } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.3, w: 9.3, h: 1.9, fontFace: 'Arial', fontSize: 38, bold: true, color: C.background1, align: 'left', valign: 'bottom', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.4, w: 9.3, h: 1.4, fontFace: 'Arial', fontSize: 18, color: C.background1, align: 'left', valign: 'top', margin: 0 }, text: '' } },
    ],
  });
  pres.defineSlideMaster({
    title: 'CONTENU', background: { color: 'FFFFFF' },
    objects: [
      { image: { path: A('groupe-bcp-bandeau.png'), x: 0.45, y: 0.28, w: 12.43, h: 0.85 } },
      { image: { path: A('groupe-bcp-logo-blanc.png'), x: 11.55, y: 0.36, w: 1.3, h: 0.68 } },
      { placeholder: { options: { name: 'title', type: 'title', x: 1.6, y: 0.28, w: 10.1, h: 0.85, fontFace: 'Arial', fontSize: 22, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0 }, text: '' } },
      { text: { text: FOOT, options: { x: 0.55, y: 7.02, w: 9.5, h: 0.3, fontFace: 'Arial', fontSize: 9, color: C.accent3, margin: 0 } } },
    ],
    slideNumber: { x: 12.25, y: 7.02, w: 0.6, h: 0.3, fontFace: 'Arial', fontSize: 9, color: C.accent3, align: 'right' },
  });

  const content = (title, section) => { const s = pres.addSlide({ masterName: 'CONTENU', sectionTitle: section }); s.addText(title, { placeholder: 'title' }); return s; };
  const T = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, color: C.text1, valign: 'top', ...o });
  const box = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: o.line ?? fill, width: o.lw ?? 0.75 }, objectName: o.name ?? 'bloc' });
  const circleIcon = async (s, name, x, y, d, bg = HEX.orange, fg = 'FFFFFF') => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg }, objectName: 'pastille icône' });
    s.addImage({ data: await icon(name, fg), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, altText: name, objectName: 'icône' });
  };

  // 1 ---------------------------------------------------------------- titre
  pres.addSection({ title: 'Ouverture' });
  let s = pres.addSlide({ masterName: 'TITRE', sectionTitle: 'Ouverture' });
  s.addText('EMPREINTE CARBONE DU GROUPE\nScopes 1, 2 et 3, avec et hors émissions financées', { placeholder: 'title' });
  s.addText('Présentation à la Direction générale et aux directeurs — novembre 2026', { placeholder: 'body' });
  s.addNotes('Objectif : partager le cadre de mesure, l\'état du dispositif et les décisions attendues. Les chiffres de l\'exercice seront présentés après la consolidation de la campagne de collecte et la validation des facteurs d\'émission.');

  // 2 ---------------------------------------------------------------- 3 messages
  s = content('L\'ESSENTIEL EN TROIS MESSAGES', 'Ouverture');
  const msgs = [
    ['1', 'Deux périmètres à distinguer', 'L\'empreinte propre (nos bâtiments, véhicules, achats, déplacements) et les émissions financées (celles des clients et des projets que nous finançons).'],
    ['2', 'Les émissions financées dominent', 'Pour une banque, elles représentent l\'essentiel du total : c\'est là que se situent le risque de transition et les leviers d\'action.'],
    ['3', 'Le dispositif est prêt', 'Canevas de collecte 2.0 aligné sur la méthodologie IFC, guides utilisateurs, et outil de calcul des émissions financées selon le standard PCAF.'],
  ];
  msgs.forEach((m, i) => {
    const x = 0.55 + i * 4.15;
    box(s, x, 1.55, 3.95, 4.9, i === 1 ? HEX.brun : HEX.tint);
    T(s, m[0], { x: x + 0.35, y: 1.8, w: 1.2, h: 1.0, fontSize: 54, bold: true, color: i === 1 ? C.background1 : C.accent1 });
    T(s, m[1], { x: x + 0.35, y: 2.95, w: 3.3, h: 0.9, fontSize: 20, bold: true, color: i === 1 ? C.background1 : C.text2 });
    T(s, m[2], { x: x + 0.35, y: 3.95, w: 3.3, h: 2.3, fontSize: 15, color: i === 1 ? C.background1 : C.text1 });
  });
  s.addNotes('Message central : mesurer l\'empreinte propre est nécessaire et attendu, mais l\'enjeu climatique d\'une banque se joue surtout dans son portefeuille de financements.');

  // 3 ---------------------------------------------------------------- pourquoi
  s = content('POURQUOI MESURER MAINTENANT', 'Ouverture');
  const why = [
    ['MdGavel', 'Régulation', 'Directive Bank Al-Maghrib n° 5/W/2021 sur les risques financiers liés au climat ; rapport ESG des émetteurs (AMMC).'],
    ['MdPublic', 'Partenaires et investisseurs', 'IFC et bailleurs internationaux, contreparties européennes soumises à la CSRD : les données carbone sont demandées.'],
    ['MdTrendingDown', 'Risque de transition', 'Les secteurs émetteurs du portefeuille sont exposés à la hausse du coût du carbone et aux nouvelles normes.'],
    ['MdInsights', 'Pilotage et opportunités', 'Fixer des objectifs, orienter le financement vers la transition, réduire nos consommations.'],
  ];
  for (let i = 0; i < why.length; i++) {
    const x = 0.55 + (i % 2) * 6.2, y = 1.5 + Math.floor(i / 2) * 2.7;
    box(s, x, y, 6.0, 2.45, HEX.tint2, { line: HEX.bord });
    await circleIcon(s, why[i][0], x + 0.35, y + 0.4, 1.0);
    T(s, why[i][1], { x: x + 1.65, y: y + 0.38, w: 4.1, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    T(s, why[i][2], { x: x + 1.65, y: y + 0.95, w: 4.1, h: 1.35, fontSize: 15 });
  }
  s.addNotes('Les références réglementaires sont à confirmer par la Conformité. Les attentes des partenaires se traduisent déjà par des questionnaires carbone dans les due diligences.');

  // 4 ---------------------------------------------------------------- scopes
  pres.addSection({ title: 'Le cadre' });
  s = content('LES SCOPES D\'UNE BANQUE', 'Le cadre');
  const scopes = [
    ['MdLocalGasStation', 'SCOPE 1', 'Émissions directes', 'Carburant des véhicules, groupes électrogènes, gaz, fuites de fluides de climatisation', HEX.orange],
    ['MdBolt', 'SCOPE 2', 'Électricité achetée', 'Agences, sièges, data centers', HEX.orange],
    ['MdShoppingCart', 'SCOPE 3', 'Nos autres activités', 'Achats, équipements, déchets, déplacements, trajets domicile-travail, visiteurs, transport de fonds', HEX.orange],
    ['MdAccountBalance', 'SCOPE 3 — CAT. 15', 'Émissions financées', 'Entreprises, projets, logements, véhicules et titres que nous finançons', HEX.brun],
  ];
  for (let i = 0; i < scopes.length; i++) {
    const x = 0.55 + i * 3.1;
    const sc = scopes[i];
    box(s, x, 1.95, 2.9, 3.85, i === 3 ? HEX.brun : HEX.tint);
    await circleIcon(s, sc[0], x + 0.95, 2.2, 1.0, i === 3 ? HEX.orange : HEX.orange);
    T(s, sc[1], { x: x + 0.15, y: 3.35, w: 2.6, h: 0.45, fontSize: 18, bold: true, align: 'center', color: i === 3 ? C.background1 : C.accent1 });
    T(s, sc[2], { x: x + 0.15, y: 3.8, w: 2.6, h: 0.45, fontSize: 15, bold: true, align: 'center', color: i === 3 ? C.background1 : C.text2 });
    T(s, sc[3], { x: x + 0.2, y: 4.3, w: 2.5, h: 1.4, fontSize: 13, align: 'center', color: i === 3 ? C.background1 : C.text1 });
  }
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: 0.55, y: 1.45, w: 9.1, h: 0.36, fill: { color: HEX.bord }, line: { color: HEX.bord }, objectName: 'accolade empreinte propre' });
  T(s, 'EMPREINTE PROPRE', { x: 3.3, y: 1.45, w: 3.6, h: 0.36, fontSize: 13, bold: true, align: 'center', valign: 'middle', color: C.text2 });
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: 9.85, y: 1.45, w: 2.9, h: 0.36, fill: { color: HEX.brun2 }, line: { color: HEX.brun2 }, objectName: 'accolade financées' });
  T(s, 'FINANCÉES', { x: 9.85, y: 1.45, w: 2.9, h: 0.36, fontSize: 13, bold: true, align: 'center', valign: 'middle', color: C.background1 });
  T(s, 'Référentiel : GHG Protocol (scopes 1, 2, 3), méthodologie IFC pour l\'empreinte propre, standard PCAF pour les émissions financées.', { x: 0.55, y: 6.1, w: 12.25, h: 0.5, fontSize: 14, italic: true, color: C.accent3 });
  s.addNotes('La catégorie 15 du scope 3 (« investissements ») regroupe les émissions des contreparties financées, au prorata de la part de la banque dans leur financement.');

  // 5 ---------------------------------------------------------------- avec / hors
  s = content('HORS ET AVEC ÉMISSIONS FINANCÉES', 'Le cadre');
  const rows = [['', 'Hors émissions financées : empreinte propre', 'Émissions financées'],
    ['Ce que l\'on mesure', 'Scopes 1, 2 et 3 de nos activités', 'Scope 3, catégorie 15 : émissions des clients et projets financés'],
    ['Méthode', 'GHG Protocol, méthodologie IFC (mesure 2022)', 'Standard PCAF, approche IFC pour les projets'],
    ['Données', 'Factures, inventaires, achats, déplacements', 'Encours, valeur des entreprises et des biens, émissions des clients'],
    ['Qui fournit', 'Sites, BPR, directions supports', 'Engagements, Finance, salle des marchés, clients'],
    ['Outil', 'Canevas de collecte 2.0 et guides', 'Outil de calcul PCAF (7 classes d\'actifs)'],
    ['Levier d\'action', 'Énergie, flotte, achats, déplacements', 'Politique de financement et accompagnement des clients']];
  s.addTable(rows.map((r, i) => r.map((v, j) => ({ text: v, options: {
    bold: i === 0 || j === 0, fontSize: i === 0 ? 15 : 14, color: i === 0 ? 'FFFFFF' : (j === 0 ? HEX.brun : HEX.gris),
    fill: { color: i === 0 ? (j === 2 ? HEX.brun : HEX.orange) : (i % 2 ? HEX.tint2 : HEX.tint) }, valign: 'middle', margin: [4, 8, 4, 8] } }))),
  { x: 0.55, y: 1.45, w: 12.25, colW: [2.45, 4.9, 4.9], rowH: 0.68, border: { type: 'solid', pt: 1, color: 'FFFFFF' } });
  s.addNotes('Les deux périmètres se publient séparément : on présente toujours l\'empreinte propre, puis les émissions financées, et un total « avec » qui les additionne.');

  // 6 ---------------------------------------------------------------- ordre de grandeur
  s = content('UN ORDRE DE GRANDEUR À AVOIR EN TÊTE', 'Le cadre');
  const big = 3.9, small = big / Math.sqrt(700);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 1.55 + big - small, w: small, h: small, fill: { color: HEX.orange }, line: { color: HEX.orange }, objectName: 'carré empreinte propre' });
  s.addShape(pres.shapes.RECTANGLE, { x: 1.6, y: 1.55, w: big, h: big, fill: { color: HEX.brun }, line: { color: HEX.brun }, objectName: 'carré émissions financées' });
  T(s, 'Empreinte propre', { x: 0.55, y: 5.6, w: 1.0, h: 0.6, fontSize: 12, bold: true, color: C.accent1, align: 'center' });
  T(s, 'Émissions financées', { x: 1.6, y: 5.6, w: big, h: 0.4, fontSize: 14, bold: true, color: C.text2, align: 'center' });
  T(s, '× 700', { x: 6.2, y: 1.6, w: 6.6, h: 1.4, fontSize: 72, bold: true, color: C.accent1 });
  T(s, 'En moyenne, les émissions financées des institutions financières sont plus de 700 fois supérieures à leurs émissions directes.', { x: 6.2, y: 3.05, w: 6.5, h: 1.2, fontSize: 18, bold: true, color: C.text2 });
  T(s, 'Le rapport exact dépend du portefeuille : il sera calculé pour le Groupe. Le message reste le même : l\'essentiel de notre impact et de notre risque climatique se trouve dans nos financements.', { x: 6.2, y: 4.3, w: 6.5, h: 1.3, fontSize: 15 });
  T(s, 'Source : CDP, Financial Services Disclosure Report « The Time to Green Finance », 2020. Surfaces des carrés proportionnelles.', { x: 6.2, y: 5.85, w: 6.5, h: 0.5, fontSize: 11, italic: true, color: C.accent3 });
  s.addNotes('Chiffre de place, issu des déclarations des institutions financières au CDP. Il ne s\'agit pas du ratio du Groupe, qui sera établi avec l\'outil PCAF.');

  // 7 ---------------------------------------------------------------- empreinte propre : postes
  pres.addSection({ title: 'Empreinte propre' });
  s = content('EMPREINTE PROPRE : CE QUE NOUS MESURONS', 'Empreinte propre');
  const postes = [['MdBolt', 'Électricité', 'Scope 2'], ['MdLocalGasStation', 'Carburants et gaz', 'Scope 1'], ['MdAcUnit', 'Climatisation', 'Scope 1'],
    ['MdShoppingCart', 'Achats', 'Scope 3'], ['MdComputer', 'Équipements et GAB', 'Scope 3'], ['MdDeleteOutline', 'Déchets', 'Scope 3'],
    ['MdFlight', 'Déplacements professionnels', 'Scope 3'], ['MdDirectionsBus', 'Domicile-travail et visiteurs', 'Scope 3'], ['MdLocalShipping', 'Transport de fonds et courrier', 'Scope 3']];
  for (let i = 0; i < postes.length; i++) {
    const x = 0.55 + (i % 3) * 4.15, y = 1.45 + Math.floor(i / 3) * 1.8;
    box(s, x, y, 3.95, 1.6, HEX.tint2, { line: HEX.bord });
    await circleIcon(s, postes[i][0], x + 0.3, y + 0.33, 0.95, postes[i][2] === 'Scope 3' ? HEX.brun2 : HEX.orange);
    T(s, postes[i][1], { x: x + 1.45, y: y + 0.32, w: 2.35, h: 0.75, fontSize: 16, bold: true, color: C.text2, valign: 'middle' });
    T(s, postes[i][2], { x: x + 1.45, y: y + 1.05, w: 2.35, h: 0.35, fontSize: 13, color: C.accent1, bold: true });
  }
  s.addNotes('Neuf postes, couverts par les neuf onglets du canevas de collecte. Les visiteurs et les loyers ont été ajoutés pour couvrir toutes les rubriques IFC.');

  // 8 ---------------------------------------------------------------- enseignements 2022
  s = content('EMPREINTE PROPRE : LES ENSEIGNEMENTS DE 2022', 'Empreinte propre');
  const stats = [['2022', 'première mesure, avec l\'IFC'], ['3', 'BPR sans aucune donnée transmise'], ['28', 'anomalies de saisie relevées'], ['40', 'rubriques IFC reprises dans le canevas 2.0']];
  stats.forEach((st, i) => {
    const x = 0.55 + i * 3.1;
    T(s, st[0], { x, y: 1.5, w: 2.9, h: 1.1, fontSize: 54, bold: true, color: C.accent1, align: 'center' });
    T(s, st[1], { x, y: 2.65, w: 2.9, h: 0.7, fontSize: 14, align: 'center', color: C.text2 });
  });
  box(s, 0.55, 3.65, 6.0, 2.95, HEX.tint);
  T(s, 'CE QUI A FREINÉ', { x: 0.85, y: 3.85, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.accent4 });
  T(s, [{ text: 'Unités et montants mélangés (« 35555 kw », TTC)', options: { bullet: true, breakLine: true } }, { text: 'Doubles comptages (chaleur, salle informatique)', options: { bullet: true, breakLine: true } },
    { text: 'Fuites de climatisation surestimées', options: { bullet: true, breakLine: true } }, { text: 'Suivi de collecte manuel, sans relance', options: { bullet: true } }], { x: 0.85, y: 4.3, w: 5.5, h: 2.2, fontSize: 15, paraSpaceAfter: 6 });
  box(s, 6.8, 3.65, 6.0, 2.95, HEX.brun);
  T(s, 'CE QUE CORRIGE LE CANEVAS 2.0', { x: 7.1, y: 3.85, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
  T(s, [{ text: 'Unités imposées, listes déroulantes, montants HT', options: { bullet: true, breakLine: true } }, { text: 'Statut obligatoire et pièce justificative', options: { bullet: true, breakLine: true } },
    { text: 'Contrôles automatiques et de cohérence', options: { bullet: true, breakLine: true } }, { text: 'Export automatique au format IFC', options: { bullet: true } }], { x: 7.1, y: 4.3, w: 5.5, h: 2.2, fontSize: 15, color: C.background1, paraSpaceAfter: 6 });
  s.addNotes('La démarche 2022 était solide sur le fond (GHG Protocol, matrice des responsables) ; ce sont les règles de saisie qui ont produit des données hétérogènes. Les BPR Laâyoune, Fès-Meknès et Nador-Al Hoceima n\'avaient rien transmis.');

  // 9 ---------------------------------------------------------------- dispositif 2.0
  s = content('EMPREINTE PROPRE : LE DISPOSITIF DE COLLECTE 2.0', 'Empreinte propre');
  const flow = [['MdStore', 'Sites et directions', 'saisissent dans le canevas'], ['MdGroups', 'Correspondants BPR', 'coordonnent et vérifient'], ['MdFactCheck', 'Contrôles', 'automatiques et de cohérence'], ['MdCalculate', 'Équipe Empreinte carbone', 'consolide, calcule, export IFC'], ['MdPublish', 'Publication', 'rapport ESG, BAM, partenaires']];
  for (let i = 0; i < flow.length; i++) {
    const x = 0.55 + i * 2.5;
    await circleIcon(s, flow[i][0], x + 0.6, 1.55, 1.1, i === 4 ? HEX.brun : HEX.orange);
    if (i < flow.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 1.85, y: 1.92, w: 0.55, h: 0.36, fill: { color: HEX.bord }, line: { color: HEX.bord }, objectName: 'flèche' });
    T(s, flow[i][1], { x, y: 2.8, w: 2.3, h: 0.7, fontSize: 15, bold: true, align: 'center', color: C.text2 });
    T(s, flow[i][2], { x, y: 3.45, w: 2.3, h: 0.7, fontSize: 13, align: 'center' });
  }
  const livr = ['Canevas de collecte 2.0', 'Guide utilisateur (Word)', 'Guide de formation (PowerPoint)', 'Note méthodologique', 'Note d\'alignement IFC'];
  T(s, 'DÉJÀ DISPONIBLES', { x: 0.55, y: 4.65, w: 12.25, h: 0.4, fontSize: 15, bold: true, color: C.accent3 });
  livr.forEach((l, i) => { box(s, 0.55 + i * 2.5, 5.15, 2.3, 1.15, HEX.tint); T(s, l, { x: 0.65 + i * 2.5, y: 5.15, w: 2.1, h: 1.15, fontSize: 14, bold: true, align: 'center', valign: 'middle', color: C.text2 }); });
  s.addNotes('Le canevas Excel est la première étape ; l\'outil web décrit dans la note méthodologique reprendra le même format, avec relances automatiques et historique.');

  // 10 --------------------------------------------------------------- PCAF
  pres.addSection({ title: 'Émissions financées' });
  s = content('ÉMISSIONS FINANCÉES : LA MÉTHODE PCAF', 'Émissions financées');
  const eq = [['Encours de la banque', HEX.tint, HEX.brun], ['÷', null], ['Valeur de la contrepartie', HEX.tint, HEX.brun], ['×', null], ['Émissions de la contrepartie', HEX.tint, HEX.brun], ['=', null], ['Émissions financées', HEX.brun, 'FFFFFF']];
  let ex = 0.55;
  eq.forEach(e => {
    if (!e[1]) { T(s, e[0], { x: ex, y: 1.55, w: 0.55, h: 1.2, fontSize: 36, bold: true, color: C.accent1, align: 'center', valign: 'middle' }); ex += 0.55; return; }
    box(s, ex, 1.55, 2.55, 1.2, e[1]); T(s, e[0], { x: ex + 0.1, y: 1.55, w: 2.35, h: 1.2, fontSize: 16, bold: true, align: 'center', valign: 'middle', color: e[2] === 'FFFFFF' ? C.background1 : C.text2 }); ex += 2.55;
  });
  T(s, 'La banque se voit attribuer la part des émissions du client correspondant à sa part dans le financement.', { x: 0.55, y: 2.95, w: 12.25, h: 0.45, fontSize: 15, italic: true, color: C.accent3 });
  T(s, '7 CLASSES D\'ACTIFS', { x: 0.55, y: 3.6, w: 6.5, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
  ['Entreprises cotées', 'Entreprises non cotées', 'Financement de projets', 'Immobilier commercial', 'Crédits habitat', 'Crédits automobiles', 'Dette souveraine'].forEach((c, i) => {
    const x = 0.55 + (i % 4) * 1.6 + (i >= 4 ? 0.8 : 0), y = 4.1 + Math.floor(i / 4) * 1.15;
    box(s, x, y, 1.5, 0.95, HEX.tint2, { line: HEX.bord }); T(s, c, { x: x + 0.05, y, w: 1.4, h: 0.95, fontSize: 12, bold: true, align: 'center', valign: 'middle', color: C.text2 });
  });
  T(s, 'SCORE DE QUALITÉ DES DONNÉES', { x: 7.4, y: 3.6, w: 5.4, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
  const dq = [['1', 'Émissions vérifiées', HEX.vert], ['2', 'Émissions déclarées', '4E8A5A'], ['3', 'Données physiques', 'C9A227'], ['4', 'Chiffre d\'affaires × secteur', HEX.orange], ['5', 'Estimation par les actifs', HEX.rouge]];
  dq.forEach((d, i) => {
    const y = 4.1 + i * 0.47;
    s.addShape(pres.shapes.RECTANGLE, { x: 7.4, y, w: 0.5, h: 0.4, fill: { color: d[2] }, line: { color: d[2] }, objectName: 'score' });
    T(s, d[0], { x: 7.4, y, w: 0.5, h: 0.4, fontSize: 14, bold: true, align: 'center', valign: 'middle', color: C.background1 });
    T(s, d[1], { x: 8.05, y, w: 4.7, h: 0.4, fontSize: 13, valign: 'middle' });
  });
  s.addNotes('Standard PCAF (Partnership for Carbon Accounting Financials), promu par l\'IFC auprès des banques. Pour les financements de projets, les émissions viennent de l\'étude d\'impact ou de l\'outil IFC CEET. L\'objectif est d\'améliorer le score de qualité d\'année en année en obtenant les émissions réelles des principaux clients.');

  // 11 --------------------------------------------------------------- secteurs
  s = content('ÉMISSIONS FINANCÉES : OÙ SE CONCENTRENT LES ENJEUX', 'Émissions financées');
  const sect = [['MdElectricBolt', 'Électricité d\'origine fossile'], ['MdFactory', 'Ciment et matériaux'], ['MdPrecisionManufacturing', 'Sidérurgie et métallurgie'], ['MdScience', 'Chimie et engrais'],
    ['MdTerrain', 'Mines et phosphates'], ['MdLocalShipping', 'Transport et logistique'], ['MdAgriculture', 'Agriculture'], ['MdApartment', 'Immobilier']];
  for (let i = 0; i < sect.length; i++) {
    const x = 0.55 + (i % 4) * 3.1, y = 1.5 + Math.floor(i / 4) * 2.05;
    box(s, x, y, 2.9, 1.8, i < 4 ? HEX.brun : HEX.tint);
    await circleIcon(s, sect[i][0], x + 0.95, y + 0.2, 0.95, HEX.orange);
    T(s, sect[i][1], { x: x + 0.15, y: y + 1.2, w: 2.6, h: 0.5, fontSize: 14, bold: true, align: 'center', color: i < 4 ? C.background1 : C.text2 });
  }
  box(s, 0.55, 5.75, 12.25, 0.95, HEX.tint2, { line: HEX.bord });
  T(s, [{ text: 'Priorité : ', options: { bold: true, color: C.accent1 } }, { text: 'obtenir les émissions réelles des principaux clients de ces secteurs (rapports ESG, questionnaire carbone joint au dossier de crédit, études d\'impact des projets).' }],
    { x: 0.85, y: 5.75, w: 11.7, h: 0.95, fontSize: 15, valign: 'middle' });
  s.addNotes('Secteurs à forte intensité carbone en général ; leur poids exact dans nos émissions financées sera mesuré avec l\'outil PCAF. Les quatre premiers sont généralement les plus intensifs.');

  // 12 --------------------------------------------------------------- tableau de bord
  pres.addSection({ title: 'Pilotage' });
  s = content('LE TABLEAU DE BORD QUE NOUS PRÉSENTERONS', 'Pilotage');
  const kpi = [['Scopes 1 et 2', 'tCO2e', 'Empreinte propre', HEX.orange], ['Scope 3 propre', 'tCO2e', 'Empreinte propre', HEX.orange], ['Intensité', 'tCO2e par collaborateur', 'Empreinte propre', HEX.orange],
    ['Émissions financées', 'tCO2e', 'Scope 3, catégorie 15', HEX.brun], ['Intensité du portefeuille', 'tCO2e par million de dirhams', 'Émissions financées', HEX.brun], ['Qualité et couverture', 'score PCAF et % de l\'encours', 'Émissions financées', HEX.brun]];
  kpi.forEach((k, i) => {
    const x = 0.55 + (i % 3) * 4.15, y = 1.5 + Math.floor(i / 3) * 2.45;
    box(s, x, y, 3.95, 2.2, k[3] === HEX.brun ? HEX.brun : HEX.tint);
    const w = k[3] === HEX.brun;
    T(s, k[2].toUpperCase(), { x: x + 0.3, y: y + 0.2, w: 3.4, h: 0.35, fontSize: 11, bold: true, color: w ? C.accent6 : C.accent3 });
    T(s, 'À consolider', { x: x + 0.3, y: y + 0.55, w: 3.4, h: 0.7, fontSize: 24, bold: true, italic: true, color: w ? C.accent1 : C.accent1, valign: 'middle' });
    T(s, k[0], { x: x + 0.3, y: y + 1.3, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: w ? C.background1 : C.text2 });
    T(s, k[1], { x: x + 0.3, y: y + 1.7, w: 3.4, h: 0.4, fontSize: 13, color: w ? C.background1 : C.text1 });
  });
  T(s, 'Valeurs présentées après la consolidation de la campagne de collecte et la validation des facteurs d\'émission. Total « avec émissions financées » = empreinte propre + émissions financées.', { x: 0.55, y: 6.4, w: 12.25, h: 0.5, fontSize: 13, italic: true, color: C.accent3 });
  s.addNotes('Ces six indicateurs constituent le socle de publication : ils répondent aux attentes de Bank Al-Maghrib, de l\'IFC et des investisseurs. La dette souveraine sera publiée à part, comme le recommande le PCAF.');

  // 13 --------------------------------------------------------------- feuille de route
  s = content('FEUILLE DE ROUTE', 'Pilotage');
  const road = [['Nov.–déc. 2026', 'Validation', 'Périmètre, référents, calendrier ; paramétrage des facteurs'], ['Janv.–mars 2027', 'Collecte', 'Campagne avec le canevas 2.0 ; relances par les BPR'],
    ['Avril 2027', 'Empreinte propre', 'Contrôles, consolidation, calcul, export IFC'], ['T2 2027', 'Émissions financées', 'Premier calcul PCAF sur les secteurs prioritaires'], ['Juin 2027', 'Publication', 'Rapport ESG, objectifs et plan d\'actions']];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.55, w: 11.6, h: 0, line: { color: HEX.bord, width: 4 }, objectName: 'frise' });
  road.forEach((r, i) => {
    const x = 0.55 + i * 2.5;
    s.addShape(pres.shapes.OVAL, { x: x + 0.85, y: 2.2, w: 0.7, h: 0.7, fill: { color: i === 4 ? HEX.brun : HEX.orange }, line: { color: 'FFFFFF', width: 3 }, objectName: 'jalon' });
    T(s, String(i + 1), { x: x + 0.85, y: 2.2, w: 0.7, h: 0.7, fontSize: 18, bold: true, align: 'center', valign: 'middle', color: C.background1 });
    T(s, r[0], { x, y: 1.45, w: 2.4, h: 0.5, fontSize: 14, bold: true, align: 'center', color: C.accent3 });
    T(s, r[1], { x, y: 3.15, w: 2.4, h: 0.5, fontSize: 17, bold: true, align: 'center', color: C.text2 });
    T(s, r[2], { x: x + 0.1, y: 3.7, w: 2.2, h: 1.4, fontSize: 13, align: 'center' });
  });
  box(s, 0.55, 5.45, 12.25, 1.15, HEX.tint);
  T(s, [{ text: 'Ensuite : ', options: { bold: true, color: C.accent1 } }, { text: 'outil web de collecte et de suivi, sur le même socle technique que les outils existants (SGES, notation), et intégration du calcul des émissions financées.' }],
    { x: 0.85, y: 5.45, w: 11.7, h: 1.15, fontSize: 15, valign: 'middle' });
  s.addNotes('Calendrier indicatif, à ajuster selon la date de lancement retenue et la disponibilité des données du portefeuille.');

  // 14 --------------------------------------------------------------- gouvernance
  s = content('GOUVERNANCE ET RÔLES', 'Pilotage');
  const gov = [['MdAccountBalance', 'Direction générale', 'Valide le périmètre, les objectifs et la publication'], ['MdShield', 'Direction des Risques', 'Pilote la mesure ; équipe Empreinte carbone et Climat'],
    ['MdApartment', 'Directions supports et BPR', 'Fournissent les données de l\'empreinte propre'], ['MdPayments', 'Engagements, Finance, marchés', 'Fournissent les encours et les données clients'], ['MdVerified', 'Contrôle permanent et audit', 'Revoient la piste d\'audit et la qualité']];
  for (let i = 0; i < gov.length; i++) {
    const x = 0.55 + i * 2.5;
    box(s, x, 1.5, 2.35, 4.9, i === 0 ? HEX.brun : HEX.tint2, { line: i === 0 ? HEX.brun : HEX.bord });
    await circleIcon(s, gov[i][0], x + 0.67, 1.85, 1.0, HEX.orange);
    T(s, gov[i][1], { x: x + 0.15, y: 3.1, w: 2.05, h: 0.95, fontSize: 16, bold: true, align: 'center', color: i === 0 ? C.background1 : C.text2 });
    T(s, gov[i][2], { x: x + 0.2, y: 4.1, w: 1.95, h: 2.1, fontSize: 13, align: 'center', color: i === 0 ? C.background1 : C.text1 });
  }
  s.addNotes('Chaque donnée a un propriétaire unique : c\'est ce qui supprime les sources contradictoires constatées en 2022.');

  // 15 --------------------------------------------------------------- décisions
  s = content('DÉCISIONS ATTENDUES', 'Pilotage');
  const dec = [['Valider le périmètre', 'Empreinte propre selon la méthodologie IFC, approche de consolidation (contrôle opérationnel), publication avec et hors émissions financées.'],
    ['Désigner les référents', 'Un référent par direction support et un correspondant par BPR ; lancement de la campagne selon le calendrier proposé.'],
    ['Engager le calcul PCAF', 'Accès aux données du portefeuille, choix des sources de facteurs d\'émission, adhésion éventuelle au PCAF.'],
    ['Fixer l\'ambition', 'Indicateurs publiés, objectifs de réduction de l\'empreinte propre et trajectoire pour les secteurs prioritaires.']];
  dec.forEach((d, i) => {
    const x = 0.55 + (i % 2) * 6.2, y = 1.5 + Math.floor(i / 2) * 2.6;
    box(s, x, y, 6.0, 2.35, HEX.tint);
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.35, w: 0.85, h: 0.85, fill: { color: HEX.brun }, line: { color: HEX.brun }, objectName: 'numéro' });
    T(s, String(i + 1), { x: x + 0.3, y: y + 0.35, w: 0.85, h: 0.85, fontSize: 24, bold: true, align: 'center', valign: 'middle', color: C.background1 });
    T(s, d[0], { x: x + 1.4, y: y + 0.35, w: 4.4, h: 0.5, fontSize: 19, bold: true, color: C.text2 });
    T(s, d[1], { x: x + 1.4, y: y + 0.9, w: 4.4, h: 1.35, fontSize: 14 });
  });
  s.addNotes('Ces quatre décisions conditionnent le lancement de la campagne et le premier calcul des émissions financées.');

  pres.addSection({ title: 'Clôture' });
  s = pres.addSlide({ masterName: 'TITRE', sectionTitle: 'Clôture' });
  s.addText('MERCI', { placeholder: 'title' });
  s.addText('Direction des Risques — équipe Empreinte carbone et Climat', { placeholder: 'body' });
  s.addNotes('Documents disponibles : note méthodologique de l\'empreinte propre, note méthodologique des émissions financées, canevas 2.0 et guides utilisateurs.');

  await pres.writeFile({ fileName: OUT });
  if (APPLY) { const { applyTheme } = require(APPLY); await applyTheme(OUT, THEME); }
  console.log('written', OUT);
})();
