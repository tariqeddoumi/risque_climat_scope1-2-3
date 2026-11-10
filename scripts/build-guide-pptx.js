/**
 * Guide utilisateur en présentation (PowerPoint) du canevas de collecte de l'empreinte carbone — charte Groupe BCP.
 *
 * Usage : NODE_PATH=<node_modules avec pptxgenjs> node scripts/build-guide-pptx.js <captures> <visuels> <sortie.pptx> [apply_theme.js]
 *   captures : produites par build-guide-captures.py ; visuels : produits par build-ppt-visuels.py
 */
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const [, , CAPS, VIS, OUT, APPLY] = process.argv;
const A = p => path.join(__dirname, 'assets', p);
const THEME = {
  name: 'Groupe BCP', headFontFace: 'Arial', bodyFontFace: 'Arial',
  colors: { dk1: '3D3D3B', lt1: 'FFFFFF', dk2: '462016', lt2: 'F2E8DB', accent1: 'E37B05', accent2: '462016', accent3: '8C5B52', accent4: 'B3261E', accent5: '146C4C', accent6: 'D9C7B0', hlink: 'E37B05', folHlink: '8C5B52' },
};
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'Guide utilisateur — canevas de collecte de l\'empreinte carbone';
pres.author = 'Groupe BCP — Direction des Risques';
pres.company = 'Groupe BCP';
pres.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial' };
const C = pres.SchemeColor;
const BRUN2 = C.accent3;

// ------------------------------------------------------------------ layouts
const FOOT = 'Guide utilisateur — Canevas de collecte de l\'empreinte carbone (version 2.0)';
pres.defineSlideMaster({
  title: 'TITRE', background: { path: A('groupe-bcp-fond-orange.png') },
  objects: [
    { image: { path: A('groupe-bcp-logo-blanc.png'), x: 0.7, y: 0.5, w: 1.6, h: 0.835 } },
    { image: { path: A('groupe-bcp-motif.png'), x: 10.35, y: 1.6, w: 2.6, h: 3.84 } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.55, w: 9.2, h: 1.6, fontFace: 'Arial', fontSize: 40, bold: true, color: C.background1, align: 'left', valign: 'bottom', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.3, w: 9.2, h: 1.3, fontFace: 'Arial', fontSize: 18, color: C.background1, align: 'left', valign: 'top', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'INTERCALAIRE', background: { path: A('groupe-bcp-fond-orange.png') },
  objects: [
    { image: { path: A('groupe-bcp-logo-blanc.png'), x: 0.7, y: 0.5, w: 1.6, h: 0.835 } },
    { image: { path: A('groupe-bcp-motif.png'), x: 10.35, y: 1.6, w: 2.6, h: 3.84 } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 3.0, w: 9.2, h: 1.2, fontFace: 'Arial', fontSize: 34, bold: true, color: C.background1, align: 'left', valign: 'bottom', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.35, w: 9.2, h: 1.2, fontFace: 'Arial', fontSize: 16, color: C.background1, align: 'left', valign: 'top', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'CONTENU', background: { color: 'FFFFFF' },
  objects: [
    { image: { path: A('groupe-bcp-bandeau.png'), x: 0.45, y: 0.28, w: 12.43, h: 0.85 } },
    { image: { path: A('groupe-bcp-logo-blanc.png'), x: 11.55, y: 0.36, w: 1.3, h: 0.68 } },
    { placeholder: { options: { name: 'title', type: 'title', x: 1.6, y: 0.28, w: 10.1, h: 0.85, fontFace: 'Arial', fontSize: 22, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0 }, text: '' } },
    { text: { text: FOOT, options: { x: 0.55, y: 7.02, w: 9, h: 0.3, fontFace: 'Arial', fontSize: 9, color: BRUN2, margin: 0 } } },
  ],
  slideNumber: { x: 12.25, y: 7.02, w: 0.6, h: 0.3, fontFace: 'Arial', fontSize: 9, color: BRUN2, align: 'right' },
});

// ------------------------------------------------------------------ helpers
function pngSize(file) { const b = fs.readFileSync(file); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function img(dir, name) { return path.join(dir, name + '.png'); }
function shot(slide, file, x = 0.45, y = 1.35, maxw = 8.4, maxh = 5.45) {
  const { w, h } = pngSize(file);
  let iw = maxw, ih = maxw * h / w;
  if (ih > maxh) { ih = maxh; iw = maxh * w / h; }
  const ix = x + (maxw - iw) / 2, iy = y;
  slide.addShape(pres.shapes.RECTANGLE, { x: ix - 0.04, y: iy - 0.04, w: iw + 0.08, h: ih + 0.08, fill: { color: 'FFFFFF' }, line: { color: 'D9C7B0', width: 1 },
    shadow: { type: 'outer', color: '7A5A40', opacity: 0.25, blur: 6, offset: 2, angle: 90 }, objectName: 'cadre capture' });
  slide.addImage({ path: file, x: ix, y: iy, w: iw, h: ih, altText: path.basename(file), objectName: 'capture' });
  return { x: ix, y: iy, w: iw, h: ih };
}
function panel(slide, steps, savoir, o = {}) {
  const x = o.x ?? 9.15, w = o.w ?? 3.75;
  slide.addText(o.l1 ?? 'ÉTAPES', { x, y: 1.35, w, h: 0.4, fontSize: 16, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  slide.addText(steps.map((t, i) => ({ text: t, options: { bullet: { type: 'number' }, breakLine: i < steps.length - 1, paraSpaceAfter: 6 } })),
    { x, y: 1.8, w, h: o.h1 ?? 2.75, fontSize: o.fs ?? 13, color: C.text1, valign: 'top', margin: 0, isTextBox: true });
  if (savoir) {
    const y2 = o.y2 ?? 4.7;
    slide.addText(o.l2 ?? 'À SAVOIR', { x, y: y2, w, h: 0.4, fontSize: 16, bold: true, color: BRUN2, margin: 0, isTextBox: true });
    slide.addText(savoir, { x, y: y2 + 0.45, w, h: 6.85 - (y2 + 0.45), fontSize: o.fs2 ?? 12, color: C.text1, valign: 'top', margin: 0, isTextBox: true });
  }
}
// Capture large (rapport largeur/hauteur > 2,3) : image sur toute la largeur, explications en deux colonnes dessous.
function screen(title, section, file, steps, savoir, o = {}) {
  const sl = content(title, section);
  const { w, h } = pngSize(file);
  if (w / h > 2.3 && !o.side) {
    const iw = 12.43, ih = Math.min(iw * h / w, o.maxh ?? 3.0);
    const r = shot(sl, file, 0.45, 1.35, iw, ih);
    const y = r.y + r.h + 0.35;
    sl.addText(o.l1 ?? 'ÉTAPES', { x: 0.55, y, w: 6.0, h: 0.38, fontSize: 16, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    sl.addText(steps.map((t, i) => ({ text: t, options: { bullet: { type: 'number' }, breakLine: i < steps.length - 1, paraSpaceAfter: 4 } })),
      { x: 0.55, y: y + 0.45, w: 6.0, h: 6.85 - (y + 0.45), fontSize: 14, color: C.text1, valign: 'top', margin: 0, isTextBox: true });
    if (savoir) {
      sl.addText(o.l2 ?? 'À SAVOIR', { x: 6.95, y, w: 5.9, h: 0.38, fontSize: 16, bold: true, color: BRUN2, margin: 0, isTextBox: true });
      sl.addText(savoir, { x: 6.95, y: y + 0.45, w: 5.9, h: 6.85 - (y + 0.45), fontSize: 14, color: C.text1, valign: 'top', margin: 0, isTextBox: true });
    }
  } else {
    shot(sl, file, 0.45, 1.35, 8.4, 5.45);
    panel(sl, steps, savoir, o);
  }
  return sl;
}
function content(title, section) {
  const s = pres.addSlide({ masterName: 'CONTENU', sectionTitle: section });
  s.addText(title, { placeholder: 'title' });
  return s;
}
function divider(title, body, section) {
  pres.addSection({ title: section });
  const s = pres.addSlide({ masterName: 'INTERCALAIRE', sectionTitle: section });
  s.addText(title, { placeholder: 'title' });
  s.addText(body, { placeholder: 'body' });
  return s;
}
function card(slide, x, y, w, h, head, text, o = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: o.fill ?? 'F2E8DB' }, line: { color: o.fill ?? 'F2E8DB', width: 0.5 }, objectName: 'carte' });
  const ic = o.icon ?? null;
  const tx = ic ? x + 0.95 : x + 0.25;
  if (ic) {
    slide.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: o.iconColor ?? 'E37B05' }, line: { color: o.iconColor ?? 'E37B05' }, objectName: 'pastille' });
    slide.addText(ic, { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fontSize: o.iconSize ?? 18, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  }
  slide.addText(head, { x: tx, y: y + 0.2, w: x + w - tx - 0.2, h: o.hh ?? 0.5, fontSize: o.hs ?? 14, bold: true, color: C.text2, valign: 'top', margin: 0, isTextBox: true });
  slide.addText(text, { x: tx, y: y + 0.2 + (o.hh ?? 0.5) + 0.05, w: x + w - tx - 0.2, h: h - 0.45 - (o.hh ?? 0.5), fontSize: o.ts ?? 12, color: C.text1, valign: 'top', margin: 0, isTextBox: true });
}
function qa(title, section, items, notes) {
  const s = content(title, section);
  const W = 6.05, H = 2.65, xs = [0.55, 6.73], ys = [1.4, 4.2];
  items.forEach((it, i) => card(s, xs[i % 2], ys[Math.floor(i / 2)], W, H, it[0], it[1], { icon: '?', hh: 0.62, hs: 15, ts: 13.5 }));
  s.addNotes(notes);
  return s;
}

// ------------------------------------------------------------------ 1. titre
pres.addSection({ title: 'Introduction' });
let s = pres.addSlide({ masterName: 'TITRE', sectionTitle: 'Introduction' });
s.addText('GUIDE UTILISATEUR\nCANEVAS DE COLLECTE DE L\'EMPREINTE CARBONE', { placeholder: 'title' });
s.addText('Pas à pas, avec les réponses aux questions et aux blocages les plus fréquents — canevas version 2.0, novembre 2026', { placeholder: 'body' });
s.addNotes('Guide de prise en main destiné aux contributeurs des sites, aux correspondants BPR et aux référents des directions centrales. Chaque diapositive montre le canevas réel ; les valeurs affichées sont des exemples.');

s = content('LE PARCOURS EN 5 ÉTAPES', 'Introduction');
const steps5 = [['1', 'S\'identifier', 'Onglet Identification : site, entité, exercice, surface, effectif'], ['2', 'Lire les cas', 'Dans chaque onglet : consignes, cas possibles, erreurs à éviter'],
  ['3', 'Saisir', 'Cellules orangées, listes déroulantes, une ligne par cas'], ['4', 'Donner un statut', 'Valeur, Néant, Non disponible ou Non applicable'], ['5', 'Contrôler et envoyer', 'Onglet Contrôle : 0 alerte, puis envoi avec les pièces']];
steps5.forEach((st, i) => {
  const x = 0.55 + i * 2.5;
  s.addShape(pres.shapes.CHEVRON, { x, y: 1.9, w: 2.35, h: 1.1, fill: { color: i % 2 ? '462016' : 'E37B05' }, line: { color: 'FFFFFF', width: 0 }, objectName: 'étape' });
  s.addText(st[0], { x: x + 0.3, y: 1.9, w: 1.75, h: 1.1, fontSize: 30, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  s.addText(st[1], { x, y: 3.25, w: 2.35, h: 0.45, fontSize: 15, bold: true, color: C.text2, align: 'center', margin: 0, isTextBox: true });
  s.addText(st[2], { x: x + 0.1, y: 3.75, w: 2.15, h: 1.2, fontSize: 12, color: C.text1, align: 'center', valign: 'top', margin: 0, isTextBox: true });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55, y: 5.4, w: 12.25, h: 1.25, rectRadius: 0.08, fill: { color: 'F2E8DB' }, line: { color: 'F2E8DB' }, objectName: 'message' });
s.addText([{ text: 'Vous n\'avez aucun calcul à faire. ', options: { bold: true, color: C.accent1 } }, { text: 'Donnez des quantités (kWh, litres, kg, nombre d\'équipements…) dans l\'unité demandée : l\'équipe Empreinte carbone applique ensuite les facteurs d\'émission.' }],
  { x: 0.85, y: 5.5, w: 11.7, h: 1.05, fontSize: 15, color: C.text1, valign: 'middle', margin: 0, isTextBox: true });
s.addNotes('Le canevas recueille des chiffres, pas des émissions. Les cinq étapes se répètent dans chaque onglet ; l\'onglet Contrôle dit quand le fichier est prêt.');

s = content('LES 14 ONGLETS DU CANEVAS', 'Introduction');
const tabs = [['Lisez-moi', 'Mode d\'emploi intégré', '462016'], ['Identification', 'Site, exercice, effectif', '462016'], ['1 Électricité', 'kWh par compteur', 'E37B05'], ['2 Carburants et gaz', 'Litres, kg, bouteilles', 'E37B05'],
  ['3 Climatisation', 'Fluides rechargés', 'E37B05'], ['4 Achats', 'Montants HT par famille', 'E37B05'], ['5 Équipements', 'Quantité et année d\'achat', 'E37B05'], ['6 Déchets', 'Quantité et traitement', 'E37B05'],
  ['7 Déplacements pro', 'Vols, trains, hôtels', 'E37B05'], ['8 Domicile et visiteurs', 'Modes, distances', 'E37B05'], ['9 Transport fonds-courrier', 'Passages, envois', 'E37B05'], ['Contrôle', 'Alertes et cohérence', 'B3261E'],
  ['Export IFC', 'Format du canevas IFC', '146C4C'], ['Listes', 'Valeurs des listes', '8A7A6A']];
tabs.forEach((t, i) => {
  const col = i % 7, row = Math.floor(i / 7);
  const x = 0.55 + col * 1.77, y = 1.55 + row * 2.45;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.62, h: 2.1, rectRadius: 0.08, fill: { color: 'FBF6F0' }, line: { color: t[2], width: 1.5 }, objectName: 'onglet' });
  s.addShape(pres.shapes.RECTANGLE, { x: x + 0.12, y: y + 0.15, w: 1.38, h: 0.12, fill: { color: t[2] }, line: { color: t[2] }, objectName: 'couleur onglet' });
  s.addText(t[0], { x: x + 0.1, y: y + 0.4, w: 1.42, h: 0.85, fontSize: 13, bold: true, color: C.text2, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  s.addText(t[1], { x: x + 0.1, y: y + 1.25, w: 1.42, h: 0.7, fontSize: 11, color: C.text1, align: 'center', valign: 'top', margin: 0, isTextBox: true });
});
s.addText('Vous remplissez Identification et les onglets 1 à 9 qui vous concernent. Contrôle et Export IFC se calculent seuls.', { x: 0.55, y: 6.5, w: 12.25, h: 0.4, fontSize: 13, italic: true, color: BRUN2, margin: 0, isTextBox: true });
s.addNotes('La couleur du trait correspond à la couleur des onglets dans Excel : brun pour l\'accueil, orange pour la saisie, rouge pour le contrôle, vert pour l\'export IFC, gris pour les listes.');

s = content('LES 6 RÈGLES D\'OR', 'Introduction');
const rules = [['1', 'Un nombre seul, sans unité', 'L\'unité se choisit dans sa colonne. ✓ 35555 — ✗ « 35555 kWh »'], ['2', 'La quantité avant le montant', 'Les émissions dépendent des litres ou des kWh. Le montant n\'est qu\'un repli.'],
  ['3', 'Des montants hors taxes', 'Si vous n\'avez que le TTC, divisez par 1,2 et dites-le en commentaire.'], ['4', 'Un statut sur chaque ligne', 'Une case vide ne dit pas si l\'activité n\'existe pas ou si l\'information manque.'],
  ['5', 'Ne compter qu\'une fois', 'Une donnée déclarée par le siège n\'est pas ressaisie par le site : « Non applicable ».'], ['6', 'Aucun calcul dans les cellules', 'Donnez la donnée brute ; expliquez vos hypothèses en commentaire. ✗ « =955*113 »']];
rules.forEach((r, i) => card(s, 0.55 + (i % 3) * 4.12, 1.45 + Math.floor(i / 3) * 2.75, 3.95, 2.5, r[1], r[2], { icon: r[0], hs: 15, ts: 13, hh: 0.62 }));
s.addNotes('Ces six règles évitent l\'essentiel des allers-retours constatés lors de la collecte 2022 : unités saisies en texte, montants TTC, cases vides, doubles comptages et extrapolations cachées.');

// ------------------------------------------------------------------ 2. Prendre en main
const S1 = '1. Prendre en main';
divider('1. PRENDRE EN MAIN LE CANEVAS', 'Ouvrir le fichier, remplir l\'identification et comprendre l\'organisation d\'un onglet de saisie.', S1);
s = screen('OUVRIR ET PRÉPARER LE FICHIER', S1, img(CAPS, '01_lisez_moi_etapes_regles'), ['Enregistrer le fichier sur votre poste avant de le remplir.', 'Cliquer sur « Activer la modification » si Excel affiche le bandeau jaune « Mode protégé ».', 'Renommer le fichier : EC_<exercice>_<code site>_<nom du site>.xlsx.', 'Lire l\'onglet Lisez-moi (2 minutes).'],
  'Sans « Activer la modification », les listes déroulantes et les contrôles ne fonctionnent pas. Ne renommez pas les onglets : l\'onglet Contrôle s\'appuie sur eux.');
s.addNotes('Exemple de nom : EC_2025_4896_Agence_Al_Boughaz.xlsx. Le fichier reçu par courriel s\'ouvre en lecture seule tant qu\'il n\'est pas enregistré.');
s = screen('L\'ONGLET IDENTIFICATION', S1, img(CAPS, '05_identification_remplie'), ['Saisir l\'exercice : l\'année des données (2025), pas l\'année d\'envoi.', 'Choisir l\'entité et le type de site dans les listes.', 'Indiquer le code et le nom du site, la surface et l\'effectif.', 'Facultatif : surface louée à des tiers, consommation d\'eau.'],
  'La surface et l\'effectif servent aux contrôles de cohérence (kWh par m², nombre d\'employés). Indiquez votre fonction, sans nom ni adresse électronique.');
s.addNotes('L\'onglet Identification se remplit une seule fois, en premier. S\'il est incomplet, l\'onglet Contrôle l\'indique.');
s = screen('UN ONGLET DE SAISIE EN 5 ZONES', S1, img(CAPS, '06_anatomie_onglet'), ['Bandeau : poste, qui remplit, où trouver l\'information.', 'Comment remplir : la règle de l\'onglet en quelques phrases.', 'Cas possibles : une ligne d\'exemple grise par situation.', 'Erreurs à éviter, sur fond rose.', 'Votre saisie : 100 lignes orangées et le contrôle automatique.'],
  'Les neuf onglets de saisie ont la même organisation : une fois l\'onglet Électricité compris, tous les autres se remplissent de la même façon.', { fs: 12.5, h1: 2.9, y2: 4.85 });
s.addNotes('Les repères 1 à 5 de la capture correspondent aux cinq zones listées à droite.');

// ------------------------------------------------------------------ 3. Saisir
const S2 = '2. Saisir pas à pas';
divider('2. SAISIR PAS À PAS', 'Lire les cas possibles, utiliser les listes, saisir un nombre, choisir le statut et la méthode.', S2);
s = screen('PARTIR DU CAS QUI VOUS RESSEMBLE', S2, img(CAPS, '08_electricite_cas_possibles'), ['Parcourir les cas possibles de l\'onglet.', 'Repérer celui qui décrit votre situation (bandeau beige).', 'Reproduire sa ligne grise dans la zone orangée, avec vos chiffres.'],
  'Les lignes grises affichent « OK » dans la colonne de contrôle : c\'est le résultat attendu sur vos propres lignes. Ne les modifiez pas et ne les recopiez pas telles quelles.');
s.addNotes('Électricité : factures mensuelles, total annuel seul, montant seul, sous-compteur, électricité payée par le bailleur, panneaux solaires, data center, réseau de chaleur.');
s = screen('LES COLONNES DE LA ZONE DE SAISIE', S2, img(CAPS, '09_electricite_zone_saisie'), ['Lire la ligne d\'aide (italique) sous chaque en-tête.', 'Laisser « Code site » vide si le fichier ne concerne qu\'un site.', 'Remplir les colonnes propres au poste (compteur, usage, mois, quantité…).', 'Terminer par Statut, Méthode, Pièce justificative, Commentaire.'],
  'Passez la souris sur un en-tête : une bulle d\'aide s\'affiche. La dernière colonne, « Contrôle automatique », ne se remplit pas : elle se calcule seule.');
s.addNotes('Les quatre dernières colonnes de saisie et la colonne de contrôle sont les mêmes dans tous les onglets.');
s = screen('CHOISIR DANS UNE LISTE DÉROULANTE', S2, img(CAPS, '12_liste_deroulante'), ['Cliquer dans la cellule orangée.', 'Cliquer sur la petite flèche à droite de la cellule.', 'Cliquer sur la valeur voulue (ou la taper exactement).'],
  'Une valeur hors liste est refusée. Si rien ne convient, choisissez « Autre (préciser) » quand elle existe et expliquez en commentaire.');
s.addNotes('Les listes évitent le texte libre qui rendait la consolidation impossible en 2022 (« Amesterdam », « MAROC-Dubai abu dhabi »).');
s = content('CHOISIR LE BON STATUT', S2);
const qs = [['Ai-je un chiffre, même égal à 0 ?', 'VALEUR', '2 850 kWh ; recharge de fluide = 0 kg selon le rapport du mainteneur', 'E37B05'],
  ['L\'activité n\'existe pas sur mon site ?', 'NÉANT', 'Pas de groupe électrogène, pas de panneaux solaires', '462016'],
  ['La donnée est déclarée par quelqu\'un d\'autre ?', 'NON APPLICABLE', 'Gardiennage payé et déclaré par la Direction des Achats', '8C5B52'],
  ['L\'activité existe mais je n\'ai pas le chiffre ?', 'NON DISPONIBLE', 'Électricité incluse dans les charges du bailleur ; indiquer qui détient la donnée', 'B3261E']];
qs.forEach((q, i) => {
  const y = 1.45 + i * 1.25;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55, y, w: 4.9, h: 1.05, rectRadius: 0.08, fill: { color: 'F2E8DB' }, line: { color: 'F2E8DB' }, objectName: 'question' });
  s.addText(`${i + 1}. ${q[0]}`, { x: 0.75, y, w: 4.6, h: 1.05, fontSize: 15, bold: true, color: C.text2, valign: 'middle', margin: 0, isTextBox: true });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 5.6, y: y + 0.32, w: 0.7, h: 0.42, fill: { color: 'D9C7B0' }, line: { color: 'D9C7B0' }, objectName: 'flèche' });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.45, y, w: 2.5, h: 1.05, rectRadius: 0.08, fill: { color: q[3] }, line: { color: q[3] }, objectName: 'statut' });
  s.addText(q[1], { x: 6.45, y, w: 2.5, h: 1.05, fontSize: 15, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  s.addText(q[2], { x: 9.15, y, w: 3.7, h: 1.05, fontSize: 12, color: C.text1, valign: 'middle', margin: 0, isTextBox: true });
});
s.addText('Posez-vous les questions dans l\'ordre : la première réponse « oui » donne le statut. Avec Néant, Non disponible ou Non applicable, laissez la quantité vide.', { x: 0.55, y: 6.45, w: 12.25, h: 0.45, fontSize: 13, italic: true, color: BRUN2, margin: 0, isTextBox: true });
s.addNotes('Le zéro n\'est admis qu\'avec le statut Valeur, quand il est vrai et justifié. Un 0 avec Néant déclenche l\'alerte « Statut incohérent avec le chiffre saisi ».');
s = screen('LES NOUVEAUTÉS DE LA VERSION 2', S2, img(CAPS, '03b_lisez_moi_nouveautes'), ['Colonne « Code site » sur chaque ligne.', 'Visiteurs et clients (onglet 8), loyers, papier recyclé, matériaux.', 'Véhicules électriques, contrat d\'électricité, fluide récupéré, zone des vols.', 'Contrôles de cohérence et onglet Export IFC.'],
  'Les feuilles sont protégées : seules les cellules orangées sont modifiables. Chaque onglet offre 100 lignes.', { fs: 12.5 });
s.addNotes('Ces nouveautés comblent les écarts avec le canevas IFC de la mesure 2022 et ajoutent les pratiques de reporting des banques.');

// ------------------------------------------------------------------ 4. Contrôler
const S3 = '3. Contrôler avant l\'envoi';
divider('3. CONTRÔLER AVANT L\'ENVOI', 'Lire le contrôle automatique de chaque ligne, le bilan de l\'onglet Contrôle, les contrôles de cohérence et l\'export IFC.', S3);
s = screen('LE CONTRÔLE AUTOMATIQUE DE CHAQUE LIGNE', S3, img(CAPS, '10_electricite_saisie_alertes'), ['Statut vide → choisir « Valeur ».', '« 2 480 kWh » est du texte → retaper 2480.', 'Valeur sans chiffre → saisir les kWh, ou choisir « Néant ».', '0 avec « Néant » → effacer le 0.'],
  'Corrigez chaque ligne tant que la dernière colonne n\'affiche pas « OK ». Le contrôle vérifie la forme de la ligne, pas l\'exactitude du chiffre.', { l1: 'LES 4 ALERTES DE L\'EXEMPLE' });
s.addNotes('Les repères 1 à 4 de la capture correspondent aux quatre corrections listées à droite.');
s = screen('APRÈS CORRECTION : TOUT EST « OK »', S3, img(CAPS, '11_electricite_saisie_corrigee'), ['Vérifier la dernière colonne de chaque ligne saisie.', 'Passer à l\'onglet suivant.', 'Ouvrir l\'onglet Contrôle pour le bilan.'], 'Une ligne non commencée reste vide dans la colonne de contrôle : ce n\'est pas une erreur.');
s.addNotes('Les mêmes lignes qu\'à la diapositive précédente, après les quatre corrections.');
s = screen('L\'ONGLET CONTRÔLE', S3, img(CAPS, '30_controle_alertes'), ['Lire, pour chaque onglet, le nombre de lignes saisies, OK et en alerte.', 'Retourner dans l\'onglet qui a des alertes et les corriger.', 'Envoyer quand l\'état du fichier indique « Prêt à envoyer ».'], 'L\'onglet Contrôle vérifie aussi que l\'onglet Identification est complet.');
s.addNotes('Ici, 4 alertes restent à corriger dans l\'onglet Électricité.');
s = screen('LES CONTRÔLES DE COHÉRENCE', S3, img(CAPS, '32_controle_coherence'), ['12 mois par compteur.', 'kWh par m² entre 20 et 500.', 'Employés déclarés proches de l\'effectif.', 'Vols avec une zone, carburants avec un prix, électricité avec un contrat.'],
  'Ces contrôles ne bloquent pas l\'envoi : corrigez la donnée, ou expliquez en commentaire une situation atypique (data center, agence partagée).', { l1: 'CE QUI EST VÉRIFIÉ' });
s.addNotes('Les bornes de 20 à 500 kWh par m² sont larges : elles repèrent surtout les erreurs d\'unité ou de surface.');
s = screen('L\'ONGLET EXPORT IFC', S3, img(CAPS, '33_export_ifc'), ['Ne rien saisir (sauf, si besoin, les jours ouvrés).', 'Parcourir les totaux par rubrique IFC.', 'Repérer un total anormal avant l\'envoi.'],
  'L\'export recompose vos saisies au format du canevas IFC de la mesure 2022 : bouteilles converties en kg, distance aller en aller-retour, année d\'achat en âge.');
s.addNotes('Seules les lignes au statut Valeur sont prises en compte. L\'équipe Empreinte carbone utilise cet onglet pour consolider les fichiers de tous les sites.');

// ------------------------------------------------------------------ 5. Blocages
const S4 = '4. Blocages et questions fréquents';
divider('4. BLOCAGES ET QUESTIONS FRÉQUENTS', 'Les messages d\'Excel qui bloquent la saisie, et les réponses aux questions posées le plus souvent.', S4);
s = screen('BLOCAGE : « VALEUR NON AUTORISÉE »', S4, img(VIS, 'v_liste_refusee'), ['Cliquer sur « Annuler ».', 'Cliquer sur la flèche de la cellule et choisir une valeur dans la liste.', 'Si aucune ne convient : « Autre (préciser) » et une explication en commentaire.'],
  'La valeur tapée n\'existe pas dans la liste de la colonne (faute de frappe, libellé différent, valeur collée depuis un autre fichier).', { l1: 'LA SOLUTION', l2: 'POURQUOI', fs: 13 });
s.addNotes('Message affiché quand on tape ou colle une valeur absente de la liste déroulante. Illustration : boîte de dialogue reconstituée avec le texte paramétré dans le canevas.');
s = screen('BLOCAGE : « NOMBRE ATTENDU »', S4, img(VIS, 'v_nombre_attendu'), ['Cliquer sur « Recommencer ».', 'Taper le nombre seul : 2850, sans espace ni unité.', 'Utiliser la virgule pour les décimales dans un Excel en français : 12,5.', 'Pour une année : 4 chiffres (2021).'],
  'La cellule attend un nombre positif. Une unité (« kWh »), un espace (« 2 850 ») ou un nombre négatif sont refusés. Un texte collé passe parfois : le contrôle affiche alors « Nombre attendu, sans unité ».', { l1: 'LA SOLUTION', l2: 'POURQUOI', fs: 13 });
s.addNotes('Le même principe s\'applique aux années (4 chiffres) et aux jours de présence par semaine (entre 0 et 7).');
s = screen('BLOCAGE : « FEUILLE PROTÉGÉE »', S4, img(VIS, 'v_feuille_protegee'), ['Cliquer sur « OK ».', 'Saisir uniquement dans les cellules orangées de la zone « Votre saisie ».', 'Besoin de plus de 100 lignes : Révision › Ôter la protection de la feuille (sans mot de passe), insérer les lignes au milieu de la zone, puis remettre la protection.'],
  'Les consignes, exemples, en-têtes et formules de contrôle sont protégés pour éviter qu\'on les efface par erreur.', { l1: 'LA SOLUTION', l2: 'POURQUOI', fs: 12.5, h1: 3.0, y2: 4.95 });
s.addNotes('Aucun mot de passe n\'est demandé pour ôter la protection. Pensez à la remettre après avoir ajouté des lignes.');
s = content('BLOCAGE : J\'AI COLLÉ ET LES LISTES ONT DISPARU', S4);
const keys = [['Ctrl', 0.9], ['Alt', 2.45], ['V', 4.0]];
s.addText('Coller les valeurs uniquement', { x: 0.55, y: 1.45, w: 8.3, h: 0.5, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
keys.forEach(([k, x]) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.2, w: 1.3, h: 1.1, rectRadius: 0.12, fill: { color: 'FFFFFF' }, line: { color: '8C5B52', width: 2 }, shadow: { type: 'outer', color: '7A5A40', opacity: 0.35, blur: 2, offset: 3, angle: 90 }, objectName: 'touche' });
  s.addText(k, { x, y: 2.2, w: 1.3, h: 1.1, fontSize: 24, bold: true, color: C.text2, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
});
['+', '+'].forEach((p, i) => s.addText(p, { x: 2.2 + i * 1.55, y: 2.2, w: 0.25, h: 1.1, fontSize: 24, bold: true, color: C.accent1, align: 'center', valign: 'middle', margin: 0, isTextBox: true }));
s.addText('puis V (Valeurs) et Entrée', { x: 5.55, y: 2.2, w: 3.3, h: 1.1, fontSize: 16, color: C.text1, valign: 'middle', margin: 0, isTextBox: true });
card(s, 0.55, 3.75, 8.3, 2.1, 'Ou avec la souris', 'Clic droit dans la cellule de destination › Options de collage › icône « Valeurs » (123). Le format orangé, la liste déroulante et le contrôle de la cellule sont conservés.', { icon: 'i', iconSize: 18, hs: 15, ts: 14 });
panel(s, ['Annuler le collage (Ctrl+Z).', 'Recoller avec « Coller les valeurs ».', 'Vérifier la colonne de contrôle.'],
  'Un collage normal remplace la cellule entière : sa couleur, sa liste déroulante et sa règle de saisie. Il peut aussi coller du texte à la place d\'un nombre.', { l1: 'LA SOLUTION', l2: 'POURQUOI' });
s.addNotes('Ce blocage arrive surtout quand on recopie une extraction ERP ou un tableau de suivi. Si les listes ont été écrasées sur beaucoup de lignes, repartir d\'un canevas vierge est souvent le plus rapide.');
s = screen('BLOCAGE : « STATUT INCOHÉRENT » ET LE ZÉRO', S4, img(CAPS, '10_electricite_saisie_alertes'), ['Activité inexistante : statut « Néant » et quantité vide.', 'Chiffre réellement nul (aucune recharge de fluide) : statut « Valeur » et 0.', 'Information manquante : « Non disponible », quantité vide, commentaire.'],
  'L\'alerte apparaît quand un chiffre, même 0, est saisi avec Néant, Non disponible ou Non applicable (ligne 6 de l\'exemple).', { l1: 'LA SOLUTION', l2: 'POURQUOI' });
s.addNotes('Le zéro est une information : il dit que l\'activité existe et vaut zéro. Le statut Néant dit que l\'activité n\'existe pas.');
qa('FAQ : DONNÉES MANQUANTES', S4, [
  ['Je n\'ai pas l\'information demandée.', 'Statut « Non disponible », quantité vide, et en commentaire qui pourrait la détenir (« factures chez la Direction du Patrimoine »).'],
  ['L\'activité n\'existe pas chez nous.', 'Une ligne au statut « Néant (aucune activité) », sans quantité. Exemple : pas de groupe électrogène.'],
  ['La donnée est déjà fournie par le siège.', 'Ne la ressaisissez pas : « Non applicable » et la direction qui la déclare en commentaire. Sinon elle serait comptée deux fois.'],
  ['Je n\'ai que le montant, ou que le TTC.', 'Montant HT, quantité vide, méthode « Estimé à partir d\'un montant ». TTC ÷ 1,2 (TVA à 20 %) à signaler en commentaire. Carburants : ajouter le prix unitaire moyen.'],
], 'Les quatre situations les plus fréquentes lors de la collecte 2022.');
qa('FAQ : ÉLECTRICITÉ ET CARBURANTS', S4, [
  ['Ma facture est bimestrielle, ou il me manque des mois.', 'Saisir chaque facture sur le mois de fin de période, avec la période en commentaire ; ou une ligne « Année entière ». Le contrôle « 12 mois par compteur » signale les manques.'],
  ['kW ou kWh ?', 'Toujours les kWh consommés (« énergie active »). Les kW ou kVA sont la puissance souscrite, pas une consommation.'],
  ['Un compteur alimente aussi un locataire ou la salle informatique.', 'Le signaler dans « Usage du compteur » : une ligne pour le compteur principal, une pour le sous-compteur. Cela évite le double comptage.'],
  ['Des bouteilles de gaz de 13 kg ?', 'Au Maroc, les formats sont 12 kg et 3 kg : choisir l\'unité « Bouteille 12 kg » ou « Bouteille 3 kg » et donner le nombre de bouteilles.'],
], 'La chaleur et le froid produits par des climatiseurs électriques sont déjà dans la facture d\'électricité : ne pas les déclarer à part.');
qa('FAQ : CLIMATISATION, ÉQUIPEMENTS, DÉCHETS', S4, [
  ['Je n\'ai que la charge du climatiseur.', 'La charge (plaque) n\'est pas une fuite. Ce qui compte, c\'est la quantité rechargée dans l\'année, sur le rapport du mainteneur. Aucune recharge : 0 avec le statut « Valeur ».'],
  ['Je ne connais que l\'âge d\'un équipement.', 'Année d\'achat = exercice − âge (2025 − 3 = 2022). Sans aucune date : estimer l\'année avec la méthode « Estimation d\'expert ». Pas de tranche (« plus de 4 ans »).'],
  ['Mon équipement n\'est pas dans la liste.', 'Catégorie « Autre équipement (préciser) » et son nom dans « Modèle / référence ».'],
  ['Les déchets ne sont pas pesés.', 'Donner le nombre de sacs ou de bacs collectés dans l\'année, ou une pesée d\'une semaine. Collecte communale sans suivi : « Non disponible ».'],
], 'Ces trois postes concentrent la plupart des données estimées : la méthode choisie permet à l\'équipe d\'attribuer un niveau de qualité.');
qa('FAQ : DÉPLACEMENTS, VISITEURS, SITES', S4, [
  ['Un aller-retour, une ou deux lignes ?', 'Une seule ligne « Aller-retour ». Un vol avec escale : une ligne par tronçon. Toujours indiquer la zone Maroc ou International.'],
  ['Ma ville n\'est pas dans la liste.', 'Choisir « Autre (préciser en commentaire) » et écrire la ville en commentaire.'],
  ['Comment déclarer les visiteurs de l\'agence ?', 'Onglet 8, catégorie « Visiteurs et clients » : une ligne par mode, le nombre de visites dans l\'année et la distance aller moyenne. Une estimation suffit.'],
  ['Je réponds pour plusieurs sites.', 'De préférence un fichier par site. Sinon, l\'indiquer dans Identification et remplir la colonne « Code site » sur chaque ligne.'],
], 'Les visiteurs sont demandés par la méthodologie IFC ; une estimation à partir de la fréquentation moyenne de l\'agence suffit.');

// ------------------------------------------------------------------ 6. Envoyer
const S5 = '5. Envoyer';
divider('5. VÉRIFIER ET ENVOYER', 'La liste de vérification, le nom du fichier et les pièces justificatives.', S5);
s = content('LISTE DE VÉRIFICATION AVANT ENVOI', S5);
shot(s, img(CAPS, '31_controle_pret'), 0.45, 1.35, 7.2, 5.45);
const checks = ['Identification complète : « Oui »', 'État du fichier : « Prêt à envoyer »', 'Une ligne au moins dans chaque onglet qui vous concerne', 'Montants hors taxes', 'Contrôles de cohérence « OK » ou expliqués', 'Pièces justificatives jointes'];
checks.forEach((c, i) => {
  const y = 1.45 + i * 0.85;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.0, y, w: 0.5, h: 0.5, rectRadius: 0.06, fill: { color: '146C4C' }, line: { color: '146C4C' }, objectName: 'case cochée' });
  s.addText('✓', { x: 8.0, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
  s.addText(c, { x: 8.7, y: y - 0.05, w: 4.2, h: 0.6, fontSize: 14, color: C.text1, valign: 'middle', margin: 0, isTextBox: true });
});
s.addNotes('Un écart de plus de 20 % avec l\'exercice précédent doit être expliqué en commentaire.');
s = content('NOMMER ET ENVOYER LE FICHIER', S5);
const send = [['1', 'Nommer', 'EC_2025_4896_Agence_Al_Boughaz.xlsx'], ['2', 'Regrouper les pièces', 'Dossier compressé du même nom : EC_2025_4896_Agence_Al_Boughaz.zip'], ['3', 'Envoyer', 'Au correspondant BPR (sites des BPR) ou à l\'équipe Empreinte carbone (directions centrales), avant la date du courriel de lancement'], ['4', 'Rester disponible', 'Les questions de l\'équipe citeront l\'onglet et le numéro de ligne']];
send.forEach((st, i) => card(s, 0.55 + i * 3.1, 1.6, 2.9, 3.6, st[1], st[2], { icon: st[0], hs: 16, ts: 13, hh: 0.8 }));
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.55, y: 5.55, w: 12.25, h: 1.1, rectRadius: 0.08, fill: { color: 'E37B05' }, line: { color: 'E37B05' }, objectName: 'contact' });
s.addText('Une question ? Votre correspondant BPR, ou l\'équipe Empreinte carbone (Direction des Risques), en précisant l\'onglet et le numéro de ligne.', { x: 0.85, y: 5.55, w: 11.7, h: 1.1, fontSize: 15, bold: true, color: C.background1, valign: 'middle', margin: 0, isTextBox: true });
s.addNotes('Le nom du fichier permet à l\'équipe de rattacher automatiquement le fichier au site et à l\'exercice.');

pres.addSection({ title: 'Conclusion' });
s = pres.addSlide({ masterName: 'TITRE', sectionTitle: 'Conclusion' });
s.addText('MERCI POUR VOTRE CONTRIBUTION', { placeholder: 'title' });
s.addText('Chaque donnée juste et justifiée rend la mesure de l\'empreinte carbone du Groupe plus fiable, et comparable d\'une année à l\'autre.', { placeholder: 'body' });
s.addNotes('Le guide utilisateur Word reprend chaque onglet en détail, avec ses cas possibles et ses colonnes.');

pres.writeFile({ fileName: OUT }).then(async () => {
  if (APPLY) { const { applyTheme } = require(APPLY); await applyTheme(OUT, THEME); }
  console.log('written', OUT);
});
