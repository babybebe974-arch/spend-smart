/**
 * 💸 SPEND SMART — Dépenses de Smartdesignforyou, tous moyens de paiement (montants en Rs)
 * ------------------------------------------------------------
 * - Lit les reçus (scan PDF, photo, capture d'écran d'un reçu numérique ou
 *   d'un abonnement) déposés dans le dossier Drive « Spend Smart - À traiter ».
 * - Remplit l'onglet « Dépenses » + l'onglet « Récap » + « Historique ».
 * - Détecte les doublons (ligne ROUGE) et les lectures douteuses (ligne JAUNE).
 * - Dépenses SANS reçu : formulaire (menu ou raccourci Android), ligne ORANGE.
 * - Envoi automatique au comptable X jours avant l'échéance de la déclaration
 *   TVA MRA (Excel + reçus en .zip).
 *
 * Quand le traitement se fait :
 *   - Menu 💸 Spend Smart → Traiter les reçus maintenant
 *   - Raccourci Android (application Web) → bouton « Traiter les reçus »
 *   - Automatiquement chaque dimanche à 20h (récap envoyé par e-mail)
 * ------------------------------------------------------------
 */

// ====== RÉGLAGES (à adapter) ======
const CONFIG = {
  EMAIL_BUREAU: '',                    // e-mail du comptable. Vide = envoi à toi seulement
  PERIODICITE_TVA: 'TRIMESTRIELLE',    // 'TRIMESTRIELLE' (CA ≤ Rs 10 M) ou 'MENSUELLE'
  JOURS_AVANT_TVA: 7,                  // envoi au comptable X jours avant l'échéance
  SIGNATURE: 'Réaz Moussa – Smartdesignforyou',

  MODELE: 'gemini-flash-latest',       // si Google retire le modèle, le script en choisit un autre seul
  NOM_ONGLET: 'Dépenses',
  DOSSIER_A_TRAITER: 'Spend Smart - À traiter',
  DOSSIER_TRAITEES: 'Spend Smart - Traitées',
  DOSSIER_ENVOIS: 'Spend Smart - Récaps',
  JOUR_AUTO: ScriptApp.WeekDay.SUNDAY, // traitement automatique hebdomadaire
  HEURE_AUTO: 20,
  PAUSE_MS: 4000,                      // pause entre 2 reçus (quota gratuit Gemini)
  DUREE_MAX_MS: 5 * 60 * 1000,         // Apps Script s'arrête à 6 min

  // Dépenses ajoutées automatiquement chaque mois (le jour indiqué)
  RECURRENTES: [
    { jour: 1, fournisseur: 'Anthropic (Claude)', description: 'Abonnement Claude',
      categorie: 'Logiciels & abonnements', ttc: 1100, paiement: 'Virement', motsCles: ['claude', 'anthropic'] }
  ]
};
CONFIG.URL_APP = 'https://babybebe974-arch.github.io/spend-smart/';
CONFIG.URL_API = 'https://script.google.com/macros/s/AKfycbzLvp065hwYgoAYhL7LfGDe4NljYYwT0yiX-IBFvqnP9F1hXaIPuJS-3vgC9RmpYvr1/exec';   // URL /exec du déploiement « Tout le monde » (API de l'app installée)
const JUSTIF_AUTO = 'Abonnement auto';
const COULEUR_AUTO = '#d9ead3';

const COLONNES = ['Date', 'Fournisseur', 'N° facture', 'Description', 'Catégorie',
  'HT (Rs)', 'TVA (Rs)', 'TTC (Rs)', 'Justificatif', 'Remarque', 'Paiement', 'Vu sur relevé'];
const PAIEMENTS = ['Espèces', 'Virement'];   // Virement = tout ce qui passe par la banque (Juice, carte, virement)
const ANCIENS_ONGLETS = ['Dépenses espèces'];

const CATEGORIES = [
  'Achat marchandises', 'Fret & transport', 'Douane & taxes import', 'Carburant',
  'Publicité & marketing', 'Loyer & charges', 'Électricité / eau / internet',
  'Téléphone', 'Fournitures & petit matériel', 'Entretien & réparations',
  'Honoraires (comptable, avocat…)', 'Banque & frais financiers',
  'Repas & déplacements', 'Salaires & sous-traitance',
  'Logiciels & abonnements', 'Assurances', 'Autre'
];

const TYPES_OK = ['application/pdf', 'image/jpeg', 'image/png', 'image/webp',
  'image/heic', 'image/heif'];

const COULEUR = { DOUBLON: '#f4cccc', A_VERIFIER: '#fff2cc', SANS_RECU: '#fce5cd' };
const MOIS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août',
  'sept.', 'oct.', 'nov.', 'déc.'];

// ====== MENU ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💸 Spend Smart')
    .addItem('Traiter les reçus maintenant', 'traiterFactures')
    .addItem('Ajouter une dépense sans reçu', 'ouvrirSaisie')
    .addItem('Mettre à jour le récap', 'majRecapMenu')
    .addSeparator()
    .addItem('Envoyer le récap du mois dernier', 'envoyerRecapMoisMenu')
    .addItem('Envoyer le récap TVA (trimestre)', 'envoyerAuComptable')
    .addItem('Marquer la période TVA comme déjà faite', 'marquerDejaFait')
    .addItem('Exporter tout en Excel', 'exporterExcel')
    .addSeparator()
    .addItem('Installer / réparer', 'installer')
    .addItem('Changer la clé Gemini', 'changerCle')
    .addToUi();
}

// ====== INSTALLATION ======
function installer() {
  const ui = SpreadsheetApp.getUi();
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY')) changerCle();
  const ss = preparer_(true);
  majRecap_(ss);
  const p = prochainePeriode_(new Date());
  ui.alert('✅ Installé',
    `Reçus à déposer dans le dossier Drive « ${CONFIG.DOSSIER_A_TRAITER} ».\n\n` +
    `Traitement automatique chaque dimanche à ${CONFIG.HEURE_AUTO}h.\n` +
    `Prochaine échéance TVA : ${jjmmaaaa_(p.echeance)} (${p.label}).\n` +
    `Envoi au comptable prévu le ${jjmmaaaa_(dateEnvoi_(p))} → ` +
    `${CONFIG.EMAIL_BUREAU || 'toi seulement (EMAIL_BUREAU vide)'}.`,
    ui.ButtonSet.OK);
}

// Crée/retrouve le Sheet, les dossiers et les déclencheurs (sans fenêtre)
function preparer_(forcer) {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) props.setProperty('SHEET_ID', ss.getId());
  if (!props.getProperty('SHEET_ID')) {
    throw new Error('Ouvre le Google Sheet et utilise le menu 💸 Spend Smart une première fois.');
  }
  ss = ss || SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  if (!props.getProperty('ID_A_TRAITER')) props.setProperty('ID_A_TRAITER', dossier_(CONFIG.DOSSIER_A_TRAITER).getId());
  if (!props.getProperty('ID_TRAITEES')) props.setProperty('ID_TRAITEES', dossier_(CONFIG.DOSSIER_TRAITEES).getId());
  feuille_(ss);

  const ts = ScriptApp.getProjectTriggers();
  const de = nom => ts.filter(t => t.getHandlerFunction() === nom);
  const hebdo = de('traiterFactures');
  if (forcer || !hebdo.length) {
    hebdo.forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('traiterFactures').timeBased()
      .onWeekDay(CONFIG.JOUR_AUTO).atHour(CONFIG.HEURE_AUTO).create();
  }
  const tva = de('verifierEnvoiTva');
  if (forcer || !tva.length) {
    tva.forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('verifierEnvoiTva').timeBased().everyDays(1).atHour(8).create();
  }
  return ss;
}

function changerCle() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Clé API Gemini',
    'Colle ta clé (aistudio.google.com → Get API key) :', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim()) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_KEY', r.getResponseText().trim());
    ui.alert('Clé enregistrée.');
  }
}

function dossier_(nom) {
  const it = DriveApp.getFoldersByName(nom);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nom);
}

function feuille_(ss) {
  let sh = ss.getSheetByName(CONFIG.NOM_ONGLET);
  if (!sh) {                                       // ancien nom « Dépenses espèces » : on renomme
    for (const n of ANCIENS_ONGLETS) {
      const o = ss.getSheetByName(n);
      if (o) { o.setName(CONFIG.NOM_ONGLET); sh = o; break; }
    }
  }
  // Ancienne version (colonnes différentes) : on la met de côté
  if (sh && sh.getLastRow() > 0) {
    let entete = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
      .filter(String).join('|');
    if (entete === COLONNES.slice(0, 10).join('|')) { migrerPaiement_(sh); entete = COLONNES.join('|'); }
    if (entete === COLONNES.join('|')) fusionnerVirements_(sh);
    if (entete !== COLONNES.join('|')) {
      if (sh.getLastRow() > 1) {
        sh.setName(CONFIG.NOM_ONGLET + ' (ancienne ' +
          Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm') + ')');
        sh = null;
      } else {
        sh.clear();
      }
    }
  }
  if (!sh) sh = ss.insertSheet(CONFIG.NOM_ONGLET, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLONNES);
    sh.getRange(1, 1, 1, COLONNES.length)
      .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
    sh.getRange('F2:H').setNumberFormat('#,##0.00');
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(4, 280);
    sh.setColumnWidth(5, 190);
    sh.setColumnWidth(10, 280);
  }
  return sh;
}

// Ajout des colonnes « Paiement » et « Vu sur relevé » à un tableau existant
function migrerPaiement_(sh) {
  sh.getRange(1, 11, 1, 2).setValues([['Paiement', 'Vu sur relevé']])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const v = sh.getRange(2, 1, n, 10).getValues();
  v.forEach((r, i) => {
    const pay = r[8] === JUSTIF_AUTO ? 'Virement' : r[8] === 'Sans reçu' ? 'Espèces'
      : /phoenix/i.test(r[1]) ? 'Virement' : '';
    marquerPaiement_(sh, i + 2, pay);
  });
}

function normPaiement_(x) {
  const t = String(x || '').toLowerCase();
  if (/esp|cash|liquide/.test(t)) return 'Espèces';
  if (/juice|vir|transfer|banking|cart|card|visa|master|cb/.test(t)) return 'Virement';
  return '';
}

// Ancien découpage Carte / Juice → Virement (une seule fois)
function fusionnerVirements_(sh) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MIGR_PAIE2') || sh.getLastRow() < 2) return;
  const r = sh.getRange(2, 11, sh.getLastRow() - 1, 1);
  r.setValues(r.getValues().map(v => [/carte|juice/i.test(v[0]) ? 'Virement' : v[0]]));
  props.setProperty('MIGR_PAIE2', '1');
}

// Colonne K = mode de paiement ; L = case à cocher quand la dépense est retrouvée sur le relevé
function marquerPaiement_(sh, r, pay) {
  sh.getRange(r, 11).setValue(pay || '');
  const c = sh.getRange(r, 12);
  if (pay && pay !== 'Espèces') { c.insertCheckboxes(); } else { c.clearDataValidations().clearContent(); }
}

function ouvrirSS_() {
  return preparer_(false);
}

// ====== LECTURE DU TABLEAU ======
function lireLignes_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const v = sh.getRange(2, 1, n, COLONNES.length).getValues();
  const liens = sh.getRange(2, 9, n, 1).getRichTextValues();
  return v.map((r, i) => ({
    date: r[0], fournisseur: r[1], numero: r[2], description: r[3], categorie: r[4],
    ht: r[5], tva: r[6], ttc: r[7], justif: r[8], remarque: r[9], paiement: r[10], releve: r[11] === true,
    lien: liens[i][0] ? liens[i][0].getLinkUrl() : null
  }));
}

// ====== TRAITEMENT DES REÇUS ======
// Menu + déclencheur du dimanche
function traiterFactures() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  const res = executer_(ui ? 'Menu' : 'Auto (dimanche)', CONFIG.DUREE_MAX_MS);
  if (ui) {
    ui.alert('💸 Spend Smart', res.msg, ui.ButtonSet.OK);
  } else if (!res.rienAFaire) {
    MailApp.sendEmail(moi_(), '💸 Spend Smart : récap des dépenses', res.msg);
  }
}

// Appelée par le bouton du raccourci Android
// ====== DEMANDES DE CLAUDE (fichiers JSON déposés dans Drive) ======
// Claude dépose {"action":"sansRecu","depense":{date,fournisseur,description,categorie,montant}}
// ou {"action":"traiter"} dans « Spend Smart - Demandes Claude » ; résultat dans « Spend Smart - Demandes traitées ».
function installerDemandesClaude() {
  dossier_('Spend Smart - Demandes Claude');
  dossier_('Spend Smart - Demandes traitées');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'traiterDemandesClaude')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('traiterDemandesClaude').timeBased().everyMinutes(1).create();
  return 'OK';
}

function traiterDemandesClaude() {
  const verrou = LockService.getUserLock();       // une seule exécution à la fois
  if (!verrou.tryLock(1000)) return;
  try {
    const src = dossier_('Spend Smart - Demandes Claude');
    const dest = dossier_('Spend Smart - Demandes traitées');
    const fichiers = [];
    const it = src.getFiles();
    while (it.hasNext()) fichiers.push(it.next());
    fichiers.forEach(f => traiterUneDemande_(f, dest));
  } finally {
    verrou.releaseLock();
  }
}

function traiterUneDemande_(f, dest) {
  {
    f.moveTo(dest);                                // retirée tout de suite : jamais traitée deux fois
    const nom = f.getName().replace(/\.json$/i, '');
    let res;
    try {
      const q = JSON.parse(f.getBlob().getDataAsString());
      if (q.action === 'sansRecu') res = { ok: true, msg: ajouterSansRecu(q.depense) };
      else if (q.action === 'traiter') res = { ok: true, msg: traiterDepuisWeb() };
      else if (q.action === 'resume') res = Object.assign({ ok: true }, resumeMois_());
      else throw new Error('Action inconnue : ' + q.action);
    } catch (e) {
      res = { ok: false, erreur: e.message };
    }
    dest.createFile(nom + '_RESULTAT.json', JSON.stringify(res, null, 2), 'application/json');
  }
}

function traiterDepuisWeb() {
  return executer_('Android', CONFIG.DUREE_MAX_MS).msg;
}

// Photos envoyées depuis le téléphone : rangées dans « À traiter » puis lues tout de suite
function recevoirPhotos(liste, paiement) {
  const pay = normPaiement_(paiement);
  if (!liste || !liste.length) throw new Error('Aucune photo reçue');
  preparer_(false);
  const src = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ID_A_TRAITER'));
  const horo = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd_HHmmss');
  liste.forEach((p, i) => {
    const mime = TYPES_OK.includes(p.mime) ? p.mime : 'image/jpeg';
    const ext = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : '.jpg';
    const f = src.createFile(Utilities.newBlob(Utilities.base64Decode(p.b64), mime, `photo_${horo}_${i + 1}${ext}`));
    if (pay) f.setDescription('paiement:' + pay);      // choix fait dans l'app : prioritaire sur la lecture
  });
  return executer_('Android (photo)', CONFIG.DUREE_MAX_MS).msg;
}

function executer_(source, dureeMax) {
  const debut = Date.now();
  const props = PropertiesService.getScriptProperties();
  const cle = (props.getProperty('GEMINI_KEY') || '').trim();
  if (!cle) {
    return { msg: '🔑 Clé Gemini manquante : Paramètres du projet → Propriétés du script → GEMINI_KEY.', rienAFaire: false };
  }
  const verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) return { msg: '⏳ Un traitement est déjà en cours, réessaie dans une minute.', rienAFaire: true };

  try {
    let ss;
    try { ss = ouvrirSS_(); } catch (e) { return { msg: '⚠️ ' + e.message, rienAFaire: false }; }
    const sh = feuille_(ss);
    const lignes = lireLignes_(sh);
    const src = DriveApp.getFolderById(props.getProperty('ID_A_TRAITER'));
    const dest = DriveApp.getFolderById(props.getProperty('ID_TRAITEES'));

    const fichiers = [];
    const it = src.getFiles();
    while (it.hasNext()) fichiers.push(it.next());

    const ajouts = [];
    const erreurs = [];
    let restants = 0;

    for (const f of fichiers) {
      if (Date.now() - debut > dureeMax) { restants++; continue; }
      const mime = f.getMimeType();
      if (!TYPES_OK.includes(mime)) { erreurs.push(`${f.getName()} (format non lu)`); continue; }
      try {
        const blob = f.getBlob();
        if (blob.getBytes().length > 15 * 1024 * 1024) throw new Error('fichier trop lourd (>15 Mo)');
        const d = lireFacture_(blob, mime, cle);
        const choix = /paiement:(\S+)/.exec(f.getDescription() || '');
        if (choix) d.paiement = choix[1];
        const info = ajouterLigne_(sh, lignes, d, f);
        f.setName(nomPropre_(d, f.getName()));
        f.moveTo(dest);
        ajouts.push(info);
        Utilities.sleep(CONFIG.PAUSE_MS);
      } catch (e) {
        erreurs.push(`${f.getName()} : ${e.message}`);
      }
    }

    if (ajouts.length) {
      trier_(sh);
      majRecap_(ss);
    }

    // ----- Récap de ce passage -----
    const total = ajouts.reduce((s, a) => s + (Number(a.ttc) || 0), 0);
    const nbDbl = ajouts.filter(a => a.doublon).length;
    const nbVerif = ajouts.filter(a => a.aVerifier && !a.doublon).length;
    let msg = ajouts.length
      ? `✅ ${ajouts.length} reçu(s) ajouté(s) — total ${rs_(total)}\n\n` +
        ajouts.map(a => `• ${a.date ? jjmm_(a.date) : '??/??'} — ${a.fournisseur || '?'} — ${rs_(a.ttc)} — ${a.categorie}` +
          (a.doublon ? '  🔴 doublon ?' : a.aVerifier ? '  🟡 à vérifier' : '')).join('\n')
      : 'Aucun nouveau reçu à traiter.';
    if (nbDbl) msg += `\n\n🔴 ${nbDbl} doublon(s) possible(s) : lignes en rouge, supprime celle en trop.`;
    if (nbVerif) msg += `\n🟡 ${nbVerif} ligne(s) à vérifier (en jaune).`;
    if (restants) msg += `\n\n${restants} reçu(s) restant(s) : relance le traitement.`;
    if (erreurs.length) msg += `\n\n⚠️ Non traités (restent dans « À traiter ») :\n- ${erreurs.join('\n- ')}`;
    const p = prochainePeriode_(new Date());
    msg += `\n\n📅 Prochaine déclaration TVA : ${jjmmaaaa_(p.echeance)} — envoi au comptable le ${jjmmaaaa_(dateEnvoi_(p))}.`;

    if (ajouts.length || erreurs.length) {
      journal_(ss, source, ajouts.length, total, nbDbl, nbVerif,
        erreurs.length ? erreurs.length + ' erreur(s)' : '');
    }
    return { msg: msg, rienAFaire: ajouts.length + restants + erreurs.length === 0 };
  } finally {
    verrou.releaseLock();
  }
}

// Modèle Gemini : celui mémorisé, sinon CONFIG.MODELE ; si Google le retire (404), on en cherche un autre tout seul
function modele_() {
  return PropertiesService.getScriptProperties().getProperty('MODELE_GEMINI') || CONFIG.MODELE;
}

function trouverModele_(cle, exclu) {
  const rep = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' +
    encodeURIComponent(cle), { muteHttpExceptions: true });
  if (rep.getResponseCode() !== 200) return null;
  const noms = (JSON.parse(rep.getContentText()).models || [])
    .filter(m => (m.supportedGenerationMethods || []).includes('generateContent'))
    .map(m => m.name.replace(/^models\//, ''))
    .filter(n => /flash/i.test(n) && !/lite|image|tts|audio|live|thinking|exp|preview/i.test(n) && n !== exclu);
  if (!noms.length) return null;
  const v = n => (n.match(/(\d+(?:\.\d+)?)/) || [0, 0])[1] * 1;
  noms.sort((a, b) => (/latest/.test(b) - /latest/.test(a)) || (v(b) - v(a)));
  return noms[0];
}

function lireFacture_(blob, mime, cle) {
  const payload = {
    contents: [{
      parts: [
        { inline_data: { mime_type: mime, data: Utilities.base64Encode(blob.getBytes()) } },
        { text: prompt_() }
      ]
    }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  const opts = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  const appel = m => UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${m}` +
    `:generateContent?key=${encodeURIComponent(cle)}`, opts);

  let modele = modele_();
  let rep = appel(modele);
  if (rep.getResponseCode() === 404) {       // modèle retiré par Google : on bascule sur le Flash actuel
    const autre = trouverModele_(cle, modele);
    if (autre) {
      PropertiesService.getScriptProperties().setProperty('MODELE_GEMINI', autre);
      modele = autre;
      rep = appel(modele);
    }
  }
  if ([429, 500, 503].includes(rep.getResponseCode())) {
    Utilities.sleep(20000);                 // quota ou serveur occupé : on réessaie une fois
    rep = appel(modele);
  }
  if (rep.getResponseCode() !== 200) {
    throw new Error(`Gemini ${rep.getResponseCode()} ${rep.getContentText().slice(0, 150)}`);
  }
  const json = JSON.parse(rep.getContentText());
  const parts = (json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts) || [];
  const texte = parts.map(p => p.text || '').join('').replace(/```json|```/g, '').trim();
  if (!texte) throw new Error('réponse vide');
  const d = JSON.parse(texte);
  return Array.isArray(d) ? d[0] : d;
}

function prompt_() {
  return `Tu es assistant comptable pour Smartdesignforyou, entreprise d'import basée à Balaclava, Maurice.
Ce document est une preuve de DÉPENSE : facture, ticket de caisse, reçu manuscrit, photo, ou capture d'écran
d'un reçu numérique (abonnement, application, e-mail de confirmation, paiement en ligne).
Les montants sont normalement en roupies mauriciennes (Rs, MUR). La TVA à Maurice est de 15 %.
Réponds UNIQUEMENT avec ce JSON :
{"date":"AAAA-MM-JJ","fournisseur":"","numero":"","description":"","categorie":"","ht":0,"tva":0,"ttc":0,"devise":"MUR","paiement":"","confiance":"haute"}

Règles :
- date : date de la facture ou du paiement (pas l'échéance). Les dates locales sont au format JJ/MM/AAAA.
- fournisseur : nom commercial de l'émetteur (jamais Smartdesignforyou).
- numero : numéro de facture, de reçu ou de commande, sinon null.
- description : en français, 3 à 10 mots, ce qui a été acheté (ex. "Abonnement mensuel Canva Pro").
- categorie : exactement une valeur parmi : ${CATEGORIES.join(' | ')}.
- ht, tva, ttc : nombres sans espace ni symbole, point décimal. ttc = total payé.
  Si la TVA n'apparaît pas, tva = null et ht = null.
- devise : "MUR" pour Rs / roupies ; sinon le code ISO réellement affiché (USD, EUR…).
- paiement : "Espèces" (cash, espèces, monnaie rendue) ou "Virement" (tout paiement passé par la banque : MCB Juice,
  carte VISA/Mastercard, paiement en ligne, virement, Internet Banking) ; null si le reçu ne l'indique pas.
- Valeur absente ou illisible : null.
- confiance : "basse" si document flou, manuscrit, ou doute sur un montant ou la date ; sinon "haute".`;
}

function ajouterLigne_(sh, lignes, d, f) {
  let date = '';
  if (d.date && /^\d{4}-\d{2}-\d{2}$/.test(d.date)) {
    date = new Date(Number(d.date.slice(0, 4)), Number(d.date.slice(5, 7)) - 1, Number(d.date.slice(8, 10)));
  }
  const cat = CATEGORIES.includes(d.categorie) ? d.categorie : 'Autre';
  const devise = String(d.devise || 'MUR').toUpperCase();
  const enRoupies = ['MUR', 'RS', 'RP'].includes(devise);
  const n = { date: date, fournisseur: d.fournisseur || '', numero: d.numero || '', ttc: num_(d.ttc) };

  const remarques = [];
  const dbl = doublon_(lignes, n) || doublonAbonnement_(lignes, n, d.description);
  if (dbl) remarques.push(`⚠️ Doublon possible : ${dbl.fournisseur || '?'} du ${dbl.date instanceof Date ? jjmmaaaa_(dbl.date) : '?'}`);
  if (!enRoupies) remarques.push(`Montant en ${devise} : à convertir en Rs`);
  if (d.confiance === 'basse') remarques.push('Lecture incertaine');
  if (!date) remarques.push('Date illisible');
  if (n.ttc === '') remarques.push('Montant illisible');
  const aVerifier = remarques.length > 0 && !dbl;

  sh.appendRow([date, n.fournisseur, n.numero, d.description || '', cat,
    num_(d.ht), num_(d.tva), n.ttc, '', remarques.join(' · ')]);
  const r = sh.getLastRow();
  marquerPaiement_(sh, r, normPaiement_(d.paiement));
  sh.getRange(r, 9).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText('📄 Voir').setLinkUrl(f.getUrl()).build());
  if (dbl) sh.getRange(r, 1, 1, COLONNES.length).setBackground(COULEUR.DOUBLON);
  else if (aVerifier) sh.getRange(r, 1, 1, COLONNES.length).setBackground(COULEUR.A_VERIFIER);

  lignes.push(Object.assign({}, n, { categorie: cat }));
  return { date: date, fournisseur: n.fournisseur, ttc: n.ttc, categorie: cat, doublon: !!dbl, aVerifier: aVerifier };
}

// Même n° de facture chez le même fournisseur, OU même jour + même montant + même fournisseur
function doublon_(lignes, n) {
  const norm = s => String(s || '').toLowerCase().normalize('NFD')
    .replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
  const f5 = norm(n.fournisseur).slice(0, 5);
  for (const l of lignes) {
    const memeFourn = f5 && norm(l.fournisseur).slice(0, 5) === f5;
    const memeNum = n.numero && l.numero && norm(n.numero) === norm(l.numero);
    const memeMontant = n.ttc !== '' && l.ttc !== '' && Math.abs(Number(l.ttc) - Number(n.ttc)) < 0.01;
    const memeJour = n.date instanceof Date && l.date instanceof Date && jour_(n.date) === jour_(l.date);
    if ((memeNum && memeFourn) || (memeMontant && memeJour && memeFourn)) return l;
  }
  return null;
}

// Reçu d'un abonnement déjà ajouté automatiquement ce mois-ci (ex. facture Claude)
function doublonAbonnement_(lignes, n, description) {
  if (!(n.date instanceof Date)) return null;
  const texte = String((n.fournisseur || '') + ' ' + (description || '')).toLowerCase();
  const mois = Utilities.formatDate(n.date, tz_(), 'yyyy-MM');
  for (const rec of CONFIG.RECURRENTES) {
    if (!rec.motsCles.some(m => texte.includes(m))) continue;
    const l = lignes.find(x => x.justif === JUSTIF_AUTO && x.fournisseur === rec.fournisseur &&
      x.date instanceof Date && Utilities.formatDate(x.date, tz_(), 'yyyy-MM') === mois);
    if (l) return l;
  }
  return null;
}

// Ajoute les abonnements mensuels (une seule fois par mois). Lancé chaque jour à 8h.
function ajouterAbonnements() {
  const props = PropertiesService.getScriptProperties();
  const auj = jourSeul_(new Date());
  const mois = Utilities.formatDate(auj, tz_(), 'yyyy-MM');
  const aFaire = CONFIG.RECURRENTES.map((rec, i) => ({ rec, cle: 'RECUR_' + i + '_' + mois }))
    .filter(x => auj.getDate() >= x.rec.jour && !props.getProperty(x.cle));
  if (!aFaire.length) return 0;
  const ss = ouvrirSS_();
  const sh = feuille_(ss);
  aFaire.forEach(({ rec, cle }) => {
    const date = new Date(auj.getFullYear(), auj.getMonth(), rec.jour);
    sh.appendRow([date, rec.fournisseur, '', rec.description, rec.categorie,
      '', '', rec.ttc, JUSTIF_AUTO, 'Ajouté automatiquement chaque mois']);
    marquerPaiement_(sh, sh.getLastRow(), rec.paiement || '');
    sh.getRange(sh.getLastRow(), 1, 1, COLONNES.length).setBackground(COULEUR_AUTO);
    props.setProperty(cle, '1');
    journal_(ss, 'Abonnement auto', 1, rec.ttc, 0, 0, rec.description);
  });
  trier_(sh);
  majRecap_(ss);
  return aFaire.length;
}

function trier_(sh) {
  if (sh.getLastRow() > 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, COLONNES.length).sort({ column: 1, ascending: true });
  }
}

function nomPropre_(d, ancien) {
  const ext = (ancien.match(/\.[a-z0-9]{2,5}$/i) || [''])[0];
  const fourn = String(d.fournisseur || 'Inconnu').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 40);
  return `${d.date || 'sans-date'}_${fourn}_${num_(d.ttc)}Rs${ext}`;
}

// ====== DÉPENSE SANS REÇU ======
function ouvrirSaisie() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(page_(false)).setWidth(440).setHeight(640), 'Spend Smart');
}

function ajouterSansRecu(o) {
  const montant = num_(o && o.montant);
  if (montant === '' || montant <= 0) throw new Error('Montant manquant');
  if (!o.description || !String(o.description).trim()) throw new Error('Description manquante');
  const verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) throw new Error('Un traitement est en cours, réessaie dans une minute.');
  try {
    const ss = ouvrirSS_();
    const sh = feuille_(ss);
    const lignes = lireLignes_(sh);
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(o.date || '');
    const date = m ? new Date(+m[1], +m[2] - 1, +m[3]) : jourSeul_(new Date());
    const cat = CATEGORIES.includes(o.categorie) ? o.categorie : 'Autre';
    const n = { date: date, fournisseur: String(o.fournisseur || '').trim(), numero: '', ttc: montant };
    const dbl = doublon_(lignes, n);
    const remarque = 'Sans justificatif — saisi le ' + jjmmaaaa_(new Date()) +
      (dbl ? ` · ⚠️ Doublon possible : ${dbl.fournisseur || '?'} du ${jjmmaaaa_(dbl.date)}` : '');

    sh.appendRow([date, n.fournisseur || '—', '', String(o.description).trim(), cat,
      '', '', montant, 'Sans reçu', remarque]);
    marquerPaiement_(sh, sh.getLastRow(), normPaiement_(o.paiement) || 'Espèces');
    sh.getRange(sh.getLastRow(), 1, 1, COLONNES.length)
      .setBackground(dbl ? COULEUR.DOUBLON : COULEUR.SANS_RECU);
    trier_(sh);
    majRecap_(ss);
    journal_(ss, 'Sans reçu', 1, montant, dbl ? 1 : 0, 0, String(o.description).trim());
    return `✅ Ajouté : ${jjmm_(date)} — ${o.description} — ${rs_(montant)}` +
      (dbl ? '\n🔴 Attention : ressemble à une dépense déjà saisie.' : '');
  } finally {
    verrou.releaseLock();
  }
}

// ====== RACCOURCI ANDROID (application Web) ======
function doGet() {
  let actif = '';
  try { actif = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!actif || actif !== moi_()) {   // déploiement public (API) : pas de page
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:20px">🔒 Accès réservé — ouvre l\'app Spend Smart.</p>');
  }
  return HtmlService.createHtmlOutput(page_(true))
    .setTitle('Spend Smart')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ====== API pour l'app installée (PWA GitHub Pages) ======
function cleApp_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('SPEND_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); props.setProperty('SPEND_KEY', k); }
  return k;
}

function doPost(e) {
  let r;
  try {
    const q = JSON.parse(e.postData.contents);
    const cle = PropertiesService.getScriptProperties().getProperty('SPEND_KEY');
    if (!cle || q.cle !== cle) throw new Error('Clé incorrecte : rouvre le lien d\'installation depuis la page Spend Smart.');
    if (q.action === 'photos') r = { msg: recevoirPhotos(q.liste, q.paiement) };
    else if (q.action === 'sansRecu') r = { msg: ajouterSansRecu(q.depense) };
    else if (q.action === 'traiter') r = { msg: traiterDepuisWeb() };
    else if (q.action === 'resume') r = resumeMois_();
    else if (q.action === 'ping') r = { categories: CATEGORIES };
    else throw new Error('Action inconnue');
    r.ok = true;
  } catch (err) {
    r = { ok: false, erreur: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function resumeMois_() {
  const ss = ouvrirSS_();
  const auj = new Date();
  const mois = Utilities.formatDate(auj, tz_(), 'yyyy-MM');
  const lignes = lireLignes_(feuille_(ss)).filter(l => l.date instanceof Date &&
    Utilities.formatDate(l.date, tz_(), 'yyyy-MM') === mois);
  const total = lignes.reduce((s, l) => s + (Number(l.ttc) || 0), 0);
  const parCat = {};
  lignes.forEach(l => { parCat[l.categorie] = (parCat[l.categorie] || 0) + (Number(l.ttc) || 0); });
  const dernieres = lignes.slice().sort((a, b) => b.date - a.date).slice(0, 10).map(l => ({
    date: jjmm_(l.date), fournisseur: String(l.fournisseur || ''), description: String(l.description || ''),
    ttc: Number(l.ttc) || 0, alerte: /Doublon/.test(l.remarque) ? 'doublon' : (l.remarque && l.justif !== JUSTIF_AUTO && l.justif !== 'Sans reçu' ? 'verifier' : '')
  }));
  const parPaiement = {};
  lignes.forEach(l => { const k = l.paiement || 'Non précisé'; parPaiement[k] = (parPaiement[k] || 0) + (Number(l.ttc) || 0); });
  const aPointer = lignes.filter(l => l.paiement && l.paiement !== 'Espèces' && !l.releve).length;
  return { mois: MOIS[auj.getMonth()] + ' ' + auj.getFullYear(), total: total, nb: lignes.length,
    parCat: parCat, parPaiement: parPaiement, aPointer: aPointer, dernieres: dernieres, urlSheet: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit' };
}

function page_(avecLienTableau) {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const urlSheet = id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : 'https://sheets.google.com';
  const aujourdhui = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  const options = CATEGORIES.map(c => '<option>' + c + '</option>').join('');
  return `<!doctype html><html><head><meta charset="utf-8">
<style>
  body{font-family:system-ui,sans-serif;margin:0;padding:16px;color:#1b2430;background:#fff}
  h2{margin:0 0 14px;font-size:22px}
  h3{margin:22px 0 10px;font-size:17px}
  .btn{display:block;width:100%;box-sizing:border-box;border:0;border-radius:12px;padding:15px;
       font-size:17px;font-weight:600;text-align:center;text-decoration:none;cursor:pointer}
  .vert{background:#2f5d45;color:#fff;margin-bottom:10px}.vert2{background:#e3eee7;color:#2f5d45}
  .bleu{background:#1f4e79;color:#fff}.orange{background:#e69138;color:#fff}
  .gris{background:#eef2f6;color:#1f4e79;margin-top:10px}
  label{display:block;font-size:14px;margin:10px 0 4px;color:#4a5563}
  input,select{width:100%;box-sizing:border-box;padding:12px;font-size:16px;border:1px solid #c9d2dc;border-radius:10px;background:#fff}
  #res{white-space:pre-wrap;font-size:15px;line-height:1.5;background:#f3f6fa;border-radius:12px;padding:14px;margin-top:14px;display:none}
  .btn:disabled{opacity:.6}
</style></head><body>
<h2>💸 Spend Smart</h2>
<button class="btn vert" id="bP" onclick="document.getElementById('fP').click()">📷 Photographier un reçu</button>
<button class="btn vert2" id="bG" onclick="document.getElementById('fG').click()">🖼️ Photo / capture / PDF du téléphone</button>
<input type="file" id="fP" accept="image/*" capture="environment" style="display:none" onchange="envoyer(this)">
<input type="file" id="fG" accept="image/*,application/pdf" multiple style="display:none" onchange="envoyer(this)">
<div style="height:10px"></div>
<button class="btn bleu" id="bT" onclick="traiter()">📥 Traiter les reçus du dossier Drive</button>
<div id="res"></div>

<h3>➕ Dépense sans reçu</h3>
<label>Date</label><input type="date" id="date" value="${aujourdhui}">
<label>Fournisseur / lieu</label><input id="fournisseur" placeholder="ex. Parking Port-Louis">
<label>Description *</label><input id="description" placeholder="ex. Parking livraison client">
<label>Catégorie</label><select id="categorie">${options}</select>
<label>Payé par</label><select id="paiement">${PAIEMENTS.map(x => '<option>' + x + '</option>').join('')}</select>
<label>Montant payé (Rs) *</label><input id="montant" type="number" inputmode="decimal" step="0.01" min="0">
<div style="height:14px"></div>
<button class="btn orange" id="bA" onclick="ajouter()">Ajouter la dépense</button>
${avecLienTableau ? '<a class="btn gris" target="_blank" href="' + urlSheet + '">📊 Ouvrir le tableau</a>' : ''}
${avecLienTableau ? '<a class="btn vert" style="margin-top:10px" target="_blank" href="' + CONFIG.URL_APP + '#k=' + cleApp_() + '&u=' + encodeURIComponent(CONFIG.URL_API) + '">📲 Installer l\'app sur ce téléphone</a>' : ''}

<script>
  function montrer(t){var r=document.getElementById('res');r.textContent=t;r.style.display='block';
    window.scrollTo({top:0,behavior:'smooth'});}
  function occupe(id,oui,txt){var b=document.getElementById(id);b.disabled=oui;if(txt)b.textContent=txt;}
  function traiter(){
    occupe('bT',true,'⏳ Lecture des reçus…');
    google.script.run
      .withSuccessHandler(function(m){montrer(m);occupe('bT',false,'📥 Traiter les reçus scannés');})
      .withFailureHandler(function(e){montrer('⚠️ '+e.message);occupe('bT',false,'📥 Traiter les reçus scannés');})
      .traiterDepuisWeb();
  }
  function ajouter(){
    var o={date:val('date'),fournisseur:val('fournisseur'),description:val('description'),
           categorie:val('categorie'),paiement:val('paiement'),montant:val('montant')};
    if(!o.description||!o.montant){montrer('Remplis au moins la description et le montant.');return;}
    occupe('bA',true,'⏳ Ajout…');
    google.script.run
      .withSuccessHandler(function(m){montrer(m);occupe('bA',false,'Ajouter la dépense');
        document.getElementById('fournisseur').value='';document.getElementById('description').value='';
        document.getElementById('montant').value='';})
      .withFailureHandler(function(e){montrer('⚠️ '+e.message);occupe('bA',false,'Ajouter la dépense');})
      .ajouterSansRecu(o);
  }
  function val(id){return document.getElementById(id).value.trim();}
  // Réduit les photos (max 1800 px, JPEG) pour un envoi rapide en 3G/4G
  function lire(f){return new Promise(function(ok,ko){
    var r=new FileReader();r.onerror=ko;
    if(f.type==='application/pdf'){r.onload=function(){ok({b64:r.result.split(',')[1],mime:'application/pdf'});};r.readAsDataURL(f);return;}
    r.onload=function(){var img=new Image();img.onerror=function(){ok({b64:r.result.split(',')[1],mime:f.type||'image/jpeg'});};
      img.onload=function(){var m=1800,k=Math.min(1,m/Math.max(img.width,img.height));
        var c=document.createElement('canvas');c.width=Math.round(img.width*k);c.height=Math.round(img.height*k);
        c.getContext('2d').drawImage(img,0,0,c.width,c.height);
        ok({b64:c.toDataURL('image/jpeg',0.85).split(',')[1],mime:'image/jpeg'});};
      img.src=r.result;};
    r.readAsDataURL(f);});}
  function envoyer(inp){
    var fs=Array.prototype.slice.call(inp.files||[]);if(!fs.length)return;
    var b=inp.id==='fP'?'bP':'bG',lib=document.getElementById(b).textContent;
    occupe(b,true,'⏳ Envoi et lecture de '+fs.length+' reçu(s)…');
    Promise.all(fs.map(lire)).then(function(l){
      google.script.run
        .withSuccessHandler(function(m){montrer(m);occupe(b,false,lib);inp.value='';})
        .withFailureHandler(function(e){montrer('⚠️ '+e.message);occupe(b,false,lib);inp.value='';})
        .recevoirPhotos(l);
    }).catch(function(e){montrer('⚠️ Photo illisible : '+e);occupe(b,false,lib);});
  }
</script></body></html>`;
}

// ====== RÉCAP ======
function majRecapMenu() {
  majRecap_(ouvrirSS_());
  SpreadsheetApp.getUi().alert('Récap mis à jour (onglet « Récap »).');
}

function majRecap_(ss) {
  const lignes = lireLignes_(feuille_(ss)).filter(l => l.date instanceof Date);
  let rc = ss.getSheetByName('Récap');
  if (!rc) rc = ss.insertSheet('Récap', 1);
  rc.clear();

  const p = prochainePeriode_(new Date());
  rc.getRange(1, 1).setValue('Récap des dépenses (Rs TTC)').setFontWeight('bold').setFontSize(14);
  rc.getRange(2, 1).setValue(`Mis à jour le ${Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm')} · ` +
    `Prochaine déclaration TVA : ${jjmmaaaa_(p.echeance)} (${p.label}) · envoi au comptable le ${jjmmaaaa_(dateEnvoi_(p))}`)
    .setFontColor('#4a5563');
  if (!lignes.length) { rc.getRange(4, 1).setValue('Aucune dépense pour le moment.'); return; }

  // Bloc 1 : catégories × mois
  const mois = [...new Set(lignes.map(l => Utilities.formatDate(l.date, tz_(), 'yyyy-MM')))].sort();
  const tab = [['Catégorie'].concat(mois.map(m => MOIS[+m.slice(5) - 1] + ' ' + m.slice(0, 4)), ['Total'])];
  const totCol = new Array(mois.length + 1).fill(0);
  CATEGORIES.forEach(c => {
    const l = lignes.filter(x => x.categorie === c);
    if (!l.length) return;
    const vals = mois.map(m => l.filter(x => Utilities.formatDate(x.date, tz_(), 'yyyy-MM') === m)
      .reduce((s, x) => s + (Number(x.ttc) || 0), 0));
    const t = vals.reduce((a, b) => a + b, 0);
    vals.forEach((v, i) => totCol[i] += v);
    totCol[mois.length] += t;
    tab.push([c].concat(vals, [t]));
  });
  tab.push(['TOTAL'].concat(totCol));
  let ligne = 4;
  rc.getRange(ligne, 1, tab.length, tab[0].length).setValues(tab);
  rc.getRange(ligne, 1, 1, tab[0].length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
  rc.getRange(ligne + tab.length - 1, 1, 1, tab[0].length).setFontWeight('bold').setBackground('#dbe5f1');
  rc.getRange(ligne + 1, 2, tab.length - 1, tab[0].length - 1).setNumberFormat('#,##0');
  ligne += tab.length + 2;

  // Bloc 1 bis : par moyen de paiement × mois + pointage relevé
  const modes = PAIEMENTS.concat(['Non précisé']);
  const tp = [['Payé par'].concat(mois.map(m => MOIS[+m.slice(5) - 1] + ' ' + m.slice(0, 4)), ['Total', 'Pas encore vu sur relevé'])];
  modes.forEach(md => {
    const l = lignes.filter(x => (x.paiement || 'Non précisé') === md);
    if (!l.length) return;
    const vals = mois.map(m => l.filter(x => Utilities.formatDate(x.date, tz_(), 'yyyy-MM') === m)
      .reduce((s, x) => s + (Number(x.ttc) || 0), 0));
    const aPointer = md === 'Espèces' || md === 'Non précisé' ? '' : l.filter(x => !x.releve).length;
    tp.push([md].concat(vals, [vals.reduce((a, b) => a + b, 0), aPointer]));
  });
  rc.getRange(ligne, 1).setValue('Espèces / Virement (cocher « Vu sur relevé » après contrôle du relevé MCB)').setFontWeight('bold');
  rc.getRange(ligne + 1, 1, tp.length, tp[0].length).setValues(tp);
  rc.getRange(ligne + 1, 1, 1, tp[0].length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
  rc.getRange(ligne + 2, 2, Math.max(tp.length - 1, 1), tp[0].length - 2).setNumberFormat('#,##0');
  ligne += tp.length + 3;

  // Bloc 2 : par période de déclaration TVA
  const periodes = {};
  lignes.forEach(l => {
    const pp = periodeContenant_(l.date);
    const k = pp.key;
    if (!periodes[k]) periodes[k] = { p: pp, nb: 0, sans: 0, ht: 0, tva: 0, ttc: 0 };
    const o = periodes[k];
    o.nb++; if (!l.lien) o.sans++;
    o.ht += Number(l.ht) || 0; o.tva += Number(l.tva) || 0; o.ttc += Number(l.ttc) || 0;
  });
  const t2 = [['Période TVA', 'Nb dépenses', 'dont sans reçu', 'HT (Rs)', 'TVA (Rs)', 'TTC (Rs)', 'Échéance', 'Envoyé au comptable']];
  const props = PropertiesService.getScriptProperties();
  Object.keys(periodes).sort().forEach(k => {
    const o = periodes[k];
    const envoi = props.getProperty(k);
    t2.push([o.p.label, o.nb, o.sans, o.ht, o.tva, o.ttc, jjmmaaaa_(o.p.echeance),
      envoi ? '✅ ' + envoi : '—']);
  });
  rc.getRange(ligne, 1).setValue('Par période de déclaration TVA (MRA)').setFontWeight('bold');
  rc.getRange(ligne + 1, 1, t2.length, t2[0].length).setValues(t2);
  rc.getRange(ligne + 1, 1, 1, t2[0].length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
  rc.getRange(ligne + 2, 4, Math.max(t2.length - 1, 1), 3).setNumberFormat('#,##0.00');
  rc.setColumnWidth(1, 230);
  rc.setFrozenRows(0);
}

// ====== HISTORIQUE ======
function journal_(ss, action, nb, total, doublons, aVerifier, detail) {
  let h = ss.getSheetByName('Historique');
  if (!h) {
    h = ss.insertSheet('Historique');
    h.appendRow(['Date / heure', 'Action', 'Nb', 'Total TTC (Rs)', 'Doublons', 'À vérifier', 'Détail']);
    h.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
    h.setFrozenRows(1);
  }
  h.insertRowAfter(1);
  h.getRange(2, 1, 1, 7).setValues([[new Date(), action, nb, total, doublons, aVerifier, detail || '']]);
  h.getRange(2, 1).setNumberFormat('dd/mm/yyyy HH:mm');
  h.getRange(2, 4).setNumberFormat('#,##0.00');
}

// ====== TVA MRA : PÉRIODES ET ÉCHÉANCES ======
// Trimestrielle : trimestres civils, déclaration au plus tard le 20 du mois suivant.
// Mensuelle (CA > Rs 10 M) : au plus tard le dernier jour du mois suivant.
function periodeContenant_(d) {
  const y = d.getFullYear(), m = d.getMonth();
  let debut, fin, echeance, label, code;
  if (CONFIG.PERIODICITE_TVA === 'MENSUELLE') {
    debut = new Date(y, m, 1); fin = new Date(y, m + 1, 0); echeance = new Date(y, m + 2, 0);
    label = MOIS[m] + ' ' + y; code = Utilities.formatDate(debut, tz_(), 'yyyy-MM');
  } else {
    const q = Math.floor(m / 3) * 3;
    debut = new Date(y, q, 1); fin = new Date(y, q + 3, 0); echeance = new Date(y, q + 3, 20);
    label = `T${q / 3 + 1} ${y} (${MOIS[q]}–${MOIS[q + 2]})`; code = `${y}-T${q / 3 + 1}`;
  }
  return { debut, fin, echeance, label, code, key: 'ENVOI_' + code };
}

// La période dont la déclaration est la prochaine à venir
function prochainePeriode_(auj) {
  const j = jourSeul_(auj);
  const cour = periodeContenant_(j);
  const prec = periodeContenant_(new Date(cour.debut.getFullYear(), cour.debut.getMonth(), 0));
  const props = PropertiesService.getScriptProperties();
  const precFaite = props.getProperty(prec.key) || props.getProperty(prec.key + '_MOI');
  return (j <= prec.echeance && !precFaite) ? prec : cour;
}

function dateEnvoi_(p) {
  return new Date(p.echeance.getFullYear(), p.echeance.getMonth(), p.echeance.getDate() - CONFIG.JOURS_AVANT_TVA);
}

// Récap mensuel : le mois précédent complet
function moisPrecedent_(auj) {
  const debut = new Date(auj.getFullYear(), auj.getMonth() - 1, 1);
  const fin = new Date(auj.getFullYear(), auj.getMonth(), 0);
  const code = Utilities.formatDate(debut, tz_(), 'yyyy-MM');
  return { debut, fin, echeance: null, mensuel: true, label: MOIS[debut.getMonth()] + ' ' + debut.getFullYear(),
    code: code, key: 'RECAP_' + code };
}

function envoyerRecapMoisMenu() {
  const ui = SpreadsheetApp.getUi();
  const p = moisPrecedent_(new Date());
  executer_('Menu (récap du mois)', 3 * 60 * 1000);
  ui.alert('💸 Spend Smart', envoyerComptable_(p), ui.ButtonSet.OK);
}

// Déclencheur quotidien (8h) : récap mensuel le 1er, récap TVA X jours avant l'échéance
function verifierEnvoiTva() {
  try { ajouterAbonnements(); } catch (e) { console.error(e); }
  const props = PropertiesService.getScriptProperties();
  const pm = moisPrecedent_(new Date());
  if (!props.getProperty(pm.key)) {
    executer_('Auto (récap du mois)', 3 * 60 * 1000);
    envoyerComptable_(pm);
  }
  const p = prochainePeriode_(new Date());
  const auj = jourSeul_(new Date());
  if (auj < dateEnvoi_(p)) return;
  const bureau = CONFIG.EMAIL_BUREAU.trim();
  if (props.getProperty(p.key)) return;                       // déjà envoyé au comptable
  if (!bureau && props.getProperty(p.key + '_MOI')) return;   // déjà envoyé à toi
  executer_('Auto (avant TVA)', 3 * 60 * 1000);               // traite d'abord les reçus en attente
  envoyerComptable_(p);
}

// Menu : envoi manuel
function envoyerAuComptable() {
  const ui = SpreadsheetApp.getUi();
  const p = prochainePeriode_(new Date());
  const dest = CONFIG.EMAIL_BUREAU.trim() || 'toi seulement (EMAIL_BUREAU est vide)';
  const r = ui.alert('Envoyer au comptable ?',
    `Période : ${p.label}\nDu ${jjmmaaaa_(p.debut)} au ${jjmmaaaa_(p.fin)}\n` +
    `Échéance déclaration TVA : ${jjmmaaaa_(p.echeance)}\nDestinataire : ${dest}\n\n` +
    `Les reçus en attente seront traités avant l'envoi.`, ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  executer_('Menu (avant envoi)', 3 * 60 * 1000);
  ui.alert('💸 Spend Smart', envoyerComptable_(p), ui.ButtonSet.OK);
}

// Menu : la période en cours a déjà été transmise au comptable par un autre moyen
function marquerDejaFait() {
  const ui = SpreadsheetApp.getUi();
  const p = prochainePeriode_(new Date());
  const r = ui.alert('Période déjà faite ?',
    `Période : ${p.label}\nÉchéance TVA : ${jjmmaaaa_(p.echeance)}\n\n` +
    `Marquer comme déjà transmise au comptable ? Aucun envoi automatique ne partira pour cette période.`,
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty(p.key, 'fait hors appli le ' + jjmmaaaa_(new Date()));
  majRecap_(ouvrirSS_());
  const suiv = prochainePeriode_(new Date(p.echeance.getFullYear(), p.echeance.getMonth(), p.echeance.getDate() + 1));
  ui.alert('✅ Noté',
    `${p.label} : marquée comme faite.\n\nProchain envoi automatique : ${jjmmaaaa_(dateEnvoi_(suiv))} ` +
    `(${suiv.label}, échéance TVA ${jjmmaaaa_(suiv.echeance)}).`, ui.ButtonSet.OK);
}

function envoyerComptable_(p) {
  const props = PropertiesService.getScriptProperties();
  const ss = ouvrirSS_();
  const moi = moi_();
  const bureau = CONFIG.EMAIL_BUREAU.trim();
  const finJour = new Date(p.fin.getFullYear(), p.fin.getMonth(), p.fin.getDate(), 23, 59, 59);
  const lignes = lireLignes_(feuille_(ss))
    .filter(l => l.date instanceof Date && l.date >= p.debut && l.date <= finJour);

  if (!lignes.length && p.mensuel) {
    props.setProperty(p.key, 'vide');                         // mois sans dépense : rien à envoyer
    return `Aucune dépense en ${p.label} : pas de récap.`;
  }
  if (!lignes.length) {
    MailApp.sendEmail(moi, `Spend Smart : aucune dépense pour ${p.label}`,
      `Aucune dépense enregistrée pour ${p.label}.\nRien n'a été envoyé au comptable.\n` +
      `Échéance de la déclaration TVA : ${jjmmaaaa_(p.echeance)}.`);
    props.setProperty(bureau ? p.key : p.key + '_MOI', jjmmaaaa_(new Date()));
    return `Aucune dépense pour ${p.label} : rien envoyé au comptable (tu as reçu un e-mail de rappel).`;
  }

  const somme = (arr, k) => arr.reduce((s, l) => s + (Number(l[k]) || 0), 0);
  const tot = { ht: somme(lignes, 'ht'), tva: somme(lignes, 'tva'), ttc: somme(lignes, 'ttc') };
  const sans = lignes.filter(l => !l.lien);
  const douteux = lignes.filter(l => /Doublon|convertir|illisible|incertaine/.test(String(l.remarque)));
  const nomBase = `Smartdesignforyou_Depenses_${p.code}`;

  // ----- Fichier Excel -----
  const tmp = SpreadsheetApp.create(nomBase);
  const s1 = tmp.getSheets()[0];
  s1.setName('Dépenses');
  const data = [COLONNES].concat(lignes.map(l => [l.date, l.fournisseur, l.numero, l.description,
    l.categorie, l.ht, l.tva, l.ttc,
    l.lien ? '=HYPERLINK("' + l.lien + '","Voir le reçu")' : (l.justif || 'Sans reçu'), l.remarque,
    l.paiement || '', l.paiement && l.paiement !== 'Espèces' ? (l.releve ? 'oui' : 'non') : '']));
  data.push(['TOTAL', '', '', '', '', tot.ht, tot.tva, tot.ttc, '', '', '', '']);
  s1.getRange(1, 1, data.length, COLONNES.length).setValues(data);
  s1.getRange(1, 1, 1, COLONNES.length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
  s1.getRange(data.length, 1, 1, COLONNES.length).setFontWeight('bold');
  s1.getRange(2, 1, data.length - 1, 1).setNumberFormat('dd/mm/yyyy');
  s1.getRange(2, 6, data.length - 1, 3).setNumberFormat('#,##0.00');
  s1.autoResizeColumns(1, COLONNES.length);

  const parCat = CATEGORIES.map(c => {
    const l = lignes.filter(x => x.categorie === c);
    return [c, l.length, somme(l, 'ht'), somme(l, 'tva'), somme(l, 'ttc')];
  }).filter(r => r[1] > 0);
  const s2 = tmp.insertSheet('Par catégorie');
  const d2 = [['Catégorie', 'Nb', 'HT (Rs)', 'TVA (Rs)', 'TTC (Rs)']].concat(parCat,
    [['TOTAL', lignes.length, tot.ht, tot.tva, tot.ttc]]);
  s2.getRange(1, 1, d2.length, 5).setValues(d2);
  s2.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#fff');
  s2.getRange(d2.length, 1, 1, 5).setFontWeight('bold');
  s2.getRange(2, 3, d2.length - 1, 3).setNumberFormat('#,##0.00');
  s2.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();

  const xlsx = UrlFetchApp.fetch(
    `https://docs.google.com/spreadsheets/d/${tmp.getId()}/export?format=xlsx`,
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName(nomBase + '.xlsx');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);

  // ----- Reçus en .zip -----
  const blobs = [];
  let manquants = 0;
  lignes.filter(l => l.lien).forEach((l, i) => {
    const m = /\/d\/([-\w]{20,})/.exec(l.lien) || /[?&]id=([-\w]{20,})/.exec(l.lien);
    try {
      const f = DriveApp.getFileById(m[1]);
      blobs.push(f.getBlob().setName(String(i + 1).padStart(3, '0') + '_' + f.getName()));
    } catch (e) { manquants++; }
  });
  const envois = dossier_(CONFIG.DOSSIER_ENVOIS);
  envois.createFile(xlsx.copyBlob());
  const pj = [xlsx];
  let lienZip = '';
  if (blobs.length) {
    const zip = Utilities.zip(blobs, nomBase + '_recus.zip');
    const zipFile = envois.createFile(zip);
    if (zip.getBytes().length < 18 * 1024 * 1024) {
      pj.push(zip);
    } else {
      try { if (bureau) zipFile.addViewer(bureau); }
      catch (e) { zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
      lienZip = zipFile.getUrl();
    }
  }

  // ----- E-mail -----
  const td = 'style="padding:6px 10px;border-bottom:1px solid #e3e8ee"';
  const tdr = 'style="padding:6px 10px;border-bottom:1px solid #e3e8ee;text-align:right"';
  const html =
    `<p>Bonjour,</p>
<p>Voici les <b>dépenses</b> (espèces et virements) de Smartdesignforyou pour <b>${p.label}</b>
(du ${jjmmaaaa_(p.debut)} au ${jjmmaaaa_(p.fin)})${p.mensuel ? '' : `, en vue de la déclaration TVA à déposer
avant le <b>${jjmmaaaa_(p.echeance)}</b>`}. Certains virements peuvent manquer : ils figurent sur les relevés bancaires.
Fichiers aussi rangés dans Drive, dossier « ${CONFIG.DOSSIER_ENVOIS} ».</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">
<tr><td ${td}>Nombre de dépenses</td><td ${tdr}>${lignes.length}</td></tr>
<tr><td ${td}>Total TTC</td><td ${tdr}><b>${rs_(tot.ttc)}</b></td></tr>
<tr><td ${td}>dont TVA indiquée sur les reçus</td><td ${tdr}>${rs_(tot.tva)}</td></tr>
<tr><td ${td}>Dépenses sans justificatif</td><td ${tdr}>${sans.length} (${rs_(somme(sans, 'ttc'))})</td></tr>
</table>
<p><b>Par catégorie (TTC)</b></p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">
${parCat.map(r => `<tr><td ${td}>${r[0]}</td><td ${tdr}>${r[1]}</td><td ${tdr}>${rs_(r[4])}</td></tr>`).join('')}
</table>
<p>Pièces jointes : le fichier Excel détaillé${blobs.length && !lienZip ? ' et les reçus (.zip)' : ''}.
${lienZip ? `<br>Reçus (fichier volumineux) : <a href="${lienZip}">télécharger le .zip</a>` : ''}</p>
${douteux.length ? `<p style="color:#9c2a00">${douteux.length} ligne(s) portent une remarque (colonne « Remarque ») à vérifier.</p>` : ''}
<p>Bien cordialement,<br>${CONFIG.SIGNATURE}</p>`;

  const options = { htmlBody: html, attachments: pj, name: 'Smartdesignforyou' };
  if (bureau) options.cc = moi;
  MailApp.sendEmail(bureau || moi,
    p.mensuel ? `Récap mensuel des dépenses ${p.label} — Smartdesignforyou`
      : `Dépenses ${p.label} — Smartdesignforyou (TVA avant le ${jjmmaaaa_(p.echeance)})`,
    `Dépenses ${p.label} : ${lignes.length} dépenses, total ${rs_(tot.ttc)}. Voir pièces jointes.`,
    options);

  props.setProperty(p.mensuel || bureau ? p.key : p.key + '_MOI', jjmmaaaa_(new Date()));
  journal_(ss, p.mensuel ? 'Récap mensuel' : (bureau ? 'Envoi comptable' : 'Envoi (à moi)'), lignes.length, tot.ttc,
    0, douteux.length, `${p.label}${manquants ? ' · ' + manquants + ' reçu(s) introuvable(s)' : ''}`);
  majRecap_(ss);

  return `📤 Envoyé à ${bureau || moi + ' (EMAIL_BUREAU vide)'} : ${p.label}, ` +
    `${lignes.length} dépenses, ${rs_(tot.ttc)}.` +
    (douteux.length ? `\n⚠️ ${douteux.length} ligne(s) avec remarque à vérifier.` : '') +
    (manquants ? `\n⚠️ ${manquants} reçu(s) introuvable(s) dans Drive.` : '');
}

// ====== EXPORT EXCEL COMPLET ======
function exporterExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const nom = `SpendSmart_Depenses_especes_${Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd')}.xlsx`;
  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(nom);
  const f = dossier_(CONFIG.DOSSIER_ENVOIS).createFile(blob);
  SpreadsheetApp.getUi().alert('✅ Fichier Excel créé dans ton Drive',
    `${f.getName()}\n\n${f.getUrl()}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ====== OUTILS ======
function tz_() { return Session.getScriptTimeZone(); }
function moi_() { return Session.getEffectiveUser().getEmail(); }
function jourSeul_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function jour_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function jjmm_(d) { return Utilities.formatDate(d, tz_(), 'dd/MM'); }
function jjmmaaaa_(d) { return Utilities.formatDate(d, tz_(), 'dd/MM/yyyy'); }

function num_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(String(v).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isNaN(n) ? '' : n;
}

function rs_(n) {
  if (n === '' || n === null || n === undefined || isNaN(Number(n))) return '— Rs';
  const parts = (Math.round(Number(n) * 100) / 100).toFixed(2).split('.');
  return parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (parts[1] === '00' ? '' : ',' + parts[1]) + ' Rs';
}
