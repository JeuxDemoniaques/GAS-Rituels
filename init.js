/***********************
 * Helpers génériques  *
 ***********************/

/** Efface en toute sécurité les données (en gardant la 1re ligne d'en-têtes).
 *  - OK si feuille absente / vide / 0 col / 1 seule ligne
 *  - Ne touche pas aux formats/validations/filtres/tableaux
 *  - Retourne un message utilisateur
 */
function safeClearDataRows_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    const msg = `ℹ️ Feuille '${sheetName}' absente (rien à purger).`;
    LOG.info(msg);
    return msg;
  }

  const lastCol = Math.max(1, sh.getLastColumn());
  const lastRow = sh.getLastRow(); // 0=vide, 1=en-têtes, >1=avec données

  if (lastRow <= 1) {
    const msg = `ℹ️ Feuille '${sheetName}' : aucune donnée à purger.`;
    LOG.info(msg);
    return msg;
  }

  try {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    const msg = `✅ Données purgées dans '${sheetName}' (en-têtes conservés).`;
    LOG.info(msg);
    return msg;
  } catch (e) {
    const err = `❌ Erreur purge '${sheetName}' : ${e && e.message ? e.message : e}`;
    LOG.error(err);
    return err;
  }
}

/** Vérifie/complète les en-têtes d’une feuille.
 *  - Si feuille vide → pose la ligne d’en-têtes
 *  - Si partielle → complète les colonnes manquantes (dans l’ordre attendu)
 *  - Log si modif → rappel de reconvertir en tableau si tu utilises cette feature
 */
function ensureHeaders(sheet, headers) {
  if (!sheet) throw new Error("❌ ensureHeaders: feuille manquante");

  // Feuille totalement vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    LOG.info(`💡 En-têtes initialisés pour '${sheet.getName()}' → pense à 'Convertir en tableau' manuellement.`);
    return;
  }

  // S'assurer qu'on a au moins autant de colonnes que d'en-têtes attendus
  const needCols = headers.length - sheet.getLastColumn();
  if (needCols > 0) sheet.insertColumnsAfter(sheet.getLastColumn() || 1, needCols);

  // Lire/compléter la ligne 1
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let changed = false;
  headers.forEach((h, i) => {
    if (existing[i] !== h) {
      sheet.getRange(1, i + 1).setValue(h);
      changed = true;
    }
  });

  if (changed) {
    LOG.info(`💡 En-têtes ajustés pour '${sheet.getName()}' → refais 'Convertir en tableau' si nécessaire.`);
  }
}

/** Ajoute une colonne d’en-tête à la fin si absente. Retourne l’index 0‑based dans getValues(). */
function ensureColumnExists_(sh, headerName) {
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(1, sh.getLastColumn());
  if (lastRow === 0) {
    sh.appendRow([headerName]);
    return 0;
  }
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = headers.indexOf(headerName);
  if (idx === -1) {
    sh.insertColumnAfter(lastCol);                   // <-- insère une vraie colonne
    sh.getRange(1, lastCol + 1).setValue(headerName);
    idx = lastCol; // nouveau dernier index (0-based)
  }
  return idx;
}

/** Pose une validation "liste de valeurs" sur une plage (remplace la validation existante). */
function setListValidation_(range, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}


/***********************
 * Initialisation base *
 ***********************/
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = [];

  // Participants_validés
  let shPart = ss.getSheetByName("Participants_validés");
  if (!shPart) {
    shPart = ss.insertSheet("Participants_validés");
    lines.push("💡 Feuille 'Participants_validés' créée → refais 'Convertir en tableau' manuellement !");
    LOG.info("Feuille 'Participants_validés' créée.");
  }
  const partHeaders = [
    "Email","Pseudo","Participants",
    // "Fromage","Charcuterie","Mixte",
    "Montant","Total payé",
    "Statut Mail","Statut Paiement","Dernière MAJ"
  ];
  ensureHeaders(shPart, partHeaders);

  // Participants_annulé
  let shAnnul = ss.getSheetByName("Participants_annulés");
  if (!shAnnul) {
    shAnnul = ss.insertSheet("Participants_annulés");
    lines.push("💡 Feuille 'Participants_annulés' créée → refais 'Convertir en tableau' manuellement !");
    LOG.info("Feuille 'Participants_annulés' créée.");
  }
  const annulHeaders = [
    "Email","Pseudo","Total payé"
  ];
  ensureHeaders(shAnnul, annulHeaders);

  // Paiements
  let shPay = ss.getSheetByName("Paiements");
  if (!shPay) {
    shPay = ss.insertSheet("Paiements");
    lines.push("💡 Feuille 'Paiements' créée → refais 'Convertir en tableau' manuellement !");
    LOG.info("Feuille 'Paiements' créée.");
  }
  const payHeaders = ["Date","Email","Montant","OrderID","TransactionID","Statut"];
  ensureHeaders(shPay, payHeaders);

  // Tarifs
  let shTarifs = ss.getSheetByName("Tarifs");
  if (!shTarifs) {
    shTarifs = ss.insertSheet("Tarifs");
    shTarifs.appendRow(["Produit", "Description", "Prix"]);
    shTarifs.appendRow(["Entrée", "Prix d'une entrée", 35]);
    shTarifs.appendRow(["Fromage", "Une planche de fromage", 7]);
    shTarifs.appendRow(["Charcuterie", "Une planche de charcuterie", 7]);
    shTarifs.appendRow(["Mixte", "Une planche mixte avec fromage et charcuterie", 12]);
    lines.push("💡 Feuille 'Tarifs' créée → refais 'Convertir en tableau' manuellement !");
    LOG.info("Feuille 'Tarifs' créée.");
  } else {
    ensureHeaders(shTarifs, ["Produit", "Description", "Prix"]);
  }


  // ---- Réponses : s'assurer des colonnes + validations (puis rappel chips/multi)
  const sheet = ss.getSheets()[0];
  sheet.setName("Réponses");

  let shRep = ss.getSheetByName("Réponses");
  if (!shRep) throw new Error("❌ ensureHeaders: feuille manquante");
  
  const repHeadersToSet = ["Validation","Action","Traité"];
  const lastRowRep = Math.max(2, shRep.getLastRow());

  for(let i = 0; i < repHeadersToSet.length; i++) {
    const lastColRep = Math.max(2,shRep.getLastColumn())
    const repHeaders = shRep.getRange(1, 1, 1, lastColRep).getValues()[0];
    let idx = repHeaders.indexOf(repHeadersToSet[i]);
    if (idx > -1) continue;
    
    idx = ensureColumnExists_(shRep, repHeadersToSet[i])

    if (repHeadersToSet[i] == "Validation") {
      try {
        setListValidation_(shRep.getRange(2, idx + 1, lastRowRep - 1, 1), ["Validé","Refusé","En attente"]);
      } catch (e) {
        LOG.info(`[init] Skip setDataValidation Réponses!Validation: ${e}`);
      }
      shRep.getRange(1, idx + 1).setNote("Astuce: Format > Liste déroulante → Style = 'Chips' (option visuelle, pas obligatoire).");
    }
    if (repHeadersToSet[i] == "Action") {
      try {
        setListValidation_(shRep.getRange(2, idx + 1, lastRowRep - 1, 1), ["PEOPLE+","PEOPLE=","CANCEL"]);
      } catch (e) {
        LOG.info(`[init] Skip setDataValidation Réponses!Action: ${e}`);
      }
      shRep.getRange(1, idx + 1).setNote("Astuce: Format > Liste déroulante → Style = 'Chips' et activer 'Plusieurs sélections'.");
    }
    if (repHeadersToSet[i] === "Traité") {
      const h = Math.max(0, shRep.getMaxRows() - 1);
      if (h > 0) shRep.getRange(2, idx + 1, h, 1).clearDataValidations();
      shRep.getRange(1, idx + 1).setNote("Marqué automatiquement (OK / IGN). Aucune validation.");
    }
  }

  // ensureAideSheet_(); // crée la feuille Aide si absente
  // lines.push("✅ Vérification/initialisation terminée (+Aide)");

  // Toast + log récap (via LOG)
  LOG.infoToast("Init Sheets", lines, 10);
}

/***********************
 * Purges avec init()  *
 ***********************/
function purgeParticipants() {
  const out = [];
  out.push(safeClearDataRows_("Participants_validés"));
  out.push(safeClearDataRows_("Participants_annulés"));

  // Réinitialiser la colonne "Traité" dans Réponses (si elle existe)
  _clearTraiteInReponses_();

  // Réassurer la structure/entêtes
  try { initSheets(); } catch (e) { LOG.error('initSheets error: ' + e); }

  LOG.infoToast("Purge Participants", out, 6);
}

function purgeParticipantsEtPaiements() {
  const out = [];
  out.push(safeClearDataRows_("Participants_validés"));
  out.push(safeClearDataRows_("Participants_annulés"));
  out.push(safeClearDataRows_("Paiements"));
  _clearTraiteInReponses_();

  initSheets();
  LOG.infoToast("Purge Participants + Paiements", out, 6);
}

function purgeComplet() {
  const out = [];
  out.push(safeClearDataRows_("Participants_validés"));
  out.push(safeClearDataRows_("Participants_annulés"));
  out.push(safeClearDataRows_("Paiements"));
  out.push(safeClearDataRows_("Réponses")); // Feuille liée au Form: on ne la supprime pas

  initSheets();
  LOG.infoToast("Purge complète", out, 8);
}

/** Crée/écrase la feuille "Aide" avec checklist + guide. */
function regenerateAideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Aide';
  const shOld = ss.getSheetByName(name);
  if (shOld) ss.deleteSheet(shOld);
  const sh = ss.insertSheet(name);

  // --- Données système
  const creds = getCredentials(); // helpers.gs
  const env = (PropertiesService.getScriptProperties().getProperty("ENV") || "dev").toLowerCase() === "prod" ? "LIVE" : "SANDBOX";
  const paypalBase = PAYPAL.baseUrl(); // paypal.gs
  const cid = creds && creds.clientId ? String(creds.clientId) : "";
  const cidPrefix = cid ? cid.slice(0, 10) : "";
  const hasWrapper = !!(creds && creds.wrapperUrl);
  const hasWebapp = !!(creds && creds.webappUrl);

  // --- Mise en page de base
  let r = 1;
  sh.getRange(r,1).setValue('Aide • QA — Munch Admin'); r++;
  sh.getRange(r,1).setValue('Tableau de bord QA : infos système, KPIs, scénarios de test et rappels.'); r+=2;

  // ===== Bloc Infos Système
  sh.getRange(r,1).setValue('🧩 Infos système'); r++;
  sh.getRange(r,1,5,2).setValues([
    ['ENV', env],
    ['PayPal base', paypalBase],
    ['ClientID (préfixe)', cidPrefix || '—'],
    ['WRAPPER_URL défini', hasWrapper ? 'oui' : 'non'],
    ['WEBAPP_URL défini', hasWebapp ? 'oui' : 'non'],
  ]);
  r+=6;

  // ===== Bloc KPI
  sh.getRange(r,1).setValue('📊 Indicateurs'); r++;
  sh.getRange(r,1,5,2).setValues([
    ['Participants (lignes)', '=MAX(ROWS(Participants_validés!A2:A)-COUNTBLANK(Participants_validés!A2:A);0)'],
    ['Total dû (Montant)',    '=IFERROR(SUM(Participants_validés!G2:G);0)'],
    ['Total payé (SUMIF)',    '=IFERROR(SUM(Paiements!C2:C);0)'],
    ['Paiements (COMPLETED)', '=IFERROR(COUNTIF(Paiements!F2:F;"COMPLETED");0)'],
    ['Statuts paiement',      'NONE / PARTIAL / COMPLETED / OVERPAID (voir Participants_validés!I)'],
  ]);
  r+=6;

  // ===== Checklist QA (progression)
  sh.getRange(r,1).setValue('✅ Checklist QA (progression)'); 
  sh.getRange(r,3).setValue('Progression').setFontWeight('bold');
  // Progression = % de cases cochées dans col A (à partir de la section ci-dessous)
  const progressStartRow = r + 3; // première ligne de la checklist
  const SEP = _sep_();
  sh.getRange(r,4).setFormula(`=IFERROR(ROUND(100*COUNTIF(A${progressStartRow}:A${SEP}TRUE)/COUNTA(A${progressStartRow}:A)${SEP}0)&" %"${SEP}"—")`);
  r+=2;

  // En-têtes checklist
  sh.getRange(r,1,1,4).setValues([['Fait','Phase','Scénario','Attendu / Notes']]).setFontWeight('bold');
  r++;

  // Lignes checklist — STRUCTURÉES
  const rows = [];

  // --- Phase 1: Setup & Init
  rows.push([false,'Setup','Init Sheets','Crée/valide Participants_validés, Paiements, Tarifs + rappels chips. Menu → 📐 Initialiser les feuilles.']);
  rows.push([false,'Setup','Tarifs présents','Fromage=7, Charcuterie=7, Mixte=12 (ou vos valeurs).']);
  rows.push([false,'Setup','Audit complet sans erreurs critiques','Menu → 🔍 Audit complet (en-têtes, validations, formules).']);

  // --- Phase 2: Consolidation & Email
  rows.push([false,'Consolidation','Réponse Validé sans planches','Ligne créée, Montant=0, email envoyé sans bouton PayPal.']);
  rows.push([false,'Consolidation','Réponse Validé avec planches 1F / 0C / 0M','Montant calculé, email avec lien PayPal, texte planche singulier.']);
  rows.push([false,'Consolidation','Réponse Validé avec planches 1F / 1C / 0M','Montant calculé, email avec lien PayPal, texte planches pluriel.']);
  rows.push([false,'Consolidation','Idempotence via "Traité"','Relancer 📤 Consolider ne duplique rien (Réponses marquées OK/IGN).']);
  rows.push([false,'Consolidation','Validation ≠ "Validé"','Réponse traitée en IGN, pas d’email ni consolidation.']);

  // --- Phase 3: Actions (modif/ajout/annulation)
  rows.push([false,'Actions','Action=PLANCHES= (remplace)','Quantités Fromage/Charcuterie/Mixte remplacées, statuts mail/paiement remis à vide.']);
  rows.push([false,'Actions','Action=PLANCHES+ (ajoute)','Quantités incrémentées, Montant mis à jour.']);
  rows.push([false,'Actions','Action=PEOPLE= (fixe)','Colonne Participants = valeur du formulaire.']);
  rows.push([false,'Actions','Action=PEOPLE+ (ajoute)','Colonne Participants += valeur du formulaire.']);
  rows.push([false,'Actions','Action=CANCEL','Ligne marquée Annulé (Statut Mail + Statut Paiement), "Traité"=OK.']);

  // --- Phase 4: Emails (cas Montant vs Total payé)
  rows.push([false,'Email','Montant>0 → bouton PayPal visible','Gabarit avec section planches + CTA "Payer avec PayPal".']);
  rows.push([false,'Email','Montant=0 → pas de bouton','Email confirmation sans paiement.']);
  rows.push([false,'Email','Total payé < Montant (PARTIAL)','Email mis à jour selon ton rendu, statut attendu PARTIAL.']);
  rows.push([false,'Email','Total payé = Montant (COMPLETED)','Email one-shot OK, statut COMPLETED.']);
  rows.push([false,'Email','Total payé > Montant (OVERPAID)','Comportement maîtrisé dans le rendu; statut OVERPAID.']);

  // --- Phase 5: Paiements (WebApp PayPal)
  rows.push([false,'Paiement','CREATE: email connu + Montant>0','approveUrl généré (SANDBOX).']);
  rows.push([false,'Paiement','CREATE: email inconnu / Montant<=0','Erreur JSON côté WebApp (create).']);
  rows.push([false,'Paiement','CAPTURE: COMPLETED → écrit dans Paiements','Ligne (Date, Email, Montant, OrderID, TransactionID, Statut).']);
  rows.push([false,'Paiement','Idempotence CAPTURE (re-jouer capture)','Aucun doublon (TransactionID / OrderID).']);
  rows.push([false,'Paiement','Statut final via formules','Participants_validés!I passe à COMPLETED quand SUMIF ≥ dû.']);

  // --- Phase 6: Edge cases & régression
  rows.push([false,'Edges','Modifier les prix dans Tarifs','Montant se recalcule, email reflète les nouveaux prix.']);
  rows.push([false,'Edges','Doublons email (multi-inscriptions)','Audit signale DOUBLON, consolidation reste cohérente.']);
  rows.push([false,'Edges','Email invalide / vide','Audit signale EMAIL_VIDE (ou à nettoyer).']);
  rows.push([false,'Régression','Réparer les formules (menu)','Pose/repare Montant/Total payé/Statut Paiement, format 0.00.']);
  rows.push([false,'Régression','Relire "Aide" depuis le menu','(Re)génère ce dashboard QA.']);

  const startChecklist = r;
  sh.getRange(startChecklist,1,rows.length,4).setValues(rows);
  sh.getRange(startChecklist,1,rows.length,1).insertCheckboxes();
  r += rows.length + 2;

  // ===== Rappels & Liens utiles
  sh.getRange(r,1).setValue('ℹ️ Rappels & raccourcis');
  sh.getRange(r+1,1,7,2).setValues([
    ['Menu', 'Munch • Admin → 📤 Consolider & envoyer confirmations'],
    ['Menu', 'Munch • Admin → 📐 Initialiser les feuilles'],
    ['Menu', 'Munch • Admin → 🔍 Audit complet'],
    ['Menu', 'Munch • Admin → 🛠️ Réparer les formules'],
    ['Menu', 'Munch • Admin → 🧹 Purges (selon besoin)'],
    ['Note', 'Les statuts paiement et le total payé proviennent de formules (SUMIF/logic).'],
    ['Note', 'Après changement d’en-têtes, relancer Init puis Réparer les formules.']
  ]);
  r += 10;

  // ===== Mise en forme rapide
  sh.setFrozenRows(3);
  sh.autoResizeColumn(1);
  sh.setColumnWidths(2, 3, 320);

  // Bandeaux
  sh.getRange(1,1,1,4).setFontSize(16).setFontWeight('bold').setBackground('#F1F5F9');
  sh.getRange(3,1,1,4).setFontWeight('bold').setBackground('#F8FAFC');          // Infos système
  sh.getRange(9,1,1,4).setFontWeight('bold').setBackground('#F8FAFC');          // KPI
  const checklistHeaderRow = progressStartRow - 1;
  sh.getRange(checklistHeaderRow,1,1,4).setBackground('#F8FAFC');               // Checklist header

  // Notes utiles
  sh.getRange(2,4).setNote("Cette feuille est régénérée via le menu 'Munch • Admin'.");
  sh.getRange(4,2).setNote("Ces infos viennent des Script Properties et des modules PAYPAL/helpers.");
  sh.getRange(checklistHeaderRow,3).setNote("Coche la colonne A → progression mise à jour automatiquement.");

  SpreadsheetApp.flush();
  LOG.toast("Feuille 'Aide' (Dashboard QA) régénérée.", "Aide • Munch Admin", 6);
}

/** Appelle regenerateAideSheet() seulement si la feuille n’existe pas */
function ensureAideSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Aide')) {
    regenerateAideSheet();
  }
}

/** Efface le contenu de la colonne 'Traité' (si présente) dans la feuille Réponses */
function _clearTraiteInReponses_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Réponses');
  if (!sh) { LOG.info("Réponses introuvable : skip reset 'Traité'"); return false; }

  const data = sh.getDataRange().getValues();
  if (data.length === 0) { LOG.info("Réponses vide : skip reset 'Traité'"); return false; }

  const headers = data[0] || [];
  const idxTraite = headers.indexOf('Traité');
  if (idxTraite === -1) { LOG.info("'Traité' absent de Réponses : rien à réinitialiser"); return false; }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) { LOG.info("Réponses sans données : rien à réinitialiser"); return true; }

  // Efface uniquement le contenu (garde la colonne et l'en-tête)
  sh.getRange(2, idxTraite + 1, lastRow - 1, 1).clearContent();
  LOG.info(`'Traité' réinitialisé dans Réponses (${lastRow - 1} lignes).`);
  return true;
}
