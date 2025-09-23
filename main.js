const creds = getCredentials();
const WRAPPER_URL = creds.wrapperUrl;

// ===== Consolidation : ne traiter que les lignes non "Traité"
function consolidation() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { LOG.warn("Lock indisponible"); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shReponse   = ss.getSheetByName("Réponses");
    const shConsolide = ss.getSheetByName("Participants_validés");
    const shCancel    = ss.getSheetByName("Participants_annulés");

    if (!shReponse || !shConsolide || !shCancel) {
      LOG.error("Feuilles manquantes (Réponses ou Participants_validés ou Participants_annulés)");
      return;
    }

    const data = shReponse.getDataRange().getValues();
    if (data.length < 2) { LOG.info("Aucune réponse."); return; }
    const headers = data[0];

    // Index colonnes Réponses
    const idxEmail       = headers.indexOf("Votre email (pour la confirmation)");
    const idxPseudo      = headers.indexOf("Pseudo");
    const idxPeople      = headers.indexOf("A combien venez-vous ? (Si vous venez à plus de 2 merci de re-remplir le formulaire )");
    // const idxPlanche     = headers.indexOf("Souhaitez-vous réserver une (ou plusieurs) planches ?");
    // const idxFromage     = headers.indexOf("Quelle(s) planche(s) souhaitez-vous réserver ? [Planche Fromage (7€)]");
    // const idxCharcuterie = headers.indexOf("Quelle(s) planche(s) souhaitez-vous réserver ? [Planche Charcuterie (7€)]");
    // const idxMixte       = headers.indexOf("Quelle(s) planche(s) souhaitez-vous réserver ? [Planche Mixte (12€)]");
    const idxValidation  = headers.indexOf("Validation");
    const idxAction      = headers.indexOf("Action");
    const idxTraite      = headers.indexOf("Traité");

    // if ([idxEmail, idxPseudo,idxPeople, idxPlanche, idxFromage, idxCharcuterie, idxMixte, idxValidation, idxAction, idxTraite, ].some(i => i === -1)) {
    //   LOG.error("Colonnes indispensables manquantes dans 'Réponses'");
    //   return;
    // }
    if ([idxEmail, idxPseudo,idxPeople, idxValidation, idxAction, idxTraite, ].some(i => i === -1)) {
      LOG.error("Colonnes indispensables manquantes dans 'Réponses'");
      return;
    }

    // En-têtes Participants_validés
    const consDataHead        = shConsolide.getDataRange().getValues()[0];
    const idxConsEmail        = consDataHead.indexOf("Email");
    const idxConsPseudo       = consDataHead.indexOf("Pseudo");
    const idxConsParticipants = consDataHead.indexOf("Participants");
    // const idxConsFromage      = consDataHead.indexOf("Fromage");
    // const idxConsCharcuterie  = consDataHead.indexOf("Charcuterie");
    // const idxConsMixte        = consDataHead.indexOf("Mixte");
    const idxConsMontant      = consDataHead.indexOf("Montant");
    const idxConsTotalPaye    = consDataHead.indexOf("Total payé");
    const idxConsStatutMail   = consDataHead.indexOf("Statut Mail");
    const idxConsStatutPayt   = consDataHead.indexOf("Statut Paiement");

    // if ([idxConsEmail, idxConsPseudo, idxConsParticipants, idxConsFromage, idxConsCharcuterie, idxConsMixte, idxConsMontant, idxConsTotalPaye, idxConsStatutMail, idxConsStatutPayt].some(i => i === -1)) {
    //   LOG.error("Colonnes indispensables manquantes dans 'Participants_validés'");
    //   return;
    // }
    if ([idxConsEmail, idxConsPseudo, idxConsParticipants, idxConsMontant, idxConsTotalPaye, idxConsStatutMail, idxConsStatutPayt].some(i => i === -1)) {
      LOG.error("Colonnes indispensables manquantes dans 'Participants_validés'");
      return;
    }

    // En-têtes Participants_annulés
    const cancelDataHead     = shCancel.getDataRange().getValues()[0];
    const idxCancelEmail     = cancelDataHead.indexOf("Email");
    const idxCancelPseudo    = cancelDataHead.indexOf("Pseudo");
    const idxCancelTotalPaye = cancelDataHead.indexOf("Total payé");

    if ([idxCancelEmail, idxCancelPseudo, idxCancelTotalPaye].some(i => i === -1)) {
      LOG.error("Colonnes indispensables manquantes dans 'Participants_annulés'");
      return;
    }

    // ===== Consolidation
    for (let i = 1; i < data.length; i++) {
      const dejaTraite = data[i][idxTraite];
      if (dejaTraite) continue;

      const validation = data[i][idxValidation];
      if (validation !== "Validé") { continue; }

      const email   = String(data[i][idxEmail]  || '').trim();
      const pseudo  = data[i][idxPseudo];
      const actions = parseActions_((data[i][idxAction] || ''));

      if (actions.length === 0) {
        // actions.push('PLANCHES=');
        actions.push('PEOPLE=');
      }

      // let plFromage = 0, plCharcuterie = 0, plMixte = 0;
      // if (idxPlanche !== -1 && data[i][idxPlanche] === "Oui") {
      //   plFromage     = getInt_(data[i][idxFromage]     || 0);
      //   plCharcuterie = getInt_(data[i][idxCharcuterie] || 0);
      //   plMixte       = getInt_(data[i][idxMixte]       || 0);
      // }
      let nbPeople = getInt_(data[i][idxPeople] || 0);

      // Chercher le participant existant
      const consolData = shConsolide.getDataRange().getValues();
      let idxConsol = -1;
      for (let j = 1; j < consolData.length; j++) {
        if (String(consolData[j][idxConsEmail]).trim().toLowerCase() === email.toLowerCase()) { idxConsol = j; break; }
      }

      // === CANCEL ===
      if (actions.includes('CANCEL')) {
        //existe déjà dans Participants_annulés ?
        const cancData = shCancel.getDataRange().getValues();
        let rCancel = -1;
        for (let rr = 1; rr < cancData.length; rr++) {
          if (String(cancData[rr][idxCancelEmail]).trim().toLowerCase() === email.toLowerCase()) { rCancel = rr + 1; break; }
        }

        if (rCancel === -1) {
          const newRowArr = new Array(shCancel.getLastColumn()).fill("");
          newRowArr[idxCancelEmail]  = email;
          newRowArr[idxCancelPseudo] = pseudo || "";
          shCancel.appendRow(newRowArr);
          rCancel = shCancel.getLastRow();
          ensureTotalPayeFormula_(shCancel, rCancel, idxCancelTotalPaye + 1); // =IFNA(SUMIF(Paiements!B:B; A{row}; Paiements!C:C); 0)
          shCancel.getRange(rCancel, idxCancelTotalPaye + 1).setNumberFormat('0.00');
        } // sinon: on ne touche rien (tu as demandé un simple "return" logique)

        // Supprimer des validés si présent
        if (idxConsol > 0) shConsolide.deleteRow(idxConsol + 1);

        // Traiter la réponse
        shReponse.getRange(i + 1, idxTraite + 1).setValue("OK");
        LOG.info("Annulation : déplacé vers 'Participants_annulés' puis supprimé des validés", { email });
        continue;
      }

      // Construire/mettre à jour la ligne
      let row;
      if (idxConsol > 0) {
        row = consolData[idxConsol].slice();
      } else {
        row = new Array(consDataHead.length).fill("");
        row[idxConsEmail]        = email;
        row[idxConsPseudo]       = pseudo;
        // row[idxConsFromage]      = 0;
        // row[idxConsCharcuterie]  = 0;
        // row[idxConsMixte]        = 0;
        row[idxConsParticipants] = 0;
      }

      // PEOPLE
      if (actions.some(a => a.startsWith('PEOPLE'))) {
        if (actions.includes('PEOPLE=')) row[idxConsParticipants] = nbPeople;
        if (actions.includes('PEOPLE+')) row[idxConsParticipants] = getInt_(row[idxConsParticipants]) + nbPeople;
        row[idxConsStatutMail]  = "";
      }
      // if (actions.includes('PLANCHES=')) {
      //   row[idxConsFromage]     = plFromage;
      //   row[idxConsCharcuterie] = plCharcuterie;
      //   row[idxConsMixte]       = plMixte;
      //   row[idxConsStatutMail]  = "";
      // }
      // if (actions.includes('PLANCHES+')) {
      //   row[idxConsFromage]     = getInt_(row[idxConsFromage])     + plFromage;
      //   row[idxConsCharcuterie] = getInt_(row[idxConsCharcuterie]) + plCharcuterie;
      //   row[idxConsMixte]       = getInt_(row[idxConsMixte])       + plMixte;
      //   row[idxConsStatutMail]  = "";
      // }

      // Écriture Participants_validés + formules
      if (idxConsol > 0) {
        shConsolide.getRange(idxConsol + 1, 1, 1, row.length).setValues([row]);
        ensureMontantFormula_(shConsolide, idxConsol + 1, idxConsMontant + 1);
        ensureTotalPayeFormula_(shConsolide, idxConsol + 1, idxConsTotalPaye + 1);
        ensureStatutPaiementFormula_(shConsolide, idxConsol + 1, idxConsStatutPayt + 1);
      } else {
        shConsolide.appendRow(row);
        const r = shConsolide.getLastRow();
        ensureMontantFormula_(shConsolide, r, idxConsMontant + 1);
        ensureTotalPayeFormula_(shConsolide, r, idxConsTotalPaye + 1);
        ensureStatutPaiementFormula_(shConsolide, r, idxConsStatutPayt + 1);
      }
      shReponse.getRange(i + 1, idxTraite + 1).setValue("OK");
    }
    LOG.info("Consolidation terminée.");
  } finally {
    lock.releaseLock();
  }
}

// ===== Envoi des mails
function sendMail() {
  const mailLock = LockService.getScriptLock();
  if (!mailLock.tryLock(30000)) { LOG.warn("Lock indisponible"); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shConsolide = ss.getSheetByName("Participants_validés");

    const consDataHead        = shConsolide.getDataRange().getValues()[0];
    const idxConsEmail        = consDataHead.indexOf("Email");
    const idxConsPseudo       = consDataHead.indexOf("Pseudo");
    const idxConsParticipants = consDataHead.indexOf("Participants");
    // const idxConsFromage      = consDataHead.indexOf("Fromage");
    // const idxConsCharcuterie  = consDataHead.indexOf("Charcuterie");
    // const idxConsMixte        = consDataHead.indexOf("Mixte");
    const idxConsMontant      = consDataHead.indexOf("Montant");
    const idxConsTotalPaye    = consDataHead.indexOf("Total payé");
    const idxConsStatutMail   = consDataHead.indexOf("Statut Mail");
    const idxConsStatutPayt   = consDataHead.indexOf("Statut Paiement");

    const consolFinal = shConsolide.getDataRange().getValues();
    const aliases     = GmailApp.getAliases();

    for (let i = 1; i < consolFinal.length; i++) {
      const statutMail = consolFinal[i][idxConsStatutMail];
      if (statutMail) continue;

      const email  = consolFinal[i][idxConsEmail];
      const pseudo = consolFinal[i][idxConsPseudo];

      ensureMontantFormula_(shConsolide, i + 1, idxConsMontant + 1);
      ensureTotalPayeFormula_(shConsolide, i + 1, idxConsTotalPaye + 1);
      ensureStatutPaiementFormula_(shConsolide, i + 1, idxConsStatutPayt + 1);
      SpreadsheetApp.flush();


      const montant = normalizeAmount(shConsolide.getRange(i + 1, idxConsMontant + 1).getValue());
      const totalPaye = normalizeAmount(shConsolide.getRange(i + 1, idxConsTotalPaye + 1).getValue());
      const nbrPeople = shConsolide.getRange(i + 1, idxConsParticipants + 1).getValue();
      // const prixUnit = getBoardPrices_();
      // const qFromage = normalizeAmount(consolFinal[i][idxConsFromage]     || 0);
      // const qCharc   = normalizeAmount(consolFinal[i][idxConsCharcuterie] || 0);
      // const qMixte   = normalizeAmount(consolFinal[i][idxConsMixte]       || 0);

      const t = HtmlService.createTemplateFromFile("email");
      t.data = {
        pseudo: pseudo,
        nbPeople : nbrPeople,
        montant: montant,
        totalPaye: totalPaye,
        paymentLink: WRAPPER_URL + "?email=" + encodeURIComponent(email),
        // qte: { Fromage: qFromage, Charcuterie: qCharc, Mixte: qMixte },
        // prixFmt: {
        //   Fromage: formatPrice_(prixUnit.Fromage),
        //   Charcuterie: formatPrice_(prixUnit.Charcuterie),
        //   Mixte: formatPrice_(prixUnit.Mixte)
        // }
      };
      const emailContent = t.evaluate().getContent();

      GmailApp.sendEmail(
        email,
        "Confirmation d'inscription à la Nuit Démoniaque",
        emailContent,
        {
          from: (aliases && aliases.length ? aliases[0] : Session.getActiveUser().getEmail()),
          name: "Les Jeux Démoniaques",
          htmlBody: emailContent
        }
      );
      shConsolide.getRange(i + 1, idxConsStatutMail + 1).setValue(new Date());
    }
    LOG.info("Envoi des mail terminés.");
  } finally {
    mailLock.releaseLock();
  }
}
