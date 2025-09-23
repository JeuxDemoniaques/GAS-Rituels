/***********************
 * AUDIT COMPLET MUNCH *
 ***********************/
function _normFormula_(s){ return String(s||'').replace(/\s+/g,'').toUpperCase(); }

function _buildMontantFormula_R1C1_() {
  const S = _sep_();
  // R1C1 indépendante de la colonne : on utilisera setFormulaR1C1 au besoin,
  // pour l’audit on compare la version A1 recalculée sur la cellule.
  return (
    `=INDEX($1:$9999${S} ROW()${S} MATCH("Fromage"${S} $1:$1${S} 0)) * VLOOKUP("Fromage"${S} Tarifs!$A:$C${S} 3${S} FALSE)` +
    ` + INDEX($1:$9999${S} ROW()${S} MATCH("Charcuterie"${S} $1:$1${S} 0)) * VLOOKUP("Charcuterie"${S} Tarifs!$A:$C${S} 3${S} FALSE)` +
    ` + INDEX($1:$9999${S} ROW()${S} MATCH("Mixte"${S} $1:$1${S} 0)) * VLOOKUP("Mixte"${S} Tarifs!$A:$C${S} 3${S} FALSE)`
  );
}
function _buildTotalPayeFormula_A1_(row){ // A{row} = email
  const S=_sep_();
  return `=IFNA(SUMIF(Paiements!B:B${S} A${row}${S} Paiements!C:C)${S} 0)`;
}
function _buildStatutFormula_R1C1_(offTot, offMont){ // offsets relatifs depuis Statut
  const S=_sep_();
  return `=IF(R[0]C[${offTot}]=0${S}"NONE"${S}` +
         `IF(R[0]C[${offTot}]<R[0]C[${offMont}]${S}"PARTIAL"${S}` +
         `IF(R[0]C[${offTot}]=R[0]C[${offMont}]${S}"COMPLETED"${S}"OVERPAID")))`;
}

function auditComplet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];

  // --- 0) Vérif présence feuilles
  const shPart = ss.getSheetByName('Participants_validés');
  const shPay  = ss.getSheetByName('Paiements');
  const shRep  = ss.getSheetByName('Réponses');

  if (!shPart) out.push(['GLOBAL','Participants_validés','ABSENT',"Créer la feuille via initSheets()"]);
  if (!shPay)  out.push(['GLOBAL','Paiements','ABSENT',"Créer la feuille via initSheets()"]);
  if (!shRep)  out.push(['GLOBAL','Réponses','ABSENT',"Connecter votre Google Form"]);

  if (!shPart || !shPay) {
    _writeAuditSheet_(out);
    LOG.info("Audit terminé (feuilles manquantes). Voir onglet 'Audit'.", 'Munch • Admin', 6);
    LOG.toast("Audit terminé (feuilles manquantes). Voir onglet 'Audit'.", 'Munch • Admin', 6);
    return;
  }

  // --- 1) En-têtes attendus
  const Hpart = shPart.getRange(1,1,1,shPart.getLastColumn()).getValues()[0];
  const needPart = ["Email","Pseudo","Fromage","Charcuterie","Mixte","Montant","Total payé","Statut Mail","Statut Paiement","Dernière MAJ"];
  const missPart = needPart.filter(h=>Hpart.indexOf(h)===-1);
  if (missPart.length) out.push(['HEADERS','Participants_validés','MANQUANTS',missPart.join(', ')]);

  const Hpay = shPay.getLastRow()? shPay.getRange(1,1,1,shPay.getLastColumn()).getValues()[0] : [];
  const needPay = ["Date","Email","Montant","OrderID","TransactionID","Statut"];
  const missPay = needPay.filter(h=>Hpay.indexOf(h)===-1);
  if (missPay.length) out.push(['HEADERS','Paiements','MANQUANTS',missPay.join(', ')]);

  // Stop si trop de manque
  if (missPart.length || missPay.length) {
    _writeAuditSheet_(out);
    LOG.warn("Audit : en-têtes manquants. Voir onglet 'Audit'.");
    LOG.toast("Audit : en-têtes manquants. Voir onglet 'Audit'.", 'Munch • Admin', 6);
    return;
  }

  // --- 2) Audit Participants_validés (lignes)
  const lastP = shPart.getLastRow();
  if (lastP >= 2) {
    const idx = {
      Email: Hpart.indexOf('Email')+1,
      Pseudo: Hpart.indexOf('Pseudo')+1,
      Fromage: Hpart.indexOf('Fromage')+1,
      Charcuterie: Hpart.indexOf('Charcuterie')+1,
      Mixte: Hpart.indexOf('Mixte')+1,
      Montant: Hpart.indexOf('Montant')+1,
      Total: Hpart.indexOf('Total payé')+1,
      SMail: Hpart.indexOf('Statut Mail')+1,
      SPay: Hpart.indexOf('Statut Paiement')+1
    };

    const vals = shPart.getRange(2,1,lastP-1, shPart.getLastColumn()).getValues();
    const emails = [];
    const emailCount = {};

    // pré-bâtir formules attendues
    const wantMont = _buildMontantFormula_R1C1_();

    for (let r = 0; r < vals.length; r++) {
      const row = r+2;
      const v = vals[r];
      const email = String(v[idx.Email-1]||'').trim().toLowerCase();
      emails.push(email);
      emailCount[email] = (emailCount[email]||0)+1;

      // Email manquant
      if (!email) out.push(['PARTICIPANTS',`L${row}`,'EMAIL_VIDE','Renseigner un email']);

      // Quantités (entiers >=0)
      [['Fromage',idx.Fromage],['Charcuterie',idx.Charcuterie],['Mixte',idx.Mixte]].forEach(([lab,c])=>{
        const q = v[c-1];
        if (q==='' || q==null) return; // vide = ok
        if (typeof q !== 'number' || !Number.isFinite(q)) out.push(['PARTICIPANTS',`L${row}`,`QTE_${lab}_NON_NUM`,'Mettre un nombre']);
        else if (q<0) out.push(['PARTICIPANTS',`L${row}`,`QTE_${lab}_NEG`,'Remplacer par 0 ou plus']);
        else if (!Number.isInteger(q)) out.push(['PARTICIPANTS',`L${row}`,`QTE_${lab}_NON_ENTIER`,'Arrondir à l’entier']);
      });

      // Formule Montant
      const fMontCell = shPart.getRange(row, idx.Montant).getFormulaR1C1();
      if (_normFormula_(fMontCell) !== _normFormula_(wantMont)) {
        out.push(['FORMULA',`L${row} Montant`,'FORMULE_DIFFERENTE',"Utiliser l’outil 'Réparer les formules'"]);
      }

      // Formule Total payé (A1 spécifique ligne)
      const wantTotA1 = _buildTotalPayeFormula_A1_(row);
      const fTot = shPart.getRange(row, idx.Total).getFormula();
      if (_normFormula_(fTot) !== _normFormula_(wantTotA1)) {
        out.push(['FORMULA',`L${row} Total payé`,'FORMULE_DIFFERENTE',"Utiliser l’outil 'Réparer les formules'"]);
      }

      // Formule Statut Paiement (R1C1 offsets)
      const offTot  = idx.Total - idx.SPay;   // ex: -3
      const offMont = idx.Montant - idx.SPay; // ex: -4
      const wantStat = _buildStatutFormula_R1C1_(offTot, offMont);
      const fStat = shPart.getRange(row, idx.SPay).getFormulaR1C1();
      if (_normFormula_(fStat) !== _normFormula_(wantStat)) {
        out.push(['FORMULA',`L${row} Statut Paiement`,'FORMULE_DIFFERENTE',"Utiliser l’outil 'Réparer les formules'"]);
      }
    }

    // Doublons email (info)
    Object.entries(emailCount).forEach(([mail,cnt])=>{
      if (mail && cnt>1) out.push(['PARTICIPANTS','Emails','DOUBLON', `${mail} apparaît ${cnt} fois (vérifie si souhaité)`]);
    });
  }

  // --- 3) Audit Paiements
  const lastPay = shPay.getLastRow();
  if (lastPay >= 2) {
    const H = shPay.getRange(1,1,1,shPay.getLastColumn()).getValues()[0];
    const idx = {
      Date: H.indexOf('Date')+1,
      Email: H.indexOf('Email')+1,
      Montant: H.indexOf('Montant')+1,
      Order: H.indexOf('OrderID')+1,
      Txn: H.indexOf('TransactionID')+1,
      Statut: H.indexOf('Statut')+1
    };
    const vals = shPay.getRange(2,1,lastPay-1, shPay.getLastColumn()).getValues();

    const setTxn = new Set();
    const setOrder = new Set();

    for (let r=0; r<vals.length; r++) {
      const row = r+2;
      const v = vals[r];
      const m = v[idx.Montant-1];
      if (m!=='' && m!=null && typeof m !== 'number') {
        out.push(['PAIEMENTS',`L${row}`,'MONTANT_NON_NUM',"Mettre un nombre (ex: 10.00)"]);
      }
      const txn = String(v[idx.Txn-1]||'').trim();
      const ord = String(v[idx.Order-1]||'').trim();
      if (txn) {
        if (setTxn.has(txn)) out.push(['PAIEMENTS',`L${row}`,'TXN_DOUBLON', txn]);
        setTxn.add(txn);
      }
      if (ord) {
        if (setOrder.has(ord)) out.push(['PAIEMENTS',`L${row}`,'ORDER_DOUBLON', ord]);
        setOrder.add(ord);
      }
    }
  }

  // --- 4) Vérifications spécifiques "Réponses" : colonnes + validations
  if (shRep) {
    const headersRep = shRep.getRange(1, 1, 1, Math.max(1, shRep.getLastColumn())).getValues()[0];

    // Présence colonnes
    const needRep = ["Validation","Action","Traité"];
    needRep.forEach(h => {
      if (headersRep.indexOf(h) === -1) {
        out.push(['HEADERS','Réponses','MANQUANT', `Colonne '${h}' absente`]);
      }
    });

    // Validations listes
    const checks = [
      { name: 'Validation', expected: ["Validé","Refusé","En attente"] },
      { name: 'Action',     expected: ["PEOPLE+","PEOPLE=","PLANCHES+","PLANCHES=","CANCEL"] }
    ];

    checks.forEach(({ name, expected }) => {
      const idx = headersRep.indexOf(name);
      if (idx === -1) return; // déjà signalé ci-dessus

      const nRows = Math.max(0, shRep.getLastRow() - 1);
      if (nRows === 0) return; // pas de données → on ne peut pas tester la validation cellulaire

      const rng = shRep.getRange(2, idx + 1, nRows, 1);
      const dvs = rng.getDataValidations();
      // On lit la première cellule de la colonne (si validation posée en colonne, elle suffira)
      const firstRule = dvs && dvs[0] && dvs[0][0];
      if (!firstRule) {
        out.push(['VALIDATION',`Réponses!${name}`,'MISSING',"Aucune validation (liste) détectée"]);
      } else {
        const crit = firstRule.getCriteriaType();
        const vals = (firstRule.getCriteriaValues() || [])[0] || [];
        const ok = (
          crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST &&
          expected.length === vals.length &&
          expected.every(v => vals.indexOf(v) !== -1)
        );
        if (!ok) {
          out.push(['VALIDATION',`Réponses!${name}`,'INCORRECT',`Attendu: ${expected.join(', ')}`]);
        }
      }
    });
    // après les checks de validation dans auditComplet()
    if (shRep && headersRep.indexOf("Action") !== -1) {
      out.push(['INFO','Réponses!Action','MANUEL', "Dans Sheets: Format > Liste déroulante → Style = 'Chips' + activer 'Plusieurs sélections'."]);
    }
  }

  // --- 5) Écrire rapport
  _writeAuditSheet_(out);
  const msg = out.length ? `Audit terminé — ${out.length} points à vérifier (voir 'Audit')` : "Audit terminé — aucun problème détecté ✅";
  LOG.info(msg);
  LOG.toast(msg, 'Munch • Admin', 6);
  if (out.length) LOG.alert('Avertissement', msg);
}

/** Écrit/écrase la feuille "Audit" avec les résultats */
function _writeAuditSheet_(rows){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name='Audit';
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(name);
  sh.getRange(1,1,1,4).setValues([['Section','Emplacement','Code','Détail / Suggestion']]).setFontWeight('bold');
  if (rows && rows.length){
    sh.getRange(2,1,rows.length,4).setValues(rows);
    sh.autoResizeColumns(1,4);
  }
}
