// -- Montant : re-pose aussi si la cellule affiche une erreur (#...)
function ensureMontantFormula_(sh, row, colMontant) {
  if (colMontant < 1) return;
  const SEP  = _sep_();
  const cell = sh.getRange(row, colMontant);
  const disp = String(cell.getDisplayValue() || '').toUpperCase();
  if (!cell.getFormula() || disp.startsWith('#')) {
    const f =
      `=INDEX($1:$9999${SEP} ROW()${SEP} MATCH("Participants"${SEP} $1:$1${SEP} 0)) * VLOOKUP("Entrée"${SEP} Tarifs!$A:$C${SEP} 3${SEP} FALSE)`
    cell.setFormula(f).setNumberFormat('0.00');
  }
}

// -- Total payé : idem, re-pose si erreur
function ensureTotalPayeFormula_(sh, row, colTotalPaye) {
  if (colTotalPaye < 1) return;
  const SEP  = _sep_();
  const cell = sh.getRange(row, colTotalPaye);
  const disp = String(cell.getDisplayValue() || '').toUpperCase();
  if (!cell.getFormula() || disp.startsWith('#')) {
    const f = `=IFNA(SUMIF(Paiements!B:B${SEP} A${row}${SEP} Paiements!C:C)${SEP} 0)`;
    cell.setFormula(f).setNumberFormat('0.00');
  }
}

// Pose Statut Paiement (R1C1) si absent ou erroné
function ensureStatutPaiementFormula_(sh, row, colStatut) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxMontant   = headers.indexOf("Montant");
  const idxTotalPaye = headers.indexOf("Total payé");
  if (idxMontant === -1 || idxTotalPaye === -1) return;

  const offTot  = (idxTotalPaye + 1) - colStatut;
  const offMont = (idxMontant   + 1) - colStatut;

  const SEP = _sep_();
  const f = `=IF(R[0]C[${offTot}]=0${SEP}"NONE"${SEP}` +
            `IF(R[0]C[${offTot}]<R[0]C[${offMont}]${SEP}"PARTIAL"${SEP}` +
            `IF(R[0]C[${offTot}]=R[0]C[${offMont}]${SEP}"COMPLETED"${SEP}"OVERPAID")))`;

  const cell = sh.getRange(row, colStatut);
  if (!cell.getFormula() || String(cell.getDisplayValue()).toUpperCase().startsWith("#ERR")) {
    cell.setFormulaR1C1(f);
  }
}

/**
 * Répare (ou pose) les formules Montant / Total payé / Statut Paiement
 * sur toutes les lignes de Participants_validés (à partir de la ligne 2).
 * - Ne touche pas aux autres cellules.
 * - Met le format 0.00 sur Montant & Total payé si absent.
 * - Affiche un toast + écrit dans les logs.
 */
function fixFormulasParticipantsV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Participants_validés');
  if (!sh) {
    const msg = "Feuille 'Participants_validés' introuvable.";
    LOG.info(msg);
    LOG.toast(msg, 'Réparer les formules', 5);
    return;
  }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    const msg = "Aucune donnée à réparer (seulement l’en-tête).";
    LOG.info(msg);
    LOG.toast(msg, 'Réparer les formules', 5);
    return;
  }

  // Récup indexes de colonnes une seule fois
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxMontant   = headers.indexOf('Montant');
  const idxTotalPaye = headers.indexOf('Total payé');
  const idxStatut    = headers.indexOf('Statut Paiement');

  const missing = [];
  if (idxMontant === -1)   missing.push('Montant');
  if (idxTotalPaye === -1) missing.push('Total payé');
  if (idxStatut === -1)    missing.push('Statut Paiement');

  if (missing.length) {
    const msg = "Colonnes manquantes dans 'Participants_validés' : " + missing.join(', ');
    LOG.info(msg);
    LOG.alert('Réparer les formules', msg);
    return;
  }

  let changed = 0;
  for (let r = 2; r <= lastRow; r++) {
    // Pose/repère les formules (helpers que tu utilises déjà)
    ensureMontantFormula_(sh, r, idxMontant + 1);        // ou ensureMontantFormula_(r) selon ta variante
    ensureTotalPayeFormula_(sh, r, idxTotalPaye + 1);    // ou ensureTotalPayeFormula_(r)
    ensureStatutPaiementFormula_(sh, r, idxStatut + 1);                     // version “sans params” fournie plus tôt

    // (Optionnel) forcer le format joli
    sh.getRange(r, idxMontant + 1).setNumberFormat('0.00');
    sh.getRange(r, idxTotalPaye + 1).setNumberFormat('0.00');

    changed++;
  }

  SpreadsheetApp.flush();

  const msg = `Formules réparées sur ${changed} ligne(s).`;
  LOG.info(msg);
  LOG.toast(msg, 'Réparer les formules', 5);
}