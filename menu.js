/***********************
 * Menu "Maintenance"  *
 ***********************/
function confirmSendMail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirmation d\'envoi',
    'Tu veux vraiment envoyer les mails de confirmation ?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    sendMail();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Nuit Démoniaques • Admin')
  .addItem('📤 Consolider les inscriptions', 'consolidation')
  .addSubMenu(ui.createMenu('📨 Emails')
    .addItem('📤 Envoyer les mails de confirmation', 'confirmSendMail')
  .addSeparator()
  .addItem('📐 Initialiser les feuilles', 'initSheets')
  // .addItem('🔍 Audit complet', 'auditComplet')
  // .addItem('🛠️ Réparer les formules', 'fixFormulasParticipantsV2')
  .addSubMenu(
    ui.createMenu('🧹 Purges')
      .addItem('Purger Participants', 'purgeParticipants')
      .addItem('Purger Participants + Paiements', 'purgeParticipantsEtPaiements')
      .addItem('Purge complète (Participants + Paiements + Réponses)', 'purgeComplet')
  )
  .addSeparator()
  // .addItem('📘 (Re)générer la feuille Aide', 'regenerateAideSheet')  // ⬅️ ici
  // .addItem('❓ Aide & mode d’emploi', 'showMunchHelp')
  .addToUi();
}


function showMunchHelp() {
  const html = HtmlService.createHtmlOutputFromFile('help_munch')
    .setTitle('Aide • Munch Admin')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Petit panneau d’infos */
function showMaintenanceHelp_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Maintenance',
    [
      '✅ Init Sheets : crée les feuilles manquantes, vérifie/complète les en-têtes, formate les prix, et affiche un récap.',
      '🧹 Purge Participants : efface les données (ligne 2→n) de Participants_validés, en-têtes conservés, puis relance Init.',
      '🧹 Purge Participants + Paiements : idem sur 2 feuilles, puis Init.',
      '🧨 Purge complète : idem + Réponses (feuille du Form), puis Init.',
      '',
      '💡 Rappel : si des en-têtes ont été modifiés/créés, pense à refaire manuellement "Convertir en tableau".'
    ].join('\n'),
    ui.ButtonSet.OK
  );
}