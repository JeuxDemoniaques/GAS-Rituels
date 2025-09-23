function getCredentials() {
  const props = PropertiesService.getScriptProperties();
  return {
    clientId: props.getProperty("PAYPAL_CLIENT_ID"),
    clientSecret: props.getProperty("PAYPAL_CLIENT_SECRET"),
    wrapperUrl: props.getProperty("WRAPPER_URL"),
    webappUrl: props.getProperty("WEBAPP_URL"),
    paiementOpen: props.getProperty("PAIEMENT_OPEN")
  };
}

/** Utils */
const eqEmail = (a,b) => String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase();

/** Lit l’onglet Tarifs et retourne les prix unitaires. */
function getBoardPrices_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Tarifs');
  if (!sh) throw new Error("Feuille 'Tarifs' introuvable");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { Fromage: 0, Charcuterie: 0, Mixte: 0 };
  }

  // A: Produit, B: Description (optionnel), C: Prix
  const rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  const map = {};
  rows.forEach(([prod, , prix]) => {
    if (!prod) return;
    const key = String(prod).trim().toLowerCase();
    const n = (typeof prix === 'number') ? prix : parseFloat(String(prix).replace(',', '.')) || 0;
    map[key] = n;
  });

  return {
    Fromage:     map['fromage']     ?? 0,
    Charcuterie: map['charcuterie'] ?? 0,
    Mixte:       map['mixte']       ?? 0
  };
}

/** Convertit proprement en nombre (gère "12,50", "", null). */
function normalizeAmount(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).trim().replace(/\s/g,'').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

/** Format FR simple "12,50 €". */
function formatPrice_(n) {
  const v = Math.round((Number(n) + Number.EPSILON) * 100) / 100;
  return v.toFixed(2).replace('.', ',') + ' €';
}

/** Parser la liste d'actions depuis la colonne Action (multi-sélection possible)*/
function parseActions_(raw) {
  if (!raw) return [];
  return String(raw)
    .split(/[;,]+/)         // sépare sur virgules ou ;
    .map(s => s.trim().toUpperCase())
    .filter(Boolean);
}

function getInt_(v) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : 0;
}

/** Séparateur de paramètres de formule (fr/de/it → ';', sinon ',') */
function _sep_() {
  const ss  = SpreadsheetApp.getActive();
  const loc = (ss && ss.getSpreadsheetLocale ? ss.getSpreadsheetLocale() : '') 
           || (Session.getActiveUserLocale() || '');
  return (/^(fr|de|it|es|pt|pl|ru)/i.test(loc) ? ';' : ',');
}

/** ---- Output JSON ---- */
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}