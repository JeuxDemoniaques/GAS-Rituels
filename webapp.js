/*************************
 *  WebApp PayPal (one-sheet)
 *************************/

const base  = PAYPAL.baseUrl();
// const token = PAYPAL.getToken();

// Une seule source de vérité
const SHEET_OK = "Participants_validés"; // email, Montant (due), Total payé, Statut Paiement

function doGet(e) {
  try {
    LOG.info('doGet params:', e && e.parameter);
    const action = (e.parameter.action || "create").toLowerCase();
    if (action === "diag") {
      const creds = getCredentials(); // helpers.gs
      const cid = creds && creds.clientId ? String(creds.clientId) : "";
      return jsonOut({
        env: PAYPAL.env().toUpperCase(),           // "SANDBOX" ou "LIVE"
        clientId_prefix: cid ? cid.slice(0, 10) : null,
        wrapperUrl_set: !!creds.wrapperUrl,        // true si WRAPPER_URL est défini
        webappUrl_set: !!creds.webappUrl,          // true si WEBAPP_URL est défini
        paypal_base: PAYPAL.baseUrl()
      });
    }

    if (action === "create")  return handleCreate(e);
    if (action === "capture") return handleCapture(e);
    return jsonOut({ error: "action inconnue" });
  } catch (err) {
    LOG.error('doGet error:', err);
    return jsonOut({ error: String(err) });
  }
}

/** ---- CREATE ----
 * Lit le Montant dû dans Participants_validés pour l'email,
 * crée l'order PayPal et renvoie approveUrl.
 */
function handleCreate(e) {
  const email = e.parameter.email;
  if (!email) return jsonOut({ error: "Email manquant" });
  if (!WRAPPER_URL) {
    return jsonOut({ error: "WRAPPER_URL manquant dans les Script Properties." });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_OK);
  if (!sh) return jsonOut({ error: `Feuille '${SHEET_OK}' introuvable` });

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return jsonOut({ error: "Aucun participant" });

  const head       = data[0];
  const idxMail    = head.indexOf('Email');
  const idxMontant = head.indexOf('Montant');
  const idxTotalPaye = head.indexOf('Total payé')
  if (idxMail === -1 || idxMontant === -1 || idxTotalPaye === -1) {
    LOG.error('Colonnes manquantes', { idxMail, idxMontant, idxTotalPaye, head });
    return jsonOut({ error: "Colonnes introuvables (Email / Montant / Total payé)" });
  }

  SpreadsheetApp.flush(); // s'il y a une formule dans Montant

  let montant = NaN;
  let totalPaye = NaN;
  for (let i = 1; i < data.length; i++) {
    if (eqEmail(data[i][idxMail], email)) {
      montant = normalizeAmount(data[i][idxMontant]);
      totalPaye = normalizeAmount(data[i][idxTotalPaye]);
      break;
    }
  }
  if (!isFinite(montant))   return jsonOut({ error: "Montant invalide" });
  if (!isFinite(totalPaye))   return jsonOut({ error: "Total invalide" });
  if (montant <= 0)         return jsonOut({ error: "Montant nul : aucun paiement à créer" });
  let reste = montant - totalPaye
  if (reste <= 0)         return jsonOut({ error: "Montant nul : aucun paiement à créer" });

  const amountStr = reste.toFixed(2);
  LOG.info('Montant trouvé:', amountStr, 'pour', email);

  // Créer l'order PayPal
  const orderPayload = {
    intent: "CAPTURE",
    purchase_units: [{
      amount: { currency_code: "EUR", value: amountStr },
      custom_id: email
    }],
    application_context: {
      brand_name: "Les Jeux Démoniaques",
      user_action: "PAY_NOW",
      return_url: WRAPPER_URL + "?paypal=return&email=" + encodeURIComponent(email),
      cancel_url: WRAPPER_URL + "?paypal=cancel&email=" + encodeURIComponent(email)
    }
  };

  const orderRes = UrlFetchApp.fetch(base + "/v2/checkout/orders", {
    method: "post",
    contentType: "application/json",
    headers: {
      ...PAYPAL.authHeader(),
      "PayPal-Request-Id": PAYPAL.requestId(),
    },
    payload: JSON.stringify(orderPayload),
    muteHttpExceptions: true
  });

  const status = orderRes.getResponseCode();
  let order; try { order = JSON.parse(orderRes.getContentText()); } catch(_) { order = {}; }
  if (status < 200 || status >= 300) {
    LOG.error('Order create HTTP error', status, orderRes.getContentText());
    return jsonOut({ error: "Erreur PayPal (create): HTTP " + status, raw: order });
  }

  // Lien d'approbation
  let approveUrl = "";
  if (order && order.links) {
    for (var j = 0; j < order.links.length; j++) {
      if (order.links[j].rel === "approve") { approveUrl = order.links[j].href; break; }
    }
  }
  if (!approveUrl) {
    LOG.error('approveUrl introuvable', order);
    return jsonOut({ error: "approveUrl introuvable", raw: order });
  }

  return jsonOut({ approveUrl });
}

/** Anti-doublons (cherche un TransactionID déjà enregistré) */
function _existsPaymentByTxn_(shPay, transId) {
  if (!transId) return false;
  const last = shPay.getLastRow();
  if (last < 2) return false;
  const vals = shPay.getRange(2, 5, last - 1, 1).getValues().flat(); // col E = TransactionID
  return vals.some(v => String(v || '').trim() === String(transId).trim());
}

/** Anti-doublons (cherche un OrderID déjà enregistré) */
function _existsPaymentByOrder_(shPay, orderId) {
  if (!orderId) return false;
  const last = shPay.getLastRow();
  if (last < 2) return false;
  const vals = shPay.getRange(2, 4, last - 1, 1).getValues().flat(); // col D = OrderID
  return vals.some(v => String(v || '').trim() === String(orderId).trim());
}


/** ---- CAPTURE ----
 * Capture PayPal, journalise dans Paiements (anti-doublons),
 */
function handleCapture(e) {
  const orderId = e.parameter.orderId || e.parameter.token;
  const email   = e.parameter.email;
  if (!orderId || !email) return jsonOut({ error: "orderId/email manquant" });

  const url   = base + "/v2/checkout/orders/" + orderId + "/capture";

  // Idempotence côté PayPal
  const capRes = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      ...PAYPAL.authHeader(),
      "PayPal-Request-Id": PAYPAL.requestId(orderId + '|' + email)
    },
    muteHttpExceptions: true
  });

  const status = capRes.getResponseCode();
  let cap; try { cap = JSON.parse(capRes.getContentText()); } catch (_) { cap = {}; }
  if (status < 200 || status >= 300) {
    LOG.error('Capture HTTP error', status, capRes.getContentText());
    return jsonOut({ error: "Erreur PayPal (capture): HTTP " + status, raw: cap });
  }

  if (cap.status !== "COMPLETED") {
    return jsonOut({ status: cap.status || "UNKNOWN" });
  }

  // Récup données fiables depuis la réponse PayPal
  const pu      = cap.purchase_units && cap.purchase_units[0];
  const capture = pu && pu.payments && pu.payments.captures && pu.payments.captures[0];

  // Email « de vérité » depuis l'order:
  const emailParam = String(email || '').trim().toLowerCase();
  const emailFromOrder = (pu && pu.custom_id) ? String(pu.custom_id).trim().toLowerCase() : null;
  const finalEmail = emailFromOrder || emailParam;

  const amountStr = capture?.amount?.value || pu?.amount?.value || null;
  // const amountStr =
  //   (capture && capture.amount && capture.amount.value) ||
  //   (pu && pu.amount && pu.amount.value) || null;

  // Conversion string → number robuste
  const amount = normalizeAmount(amountStr);
  if (!isFinite(amount)) {
    LOG.error("Montant de capture invalide:", amountStr);
    return jsonOut({ error: "Montant de capture invalide" });
  }

  // Sécurise devise
  const cc = capture?.amount?.currency_code || pu?.amount?.currency_code || null;
  // const cc = capture && capture.amount && capture.amount.currency_code;
  if (cc && cc !== 'EUR') return jsonOut({ error: "Devise inattendue sur la capture", currency: cc });

  const transId = capture && capture.id;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Journal "Paiements" (anti-doublons)
  let shPay = ss.getSheetByName("Paiements");
  if (!shPay) {
    shPay = ss.insertSheet("Paiements");
    shPay.appendRow(["Date","Email","Montant","OrderID","TransactionID","Statut"]);
  }

  const dupByTxn  = _existsPaymentByTxn_(shPay, transId);
  const dupByOrd  = _existsPaymentByOrder_(shPay, orderId);

  if (!dupByTxn && !dupByOrd) {
    shPay.appendRow([new Date(), finalEmail, amount, orderId, transId, "COMPLETED"]);
  } else {
    LOG.info('Paiement déjà enregistré, saut de l’append.', { orderId, transId });
  }

  SpreadsheetApp.flush();
  return jsonOut({ status: "COMPLETED" });
}
