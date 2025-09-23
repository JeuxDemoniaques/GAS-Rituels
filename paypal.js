/** paypal.gs — module PayPal autonome (pas de globals partagés) */
var PAYPAL = (()=>{
  function env_() {
    const e = (PropertiesService.getScriptProperties().getProperty("ENV") || "dev").toLowerCase();
    return e === "prod" ? "live" : "sandbox";
  }
  function baseUrl_() {
    return env_() === "live" ? "https://api-m.paypal.com" : "https://api-m.sandbox.paypal.com";
  }

  // Mise en cache du token pour éviter de frapper l’API à chaque appel
  function getToken_() {
    const cache = CacheService.getScriptCache();
    const cacheKey = "PAYPAL_TOKEN_" + env_();
    const cached = cache.get(cacheKey);
    if (cached) return cached;

    const creds = getCredentials(); // helpers.gs
    const clientId = creds.clientId;
    const clientSecret = creds.clientSecret;
    if (!clientId || !clientSecret) {
      throw new Error("PAYPAL_CLIENT_ID / PAYPAL_CLIENT_SECRET manquants dans Script Properties");
    }

    const resp = UrlFetchApp.fetch(baseUrl_() + "/v1/oauth2/token", {
      method: "post",
      headers: { "Authorization": "Basic " + Utilities.base64Encode(clientId + ":" + clientSecret) },
      payload: { "grant_type": "client_credentials" },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    const txt  = resp.getContentText();
    if (code < 200 || code >= 300) {
      LOG.error('Token HTTP error', code, txt);
      throw new Error("Erreur PayPal (token): HTTP " + code);
    }
    const json = JSON.parse(txt);
    const token = json && json.access_token;
    if (!token) {
      LOG.error('Token payload invalide', txt);
      throw new Error("Erreur PayPal (token): payload invalide");
    }

    // Token TTL typique 8h ; on met 20 minutes pour rester conservateur et éviter les 401
    cache.put(cacheKey, token, 20 * 60);
    return token;
  }

  // Petit helper pratique pour composer les headers d’API PayPal
  function authHeader_() {
    return { "Authorization": "Bearer " + getToken_() };
  }

  // Helper pour header d’idempotence
  function requestId_(seed) {
    return Utilities.base64Encode(String(seed || Utilities.getUuid()));
  }

  return {
    env: env_,
    baseUrl: baseUrl_,
    getToken: getToken_,
    authHeader: authHeader_,
    requestId: requestId_
  };
})();
