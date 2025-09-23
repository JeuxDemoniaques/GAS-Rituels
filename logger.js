/** logger.gs — utilitaire de logs unifié */
var LOG = (() => {
  const props = PropertiesService.getScriptProperties();
  const DEBUG =
    String(props.getProperty('DEBUG') || 'false').toLowerCase() === 'true';

  function fmt(parts) {
    return parts
      .map(p => {
        if (p == null) return String(p);
        if (typeof p === 'string') return p;
        try { return JSON.stringify(p); } catch(_) { return String(p); }
      })
      .join(' ');
  }

  function _out(level, ...parts) {
    const msg = `[${level}] ${fmt(parts)}`;
    // Console (visible dans Exécutions / journaux) + fallback Logger
    try {
      if (level === 'ERROR' && console && console.error) console.error(msg);
      else if (level === 'WARN'  && console && console.warn) console.warn(msg);
      else if (level === 'INFO'  && console && console.log)  console.log(msg);
      else if (console && console.log) console.log(msg);
      else Logger.log(msg);
    } catch (_) {
      Logger.log(msg);
    }
    return msg;
  }

  function info(...a)  { return _out('INFO',  ...a); }
  function warn(...a)  { return _out('WARN',  ...a); }
  function error(...a) { return _out('ERROR', ...a); }
  function debug(...a) { if (DEBUG) return _out('DEBUG', ...a); }

  function toast(message, title, seconds) {
    try {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast(String(message), String(title || 'Info'), Number(seconds || 5));
    } catch (_) {/* hors contexte Sheets */}
  }

  function alert(title, message) {
    try {
      SpreadsheetApp.getUi()
        .alert(String(title || 'Info'), String(message), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (_) {/* hors contexte Sheets */}
  }

  /** Log + toast pratique (liste ou texte) */
  function infoToast(title, lines, seconds) {
    const msg = Array.isArray(lines) ? lines.join('\n') : String(lines);
    info(title, '\n' + msg);
    toast(msg, title, seconds);
  }

  return { info, warn, error, debug, toast, alert, infoToast };
})();
