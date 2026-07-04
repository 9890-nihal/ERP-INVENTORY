/**
 * Code.gs
 * MAIN ROUTER FILE
 *
 * ============================================================
 * ANDROID APP URL FIX (IMPORTANT!)
 * ============================================================
 * Pehle ka doGet() hamesha Material List JSON return karta tha.
 * Ab doGet() ERP website (HTML) serve karta hai by default.
 *
 * Android app ka Material List download URL update karo:
 *   OLD (kaam nahi karta): https://script.google.com/macros/s/YOUR_ID/exec
 *   NEW (sahi URL):        https://script.google.com/macros/s/YOUR_ID/exec?api=material-list
 *
 * Android app mein sirf URL ke aant mein ?api=material-list add karo.
 * POST (Logistics upload) ke liye koi change nahi — woh pehle ki tarah kaam karta hai.
 * ============================================================
 *
 * ROUTES:
 *   GET  /exec                    -> ERP Web Page (HTML)
 *   GET  /exec?api=material-list  -> Android app Material List JSON
 *   GET  /exec?api=materials      -> Same as above (alias)
 *   GET  /exec?api=items          -> Same as above (alias)
 *   GET  /exec?api=health         -> Deployment test JSON
 *   POST /exec                    -> Android app Logistics data save
 */

function doGet(e) {
  var api = '';

  if (e && e.parameter) {
    api = String(
      e.parameter.api ||
      e.parameter.action ||
      e.parameter.mode ||
      ''
    ).trim().toLowerCase();
  }

  // --- Health check ---
  if (api === 'health') {
    return jsonOutput_({
      status: 'ok',
      app: 'ERP',
      route: 'health',
      message: 'ERP router is active',
      date: new Date().toISOString()
    });
  }

  // --- Material List API (Android app use karta hai) ---
  if (
    api === 'material-list' ||
    api === 'materiallist' ||
    api === 'materials' ||
    api === 'items'
  ) {
    return getMaterialListJsonApi_();
  }

  // --- Default: ERP Web Page (HTML) ---
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ERP System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST handler — Android app Logistics data yahan aata hai.
 * Koi change nahi, pehle ki tarah kaam karta hai.
 */
function doPost(e) {
  return saveLogisticsRowsApi_(e);
}

/**
 * GAS template include helper
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
