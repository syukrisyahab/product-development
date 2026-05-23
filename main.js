/**
 * main.js — Entry point Web App & konstanta global
 */

const SHEET_PD          = 'ProductDev';
const SHEET_STATUS      = 'MasterStatus';
const SHEET_USERS       = 'Users';
const SHEET_PD_STATUS   = 'PD_StatusLog';
const SHEET_CONFIG      = 'Config';
const SHEET_LOG         = 'AppLog';

// Fallback defaults — di-override oleh Config sheet jika ada
const LPP_FOLDER_ID_DEFAULT     = '1io0ZXlbFXwtZg89qCFCSwbUctIMkRyIe';
const TOKEN_EXPIRY_DAYS_DEFAULT  = 30;


function doGet(e) {
  const action = (e?.parameter?.action || '').toString();
  if (action === 'issueToken') return issueToken_(e);

  const token  = (e?.parameter?.token || '').toString().trim();
  const access = token ? getUserAccessByToken_(token) : { allowed: false };

  if (!access.allowed) {
    return redirectToLogin_();
  }

  const t     = HtmlService.createTemplateFromFile('index');
  t.userEmail = (access.email || '').toLowerCase();
  t.userName  = access.name  || access.email || 'Guest';
  t.isAllowed = true;
  t.userRole  = access.role  || '';
  t.token     = token;

  return t.evaluate()
    .setTitle('Product Development Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function redirectToLogin_() {
  const loginUrl = getConfig_().DASHBOARD_URL || '';
  // target="_top" memastikan link dibuka di top-level frame,
  // bukan di dalam iframe Apps Script.
  // Jangan gunakan location.replace() sebagai fallback karena
  // script.google.com menolak dimuat di dalam iframe (X-Frame-Options: SAMEORIGIN).
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html>' +
    '<head><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="font-family:Arial;padding:32px;text-align:center;color:#374151">' +
    '<p style="margin-bottom:20px">Sesi tidak valid atau belum login.</p>' +
    '<a href="' + loginUrl + '" target="_top"' +
    ' style="display:inline-block;padding:12px 28px;background:#1F6B3E;color:#fff;' +
    'border-radius:10px;font-weight:700;text-decoration:none;font-size:15px">Login</a>' +
    '<script>try{top.location.replace(' + JSON.stringify(loginUrl) + ')}catch(e){}</script>' +
    '</body></html>'
  );
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
