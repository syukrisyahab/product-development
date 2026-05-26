/**
 * userAccess.js — Autentikasi hybrid (token mode + session mode)
 *
 * Sheet Users:
 *   A=Email, B=Nama, C=Role, D=Token(UUID),
 *   E=TokenCreatedAt, F=TokenExpiredAt, G=LastLogin, H=Status(Active/Inactive)
 */


// ─────────────────────────────────────────────────
// CEK USER DI SHEET (by email)
// ─────────────────────────────────────────────────

function getUserAccess_(email) {
  const result = {
    allowed: false, name: '', email: (email || '').toString().toLowerCase(), role: ''
  };
  if (!result.email) return result;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return result;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowEmail = (values[i][0] || '').toString().toLowerCase().trim();
    if (!rowEmail || rowEmail !== result.email) continue;

    // Cek Status Active/Inactive (col H)
    const status = (values[i][7] || 'Active').toString().trim().toLowerCase();
    if (status === 'inactive') return result; // allowed tetap false

    result.allowed = true;
    result.name    = (values[i][1] || '').toString().trim();
    result.role    = (values[i][2] || '').toString().trim();
    return result;
  }
  return result;
}


// ─────────────────────────────────────────────────
// CEK USER DI SHEET (by token UUID)
// ─────────────────────────────────────────────────

function getUserAccessByToken_(token) {
  token = (token || '').toString().trim();
  if (!token) return { allowed: false, reason: 'no_token' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return { allowed: false, reason: 'no_sheet' };

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < values.length; i++) {
    const tok = (values[i][3] || '').toString().trim(); // col D
    if (!tok || tok !== token) continue;

    // Cek Status Active/Inactive (col H)
    const status = (values[i][7] || 'Active').toString().trim().toLowerCase();
    if (status === 'inactive') return { allowed: false, reason: 'inactive' };

    // Cek token expired (col F)
    const expiredAt = values[i][5];
    if (expiredAt instanceof Date && expiredAt.getTime() < Date.now()) {
      return { allowed: false, reason: 'token_expired' };
    }

    const email = (values[i][0] || '').toString().toLowerCase().trim();
    const name  = (values[i][1] || '').toString().trim();
    const role  = (values[i][2] || '').toString().trim();

    // Update LastLogin (col G) — non-critical
    try { sh.getRange(i + 2, 7).setValue(new Date()); } catch (e) {}

    return { allowed: true, email, name, role };
  }
  return { allowed: false, reason: 'not_found' };
}


// ─────────────────────────────────────────────────
// CEK DOMAIN INTERNAL
// ─────────────────────────────────────────────────

function isInternalEmail_(email) {
  email = (email || '').toString().toLowerCase();
  const domains = getConfig_().INTERNAL_DOMAINS || INTERNAL_DOMAINS_DEFAULT;
  return domains.some(d => email.endsWith(d));
}


// ─────────────────────────────────────────────────
// ACTIVE USER — HYBRID (token | session)
// ─────────────────────────────────────────────────

function getActiveUser_(eOrToken) {
  let token        = '';
  let emailSession = '';

  if (typeof eOrToken === 'string') {
    token = eOrToken.toString().trim();
  } else {
    // Dipanggil dari doGet — konteks web app, Session.getActiveUser() reliable
    token        = (eOrToken?.parameter?.token || '').toString().trim();
    emailSession = getSessionEmail_();
  }

  // ── TOKEN MODE ──
  if (token) {
    const access = getUserAccessByToken_(token);
    return {
      mode:    'token',
      token,
      email:   (access.email  || '').toString().toLowerCase(),
      name:    access.name    || '',
      role:    access.role    || '',
      allowed: !!access.allowed,
      reason:  access.reason  || ''
    };
  }

  // ── SESSION MODE (internal domain) ──
  // Hanya dipakai di doGet. google.script.run selalu lewat token mode
  // setelah ensureUserToken_() di doGet mengisi SERVER.token.
  const email = (emailSession || getSessionEmail_()).toString().toLowerCase();
  if (!email || !isInternalEmail_(email)) {
    return { mode: 'none', token: '', email: '', name: '', role: '', allowed: false };
  }

  const access = getUserAccess_(email);
  return {
    mode:    'session',
    token:   '',
    email,
    name:    access.name || email,
    role:    access.role || 'internal',
    allowed: true
  };
}


// Session.getActiveUser() butuh scope userinfo.email.
// Wrap dengan try/catch agar tidak throw jika scope belum granted
// atau dipanggil di konteks yang salah.
function getSessionEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').toString().toLowerCase().trim();
  } catch (e) {
    return '';
  }
}


// ─────────────────────────────────────────────────
// ENSURE TOKEN — buat/reuse token UUID untuk internal user
// Dipanggil dari doGet agar google.script.run selalu pakai token mode
// ─────────────────────────────────────────────────

function ensureUserToken_(email, name, role) {
  email = (email || '').toString().toLowerCase().trim();
  if (!email) return '';

  const cfg        = getConfig_();
  const expiryDays = Number(cfg.TOKEN_EXPIRY_DAYS) || TOKEN_EXPIRY_DAYS_DEFAULT;
  const now        = new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) return '';

  const last   = sh.getLastRow();
  const values = last >= 2 ? sh.getRange(2, 1, last - 1, 8).getValues() : [];

  for (let i = 0; i < values.length; i++) {
    const rowEmail = (values[i][0] || '').toString().toLowerCase().trim();
    if (rowEmail !== email) continue;

    // Cek token yang sudah ada di col D — reuse jika masih valid
    const existingToken = (values[i][3] || '').toString().trim();
    const expiredAt     = values[i][5]; // col F
    const isValid       = existingToken &&
                          expiredAt instanceof Date &&
                          expiredAt.getTime() > now.getTime();
    if (isValid) return existingToken;

    // Token tidak ada atau sudah expired — terbitkan yang baru
    const newToken  = Utilities.getUuid().replace(/-/g, '');
    const newExpiry = new Date(now.getTime() + expiryDays * 86400000);
    const rowNum    = i + 2;
    sh.getRange(rowNum, 4).setValue(newToken);
    sh.getRange(rowNum, 5).setValue(now);
    sh.getRange(rowNum, 6).setValue(newExpiry);
    sh.getRange(rowNum, 7).setValue(now);
    return newToken;
  }

  // User belum ada di sheet (internal baru) — buat baris baru
  const newToken  = Utilities.getUuid().replace(/-/g, '');
  const newExpiry = new Date(now.getTime() + expiryDays * 86400000);
  sh.appendRow([email, name || '', role || 'internal', newToken, now, newExpiry, now, 'Active']);
  return newToken;
}


// ─────────────────────────────────────────────────
// ISSUE TOKEN (dipanggil dari doGet action=issueToken)
// ─────────────────────────────────────────────────

function issueToken_(e) {
  try {
    const idToken  = (e?.parameter?.idToken  || '').toString().trim();
    const redirect = (e?.parameter?.redirect || '').toString().trim();
    if (!idToken) return HtmlService.createHtmlOutput('Missing idToken');

    const decoded = decodeJwtPayload_(idToken);
    const email   = (decoded.email || '').toString().toLowerCase().trim();
    const name    = (decoded.name  || decoded.displayName || '').toString().trim();
    if (!email) return HtmlService.createHtmlOutput('Email not found in token');

    const cfg         = getConfig_();
    const expiryDays  = Number(cfg.TOKEN_EXPIRY_DAYS) || 30;
    const now         = new Date();
    const expiredAt   = new Date(now.getTime() + expiryDays * 86400000);
    const accessToken = Utilities.getUuid().replace(/-/g, '');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_USERS);
    if (!sh) return HtmlService.createHtmlOutput('Sheet Users not found');

    const last   = sh.getLastRow();
    const values = last >= 2 ? sh.getRange(2, 1, last - 1, 8).getValues() : [];
    let   foundRow = -1;

    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || '').toString().toLowerCase().trim() === email) {
        foundRow = i + 2; break;
      }
    }

    const internal = isInternalEmail_(email);

    if (foundRow === -1 && !internal) {
      return HtmlService.createHtmlOutput('Akun Anda belum terdaftar. Hubungi administrator.');
    }

    if (foundRow === -1 && internal) {
      // Internal user baru — buat baris
      sh.appendRow([email, name || '', 'internal', accessToken, now, expiredAt, now, 'Active']);
    } else {
      // Cek status sebelum issue token baru
      const currentStatus = (values[foundRow - 2][7] || 'Active').toString().trim().toLowerCase();
      if (currentStatus === 'inactive') {
        return HtmlService.createHtmlOutput('Akun Anda tidak aktif. Hubungi administrator.');
      }
      sh.getRange(foundRow, 4).setValue(accessToken); // D: token
      sh.getRange(foundRow, 5).setValue(now);          // E: created at
      sh.getRange(foundRow, 6).setValue(expiredAt);    // F: expired at
      sh.getRange(foundRow, 7).setValue(now);          // G: last login
      if (name && !sh.getRange(foundRow, 2).getValue()) {
        sh.getRange(foundRow, 2).setValue(name);
      }
    }

    logActivity_(email, 'issueToken', 'Expires: ' + formatDateOnly_(expiredAt));

    // Selalu gunakan URL generik (/macros/s/...) agar user non-workspace bisa akses.
    // toGenericUrl_ menghapus prefix /a/macros/domain.com/ dari URL manapun.
    const dash      = toGenericUrl_(redirect) || getGenericAppUrl_();
    const targetUrl = dash + (dash.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(accessToken);

    // target="_top" memastikan link dibuka di top-level frame.
    // Jangan gunakan location.replace() sebagai fallback — script.google.com
    // menolak dimuat di dalam iframe (X-Frame-Options: SAMEORIGIN).
    const safeUrl = JSON.stringify(targetUrl);
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html>' +
      '<head><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
      '<body style="font-family:Arial;padding:24px;text-align:center;color:#374151">' +
      '<p>Autentikasi berhasil. Mengalihkan...</p>' +
      '<a href="' + targetUrl + '" target="_top"' +
      ' style="display:inline-block;margin-top:12px;padding:10px 24px;background:#1F6B3E;' +
      'color:#fff;border-radius:8px;text-decoration:none;font-weight:700">Lanjutkan ke Dashboard</a>' +
      '<script>try{top.location.replace(' + safeUrl + ')}catch(e){}</script>' +
      '</body></html>'
    );

  } catch (err) {
    Logger.log('issueToken_ ERROR: ' + err);
    return HtmlService.createHtmlOutput('issueToken error: ' + err);
  }
}


// ─────────────────────────────────────────────────
// DECODE JWT PAYLOAD (tanpa verify signature)
// ─────────────────────────────────────────────────

function decodeJwtPayload_(jwt) {
  jwt = (jwt || '').toString().trim();
  const parts = jwt.split('.');
  if (parts.length < 2) throw new Error('Invalid JWT format');

  const b64  = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  const pad  = '='.repeat((4 - (b64.length % 4)) % 4);
  return JSON.parse(
    Utilities.newBlob(Utilities.base64Decode(b64 + pad)).getDataAsString('UTF-8')
  );
}
