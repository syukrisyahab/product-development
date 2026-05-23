/**
 * Jobs.js — Email notifications (circle + PIC berikutnya)
 * Email circle dibaca dari Config sheet, bukan hardcoded.
 */


// ─────────────────────────────────────────────────
// HELPERS (baca dari sheet)
// ─────────────────────────────────────────────────

function getStatusNameByOrderIndex_(orderIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STATUS);
  if (!sh || sh.getLastRow() < 2) return '';
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (Number(values[i][1]) === Number(orderIndex)) return String(values[i][0] || '').trim();
  }
  return '';
}

function getNamaProdukById_(pdId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_PD);
  if (!sh || sh.getLastRow() < 2) return '';
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(pdId).trim()) return String(values[i][4] || '').trim();
  }
  return '';
}

function getPicEmailByOrderIndex_(orderIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STATUS);
  if (!sh || sh.getLastRow() < 2) return '';
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < values.length; i++) {
    if (Number(values[i][1]) === Number(orderIndex)) return String(values[i][2] || '').trim();
  }
  return '';
}

function buildEmailTable_(rows) {
  const trs = rows.map(([label, value]) =>
    `<tr>
       <td style="border:1px solid #ddd;padding:7px 10px;font-weight:700;white-space:nowrap">${label}</td>
       <td style="border:1px solid #ddd;padding:7px 10px">${value || '-'}</td>
     </tr>`
  ).join('');
  return `<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd;width:100%">${trs}</table>`;
}

function buildDashboardButton_(url) {
  return `<p style="margin-top:14px">
    <a href="${url}" target="_blank"
       style="display:inline-block;padding:10px 18px;background:#00B050;color:#fff;
              text-decoration:none;border-radius:8px;font-weight:700;font-size:13px">
      Buka Dashboard
    </a>
  </p>`;
}

function emailWrapper_(title, body) {
  return `<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5;color:#1f2937;max-width:600px">
    <h3 style="margin:0 0 14px;color:#007A33">${title}</h3>
    ${body}
  </div>`;
}


// ─────────────────────────────────────────────────
// EMAIL: STATUS UPDATE (ke circle)
// ─────────────────────────────────────────────────

function sendPDStatusEmail_(payload) {
  payload = payload || {};
  const id        = (payload.id        || '').toString().trim();
  const orderIdx  = Number(payload.orderIndex) || 0;
  const result    = (payload.result    || '').toString().trim();
  const note      = (payload.note      || '').toString().trim();
  const updatedBy = (payload.updatedBy || '').toString().trim();

  if (!id || !orderIdx || !result) return;

  const cfg        = getConfig_();
  const statusName = getStatusNameByOrderIndex_(orderIdx) || ('Step ' + orderIdx);
  const namaProduk = getNamaProdukById_(id) || '-';
  const dashUrl    = cfg.DASHBOARD_URL || 'https://pd-dashboard-rbg.web.app/';
  const whenStr    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  const to      = (cfg.CIRCLE_EMAILS || PD_CIRCLE_EMAILS_DEFAULT).join(',');
  const subject = `[ProductDev] ${id} — ${statusName} (${result})`;
  const html    = emailWrapper_('Product Development Status Update',
    buildEmailTable_([
      ['ID',          id],
      ['Nama Produk', namaProduk],
      ['Status',      statusName],
      ['Decision',    result],
      ['Keterangan',  note || '-'],
      ['Updated By',  updatedBy],
      ['Waktu',       whenStr]
    ]) + buildDashboardButton_(dashUrl)
  );

  MailApp.sendEmail({ to, subject, htmlBody: html });
}


// ─────────────────────────────────────────────────
// EMAIL: NOTIFIKASI KE PIC BERIKUTNYA
// ─────────────────────────────────────────────────

function sendPDNextPicEmail_(payload) {
  payload = payload || {};
  const id        = (payload.id        || '').toString().trim();
  const orderIdx  = Number(payload.orderIndex) || 0;
  const result    = (payload.result    || '').toString().trim();
  const updatedBy = (payload.updatedBy || '').toString().trim();

  // Hanya kirim jika decision = OK
  if (!id || !orderIdx || result.toLowerCase() !== 'ok') return;

  const nextIdx    = orderIdx + 1;
  const picEmail   = getPicEmailByOrderIndex_(nextIdx);
  if (!picEmail) return;

  const cfg          = getConfig_();
  const nextStepName = getStatusNameByOrderIndex_(nextIdx) || ('Step ' + nextIdx);
  const namaProduk   = getNamaProdukById_(id) || '-';
  const dashUrl      = cfg.DASHBOARD_URL || 'https://pd-dashboard-rbg.web.app/';
  const whenStr      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  const subject = `[ACTION NEEDED] ${id} — ${namaProduk} menunggu review Anda: ${nextStepName}`;
  const html    = emailWrapper_('Tindakan Dibutuhkan',
    `<p>Anda ditunjuk sebagai PIC untuk tahap <b>${nextStepName}</b> pada request berikut:</p>` +
    buildEmailTable_([
      ['ID',          id],
      ['Nama Produk', namaProduk],
      ['Tahap',       nextStepName],
      ['Approved By', updatedBy],
      ['Waktu',       whenStr]
    ]) + buildDashboardButton_(dashUrl)
  );

  MailApp.sendEmail({ to: picEmail, subject, htmlBody: html });
}


// ─────────────────────────────────────────────────
// EMAIL: NEW REQUEST (ke circle)
// ─────────────────────────────────────────────────

function sendPDCreateEmail_(payload) {
  payload = payload || {};
  const id             = (payload.id             || '').toString().trim();
  const namaProduk     = (payload.namaProduk     || '').toString().trim();
  const kategori       = (payload.kategori       || '').toString().trim();
  const deskripsi      = (payload.deskripsi      || '').toString().trim();
  const estimasiQty    = (payload.estimasiQty    || '').toString().trim();
  const estimasiOrder  = payload.estimasiOrder   || '';
  const lppUrl         = (payload.lppUrl         || '').toString().trim();
  const createdByEmail = (payload.createdByEmail || '').toString().trim();
  const createdByName  = (payload.createdByName  || '').toString().trim();

  if (!id || !namaProduk) return;

  const cfg     = getConfig_();
  const to      = (cfg.CIRCLE_EMAILS || PD_CIRCLE_EMAILS_DEFAULT).join(',');
  const dashUrl = cfg.DASHBOARD_URL || 'https://pd-dashboard-rbg.web.app/';
  const whenStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  let estOrderStr = '-';
  try {
    const d = estimasiOrder instanceof Date ? estimasiOrder : new Date(estimasiOrder);
    if (!isNaN(d.getTime())) estOrderStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {}

  const lppCell = lppUrl
    ? `<a href="${lppUrl}" style="color:#007A33;font-weight:700">Open LPP</a>`
    : '-';

  const subject = `[ProductDev] NEW REQUEST ${id} — ${namaProduk}`;
  const html    = emailWrapper_('New Product Development Request',
    buildEmailTable_([
      ['ID',                   id],
      ['Nama Produk',          namaProduk],
      ['Kategori',             kategori    || '-'],
      ['Deskripsi',            deskripsi   || '-'],
      ['Estimasi Qty',         estimasiQty || '-'],
      ['Estimasi Order Date',  estOrderStr],
      ['LPP Document',         lppCell],
      ['Created By',           (createdByName ? createdByName + ' — ' : '') + createdByEmail],
      ['Waktu',                whenStr]
    ]) + buildDashboardButton_(dashUrl)
  );

  MailApp.sendEmail({ to, subject, htmlBody: html });
}


// ─────────────────────────────────────────────────
// TEST FUNCTION
// ─────────────────────────────────────────────────

function testSendEmail() {
  MailApp.sendEmail('syukrisyahab@gmail.com', '[Test] PD Dashboard', 'Tes izin kirim email dari Apps Script.');
}
