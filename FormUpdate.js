/**
 * FormUpdate.js — Submit, edit, cancel request PD + update status
 */


// ─────────────────────────────────────────────────
// SUBMIT FORM BARU
// ─────────────────────────────────────────────────

function submitPDForm(formObj) {
  try {
    const au = getActiveUser_((formObj && formObj.token) ? formObj.token : '');
    if (!au.allowed || !au.email) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');

    const role = (au.role || '').toLowerCase();
    if (role === 'viewer') return errResponse_('Role Viewer tidak bisa membuat request.', 'FORBIDDEN');

    const namaProduk  = (formObj.namaProduk   || '').toString().trim();
    const kategori    = (formObj.kategori     || '').toString().trim();
    const deskripsi   = (formObj.deskripsi    || '').toString().trim();
    const estimasiQty = formObj.estimasiQty   || '';
    const estOrderRaw = formObj.estimasiOrder || '';
    const lppBlob     = formObj.lppFile;

    if (!namaProduk) return errResponse_('Nama Produk / Project harus diisi.', 'VALIDATION');
    if (!kategori)   return errResponse_('Kategori harus dipilih.', 'VALIDATION');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PD);
    if (!sh) return errResponse_("Sheet 'ProductDev' tidak ditemukan.", 'SHEET_NOT_FOUND');

    const cfg = getConfig_();
    const now = new Date();
    let   idStr = '';
    let   lppUrl = '';

    // ── LockService: cegah race condition pada ID generation ──
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const lastRow = sh.getLastRow();
      let   newSeq  = 1;
      if (lastRow > 1) {
        const lastId = sh.getRange(lastRow, 1).getValue();
        const m      = /PD-\d{8}-(\d+)/.exec(lastId);
        if (m) newSeq = parseInt(m[1], 10) + 1;
      }
      const ymd = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
      idStr     = 'PD-' + ymd + '-' + ('000' + newSeq).slice(-3);

      // Upload LPP ke Drive
      if (lppBlob && lppBlob.getName) {
        const folder   = DriveApp.getFolderById(cfg.LPP_FOLDER_ID || LPP_FOLDER_ID_DEFAULT);
        const fileName = idStr + ' - ' + lppBlob.getName();
        const file     = folder.createFile(lppBlob.copyBlob().setName(fileName));
        lppUrl         = file.getUrl();
      }

      const estOrderDate = estOrderRaw ? new Date(estOrderRaw) : '';
      sh.appendRow([
        idStr, now, au.email, au.name || '',
        namaProduk, kategori, deskripsi,
        estimasiQty, estOrderDate, lppUrl,
        '', '', '', ''
      ]);
    } finally {
      lock.releaseLock();
    }

    sendPDCreateEmail_({
      id: idStr, namaProduk, kategori, deskripsi,
      estimasiQty, estimasiOrder: estOrderRaw ? new Date(estOrderRaw) : '',
      lppUrl, createdByEmail: au.email, createdByName: au.name || ''
    });

    invalidatePdDataCache_(au.email);
    logActivity_(au.email, 'submitPDForm', 'ID: ' + idStr + ' | Produk: ' + namaProduk);

    return okResponse_({ id: idStr }, 'Request berhasil dibuat dengan ID ' + idStr);

  } catch (err) {
    Logger.log('submitPDForm ERROR: ' + err);
    return errResponse_(String(err.message || err), 'INTERNAL_ERROR');
  }
}


// ─────────────────────────────────────────────────
// UPDATE STATUS
// ─────────────────────────────────────────────────

function submitPDStatus(formObj) {
  try {
    const au = getActiveUser_((formObj && formObj.token) ? formObj.token : '');
    if (!au.allowed || !au.email) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');

    const role = (au.role || '').toLowerCase();
    if (role === 'viewer' || role === 'requester') {
      return errResponse_('Role Anda tidak bisa melakukan update status.', 'FORBIDDEN');
    }

    const id         = (formObj.pdId       || '').toString().trim();
    const orderIndex = Number(formObj.orderIndex || 0);
    const result     = (formObj.result     || '').toString().trim();
    const keterangan = (formObj.keterangan || '').toString().trim();

    if (!id || !orderIndex) return errResponse_('ID atau OrderIndex tidak valid.', 'VALIDATION');
    if (!result)            return errResponse_('Result (OK / Not OK) harus dipilih.', 'VALIDATION');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PD_STATUS);
    if (!sh) return errResponse_("Sheet 'PD_StatusLog' tidak ditemukan.", 'SHEET_NOT_FOUND');

    sh.appendRow([id, orderIndex, result, new Date(), keterangan, au.email]);

    sendPDStatusEmail_({ id, orderIndex, result, note: keterangan, updatedBy: au.email });
    sendPDNextPicEmail_({ id, orderIndex, result, updatedBy: au.email });

    invalidatePdDataCache_(au.email);
    logActivity_(au.email, 'submitPDStatus', `ID: ${id} | Step: ${orderIndex} | Result: ${result}`);

    return okResponse_(null, 'Status berhasil disimpan.');

  } catch (err) {
    Logger.log('submitPDStatus ERROR: ' + err);
    return errResponse_(String(err.message || err), 'INTERNAL_ERROR');
  }
}


// ─────────────────────────────────────────────────
// EDIT REQUEST (field non-kritis)
// ─────────────────────────────────────────────────

function updatePDForm(formObj) {
  try {
    const au = getActiveUser_((formObj && formObj.token) ? formObj.token : '');
    if (!au.allowed || !au.email) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');

    const id   = (formObj.id || '').toString().trim();
    const role = (au.role || '').toLowerCase();
    if (!id) return errResponse_('ID tidak valid.', 'VALIDATION');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PD);
    if (!sh) return errResponse_("Sheet 'ProductDev' tidak ditemukan.", 'SHEET_NOT_FOUND');

    const last   = sh.getLastRow();
    if (last < 2) return errResponse_('ID tidak ditemukan.', 'NOT_FOUND');

    const values = sh.getRange(2, 1, last - 1, 10).getValues();
    let foundRow = -1;
    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || '').toString().trim() === id) { foundRow = i + 2; break; }
    }
    if (foundRow === -1) return errResponse_('ID tidak ditemukan.', 'NOT_FOUND');

    const rowRequester = (values[foundRow - 2][2] || '').toString().toLowerCase();
    if (role !== 'admin' && rowRequester !== au.email) {
      return errResponse_('Anda tidak memiliki izin edit request ini.', 'FORBIDDEN');
    }

    // Hanya field non-kritis yang boleh diubah
    if (formObj.deskripsi    !== undefined) sh.getRange(foundRow, 7).setValue(formObj.deskripsi    || '');
    if (formObj.estimasiQty  !== undefined) sh.getRange(foundRow, 8).setValue(formObj.estimasiQty  || '');
    if (formObj.estimasiOrder !== undefined) {
      sh.getRange(foundRow, 9).setValue(formObj.estimasiOrder ? new Date(formObj.estimasiOrder) : '');
    }
    if (formObj.lppFile && formObj.lppFile.getName) {
      const cfg    = getConfig_();
      const folder = DriveApp.getFolderById(cfg.LPP_FOLDER_ID || LPP_FOLDER_ID_DEFAULT);
      const file   = folder.createFile(
        formObj.lppFile.copyBlob().setName(id + '_rev_' + formObj.lppFile.getName())
      );
      sh.getRange(foundRow, 10).setValue(file.getUrl());
    }

    invalidatePdDataCache_(au.email);
    logActivity_(au.email, 'updatePDForm', 'ID: ' + id);

    return okResponse_(null, 'Request berhasil diperbarui.');

  } catch (err) {
    Logger.log('updatePDForm ERROR: ' + err);
    return errResponse_(String(err.message || err), 'INTERNAL_ERROR');
  }
}


// ─────────────────────────────────────────────────
// CANCEL REQUEST
// ─────────────────────────────────────────────────

function cancelPD(formObj) {
  try {
    const au = getActiveUser_((formObj && formObj.token) ? formObj.token : '');
    if (!au.allowed || !au.email) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');

    const id     = (formObj.id     || '').toString().trim();
    const reason = (formObj.reason || '').toString().trim();
    const role   = (au.role || '').toLowerCase();

    if (!id)     return errResponse_('ID tidak valid.', 'VALIDATION');
    if (!reason) return errResponse_('Alasan cancel harus diisi.', 'VALIDATION');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PD);
    if (!sh) return errResponse_("Sheet 'ProductDev' tidak ditemukan.", 'SHEET_NOT_FOUND');

    const last   = sh.getLastRow();
    if (last < 2) return errResponse_('ID tidak ditemukan.', 'NOT_FOUND');

    const values = sh.getRange(2, 1, last - 1, 3).getValues();
    let foundRow = -1;
    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || '').toString().trim() === id) { foundRow = i + 2; break; }
    }
    if (foundRow === -1) return errResponse_('ID tidak ditemukan.', 'NOT_FOUND');

    const rowRequester = (values[foundRow - 2][2] || '').toString().toLowerCase();
    if (role !== 'admin' && rowRequester !== au.email) {
      return errResponse_('Anda tidak memiliki izin cancel request ini.', 'FORBIDDEN');
    }

    // Tandai di col K
    sh.getRange(foundRow, 11).setValue('cancelled: ' + reason);

    invalidatePdDataCache_(au.email);
    logActivity_(au.email, 'cancelPD', `ID: ${id} | Reason: ${reason}`);

    return okResponse_(null, 'Request ' + id + ' berhasil dibatalkan.');

  } catch (err) {
    Logger.log('cancelPD ERROR: ' + err);
    return errResponse_(String(err.message || err), 'INTERNAL_ERROR');
  }
}


// ─────────────────────────────────────────────────
// EXPORT DATA KE CSV
// ─────────────────────────────────────────────────

function exportPDData(token) {
  try {
    const au = getActiveUser_(token);
    if (!au.allowed) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');

    const result = buildPDData_(au.email, au.role);
    const rows   = result.rows || [];

    const header = [
      'ID', 'Tanggal Create', 'Requester Email', 'Requester Name',
      'Nama Produk', 'Kategori', 'Deskripsi', 'Estimasi Qty',
      'Estimasi Order Date', 'LPP Link', 'Status', 'Last Status Update',
      'Last Updated By', 'Status Note', 'Decision Status', 'SLA Status'
    ];

    const csvRows = rows.map(r =>
      [
        r.id, r.tanggalCreate, r.requesterEmail, r.requesterName,
        r.namaProduk, r.kategori, r.deskripsi, r.estimasiQty,
        r.estimasiOrder, r.lppLink, r.status, r.lastStatusUpdate,
        r.lastUpdatedBy, r.statusNote, r.decisionStatus, r.slaStatus
      ].map(v => '"' + String(v || '').replace(/"/g, '""') + '"').join(',')
    );

    const csv      = [header.join(','), ...csvRows].join('\n');
    const filename = 'PD_Export_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '.csv';

    logActivity_(au.email, 'exportPDData', rows.length + ' rows exported');
    return okResponse_({ csv, filename }, '');

  } catch (err) {
    Logger.log('exportPDData ERROR: ' + err);
    return errResponse_(String(err.message || err), 'INTERNAL_ERROR');
  }
}
