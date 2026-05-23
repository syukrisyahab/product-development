/**
 * AdminPanel.js — Fungsi admin (manage users, config, logs)
 * Semua fungsi publik wajib validasi role = 'admin'
 */

function requireAdmin_(token) {
  const au = getActiveUser_(token);
  if (!au.allowed || au.role !== 'admin') return null;
  return au;
}


// ─────────────────────────────────────────────────
// GET ALL ADMIN DATA (users + config + logs)
// ─────────────────────────────────────────────────

function getAdminData(token) {
  if (!requireAdmin_(token)) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shUsers  = ss.getSheetByName(SHEET_USERS);
  const shConfig = ss.getSheetByName(SHEET_CONFIG);
  const cfg      = getConfig_();

  // ── Users ──
  const users = [];
  if (shUsers && shUsers.getLastRow() >= 2) {
    shUsers.getRange(2, 1, shUsers.getLastRow() - 1, 8).getValues().forEach((r, i) => {
      const email = (r[0] || '').toString().trim();
      if (!email) return;
      users.push({
        rowNum:         i + 2,
        email,
        name:           (r[1] || '').toString().trim(),
        role:           (r[2] || '').toString().trim(),
        tokenCreatedAt: r[4] instanceof Date ? formatDateTime_(r[4]) : String(r[4] || ''),
        tokenExpiredAt: r[5] instanceof Date ? formatDateTime_(r[5]) : String(r[5] || ''),
        lastLogin:      r[6] instanceof Date ? formatDateTime_(r[6]) : String(r[6] || ''),
        status:         (r[7] || 'Active').toString().trim()
      });
    });
  }

  // ── Config rows ──
  const configItems = [];
  if (shConfig && shConfig.getLastRow() >= 2) {
    shConfig.getRange(2, 1, shConfig.getLastRow() - 1, 2).getValues().forEach((r, i) => {
      if (r[0]) configItems.push({ rowNum: i + 2, key: String(r[0]), value: String(r[1] || '') });
    });
  }

  const logs = getRecentLogs_(100);

  return okResponse_({ users, configItems, circleEmails: cfg.CIRCLE_EMAILS, logs }, '');
}


// ─────────────────────────────────────────────────
// UPDATE USER (status, role, name)
// ─────────────────────────────────────────────────

function adminUpdateUser(token, payload) {
  const au = requireAdmin_(token);
  if (!au) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');

  const targetEmail = (payload.email || '').toString().toLowerCase().trim();
  const field       = (payload.field || '').toString().trim();
  const value       = (payload.value || '').toString().trim();

  if (!targetEmail || !field) return errResponse_('Parameter tidak lengkap.', 'VALIDATION');

  const COL_MAP = { name: 2, role: 3, status: 8 };
  const col = COL_MAP[field];
  if (!col) return errResponse_('Field tidak dikenali: ' + field, 'VALIDATION');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return errResponse_('Sheet Users tidak ditemukan.', 'SHEET_NOT_FOUND');

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  let foundRow = -1;
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] || '').toString().toLowerCase().trim() === targetEmail) { foundRow = i + 2; break; }
  }
  if (foundRow === -1) return errResponse_('User tidak ditemukan.', 'NOT_FOUND');

  sh.getRange(foundRow, col).setValue(value);
  invalidatePdDataCache_(targetEmail);
  logActivity_(au.email, 'adminUpdateUser', `${targetEmail} | ${field} = ${value}`);

  return okResponse_(null, `User ${targetEmail} berhasil diupdate.`);
}


// ─────────────────────────────────────────────────
// RESET TOKEN USER (force logout)
// ─────────────────────────────────────────────────

function adminResetToken(token, targetEmail) {
  const au = requireAdmin_(token);
  if (!au) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');

  targetEmail = (targetEmail || '').toString().toLowerCase().trim();
  if (!targetEmail) return errResponse_('Email tidak valid.', 'VALIDATION');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return errResponse_('Sheet Users tidak ditemukan.', 'SHEET_NOT_FOUND');

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  let foundRow = -1;
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] || '').toString().toLowerCase().trim() === targetEmail) { foundRow = i + 2; break; }
  }
  if (foundRow === -1) return errResponse_('User tidak ditemukan.', 'NOT_FOUND');

  // Hapus token + expiry — paksa login ulang
  sh.getRange(foundRow, 4).setValue('');
  sh.getRange(foundRow, 5).setValue('');
  sh.getRange(foundRow, 6).setValue('');

  invalidatePdDataCache_(targetEmail);
  logActivity_(au.email, 'adminResetToken', 'target: ' + targetEmail);

  return okResponse_(null, `Token ${targetEmail} direset. User perlu login ulang.`);
}


// ─────────────────────────────────────────────────
// TAMBAH USER BARU
// ─────────────────────────────────────────────────

function adminAddUser(token, payload) {
  const au = requireAdmin_(token);
  if (!au) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');

  const email  = (payload.email  || '').toString().toLowerCase().trim();
  const name   = (payload.name   || '').toString().trim();
  const role   = (payload.role   || 'viewer').toString().trim();
  const status = (payload.status || 'Active').toString().trim();

  if (!email) return errResponse_('Email tidak boleh kosong.', 'VALIDATION');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) return errResponse_('Sheet Users tidak ditemukan.', 'SHEET_NOT_FOUND');

  // Cek duplikat
  if (sh.getLastRow() >= 2) {
    const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if ((existing[i][0] || '').toString().toLowerCase().trim() === email) {
        return errResponse_('Email ' + email + ' sudah terdaftar.', 'DUPLICATE');
      }
    }
  }

  sh.appendRow([email, name, role, '', '', '', '', status]);
  logActivity_(au.email, 'adminAddUser', `${email} | role: ${role}`);

  return okResponse_(null, `User ${email} berhasil ditambahkan.`);
}


// ─────────────────────────────────────────────────
// UPDATE CONFIG ITEM (key-value di sheet Config)
// ─────────────────────────────────────────────────

function adminUpdateConfig(token, payload) {
  const au = requireAdmin_(token);
  if (!au) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');

  const key   = (payload.key   || '').toString().trim();
  const value = (payload.value || '').toString().trim();
  if (!key) return errResponse_('Key tidak boleh kosong.', 'VALIDATION');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   sh = ss.getSheetByName(SHEET_CONFIG);

  if (!sh) {
    sh = ss.insertSheet(SHEET_CONFIG);
    sh.appendRow(['Key', 'Value']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  // Cari baris existing
  let foundRow = -1;
  if (sh.getLastRow() >= 2) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || '').toString().trim() === key) { foundRow = i + 2; break; }
    }
  }

  if (foundRow !== -1) {
    sh.getRange(foundRow, 2).setValue(value);
  } else {
    sh.appendRow([key, value]);
  }

  invalidateConfigCache_();
  logActivity_(au.email, 'adminUpdateConfig', `${key} = ${value}`);

  return okResponse_(null, `Config '${key}' berhasil disimpan.`);
}


// ─────────────────────────────────────────────────
// GET LOGS (untuk admin panel)
// ─────────────────────────────────────────────────

function adminGetLogs(token, limit) {
  if (!requireAdmin_(token)) return errResponse_('Hanya admin yang bisa mengakses fungsi ini.', 'FORBIDDEN');
  return okResponse_(getRecentLogs_(limit || 100), '');
}
