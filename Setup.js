/**
 * Setup.js — Inisialisasi / perapian struktur Google Spreadsheet
 *
 * Fungsi utama : setupSpreadsheetSafe()
 * Aman dijalankan berkali-kali. Tidak pernah menghapus sheet, clear data,
 * atau menimpa isi baris 2 ke bawah.
 *
 * Cara pakai  : Apps Script Editor → pilih setupSpreadsheetSafe → Run
 */


// ─────────────────────────────────────────────────
// KONFIGURASI SHEET & HEADER
// ─────────────────────────────────────────────────

const SETUP_SHEETS_ = [
  {
    name: SHEET_PD,
    headers: [
      'ID', 'TanggalCreate', 'RequesterEmail', 'RequesterName', 'NamaProduk',
      'Kategori', 'Deskripsi', 'EstimasiQty', 'EstimasiOrderDate', 'LPP_URL',
      'CancelFlag', 'LastUpdate', 'LastUpdatedBy', 'Remarks'
    ]
  },
  {
    name: SHEET_STATUS,
    headers: ['NamaStatus', 'OrderIndex', 'PIC_Email', 'SLA_Days']
  },
  {
    name: SHEET_PD_STATUS,
    headers: ['ID', 'OrderIndex', 'Result', 'TanggalUpdate', 'Keterangan', 'UpdatedBy', 'CommitmentDate']
  },
  {
    name: SHEET_USERS,
    headers: ['Email', 'Nama', 'Role', 'Token', 'TokenCreatedAt', 'TokenExpiredAt', 'LastLogin', 'Status']
  },
  {
    name: SHEET_CONFIG,
    headers: ['Key', 'Value'],
    defaultConfig: true
  },
  {
    name: SHEET_LOG,
    headers: ['Timestamp', 'UserEmail', 'Action', 'Detail', 'Status']
  }
];

const DEFAULT_CONFIG_ROWS_ = [
  ['LPP_FOLDER_ID',     ''],
  ['APP_TITLE',         'Product Development Dashboard'],
  ['DASHBOARD_URL',     'https://pd-dashboard-rbg.web.app/'],
  // APP_URL: URL generik Apps Script TANPA prefix /a/macros/domain.com/
  // Isi dengan: https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec
  ['APP_URL',           ''],
  ['TOKEN_EXPIRY_DAYS', '30'],
  ['SLA_DEFAULT_DAYS',  '3'],
  ['INTERNAL_DOMAIN_1', '@rbg.co.id'],
  ['INTERNAL_DOMAIN_2', '@rbshera.co.id'],
  // Circle email — penerima notifikasi semua event PD
  ['CIRCLE_EMAIL_1',    'ppic.board@rbg.co.id'],
  ['CIRCLE_EMAIL_2',    'robby.priandra@rbg.co.id'],
  ['CIRCLE_EMAIL_3',    'nanda.dewangga@rbg.co.id'],
  ['CIRCLE_EMAIL_4',    'novita.dian@rbg.co.id'],
  ['CIRCLE_EMAIL_5',    'nurus.safingi@rbg.co.id'],
  ['CIRCLE_EMAIL_6',    'lambok.sinaga@rbg.co.id'],
  ['CIRCLE_EMAIL_7',    'nopriandi.wasatha@rbg.co.id'],
  ['CIRCLE_EMAIL_8',    'apud.mahpudin@rbg.co.id'],
  ['CIRCLE_EMAIL_9',    'galih.murti@rbg.co.id'],
  ['CIRCLE_EMAIL_10',   'rendra.fatmaja@rbg.co.id'],
  ['CIRCLE_EMAIL_11',   'ayu.pratiwi@rbshera.co.id'],
  ['CIRCLE_EMAIL_12',   'nova.andi@rbg.co.id']
];

const HEADER_BG_COLOR_   = '#1F6B3E'; // hijau
const HEADER_FONT_COLOR_ = '#FFFFFF'; // putih


// ─────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────

function setupSpreadsheetSafe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('══ setupSpreadsheetSafe() START ══');

  for (const def of SETUP_SHEETS_) {
    Logger.log('── ' + def.name);
    const sheet = ensureSheet_(ss, def.name);
    ensureHeadersSafe_(sheet, def.headers);
    styleHeader_(sheet);
    if (def.defaultConfig) ensureDefaultConfig_(sheet);
    Logger.log('   ✓ selesai');
  }

  Logger.log('══ setupSpreadsheetSafe() COMPLETE ══');
}


// ─────────────────────────────────────────────────
// HELPER: ensureSheet_
// ─────────────────────────────────────────────────

/**
 * Mengembalikan sheet yang sudah ada, atau membuat sheet baru.
 * TIDAK PERNAH menghapus sheet manapun.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('   [baru] sheet dibuat: ' + sheetName);
  }
  return sheet;
}


// ─────────────────────────────────────────────────
// HELPER: ensureHeadersSafe_
// ─────────────────────────────────────────────────

/**
 * Memastikan row 1 memiliki header sesuai parameter `headers`.
 *
 * Aturan keamanan:
 *  - Sel yang sudah berisi nilai  → TIDAK diubah sama sekali
 *  - Sel yang kosong              → diisi dengan header yang diharapkan
 *  - Header melebihi kolom existing → ditambahkan di kanan
 *  - Baris 2 ke bawah            → TIDAK disentuh
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 */
function ensureHeadersSafe_(sheet, headers) {
  if (!headers || headers.length === 0) return;

  const needCols  = headers.length;
  const lastCol   = sheet.getLastColumn();
  const readCols  = Math.max(lastCol, needCols);

  // Baca baris 1 selebar kolom yang diperlukan
  // (Membaca melampaui lastColumn di Apps Script mengembalikan string kosong — aman)
  const currentRow = sheet.getRange(1, 1, 1, readCols).getValues()[0];

  let changed = false;

  for (let i = 0; i < needCols; i++) {
    const current = (currentRow[i] !== undefined ? currentRow[i] : '').toString().trim();
    if (current === '') {
      currentRow[i] = headers[i];
      changed = true;
    }
    // Sel tidak kosong → biarkan apa adanya (tidak diubah)
  }

  if (changed) {
    sheet.getRange(1, 1, 1, readCols).setValues([currentRow]);
    Logger.log('   [header] diperbarui: ' + sheet.getName());
  }
}


// ─────────────────────────────────────────────────
// HELPER: ensureDefaultConfig_
// ─────────────────────────────────────────────────

/**
 * Menambahkan baris default ke sheet Config untuk key yang belum ada.
 * Tidak pernah menimpa baris yang sudah ada.
 * Deteksi duplikat berdasarkan nilai kolom A (Key), case-insensitive.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  — sheet Config
 */
function ensureDefaultConfig_(sheet) {
  const lastRow = sheet.getLastRow();

  // Kumpulkan key yang sudah ada (baris 2 ke bawah, kolom A)
  const existingKeys = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => {
      const k = (r[0] || '').toString().trim().toLowerCase();
      if (k) existingKeys.add(k);
    });
  }

  // Tambahkan hanya key yang belum ada
  let added = 0;
  for (const row of DEFAULT_CONFIG_ROWS_) {
    const key = (row[0] || '').toString().trim().toLowerCase();
    if (key && !existingKeys.has(key)) {
      sheet.appendRow(row);
      existingKeys.add(key);
      added++;
    }
  }

  if (added > 0) {
    Logger.log('   [config] ' + added + ' default row ditambahkan');
  }
}


// ─────────────────────────────────────────────────
// HELPER: styleHeader_
// ─────────────────────────────────────────────────

/**
 * Memberi style pada baris 1: bold, background hijau, text putih.
 * Freeze baris pertama. Auto-resize semua kolom.
 * Idempoten — aman dipanggil berkali-kali.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function styleHeader_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header  = sheet.getRange(1, 1, 1, lastCol);

  header
    .setFontWeight('bold')
    .setBackground(HEADER_BG_COLOR_)
    .setFontColor(HEADER_FONT_COLOR_)
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);

  try {
    sheet.autoResizeColumns(1, lastCol);
  } catch (e) {
    // autoResize bisa gagal di sheet yang sangat baru — tidak kritikal
    Logger.log('   [warn] autoResize gagal untuk ' + sheet.getName() + ': ' + e);
  }
}


// ─────────────────────────────────────────────────
// SETUP CIRCLE + USERS
// ─────────────────────────────────────────────────

/**
 * Role assignment:
 *   approver → email yang terdaftar sebagai PIC_Email di MasterStatus (col C)
 *   viewer   → semua circle email lainnya
 *   requester / admin → hanya bisa di-set manual oleh admin, tidak disentuh fungsi ini
 *
 * Lalu:
 *   A. Tambahkan email yang belum ada ke Config sheet sebagai CIRCLE_EMAIL_*
 *   B. Tambahkan email yang belum ada ke Users sheet dengan role yang tepat
 *   C. Email yang SUDAH ADA di Users sheet tidak diubah sama sekali
 *
 * Aman dijalankan berkali-kali.
 * Cara pakai: Apps Script Editor → pilih setupCircleAndUsers → Run
 */
function setupCircleAndUsers() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shConfig = ss.getSheetByName(SHEET_CONFIG);
  const shUsers  = ss.getSheetByName(SHEET_USERS);
  const shStatus = ss.getSheetByName(SHEET_STATUS);

  if (!shConfig) { Logger.log('ERROR: Config sheet tidak ditemukan.'); return; }
  if (!shUsers)  { Logger.log('ERROR: Users sheet tidak ditemukan.'); return; }

  Logger.log('══ setupCircleAndUsers() START ══');

  // ── 1. Kumpulkan PIC_Email dari MasterStatus → ini yang dapat role approver ──
  const picEmails = new Set();
  if (shStatus && shStatus.getLastRow() >= 2) {
    shStatus.getRange(2, 1, shStatus.getLastRow() - 1, 3).getValues().forEach(r => {
      const em = (r[2] || '').toString().toLowerCase().trim();
      if (em && em.includes('@')) picEmails.add(em);
    });
  }
  Logger.log('   PIC emails dari MasterStatus: ' + picEmails.size);

  // ── 2. Kumpulkan semua circle email (default + PIC) ──
  const allEmails = new Set();
  PD_CIRCLE_EMAILS_DEFAULT.forEach(e => {
    const em = e.toString().toLowerCase().trim();
    if (em) allEmails.add(em);
  });
  picEmails.forEach(em => allEmails.add(em));
  Logger.log('   Total email dikumpulkan: ' + allEmails.size);

  // ── 3. Sinkronisasi ke Config (CIRCLE_EMAIL_*) ──
  const existingCircleEmails = new Set();
  let   maxCircleIdx = 0;

  if (shConfig.getLastRow() >= 2) {
    shConfig.getRange(2, 1, shConfig.getLastRow() - 1, 2).getValues().forEach(r => {
      const key = (r[0] || '').toString().trim();
      const val = (r[1] || '').toString().toLowerCase().trim();
      if (key.startsWith('CIRCLE_EMAIL')) {
        if (val) existingCircleEmails.add(val);
        const m = /CIRCLE_EMAIL_(\d+)/i.exec(key);
        if (m) maxCircleIdx = Math.max(maxCircleIdx, parseInt(m[1], 10));
      }
    });
  }

  let addedCircle = 0;
  allEmails.forEach(email => {
    if (!existingCircleEmails.has(email)) {
      maxCircleIdx++;
      shConfig.appendRow(['CIRCLE_EMAIL_' + maxCircleIdx, email]);
      existingCircleEmails.add(email);
      addedCircle++;
      Logger.log('   [circle +] CIRCLE_EMAIL_' + maxCircleIdx + ' = ' + email);
    } else {
      Logger.log('   [circle =] sudah ada: ' + email);
    }
  });

  // ── 4. Sinkronisasi ke Users sheet ──
  // Kumpulkan email yang sudah ada agar tidak ditimpa
  const existingUsers = new Set();
  if (shUsers.getLastRow() >= 2) {
    shUsers.getRange(2, 1, shUsers.getLastRow() - 1, 1).getValues().forEach(r => {
      const em = (r[0] || '').toString().toLowerCase().trim();
      if (em) existingUsers.add(em);
    });
  }

  let addedUsers = 0;
  allEmails.forEach(email => {
    if (existingUsers.has(email)) {
      Logger.log('   [user =] sudah ada (tidak diubah): ' + email);
      return;
    }
    // PIC MasterStatus → approver, sisanya → viewer
    const role = picEmails.has(email) ? 'approver' : 'viewer';
    const name = nameFromEmail_(email);
    shUsers.appendRow([email, name, role, '', '', '', '', 'Active']);
    existingUsers.add(email);
    addedUsers++;
    Logger.log('   [user +] ' + email + ' | role: ' + role + ' | nama: ' + name);
  });

  invalidateConfigCache_();
  Logger.log('══ SELESAI: ' + addedCircle + ' circle baru, ' + addedUsers + ' user baru ══');
}


/**
 * Derivasi nama dari email.
 * Contoh: robby.priandra@rbg.co.id → "Robby Priandra"
 */
function nameFromEmail_(email) {
  const local = (email || '').split('@')[0] || '';
  return local.split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
}
