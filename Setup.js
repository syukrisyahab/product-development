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
    headers: ['ID', 'OrderIndex', 'Result', 'TanggalUpdate', 'Keterangan', 'UpdatedBy']
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
  ['INTERNAL_DOMAIN_2', '@rbshera.co.id']
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
