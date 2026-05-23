/**
 * Utils.js — Helper umum: format, API response, SLA
 */

const TZ = 'Asia/Jakarta';


// ─────────────────────────────────────────────────
// FORMAT HELPERS
// ─────────────────────────────────────────────────

function formatDateOnly_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  const s = String(value);
  return s.includes('T') ? s.split('T')[0] : s;
}

function formatDateTime_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, TZ, 'yyyy-MM-dd HH:mm');
  const s = String(value);
  if (!s.includes('T')) return s;
  const [d, t] = s.split('T');
  return d + ' ' + t.substring(0, 5);
}


// ─────────────────────────────────────────────────
// API RESPONSE HELPERS
// ─────────────────────────────────────────────────

function okResponse_(data, message) {
  return { success: true, message: message || '', data: data || null };
}

function errResponse_(message, code) {
  return { success: false, message: message || 'Terjadi kesalahan.', code: code || 'ERROR', data: null };
}


// ─────────────────────────────────────────────────
// SLA HELPER
// ─────────────────────────────────────────────────

/**
 * Hitung status SLA.
 * Return: 'final' | 'new' | 'overdue' | 'warning' | 'ok'
 */
function getSlaStatus_(lastUpdateDate, slaDays, isNew, isFinal) {
  if (isFinal) return 'final';
  if (isNew)   return 'new';
  if (!lastUpdateDate || !(lastUpdateDate instanceof Date)) return 'new';

  const diffDays = (new Date().getTime() - lastUpdateDate.getTime()) / 86400000;

  if (diffDays > slaDays)          return 'overdue';
  if (diffDays > slaDays * 0.75)   return 'warning';
  return 'ok';
}
