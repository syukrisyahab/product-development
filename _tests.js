/**
 * _tests.js — Unit test sederhana untuk Apps Script
 * Jalankan dari Script Editor: Run > runAllTests_
 * Hasil dilihat di View > Logs
 */

function runAllTests_() {
  Logger.log('═══ RUNNING ALL TESTS ═══');
  let passed = 0, failed = 0;

  function assert_(label, condition) {
    if (condition) {
      Logger.log('  ✓ PASS: ' + label);
      passed++;
    } else {
      Logger.log('  ✕ FAIL: ' + label);
      failed++;
    }
  }

  try { testFormatHelpers_(assert_); }   catch(e) { Logger.log('ERROR in testFormatHelpers_: ' + e); failed++; }
  try { testGetConfig_(assert_); }       catch(e) { Logger.log('ERROR in testGetConfig_: ' + e); failed++; }
  try { testGetSlaStatus_(assert_); }    catch(e) { Logger.log('ERROR in testGetSlaStatus_: ' + e); failed++; }
  try { testApiResponses_(assert_); }    catch(e) { Logger.log('ERROR in testApiResponses_: ' + e); failed++; }
  try { testIsInternalEmail_(assert_); } catch(e) { Logger.log('ERROR in testIsInternalEmail_: ' + e); failed++; }

  Logger.log('═══ RESULTS: ' + passed + ' passed, ' + failed + ' failed ═══');
  if (failed > 0) throw new Error(failed + ' test(s) failed — lihat Logs');
}


function testFormatHelpers_(assert_) {
  Logger.log('── testFormatHelpers_');

  assert_('formatDateOnly_ null returns empty',
    formatDateOnly_(null) === '');

  assert_('formatDateOnly_ Date returns yyyy-MM-dd',
    /^\d{4}-\d{2}-\d{2}$/.test(formatDateOnly_(new Date(2026, 0, 15))));

  assert_('formatDateOnly_ ISO string strips time',
    formatDateOnly_('2026-01-15T10:30:00Z') === '2026-01-15');

  assert_('formatDateTime_ null returns empty',
    formatDateTime_(null) === '');

  assert_('formatDateTime_ Date returns yyyy-MM-dd HH:mm',
    /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(formatDateTime_(new Date())));
}


function testGetConfig_(assert_) {
  Logger.log('── testGetConfig_');

  const cfg = getConfig_();

  assert_('getConfig_ returns object',
    typeof cfg === 'object' && cfg !== null);

  assert_('getConfig_ has LPP_FOLDER_ID',
    typeof cfg.LPP_FOLDER_ID === 'string' && cfg.LPP_FOLDER_ID.length > 0);

  assert_('getConfig_ has CIRCLE_EMAILS array',
    Array.isArray(cfg.CIRCLE_EMAILS) && cfg.CIRCLE_EMAILS.length > 0);

  assert_('getConfig_ has INTERNAL_DOMAINS array',
    Array.isArray(cfg.INTERNAL_DOMAINS) && cfg.INTERNAL_DOMAINS.length > 0);

  assert_('getConfig_ TOKEN_EXPIRY_DAYS is number',
    typeof cfg.TOKEN_EXPIRY_DAYS === 'number' && cfg.TOKEN_EXPIRY_DAYS > 0);
}


function testGetSlaStatus_(assert_) {
  Logger.log('── testGetSlaStatus_');

  const now  = new Date();
  const old  = new Date(now.getTime() - 5 * 86400000); // 5 hari lalu
  const very = new Date(now.getTime() - 10 * 86400000); // 10 hari lalu

  assert_('getSlaStatus_ new request returns new',
    getSlaStatus_(null, 3, true, false) === 'new');

  assert_('getSlaStatus_ final returns final',
    getSlaStatus_(now, 3, false, true) === 'final');

  assert_('getSlaStatus_ recent update ok',
    getSlaStatus_(now, 3, false, false) === 'ok');

  assert_('getSlaStatus_ 5 days overdue (sla=3)',
    getSlaStatus_(old, 3, false, false) === 'overdue');

  assert_('getSlaStatus_ warning at 75% of SLA',
    getSlaStatus_(new Date(now.getTime() - 2.5 * 86400000), 3, false, false) === 'warning');
}


function testApiResponses_(assert_) {
  Logger.log('── testApiResponses_');

  const ok  = okResponse_({ id: 'test' }, 'Berhasil');
  const err = errResponse_('Gagal', 'TEST_ERR');

  assert_('okResponse_ success=true',  ok.success  === true);
  assert_('okResponse_ has data',      ok.data && ok.data.id === 'test');
  assert_('okResponse_ has message',   ok.message === 'Berhasil');

  assert_('errResponse_ success=false', err.success === false);
  assert_('errResponse_ has code',      err.code    === 'TEST_ERR');
  assert_('errResponse_ has message',   err.message === 'Gagal');
}


function testIsInternalEmail_(assert_) {
  Logger.log('── testIsInternalEmail_');

  assert_('rbg.co.id is internal',     isInternalEmail_('user@rbg.co.id'));
  assert_('rbshera.co.id is internal', isInternalEmail_('user@rbshera.co.id'));
  assert_('gmail.com is external',     !isInternalEmail_('user@gmail.com'));
  assert_('empty string is external',  !isInternalEmail_(''));
}


// ─────────────────────────────────────────────────
// HELPERS untuk integration test (run manual)
// ─────────────────────────────────────────────────

/** Test getPDDataJson dengan token kosong — harus return error */
function testGetPDDataJsonNoToken_() {
  const res = getPDDataJson('');
  Logger.log('getPDDataJson no token: ' + JSON.stringify(res));
}

/** Test getConfig_ — lihat isi config dari sheet */
function testPrintConfig_() {
  invalidateConfigCache_(); // pastikan baca dari sheet, bukan cache
  const cfg = getConfig_();
  Logger.log('Config: ' + JSON.stringify(cfg, null, 2));
}

/** Test getRecentLogs_ */
function testPrintLogs_() {
  const logs = getRecentLogs_(10);
  Logger.log('Recent logs (' + logs.length + '): ' + JSON.stringify(logs, null, 2));
}
