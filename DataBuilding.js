/**
 * DataBuilding.js — Bangun data dashboard dari 3 sheet
 *
 * Join: ProductDev + MasterStatus + PD_StatusLog
 * Fitur: RBAC per role, SLA tracking, CacheService, history log
 */


// ─────────────────────────────────────────────────
// CORE DATA BUILDER
// ─────────────────────────────────────────────────

function buildPDData_(userEmailRaw, userRole) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shPD     = ss.getSheetByName(SHEET_PD);
  const shStatus = ss.getSheetByName(SHEET_STATUS);
  const shLog    = ss.getSheetByName(SHEET_PD_STATUS);

  if (!shPD) throw new Error("Sheet 'ProductDev' tidak ditemukan.");

  const pdValues     = shPD.getDataRange().getValues();
  const statusValues = shStatus ? shStatus.getDataRange().getValues() : [];
  const logValues    = shLog    ? shLog.getDataRange().getValues()    : [];

  const currentUserEmail = (userEmailRaw || '').toString().toLowerCase().trim();
  const isRequesterRole  = (userRole || '').toLowerCase() === 'requester';
  const cfg              = getConfig_();
  const slaDefaultDays   = Number(cfg.SLA_DEFAULT_DAYS) || 3;

  // ── MasterStatus: map by OrderIndex ──
  const statusMap  = {};
  const statusList = [];

  if (statusValues.length > 1) {
    for (let i = 1; i < statusValues.length; i++) {
      const r        = statusValues[i];
      const name     = (r[0] || '').toString().trim();
      const idx      = Number(r[1]) || 0;
      const email    = (r[2] || '').toString().toLowerCase().trim();
      const slaDays  = Number(r[3]) || slaDefaultDays; // col D opsional per-step SLA
      if (!name || !idx) continue;

      const obj = { orderIndex: idx, name, userEmail: email, slaDays };
      statusMap[idx] = obj;
      statusList.push(obj);
    }
  }

  // Cari index FINAL (status yang namanya dimulai "Final")
  let FINAL_STATUS_INDEX = 0;
  statusList.forEach(s => {
    if (String(s.name).toLowerCase().startsWith('final') && s.orderIndex > FINAL_STATUS_INDEX) {
      FINAL_STATUS_INDEX = s.orderIndex;
    }
  });

  // ── PD_StatusLog: index ID -> { byIndex, maxIndex } ──
  const logMap = {};
  if (logValues.length > 1) {
    for (let i = 1; i < logValues.length; i++) {
      const r  = logValues[i];
      const id = r[0]; if (!id) continue;

      const idx      = Number(r[1]) || 0;
      const result   = (r[2] || '').toString();
      const tanggal  = r[3] instanceof Date ? r[3] : null;
      const ket      = r[4] || '';
      const updBy    = r[5] || '';

      if (!logMap[id]) logMap[id] = { byIndex: {}, maxIndex: 0 };
      const lm       = logMap[id];
      const newEntry = { id, orderIndex: idx, result, tanggal, keterangan: ket, updatedBy: updBy };

      const prev = lm.byIndex[idx];
      if (!prev || (tanggal && prev.tanggal && tanggal > prev.tanggal)) {
        lm.byIndex[idx] = newEntry;
      }
      if (idx > lm.maxIndex) lm.maxIndex = idx;
    }
  }

  // ── Bangun rows ──
  const rows = [];
  for (let i = 1; i < pdValues.length; i++) {
    const r  = pdValues[i];
    const id = r[0]; if (!id) continue;

    const requesterEmail = (r[2] || '').toString().toLowerCase().trim();

    // RBAC: role=requester hanya lihat data milik sendiri
    if (isRequesterRole && requesterEmail !== currentUserEmail) continue;

    const tanggalCreate  = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm') : (r[1] || '');
    const requesterName  = r[3] || '';
    const namaProduk     = r[4] || '';
    const kategori       = r[5] || '';
    const deskripsi      = r[6] || '';
    const estimasiQty    = r[7] || '';
    const estimasiOrder  = r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, 'yyyy-MM-dd') : (r[8] || '');
    const lppLink        = r[9] || '';

    // Col K dipakai untuk flag cancel/archive
    const colK       = (r[10] || '').toString().toLowerCase();
    const isArchived = colK.startsWith('cancelled') || colK === 'archived';
    const cancelReason = isArchived ? (r[10] || '').toString() : '';

    // ── Status dari log ──
    const lm           = logMap[id] || { byIndex: {}, maxIndex: 0 };
    const currentIndex = lm.maxIndex || 0;
    const currentEntry = currentIndex ? lm.byIndex[currentIndex] : null;

    let statusName        = '';
    let lastStatusDate    = '';
    let lastStatusDateRaw = null;
    let lastUpdatedBy     = '';
    let statusNote        = '';
    let decisionStatus    = '';

    if (currentEntry) {
      statusName = (statusMap[currentIndex] || {}).name || '';
      if (currentEntry.tanggal instanceof Date) {
        lastStatusDate    = Utilities.formatDate(currentEntry.tanggal, TZ, 'yyyy-MM-dd HH:mm');
        lastStatusDateRaw = currentEntry.tanggal;
      }
      lastUpdatedBy  = currentEntry.updatedBy  || '';
      statusNote     = currentEntry.keterangan || '';
      decisionStatus = currentEntry.result     || '';
    }

    // ── Next step ──
    const nextIndexObj = (function () {
      if (isArchived) return null;
      const dec = (decisionStatus || '').toLowerCase();
      let targetIdx;

      if (!currentIndex)                              targetIdx = 1;
      else if (dec === 'ok') {
        if (FINAL_STATUS_INDEX && currentIndex >= FINAL_STATUS_INDEX) return null;
        targetIdx = currentIndex + 1;
      }
      else if (dec === 'not ok' || dec === 'notok')  targetIdx = currentIndex;
      else                                            targetIdx = currentIndex || 1;

      const st = statusMap[targetIdx];
      return st ? { index: targetIdx, name: st.name, userEmail: st.userEmail } : null;
    })();

    const canUpdate = !!(
      nextIndexObj &&
      currentUserEmail &&
      nextIndexObj.userEmail &&
      nextIndexObj.userEmail === currentUserEmail
    );

    // ── SLA ──
    const isFinal  = !!(FINAL_STATUS_INDEX && currentIndex >= FINAL_STATUS_INDEX && decisionStatus.toLowerCase() === 'ok');
    const isNew    = !currentIndex;
    const stepSla  = (statusMap[currentIndex] || {}).slaDays || slaDefaultDays;
    const slaStatus = isArchived ? 'cancelled' : getSlaStatus_(lastStatusDateRaw, stepSla, isNew, isFinal);

    rows.push({
      id,
      tanggalCreate,
      requesterEmail,
      requesterName,
      namaProduk,
      kategori,
      deskripsi,
      estimasiQty,
      estimasiOrder,
      lppLink,
      status:          statusName,
      lastStatusUpdate: lastStatusDate,
      lastUpdatedBy,
      statusNote,
      decisionStatus,
      currentIndex,
      nextIndex:       nextIndexObj ? nextIndexObj.index : 0,
      nextStatusName:  nextIndexObj ? nextIndexObj.name  : '',
      canUpdate,
      slaStatus,
      isArchived,
      cancelReason
    });
  }

  return { rows, statusList, finalStatusIndex: FINAL_STATUS_INDEX };
}


// ─────────────────────────────────────────────────
// PUBLIC API: DATA DASHBOARD
// ─────────────────────────────────────────────────

function getPDDataJson(token) {
  const au = getActiveUser_(token);
  Logger.log('getPDDataJson: ' + JSON.stringify({ email: au.email, role: au.role, allowed: au.allowed }));

  if (!au.allowed) {
    const reason   = au.reason || '';
    const loginUrl = getConfig_().DASHBOARD_URL || '';
    const msg =
      reason === 'inactive' ? 'Akun Anda tidak aktif. Hubungi admin.' :
                              'Sesi tidak valid. Silakan login ulang.';
    return Object.assign(errResponse_(msg, 'TOKEN_EXPIRED'), { loginUrl });
  }

  const email = (au.email || '').toLowerCase().trim();
  if (!email) return errResponse_('Email tidak ditemukan.', 'NO_EMAIL');

  // ── Cache ──
  const cacheKey = CACHE_KEY_PDDATA + email;
  const cache    = CacheService.getScriptCache();
  const cached   = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* stale */ }
  }

  const result = buildPDData_(email, au.role);
  const payload = {
    success:          true,
    rows:             result.rows,
    statusList:       result.statusList,
    finalStatusIndex: result.finalStatusIndex,
    userEmail:        email,
    userRole:         au.role || ''
  };

  try { cache.put(cacheKey, JSON.stringify(payload), CACHE_TTL_PD); } catch (e) {}
  return payload;
}


// ─────────────────────────────────────────────────
// PUBLIC API: HISTORY LOG PER ID
// ─────────────────────────────────────────────────

function getPDHistory(id, token) {
  const au = getActiveUser_(token);
  if (!au.allowed) return errResponse_('Akses ditolak.', 'NOT_ALLOWED');
  if (!id)         return errResponse_('ID tidak valid.', 'INVALID_ID');

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const shLog  = ss.getSheetByName(SHEET_PD_STATUS);
  const shStat = ss.getSheetByName(SHEET_STATUS);
  if (!shLog)  return okResponse_([], '');

  // Build status name map
  const nameMap = {};
  if (shStat && shStat.getLastRow() >= 2) {
    shStat.getRange(2, 1, shStat.getLastRow() - 1, 2).getValues().forEach(r => {
      const nm  = (r[0] || '').toString().trim();
      const idx = Number(r[1]) || 0;
      if (nm && idx) nameMap[idx] = nm;
    });
  }

  const idStr  = String(id).trim();
  const logVals = shLog.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < logVals.length; i++) {
    const r = logVals[i];
    if ((r[0] || '').toString().trim() !== idStr) continue;
    const idx = Number(r[1]) || 0;
    history.push({
      orderIndex: idx,
      statusName: nameMap[idx] || ('Step ' + idx),
      result:     (r[2] || '').toString(),
      tanggal:    r[3] instanceof Date ? Utilities.formatDate(r[3], TZ, 'yyyy-MM-dd HH:mm') : String(r[3] || ''),
      keterangan: (r[4] || '').toString(),
      updatedBy:  (r[5] || '').toString()
    });
  }

  history.sort((a, b) => a.orderIndex - b.orderIndex);
  return okResponse_(history, '');
}
