/**
 * Config.js — Konfigurasi terpusat dari sheet "Config"
 *
 * Sheet Config: A=Key, B=Value
 * Keys yang dikenali:
 *   LPP_FOLDER_ID, APP_TITLE, DASHBOARD_URL,
 *   TOKEN_EXPIRY_DAYS, SLA_DEFAULT_DAYS,
 *   CIRCLE_EMAIL_* (multiple rows),
 *   INTERNAL_DOMAIN_* (multiple rows)
 */

const CACHE_KEY_CFG    = 'v1_cfg';
const CACHE_TTL_CFG    = 600; // 10 menit
const CACHE_KEY_PDDATA = 'v1_pd_';
const CACHE_TTL_PD     = 60;  // 60 detik

// ── Fallback defaults (dipakai jika Config sheet kosong / tidak ada) ──
const PD_CIRCLE_EMAILS_DEFAULT = [
  'ppic.board@rbg.co.id',
  'robby.priandra@rbg.co.id',
  'nanda.dewangga@rbg.co.id',
  'novita.dian@rbg.co.id',
  'nurus.safingi@rbg.co.id',
  'lambok.sinaga@rbg.co.id',
  'nopriandi.wasatha@rbg.co.id',
  'apud.mahpudin@rbg.co.id',
  'galih.murti@rbg.co.id',
  'rendra.fatmaja@rbg.co.id',
  'ayu.pratiwi@rbshera.co.id',
  'nova.andi@rbg.co.id'
];

const INTERNAL_DOMAINS_DEFAULT = ['@rbg.co.id', '@rbshera.co.id'];


function getConfig_() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_CFG);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* stale — rebuild */ }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG);

  const cfg = {
    LPP_FOLDER_ID:     LPP_FOLDER_ID_DEFAULT,
    APP_TITLE:         'Product Development Dashboard',
    DASHBOARD_URL:     'https://pd-dashboard-rbg.web.app/',
    APP_URL:           '',
    TOKEN_EXPIRY_DAYS: TOKEN_EXPIRY_DAYS_DEFAULT,
    SLA_DEFAULT_DAYS:  3,
    CIRCLE_EMAILS:     PD_CIRCLE_EMAILS_DEFAULT.slice(),
    INTERNAL_DOMAINS:  INTERNAL_DOMAINS_DEFAULT.slice()
  };

  if (sh && sh.getLastRow() >= 2) {
    const data    = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const circles = [];
    const domains = [];

    for (let i = 0; i < data.length; i++) {
      const key = (data[i][0] || '').toString().trim();
      const val = (data[i][1] || '').toString().trim();
      if (!key || !val) continue;

      if      (key.startsWith('CIRCLE_EMAIL'))    circles.push(val);
      else if (key.startsWith('INTERNAL_DOMAIN')) domains.push(val);
      else {
        switch (key) {
          case 'LPP_FOLDER_ID':     cfg.LPP_FOLDER_ID     = val;                  break;
          case 'APP_TITLE':         cfg.APP_TITLE          = val;                  break;
          case 'DASHBOARD_URL':     cfg.DASHBOARD_URL      = val;                  break;
          case 'APP_URL':           cfg.APP_URL            = val;                  break;
          case 'TOKEN_EXPIRY_DAYS': cfg.TOKEN_EXPIRY_DAYS  = parseInt(val) || 30;  break;
          case 'SLA_DEFAULT_DAYS':  cfg.SLA_DEFAULT_DAYS   = parseInt(val) || 3;   break;
        }
      }
    }

    if (circles.length) cfg.CIRCLE_EMAILS    = circles;
    if (domains.length) cfg.INTERNAL_DOMAINS = domains;
  }

  cache.put(CACHE_KEY_CFG, JSON.stringify(cfg), CACHE_TTL_CFG);
  return cfg;
}


function invalidateConfigCache_() {
  CacheService.getScriptCache().remove(CACHE_KEY_CFG);
}


function invalidatePdDataCache_(email) {
  CacheService.getScriptCache().remove(CACHE_KEY_PDDATA + (email || '').toLowerCase().trim());
}
