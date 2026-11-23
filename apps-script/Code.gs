/**
 * PROMOTER SHARE PORTFOLIO — Google Apps Script backend
 * ------------------------------------------------------
 * Container-bound to a Google Sheet. Serves the web app (Index.html), stores
 * holdings in the "Holdings" sheet and pulls LTP from nepalstock.com.
 *
 * Quick start
 *   1. Paste Code.gs, Index.html and appsscript.json into Extensions > Apps Script.
 *   2. Run setup() once (authorise when asked).
 *   3. Run testPriceSources() to see which price sources respond (Execution log).
 *   4. Deploy > New deployment > Web app  (Execute as: Me, Access: Only myself) — app hosted by Google.
 *   5. For the GitHub Pages / mobile app: Sheet menu → Promoter Portfolio → Set app PIN / password, then a second deployment with Access: Anyone.
 *
 * Price sources, tried in order: GitHub feed (optional) → nepalstock.com → nepalstock.com.np → merolagani.com.
 * nepalstock.com refuses Google's servers, so use the GitHub feed or the in-app browser bridge for nepalstock.com prices.
 *
 * nepalstock.com notes
 *   nepalstock.com's API needs a short-lived token. GET /api/authenticate/prove
 *   returns an obfuscated accessToken + five salts; the site's css.wasm computes
 *   five character positions to strip from it. decodeNepseToken_() is a plain-JS
 *   port of that wasm (only salt2 is used for the access token). The token is then
 *   sent as "Authorization: Salter <token>".
 */

// ============================== CONFIG ==============================
const APP_TITLE = 'Promoter Share Portfolio';
const SH_HOLDINGS = 'Holdings';
const SH_PRICES = 'Prices';
const SH_COMPANIES = 'Companies';
const SH_LOG = 'Log';
const NEPSE_BASES = ['https://www.nepalstock.com', 'https://www.nepalstock.com.np'];
const MEROLAGANI_URL = 'https://merolagani.com/handlers/webrequesthandler.ashx?type=market_summary';
// Optional: URL of prices.json published by the GitHub Action (raw.githubusercontent.com/...). Set from the app's Settings.
const PROP_FEED = 'PRICE_FEED_URL';
const TZ = 'Asia/Kathmandu';

// Input columns (key, header). Order = sheet column order. Do not rename keys.
const COLS = [
  ['id', 'ID'],
  ['company', 'Company Name'],
  ['symbol', 'NEPSE Symbol'],
  ['sector', 'Sector'],
  ['listing', 'Listing / IPO Status'],
  ['listingDate', 'Listing Date'],
  ['projectStatus', 'Project Status'],
  ['progressPct', 'Construction Progress %'],
  ['dematHolder', 'Demat Account Holder'],
  ['owner', 'Ownership (owner or split e.g. Self 50%, Partner 50%)'],
  ['boid', 'BOID / Demat No.'],
  ['shares', 'No. of Shares'],
  ['bonusShares', 'Bonus Shares'],
  ['costPerShare', 'Rate (Cost per Share)'],
  ['investDate', 'Date of Investment'],
  ['lockinDate', 'Lock-in Release Date (date or 2030E)'],
  ['exitStatus', 'Hold / Partial Exit / Exit'],
  ['exitedQty', 'Sold Qty'],
  ['exitPrice', 'Sold Rate (avg)'],
  ['exitDate', 'Date of Exit'],
  ['fairValue', 'Current / Expected Rate'],
  ['manualLtp', 'Manual LTP Override'],
  ['notes', 'Notes'],
  ['createdAt', 'Created'],
  ['updatedAt', 'Updated']
];
// Computed columns written by the script (read-only reference for sheet users)
const CALC_COLS = [
  ['totalCost', 'Total Cost'],
  ['remainingQty', 'Shares Held (incl. bonus)'],
  ['priceUsed', 'Price Used (LTP / Expected / Cost)'],
  ['currentValue', 'Current Value'],
  ['realizedPL', 'Realized P/L'],
  ['unrealizedPL', 'Unrealized P/L'],
  ['lockStatus', 'Lock-in Status'],
  ['lockedQty', 'Locked Shares'],
  ['liquidQty', 'Liquid Shares'],
  ['lockedValue', 'Locked Value'],
  ['liquidValue', 'Liquid Value'],
  ['totalReturnPct', 'Total Return %']
];
const DATE_KEYS = ['investDate', 'listingDate', 'lockinDate', 'exitDate'];
const NUM_KEYS = ['progressPct', 'shares', 'bonusShares', 'costPerShare', 'exitedQty', 'exitPrice', 'fairValue', 'manualLtp'];
const LISTINGS = ['Listed', 'Pipeline', 'Not Applied'];
const EXIT_STATUSES = ['Hold', 'Partial Exit', 'Exit'];

// ============================== WEB APP ==============================
// ============================== ACCESS CONTROL ==============================
// Two ways in:
//  1. The Apps Script-hosted page (google.script.run): allowed only for the script owner.
//  2. The JSON API (doPost) used by the GitHub Pages / mobile app: password login → 30-day device session.
//     Set the password from the Sheet menu: Promoter Portfolio → Set app PIN / password (owner only).
let API_AUTHED_ = false;
const PROP_PW = 'APP_PW_HASH', PROP_SESS = 'APP_SESSIONS';
const SESSION_DAYS = 30, MAX_FAILS = 5, LOCK_MIN = 60;   // 5 wrong tries → locked for 1 hour

function isOwner_() {
  try { const me = Session.getActiveUser().getEmail(); return !!me && me === Session.getEffectiveUser().getEmail(); } catch (e) { return false; }
}
function authorize_() {
  if (API_AUTHED_ || isOwner_()) return;
  throw new Error('Not authorised');
}
function sha_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function safeEq_(a, b) { a = String(a); b = String(b); if (a.length !== b.length) return false; let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; }
function sessions_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_SESS) || '{}'); } catch (e) { return {}; } }
function saveSessions_(m) {
  const now = Date.now();
  const live = Object.keys(m).filter(k => m[k] > now).sort((x, y) => m[y] - m[x]).slice(0, 25);   // drop expired, keep newest 25
  const out = {}; live.forEach(k => out[k] = m[k]);
  PropertiesService.getScriptProperties().setProperty(PROP_SESS, JSON.stringify(out));
}

/** Sheet menu → Set app password (owner only). Changing it signs out every device. */
function setAppPassword() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Set app password', 'New PIN or password for the Promoter Portfolio app (at least 4 characters):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const pw = r1.getResponseText();
  if (pw.length < 4) { ui.alert('PIN / password must be at least 4 characters.'); return; }
  const r2 = ui.prompt('Confirm', 'Type the same PIN / password again:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (r2.getResponseText() !== pw) { ui.alert('Passwords did not match. Nothing changed.'); return; }
  const salt = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_PW, salt + ':' + sha_(salt + pw));
  props.deleteProperty(PROP_SESS);
  log_('password', 'app password set; all devices signed out');
  ui.alert('Saved. All devices have been signed out; sign in again on each.');
}
/** Sheet menu → Sign out all devices. */
function signOutAllDevices() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SESS);
  log_('password', 'all devices signed out');
  SpreadsheetApp.getUi().alert('All devices signed out.');
}

function login_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pw_fails') || 0);
  if (fails >= MAX_FAILS) throw new Error('Too many wrong attempts. Try again in ' + LOCK_MIN + ' minutes.');
  const stored = PropertiesService.getScriptProperties().getProperty(PROP_PW);
  if (!stored) throw new Error('No app PIN set yet. In the Google Sheet use the menu Promoter Portfolio → Set app PIN / password.');
  const [salt, hash] = stored.split(':');
  if (!safeEq_(sha_(salt + String(pw || '')), hash)) {
    cache.put('pw_fails', String(fails + 1), LOCK_MIN * 60);
    Utilities.sleep(1000);
    throw new Error('Wrong PIN' + (MAX_FAILS - fails - 1 > 0 ? ' (' + (MAX_FAILS - fails - 1) + ' attempts left)' : ''));
  }
  cache.remove('pw_fails');
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { const m = sessions_(); m[sha_(token)] = Date.now() + SESSION_DAYS * 864e5; saveSessions_(m); } finally { lock.releaseLock(); }
  log_('login', 'device signed in');
  return { token: token, days: SESSION_DAYS };
}
function tokenOk_(token) {
  if (!token) return false;
  const exp = sessions_()[sha_(String(token))];
  return !!exp && exp > Date.now();
}
function logout_(token) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { const m = sessions_(); delete m[sha_(String(token || ''))]; saveSessions_(m); } finally { lock.releaseLock(); }
  return { ok: true };
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** JSON API. Body (text/plain): {"fn": "login", "args": [password]}  or  {"token": "...", "fn": "api_getData", "args": [...]} */
function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.fn === 'login') return json_({ ok: true, result: login_((req.args || [])[0]) });
    if (!tokenOk_(req.token)) return json_({ ok: false, code: 'AUTH', error: 'Please sign in again' });
    if (req.fn === 'logout') return json_({ ok: true, result: logout_(req.token) });
    const fns = { api_getData, api_saveHolding, api_deleteHolding, api_importHoldings, api_segregateHolding, api_refreshPrices,
      api_syncCompanies, api_matchSymbols, api_setAutoRefresh, api_setFeedUrl, api_importPrices };
    const fn = Object.prototype.hasOwnProperty.call(fns, req.fn) ? fns[req.fn] : null;
    if (!fn) return json_({ ok: false, error: 'Unknown action' });
    API_AUTHED_ = true;
    return json_({ ok: true, result: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return json_({ ok: true, app: APP_TITLE });
  if (!isOwner_()) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">Promoter Share Portfolio API is running. Open the app from your GitHub Pages link.</p>').setTitle(APP_TITLE);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Promoter Portfolio')
    .addItem('Setup / repair sheets', 'setup')
    .addItem('Refresh LTP from nepalstock.com', 'menuRefresh_')
    .addItem('Sync company & sector list', 'menuSyncCompanies_')
    .addSeparator()
    .addItem('Turn ON auto-refresh (market hours)', 'enableAutoRefresh')
    .addItem('Turn OFF auto-refresh', 'disableAutoRefresh')
    .addSeparator()
    .addItem('Set app PIN / password', 'setAppPassword')
    .addItem('Sign out all devices', 'signOutAllDevices')
    .addToUi();
}
function menuRefresh_() { const r = api_refreshPrices(); SpreadsheetApp.getActive().toast(r.message, APP_TITLE, 6); }
function menuSyncCompanies_() { const r = api_syncCompanies(); SpreadsheetApp.getActive().toast(r.message, APP_TITLE, 6); }

// ============================== SETUP ==============================
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  const h = ensureSheet_(SH_HOLDINGS, COLS.map(c => c[1]).concat(CALC_COLS.map(c => c[1])));
  h.setFrozenRows(1); h.setFrozenColumns(2);
  h.getRange(1, 1, 1, COLS.length).setBackground('#0C4DA1').setFontColor('#ffffff').setFontWeight('bold');
  h.getRange(1, COLS.length + 1, 1, CALC_COLS.length).setBackground('#ED1B24').setFontColor('#ffffff').setFontWeight('bold');
  h.getRange(2, colOf_('lockinDate'), h.getMaxRows() - 1, 1).setNumberFormat('@'); // holds dates or '2030E'
  ensureSheet_(SH_PRICES, ['Symbol', 'Security Name', 'LTP', 'Previous Close', '% Change', 'Business Date', 'Fetched At', 'Source']).setFrozenRows(1);
  ensureSheet_(SH_COMPANIES, ['Symbol', 'Company Name', 'Sector', 'Instrument', 'Status']).setFrozenRows(1);
  ensureSheet_(SH_LOG, ['Time', 'Event', 'Detail']).setFrozenRows(1);
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  log_('setup', 'Sheets ready');
  return 'Setup complete';
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const cur = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  headers.forEach((hd, i) => { if (cur[i] !== hd) sh.getRange(1, i + 1).setValue(hd); });
  return sh;
}

// ============================== DATA API (called from UI) ==============================
function api_getData() {
  authorize_();
  const holdings = readHoldings_();
  const prices = readPrices_();
  const props = PropertiesService.getScriptProperties();
  return {
    holdings: holdings,
    prices: prices,
    companies: readCompanies_(),
    meta: {
      lastPriceRefresh: props.getProperty('lastPriceRefresh') || '',
      lastPriceStatus: props.getProperty('lastPriceStatus') || '',
      feedUrl: props.getProperty(PROP_FEED) || '',
      autoRefresh: hasTrigger_('scheduledRefresh'),
      today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
      sheetUrl: SpreadsheetApp.getActive().getUrl()
    }
  };
}

function api_saveHolding(h) {
  authorize_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH_HOLDINGS);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    h = sanitize_(h);
    let row = 0;
    if (h.id) row = findRow_(sh, h.id);
    if (!row) { h.id = h.id || ('PS-' + Utilities.getUuid().slice(0, 8).toUpperCase()); h.createdAt = now; row = sh.getLastRow() + 1; }
    else { h.createdAt = sh.getRange(row, headers.indexOf('Created') + 1).getDisplayValue() || now; }
    h.updatedAt = now;
    sh.getRange(row, 1, 1, COLS.length).setValues([COLS.map(c => toCell_(c[0], h[c[0]]))]);
    recalcSheet_();
    log_('save', h.id + ' ' + h.company);
    return { ok: true, id: h.id, data: api_getData() };
  } finally { lock.releaseLock(); }
}

function api_deleteHolding(id) {
  authorize_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH_HOLDINGS);
    const row = findRow_(sh, id);
    if (row) { sh.deleteRow(row); log_('delete', id); }
    return { ok: !!row, data: api_getData() };
  } finally { lock.releaseLock(); }
}

/**
 * Splits a joint holding into one row per owner (UI computes the share split).
 * The original row is replaced in place so the sheet keeps its order.
 */
function api_segregateHolding(id, parts) {
  authorize_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH_HOLDINGS);
    const row = findRow_(sh, id);
    if (!row) throw new Error('Holding not found: ' + id);
    if (!parts || parts.length < 2) throw new Error('Need at least two owners to segregate');
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const out = parts.map(p => {
      const h = sanitize_(p);
      h.id = 'PS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      h.createdAt = now; h.updatedAt = now;
      return COLS.map(c => toCell_(c[0], h[c[0]]));
    });
    sh.insertRowsAfter(row, out.length);
    sh.getRange(row + 1, 1, out.length, COLS.length).setValues(out);
    sh.deleteRow(row);
    recalcSheet_();
    log_('segregate', id + ' → ' + out.length + ' holdings');
    return { ok: true, count: out.length, data: api_getData() };
  } finally { lock.releaseLock(); }
}

/** Bulk import (from CSV in the UI). rows = array of holding objects. */
function api_importHoldings(rows) {
  authorize_();
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = sheet_(SH_HOLDINGS);
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const out = rows.map(r => {
      const h = sanitize_(r);
      h.id = 'PS-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      h.createdAt = now; h.updatedAt = now;
      return COLS.map(c => toCell_(c[0], h[c[0]]));
    });
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, COLS.length).setValues(out);
    try { autoMatchSymbols_(); } catch (e) { }
    recalcSheet_();
    log_('import', out.length + ' rows');
    return { ok: true, count: out.length, data: api_getData() };
  } finally { lock.releaseLock(); }
}

// ============================== PRICES (nepalstock.com) ==============================
function api_refreshPrices() {
  authorize_();
  const props = PropertiesService.getScriptProperties();
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  try {
    const res = fetchPrices_();
    writePrices_(res.rows, stamp, res.source);
    if (res.companies && res.companies.length) writeCompanies_(res.companies);
    else if (sheet_(SH_COMPANIES).getLastRow() < 2 && res.source.indexOf('nepalstock') === 0) { try { api_syncCompanies(); } catch (e) { log_('companies', e.message); } }
    recalcSheet_();
    const msg = 'LTP updated for ' + res.rows.length + ' securities via ' + res.source + (res.asOf ? ' (data ' + res.asOf + ')' : '');
    props.setProperties({ lastPriceRefresh: stamp, lastPriceStatus: 'OK: ' + msg });
    log_('prices', msg + (res.errors.length ? ' | skipped: ' + res.errors.join(' | ') : ''));
    return { ok: true, message: msg, data: api_getData() };
  } catch (e) {
    const msg = 'LTP refresh failed: ' + e.message;
    props.setProperty('lastPriceStatus', 'ERR ' + stamp + ': ' + e.message);
    log_('prices-error', e.message);
    return { ok: false, message: msg, data: api_getData() };
  }
}

function writePrices_(rows, stamp, source) {
  const sh = sheet_(SH_PRICES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 8).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows.map(r => [r.symbol, r.name || '', r.ltp, r.prevClose || '', r.pct === 0 ? 0 : (r.pct || ''), r.businessDate || '', stamp, source]));
}
function writeCompanies_(list) {
  const sh = sheet_(SH_COMPANIES);
  const cr = list.filter(c => c && c[0]).map(c => [String(c[0]).toUpperCase(), c[1] || '', c[2] || '', '', '']);
  if (!cr.length) return 0;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  sh.getRange(2, 1, cr.length, 5).setValues(cr);
  autoMatchSymbols_();
  return cr.length;
}
/** Payload format shared by the GitHub feed, the browser bridge and pasted data:
 *  { src, at, p: [[symbol, ltp, prevClose, pct, businessDate], ...], c: [[symbol, name, sector], ...] } */
function payloadRows_(payload) {
  return (payload.p || []).filter(r => r && r[0] && num_(r[1])).map(r => ({ symbol: String(r[0]).toUpperCase().trim(), name: '', ltp: num_(r[1]), prevClose: num_(r[2]), pct: r[3] === '' || r[3] === null || r[3] === undefined ? '' : num_(r[3]), businessDate: r[4] || '' }));
}

/** Prices pushed from the browser (nepalstock.com bridge bookmark, or a pasted Today's Price table). */
function api_importPrices(payload) {
  authorize_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const rows = payloadRows_(payload);
    if (!rows.length) throw new Error('No prices found in the pasted data');
    const src = String(payload.src || 'pasted');
    writePrices_(rows, stamp, src);
    const cos = payload.c && payload.c.length ? writeCompanies_(payload.c) : 0;
    recalcSheet_();
    const msg = 'LTP updated for ' + rows.length + ' securities from ' + src + (cos ? ' · ' + cos + ' companies synced' : '');
    PropertiesService.getScriptProperties().setProperties({ lastPriceRefresh: stamp, lastPriceStatus: 'OK: ' + msg });
    log_('prices-import', msg);
    return { ok: true, message: msg, data: api_getData() };
  } finally { lock.releaseLock(); }
}

function api_setFeedUrl(url) {
  authorize_();
  url = String(url || '').trim();
  if (url && !/^https:\/\//.test(url)) throw new Error('Feed URL must start with https://');
  const props = PropertiesService.getScriptProperties();
  if (url) props.setProperty(PROP_FEED, url); else props.deleteProperty(PROP_FEED);
  return { ok: true, message: url ? 'Price feed saved' : 'Price feed removed', data: api_getData() };
}

/** Pulls company list (symbol, name, sector) — used to auto-fill sector in the form. */
function api_syncCompanies() {
  authorize_();
  let list, lastErr;
  for (const base of NEPSE_BASES) {
    try { list = nepseGet_(base, '/api/nots/company/list', getNepseToken_(base)); break; } catch (e) { lastErr = e; }
  }
  if (!list) {
    try { const f = fetchFeed_(); if (f.companies.length) { const n = writeCompanies_(f.companies); return { ok: true, message: 'Company list synced from price feed: ' + n + ' companies', data: api_getData() }; } } catch (e) { }
    throw new Error('nepalstock.com is not reachable from Google servers (' + (lastErr && lastErr.message) + '). Use the GitHub price feed or the nepalstock.com bridge in Prices & Settings.');
  }
  const arr = Array.isArray(list) ? list : (list && list.content) || [];
  const rows = arr.filter(c => c.symbol).map(c => [String(c.symbol).toUpperCase(), c.companyName || c.securityName || '', c.sectorName || '', c.instrumentType || '', c.status || '']);
  const sh = sheet_(SH_COMPANIES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
  log_('companies', rows.length + ' companies');
  const matched = autoMatchSymbols_();
  return { ok: true, message: 'Company list synced: ' + rows.length + ' companies' + (matched ? ' · symbols matched for ' + matched + ' holdings' : ''), data: api_getData() };
}

/** Tries each price source in order and returns the first that works. */
function fetchPrices_() {
  const errors = [];
  if (PropertiesService.getScriptProperties().getProperty(PROP_FEED)) {
    try { const f = fetchFeed_(); if (f.rows.length) return { rows: f.rows, companies: f.companies, source: f.src + ' via GitHub feed', asOf: f.at, errors: errors }; errors.push('feed: no rows'); }
    catch (e) { errors.push('GitHub feed: ' + e.message); }
  }
  for (const base of NEPSE_BASES) {
    try { const rows = fetchNepsePrices_(base); if (rows.length) return { rows: rows, source: base.replace('https://www.', ''), errors: errors }; errors.push(base + ': no rows'); }
    catch (e) { errors.push(base.replace('https://www.', '') + ': ' + e.message); }
  }
  try { const rows = fetchMerolaganiPrices_(); if (rows.length) return { rows: rows, source: 'merolagani.com', errors: errors }; errors.push('merolagani: no rows'); }
  catch (e) { errors.push('merolagani.com: ' + e.message); }
  throw new Error(errors.join(' | '));
}

/** prices.json written by the GitHub Action (same payload format as the browser bridge). */
function fetchFeed_() {
  const url = PropertiesService.getScriptProperties().getProperty(PROP_FEED);
  if (!url) throw new Error('no feed URL set');
  const res = UrlFetchApp.fetch(url + (url.indexOf('?') < 0 ? '?' : '&') + 't=' + Date.now(), { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache' } });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  const j = JSON.parse(res.getContentText());
  const at = j.at ? Utilities.formatDate(new Date(j.at), TZ, 'yyyy-MM-dd HH:mm') : '';
  return { rows: payloadRows_(j), companies: j.c || [], src: j.src || 'nepalstock.com', at: at };
}

function fetchNepsePrices_(base) {
  const token = getNepseToken_(base);
  let data = nepseGet_(base, '/api/nots/securityDailyTradeStat/58', token);
  if (!Array.isArray(data)) data = (data && (data.content || data.data)) || [];
  return data.filter(r => r && r.symbol).map(r => {
    const ltp = num_(firstDefined_(r.lastTradedPrice, r.ltp, r.closePrice, r.lastUpdatedPrice));
    const prev = num_(firstDefined_(r.previousClose, r.previousDayClosePrice));
    return {
      symbol: String(r.symbol).toUpperCase().trim(),
      name: r.securityName || r.companyName || '',
      ltp: ltp || prev || '',
      prevClose: prev || '',
      pct: num_(firstDefined_(r.percentageChange, r.perChange)) || (ltp && prev ? +(((ltp - prev) / prev) * 100).toFixed(2) : ''),
      businessDate: r.businessDate || r.lastUpdatedDateTime || r.lastUpdatedTime || ''
    };
  });
}

/** Fallback: Merolagani live market summary (stock.detail = [{s: symbol, lp: last price, c: % change, q: qty}]). */
function fetchMerolaganiPrices_() {
  const res = UrlFetchApp.fetch(MEROLAGANI_URL, { muteHttpExceptions: true, followRedirects: true,
    headers: { 'Accept': 'application/json, text/plain, */*', 'Referer': 'https://merolagani.com/LatestMarket.aspx', 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36' } });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  const j = JSON.parse(res.getContentText());
  const det = (j && j.stock && j.stock.detail) || [];
  const date = (j && j.mt) || '';
  return det.filter(r => r && r.s && num_(r.lp)).map(r => {
    const ltp = num_(r.lp), pct = num_(r.c);
    return { symbol: String(r.s).toUpperCase().trim(), name: '', ltp: ltp, prevClose: pct > -100 ? +(ltp / (1 + pct / 100)).toFixed(2) : '', pct: pct, businessDate: String(date) };
  });
}

function getNepseToken_(base) {
  const prove = nepseGet_(base, '/api/authenticate/prove', null);
  if (!prove || !prove.accessToken) throw new Error('Token handshake failed (no accessToken)');
  return decodeNepseToken_(prove.accessToken, prove.salt2);
}

/** Port of nepalstock.com css.wasm (cdx/rdx/bdx/ndx/mdx). Only salt2 drives the access token. */
function decodeNepseToken_(token, salt2) {
  const T = [5, 8, 4, 7, 9, 4, 6, 9, 5, 5, 6, 5, 3, 5, 4, 4, 9, 6, 6, 8, 8, 6, 8, 6, 5, 8, 4, 9, 5, 9, 8, 5, 3, 4, 7, 7, 4, 7, 3, 9];
  const s = parseInt(salt2, 10);
  const A = Math.trunc(s / 100) % 10, B = Math.trunc(s / 10) % 10, C = s % 10;
  const t = T[A + B + C];
  const n = t + 22, l = A + B + t + 32, o = A + B + t + 60, p = B + t + 88, q = A + t + 110;
  return token.slice(0, n) + token.slice(n + 1, l) + token.slice(l + 1, o) + token.slice(o + 1, p) + token.slice(p + 1, q) + token.slice(q + 1);
}

function nepseGet_(base, path, token, attempt) {
  attempt = attempt || 0;
  const headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.5',
    'Referer': base + '/',
    'Pragma': 'no-cache',
    'Cache-Control': 'no-cache',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
  };
  if (token) { headers['Authorization'] = 'Salter ' + token; headers['Content-Type'] = 'application/json'; }
  const opts = { method: 'get', headers: headers, muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true };
  let res;
  try { res = UrlFetchApp.fetch(base + path, opts); }
  catch (e) {
    // nepalstock.com has at times served an incomplete certificate chain. Public price data only, so retry once without strict validation.
    if (/certificate|ssl|handshake/i.test(e.message)) { opts.validateHttpsCertificates = false; res = UrlFetchApp.fetch(base + path, opts); }
    else throw new Error(/Address unavailable|DNS|timed out/i.test(e.message) ? 'not reachable from Google servers' : e.message);
  }
  const code = res.getResponseCode();
  if (code === 401 && token && attempt < 2) return nepseGet_(base, path, getNepseToken_(base), attempt + 1);
  if ((code === 502 || code === 503 || code === 429) && attempt < 2) { Utilities.sleep(2000); return nepseGet_(base, path, token, attempt + 1); }
  if (code < 200 || code >= 300) throw new Error(path + ' → HTTP ' + code);
  const txt = res.getContentText();
  try { return JSON.parse(txt); } catch (e) { throw new Error('Unexpected (non-JSON) reply: ' + txt.slice(0, 120)); }
}

/** Run from the editor: checks every price source and logs which ones work. */
function testPriceSources() {
  const feed = PropertiesService.getScriptProperties().getProperty(PROP_FEED);
  if (feed) { try { const f = fetchFeed_(); Logger.log('✅ GitHub feed: ' + f.rows.length + ' rows, data as of ' + f.at); } catch (e) { Logger.log('❌ GitHub feed: ' + e.message); } }
  else Logger.log('ℹ️ GitHub feed: not set (optional; see SETUP.md)');
  NEPSE_BASES.forEach(base => {
    try { const r = fetchNepsePrices_(base); Logger.log('✅ ' + base + ': ' + r.length + ' rows, e.g. ' + JSON.stringify(r.slice(0, 2))); }
    catch (e) { Logger.log('❌ ' + base + ': ' + e.message); }
  });
  try { const r = fetchMerolaganiPrices_(); Logger.log('✅ merolagani.com: ' + r.length + ' rows, e.g. ' + JSON.stringify(r.slice(0, 2))); }
  catch (e) { Logger.log('❌ merolagani.com: ' + e.message); }
}
function testNepseConnection() { testPriceSources(); }

/** Fills NEPSE symbol for Listed holdings that have none, by matching company name to the Companies list. */
function autoMatchSymbols_() {
  const cos = readCompanies_(); if (!cos.length) return 0;
  const sh = sheet_(SH_HOLDINGS); if (sh.getLastRow() < 2) return 0;
  const norm = s => String(s || '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\b(ltd|limited|pvt|private|co|company|the|public)\b/g, ' ').replace(/\s+/g, '');
  const cmap = cos.map(c => ({ sym: c.symbol, n: norm(c.name) })).filter(c => c.n);
  const symCol = colOf_('symbol'), nameCol = colOf_('company'), listCol = colOf_('listing');
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues();
  let n = 0;
  vals.forEach((r, i) => {
    if (r[symCol - 1] || r[listCol - 1] !== 'Listed') return;
    const hn = norm(r[nameCol - 1]); if (hn.length < 5) return;
    const hits = cmap.filter(c => c.n === hn || c.n.indexOf(hn) === 0 || hn.indexOf(c.n) === 0);
    if (hits.length === 1) { sh.getRange(i + 2, symCol).setValue(hits[0].sym); n++; }
  });
  if (n) log_('symbols', 'auto-matched ' + n);
  return n;
}
function api_matchSymbols() {
  authorize_();
  const n = autoMatchSymbols_(); recalcSheet_();
  return { ok: true, message: n ? 'Matched symbols for ' + n + ' holdings' : 'No new symbol matches (sync the company list first, or enter symbols manually)', data: api_getData() };
}

// ============================== AUTO REFRESH ==============================
function enableAutoRefresh() {
  disableAutoRefresh();
  ScriptApp.newTrigger('scheduledRefresh').timeBased().everyMinutes(15).create();
  return 'Auto-refresh ON (every 15 min, Sun–Thu 11:00–15:15 NPT)';
}
function disableAutoRefresh() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'scheduledRefresh').forEach(t => ScriptApp.deleteTrigger(t));
  return 'Auto-refresh OFF';
}
function api_setAutoRefresh(on) {
  authorize_(); const m = on ? enableAutoRefresh() : disableAutoRefresh(); return { ok: true, message: m, data: api_getData() }; }

function scheduledRefresh() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1=Mon … 7=Sun
  const hm = Number(Utilities.formatDate(now, TZ, 'HHmm'));
  const tradingDay = (dow === 7 || dow <= 4); // Sun–Thu
  if (tradingDay && hm >= 1100 && hm <= 1515) api_refreshPrices();
  else if (tradingDay && hm > 1515 && hm <= 1535) api_refreshPrices(); // one close-of-day pull
}
function hasTrigger_(fn) { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn); }

// ============================== SHEET HELPERS ==============================
function sheet_(name) { const sh = SpreadsheetApp.getActive().getSheetByName(name); if (!sh) { setup(); return SpreadsheetApp.getActive().getSheetByName(name); } return sh; }

function readHoldings_() {
  const sh = sheet_(SH_HOLDINGS);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = values[0];
  const idx = {}; COLS.forEach((c, i) => { const j = headers.indexOf(c[1]); idx[c[0]] = j >= 0 ? j : (i < headers.length ? i : -1); });   // falls back to column position if a header was renamed
  return values.slice(1).filter(r => r[idx.id] !== '' || r[idx.company] !== '').map(r => {
    const o = {};
    COLS.forEach(c => { const i = idx[c[0]]; o[c[0]] = i < 0 ? '' : fromCell_(c[0], r[i]); });
    if (!o.id) o.id = '';
    return o;
  });
}

function readPrices_() {
  const sh = sheet_(SH_PRICES);
  if (sh.getLastRow() < 2) return {};
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => {
    if (r[0]) out[String(r[0]).toUpperCase()] = { name: r[1], ltp: num_(r[2]), prevClose: num_(r[3]), pct: r[4] === '' ? null : num_(r[4]), businessDate: r[5] instanceof Date ? Utilities.formatDate(r[5], TZ, 'yyyy-MM-dd') : String(r[5] || ''), fetchedAt: r[6] instanceof Date ? Utilities.formatDate(r[6], TZ, 'yyyy-MM-dd HH:mm') : String(r[6] || ''), source: String(r[7] || '') };
  });
  return out;
}

function readCompanies_() {
  const sh = sheet_(SH_COMPANIES);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().filter(r => r[0]).map(r => ({ symbol: String(r[0]), name: String(r[1]), sector: String(r[2]) }));
}

function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/** Writes the computed columns for every row (values, so the sheet stays readable/sortable). */
function recalcSheet_() {
  const sh = sheet_(SH_HOLDINGS);
  if (sh.getLastRow() < 2) return;
  const prices = readPrices_();
  const hs = readHoldings_();
  const out = hs.map(h => {
    const c = calc_(h, prices);
    return [c.totalCost, c.remainingQty, c.priceUsed, c.currentValue, c.realizedPL, c.unrealizedPL, c.lockStatus, c.lockedQty, c.liquidQty, c.lockedValue, c.liquidValue, c.totalReturnPct === '' ? '' : c.totalReturnPct / 100];
  });
  const start = COLS.length + 1;
  sh.getRange(2, start, out.length, CALC_COLS.length).setValues(out);
  sh.getRange(2, start, out.length, 6).setNumberFormat('#,##0.00');
  sh.getRange(2, start + 7, out.length, 2).setNumberFormat('#,##0');
  sh.getRange(2, start + 9, out.length, 2).setNumberFormat('#,##0.00');
  sh.getRange(2, start + 11, out.length, 1).setNumberFormat('0.00%');
}

function calc_(h, prices) {
  const orig = num_(h.shares) || 0, bonus = num_(h.bonusShares) || 0, total = orig + bonus;
  const totalCost = orig * (num_(h.costPerShare) || 0);
  const cb = total ? totalCost / total : 0; // cost basis per share after bonus
  const exQ = Math.min(num_(h.exitedQty) || 0, total), exP = num_(h.exitPrice) || 0;
  const rem = total - exQ;
  const px = prices[String(h.symbol || '').toUpperCase()];
  let price = cb;
  if (num_(h.manualLtp)) price = num_(h.manualLtp);
  else if (h.listing === 'Listed' && px && px.ltp) price = px.ltp;
  else if (num_(h.fairValue)) price = num_(h.fairValue);
  const realized = exQ * (exP - cb);
  const unreal = rem * (price - cb);
  // Locked = still held and lock-in release (exact date, or expected year 'YYYYE' = end of that year) is in the future.
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const rel = releaseDate_(h.lockinDate);
  const locked = !!(rel && rel > today);
  const lockedQty = locked ? rem : 0, liquidQty = locked ? 0 : rem;
  return { totalCost: totalCost, remainingQty: rem, priceUsed: price, currentValue: rem * price, realizedPL: realized, unrealizedPL: unreal,
    lockStatus: rem <= 0 ? 'Exited' : (locked ? 'Locked' : 'Released'), lockedQty: lockedQty, liquidQty: liquidQty, lockedValue: lockedQty * price, liquidValue: liquidQty * price,
    totalReturnPct: totalCost ? ((realized + unreal) / totalCost) * 100 : '' };
}
function releaseDate_(v) {
  v = String(v || '').trim();
  let m = v.match(/^(\d{4})\s*E$/i); if (m) return m[1] + '-12-31';
  m = v.match(/^(\d{4}-\d{2}-\d{2})/); return m ? m[1] : '';
}
function colOf_(key) { for (let i = 0; i < COLS.length; i++) if (COLS[i][0] === key) return i + 1; return 1; }

function sanitize_(h) {
  const o = {};
  COLS.forEach(c => { o[c[0]] = h[c[0]] === undefined || h[c[0]] === null ? '' : h[c[0]]; });
  NUM_KEYS.forEach(k => { o[k] = o[k] === '' ? '' : num_(o[k]); });
  ['company', 'symbol', 'sector', 'dematHolder', 'owner', 'boid', 'notes', 'lockinDate'].forEach(k => o[k] = String(o[k]).trim());
  o.symbol = o.symbol.toUpperCase();
  if (o.listing === 'Unlisted' || /not\s*applied/i.test(o.listing)) o.listing = 'Not Applied';
  if (/pipeline/i.test(o.listing)) o.listing = 'Pipeline';
  if (LISTINGS.indexOf(o.listing) < 0) o.listing = o.symbol ? 'Listed' : 'Not Applied';
  if (['Under Construction', 'Operational'].indexOf(o.projectStatus) < 0) o.projectStatus = num_(o.progressPct) && num_(o.progressPct) < 100 ? 'Under Construction' : 'Operational';
  if (/^\d{4}\s*e$/i.test(o.lockinDate)) o.lockinDate = o.lockinDate.replace(/\s/g, '').toUpperCase();
  const total = (num_(o.shares) || 0) + (num_(o.bonusShares) || 0), ex = num_(o.exitedQty) || 0;
  if (o.exitStatus === 'Exit' && !ex && total) o.exitedQty = total;
  if (EXIT_STATUSES.indexOf(o.exitStatus) < 0) { const q = num_(o.exitedQty) || 0; o.exitStatus = q <= 0 ? 'Hold' : (q >= total ? 'Exit' : 'Partial Exit'); }
  if (!o.company) throw new Error('Company name is required');
  if (!(num_(o.shares) > 0)) throw new Error('No. of shares must be greater than 0');
  if ((num_(o.exitedQty) || 0) > total) throw new Error('Sold qty cannot exceed shares + bonus');
  return o;
}

function toCell_(key, v) {
  if (key === 'lockinDate') return String(v || ''); // text column: '2027-08-14' or '2030E'
  if (DATE_KEYS.indexOf(key) >= 0) {
    if (!v) return '';
    const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : v;
  }
  return v;
}
function fromCell_(key, v) {
  if (key === 'lockinDate' && v && !(v instanceof Date)) {
    const t = String(v).trim();
    if (/^\d{4}\s*E$/i.test(t)) return t.replace(/\s/g, '').toUpperCase();
    if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.slice(0, 10);
    const d = new Date(t); return isNaN(d) ? t : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  }
  if (v instanceof Date) return Utilities.formatDate(v, TZ, DATE_KEYS.indexOf(key) >= 0 ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm');
  if (NUM_KEYS.indexOf(key) >= 0) return v === '' ? '' : num_(v);
  return v === null || v === undefined ? '' : String(v);
}
function num_(v) { if (v === '' || v === null || v === undefined) return 0; const n = Number(String(v).replace(/[, ]/g, '')); return isFinite(n) ? n : 0; }
function firstDefined_() { for (let i = 0; i < arguments.length; i++) if (arguments[i] !== undefined && arguments[i] !== null && arguments[i] !== '') return arguments[i]; return ''; }
function log_(ev, detail) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH_LOG); if (!sh) return;
    sh.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), ev, String(detail).slice(0, 500)]);
    if (sh.getLastRow() > 600) sh.deleteRows(2, 100);
  } catch (e) { }
}
