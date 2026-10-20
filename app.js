'use strict';

// ── CONFIG ──────────────────────────────────────────────────────────────────
const CLIENT_ID = '37e2a0cc-37b9-4eb5-b2a1-2ca0d9a62db0';
const TENANT_ID = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES    = ['User.Read', 'Sites.ReadWrite.All'];
const SP_SITE   = 'dihag.sharepoint.com:/sites/gruppe_shb';
const SP_LIST   = 'Bedarfsanfrage';
const API       = 'https://graph.microsoft.com/v1.0';
const SP_BASE   = 'https://' + SP_SITE.split(':/')[0] + '/' + SP_SITE.split(':/')[1]; // split(':/')[1] = 'sites/...' (no leading slash)

// ── BPMN-KONFORME FORMFELD-DEFINITION ───────────────────────────────────────
// key        = interner SP-Spaltenname (wird beim POST verwendet)
// alsoTry    = alternative interne Namen falls key nicht gefunden
// step       = Wizard-Schritt (1-3)
const FORM_FIELDS = [
  // Step 1: Bedarf
  { key:'Title',                 label:'Bezeichnung',                    step:1, required:true  },
  { key:'Beschreibung',          label:'Beschreibung',                   step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung','Grund'] },
  { key:'Warengruppe',           label:'Warengruppe',                    step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',            label:'Priorität',                      step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  { key:'AutoBedarfDatum',       label:'Datum zur Automatische Bedarfsanfrage', step:1, alsoTry:['AutomatischeBedarfsanfrageDatum','DatumAutomatischeBedarfsanfrage','DatumZurAutomatischenBedarfsanfrage','AutoBANFDatum'] },
  // Step 2: Menge
  { key:'Menge',                 label:'Menge',                          step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',         label:'Mengeneinheit',                  step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge',     label:'Mindestlagermenge',              step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',                label:'Benötigt bis',                   step:2, required:true, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis','Ben_x00f6_tigtbis'] },
  // Step 3: Beschaffung
  { key:'Artikelnummer',         label:'Artikelnummer',                  step:1, alsoTry:['MaterialNumber','ItemNumber','Artikelnummer_x002f_Nummernangab'] },
  { key:'ExterneArtikelnummer',  label:'Externe Artikelnummer',          step:1, alsoTry:['ExternalItemNumber','ExtArtikelNr','ExterneArtikelnr'] },
  { key:'Beschaffungslogik',     label:'Beschaffungsart',                step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Lieferant',             label:'Lieferant 1',                    step:3, alsoTry:['Vendor','Supplier'] },
  { key:'Lieferant2',            label:'Lieferant 2 (Alternative)',      step:3, alsoTry:['Vendor2','Supplier2','Lieferant_2'] },
  { key:'Lieferant3',            label:'Lieferant 3 (Alternative)',      step:3, alsoTry:['Vendor3','Supplier3','Lieferant_3'] },
  { key:'Lieferant4',            label:'Lieferant 4 (Alternative)',      step:3, alsoTry:['Vendor4','Supplier4','Lieferant_4'] },
  { key:'GeschaetzterPreis',     label:'Bestellvolumen in €',            step:3, alsoTry:['EstimatedPrice','Preis','Price','Gesch_x00e4_tzterPreisnetto_x002'] },
  { key:'Kostenstelle',          label:'Kostenstelle',                   step:3, alsoTry:['CostCenter'] },
  { key:'LeadBuyerAbschluss',    label:'Lead-Buyer-Abschluss',           step:3, alsoTry:['LeadBuyer','LeadBuyerAbschlus'] },
  { key:'Positionen',            label:'Positionen (JSON)',               step:1, alsoTry:['Positions'] },
];

// Felder die Einkauf nach der Einreichung befüllt
const EINKAUF_FIELDS = [
  { key:'Bestellnummer',  label:'Bestellnummer',        alsoTry:['OrderNumber','PO_Number'] },
  { key:'Lieferdatum',    label:'Lieferdatum',          alsoTry:['DeliveryDate'] },
  { key:'TatsaechlicherPreis', label:'Tatsächlicher Preis (€)', alsoTry:['ActualPrice','FinalPrice'] },
];

// Kommentarfeld — separates SP-Listenfeld "Kommentare" (Mehrere Textzeilen)
// Format jedes Eintrags: "[DD.MM.YYYY HH:MM – Autor]: Text\n---\n"
const KOMMENTAR_FIELD = { key: 'Kommentare', alsoTry: ['Comments', 'Kommentar'] };

// Status-Werte → Darstellung (kommt von Power Automate)
// Reihenfolge wichtig: spezifischere (exakte) Werte zuerst, Fallback-Partials am Ende
const STATUS_STYLES = {
  // ── Exakte SP-Auswahlwerte ──
  'eingereicht':                        { bg:'#fce7f3', color:'#be185d' },
  'in prüfung (einkauf)':               { bg:'#ccfbf1', color:'#0f766e' },
  'in prüfung (werkleitung)':           { bg:'#ffe4e6', color:'#be123c' },
  'in prüfung (strategischer einkauf)': { bg:'#ccfbf1', color:'#0f766e' },
  'in prüfung (controlling)':           { bg:'#ccfbf1', color:'#0f766e' },
  'freigegeben':                        { bg:'#f3e8ff', color:'#7e22ce' },
  'abgelehnt':                          { bg:'#f3f4f6', color:'#374151' },
  'bestellt':                           { bg:'#dbeafe', color:'#1d4ed8' },
  'erledigt':                           { bg:'#f3f4f6', color:'#374151' },
  // ── Generische Fallbacks (Teilübereinstimmung) ──
  'in prüfung':                         { bg:'#ccfbf1', color:'#0f766e' },
  'angefragt':                          { bg:'#f0f9ff', color:'#0369a1' },
  'in bearbeitung':                     { bg:'#fffbeb', color:'#b45309' },
  'offen':                              { bg:'#eff6ff', color:'#1d4ed8' },
};

const SYSTEM_FIELDS = new Set([
  '@odata.etag','@odata.id','id','ContentType','Modified','Created',
  'AuthorLookupId','EditorLookupId','Attachments','Edit','LinkTitleNoMenu','LinkTitle',
  'ItemChildCount','FolderChildCount','_UIVersionString','ComplianceAssetId',
  'OData__ColorTag','AppAuthorLookupId','AppEditorLookupId'
]);

// ── ARTIKEL-STAMMDATEN (TID → Bezeichnung / Warengruppe) ─────────────────────
const TID_MAP = {
};


// ── KOSTENSTELLEN-STAMMDATEN ────────────────────────────────────────────────
// Format: { nr: "10011", label: "Fertigungsleitung" }
// Overrideable via localStorage key DIHAG_KOSTENST
const KOSTENST_DEFAULT = [
];

// Active lookup tables — reloaded from localStorage by reloadLookupData()
let TID_MAP_ACTIVE = TID_MAP;         // may be replaced by imported data
let TID_ENTRIES    = Object.entries(TID_MAP);
let KOSTENST_DATA  = KOSTENST_DEFAULT; // may be replaced by imported data

const LS_TID_KEY    = 'DIHAG_TID_DATA';
const LS_KOSTENST_KEY = 'DIHAG_KOSTENST';

function reloadLookupData() {
  try {
    const rawTid = localStorage.getItem(LS_TID_KEY);
    if (rawTid) {
      const obj = JSON.parse(rawTid);
      TID_MAP_ACTIVE = Object.assign({}, TID_MAP, obj); // merge: imported wins over default for same keys
      TID_ENTRIES = Object.entries(TID_MAP_ACTIVE);
    } else {
      TID_MAP_ACTIVE = TID_MAP;
      TID_ENTRIES = Object.entries(TID_MAP);
    }
  } catch(e) { console.warn('reloadLookupData TID:', e); }
  try {
    const rawKst = localStorage.getItem(LS_KOSTENST_KEY);
    if (rawKst) KOSTENST_DATA = JSON.parse(rawKst);
    else        KOSTENST_DATA = KOSTENST_DEFAULT;
  } catch(e) { console.warn('reloadLookupData KOSTENST:', e); }
}

// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp, account;
let siteId = null, listId = null;
let allItems = [];
let colByKey   = {};  // internal name → column definition (from SP)
let spUserMap  = {};  // SP user id (string) → display name (for Person-column LookupId resolution)
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let statusChoices = []; // All SP Status column choices in order (populated in discoverSP)

// ── EDIT LOCK ─────────────────────────────────────────────────────────────────
// Stores a "user|ISO-expires" value in a SP list field to prevent concurrent edits.
const LOCK_FIELD_CANDIDATES = ['BearbeitungsSperre', 'EditLock', 'Bearbeitungssperre'];
let _lockField  = undefined; // undefined = not yet resolved; '' = field not in SP
let lockTimers  = {};        // itemId (string) → timeout handle
let currentView = 'dashboard';
let prevView    = 'dashboard';
let wizardData  = {};
let wizardFilesArr = []; // Drag-and-drop file store for wizard attachments
// Einzelanfrage: mehrere (Menge × Einzelpreis)-Zeilen; Bestellvolumen = Σ(Menge×EP)
let einzelLines = [{ stueck: '', preis: '' }];
const _ezNum = v => parseFloat(String(v ?? '').replace(',', '.')) || 0;
let panelItemId = null;
let panelEditMode = false; // true, solange ein Panel im Bearbeiten-Modus offen ist (pausiert Auto-Refresh)
// SP column may be misspelled "Stauts" in some tenants → always try both
const getStatusVal = item => getField(item,'Status') || getField(item,'Stauts') || '';

// ── USER SETTINGS ────────────────────────────────────────────────────────────
const SETTINGS_KEY = 'bedarfsanfrage_settings_v1';
const ADMIN_EMAIL  = 'administrator@dihag.com';

function getSettings(email) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  return Object.assign({
    autoRefresh:        true,
    autoRefreshGranted: true,
    pageSize:           100,
    compactView:        false,   // dense single-line cards in list views
    hideCompleted:      false,   // hide bestellt/erledigt/abgelehnt in Meine Anfragen
    defaultSort:        'date-desc', // default sort for Meine Anfragen
    canSeeDashboard:    false,   // admin-granted: show Dashboard tab (all items visible)
  }, all[(email||'').toLowerCase()] || {});
}
// Persönliche Einstellungen (pageSize, compactView, Sortierung …) liegen lokal
// im Browser. Zentrale, geräteübergreifende Steuerung läuft über die
// Zugriffsverwaltung (BANF_Konfiguration), nicht hier.
function saveUserSettings(email, patch) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  const key = (email||'').toLowerCase();
  all[key] = Object.assign(getSettings(email), patch);
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(all));
  return all[key];
}

// ── AUTO-REFRESH ─────────────────────────────────────────────────────────────
// autoRefreshTimer: 1-second tick; arCountdown: seconds until next refresh
let autoRefreshTimer = null;
let arCountdown = 30;
let arPaused = false;   // user can pause without admin losing the feature-enable
let arPausedByWizard = false; // true, wenn die Pause nur durch einen offenen Wizard entstand

function startAutoRefresh() {
  stopAutoRefresh();
  if (!account) return;
  arCountdown = 20;
  arPaused    = false;
  autoRefreshTimer = setInterval(() => {
    if (arPaused) return;
    // Tab nicht sichtbar → kein API-Aufruf (spart Graph-Last/Throttling)
    if (document.hidden) return;
    // Panel im Bearbeiten-Modus offen → Refresh aussetzen (Eingaben nicht überschreiben)
    if (panelEditMode) return;
    arCountdown--;
    if (arCountdown <= 0) {
      arCountdown = 20;
      _autoRefreshData();    // inkrementell (bzw. periodisch voll)
      refreshAccessConfig(); // Zugriffsfreigaben automatisch übernehmen (~20s)
    }
    updateARBtn();
  }, 1000);
  updateARBtn();
}
function stopAutoRefresh() {
  if (autoRefreshTimer) { clearInterval(autoRefreshTimer); autoRefreshTimer = null; }
  updateARBtn();
}
// Beim Zurückkehren auf den Tab sofort einmal aktualisieren und Countdown neu starten
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && autoRefreshTimer && !arPaused && !panelEditMode) {
    arCountdown = 20;
    loadItems(false);
    refreshAccessConfig(); // Freigaben beim Zurückkehren sofort übernehmen
    updateARBtn();
  }
});
function updateARBtn() {
  const btn = $id('btn-autorefresh');
  if (!btn) return;
  const running = !!autoRefreshTimer && !arPaused;
  btn.classList.toggle('ar-on', running);
  if (running) {
    btn.title       = `Auto-Aktualisierung AN – nächste in ${arCountdown}s – klicken zum Pausieren`;
    btn.textContent = `⏱ ${arCountdown}s`;
  } else {
    btn.title       = 'Auto-Aktualisierung pausiert – klicken zum Fortsetzen';
    btn.textContent = '⏱ Pause';
  }
}
function toggleAutoRefresh() {
  if (!account) return;
  if (!autoRefreshTimer) { startAutoRefresh(); return; }
  arPaused = !arPaused;
  arPausedByWizard = false; // manuelle Steuerung hat Vorrang
  updateARBtn();
}

// ── APPROVER DISPLAY ─────────────────────────────────────────────────────────
// Confirmed SP internal name: Entscheider_x002a_in (display: "Genehmiger")
// Also checks colByKey display names for future-proofing.
const APPROVER_DIRECT_KEYS = ['Entscheider_x002a_in', 'Genehmiger', 'CurrentApprover'];
const APPROVER_COL_RE = /\bgenehmiger\b|\bentscheider\b|\bbearbeiter\b|aktuelle[rs]?\s*(bearbeiter|zugewiesen|genehmiger)|current.?approver/i;

function getApproverVal(item) {
  // Helper: extract a printable string from a field value (handles Person/Lookup objects).
  function extractStr(raw) {
    if (raw == null) return null;
    if (typeof raw === 'object') {
      const name = raw.displayName || raw.LookupValue || raw.Title || raw.title || raw.text;
      return (name && name !== '[object Object]') ? String(name).trim() || null : null;
    }
    const s = String(raw).trim();
    return (s && s !== '[object Object]') ? s : null;
  }

  if (!item?.fields) return null;

  // 1. Direct known keys (fastest path — confirmed SP internal names)
  for (const k of APPROVER_DIRECT_KEYS) {
    // Text/object value
    const v = extractStr(item.fields[k]);
    if (v) return v;
    // Person-column LookupId variant (e.g. "Entscheider_x002a_inLookupId" = 42)
    const lid = item.fields[k + 'LookupId'];
    if (lid != null) {
      const name = spUserMap[String(lid)];
      if (name) return name;
    }
  }

  // 2. Match via colByKey display names (catches renamed/future columns)
  for (const [k, c] of Object.entries(colByKey)) {
    if (SYSTEM_FIELDS.has(k)) continue;
    if (APPROVER_COL_RE.test(c.displayName || k)) {
      const v = extractStr(getField(item, k));
      if (v) return v;
      // Also try LookupId variant for Person columns
      const lid = item.fields[k + 'LookupId'];
      if (lid != null) {
        const name = spUserMap[String(lid)];
        if (name) return name;
      }
    }
  }

  // 3. Scan item.fields keys directly as last resort
  for (const [k, raw] of Object.entries(item.fields)) {
    if (SYSTEM_FIELDS.has(k)) continue;
    if (APPROVER_COL_RE.test(k)) {
      const v = extractStr(raw);
      if (v) return v;
    }
  }

  return null;
}

// ── AUTH ────────────────────────────────────────────────────────────────────
async function initAuth() {
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth: { clientId:CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation:'localStorage', storeAuthStateInCookie:true }
  });
  await msalApp.initialize();
  const r = await msalApp.handleRedirectPromise();
  if (r) { account = r.account; msalApp.setActiveAccount(r.account); return true; }
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) { account = accounts[0]; msalApp.setActiveAccount(accounts[0]); return true; }
  return false;
}

async function doLogin() {
  $id('boot-btn').style.display = 'none';
  $id('boot-sub').textContent = 'Weiterleitung zur Anmeldung…';
  $id('boot-spinner').style.display = 'block';
  try {
    await msalApp.loginRedirect({ scopes: SCOPES });
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
  }
}

async function bootDone() {
  $id('boot-sub').textContent = 'Daten werden geladen…';
  try {
    await discoverSP();
    await loadItems(false);
    // Echte Identitäten (UPN/Mail/Aliase) laden – Basis für die Zugriffsprüfung
    await loadMyIdentities();
    // Zentrale Zugriffssteuerung laden (bestimmt Sichtbarkeit von Dashboard/Reports/Importer)
    await loadAccessConfig();
    // Wiedervorlagen aus SharePoint laden (nicht blockierend für den Start)
    initFavorites().catch(() => {});
    // Hinweis: Automatische Entwürfe werden jetzt vom geplanten GitHub-Actions-Cron
    // (scripts/auto-drafts.mjs) angelegt – die App-seitige Erstellung ist deaktiviert,
    // um Doppelanlagen zu vermeiden.
    // processAutoDrafts().catch(() => {});
    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';
    applyNavVisibility(); // Nav-Sichtbarkeit nach geladener Zugriffskonfiguration setzen
    startAutoRefresh();
    // Set user info in sidebar
    const name = account?.name || account?.username || '?';
    $id('hdr-av').textContent   = name.split(' ').map(p=>p[0]||'').join('').substring(0,2).toUpperCase();
    $id('hdr-name').textContent = name;
    $id('hdr-mail').textContent = account?.username || '';
    navigate('dashboard');
  } catch(e) {
    $id('boot-sub').textContent = 'Fehler beim Laden: ' + e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-err').textContent = e.message;
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
    $id('boot-btn').onclick = bootDone;
  }
}

function doLogout() {
  msalApp?.logoutPopup({ account }).catch(()=>{});
  location.reload();
}

async function getToken(forceRefresh = false) {
  if (!account) throw new Error('Nicht angemeldet');
  try { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account, forceRefresh})).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({scopes:SCOPES});
    }
    throw e;
  }
}

// Authentifizierter fetch mit Retry.
//  • 401 (abgelaufenes/ungültiges Token): einmal mit frisch erzwungenem Token wiederholen.
//  • 429/503/504 (Graph-Throttling bzw. kurzzeitige Serverfehler): die vom Dienst
//    vorgegebene Retry-After-Wartezeit beachten und begrenzt wiederholen. Ohne das
//    flog eine gedrosselte Anfrage (z. B. bei 20-s-Polling mehrerer Nutzer) als Fehler hoch.
async function _authedFetch(url, opts = {}, retries = { auth: 1, throttle: 2 }) {
  const tok = await getToken();
  const r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + tok } });
  if (r.status === 401 && retries.auth > 0) {
    try { await getToken(true); } catch {}
    return _authedFetch(url, opts, { ...retries, auth: retries.auth - 1 });
  }
  if ((r.status === 429 || r.status === 503 || r.status === 504) && retries.throttle > 0) {
    const ra = parseInt(r.headers.get('Retry-After') || '', 10);
    const waitMs = Math.min((Number.isFinite(ra) && ra > 0 ? ra : 2) * 1000, 30000);
    await new Promise(res => setTimeout(res, waitMs));
    return _authedFetch(url, opts, { ...retries, throttle: retries.throttle - 1 });
  }
  return r;
}

// SharePoint REST API needs a token for the SP resource (different from Graph)
const SP_SCOPES = [`https://${SP_SITE.split(':/')[0]}/Sites.ReadWrite.All`];
async function getSpToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try { return (await msalApp.acquireTokenSilent({ scopes: SP_SCOPES, account })).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({ scopes: SP_SCOPES });
    }
    throw e;
  }
}

// ── GRAPH API ────────────────────────────────────────────────────────────────
async function gGet(path) {
  const url = path.startsWith('http') ? path : API + path; // absolute URL = @odata.nextLink
  const r   = await _authedFetch(url, { headers: { 'Cache-Control': 'no-cache', Pragma: 'no-cache' } });
  if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
// Folgt @odata.nextLink und sammelt ALLE Seiten (mit Sicherheitslimit gegen Endlosschleifen).
async function gGetAll(path, maxPages = 100) {
  const out = [];
  let next = path, pages = 0;
  while (next && pages < maxPages) {
    pages++;
    const data = await gGet(next);
    if (Array.isArray(data.value)) out.push(...data.value);
    next = data['@odata.nextLink'] || null;
  }
  return out;
}
async function gPost(path, body) {
  const r = await _authedFetch(API + path, {
    method:'POST', headers:{ 'Content-Type':'application/json' }, body: JSON.stringify(body)
  });
  if (!r.ok) {
    const txt = await r.text().catch(()=>'');
    console.error('[gPost] error', r.status, txt);
    throw new Error(`Graph POST ${r.status}: ${txt}`);
  }
  return r.json();
}
// Retry helper: SP sometimes reports 404 immediately after item creation (propagation delay).
// Wait delayMs and retry up to `retries` times when itemNotFound.
async function retryOn404(fn, retries = 4, delayMs = 1200) {
  for (let i = 0; i <= retries; i++) {
    try { return await fn(); }
    catch(e) {
      const is404 = (e.message.includes('404') || e.message.includes('itemNotFound')) && !e._noRetry;
      if (is404 && i < retries) {
        console.warn(`[retryOn404] Versuch ${i+1} – warte ${delayMs}ms…`, e.message);
        await new Promise(r => setTimeout(r, delayMs));
        continue;
      }
      throw e;
    }
  }
}

async function gPatch(path, body) {
  const r = await _authedFetch(API + path, {
    method:'PATCH', headers:{ 'Content-Type':'application/json' }, body: JSON.stringify(body)
  });
  if (r.status === 204) return {};
  if (!r.ok) throw new Error(`Graph PATCH ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json().catch(()=>({}));
}

async function gDelete(path) {
  const r = await _authedFetch(API + path, { method:'DELETE' });
  if (!r.ok && r.status !== 404) throw new Error(`Graph DELETE ${r.status}: ${await r.text().catch(()=>'')}`);
  return {};
}

// ── SP DISCOVERY ─────────────────────────────────────────────────────────────
async function discoverSP() {
  // 1. Site ID
  const site = await gGet(`/sites/${SP_SITE}`);
  siteId = site.id;

  // 2. List ID
  const list = await gGet(`/sites/${siteId}/lists/${encodeURIComponent(SP_LIST)}`);
  listId = list.id;

  // 3. Columns
  const colsRes = await gGet(`/sites/${siteId}/lists/${listId}/columns`);
  colByKey = {};
  for (const c of colsRes.value) {
    colByKey[c.name] = c;  // by internal name
  }

  // 4. Resolve each form field to actual SP internal name
  resolvedFields = {};
  const missing = [];
  for (const fd of [...FORM_FIELDS, ...EINKAUF_FIELDS]) {
    const found = resolveColName(fd);
    resolvedFields[fd.key] = found;
    if (!found && fd.key !== 'Title') missing.push(fd.label);
  }
  // Title always exists
  resolvedFields['Title'] = 'Title';

  // Show warning for missing columns
  if (missing.length) {
    const w = $id('col-warning');
    w.innerHTML = `<strong>⚠ Folgende Spalten fehlen noch in der SharePoint-Liste</strong> – bitte in der SP-Liste anlegen, damit alle Felder gespeichert werden können:<br>
      <em>${missing.join(' · ')}</em>`;
    w.style.display = 'block';
  }

  // Populate Choice selects from actual SP column definitions
  for (const fd of FORM_FIELDS) {
    const colName = resolvedFields[fd.key];
    if (!colName) continue;
    const col = colByKey[colName];
    if (!col?.choice?.choices?.length) continue;
    const el = document.getElementById('f-' + fd.key);
    if (!el || el.tagName !== 'SELECT') continue;
    el.innerHTML = fd.required ? '' : '<option value="">– bitte wählen –</option>';
    for (const c of col.choice.choices) el.add(new Option(c, c));
  }

  // ── Status-Spalte mit SP-Liste synchronisieren ────────────────────────────
  // Findet die Status-Spalte (auch "Stauts" als Tippfehler-Variante), liest die
  // Auswahlwerte und baut STATUS_STYLES sowie WORKFLOW_STAGES automatisch auf.
  const statusCol = Object.values(colByKey).find(c =>
    /^status$/i.test(c.name) || /^stauts$/i.test(c.name)
  );
  // Store the real internal name so all readers (loadVersionHistory etc.) use the correct key
  if (statusCol) resolvedFields['Status'] = statusCol.name;

  if (statusCol?.choice?.choices?.length) {
    const choices = statusCol.choice.choices;
    statusChoices = choices; // store all choices (incl. end states) for the timeline
    // Rebuild STATUS_STYLES
    for (const v of choices) {
      const key = v.toLowerCase().trim();
      if (!STATUS_STYLES[key]) STATUS_STYLES[key] = statusColorFor(v);
    }
    // Rebuild WORKFLOW_STAGES: preserve order from SP list, skip "Abgelehnt"/"Erledigt" (end states)
    const endStates = /abgelehnt|erledigt|abgebrochen/i;
    const stages = choices.filter(c => !endStates.test(c));
    if (stages.length) {
      WORKFLOW_STAGES.length = 0;
      WORKFLOW_STAGES.push(...stages);
    }
  }

  // 5. Build SP user map (id → display name) for Person-column LookupId resolution.
  //    Person columns via Graph API return only a numeric LookupId; we need this map
  //    to show the approver's name instead of a raw ID.
  try {
    const usersRes = await gGet(
      `/sites/${siteId}/lists('User Information List')/items?$expand=fields($select=Title,EMail,Name)&$top=500`
    );
    spUserMap = {};
    for (const u of (usersRes.value || [])) {
      const name = u.fields?.Title || u.fields?.Name || u.fields?.EMail;
      if (name) spUserMap[String(u.id)] = name;
    }
  } catch(e) {
    console.warn('[discoverSP] Could not load User Information List:', e.message);
  }
}

// Decode SP hex-encoded field names like "Gesch_x00e4_tzter" → "Geschätzter"
function decodeSpFieldName(name) {
  return (name || '').replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

function resolveColName(fd) {
  if (fd.key === 'Title') return 'Title';
  // Only consider writable, non-sealed columns
  const ok = c => c && !c.readOnly && !c.sealed;
  if (ok(colByKey[fd.key])) return fd.key;
  for (const alt of (fd.alsoTry || [])) {
    if (ok(colByKey[alt])) return alt;
  }
  // Case-insensitive display-name match (writable only)
  const labelNorm = (fd.label || '').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
  for (const [k, c] of Object.entries(colByKey)) {
    if (!ok(c)) continue;
    const dn = (c.displayName||'').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
    if (dn === labelNorm) return k;
  }
  // Fuzzy hex-decode fallback: decode SP internal name and compare normalised
  // e.g. "Gesch_x00e4_tzterPreisnetto_x002" decoded ≈ "GeschätzterPreisnetto..."
  const keyNorm = fd.key.toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
  for (const [k, c] of Object.entries(colByKey)) {
    if (!ok(c)) continue;
    const decoded = decodeSpFieldName(k).toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
    if (decoded === keyNorm) return k;
    // Prefix match (SP truncates names at 32 chars)
    if (decoded.length >= 10 && (decoded.startsWith(keyNorm) || keyNorm.startsWith(decoded))) return k;
  }
  return null;
}

// ── LOAD ITEMS ───────────────────────────────────────────────────────────────
async function loadItems(showToast = true) {
  const btn = $id('btn-reload');
  if (btn) btn.disabled = true;
  try {
    // Vollständig laden: allen Seiten folgen (Graph max. 200/Seite), nicht nur die ersten N
    const ps  = account ? (getSettings(account.username).pageSize || 100) : 100;
    const top = Math.min(Math.max(ps, 50), 200);
    allItems = await gGetAll(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*)&$top=${top}&$orderby=createdDateTime desc`
    );
    if (showToast) toast('Daten aktualisiert', 'success');
    _rerenderAfterData();
  } catch(e) {
    toast('Fehler beim Laden: ' + e.message, 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Aktuelle Ansicht (+ offenes Panel) nach Datenänderung neu rendern
function _rerenderAfterData() {
  if (currentView === 'dashboard')    renderDashboard();
  else if (currentView === 'mine')    renderList('mine');
  else if (currentView === 'all')     renderList('all');
  else if (currentView === 'reports') renderReportsTable();
  // Offenes Panel aktualisieren – aber NICHT, solange bearbeitet wird
  if (panelItemId && !panelEditMode && ['mine','all','dashboard'].includes(currentView)) {
    const pi = allItems.find(i => String(i.id) === panelItemId);
    if (pi) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(pi); bindPanelEvents(panelItemId); }
  }
}

// Inkrementelle Aktualisierung für den Auto-Refresh: holt nur die zuletzt
// geänderten Einträge (kleine Seite) und merged sie – statt der ganzen Liste.
// Löschungen fängt der periodische Voll-Reload (siehe _autoRefreshData) + Tabwechsel.
async function refreshItemsIncremental() {
  if (!siteId || !listId) return;
  try {
    const page = await gGet(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*)&$top=50&$orderby=lastModifiedDateTime desc`
    );
    const recent = page.value || [];
    if (!recent.length) return;
    const byId = new Map(allItems.map(i => [String(i.id), i]));
    let changed = false;
    for (const it of recent) {
      const id = String(it.id);
      const ex = byId.get(id);
      if (!ex || ex.lastModifiedDateTime !== it.lastModifiedDateTime) { byId.set(id, it); changed = true; }
    }
    if (changed) { allItems = [...byId.values()]; _rerenderAfterData(); }
  } catch {
    // z. B. orderby nicht unterstützt → sicherer Voll-Reload
    try { await loadItems(false); } catch {}
  }
}

// Auto-Refresh-Datenstrategie: meist inkrementell, ~alle 5 min voll (fängt Löschungen)
let _arCycles = 0;
function _autoRefreshData() {
  _arCycles++;
  if (_arCycles % 15 === 0) loadItems(false);
  else                      refreshItemsIncremental();
}

// ── AUTOMATISCHE ENTWÜRFE (Erinnerung) ───────────────────────────────────────
// Erstellt zu fälligen „Datum zur automatischen Bedarfsanfrage" eine Kopie als
// Entwurf (Status „Automatisch erstellter Entwurf"). Läuft NUR als Admin → eine
// Sitzung, keine Doppelanlage. Die Quelle wird nach dem Auslösen geleert, damit
// sie nicht erneut feuert.
function _buildAutoDraftFields(src) {
  const out = {};
  for (const fd of FORM_FIELDS) {
    if (fd.key === 'AutoBedarfDatum') continue;          // Erinnerungsdatum nicht kopieren
    const col = resolvedFields[fd.key];
    if (!col) continue;
    const v = getField(src, col);
    if (v == null || v === '' || typeof v === 'object') continue; // Personen/Arrays/Objekte überspringen
    out[col] = v;
  }
  const srcTitle = cleanTitle(getField(src,'Title') || '') || 'Bedarfsanfrage';
  out['Title'] = 'Auto-Entwurf: ' + srcTitle;
  return out;
}

async function processAutoDrafts() {
  if (!isAdmin()) return;                                // nur Admin erzeugt Auto-Entwürfe
  const dateCol = resolvedFields['AutoBedarfDatum'];
  if (!dateCol) return;                                  // Spalte fehlt → nichts zu tun
  const statusCol = resolvedFields['Status'] || 'Status';
  const draftStatus = statusChoices.find(c => /automatisch erstellt/i.test(c))
                   || statusChoices.find(c => /entwurf/i.test(c))
                   || 'Automatisch erstellter Entwurf';
  const todayEnd = new Date(); todayEnd.setHours(23,59,59,999);
  const due = allItems.filter(i => {
    const d = getField(i, dateCol);
    if (!d) return false;
    const dd = new Date(d);
    if (isNaN(dd.getTime()) || dd > todayEnd) return false;
    const st = (getStatusVal(i) || '').toLowerCase();
    return !/entwurf|automatisch erstellt/.test(st);     // Entwürfe nicht erneut kopieren
  });
  if (!due.length) return;

  const listPath = `/sites/${siteId}/lists/${listId}`;
  let created = 0;
  for (const src of due) {
    try {
      // 1) Quelle: Erinnerungsdatum leeren (verhindert erneutes Auslösen)
      await gPatch(`${listPath}/items/${src.id}/fields`, { [dateCol]: null });
      const sc = allItems.find(i => String(i.id) === String(src.id));
      if (sc?.fields) sc.fields[dateCol] = null;
      // 2) Entwurf als Kopie anlegen (robust: unbekannte Felder werden verworfen)
      const fields  = _buildAutoDraftFields(src);
      const newItem = await gPost(`${listPath}/items`, { fields: { Title: fields.Title } });
      const rest    = { ...fields }; delete rest.Title;
      rest[statusCol] = draftStatus;
      await patchRetry(`${listPath}/items/${newItem.id}/fields`, rest);
      created++;
    } catch(e) { console.warn('[AutoDraft] #' + src.id + ' fehlgeschlagen:', e.message); }
  }
  if (created) {
    toast(`${created} automatische${created === 1 ? 'r' : ''} Entwurf${created === 1 ? '' : 'e'} erstellt.`, 'success');
    await loadItems(false);
  }
}

// ── ROUTING ──────────────────────────────────────────────────────────────────
// Dashboard title differs by role: admins see "Dashboard (Alle)", others see "Meine Anfragen"
function VIEW_TITLES(view) {
  const map = { new:'Neue Bedarfsanfrage', multi:'Sammelanfrage', mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details', importer:'Tabellen-Importer', reports:'Reports & Auswertungen', schulung:'Schulung (Beta)' };
  if (view === 'dashboard') return canAccess('dashboard') ? 'Dashboard (Alle Anfragen)' : 'Meine Anfragen';
  return map[view] || view;
}

function isAdmin() { return account?.username?.toLowerCase() === ADMIN_EMAIL; }

// ── ZENTRALE ZUGRIFFSSTEUERUNG (nur vom Admin pflegbar, in SharePoint) ─────────
// Gated-Bereiche: Dashboard, Reports, Tabellen-Importer. Standard: nur Admin sieht
// sie; der Admin gibt einzelne UPNs (z. B. Einkauf) zentral frei.
// Speicherung: dedizierte SP-Liste 'BANF_Konfiguration'. Ein Element identifiziert
// über Title='access' hält die gesamte Konfiguration als JSON in der Spalte
// 'ConfigValue' (mehrzeiliger Text). Die Liste wird manuell angelegt (Listen-
// erstellung per Graph ist im Tenant gesperrt) – die App findet sie nur.
const ACCESS_LIST_NAME  = 'BANF_Konfiguration';
const ACCESS_ITEM_TITLE = 'access';
const GATED_FEATURES    = ['dashboard', 'reports', 'importer'];
let accessListId        = null;
let accessConfigItemId  = null;
let accessUsers         = {}; // { 'upn@dihag.com': { dashboard:true, reports:true, importer:false } }

function myUPN() { return (account?.username || '').trim().toLowerCase(); }

// Identitäten aus Graph /me (UPN, Mail, Alias-/Proxy-Adressen) – beim Start geladen.
// Entscheidend bei Membern mit Zweitdomäne: dort ist der Login-UPN oft ≠ E-Mail,
// und der E-Mail-Claim fehlt im Token. /me liefert die echten Adressen zuverlässig.
let _meIds = null;
async function loadMyIdentities() {
  _meIds = new Set();
  try {
    const me = await gGet('/me?$select=userPrincipalName,mail,otherMails,proxyAddresses');
    const add = v => { if (v) _meIds.add(String(v).replace(/^smtp:/i, '').trim().toLowerCase()); };
    add(me.userPrincipalName); add(me.mail);
    (me.otherMails || []).forEach(add);
    (me.proxyAddresses || []).forEach(add);
  } catch (e) { console.warn('[ident] /me konnte nicht geladen werden:', e.message); }
}

// Alle Identitäts-Bezeichner des aktuellen Nutzers (kleingeschrieben):
// Graph /me-Adressen + angemeldeter „username" + Token-Claims.
function _myIdentities() {
  const ids = new Set(_meIds || []);
  const add = v => { if (v) ids.add(String(v).trim().toLowerCase()); };
  add(account?.username);
  const c = account?.idTokenClaims || {};
  add(c.email); add(c.preferred_username); add(c.upn);
  if (Array.isArray(c.emails)) c.emails.forEach(add);
  return ids;
}

function canAccess(feature) {
  if (isAdmin()) return true;
  for (const id of _myIdentities()) {
    const u = accessUsers[id];
    if (u && u[feature]) return true;
  }
  return false;
}
function canSeeDashboard() { return canAccess('dashboard'); }
// Einkäufer: darf beim Öffnen den Status auf „In Prüfung (Einkauf)" setzen und
// Anfragen genehmigen/ablehnen. Admin immer; sonst per Zugriffsverwaltung (einkauf).
function canApprove() { return isAdmin() || canAccess('einkauf'); }

// Rich-Text-Spalten liefern HTML zurück → Tags entfernen + Entities dekodieren,
// damit das JSON unabhängig vom Spaltentyp (einfacher Text / Rich-Text) parsebar ist.
function _decodeSpText(s) {
  if (s == null) return '';
  const noTags = String(s).replace(/<[^>]*>/g, '');
  const ta = document.createElement('textarea');
  ta.innerHTML = noTags;
  return ta.value;
}

async function _findConfigList() {
  if (accessListId) return accessListId;
  try {
    const all = await gGet(`/sites/${siteId}/lists?$select=id,displayName&$top=500`);
    accessListId = (all.value || []).find(l =>
      (l.displayName || '').trim().toLowerCase() === ACCESS_LIST_NAME.toLowerCase())?.id || null;
  } catch { accessListId = null; }
  return accessListId;
}

// Beim Start aufrufen (nach discoverSP). Niemals werfen.
async function loadAccessConfig() {
  accessUsers = {}; accessConfigItemId = null; accessListId = null;
  try {
    if (!siteId) return;
    const listId = await _findConfigList();
    if (!listId) return; // Liste fehlt → nur Admin sieht gated Bereiche
    const res  = await gGet(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`);
    const item = (res.value || []).find(it => (it.fields?.Title || '') === ACCESS_ITEM_TITLE);
    if (item) {
      accessConfigItemId = item.id;
      try { accessUsers = JSON.parse(_decodeSpText(item.fields?.ConfigValue) || '{}').users || {}; } catch {}
    }
  } catch (e) {
    console.warn('[Zugriff] Konfiguration konnte nicht geladen werden:', e.message);
  }
}

// Leichte periodische Aktualisierung (an Auto-Refresh gekoppelt): liest die
// Konfiguration neu und wendet Änderungen ohne manuelles Neuladen an. Nutzt die
// gecachte Listen-ID; bei fehlendem Lesezugriff still wirkungslos.
async function refreshAccessConfig() {
  if (!account || isAdmin()) return; // Admin sieht ohnehin alles
  try {
    const listId = await _findConfigList();
    if (!listId) return;
    const res  = await gGet(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`);
    const item = (res.value || []).find(it => (it.fields?.Title || '') === ACCESS_ITEM_TITLE);
    let next = {};
    if (item) {
      accessConfigItemId = item.id;
      try { next = JSON.parse(_decodeSpText(item.fields?.ConfigValue) || '{}').users || {}; } catch {}
    }
    if (JSON.stringify(next) !== JSON.stringify(accessUsers)) {
      accessUsers = next;
      applyNavVisibility();
      // Aktueller Bereich nicht mehr erlaubt → auf „Meine Anfragen" umleiten
      if (GATED_FEATURES.includes(currentView) && !canAccess(currentView)) navigate('mine');
    }
  } catch { /* still: kein Lesezugriff o. Ä. */ }
}
async function saveAccessConfig() {
  if (!isAdmin()) { console.warn('[Zugriff] Speichern abgelehnt – kein Admin.'); return; }
  try {
    const listId = await _findConfigList();
    if (!listId) { toast("Liste 'BANF_Konfiguration' nicht gefunden – siehe Konsole.", 'error'); return; }
    const fields = { Title: ACCESS_ITEM_TITLE, ConfigValue: JSON.stringify({ users: accessUsers }) };
    if (accessConfigItemId) {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${accessConfigItemId}/fields`, fields);
    } else {
      const created = await gPost(`/sites/${siteId}/lists/${listId}/items`, { fields });
      accessConfigItemId = created.id;
    }
    toast('Zugriffsrechte gespeichert ✓', 'success');
  } catch (e) {
    const hint = /ConfigValue/i.test(e.message)
      ? " – Spalte 'ConfigValue' (Mehrere Textzeilen) in der Liste anlegen."
      : '';
    toast('Speichern der Zugriffsrechte fehlgeschlagen: ' + e.message + hint, 'error');
  }
}
// Admin-UI-Helfer
function setAccess(upn, feature, on) {
  upn = (upn || '').trim().toLowerCase();
  if (!upn) return;
  if (!accessUsers[upn]) accessUsers[upn] = { dashboard: false, reports: false, importer: false, einkauf: false };
  accessUsers[upn][feature] = !!on;
  saveAccessConfig();
  applyNavVisibility();
}
function addAccessUser() {
  const inp = $id('access-new-upn');
  const upn = (inp?.value || '').trim().toLowerCase();
  if (!upn || !/@/.test(upn)) { toast('Bitte gültige UPN/E-Mail angeben.', 'error'); return; }
  if (!accessUsers[upn]) accessUsers[upn] = { dashboard: true, reports: false, importer: false, einkauf: true };
  saveAccessConfig();
  openSettings();
}
function removeAccessUser(upn) {
  delete accessUsers[(upn || '').trim().toLowerCase()];
  saveAccessConfig();
  applyNavVisibility();
  openSettings();
}

// Nav-Items nach zentraler Zugriffssteuerung ein-/ausblenden.
function applyNavVisibility() {
  const setVis = (view, vis) => {
    const el = document.querySelector(`.nav-item[data-view="${view}"]`);
    if (el) el.style.display = vis ? '' : 'none';
  };
  GATED_FEATURES.forEach(f => setVis(f, canAccess(f)));
  // new, multi, mine, schulung sind für alle sichtbar
}

function navigate(view, id) {
  // Gated-Bereiche nur mit Zugriff betreten; sonst auf „Meine Anfragen" umleiten
  if (GATED_FEATURES.includes(view) && !canAccess(view)) view = 'mine';
  // Always close panels of all split views when navigating
  ['mine','all','dashboard'].forEach(v => {
    $id('panel-' + v)?.classList.add('hidden');
    $id('split-' + v)?.classList.remove('has-panel');
  });
  panelItemId = null;
  panelEditMode = false;

  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item[data-view]').forEach(n => n.classList.remove('active'));

  const el = $id('view-' + view);
  if (!el) return;
  el.classList.add('active');

  const nav = document.querySelector(`.nav-item[data-view="${view}"]`);
  if (nav) nav.classList.add('active');

  $id('page-title').textContent = VIEW_TITLES(view);
  prevView = currentView;
  currentView = view;

  // Pause auto-refresh while filling in a form; resume when leaving
  if (view === 'new' || view === 'multi') {
    // Nur pausieren, wenn nicht ohnehin (manuell) pausiert; die Pause als wizard-bedingt merken.
    if (!arPaused) { arPaused = true; arPausedByWizard = true; updateARBtn(); }
  } else if (arPausedByWizard) {
    // Eine wizard-bedingte Pause beim Verlassen wieder aufheben; eine manuelle bleibt bestehen.
    arPausedByWizard = false; arPaused = false;
    if (!autoRefreshTimer) startAutoRefresh();
    updateARBtn();
  }

  if      (view === 'dashboard') renderDashboard();
  else if (view === 'mine')      renderList('mine');
  else if (view === 'all')       renderList('all');
  else if (view === 'new')       initWizard();
  else if (view === 'multi')     initMultiWizard();
  else if (view === 'importer')  initImporter();
  else if (view === 'reports')   initReports();
  else if (view === 'schulung')  initSchulung();
  else if (view === 'detail' && id) renderDetail(id);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Nav clicks
  document.querySelectorAll('.nav-item[data-view]').forEach(a => {
    a.addEventListener('click', e => { e.preventDefault(); navigate(a.dataset.view); });
  });
  // Dashboard "Neue Anfrage" button
  document.querySelectorAll('[data-view="new"]').forEach(b => {
    b.addEventListener('click', e => { e.preventDefault(); navigate('new'); });
  });
  // Detail back button
  $id('detail-back').addEventListener('click', () => navigate(prevView));
  // Menu toggle
  $id('menu-toggle').addEventListener('click', () => $id('sidebar').classList.toggle('open'));

  // Boot — load persisted lookup table overrides (TID + Kostenstellen)
  reloadLookupData();

  $id('boot-spinner').style.display = 'block';
  try {
    const loggedIn = await initAuth();
    if (loggedIn) await bootDone();
    else {
      $id('boot-sub').textContent = 'Bitte melden Sie sich an.';
      $id('boot-spinner').style.display = 'none';
      $id('boot-btn').style.display = 'block';
    }
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
  }
});

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
let dashStatusFilter = '';
let dashSortOrder   = 'date-desc';
let dashWGFilter = '';
let dashMAFilter = '';

function renderDashboard() {
  renderStatusChips();
  filterDashboard();
}

function setDashSort(val) {
  dashSortOrder = val;
  filterDashboard();
}

function renderStatusChips() {
  const el = $id('status-chips');
  if (!el) return;

  const baseItems = allItems;

  // Count per status
  const counts = {};
  let total = 0;
  for (const item of baseItems) {
    const s = getStatusVal(item) || '';
    counts[s] = (counts[s] || 0) + 1;
    total++;
  }

  // Volume from filtered items only
  const filteredForVol = dashStatusFilter
    ? baseItems.filter(i => (getStatusVal(i) || '') === dashStatusFilter)
    : baseItems;
  const volume = filteredForVol.reduce((sum, i) =>
    sum + (parseFloat(getField(i, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0), 0);

  // Build chips: "Alle" + one per status
  const allChip = `<button class="sc-chip${dashStatusFilter==='' ? ' active' : ''}" onclick="setDashFilter('')">
    <span class="sc-label">Alle</span>
    <span class="sc-count">${total}</span>
  </button>`;

  const statusChips = Object.entries(counts).sort((a,b) => b[1]-a[1]).map(([s, n]) => {
    const sl = s.toLowerCase().trim();
    let st = STATUS_STYLES[sl];
    if (!st) {
      for (const [k, style] of Object.entries(STATUS_STYLES)) {
        if (sl.includes(k)) { st = style; break; }
      }
    }
    st = st || { bg:'#f3f4f6', color:'#374151' };
    let icon = '';
    if      (sl.includes('abgelehnt') || sl.includes('rejected'))                icon = '✗';
    else if (sl.includes('freigegeben') || sl.includes('bestellt') || sl.includes('erledigt')) icon = '✓';
    else if (sl.includes('prüfung') || sl.includes('bearbeitung'))                icon = '⏳';
    else if (sl.includes('eingereicht') || sl.includes('angefragt'))              icon = '📋';
    else if (!s)                                                                   icon = '📋';
    const isActive = dashStatusFilter === s;
    const labelText = s || 'Eingereicht';
    return `<button class="sc-chip${isActive ? ' active' : ''}"
      style="--sc-bg:${st.bg};--sc-color:${st.color}"
      onclick="setDashFilter('${escJs(s)}')">
      ${icon ? `<span class="sc-icon">${icon}</span>` : ''}
      <span class="sc-label">${esc(labelText)}</span>
      <span class="sc-count" style="background:${st.bg};color:${st.color}">${n}</span>
    </button>`;
  }).join('');

  const vol = `<span class="sc-volume">${fmtEuro(volume)}<small>Volumen</small></span>`;
  el.innerHTML = allChip + statusChips + vol;
  renderExtraFilters();
}

function setDashFilter(status) {
  dashStatusFilter = (dashStatusFilter === status) ? '' : status;
  renderStatusChips();
  filterDashboard();
}

function renderExtraFilters() {
  const el = $id('dash-extra-filters');
  if (!el) return;

  // Unique WG values
  const wgField = resolvedFields['Warengruppe'] || 'Warengruppe';
  const wgSet = new Set();
  for (const i of allItems) {
    const v = getField(i, wgField) || '';
    if (v) wgSet.add(v);
  }
  const wgOpts = [...wgSet].sort().map(v =>
    `<option value="${esc(v)}"${dashWGFilter === v ? ' selected' : ''}>${esc(v)}</option>`
  ).join('');

  // Unique Mitarbeiter values
  const maSet = new Set();
  for (const i of allItems) {
    const v = i.createdBy?.user?.displayName || i.createdBy?.user?.email || '';
    if (v) maSet.add(v);
  }
  const maOpts = [...maSet].sort().map(v =>
    `<option value="${esc(v)}"${dashMAFilter === v ? ' selected' : ''}>${esc(v)}</option>`
  ).join('');

  el.innerHTML = `
    <select class="dash-filter-select" onchange="setDashWGFilter(this.value)">
      <option value="">Alle Warengruppen</option>
      ${wgOpts}
    </select>
    <select class="dash-filter-select" onchange="setDashMAFilter(this.value)">
      <option value="">Alle Mitarbeiter</option>
      ${maOpts}
    </select>`;
}

function setDashWGFilter(val) {
  dashWGFilter = val;
  filterDashboard();
}

function setDashMAFilter(val) {
  dashMAFilter = val;
  filterDashboard();
}

// Entprellte Suche: die Liste erst nach kurzer Tipp-Pause neu bauen, nicht bei jedem
// Tastendruck. Spart bei großen Listen das wiederholte Filtern, Sortieren und Neu-Rendern.
let _searchTimer = null;
function searchDebounced(fn) {
  clearTimeout(_searchTimer);
  _searchTimer = setTimeout(fn, 180);
}

function filterDashboard() {
  const search = ($id('search-dashboard')?.value || '').toLowerCase();
  let items = [...allItems];
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) || String(i.id||'').includes(search)
  );
  if (dashStatusFilter !== '') items = items.filter(i => (getStatusVal(i)||'') === dashStatusFilter);
  if (dashWGFilter) items = items.filter(i => (getField(i, resolvedFields['Warengruppe']||'Warengruppe')||'') === dashWGFilter);
  if (dashMAFilter) items = items.filter(i => {
    const creator = i.createdBy?.user?.displayName || i.createdBy?.user?.email || '';
    return creator === dashMAFilter;
  });

  const priceKey = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
  switch (dashSortOrder) {
    case 'date-asc':
      items.sort((a,b) => new Date(a.createdDateTime) - new Date(b.createdDateTime)); break;
    case 'price-desc':
      items.sort((a,b) => (parseFloat(getField(b,priceKey))||0) - (parseFloat(getField(a,priceKey))||0)); break;
    case 'price-asc':
      items.sort((a,b) => (parseFloat(getField(a,priceKey))||0) - (parseFloat(getField(b,priceKey))||0)); break;
    case 'status':
      items.sort((a,b) => (getStatusVal(a)||'').localeCompare(getStatusVal(b)||'')); break;
    default: // date-desc
      items.sort((a,b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
  }

  const container = $id('list-dashboard');
  if (container) container.innerHTML = items.length
    ? _cardsHtml(items)
    : emptyState('Keine Anfragen für diesen Status.');
}

// Render-Performance: nur die ersten LIST_ROW_CAP Karten darstellen
const LIST_ROW_CAP = 200;
function _cardsHtml(items) {
  const shown = items.slice(0, LIST_ROW_CAP);
  let html = shown.map(i => itemCard(i)).join('');
  if (items.length > LIST_ROW_CAP) {
    html += `<div class="list-cap-note">Erste ${LIST_ROW_CAP} von ${items.length} angezeigt – nutze Suche/Filter zum Eingrenzen.</div>`;
  }
  return html;
}

function isOpenStatus(s) {
  if (!s) return true;
  const sl = s.toLowerCase();
  return sl.includes('eingereicht') || sl.includes('prüfung') || sl.includes('offen');
}

const KANBAN_COLS = [
  { id:'eingereicht', label:'Eingereicht',  icon:'📋', match: s => !s || /eingereicht|angefragt|offen/.test(s.toLowerCase()),  bg:'#eff6ff', border:'#bfdbfe', dot:'#1d4ed8' },
  { id:'pruefung',    label:'In Prüfung',   icon:'⏳', match: s => /prüfung|bearbeitung/.test((s||'').toLowerCase()),           bg:'#fffbeb', border:'#fde68a', dot:'#b45309' },
  { id:'freigegeben', label:'Freigegeben',  icon:'✓',  match: s => /freigegeben/.test((s||'').toLowerCase()),                  bg:'#f0fdf4', border:'#bbf7d0', dot:'#15803d' },
  { id:'bestellt',    label:'Bestellt',     icon:'📦', match: s => /bestellt|erledigt/.test((s||'').toLowerCase()),             bg:'#faf5ff', border:'#e9d5ff', dot:'#7e22ce' },
  { id:'abgelehnt',   label:'Abgelehnt',   icon:'✗',  match: s => /abgelehnt|rejected/.test((s||'').toLowerCase()),            bg:'#fef2f2', border:'#fecaca', dot:'#b91c1c' },
];

function renderKanban(items) {
  if (!items.length) return `<div style="padding:24px 18px;color:#9ca3af;font-size:.88rem">Noch keine Anfragen vorhanden.</div>`;
  const cols = KANBAN_COLS.map(col => ({
    ...col,
    items: items.filter(i => col.match(getStatusVal(i)))
  }));
  return `<div class="kb-board">` +
    cols.map(col => {
      const cards = col.items.slice(0, 15).map(item => {
        const title  = getField(item,'Title') || '–';
        const preis  = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
        const wg     = getField(item, resolvedFields['Warengruppe'] || 'Warengruppe') || '';
        const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
        return `<div class="kb-mini" onclick="navigate('detail','${item.id}')">
          <div class="kb-mini-title">${esc(title)}</div>
          ${wg ? `<div class="kb-mini-wg">${esc(wg)}</div>` : ''}
          <div class="kb-mini-foot">#${item.id}${preis ? ` · ${fmtEuro(preis)}` : ''}${created ? ` · ${created}` : ''}</div>
        </div>`;
      }).join('');
      const moreCount = col.items.length - 15;
      return `
        <div class="kb-col" style="--kb-border:${col.border};--kb-bg:${col.bg}">
          <div class="kb-col-hdr">
            <span class="kb-col-icon">${col.icon}</span>
            <span class="kb-col-label">${col.label}</span>
            <span class="kb-col-count" style="background:${col.bg};color:${col.dot};border-color:${col.border}">${col.items.length}</span>
          </div>
          <div class="kb-col-body">
            ${cards || `<div class="kb-empty">–</div>`}
            ${moreCount > 0 ? `<div class="kb-more">+${moreCount} weitere</div>` : ''}
          </div>
        </div>`;
    }).join('') +
  `</div>`;
}

// ── LIST VIEWS ────────────────────────────────────────────────────────────────
let mineStatusFilter  = '';
let _mineInitialized  = false; // so defaultSort only applies on first load
let mineSortOrder    = 'date-desc';

function myItems() {
  const myEmail = (account?.username || '').toLowerCase();
  return allItems.filter(i =>
    (i.createdBy?.user?.email || '').toLowerCase() === myEmail ||
    (i.createdBy?.user?.displayName || '').toLowerCase() === (account?.name || '').toLowerCase()
  );
}

function renderList(type) {
  if (type === 'mine') {
    // Apply user's defaultSort on first load (not on every re-render so user can override)
    if (account) {
      const s = getSettings(account.username);
      if (!_mineInitialized) { mineSortOrder = s.defaultSort || 'date-desc'; _mineInitialized = true; }
    }
    renderStatusChipsMine();
    filterView('mine');
    return;
  }
  const sel = $id(`filter-${type}-status`);
  if (sel && sel.options.length <= 1) {
    const statuses = [...new Set(allItems.map(i => getStatusVal(i)).filter(Boolean))];
    statuses.forEach(s => sel.add(new Option(s, s)));
  }
  filterView(type);
}

function renderStatusChipsMine() {
  const el = $id('status-chips-mine');
  if (!el) return;
  const mine = myItems();
  const counts = {};
  for (const item of mine) {
    const s = getStatusVal(item) || '';
    counts[s] = (counts[s] || 0) + 1;
  }
  const filteredVol = (mineStatusFilter
    ? mine.filter(i => (getStatusVal(i)||'') === mineStatusFilter)
    : mine
  ).reduce((sum, i) => sum + (parseFloat(getField(i, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0), 0);

  const allChip = `<button class="sc-chip${mineStatusFilter==='' ? ' active' : ''}" onclick="setMineFilter('')">
    <span class="sc-label">Alle</span><span class="sc-count">${mine.length}</span></button>`;

  const statusChips = Object.entries(counts).sort((a,b) => b[1]-a[1]).map(([s, n]) => {
    const sl = (s||'').toLowerCase().trim();
    let st = STATUS_STYLES[sl];
    if (!st) for (const [k, style] of Object.entries(STATUS_STYLES)) { if (sl.includes(k)) { st = style; break; } }
    st = st || { bg:'#f3f4f6', color:'#374151' };
    let icon = sl.includes('abgelehnt') ? '✗' : sl.includes('freigegeben') || sl.includes('bestellt') ? '✓'
             : sl.includes('prüfung') || sl.includes('bearbeitung') ? '⏳' : '📋';
    const isActive = mineStatusFilter === s;
    return `<button class="sc-chip${isActive ? ' active' : ''}" style="--sc-bg:${st.bg};--sc-color:${st.color}"
      onclick="setMineFilter('${escJs(s)}')">
      <span class="sc-icon">${icon}</span>
      <span class="sc-label">${esc(s||'Eingereicht')}</span>
      <span class="sc-count" style="background:${st.bg};color:${st.color}">${n}</span>
    </button>`;
  }).join('');

  const vol = `<span class="sc-volume">${fmtEuro(filteredVol)}<small>Volumen</small></span>`;
  el.innerHTML = allChip + statusChips + vol;
}

function setMineFilter(status) {
  mineStatusFilter = (mineStatusFilter === status) ? '' : status;
  renderStatusChipsMine();
  filterView('mine');
}

function setMineSort(val) {
  mineSortOrder = val;
  filterView('mine');
}

function filterView(type) {
  const search = ($id(`search-${type}`)?.value || '').toLowerCase();
  const priceKey = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';

  const userSettings = account ? getSettings(account.username) : {};
  let items = type === 'mine' ? myItems() : [...allItems];

  // hideCompleted: filter out terminal states in Meine Anfragen
  if (type === 'mine' && userSettings.hideCompleted && !mineStatusFilter) {
    items = items.filter(i => !/^(bestellt|erledigt|abgelehnt)$/i.test(getStatusVal(i)));
  }

  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) ||
    String(i.id||'').includes(search)
  );

  if (type === 'mine' && mineStatusFilter)
    items = items.filter(i => (getStatusVal(i)||'') === mineStatusFilter);
  else if (type !== 'mine') {
    const status = $id(`filter-${type}-status`)?.value || '';
    if (status) items = items.filter(i => (getStatusVal(i)||'') === status);
  }

  if (type === 'mine') {
    switch (mineSortOrder) {
      case 'date-asc':   items.sort((a,b) => new Date(a.createdDateTime)-new Date(b.createdDateTime)); break;
      case 'price-desc': items.sort((a,b) => (parseFloat(getField(b,priceKey))||0)-(parseFloat(getField(a,priceKey))||0)); break;
      case 'price-asc':  items.sort((a,b) => (parseFloat(getField(a,priceKey))||0)-(parseFloat(getField(b,priceKey))||0)); break;
      case 'status':     items.sort((a,b) => (getStatusVal(a)||'').localeCompare(getStatusVal(b)||'')); break;
      default:           items.sort((a,b) => new Date(b.createdDateTime)-new Date(a.createdDateTime));
    }
  }

  const container = $id(`list-${type}`);
  if (container) {
    // compactView: toggle dense layout class
    container.classList.toggle('compact-list', !!(userSettings.compactView));
    container.innerHTML = items.length
      ? _cardsHtml(items)
      : emptyState(type === 'mine' ? 'Sie haben noch keine Anfragen erstellt.' : 'Keine Anfragen gefunden.');
  }
}

// ── DETAIL VIEW ───────────────────────────────────────────────────────────────
function renderDetail(id) {
  const item = allItems.find(i => String(i.id) === String(id));
  if (!item) { $id('detail-content').innerHTML = '<p>Eintrag nicht gefunden.</p>'; return; }
  const f = item.fields || {};

  // Group known fields
  const bedarfKeys   = FORM_FIELDS.filter(fd => fd.step === 1).map(fd => resolvedFields[fd.key] || fd.key);
  const mengeKeys    = FORM_FIELDS.filter(fd => fd.step === 2).map(fd => resolvedFields[fd.key] || fd.key);
  const beschKeys    = FORM_FIELDS.filter(fd => fd.step === 3).map(fd => resolvedFields[fd.key] || fd.key);
  const einkaufKeys  = EINKAUF_FIELDS.map(fd => resolvedFields[fd.key] || fd.key);
  const knownKeys    = new Set([...bedarfKeys,...mengeKeys,...beschKeys,...einkaufKeys,'Status','Title']);

  // Remaining SP fields not in our definition
  const extraFields  = Object.entries(f)
    .filter(([k]) => !SYSTEM_FIELDS.has(k) && !knownKeys.has(k))
    .filter(([,v]) => v !== null && v !== undefined && v !== '');

  const createdBy    = item.createdBy?.user?.displayName || item.createdBy?.user?.email || getField(item,'Author') || '–';
  const createdAt    = item.createdDateTime ? fmtDateTime(item.createdDateTime) : '–';
  const statusVal    = getStatusVal(item) || 'Eingereicht';
  const preis        = parseFloat(getField(item,'GeschaetzterPreis') || getField(item, resolvedFields['GeschaetzterPreis'])) || null;
  const menge        = getField(item,'Menge') || getField(item, resolvedFields['Menge']);
  const isEinkauf    = true; // for now, all logged-in users can add order info (adjust if needed)

  $id('detail-content').innerHTML = `
    <div class="detail-header">
      <h2>${esc(getField(item,'Title') || '–')}</h2>
      <div class="detail-meta">
        <span class="item-id">ID ${item.id}</span>
        ${statusBadge(statusVal)}
        ${prioTag(getField(item,'Prioritaet') || getField(item, resolvedFields['Prioritaet']))}
      </div>
      <div class="detail-byline">Erstellt von <strong>${esc(createdBy)}</strong> am ${createdAt}</div>
      <div class="detail-actions" style="margin-top:10px">
        <button class="btn btn-sm btn-outline" id="detail-fav" title="Als Wiedervorlage / Favorit speichern">⭐ Als Wiedervorlage</button>
      </div>
    </div>

    <div class="detail-grid">
      <div class="detail-left">
        ${detailSection('Bedarf', [
          ...FORM_FIELDS.filter(fd=>fd.step===1).map(fd => detailRow(fd.label, getField(item, resolvedFields[fd.key] || fd.key)))
        ])}
        ${detailSection('Mengengaben', [
          ...FORM_FIELDS.filter(fd=>fd.step===2).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${detailSection('Beschaffungsdetails', [
          ...FORM_FIELDS.filter(fd=>fd.step===3).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${extraFields.length ? detailSection('Weitere Felder aus SharePoint', extraFields.map(([k,v]) => detailRow(k, v))) : ''}
      </div>

      <div class="detail-right">
        ${detailSidebar(item, isEinkauf)}
      </div>
    </div>`;

  // Bind update button
  const btnOrder = document.getElementById('btn-add-order');
  if (btnOrder) btnOrder.onclick = () => openOrderModal(item.id);
  const btnFav = document.getElementById('detail-fav');
  if (btnFav) btnFav.onclick = () => saveItemAsFavorite(item.id);

  // Load attachments async — cache-busting ensures freshly uploaded files appear immediately
  const attachEl = document.getElementById('detail-attachments');
  if (attachEl) {
    getSpToken().then(tok =>
      fetch(`${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${id})/AttachmentFiles?_=${Date.now()}`, {
        headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata',
          'Cache-Control': 'no-cache', Pragma: 'no-cache' }
      })
    ).then(r => r.ok ? r.json() : { value: [] })
     .then(data => {
       const files = data.value || [];
       attachEl.innerHTML = files.length
         ? files.map(attachmentLink).join('')
         : '<span class="no-order">Keine Anhänge.</span>';
     })
     .catch(() => { attachEl.innerHTML = ''; });
  }
}

function detailSection(title, rows) {
  const content = rows.filter(r => r).join('');
  if (!content) return '';
  return `<div class="detail-card"><div class="detail-card-header">${title}</div><div class="detail-card-body">${content}</div></div>`;
}
function detailRow(label, value) {
  if (!value && value !== 0) return '';
  return `<div class="detail-row"><span class="detail-label">${esc(String(label))}</span><span class="detail-value">${esc(String(value))}</span></div>`;
}
function formatFieldValue(fd, v) {
  if (!v && v !== 0) return '';
  if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') return fmtEuro(v);
  if (fd.key === 'Termin' || fd.key === 'Lieferdatum' || fd.key === 'AutoBedarfDatum') return v ? fmtDate(v) : '';
  return v;
}

function detailSidebar(item, isEinkauf) {
  const orderNr   = getField(item,'Bestellnummer')      || getField(item, resolvedFields['Bestellnummer'])      || '';
  const lieferd   = getField(item,'Lieferdatum')        || getField(item, resolvedFields['Lieferdatum'])        || '';
  const tatPreis  = getField(item,'TatsaechlicherPreis')|| getField(item, resolvedFields['TatsaechlicherPreis'])|| '';

  return `
    ${renderApprovalCard(item)}

    <div class="detail-card">
      <div class="detail-card-header">Bestellung (Einkauf)
        ${isEinkauf && /^freigegeben$/i.test((getStatusVal(item)||'').trim()) ? `<button class="btn btn-sm btn-outline" id="btn-add-order">Bearbeiten</button>` : ''}
      </div>
      <div class="detail-card-body">
        ${orderNr  ? detailRow('Bestellnummer', orderNr) : '<p class="no-order">Noch keine Bestellnummer eingetragen.</p>'}
        ${lieferd  ? detailRow('Lieferdatum', fmtDate(lieferd)) : ''}
        ${tatPreis ? detailRow('Tatsächlicher Preis', fmtEuro(tatPreis)) : ''}
        <div id="detail-attachments" class="detail-attachments">Lade Anhänge …</div>
      </div>
    </div>`;
}

// Keywords to detect approval-related SP columns by displayName or internal name
const APPROVAL_RE   = /genehmig|freigab|entscheid|ablehn|kommentar.*genehm|genehm.*kommentar/i;
const STAGE_MAP = [
  { label: 'Einkauf',                    re: /einkauf/i },
  { label: 'Werkleitung',                re: /werkleitung/i },
  { label: 'Strategischer Einkauf',      re: /strategisch/i },
  { label: 'Controlling',                re: /controlling/i },
  { label: 'Geschäftsführung',           re: /\bgf\b|geschäftsführ|geschaeftsfuehr/i },
];

// Timeline stages with direct SP field mappings (from Power Automate workflow).
// approverField / commentField are read directly from the list item — no version
// history needed.  A stage is "visible" if its approverField is non-empty
// (meaning that stage was actually reached) or it is the current status.
const TIMELINE_STAGES = [
  { label: 'Eingereicht',
    test: v => /^eingereicht$/i.test(v),
    approverField: null, commentField: null },
  { label: 'In Prüfung (Einkauf)',
    test: v => /pr[üu]fung/i.test(v) && /einkauf/i.test(v) && !/strategisch/i.test(v),
    approverField: 'Entscheider_x002a_in', commentField: 'Genehmigungskommentar' },
  { label: 'Genehmigt (Einkauf)',
    test: v => /genehmigt/i.test(v) && /einkauf/i.test(v),
    approverField: 'Entscheider_x002a_in', commentField: 'Genehmigungskommentar' },
  // Reihenfolge laut Power Automate-Freigabeprozess:
  // Controlling (nur Dienstleistung) → Werkleitung (ab 1.500 €) → strategischer Einkauf (ab 10.000 €)
  { label: 'In Prüfung (Controlling)',
    test: v => /controlling/i.test(v),
    approverField: 'Genehmiger2', commentField: 'Genehmigungskommentar2' },
  { label: 'In Prüfung (Werkleitung)',
    test: v => /werkleitung/i.test(v),
    approverField: 'Genehmiger3', commentField: 'Genehmigungskommentar3' },
  { label: 'In Prüfung (strategischer Einkauf)',
    test: v => /strategisch/i.test(v),
    approverField: 'Genehmiger4', commentField: 'Genehmigungskommentar4' },
  { label: 'Freigegeben',
    test: v => /^freigegeben$/i.test(v),
    approverField: null, commentField: null },
  { label: 'In Bestellung',
    test: v => /in bestellung/i.test(v),
    approverField: null, commentField: null },
  { label: 'Bestellt',
    test: v => /^bestellt$/i.test(v),
    approverField: null, commentField: null },
  { label: 'Abgelehnt',
    test: v => /^abgelehnt$/i.test(v),
    approverField: null, commentField: null },
];

// Resolve a person/lookup field from item.fields to a display name string.
function resolvePersonField(fields, fieldName) {
  if (!fieldName || !fields) return null;
  function tryExtract(raw) {
    if (raw == null || raw === '') return null;
    if (typeof raw === 'object') {
      const n = raw.displayName || raw.LookupValue || raw.Title || raw.title;
      return (n && n !== '[object Object]') ? String(n).trim() || null : null;
    }
    const s = String(raw).trim();
    if (!s || s === '[object Object]') return null;
    if (/^\d+$/.test(s)) return spUserMap[s] || null; // numeric LookupId
    return s;
  }
  return tryExtract(fields[fieldName])
      || tryExtract(spUserMap[String(fields[fieldName + 'LookupId'] ?? '')])
      || null;
}

// Workflow-Reihenfolge für die Status-Zeitleiste (wird in discoverSP() aus SP-Spalte befüllt)
const WORKFLOW_STAGES = [
  'Eingereicht',
  'In Prüfung (Einkauf)',
  'Genehmigt (Einkauf)',
  'In Prüfung (Controlling)',
  'In Prüfung (Werkleitung)',
  'In Prüfung (strategische Einkauf)',
  'Freigegeben',
  'In Bestellung',
  'Bestellt',
];

// Farbe für einen Status anhand von Schlüsselwörtern bestimmen
function statusColorFor(val) {
  const v = (val || '').toLowerCase();
  if (/entwurf|automatisch erstellt/.test(v)) return { bg:'#f1f5f9', color:'#64748b' }; // grau
  if (/eingereicht/.test(v))            return { bg:'#fce7f3', color:'#be185d' };
  if (/werkleitung/.test(v))            return { bg:'#ffe4e6', color:'#be123c' };
  if (/prüf|pruef|prüf/.test(v))        return { bg:'#ccfbf1', color:'#0f766e' };
  if (/freigegeben/.test(v))            return { bg:'#f3e8ff', color:'#7e22ce' };
  if (/abgelehnt/.test(v))              return { bg:'#f3f4f6', color:'#374151' };
  if (/bestellt/.test(v))               return { bg:'#dbeafe', color:'#1d4ed8' };
  if (/erledigt|abgeschlossen/.test(v)) return { bg:'#f3f4f6', color:'#374151' };
  return { bg:'#f0f9ff', color:'#0369a1' };
}

// Prädiktive Status-Zeitleiste.
// Zeigt NUR die Stufen, die für DIESE Anfrage laut Power-Automate-Freigabeprozess
// tatsächlich getriggert werden (abhängig von Beschaffungsart, Preisschwellen,
// Lead-Buyer) — vergangene, aktuelle und kommende. Nicht zutreffende Verzweigungen
// (z. B. Controlling bei Nicht-Dienstleistung) werden nie angezeigt.
//
// Genehmigungslogik beginnt in der APP (Eingereicht → In Prüfung (Einkauf) →
// Genehmigt (Einkauf)); danach übernimmt Power Automate die weiteren Stufen.
function statusTimeline(statusVal, item) {
  const sv     = (statusVal || '').trim();
  const fields = item?.fields || item || {};

  // Attribute dieser Anfrage → welche Stufen werden überhaupt durchlaufen?
  const beschaffung = String(getField(item, resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik') || '');
  const preis       = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0;
  const lbRaw       = getField(item, resolvedFields['LeadBuyerAbschluss'] || 'LeadBuyerAbschluss');
  const leadBuyer   = lbRaw === true || /^(true|ja|yes|1)$/i.test(String(lbRaw || ''));
  const isDienstleistung = /dienstleistung/i.test(beschaffung);
  const needsWerkleitung = preis >= 1500;
  const needsStrategisch = preis >= 10000 && !leadBuyer;

  // Anwendbare Stufen in Workflow-Reihenfolge (nur tatsächlich getriggerte)
  const applies = st => {
    if (/in bestellung/i.test(st.label)) return false;
    if (/^abgelehnt$/i.test(st.label))   return false;
    if (/controlling/i.test(st.label))   return isDienstleistung;
    if (/werkleitung/i.test(st.label))   return needsWerkleitung;
    if (/strategisch/i.test(st.label))   return needsStrategisch;
    return true; // Eingereicht, In Prüfung (Einkauf), Genehmigt (Einkauf), Freigegeben, Bestellt
  };
  const stages = TIMELINE_STAGES.filter(applies);

  // Anzeige-Label aus echten SP-Choices (falls geladen), sonst aus TIMELINE_STAGES
  const spLabel = st => (statusChoices.find(c => st.test(c.trim())) || st.label).trim();

  const isRej = /^abgelehnt$/i.test(sv);

  // Eine Zeitleisten-Zeile rendern (mit Genehmiger-Dedup gegen Doppelnennung)
  let lastApprover = null;
  const row = (label, dot, cls, approverField, commentField, bold) => {
    let approverHtml = '';
    const approver = resolvePersonField(fields, approverField);
    const comment  = commentField ? String(fields[commentField] || '').trim() : '';
    if (approver && approver !== lastApprover) {
      lastApprover = approver;
      const commentHtml = comment ? `<div class="ap-inline-comment">💬 ${esc(comment)}</div>` : '';
      const pastCls = cls === 'ap-ok' ? ' ap-approver-past' : '';
      approverHtml = `<div class="ap-approver${pastCls}">👤 ${esc(approver)}${commentHtml}</div>`;
    }
    const boldAttr = bold ? ' style="font-weight:600"' : '';
    return `<div class="approval-stage"><div class="ap-dot ${cls}">${dot}</div>`
         + `<div class="ap-body"><div class="ap-stage-label"${boldAttr}>${esc(label)}</div>${approverHtml}</div></div>`;
  };

  const out = [];

  // Abgelehnt: erreichte Stufen (Genehmiger hinterlegt) als erledigt, dann „Abgelehnt"
  if (isRej) {
    for (const st of stages) {
      if (/^(freigegeben|bestellt)$/i.test(st.label)) continue;
      if (/^eingereicht$/i.test(st.label)) { out.push(row(spLabel(st), '✓', 'ap-ok', null, null, false)); continue; }
      if (resolvePersonField(fields, st.approverField)) {
        out.push(row(spLabel(st), '✓', 'ap-ok', st.approverField, st.commentField, false));
      }
    }
    const rejLabel = (statusChoices.find(c => /^abgelehnt$/i.test(c.trim())) || 'Abgelehnt').trim();
    out.push(row(rejLabel, '✗', 'ap-no', null, null, false));
    return out.join('');
  }

  // Normaler Verlauf: erledigt (✓) / aktuell (●) / kommend (○)
  const curIdx = stages.findIndex(st => st.test(sv));
  stages.forEach((st, i) => {
    const label        = spLabel(st);
    const isTerminalOk = /^(freigegeben|bestellt)$/i.test(st.label);
    let dot, cls, bold = false;
    if (curIdx === -1) {                       // unbekannter Status → erste Stufe erledigt, Rest kommend
      if (i === 0) { dot = '✓'; cls = 'ap-ok'; } else { dot = '○'; cls = 'ap-circle-future'; }
    } else if (i < curIdx) {                   // erledigt
      dot = '✓'; cls = 'ap-ok';
    } else if (i === curIdx) {                 // aktuell
      if (isTerminalOk) { dot = '✓'; cls = 'ap-ok'; } else { dot = '●'; cls = 'ap-pending'; bold = true; }
    } else {                                   // kommend
      dot = '○'; cls = 'ap-circle-future';
    }
    // Genehmiger nur bei erledigten/aktuellen Stufen anzeigen (kommende noch ohne)
    const show = curIdx !== -1 && i <= curIdx;
    out.push(row(label, dot, cls, show ? st.approverField : null, show ? st.commentField : null, bold));
  });

  return out.join('');
}

function approvalStyle(val) {
  const v = (val || '').toLowerCase();
  if (/freigegeben|genehmigt|approved|ja\b/.test(v)) return { bg:'#f0fdf4', color:'#15803d', dot:'✓', cls:'ap-ok' };
  if (/abgelehnt|rejected|nein\b/.test(v))           return { bg:'#fef2f2', color:'#b91c1c', dot:'✗', cls:'ap-no' };
  return { bg:'#fffbeb', color:'#b45309', dot:'…', cls:'ap-pending' };
}

function renderApprovalCard(item) {
  // Delegate to buildApprovalInner so detail view and panel share identical logic
  // (same timeline, same inline comments, same approver placement).
  const statusVal = getStatusVal(item) || 'Eingereicht';
  return `
    <div class="detail-card">
      <div class="detail-card-header">Status &amp; Genehmigung</div>
      <div class="detail-card-body">
        <div class="ap-current-status">${statusWithApprover(item)}</div>
        ${buildApprovalInner(item, statusVal)}
      </div>
    </div>`;
}

// ── EINKAUF ORDER MODAL ───────────────────────────────────────────────────────
async function openOrderModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  if (!item || !/^freigegeben$/i.test((getStatusVal(item)||'').trim())) {
    toast('Einkauf-Daten können nur bei Status „Freigegeben" eingetragen werden.', 'error');
    return;
  }
  const orderNr    = getField(item,'Bestellnummer')       || getField(item, resolvedFields['Bestellnummer'])       || '';
  const lieferd    = getField(item,'Lieferdatum')         || getField(item, resolvedFields['Lieferdatum'])         || '';
  const tatPreis   = getField(item,'TatsaechlicherPreis') || getField(item, resolvedFields['TatsaechlicherPreis']) || '';
  const termin     = getField(item,'Termin')              || getField(item, resolvedFields['Termin'])              || '';
  const geschPreis = getField(item,'GeschaetzterPreis')   || getField(item, resolvedFields['GeschaetzterPreis'])   || '';
  const lief1      = getField(item,'Lieferant')           || getField(item, resolvedFields['Lieferant'])           || '';
  const lief2      = getField(item,'Lieferant2')          || getField(item, resolvedFields['Lieferant2'])          || '';
  const lief3      = getField(item,'Lieferant3')          || getField(item, resolvedFields['Lieferant3'])          || '';
  const lief4      = getField(item,'Lieferant4')          || getField(item, resolvedFields['Lieferant4'])          || '';

  // Fetch existing attachments
  let existingFiles = [];
  try {
    const tok = await getSpToken();
    const r = await fetch(`${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles`, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' }
    });
    if (r.ok) existingFiles = (await r.json()).value || [];
  } catch(e) { /* non-critical */ }

  const existingHtml = existingFiles.length ? `
    <div class="attach-existing">
      ${existingFiles.map(attachmentLink).join('')}
    </div>` : '';

  const required   = angeboteAnzahl(parseFloat(geschPreis) || 0);
  const attachNote = required === 0
    ? 'Angebotsunterlagen (optional)'
    : `Angebote – Regelwerk: mind. ${required} erforderlich`;
  // Always show 1 slot; extra slots (up to required-1 or 4 more) go into a collapsed <details>
  const extraN     = Math.max(0, required - 1);
  const extraSlots = Array.from({length: extraN}, (_, i) =>
    `<input type="file" class="attach-file-input" accept=".pdf,.PDF" style="margin-bottom:6px" data-slot="${i+2}"/>`
  ).join('');

  $id('modal-title').textContent = 'Einkauf-Daten eintragen';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellnummer</label>
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. BE252093"
        oninput="checkBeDuplicate(this.value, ${itemId})"/>
      <div id="be-warn" class="be-warn" style="display:none"></div>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferant 1</label>
      <input type="text" id="m-lief1" value="${esc(lief1)}" placeholder="Firmenname oder Lieferanten-Nr."/>
      <details style="margin-top:6px" ${(lief2 || lief3 || lief4) ? 'open' : ''}>
        <summary class="attach-more-toggle">+ Weitere Lieferanten</summary>
        <div style="margin-top:6px;display:flex;flex-direction:column;gap:6px">
          <input type="text" id="m-lief2" value="${esc(lief2)}" placeholder="Lieferant 2 (Alternative)"/>
          <input type="text" id="m-lief3" value="${esc(lief3)}" placeholder="Lieferant 3 (Alternative)"/>
          <input type="text" id="m-lief4" value="${esc(lief4)}" placeholder="Lieferant 4 (Alternative)"/>
        </div>
      </details>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label style="display:flex;align-items:center;justify-content:space-between">
        Lieferdatum
        ${termin ? `<button type="button" class="btn-prefill" onclick="document.getElementById('m-delivery').value='${toLocalInputDate(termin)}'">← Aus Anfrage (${fmtDate(termin)})</button>` : ''}
      </label>
      <input type="date" id="m-delivery" value="${toLocalInputDate(lieferd)}"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label style="display:flex;align-items:center;justify-content:space-between">
        Tatsächlicher Preis netto (€)
        ${geschPreis ? `<button type="button" class="btn-prefill" onclick="document.getElementById('m-price').value='${geschPreis}'">← Aus Anfrage (${fmtEuro(geschPreis)})</button>` : ''}
      </label>
      <input type="number" id="m-price" value="${tatPreis}" min="0" step="0.01" placeholder="0,00"/>
    </div>
    <div class="form-group">
      <label>${esc(attachNote)}</label>
      ${existingHtml}
      <div id="attach-inputs">
        <input type="file" class="attach-file-input" accept=".pdf,.PDF" style="margin-bottom:6px" data-slot="1"/>
      </div>
      <details id="attach-extra-wrap" style="margin-top:4px">
        <summary class="attach-more-toggle">+ Weitere Angebote${required > 1 ? ` (${extraN} gem. Regelwerk vorausgefüllt)` : ''}</summary>
        <div id="attach-extra-inputs" style="margin-top:6px">${extraSlots}</div>
        <button type="button" class="btn btn-sm btn-ghost" id="btn-add-attach" onclick="addAttachInput()" style="margin-top:4px">+ Weiteres hinzufügen</button>
      </details>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveOrderData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
  // Initiale Prüfung, falls bereits eine BE eingetragen ist
  checkBeDuplicate(orderNr, itemId);
}

// Warnt, wenn die eingegebene Bestellnummer bereits bei einem anderen Item verwendet wird
function checkBeDuplicate(val, currentId) {
  const el = $id('be-warn');
  if (!el) return;
  const v = (val || '').trim().toLowerCase();
  if (!v) { el.style.display = 'none'; el.innerHTML = ''; return; }
  const beCol = resolvedFields['Bestellnummer'] || 'Bestellnummer';
  const dupes = allItems.filter(i =>
    String(i.id) !== String(currentId) &&
    (getField(i, beCol) || '').trim().toLowerCase() === v
  );
  if (dupes.length) {
    el.style.display = 'block';
    const refs = dupes.map(d => {
      const t = cleanTitle(getField(d, 'Title') || '');
      return `#${d.id}${t ? ' – ' + esc(t) : ''}`;
    }).join(', ');
    el.innerHTML = `⚠️ Achtung: Bestellnummer „${esc(val.trim())}" wird bereits verwendet (${refs}).`;
  } else {
    el.style.display = 'none';
    el.innerHTML = '';
  }
}

function addAttachInput() {
  const extra  = $id('attach-extra-inputs') || $id('attach-inputs');
  const addBtn = $id('btn-add-attach');
  if (!extra) return;
  const total = document.querySelectorAll('.attach-file-input').length;
  if (total >= 5) { if (addBtn) addBtn.disabled = true; return; }
  const inp = document.createElement('input');
  inp.type = 'file'; inp.accept = '.pdf,.PDF'; inp.className = 'attach-file-input';
  inp.style.marginBottom = '6px';
  extra.appendChild(inp);
  if (document.querySelectorAll('.attach-file-input').length >= 5 && addBtn) addBtn.disabled = true;
}

async function saveOrderData(itemId) {
  const orderNr  = $id('m-ordernr').value.trim();
  const delivery = $id('m-delivery').value;
  const price    = $id('m-price').value;

  // Collect selected files
  const fileInputs = [...document.querySelectorAll('.attach-file-input')];
  const files = fileInputs.map(i => i.files[0]).filter(Boolean);

  // Build patch using resolved names; fall back to raw key if column was never resolved
  const patch = {};
  const add = (key, val) => {
    if (!val) return;
    const col = resolvedFields[key] || key;
    if (col) patch[col] = val;
  };
  add('Bestellnummer', orderNr);
  if (delivery) add('Lieferdatum', toSpDate(delivery, resolvedFields['Lieferdatum'] || 'Lieferdatum'));
  if (price) add('TatsaechlicherPreis', parseFloat(price));
  // Lieferanten sind voll editierbar (ändern UND entfernen). Nur tatsächlich geänderte
  // Felder schreiben – so wird ein geleertes Feld auch wirklich in SharePoint geleert,
  // unveränderte Felder lösen aber keinen unnötigen Write aus.
  const _item   = allItems.find(i => String(i.id) === String(itemId));
  const curLief = key => String(getField(_item, key) || getField(_item, resolvedFields[key]) || '').trim();
  const setLief = (key, elId) => {
    const el = $id(elId); if (!el) return;
    const val = el.value.trim();
    if (val !== curLief(key)) { const col = resolvedFields[key] || key; if (col) patch[col] = val; }
  };
  setLief('Lieferant',  'm-lief1');
  setLief('Lieferant2', 'm-lief2');
  setLief('Lieferant3', 'm-lief3');
  setLief('Lieferant4', 'm-lief4');
  // Wenn Bestellnummer vergeben → Status auf "Bestellt" setzen
  if (orderNr) {
    const statusCol = resolvedFields['Status'] || 'Status';
    if (statusCol) patch[statusCol] = 'Bestellt';
  }

  const hasFields = Object.keys(patch).length > 0;
  if (!hasFields && !files.length) { closeModal(); return; }

  // PATCH fields if any
  if (hasFields) {
    const skipped = [];
    for (let i = 0; i < 10; i++) {
      try {
        await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
        break;
      } catch(e) {
        const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
        if (!m) { toast('Fehler: ' + e.message, 'error'); return; }
        skipped.push(m[1]);
        delete patch[m[1]];
      }
    }
  }

  // Upload attachments
  if (files.length) {
    const tok = await getSpToken();
    const errors = [];
    for (const file of files) {
      try {
        // Einfache Anführungszeichen müssen im OData-Stringliteral verdoppelt
        // werden ('' = ein '); encodeURIComponent kodiert sie nicht. Ohne das
        // schlägt der Upload bei Dateinamen wie "Angebot'25.pdf" mit 400 fehl.
        const fname = encodeURIComponent(file.name.replace(/'/g, "''"));
        const r = await fetch(
          `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${fname}')`,
          { method: 'POST', headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' }, body: await file.arrayBuffer() }
        );
        if (!r.ok) errors.push(file.name + ' (' + r.status + ')');
      } catch(e) { errors.push(file.name); }
    }
    if (errors.length) toast('Fehler beim Hochladen: ' + errors.join(', '), 'error');
    else toast('Einkauf-Daten und Anhänge gespeichert ✓', 'success');
  } else {
    toast('Einkauf-Daten gespeichert ✓', 'success');
  }

  closeModal();
  await loadItems(false);
  renderDetail(itemId);
}

// ── BESCHAFFUNGSDETAILS MODAL (Einkauf) ───────────────────────────────────────
function openBeschModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  if (!item) return;

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const BL_OPTS = (colByKey[resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik']?.choice?.choices)
    || ['Bestandsmaterial (bestandsgeführt)','Nicht-bestandsgeführtes Material','Direktes Material','Indirektes Material / Dienstleistung'];

  const selOpts = (opts, cur) => opts.map(o =>
    `<option value="${esc(o)}"${cur === o ? ' selected' : ''}>${esc(o)}</option>`
  ).join('');

  $id('modal-title').textContent = 'Beschaffungsdetails bearbeiten';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Beschaffungsart</label>
      <select id="b-beschaffungslogik" class="form-control" style="width:100%">
        <option value="">–</option>${selOpts(BL_OPTS, gv('Beschaffungslogik'))}
      </select>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Artikelnummer</label>
      <input type="text" id="b-artikelnummer" value="${esc(gv('Artikelnummer'))}" placeholder="z. B. 4001-00010"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferant 1</label>
      <input type="text" id="b-lieferant" value="${esc(gv('Lieferant'))}" placeholder="Firmenname oder Lieferanten-Nr."/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferant 2</label>
      <input type="text" id="b-lieferant2" value="${esc(gv('Lieferant2'))}" placeholder="optional"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellvolumen in €</label>
      <input type="number" id="b-preis" value="${esc(String(gv('GeschaetzterPreis')))}" min="0" step="0.01" placeholder="0,00"/>
    </div>
    <div class="form-group">
      <label>Kostenstelle</label>
      <input type="text" id="b-kostenstelle" value="${esc(gv('Kostenstelle'))}" placeholder="z. B. 4200"/>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveBeschData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
  initKostenstAuto($id('b-kostenstelle'));
}

async function saveBeschData(itemId) {
  const patch = {};
  const set = (key, val) => {
    if (val === null || val === undefined || val === '') return;
    const col = resolvedFields[key] || key;
    if (col) patch[col] = val;
  };
  set('Beschaffungslogik', $id('b-beschaffungslogik').value);
  set('Artikelnummer',     $id('b-artikelnummer').value.trim());
  set('Lieferant',         $id('b-lieferant').value.trim());
  set('Lieferant2',        $id('b-lieferant2').value.trim());
  const preis = parseFloat($id('b-preis').value);
  if (!isNaN(preis) && preis > 0) {
    const col = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
    const colDef = colByKey[col];
    patch[col] = colDef?.number ? preis : String(preis);
  }
  set('Kostenstelle', $id('b-kostenstelle').value.trim());
  if (!Object.keys(patch).length) { closeModal(); return; }

  const btn = $id('modal-footer')?.querySelector('.btn-primary');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }
  const skipped = [];
  for (let i = 0; i < 10; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      toast('Beschaffungsdetails gespeichert ✓', 'success');
      closeModal();
      await loadItems(false);
      const pi = allItems.find(x => String(x.id) === String(itemId));
      if (pi && panelItemId) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(pi); bindPanelEvents(itemId); }
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
}

// ── ARTIKELNUMMER AUTOCOMPLETE ────────────────────────────────────────────────
// Shared between wizard (lookupTID) and panel edit mode (initTidAutocomplete).
// Searches TID_MAP_ACTIVE by number prefix OR name substring.

function initTidAutocomplete(inputEl, getRelatedInput) {
  // Dropdown appended to <body> with position:fixed to escape panel overflow clipping
  let dropdown = null;
  let activeIdx = -1;

  function getOrCreateDropdown() {
    if (!dropdown) {
      dropdown = document.createElement('div');
      dropdown.className = 'tid-ac-dropdown';
      dropdown.style.cssText = 'display:none;position:fixed;z-index:9999';
      document.body.appendChild(dropdown);
    }
    return dropdown;
  }

  function reposition() {
    if (!dropdown || dropdown.style.display === 'none') return;
    const r = inputEl.getBoundingClientRect();
    // Vorschlagsliste mind. 380px breit (damit auch der beschreibende Text sichtbar ist),
    // aber nie breiter als der Viewport; ragt sie rechts heraus, nach links verschieben.
    const w = Math.min(Math.max(r.width, 380), window.innerWidth - 16);
    let   left = r.left;
    if (left + w > window.innerWidth - 8) left = Math.max(8, window.innerWidth - 8 - w);
    dropdown.style.left  = left + 'px';
    dropdown.style.top   = (r.bottom + 2) + 'px';
    dropdown.style.width = w + 'px';
  }

  function hide() {
    if (dropdown) dropdown.style.display = 'none';
    activeIdx = -1;
  }

  function show(items) {
    if (!items.length) { hide(); return; }
    activeIdx = -1;
    const dd = getOrCreateDropdown();
    dd.innerHTML = items.slice(0, 12).map(([ tid, h ], i) =>
      `<div class="tid-ac-item" data-tid="${esc(tid)}" data-idx="${i}">
        <span class="tid-ac-nr">${esc(tid)}</span>
        <span class="tid-ac-name">${esc(h.b)}</span>
        <span class="tid-ac-wg">${esc(h.w)}</span>
      </div>`
    ).join('');
    dd.querySelectorAll('.tid-ac-item').forEach(el => {
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(el.dataset.tid); });
    });
    dd.style.display = 'block';
    reposition();
  }

  function selectItem(tid) {
    const hit = TID_MAP_ACTIVE[tid];
    if (!hit) return;
    inputEl.value = tid;
    const titleEl = getRelatedInput?.('Title');
    const wgEl    = getRelatedInput?.('Warengruppe');
    if (titleEl && !titleEl.dataset?.manual) titleEl.value = hit.b;
    if (wgEl) {
      let found = false;
      for (const opt of (wgEl.options || [])) {
        if (opt.value === hit.w) { wgEl.value = hit.w; found = true; break; }
      }
      if (!found && wgEl.tagName === 'SELECT') {
        wgEl.add(new Option(hit.w, hit.w)); wgEl.value = hit.w;
      } else if (wgEl.tagName === 'INPUT') { wgEl.value = hit.w; }
    }
    const confirm = inputEl.closest('.tid-ac-wrap')?.querySelector('.tid-ac-confirm');
    if (confirm) confirm.textContent = '✓ ' + hit.b + ' · ' + hit.w;
    // Also update wizard tid-match if present
    if (typeof lookupTID === 'function' && $id('tid-match')) lookupTID(tid);
    hide();
  }

  inputEl.addEventListener('input', () => {
    const q = inputEl.value.trim().toLowerCase();
    if (q.length < 2) { hide(); return; }
    const matches = TID_ENTRIES.filter(([tid, h]) =>
      tid.toLowerCase().includes(q) || h.b.toLowerCase().includes(q)
    );
    show(matches);
  });

  inputEl.addEventListener('keydown', e => {
    if (!dropdown || dropdown.style.display === 'none') return;
    const items = dropdown.querySelectorAll('.tid-ac-item');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'Enter' && activeIdx >= 0) {
      e.preventDefault();
      selectItem(items[activeIdx]?.dataset.tid);
    } else if (e.key === 'Escape') { hide(); }
  });

  // Reposition on scroll/resize; hide on blur
  inputEl.addEventListener('blur', () => setTimeout(hide, 150));
  window.addEventListener('scroll', reposition, true);
  window.addEventListener('resize', reposition);

  // Cleanup, sobald das Feld aus dem DOM fliegt (Panel-Schließen/Neu-Rendern): Dropdown
  // entfernen UND window-Listener + Observer abmelden, sonst sammeln sie sich bei
  // wiederholtem Neu-Rendern (z. B. Sammelanfrage-Positionen) an.
  const _mo = new MutationObserver(() => {
    if (document.contains(inputEl)) return;
    hide(); if (dropdown) dropdown.remove(); dropdown = null;
    window.removeEventListener('scroll', reposition, true);
    window.removeEventListener('resize', reposition);
    _mo.disconnect();
  });
  _mo.observe(document.body, { childList: true, subtree: true });
}

// ── KOSTENSTELLEN AUTOCOMPLETE ────────────────────────────────────────────────
// Wraps an <input> with a dropdown showing matching Kostenstellen entries.
// On select: fills input with "NUMMER – Bezeichnung"
function initKostenstAuto(inputEl) {
  if (!inputEl) return;
  let dropdown = null;
  let activeIdx = -1;

  function getOrCreate() {
    if (!dropdown) {
      dropdown = document.createElement('div');
      dropdown.className = 'tid-ac-dropdown';
      dropdown.style.cssText = 'display:none;position:fixed;z-index:9999';
      document.body.appendChild(dropdown);
    }
    return dropdown;
  }
  function reposition() {
    if (!dropdown || dropdown.style.display === 'none') return;
    const r = inputEl.getBoundingClientRect();
    // Vorschlagsliste mind. 380px breit (damit auch der beschreibende Text sichtbar ist),
    // aber nie breiter als der Viewport; ragt sie rechts heraus, nach links verschieben.
    const w = Math.min(Math.max(r.width, 380), window.innerWidth - 16);
    let   left = r.left;
    if (left + w > window.innerWidth - 8) left = Math.max(8, window.innerWidth - 8 - w);
    dropdown.style.left  = left + 'px';
    dropdown.style.top   = (r.bottom + 2) + 'px';
    dropdown.style.width = w + 'px';
  }
  function hide() { if (dropdown) dropdown.style.display = 'none'; activeIdx = -1; }
  function show(items) {
    if (!items.length) { hide(); return; }
    activeIdx = -1;
    const dd = getOrCreate();
    dd.innerHTML = items.slice(0, 50).map((e, i) =>
      `<div class="tid-ac-item" data-nr="${esc(e.nr)}" data-idx="${i}">
        <span class="tid-ac-nr">${esc(e.nr)}</span>
        <span class="tid-ac-name">${esc(e.label)}</span>
      </div>`
    ).join('');
    dd.querySelectorAll('.tid-ac-item').forEach(el => {
      el.addEventListener('mousedown', ev => { ev.preventDefault(); selectItem(el.dataset.nr); });
    });
    dd.style.display = 'block';
    reposition();
  }
  function selectItem(nr) {
    const entry = KOSTENST_DATA.find(e => e.nr === nr);
    if (!entry) return;
    inputEl.value = entry.nr + ' – ' + entry.label;
    hide();
  }
  // Liefert die (sortierte) Trefferliste. Leere Eingabe = komplette Liste zum
  // Durchblättern – so kann die Kostenstelle ausgewählt werden, ohne sie zu kennen.
  function _kstMatches(q) {
    q = (q || '').trim().toLowerCase();
    const list = q.length < 1
      ? KOSTENST_DATA.slice()
      : KOSTENST_DATA.filter(e => e.nr.toLowerCase().includes(q) || e.label.toLowerCase().includes(q));
    return list.sort((a, b) =>
      (parseInt(a.nr, 10) || 0) - (parseInt(b.nr, 10) || 0) || a.nr.localeCompare(b.nr));
  }
  // Fokus/Klick: ganze Liste zum Auswählen anzeigen. Tippen: nach Nr. ODER Bezeichnung filtern.
  inputEl.addEventListener('focus', () => show(_kstMatches('')));
  inputEl.addEventListener('input', () => show(_kstMatches(inputEl.value)));
  inputEl.addEventListener('keydown', e => {
    if (!dropdown || dropdown.style.display === 'none') return;
    const items = dropdown.querySelectorAll('.tid-ac-item');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'Enter' && activeIdx >= 0) {
      e.preventDefault();
      selectItem(items[activeIdx]?.dataset.nr);
    } else if (e.key === 'Escape') { hide(); }
  });
  inputEl.addEventListener('blur', () => setTimeout(hide, 150));
  window.addEventListener('scroll', reposition, true);
  window.addEventListener('resize', reposition);
  // Beim Entfernen des Felds Dropdown, window-Listener und Observer abmelden (kein Leck).
  const _mo = new MutationObserver(() => {
    if (document.contains(inputEl)) return;
    hide(); if (dropdown) dropdown.remove(); dropdown = null;
    window.removeEventListener('scroll', reposition, true);
    window.removeEventListener('resize', reposition);
    _mo.disconnect();
  });
  _mo.observe(document.body, { childList: true, subtree: true });
}

// ── WIZARD ────────────────────────────────────────────────────────────────────
function lookupTID(val) {
  const tid = (val || '').trim().toUpperCase();
  const matchEl = $id('tid-match');
  // Normalize: user might type "4001-00010" or "400100010" — try exact then padded
  const hit = TID_MAP_ACTIVE[tid] || TID_MAP_ACTIVE[tid.replace(/^(\d{4})(\d{5})$/, '$1-$2')] || null;
  if (hit) {
    // Auto-fill Bezeichnung (Title) and Warengruppe
    const titleEl = $id('f-Title');
    const wgEl    = $id('f-Warengruppe');
    if (titleEl && !titleEl.dataset.manual) titleEl.value = hit.b;
    if (wgEl) {
      // Try to select matching option
      let found = false;
      for (const opt of wgEl.options) {
        if (opt.value === hit.w) { wgEl.value = hit.w; found = true; break; }
      }
      if (!found) {
        const newOpt = new Option(hit.w, hit.w);
        wgEl.add(newOpt);
        wgEl.value = hit.w;
      }
    }
    if (matchEl) {
      matchEl.innerHTML = `<span class="tid-ok">✓ ${esc(hit.b)}</span> <span class="tid-wg">${esc(hit.w)}</span>`;
      matchEl.style.display = 'block';
    }
  } else if (tid.length >= 4) {
    if (matchEl) {
      matchEl.innerHTML = '<span class="tid-miss">Artikelnummer nicht im Stamm – bitte Bezeichnung und Warengruppe manuell eintragen.</span>';
      matchEl.style.display = 'block';
    }
  } else {
    if (matchEl) matchEl.style.display = 'none';
  }
}

// ── WIEDERVORLAGE / FAVORITEN ────────────────────────────────────────────────
// Primär in einer SharePoint-Liste gespeichert (geräteübergreifend); fällt bei
// fehlender Liste/Berechtigung automatisch auf localStorage zurück.
const LS_FAVORITES   = 'DIHAG_FAVORITES';
const FAV_LIST_NAME  = 'BANF_Favoriten';
let pendingPrefill = null; // {type, data, positions} – wird beim nächsten Wizard-Init angewendet
let favUseSP   = false;    // true, sobald die SP-Liste verfügbar ist
let favListId  = null;
let favCache   = [];       // In-Memory-Spiegel der SP-Favoriten des aktuellen Nutzers

function _favOwner() { return (account?.username || '').trim().toLowerCase(); }

// Sucht die Favoriten-Liste (case-insensitiv, ohne Fehler zu werfen)
async function _findFavList() {
  try {
    const all = await gGet(`/sites/${siteId}/lists?$select=id,displayName&$top=500`);
    const hit = (all.value || []).find(l =>
      (l.displayName || '').trim().toLowerCase() === FAV_LIST_NAME.toLowerCase());
    return hit?.id || null;
  } catch { return null; }
}

// Versucht, die Liste anzulegen. Gibt bei fehlender Berechtigung (403) leise null
// zurück – KEIN console.error (das Anlegen erfordert Site-Owner-Rechte).
async function _createFavList() {
  try {
    const tok = await getToken();
    const r = await fetch(API + `/sites/${siteId}/lists`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        displayName: FAV_LIST_NAME,
        list: { template: 'genericList' },
        columns: [
          { name: 'FavOwner',     text: {} },
          { name: 'FavType',      text: {} },
          { name: 'FavData',      text: { allowMultipleLines: true } },
          { name: 'FavPositions', text: { allowMultipleLines: true } },
        ],
      }),
    });
    if (r.ok) return (await r.json()).id;
    console.info(`[Favoriten] Liste konnte nicht automatisch angelegt werden (HTTP ${r.status}).`);
    return null;
  } catch { return null; }
}

// Beim Start aufrufen (nach discoverSP). Niemals werfen – Fallback ist localStorage.
async function initFavorites() {
  favUseSP = false; favListId = null; favCache = [];
  try {
    if (!siteId) return;
    // 1) Bestehende Liste suchen (Normalfall, seit sie angelegt wurde)
    favListId = await _findFavList();
    // 2) Nur wenn nicht vorhanden: leise versuchen anzulegen, danach erneut suchen
    if (!favListId) {
      favListId = await _createFavList();
      if (!favListId) favListId = await _findFavList(); // evtl. parallel angelegt
    }
    favUseSP = !!favListId;
    if (favUseSP) {
      await loadFavoritesSP();
      _refreshFavBars();
    } else {
      console.info('[Favoriten] Keine SP-Liste verfügbar – nutze lokalen Speicher.');
    }
  } catch (e) {
    console.warn('[Favoriten] SP nicht verfügbar, nutze localStorage:', e.message);
    favUseSP = false;
  }
}

async function loadFavoritesSP() {
  if (!favUseSP) return;
  const owner = _favOwner();
  const res = await gGet(`/sites/${siteId}/lists/${favListId}/items?$expand=fields&$top=500`);
  favCache = (res.value || [])
    .filter(it => (it.fields?.FavOwner || '').trim().toLowerCase() === owner)
    .map(it => {
      let data = {}, positions = null;
      try { data = JSON.parse(it.fields?.FavData || '{}'); } catch {}
      try { const p = JSON.parse(it.fields?.FavPositions || 'null'); if (Array.isArray(p) && p.length) positions = p; } catch {}
      return {
        id: String(it.id),
        name: it.fields?.Title || data.Title || 'Anfrage',
        type: it.fields?.FavType || (positions ? 'multi' : 'single'),
        data, positions,
      };
    });
}

// Liefert Favoriten synchron (SP-Cache oder localStorage)
function getFavorites() {
  if (favUseSP) return favCache.slice();
  try { return JSON.parse(localStorage.getItem(LS_FAVORITES) || '[]'); }
  catch { return []; }
}
function saveFavorites(arr) {
  localStorage.setItem(LS_FAVORITES, JSON.stringify(arr));
}

function _refreshFavBars() {
  if (currentView === 'new')   renderFavBar('single');
  if (currentView === 'multi') renderFavBar('multi');
}

async function saveItemAsFavorite(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  if (!item) return;
  let positions = null;
  const posRaw = _rf(item, 'Positionen');
  if (posRaw) { try { const p = JSON.parse(posRaw); if (Array.isArray(p) && p.length) positions = p; } catch {} }
  const type  = positions ? 'multi' : 'single';
  const title = cleanTitle(getField(item, 'Title') || '') || 'Anfrage';
  const data = {};
  ['Title','Beschreibung','Warengruppe','Prioritaet','Artikelnummer','ExterneArtikelnummer',
   'Menge','Mengeneinheit','Mindestlagermenge','Termin','Beschaffungslogik',
   'Lieferant','Lieferant2','Lieferant3','Lieferant4','GeschaetzterPreis','Kostenstelle']
    .forEach(k => { data[k] = k === 'Title' ? title : _rf(item, k); });

  if (favUseSP) {
    try {
      const created = await gPost(`/sites/${siteId}/lists/${favListId}/items`, {
        fields: {
          Title: title, FavOwner: _favOwner(), FavType: type,
          FavData: JSON.stringify(data),
          FavPositions: positions ? JSON.stringify(positions) : '',
        },
      });
      favCache.unshift({ id: String(created.id), name: title, type, data, positions });
      toast('Als Wiedervorlage gespeichert ⭐', 'success');
      _refreshFavBars();
      return;
    } catch (e) {
      toast('Speichern in SharePoint fehlgeschlagen, lokal gesichert.', 'info');
      // Fallback unten
    }
  }
  const favs = getFavorites();
  favs.unshift({ id: 'fav_' + Date.now(), name: title, type, data, positions, savedAt: new Date().toISOString() });
  saveFavorites(favs);
  toast('Als Wiedervorlage gespeichert ⭐', 'success');
  _refreshFavBars();
}

async function deleteFavorite(favId, ev) {
  if (ev) ev.stopPropagation();
  if (favUseSP && !String(favId).startsWith('fav_')) {
    try {
      await gDelete(`/sites/${siteId}/lists/${favListId}/items/${favId}`);
      favCache = favCache.filter(f => f.id !== favId);
      _refreshFavBars();
      return;
    } catch (e) {
      toast('Löschen fehlgeschlagen: ' + e.message, 'error');
      return;
    }
  }
  saveFavorites(getFavorites().filter(f => f.id !== favId));
  _refreshFavBars();
}

function useFavorite(favId) {
  const fav = getFavorites().find(f => f.id === favId);
  if (!fav) return;
  pendingPrefill = fav;
  navigate(fav.type === 'multi' ? 'multi' : 'new');
}

function renderFavBar(kind) {
  const el = $id(kind === 'multi' ? 'fav-bar-multi' : 'fav-bar-single');
  if (!el) return;
  const favs = getFavorites().filter(f => (f.type || 'single') === (kind === 'multi' ? 'multi' : 'single'));
  if (!favs.length) { el.innerHTML = ''; el.style.display = 'none'; return; }
  el.style.display = 'block';
  el.innerHTML = `
    <div class="fav-bar">
      <div class="fav-bar-title">⭐ Wiedervorlagen</div>
      <div class="fav-chips">
        ${favs.map(f => `
          <div class="fav-chip" onclick="useFavorite('${f.id}')" title="Übernehmen: ${esc(f.name)}">
            <span class="fav-chip-name">${esc(f.name)}</span>
            <button class="fav-chip-del" onclick="deleteFavorite('${f.id}', event)" title="Entfernen">✕</button>
          </div>`).join('')}
      </div>
    </div>`;
}

// Füllt den Einzel-Wizard mit Favoritendaten
// Setzt einen <select> tolerant: exakt, sonst über den beschreibenden Teil
// (ohne „NNNN – "-Präfix), sonst per Teilstring. So greift z. B. altes
// „Ersatzteile" auch auf die Option „4320 – Ersatzteile" (und umgekehrt).
function _setSelectFlexible(id, val) {
  const el = $id(id);
  if (!el || val == null || String(val).trim() === '') return;
  const v   = String(val).trim();
  const opts = [...el.options].filter(o => o.value !== '');
  const desc = s => String(s).replace(/^\s*\d+\s*[–-]\s*/, '').trim().toLowerCase();
  let opt = opts.find(o => o.value.trim() === v)
         || opts.find(o => desc(o.value) === desc(v))
         || opts.find(o => o.value.toLowerCase().includes(v.toLowerCase()) || v.toLowerCase().includes(desc(o.value)));
  if (opt) el.value = opt.value;
}

// Beschaffungslogik wiederherstellen: erster Teil = Radio (Beschaffungsart),
// weitere Teile = „Zusätzlich kombinierbar"-Karten.
function _applyBeschaffungPrefill(beschaffung, radioName, extraContainerId) {
  if (!beschaffung) return;
  const parts    = String(beschaffung).split(',').map(s => s.trim()).filter(Boolean);
  const radioVal = (parts[0] || '').toLowerCase();
  const extraVals = parts.slice(1).map(s => s.toLowerCase());
  const radios = [...document.querySelectorAll(`input[name=${radioName}]`)];
  const rm = radios.find(r => r.value.trim().toLowerCase() === radioVal);
  if (rm) rm.checked = true;
  document.querySelectorAll(`#${extraContainerId} .check-card`).forEach(c => {
    const cv = (c.dataset.value || '').trim().toLowerCase();
    c.classList.toggle('selected', extraVals.includes(cv));
  });
}

function applyWizardPrefill(data) {
  if (!data) return;
  const setVal = (id, v) => { const el = $id(id); if (el && v != null && v !== '') el.value = v; };
  setVal('f-Title', data.Title);
  setVal('f-Beschreibung', data.Beschreibung);
  _setSelectFlexible('f-Warengruppe', data.Warengruppe);
  _setSelectFlexible('f-Prioritaet', data.Prioritaet);
  setVal('f-Artikelnummer', data.Artikelnummer);
  setVal('f-ExterneArtikelnummer', data.ExterneArtikelnummer);
  setVal('f-Menge', data.Menge);
  _setSelectFlexible('f-Mengeneinheit', data.Mengeneinheit);
  setVal('f-Mindestlagermenge', data.Mindestlagermenge);
  setVal('f-Termin', data.Termin ? String(data.Termin).slice(0,10) : '');
  setVal('f-Lieferant', data.Lieferant);
  setVal('f-Lieferant2', data.Lieferant2);
  setVal('f-Lieferant3', data.Lieferant3);
  setVal('f-Lieferant4', data.Lieferant4);
  setVal('f-GeschaetzterPreis', data.GeschaetzterPreis);
  setVal('f-Kostenstelle', data.Kostenstelle);
  // Beschaffungslogik: Radio + Zusatz-Karten
  _applyBeschaffungPrefill(data.Beschaffungslogik, 'Beschaffungslogik', 'beschaffungslogik-extra-cards');
  // Zusatz-Lieferanten einblenden, wenn befüllt
  [2,3,4].forEach(n => {
    if (data['Lieferant' + n]) { const grp = $id('lieferant-extra-' + n); if (grp) grp.style.display = ''; }
  });
  // Preiszeile aus der Vorlage ableiten: eine Zeile mit Menge + Einzelpreis
  // (= Gesamtwert ÷ Menge). Bestellvolumen rechnet sich daraus neu; eine geänderte
  // Menge ergibt sofort einen neuen Gesamtwert.
  const mengeP = _ezNum(data.Menge);
  const totalP = _ezNum(data.GeschaetzterPreis);
  const epP    = mengeP > 0 && totalP > 0 ? Math.round((totalP / mengeP) * 100) / 100 : 0;
  einzelLines = [{ stueck: mengeP > 0 ? data.Menge : '', preis: epP > 0 ? epP : '' }];
  renderEinzelLines();
}

function initWizard() {
  wizardData = {};
  showStep(1);
  // Artikelnummer + Kostenstelle autocomplete in wizard
  const wArtNr = $id('f-Artikelnummer');
  if (wArtNr) initTidAutocomplete(wArtNr, key => $id('f-' + key));
  initKostenstAuto($id('f-Kostenstelle'));
  // Show field availability hints
  for (const fd of FORM_FIELDS) {
    const hint = $id('hint-' + fd.key);
    if (!hint) continue;
    if (resolvedFields[fd.key]) {
      hint.textContent = '';
      hint.className = 'field-hint';
    } else if (fd.key !== 'Title') {
      hint.textContent = '⚠ Spalte „' + fd.label + '" existiert noch nicht in der SP-Liste – wird beim Speichern übersprungen.';
      hint.className = 'field-hint warn';
    }
  }
  // Reset fields
  ['Title','Beschreibung','Warengruppe','Mengeneinheit',
   'Mindestlagermenge','Termin','Artikelnummer','ExterneArtikelnummer','AutoBedarfDatum',
   'Lieferant','Lieferant2','Lieferant3','Lieferant4',
   'GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  // Mengen×Einzelpreis-Zeilen zurücksetzen (Bestellvolumen bleibt schreibgeschützt)
  einzelLines = [{ stueck: '', preis: '' }];
  renderEinzelLines();
  const prioEl = $id('f-Prioritaet');
  if (prioEl) {
    // SP populate-code may have replaced HTML options with SP choice values (e.g. "HIGH").
    // Find a "Normal"-like option case-insensitively; if not found, pick first non-empty option.
    const prioOpts = [...prioEl.options];
    const normalOpt = prioOpts.find(o => /^(normal|standard)$/i.test(o.value.trim()));
    if (normalOpt) prioEl.value = normalOpt.value;
    else {
      const first = prioOpts.find(o => o.value !== '');
      prioEl.value = first ? first.value : '';
    }
  }
  const mengeEl = $id('f-Menge'); if (mengeEl) mengeEl.value = '1'; // Standardwert
  // Reset submit button in case previous submission left it in a loading state
  const submitBtn = $id('btn-submit');
  if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = '✓ Anfrage einreichen'; }
  const firstRadio = document.querySelector('input[name=Beschaffungslogik]');
  if (firstRadio) firstRadio.checked = true;
  document.querySelectorAll('#beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  // Reset wizard attachments
  wizardFilesArr = [];
  renderWizardFiles();
  const wSec = $id('wizard-attachments');
  if (wSec) wSec.style.display = 'none';
  const wNote = $id('attach-required-note');
  if (wNote) wNote.textContent = '';
  initWizardDrop();
  const lbaEl = $id('f-LeadBuyerAbschluss');
  if (lbaEl) lbaEl.checked = false;
  // Reset progressive Lieferant disclosure
  [2,3,4].forEach(n => {
    const grp = $id('lieferant-extra-' + n);
    if (grp) grp.style.display = 'none';
  });
  // Restore "+" button on Lieferant 1 if it was removed
  const lief1 = $id('f-Lieferant')?.closest('.form-group');
  if (lief1 && !$id('btn-add-lieferant-1')) {
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'btn-add-lief';
    btn.id = 'btn-add-lieferant-1'; btn.onclick = () => addLieferant(2);
    btn.textContent = '+ Weiteren Lieferanten hinzufügen';
    lief1.appendChild(btn);
  }
  $id('preis-route-hint').style.display = 'none';
  // Favoriten-Leiste + ggf. ausstehende Wiedervorlage anwenden
  renderFavBar('single');
  if (pendingPrefill && pendingPrefill.type === 'single') {
    applyWizardPrefill(pendingPrefill.data);
    toast('Wiedervorlage übernommen – bitte prüfen und absenden.', 'success');
  }
  pendingPrefill = null;
}

function showStep(n) {
  [1,2,3,4].forEach(i => {
    $id('wstep-'+i).classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-step="${i}"]`);
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
  // Beim Wechsel auf die Beschaffungsdetails: erste Preiszeile mit der Menge aus
  // Schritt 1 vorbelegen (Komfort), dann Zeilen rendern + Bestellvolumen berechnen.
  if (n === 3) {
    if (einzelLines.length === 1 && !String(einzelLines[0].stueck || '').trim()) {
      const m = ($id('f-Menge')?.value || '').trim();
      if (m) einzelLines[0].stueck = m;
    }
    renderEinzelLines();
  }
}

function wNext(step) {
  if (step === 1) {
    const title = $id('f-Title').value.trim();
    const wg    = $id('f-Warengruppe').value;
    if (!title) { toast('Bitte Bezeichnung angeben.', 'error'); return; }
    if (!wg)    { toast('Bitte Warengruppe wählen.', 'error'); return; }
    wizardData.step1 = {
      Title:                title,
      Beschreibung:         $id('f-Beschreibung').value.trim(),
      Warengruppe:          wg,
      Prioritaet:           $id('f-Prioritaet').value,
      Artikelnummer:        $id('f-Artikelnummer').value.trim(),
      ExterneArtikelnummer: $id('f-ExterneArtikelnummer')?.value.trim() || '',
      AutoBedarfDatum:      $id('f-AutoBedarfDatum')?.value || '',
    };
  } else if (step === 2) {
    const menge  = $id('f-Menge').value;
    const me     = $id('f-Mengeneinheit').value;
    const termin = $id('f-Termin').value;
    if (!menge || parseFloat(menge) <= 0) { toast('Bitte gültige Menge eingeben.', 'error'); return; }
    if (!me)     { toast('Bitte Mengeneinheit wählen.', 'error'); return; }
    if (!termin) { toast('Bitte Benötigt-bis-Datum angeben.', 'error'); return; }
    wizardData.step2 = {
      Menge:             menge,
      Mengeneinheit:     me,
      Mindestlagermenge: $id('f-Mindestlagermenge').value || null,
      Termin:            termin,
    };
  } else if (step === 3) {
    const lieferant = $id('f-Lieferant').value.trim();
    if (!lieferant) { toast('Bitte mindestens Lieferant 1 angeben.', 'error'); return; }
    const extraSelected = document.querySelector('#beschaffungslogik-extra-cards .check-card.selected');
    if (!extraSelected) { toast('Bitte unter „Zusätzlich kombinierbar" eine Option auswählen.', 'error'); return; }
    const total    = einzelLines.reduce((s, ln) => s + _ezNum(ln.stueck) * _ezNum(ln.preis), 0);
    const totalQty = einzelLines.reduce((s, ln) => s + _ezNum(ln.stueck), 0);
    if (!(total > 0)) { toast('Bitte mindestens eine Zeile mit Menge und Einzelpreis ausfüllen.', 'error'); return; }
    wizardData.step3 = {
      Beschaffungslogik: [
        document.querySelector('input[name=Beschaffungslogik]:checked')?.value || '',
        ...[...document.querySelectorAll('#beschaffungslogik-extra-cards .check-card.selected')].map(c => c.dataset.value)
      ].filter(Boolean).join(', '),
      Lieferant:         $id('f-Lieferant').value.trim(),
      Lieferant2:        $id('f-Lieferant2').value.trim(),
      Lieferant3:        $id('f-Lieferant3').value.trim(),
      Lieferant4:        $id('f-Lieferant4').value.trim(),
      GeschaetzterPreis:  Math.round(total * 100) / 100,
      Kostenstelle:       $id('f-Kostenstelle').value.trim(),
      LeadBuyerAbschluss: $id('f-LeadBuyerAbschluss')?.checked ?? false,
    };
    // Gesamtmenge aus den Zeilen übernehmen (überschreibt die Menge aus Schritt 2)
    if (totalQty > 0) wizardData.step3.Menge = String(totalQty);
    buildReview();
  }
  showStep(step + 1);
}

function toggleBLExtra(el) {
  const wasSelected = el.classList.contains('selected');
  document.querySelectorAll('#beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  if (!wasSelected) el.classList.add('selected');
}

function angeboteAnzahl(gesamt) {
  if (!gesamt || gesamt <= 500)  return 0;
  if (gesamt <= 2500)            return 2;
  if (gesamt <= 5000)            return 3;
  if (gesamt <= 10000)           return 3;
  if (gesamt <= 50000)           return 4;
  return 5;
}

function wBack(step) { showStep(step - 1); }

function addLieferant(n) {
  if (n > 4) return;
  const grp = $id('lieferant-extra-' + n);
  if (grp) grp.style.display = 'block';
  $id('btn-add-lieferant-' + (n - 1))?.remove();
  $id('f-Lieferant' + n)?.focus();
}

// ── EINZELANFRAGE: MENGEN × EINZELPREIS-ZEILEN ───────────────────────────────
// Mehrere (Menge × Einzelpreis)-Zeilen; das Bestellvolumen wird daraus summiert
// und ist schreibgeschützt ("nicht mehr beschreibbar"). Greift in JEDEM Fall –
// auch aus Wiedervorlage/Auto-Entwurf (dort als eine Zeile: Menge + Gesamtwert÷Menge).
function renderEinzelLines() {
  const wrap = $id('einzel-lines');
  if (!wrap) return;
  if (!einzelLines.length) einzelLines.push({ stueck: '', preis: '' });
  wrap.innerHTML = einzelLines.map((ln, i) => {
    const sum = _ezNum(ln.stueck) * _ezNum(ln.preis);
    const del = einzelLines.length > 1
      ? `<button type="button" class="ezl-del" title="Zeile entfernen" onclick="einzelLineRemove(${i})">✕</button>`
      : `<span class="ezl-del-ph"></span>`;
    return `<div class="ezl-row">
      <input type="number" class="ezl-stueck" min="0" step="any" placeholder="Menge"
        value="${esc(String(ln.stueck ?? ''))}" oninput="einzelLineChange(${i},'stueck',this.value)"/>
      <span class="ezl-x">×</span>
      <input type="number" class="ezl-preis" min="0" step="0.01" placeholder="Einzelpreis €"
        value="${esc(String(ln.preis ?? ''))}" oninput="einzelLineChange(${i},'preis',this.value)"/>
      <span class="ezl-eq" data-idx="${i}">= ${fmtEuro(sum)}</span>
      ${del}
    </div>`;
  }).join('');
  recalcEinzelPreis();
}

function einzelLineAdd()    { einzelLines.push({ stueck: '', preis: '' }); renderEinzelLines(); }
function einzelLineRemove(i){ einzelLines.splice(i, 1); if (!einzelLines.length) einzelLines.push({ stueck: '', preis: '' }); renderEinzelLines(); }
function einzelLineChange(i, field, val) {
  if (!einzelLines[i]) return;
  einzelLines[i][field] = val;
  // Zeilensumme live aktualisieren (kein Re-Render → Fokus bleibt erhalten)
  const eq = document.querySelector(`#einzel-lines .ezl-eq[data-idx="${i}"]`);
  if (eq) eq.textContent = '= ' + fmtEuro(_ezNum(einzelLines[i].stueck) * _ezNum(einzelLines[i].preis));
  recalcEinzelPreis();
}

function recalcEinzelPreis() {
  const totalEl = $id('f-GeschaetzterPreis');
  if (!totalEl) return;
  const total = einzelLines.reduce((s, ln) => s + _ezNum(ln.stueck) * _ezNum(ln.preis), 0);
  const rounded = Math.round(total * 100) / 100;
  totalEl.value    = rounded > 0 ? rounded : '';
  totalEl.readOnly = true;
  totalEl.classList.add('input-computed');
  const note = $id('bv-auto-note');
  if (note) {
    const qty = einzelLines.reduce((s, ln) => s + _ezNum(ln.stueck), 0);
    note.textContent = rounded > 0
      ? `Σ aus ${einzelLines.length} Zeile(n) · Menge gesamt ${qty}`
      : 'automatisch berechnet';
  }
  updatePreisHint();
}

function updatePreisHint() {
  const preis = parseFloat($id('f-GeschaetzterPreis').value) || 0;
  const lba   = $id('f-LeadBuyerAbschluss')?.checked ?? false;
  const hint  = $id('preis-route-hint');
  if (preis > 0) {
    hint.textContent = genehmigungsweg(preis, lba);
    hint.style.display = 'block';
  } else {
    hint.style.display = 'none';
  }
  // Show/hide attachment section and update required-count note
  const req     = lba ? 0 : angeboteAnzahl(preis);
  const section = $id('wizard-attachments');
  const note    = $id('attach-required-note');
  if (!section) return;
  section.style.display = 'block';
  if (note) note.textContent = req === 0 ? '(optional)' : `mind. ${req} erforderlich gemäß Regelwerk`;
}

// ── WIZARD DRAG-AND-DROP ATTACHMENTS ─────────────────────────────────────────
function initWizardDrop() {
  const zone  = $id('wizard-drop-zone');
  const input = $id('wizard-file-input');
  if (!zone || !input) return;
  // Guard: only attach listeners once (called on every wizard reset)
  if (zone._dropReady) return;
  zone._dropReady = true;
  zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    addWizardFiles([...e.dataTransfer.files]);
  });
  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => { addWizardFiles([...input.files]); input.value = ''; });
}

function addWizardFiles(files) {
  for (const f of files) {
    if (wizardFilesArr.length >= 5) break;
    if (!/\.pdf$/i.test(f.name)) { toast('Nur PDF-Dateien sind erlaubt.', 'error'); continue; }
    if (wizardFilesArr.some(x => x.name === f.name && x.size === f.size)) continue; // skip duplicate
    wizardFilesArr.push(f);
  }
  renderWizardFiles();
}

function removeWizardFile(idx) {
  wizardFilesArr.splice(idx, 1);
  renderWizardFiles();
}

function renderWizardFiles() {
  const list = $id('wizard-file-list');
  if (!list) return;
  list.innerHTML = wizardFilesArr.map((f, i) =>
    `<div class="wizard-file-item">
      <span class="wf-icon">📎</span>
      <span class="wf-name">${esc(f.name)}</span>
      <span class="wf-size">${(f.size/1024).toFixed(0)} KB</span>
      <button class="wf-remove" onclick="removeWizardFile(${i})" title="Entfernen">✕</button>
    </div>`
  ).join('');
  const zone = $id('wizard-drop-zone');
  if (zone) zone.classList.toggle('has-files', wizardFilesArr.length > 0);
}

function genehmigungsweg(gesamt, leadBuyer = false) {
  let stufe, freigeber, angebote;
  if (gesamt <= 500) {
    stufe = 1; freigeber = ['Anforderer Fachabt.', 'Einkäufer Werk']; angebote = 'kein Angebot nötig';
  } else if (gesamt <= 2500) {
    stufe = 2; freigeber = ['Einkäufer Werk', 'Werkleiter'];           angebote = 'mind. 2 Angebote';
  } else if (gesamt <= 5000) {
    stufe = 2; freigeber = ['Einkäufer Werk', 'Werkleiter'];           angebote = 'mind. 2–3 Angebote';
  } else if (gesamt <= 10000) {
    stufe = 3; freigeber = ['Werkleiter', 'Einkaufsleitung Holding'];  angebote = 'mind. 3 Angebote';
  } else if (gesamt <= 50000) {
    stufe = 3; freigeber = ['Werkleiter', 'Einkaufsleitung Holding'];  angebote = 'mind. 3–4 Angebote';
  } else {
    stufe = 4; freigeber = ['Einkaufsleitung Holding', 'GF kfm. Leitung Holding'];
    angebote = 'Europ. Ausschreibung (mind. 5 Angebote)';
  }
  // Angebotspflicht: ab Stufe 2 entfällt sie beim Lead-Buyer-Abschluss.
  let angeboteText = angebote;
  if (stufe > 1) angeboteText += leadBuyer ? ' ✓ entfällt' : ' (entfällt bei Lead-Buyer-Abschluss)';
  // Lead-Buyer-Abschluss macht die Holding-Freigabe (Einkaufsleitung Holding) überflüssig.
  let fg = freigeber;
  if (leadBuyer) {
    const filtered = freigeber.filter(f => !/Einkaufsleitung Holding/i.test(f));
    if (filtered.length) fg = filtered;
  }
  return `Stufe ${stufe} · ${fmtEuro(gesamt)} · ${angeboteText} · Freigabe: ${fg.join(' + ')}`;
}

function buildReview() {
  const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
  const missingCols = FORM_FIELDS
    .filter(fd => fd.key !== 'Title' && !resolvedFields[fd.key] && d[fd.key])
    .map(fd => fd.label);

  $id('review-content').innerHTML = `
    <div class="review-grid">
      ${reviewSection('Bedarf', [
        ['Bezeichnung', d.Title],
        ['Beschreibung', d.Beschreibung],
        ['Warengruppe', d.Warengruppe],
        ['Priorität', d.Prioritaet],
      ])}
      ${reviewSection('Mengengaben', [
        ['Menge', d.Menge && d.Mengeneinheit ? `${d.Menge} ${d.Mengeneinheit}` : d.Menge],
        ['Mindestlagermenge', d.Mindestlagermenge],
        ['Benötigt bis', d.Termin ? fmtDate(d.Termin) : null],
      ])}
      ${reviewSection('Beschaffungsdetails', [
        ['Beschaffungsart', d.Beschaffungslogik],
        ['Artikelnummer', d.Artikelnummer],
        ['Lieferant 1', d.Lieferant],
        ['Lieferant 2', d.Lieferant2],
        ['Lieferant 3', d.Lieferant3],
        ['Lieferant 4', d.Lieferant4],
        ['Mengen × Einzelpreise', einzelLines.filter(l => _ezNum(l.stueck) && _ezNum(l.preis)).map(l => `${l.stueck} × ${fmtEuro(_ezNum(l.preis))}`).join('  ·  ') || null],
        ['Bestellvolumen in €', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
        ['Kostenstelle', d.Kostenstelle],
      ])}
    </div>
    ${d.GeschaetzterPreis ? `<div class="info-box info" style="margin-top:12px">${genehmigungsweg(d.GeschaetzterPreis, d.LeadBuyerAbschluss === true || /^(true|ja|yes|1)$/i.test(String(d.LeadBuyerAbschluss || '')))}</div>` : ''}
    <div class="info-box info" style="margin-top:10px">
      Nach dem Einreichen wird <strong>Power Automate</strong> automatisch den Genehmigungsprozess starten und den zuständigen Genehmiger per E-Mail / Teams benachrichtigen.
    </div>`;

  const mw = $id('missing-cols-warn');
  if (missingCols.length) {
    mw.innerHTML = `⚠ Folgende Felder werden <strong>nicht gespeichert</strong> (Spalten fehlen in SP-Liste): <em>${missingCols.join(', ')}</em>`;
    mw.style.display = 'block';
  } else {
    mw.style.display = 'none';
  }
}

function reviewSection(title, rows) {
  const content = rows.filter(([,v]) => v).map(([l,v]) =>
    `<div class="review-row"><span class="review-label">${esc(l)}</span><span class="review-value">${esc(String(v))}</span></div>`
  ).join('');
  if (!content) return '';
  return `<div class="review-section"><h3>${title}</h3>${content}</div>`;
}

// ── SUBMIT ───────────────────────────────────────────────────────────────────
const NUMBER_FIELDS = new Set(['GeschaetzterPreis','Menge','Mindestlagermenge','TatsaechlicherPreis']);
const DATE_FIELDS   = new Set(['Termin','Lieferdatum','AutoBedarfDatum']);
const BOOL_FIELDS   = new Set(['LeadBuyerAbschluss']);

// SP Graph requires ISO 8601 for DateTime columns: "2026-05-10T00:00:00Z"
// "Date Only" columns want just the date: "2026-05-10"
// <input type="date"> yields "2026-05-10" → append time only if column is dateTime
function toSpDate(val, spColName) {
  if (!val) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(val)) return val; // already has time component
  const col = spColName ? colByKey[spColName] : null;
  const isDateOnly = col?.dateTime?.format === 'dateOnly';
  return isDateOnly ? val : val + 'T00:00:00Z';
}

function buildFields(data, fieldDefs) {
  const fields = {};
  for (const fd of fieldDefs) {
    const spCol = resolvedFields[fd.key];
    const val   = data[fd.key];
    if (BOOL_FIELDS.has(fd.key)) {
      // Only send when explicitly set — avoids sending false to SP when user never touched the field.
      if (spCol && val != null && val !== '') fields[spCol] = Boolean(val);
      continue;
    }
    if (!spCol || val === null || val === undefined || val === '') continue;
    if (DATE_FIELDS.has(fd.key)) {
      const d = toSpDate(val, spCol);
      if (d) fields[spCol] = d;
    } else if (NUMBER_FIELDS.has(fd.key)) {
      const n = parseFloat(val);
      // Always send as JS number — Graph accepts numbers for all numeric/currency columns.
      // Sending String(n) to a Currency column causes 404 "itemNotFound" in some Graph versions.
      if (!isNaN(n)) fields[spCol] = n;
    } else {
      fields[spCol] = val;
    }
  }
  return fields;
}

// Retry POST removing offending fields until SP accepts the payload.
// Handles 400 "not recognized", generic 400 "invalidRequest", and 500 "generalException".
async function postRetry(path, fields) {
  const skipped = [];
  // Drop order: optional fields first, then non-Title required fields
  const optionalKeys = Object.entries(resolvedFields)
    .filter(([k]) => !FORM_FIELDS.find(f => f.key === k && f.required))
    .map(([, v]) => v).filter(Boolean);
  const requiredKeys = Object.entries(resolvedFields)
    .filter(([k]) => FORM_FIELDS.find(f => f.key === k && f.required && f.key !== 'Title'))
    .map(([, v]) => v).filter(Boolean);
  const dropQueue = [...optionalKeys, ...requiredKeys];

  for (let i = 0; i < dropQueue.length + 5; i++) {
    try {
      const newItem = await gPost(path, { fields });
      return { skipped, newItem };
    } catch(e) {
      // 400 with named field: SP says exactly which field is wrong
      const m400 = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (m400) {
        const bad = m400[1];
        skipped.push(bad);
        delete fields[bad];
        for (const [k, v] of Object.entries(resolvedFields)) {
          if (v === bad) resolvedFields[k] = null;
        }
        continue;
      }
      // 500 or generic 400 "invalidRequest": SP won't say which field is wrong
      // → drop the next field from the queue and retry
      const isRetryable = e.message.includes('500') || e.message.includes('generalException') ||
        e.message.includes('invalidRequest');
      if (isRetryable) {
        const dropKey = dropQueue.find(k => k && fields[k] !== undefined);
        if (dropKey) {
          const code = e.message.includes('400') ? '400' : '500';
          console.warn(`[postRetry] ${code} – dropping field:`, dropKey, 'value:', fields[dropKey]);
          skipped.push(dropKey + `(${code})`);
          delete fields[dropKey];
          continue;
        }
      }
      throw e;
    }
  }
  throw new Error('Einreichen fehlgeschlagen – zu viele problematische Felder.');
  return { skipped: [], newItem: null }; // unreachable, satisfies linters
}

// patchRetry: updates item fields with 404-wait (SP propagation) + 400/500 field-dropping.
// On generic 400 "invalidRequest" SP won't say which field is wrong → drop fields one by
// one (optional first, required last) until the PATCH succeeds. Returns array of dropped keys.
async function patchRetry(path, fields, retries404 = 6, delay404 = 2000) {
  const skipped = [];
  // Build drop queue: optional fields first, then required (excluding Title equivalents)
  const optKeys = Object.entries(resolvedFields)
    .filter(([k]) => !FORM_FIELDS.find(f => f.key === k && f.required))
    .map(([, v]) => v).filter(v => v && fields[v] !== undefined);
  const reqKeys = Object.entries(resolvedFields)
    .filter(([k]) => FORM_FIELDS.find(f => f.key === k && f.required && f.key !== 'Title'))
    .map(([, v]) => v).filter(v => v && fields[v] !== undefined);
  const dropQueue = [...optKeys, ...reqKeys];

  let notFound404s = 0;
  for (let attempt = 0; attempt < retries404 + dropQueue.length + 5; attempt++) {
    try {
      await gPatch(path, fields);
      if (skipped.length) console.warn('[patchRetry] succeeded after dropping:', skipped);
      return skipped;
    } catch(e) {
      const msg = e.message || '';
      // 404 / itemNotFound: item not yet propagated → wait and retry
      if ((msg.includes('404') || msg.includes('itemNotFound')) && !e._noRetry) {
        if (notFound404s < retries404) { notFound404s++; await new Promise(r => setTimeout(r, delay404)); continue; }
      }
      // Named field error: SP says exactly which field
      const mField = msg.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (mField) {
        const bad = mField[1];
        console.warn('[patchRetry] SP rejected field:', bad);
        skipped.push(bad); delete fields[bad];
        for (const [k, v] of Object.entries(resolvedFields)) { if (v === bad) resolvedFields[k] = null; }
        continue;
      }
      // Generic 400 "invalidRequest" / 500: drop next candidate and retry
      const isRetryable = msg.includes('invalidRequest') || msg.includes('400') ||
        msg.includes('500') || msg.includes('generalException');
      if (isRetryable) {
        const dropKey = dropQueue.find(k => fields[k] !== undefined);
        if (dropKey) {
          console.warn('[patchRetry] 400/500 → dropping field:', dropKey, '=', JSON.stringify(fields[dropKey]));
          skipped.push(dropKey); delete fields[dropKey]; continue;
        }
      }
      throw e;
    }
  }
  throw new Error('PATCH fehlgeschlagen – alle Kandidatenfelder versucht.');
}

async function submitRequest() {
  const btn = $id('btn-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';

  try {
    const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
    const allFields = buildFields(d, FORM_FIELDS);

    if (!allFields['Title']) { toast('Titel fehlt.', 'error'); btn.disabled=false; btn.textContent='✓ Anfrage einreichen'; return; }

    const wizardFiles = [...wizardFilesArr];

    if (!siteId || !listId) await discoverSP();

    // Use the list GUID (listId) instead of the display name — more reliable, avoids
    // name-resolution overhead and potential encoding issues in Graph API paths.
    const listPath = `/sites/${siteId}/lists/${listId}`;

    // ── Schritt 1: Element mit nur Title anlegen (minimal, kann nicht an Feldern scheitern) ──
    btn.textContent = 'Anfrage wird angelegt…';
    const newItem = await gPost(`${listPath}/items`, { fields: { Title: allFields['Title'] } });
    const itemId  = newItem.id;

    // ── Schritt 2: Restliche Felder per PATCH nachpflegen ──
    // SP Online has a propagation delay: the new item may not be addressable via PATCH
    // immediately after creation. Retry with generous delay (6 × 2 s = up to 12 s total).
    const patchFields = { ...allFields };
    delete patchFields['Title'];
    if (Object.keys(patchFields).length) {
      btn.textContent = 'Felder werden gespeichert…';
      try {
        const skipped = await patchRetry(`${listPath}/items/${itemId}/fields`, patchFields);
        if (skipped.length) console.warn('[submitRequest] übersprungene Felder (SP ablehnte):', skipped);
      } catch(patchErr) {
        console.warn('[submitRequest] PATCH endgültig fehlgeschlagen:', patchErr.message);
        toast('Anfrage erstellt, aber einige Felder konnten nicht gespeichert werden.', 'error');
      }
    }

    // ── Schritt 3: Anhänge hochladen ──
    if (wizardFiles.length) {
      btn.textContent = 'Anhänge werden hochgeladen…';
      let tok;
      try { tok = await getSpToken(); }
      catch(tokErr) {
        console.warn('[submitRequest] SP-Token:', tokErr);
        toast('Anfrage erstellt. Anhänge konnten nicht hochgeladen werden (Token-Fehler).', 'error');
        await loadItems(false);
        navigate('mine');
        return;
      }
      const errors = [];
      for (const file of wizardFiles) {
        try {
          // Read the file buffer ONCE before entering the retry loop so a
          // second attempt doesn't try to re-read a potentially consumed stream.
          const buf = await file.arrayBuffer();
          // SP REST may not yet have propagated the new item (same lag as PATCH 404).
          // Wrap in retryOn404 so we wait up to ~4.8s before giving up.
          // Filename: single-quote chars must be escaped as %27 inside the OData string literal.
          const fname = file.name.replace(/'/g, "''");
          const fnameUrl = encodeURIComponent(fname);
          const uploadUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${fnameUrl}')`;
          await retryOn404(async () => {
            const r = await fetch(uploadUrl, {
              method: 'POST',
              headers: {
                Authorization: 'Bearer ' + tok,
                Accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/octet-stream',
              },
              body: buf,
            });
            if (!r.ok) {
              const txt = await r.text().catch(() => '');
              const err = new Error(`${r.status}: ${txt}`);
              // Only retry on 404; other errors are immediate failures
              if (r.status !== 404) err._noRetry = true;
              throw err;
            }
          }, 3, 1200);
        } catch(e) { console.warn('[submitRequest] Anhang:', e); errors.push(`${file.name} (${e.message})`); }
      }
      if (errors.length) toast(`Anfrage erstellt. Anhänge fehlgeschlagen: ${errors.join(', ')}`, 'error');
      else toast(`Anfrage eingereicht! ${wizardFiles.length} Angebot(e) angehängt.`, 'success');
    } else {
      toast('Anfrage eingereicht! Power Automate startet den Genehmigungsprozess.', 'success');
    }

    await loadItems(false);
    navigate('mine');
  } catch(e) {
    toast('Fehler: ' + e.message, 'error');
    btn.disabled = false;
    btn.textContent = '✓ Anfrage einreichen';
  }
}

// ── TABELLEN-IMPORTER ────────────────────────────────────────────────────────

function initImporter() {
  renderImporter();
}

function downloadCSV(csvText, filename) {
  const bom  = '﻿'; // UTF-8 BOM for Excel
  const blob = new Blob([bom + csvText], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function exportTidCSV() {
  const rows = ['Artikelnummer;Bezeichnung;Warengruppe'];
  for (const [nr, h] of Object.entries(TID_MAP_ACTIVE)) {
    rows.push(`${nr};${(h.b||'').replace(/;/g,'')};${(h.w||'').replace(/;/g,'')}`);
  }
  downloadCSV(rows.join('\r\n'), 'Artikelnummern_Export.csv');
}

function exportKstCSV() {
  const rows = ['Kostenstelle;Bezeichnung'];
  for (const e of KOSTENST_DATA) {
    rows.push(`${e.nr};${(e.label||'').replace(/;/g,'')}`);
  }
  downloadCSV(rows.join('\r\n'), 'Kostenstellen_Export.csv');
}

function downloadTidTemplate() {
  downloadCSV('Artikelnummer;Bezeichnung;Warengruppe\r\n1234-00001;Beispielartikel;Ersatzteile', 'Artikelnummern_Vorlage.csv');
}

function downloadKstTemplate() {
  downloadCSV('Kostenstelle;Bezeichnung\r\n10011;Fertigungsleitung', 'Kostenstellen_Vorlage.csv');
}

function renderImporter() {
  const tidCount  = Object.keys(TID_MAP_ACTIVE).length;
  const kstCount  = KOSTENST_DATA.length;
  const tidIsCustom = !!localStorage.getItem(LS_TID_KEY);
  const kstIsCustom = !!localStorage.getItem(LS_KOSTENST_KEY);

  $id('view-importer').innerHTML = `
    <div style="max-width:700px;margin:0 auto">
      <p style="color:#6b7280;margin-bottom:20px;font-size:.9rem">
        Laden Sie CSV- oder XLSX-Dateien hoch, um die Artikelnummer- und Kostenstellentabellen zu aktualisieren.
        Die Daten werden lokal im Browser gespeichert.
      </p>

      <!-- Artikelnummern -->
      <div class="imp-card">
        <div class="imp-card-header">
          <div>
            <div class="imp-card-title">📦 Artikelnummern (TID-Tabelle)</div>
            <div class="imp-card-sub">${tidCount} Einträge geladen${tidIsCustom ? ' <span class="imp-badge-custom">Angepasst</span>' : ' <span class="imp-badge-default">Standard</span>'}</div>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <button class="btn btn-sm btn-outline" onclick="exportTidCSV()" title="Aktuelle Tabelle als CSV herunterladen">⬇ Export</button>
            <button class="btn btn-sm btn-ghost" onclick="downloadTidTemplate()" title="Leere Vorlage herunterladen">📄 Vorlage</button>
            ${tidIsCustom ? `<button class="btn btn-sm btn-ghost" onclick="resetTable('tid')">↺ Zurücksetzen</button>` : ''}
          </div>
        </div>
        <div class="imp-format-hint">
          Format: CSV mit Semikolon <code>Artikelnummer;Bezeichnung;Warengruppe</code><br>
          oder XLSX mit denselben Spalten in Zeile 1. Bestehende Daten werden durch Import <strong>ergänzt</strong> (gleiche Nr. wird überschrieben).
        </div>
        <div class="imp-drop-zone" id="imp-drop-tid"
          ondragover="event.preventDefault();this.classList.add('drag-over')"
          ondragleave="this.classList.remove('drag-over')"
          ondrop="handleImpDrop(event,'tid')">
          <span>📂 CSV oder XLSX hier ablegen oder</span>
          <label class="btn btn-sm btn-outline" style="cursor:pointer;margin-left:6px">
            Datei auswählen
            <input type="file" accept=".csv,.xlsx" style="display:none" onchange="handleImpFile(this,'tid')">
          </label>
        </div>
        <div id="imp-result-tid" class="imp-result"></div>
      </div>

      <!-- Kostenstellen -->
      <div class="imp-card" style="margin-top:16px">
        <div class="imp-card-header">
          <div>
            <div class="imp-card-title">🏢 Kostenstellen</div>
            <div class="imp-card-sub">${kstCount} Einträge geladen${kstIsCustom ? ' <span class="imp-badge-custom">Angepasst</span>' : ' <span class="imp-badge-default">Standard</span>'}</div>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <button class="btn btn-sm btn-outline" onclick="exportKstCSV()" title="Aktuelle Tabelle als CSV herunterladen">⬇ Export</button>
            <button class="btn btn-sm btn-ghost" onclick="downloadKstTemplate()" title="Leere Vorlage herunterladen">📄 Vorlage</button>
            ${kstIsCustom ? `<button class="btn btn-sm btn-ghost" onclick="resetTable('kst')">↺ Zurücksetzen</button>` : ''}
          </div>
        </div>
        <div class="imp-format-hint">
          Format: CSV mit Semikolon <code>Kostenstelle;Bezeichnung</code> oder eine Spalte <code>10011 Fertigungsleitung</code><br>
          oder XLSX mit denselben Spalten. Hochladen <strong>ersetzt</strong> die gesamte Tabelle.
        </div>
        <div class="imp-drop-zone" id="imp-drop-kst"
          ondragover="event.preventDefault();this.classList.add('drag-over')"
          ondragleave="this.classList.remove('drag-over')"
          ondrop="handleImpDrop(event,'kst')">
          <span>📂 CSV oder XLSX hier ablegen oder</span>
          <label class="btn btn-sm btn-outline" style="cursor:pointer;margin-left:6px">
            Datei auswählen
            <input type="file" accept=".csv,.xlsx" style="display:none" onchange="handleImpFile(this,'kst')">
          </label>
        </div>
        <div id="imp-result-kst" class="imp-result"></div>
      </div>
    </div>`;
}

function handleImpDrop(e, type) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over');
  const file = e.dataTransfer?.files?.[0];
  if (file) parseImportFile(file, type);
}

function handleImpFile(input, type) {
  const file = input.files?.[0];
  if (file) parseImportFile(file, type);
  input.value = '';
}

// SheetJS (xlsx) erst bei Bedarf nachladen – spart ~900 KB beim Start
let _xlsxPromise = null;
function ensureXLSX() {
  if (typeof XLSX !== 'undefined') return Promise.resolve();
  if (_xlsxPromise) return _xlsxPromise;
  _xlsxPromise = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    s.onload  = () => resolve();
    s.onerror = () => { _xlsxPromise = null; reject(new Error('SheetJS konnte nicht geladen werden')); };
    document.head.appendChild(s);
  });
  return _xlsxPromise;
}

async function parseImportFile(file, type) {
  const resultEl = $id('imp-result-' + type);
  resultEl.textContent = '⏳ Wird verarbeitet…';
  try {
    const ext = file.name.split('.').pop().toLowerCase();
    let rows = []; // [{col0, col1, col2, ...}]

    if (ext === 'csv') {
      const text = await file.text();
      const lines = text.split(/\r?\n/).filter(l => l.trim());
      rows = lines.map(l => l.split(/;|\t/).map(c => c.trim().replace(/^"|"$/g, '')));
    } else if (ext === 'xlsx') {
      await ensureXLSX().catch(() => {}); // bei Fehlschlag greift der eingebaute Fallback-Parser
      rows = await parseXlsxFile(file);
    } else {
      resultEl.innerHTML = '<span style="color:#ef4444">❌ Nur CSV oder XLSX-Dateien unterstützt.</span>';
      return;
    }

    if (type === 'tid') applyTidImport(rows, resultEl);
    else                applyKstImport(rows, resultEl);

  } catch(e) {
    resultEl.innerHTML = `<span style="color:#ef4444">❌ Fehler: ${esc(e.message)}</span>`;
  }
}

// Parse XLSX in-browser using built-in zip (no external library needed for simple cases)
async function parseXlsxFile(file) {
  // Use SheetJS if available, else fallback to manual parse
  if (typeof XLSX !== 'undefined') {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    return data.map(r => r.map(c => String(c ?? '').trim()));
  }
  // Fallback: manual zip parse (shared strings + sheet1)
  const buf  = await file.arrayBuffer();
  const u8   = new Uint8Array(buf);
  // Minimal zip extractor
  function readZipEntry(u8, name) {
    const enc = new TextEncoder();
    const nameBytes = enc.encode(name);
    for (let i = 0; i < u8.length - 30; i++) {
      if (u8[i] !== 0x50 || u8[i+1] !== 0x4B || u8[i+2] !== 0x03 || u8[i+3] !== 0x04) continue;
      const fnLen = u8[i+26] | (u8[i+27] << 8);
      const exLen = u8[i+28] | (u8[i+29] << 8);
      const fn = new TextDecoder().decode(u8.slice(i+30, i+30+fnLen));
      const dataStart = i + 30 + fnLen + exLen;
      const compSize  = u8[i+18] | (u8[i+19]<<8) | (u8[i+20]<<16) | (u8[i+21]<<24);
      const method    = u8[i+8] | (u8[i+9]<<8);
      if (fn === name) {
        if (method === 0) return new TextDecoder('utf-8').decode(u8.slice(dataStart, dataStart + compSize));
        // deflate — skip; user should use SheetJS for compressed entries
        return null;
      }
    }
    return null;
  }
  const ssXml  = readZipEntry(u8, 'xl/sharedStrings.xml');
  const shXml  = readZipEntry(u8, 'xl/worksheets/sheet1.xml');
  if (!ssXml || !shXml) throw new Error('XLSX-Datei konnte nicht gelesen werden. Bitte als CSV exportieren.');
  const parser = new DOMParser();
  const ssDoc  = parser.parseFromString(ssXml, 'text/xml');
  const shDoc  = parser.parseFromString(shXml, 'text/xml');
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const strings = [...ssDoc.getElementsByTagNameNS(ns, 'si')].map(si =>
    [...si.getElementsByTagNameNS(ns, 't')].map(t => t.textContent || '').join('')
  );
  const result = [];
  for (const row of shDoc.getElementsByTagNameNS(ns, 'row')) {
    const cells = [];
    for (const c of row.getElementsByTagNameNS(ns, 'c')) {
      const t  = c.getAttribute('t');
      const v  = c.getElementsByTagNameNS(ns, 'v')[0]?.textContent || '';
      cells.push(t === 's' ? (strings[parseInt(v)] || '') : v);
    }
    result.push(cells);
  }
  return result;
}

function applyTidImport(rows, resultEl) {
  // Detect header row: if first row looks like a header (non-numeric first cell)
  let start = 0;
  if (rows.length && !/^\d/.test(rows[0][0] || '')) start = 1;
  const imported = {};
  let count = 0;
  for (const row of rows.slice(start)) {
    const nr = String(row[0] || '').trim().toUpperCase();
    const b  = String(row[1] || '').trim();
    const w  = String(row[2] || '').trim();
    if (!nr || !b) continue;
    imported[nr] = { b, w: w || 'Sonstiges', g: w || 'Sonstiges' };
    count++;
  }
  if (!count) { resultEl.innerHTML = '<span style="color:#ef4444">❌ Keine gültigen Zeilen gefunden. Prüfen Sie das Format.</span>'; return; }
  const merged = Object.assign({}, TID_MAP, imported);
  localStorage.setItem(LS_TID_KEY, JSON.stringify(imported)); // store only the delta
  reloadLookupData();
  resultEl.innerHTML = `<span style="color:#16a34a">✓ ${count} Artikelnummern importiert. Gesamt: ${Object.keys(TID_MAP_ACTIVE).length} Einträge.</span>`;
  setTimeout(renderImporter, 800);
}

function applyKstImport(rows, resultEl) {
  let start = 0;
  const firstCell = String(rows[0]?.[0] || '').trim().toLowerCase();
  if (!firstCell || /kostenstelle|nr|nummer/i.test(firstCell)) start = 1;
  const result = [];
  for (const row of rows.slice(start)) {
    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim();
    if (!col0) continue;
    if (col1) {
      // Two-column format: col0=nr, col1=label
      result.push({ nr: col0, label: col1 });
    } else {
      // Single-column: "10011 Fertigungsleitung"
      const m = col0.match(/^(\d+)\s+(.+)$/);
      if (m) result.push({ nr: m[1], label: m[2].trim() });
    }
  }
  if (!result.length) { resultEl.innerHTML = '<span style="color:#ef4444">❌ Keine gültigen Zeilen gefunden. Prüfen Sie das Format.</span>'; return; }
  localStorage.setItem(LS_KOSTENST_KEY, JSON.stringify(result));
  reloadLookupData();
  resultEl.innerHTML = `<span style="color:#16a34a">✓ ${result.length} Kostenstellen importiert.</span>`;
  setTimeout(renderImporter, 800);
}

function resetTable(type) {
  if (type === 'tid')  localStorage.removeItem(LS_TID_KEY);
  if (type === 'kst')  localStorage.removeItem(LS_KOSTENST_KEY);
  reloadLookupData();
  renderImporter();
}

// ── REPORTS & AUSWERTUNGEN ───────────────────────────────────────────────────
let reportFilters = {};
let reportSort    = { col: 'id', dir: 'desc' };

// rf = resolved field reader
function _rf(i, key) { return getField(i, resolvedFields[key] || key) || ''; }

const REPORT_COLS = [
  { key:'id',               label:'ID',             type:'num',  get:i => parseInt(i.id, 10) },
  { key:'Title',            label:'Bezeichnung',    type:'text', get:i => cleanTitle(getField(i,'Title') || '') },
  { key:'Status',           label:'Status',         type:'text', get:i => getStatusVal(i) || 'Eingereicht' },
  { key:'Warengruppe',      label:'Warengruppe',    type:'text', get:i => _rf(i,'Warengruppe') },
  { key:'Prioritaet',       label:'Priorität',      type:'text', get:i => _rf(i,'Prioritaet') },
  { key:'Menge',            label:'Menge',          type:'num',  get:i => { const n = parseFloat(_rf(i,'Menge')); return isNaN(n) ? null : n; } },
  { key:'Mengeneinheit',    label:'ME',             type:'text', get:i => _rf(i,'Mengeneinheit') },
  { key:'GeschaetzterPreis',label:'Volumen (€)',    type:'num',  get:i => { const n = parseFloat(_rf(i,'GeschaetzterPreis')); return isNaN(n) ? null : n; } },
  { key:'Beschaffungslogik',label:'Beschaffungsart',type:'text', get:i => _rf(i,'Beschaffungslogik') },
  { key:'Lieferant',        label:'Lieferant',      type:'text', get:i => _rf(i,'Lieferant') },
  { key:'Kostenstelle',     label:'Kostenstelle',   type:'text', get:i => _rf(i,'Kostenstelle') },
  { key:'Bestellnummer',    label:'Bestell-Nr.',    type:'text', get:i => _rf(i,'Bestellnummer') },
  { key:'Artikelnummer',    label:'Artikel-Nr.',    type:'text', get:i => _rf(i,'Artikelnummer') },
  { key:'Termin',           label:'Benötigt bis',   type:'date', get:i => _rf(i,'Termin') },
  { key:'creator',          label:'Ersteller',      type:'text', get:i => i.createdBy?.user?.displayName || i.createdBy?.user?.email || '' },
  { key:'created',          label:'Erstellt am',    type:'date', get:i => i.createdDateTime || '' },
];

function _uniqueReportVals(colKey) {
  const col = REPORT_COLS.find(c => c.key === colKey);
  if (!col) return [];
  const set = new Set();
  for (const i of allItems) { const v = col.get(i); if (v != null && String(v).trim() !== '') set.add(String(v)); }
  return [...set].sort((a, b) => a.localeCompare(b, 'de'));
}

function initReports() {
  reportFilters = {};
  renderReports();
}

function renderReports() {
  const statusOpts = _uniqueReportVals('Status');
  const wgOpts     = _uniqueReportVals('Warengruppe');
  const maOpts     = _uniqueReportVals('creator');
  const prioOpts   = _uniqueReportVals('Prioritaet');
  const baOpts     = _uniqueReportVals('Beschaffungslogik');

  // Mehrfachauswahl-Dropdown (Excel-Stil mit Checkboxen)
  const sel = (id, label, opts) => `
    <div class="rep-filter">
      <label>${label}</label>
      <div class="rep-ms" data-field="${id}">
        <button type="button" class="rep-ms-btn" onclick="toggleRepMs(this)">
          <span class="rep-ms-label">Alle</span>
          <span class="rep-ms-caret">▾</span>
        </button>
        <div class="rep-ms-panel" style="display:none">
          ${opts.length ? opts.map(o =>
            `<label class="rep-ms-opt"><input type="checkbox" value="${esc(o)}" onchange="onRepMsChange(this)"/> <span>${esc(o)}</span></label>`
          ).join('') : '<div class="rep-ms-empty">– keine Werte –</div>'}
        </div>
      </div>
    </div>`;
  const txt = (id, label, ph) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="text" class="rep-input" id="${id}" placeholder="${ph||''}" oninput="applyReportFiltersDebounced()"/>
    </div>`;
  const num = (id, label, ph) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="number" class="rep-input" id="${id}" placeholder="${ph||''}" min="0" step="0.01" oninput="applyReportFiltersDebounced()"/>
    </div>`;
  const dat = (id, label) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="date" class="rep-input" id="${id}" onchange="applyReportFilters()"/>
    </div>`;

  $id('view-reports').innerHTML = `
    <div class="rep-wrap">
      <div class="rep-toolbar">
        <div class="rep-toolbar-left">
          <span class="rep-count" id="rep-count"></span>
          <span class="rep-sum" id="rep-sum"></span>
        </div>
        <div class="rep-toolbar-right">
          <button class="btn btn-sm btn-ghost" onclick="resetReportFilters()">↺ Filter zurücksetzen</button>
          <button class="btn btn-sm btn-primary" onclick="exportReportsExcel()">⬇ Als Excel exportieren</button>
        </div>
      </div>
      <div class="rep-filters">
        ${txt('rep-search','Suche','Titel, ID, Artikel…')}
        ${sel('rep-status','Status', statusOpts)}
        ${sel('rep-wg','Warengruppe', wgOpts)}
        ${sel('rep-ma','Mitarbeiter', maOpts)}
        ${sel('rep-prio','Priorität', prioOpts)}
        ${sel('rep-ba','Beschaffungsart', baOpts)}
        ${txt('rep-lieferant','Lieferant','enthält…')}
        ${txt('rep-kst','Kostenstelle','enthält…')}
        ${txt('rep-be','Bestell-Nr.','enthält…')}
        ${num('rep-preis-min','Volumen ab €','min')}
        ${num('rep-preis-max','Volumen bis €','max')}
        ${dat('rep-created-from','Erstellt von')}
        ${dat('rep-created-to','Erstellt bis')}
        ${dat('rep-termin-from','Benötigt von')}
        ${dat('rep-termin-to','Benötigt bis')}
      </div>
      <div class="rep-active" id="rep-active"></div>
      <div class="rep-table-wrap">
        <table class="rep-table" id="rep-table"></table>
      </div>
    </div>`;
  _restoreReportFilters();   // zuletzt genutzte Filter wiederherstellen
  applyReportFilters();      // liest Werte + rendert Tabelle
}

// Liest die angehakten Werte eines Mehrfachauswahl-Dropdowns
function repMsValues(field) {
  const wrap = document.querySelector(`.rep-ms[data-field="${field}"]`);
  if (!wrap) return [];
  return [...wrap.querySelectorAll('input[type=checkbox]:checked')].map(c => c.value);
}

// Öffnet/schließt ein Dropdown; schließt andere offene
function toggleRepMs(btn) {
  const panel = btn.parentElement.querySelector('.rep-ms-panel');
  const isOpen = panel.style.display !== 'none';
  document.querySelectorAll('.rep-ms-panel').forEach(p => p.style.display = 'none');
  document.querySelectorAll('.rep-ms-btn').forEach(b => b.classList.remove('open'));
  if (!isOpen) { panel.style.display = 'block'; btn.classList.add('open'); }
}

// Aktualisiert Button-Label + filtert
function onRepMsChange(cb) {
  const wrap = cb.closest('.rep-ms');
  const btn  = wrap.querySelector('.rep-ms-label');
  const vals = [...wrap.querySelectorAll('input[type=checkbox]:checked')].map(c => c.value);
  if (btn) {
    btn.textContent = vals.length === 0 ? 'Alle'
      : vals.length === 1 ? vals[0]
      : `${vals.length} ausgewählt`;
    btn.classList.toggle('active', vals.length > 0);
  }
  applyReportFilters();
}

// Schließt offene Dropdowns bei Klick außerhalb
document.addEventListener('click', e => {
  if (!e.target.closest('.rep-ms')) {
    document.querySelectorAll('.rep-ms-panel').forEach(p => p.style.display = 'none');
    document.querySelectorAll('.rep-ms-btn').forEach(b => b.classList.remove('open'));
  }
});

function applyReportFilters() {
  reportFilters = {
    search:      ($id('rep-search')?.value || '').toLowerCase().trim(),
    status:      repMsValues('rep-status'),
    wg:          repMsValues('rep-wg'),
    ma:          repMsValues('rep-ma'),
    prio:        repMsValues('rep-prio'),
    ba:          repMsValues('rep-ba'),
    lieferant:   ($id('rep-lieferant')?.value || '').toLowerCase().trim(),
    kst:         ($id('rep-kst')?.value || '').toLowerCase().trim(),
    be:          ($id('rep-be')?.value || '').toLowerCase().trim(),
    preisMin:    parseFloat($id('rep-preis-min')?.value) || null,
    preisMax:    parseFloat($id('rep-preis-max')?.value) || null,
    createdFrom: $id('rep-created-from')?.value || '',
    createdTo:   $id('rep-created-to')?.value || '',
    terminFrom:  $id('rep-termin-from')?.value || '',
    terminTo:    $id('rep-termin-to')?.value || '',
  };
  try { localStorage.setItem(LS_REPORT_FILTERS, JSON.stringify(reportFilters)); } catch {}
  renderReportsTable();
}

// Entprellte Variante für Texteingaben (verhindert Neurendern bei jedem Tastendruck)
let _repFilterTimer = null;
function applyReportFiltersDebounced() {
  clearTimeout(_repFilterTimer);
  _repFilterTimer = setTimeout(applyReportFilters, 200);
}

// Filter-Persistenz + Wiederherstellung
const LS_REPORT_FILTERS = 'DIHAG_REPORT_FILTERS';
const REP_TEXT_INPUTS = {
  search:'rep-search', lieferant:'rep-lieferant', kst:'rep-kst', be:'rep-be',
  preisMin:'rep-preis-min', preisMax:'rep-preis-max',
  createdFrom:'rep-created-from', createdTo:'rep-created-to',
  terminFrom:'rep-termin-from', terminTo:'rep-termin-to',
};
const REP_MS_FIELDS = { status:'rep-status', wg:'rep-wg', ma:'rep-ma', prio:'rep-prio', ba:'rep-ba' };
const REP_LABELS = {
  search:'Suche', status:'Status', wg:'Warengruppe', ma:'Mitarbeiter', prio:'Priorität',
  ba:'Beschaffungsart', lieferant:'Lieferant', kst:'Kostenstelle', be:'Bestell-Nr.',
  preisMin:'Volumen ab', preisMax:'Volumen bis', createdFrom:'Erstellt von', createdTo:'Erstellt bis',
  terminFrom:'Benötigt von', terminTo:'Benötigt bis',
};

function _restoreReportFilters() {
  let saved;
  try { saved = JSON.parse(localStorage.getItem(LS_REPORT_FILTERS) || '{}'); } catch { return; }
  if (!saved || typeof saved !== 'object') return;
  // Textfelder/Zahlen/Daten
  for (const [key, id] of Object.entries(REP_TEXT_INPUTS)) {
    const v = saved[key];
    const el = $id(id);
    if (el && v != null && v !== '') el.value = v;
  }
  // Mehrfachauswahl
  for (const [key, field] of Object.entries(REP_MS_FIELDS)) {
    const vals = saved[key];
    if (!Array.isArray(vals) || !vals.length) continue;
    const wrap = document.querySelector(`.rep-ms[data-field="${field}"]`);
    if (!wrap) continue;
    wrap.querySelectorAll('input[type=checkbox]').forEach(c => { if (vals.includes(c.value)) c.checked = true; });
    const lbl = wrap.querySelector('.rep-ms-label');
    if (lbl) {
      lbl.textContent = vals.length === 1 ? vals[0] : `${vals.length} ausgewählt`;
      lbl.classList.add('active');
    }
  }
}

function clearReportFilter(key) {
  if (REP_TEXT_INPUTS[key]) {
    const el = $id(REP_TEXT_INPUTS[key]); if (el) el.value = '';
  } else if (REP_MS_FIELDS[key]) {
    const wrap = document.querySelector(`.rep-ms[data-field="${REP_MS_FIELDS[key]}"]`);
    if (wrap) {
      wrap.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = false);
      const lbl = wrap.querySelector('.rep-ms-label');
      if (lbl) { lbl.textContent = 'Alle'; lbl.classList.remove('active'); }
    }
  }
  applyReportFilters();
}

function _renderActiveFilters() {
  const el = $id('rep-active');
  if (!el) return;
  const f = reportFilters || {};
  const chips = [];
  const chip = (key, val) => chips.push(
    `<span class="rep-chip">${esc(REP_LABELS[key] || key)}: <b>${esc(val)}</b>` +
    `<button title="Filter entfernen" onclick="clearReportFilter('${key}')">✕</button></span>`);
  ['status','wg','ma','prio','ba'].forEach(k => { if (Array.isArray(f[k]) && f[k].length) chip(k, f[k].join(', ')); });
  ['search','lieferant','kst','be','createdFrom','createdTo','terminFrom','terminTo'].forEach(k => { if (f[k]) chip(k, f[k]); });
  if (f.preisMin != null) chip('preisMin', fmtEuro(f.preisMin));
  if (f.preisMax != null) chip('preisMax', fmtEuro(f.preisMax));
  el.innerHTML = chips.length
    ? chips.join('') + `<button class="rep-chip-clear" onclick="resetReportFilters()">Alle löschen</button>`
    : '';
  el.style.display = chips.length ? 'flex' : 'none';
}

function resetReportFilters() {
  reportFilters = {};
  try { localStorage.removeItem(LS_REPORT_FILTERS); } catch {}
  ['rep-search','rep-lieferant','rep-kst','rep-be','rep-preis-min','rep-preis-max',
   'rep-created-from','rep-created-to','rep-termin-from','rep-termin-to']
    .forEach(id => { const el = $id(id); if (el) el.value = ''; });
  // Mehrfachauswahl zurücksetzen
  document.querySelectorAll('.rep-ms').forEach(wrap => {
    wrap.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = false);
    const lbl = wrap.querySelector('.rep-ms-label');
    if (lbl) { lbl.textContent = 'Alle'; lbl.classList.remove('active'); }
  });
  renderReportsTable();
}

function getReportRows() {
  const f = reportFilters;
  const get = (i, key) => { const c = REPORT_COLS.find(x => x.key === key); return c ? c.get(i) : ''; };
  const has = arr => Array.isArray(arr) && arr.length > 0;
  let rows = allItems.filter(i => {
    if (has(f.status) && !f.status.includes(get(i,'Status') || '')) return false;
    if (has(f.wg)     && !f.wg.includes(get(i,'Warengruppe') || '')) return false;
    if (has(f.ma)     && !f.ma.includes(get(i,'creator') || '')) return false;
    if (has(f.prio)   && !f.prio.includes(get(i,'Prioritaet') || '')) return false;
    if (has(f.ba)     && !f.ba.includes(get(i,'Beschaffungslogik') || '')) return false;
    if (f.lieferant && !String(get(i,'Lieferant') || '').toLowerCase().includes(f.lieferant)) return false;
    if (f.kst       && !String(get(i,'Kostenstelle') || '').toLowerCase().includes(f.kst)) return false;
    if (f.be        && !String(get(i,'Bestellnummer') || '').toLowerCase().includes(f.be)) return false;
    const preis = get(i,'GeschaetzterPreis');
    if (f.preisMin != null && (preis == null || preis < f.preisMin)) return false;
    if (f.preisMax != null && (preis == null || preis > f.preisMax)) return false;
    const created = (get(i,'created') || '').slice(0,10);
    if (f.createdFrom && (!created || created < f.createdFrom)) return false;
    if (f.createdTo   && (!created || created > f.createdTo)) return false;
    const termin = (get(i,'Termin') || '').slice(0,10);
    if (f.terminFrom && (!termin || termin < f.terminFrom)) return false;
    if (f.terminTo   && (!termin || termin > f.terminTo)) return false;
    if (f.search) {
      const hay = [get(i,'id'), get(i,'Title'), get(i,'Artikelnummer'), get(i,'Lieferant'),
                   get(i,'Bestellnummer'), get(i,'Kostenstelle'), get(i,'creator')]
                  .map(v => String(v ?? '').toLowerCase()).join(' ');
      if (!hay.includes(f.search)) return false;
    }
    return true;
  });
  // Sort
  const col = REPORT_COLS.find(c => c.key === reportSort.col) || REPORT_COLS[0];
  const dir = reportSort.dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let va = col.get(a), vb = col.get(b);
    if (col.type === 'num')  { va = va == null ? -Infinity : va; vb = vb == null ? -Infinity : vb; return (va - vb) * dir; }
    if (col.type === 'date') { return (String(va).localeCompare(String(vb))) * dir; }
    return String(va).localeCompare(String(vb), 'de') * dir;
  });
  return rows;
}

function reportSortBy(colKey) {
  if (reportSort.col === colKey) reportSort.dir = reportSort.dir === 'asc' ? 'desc' : 'asc';
  else { reportSort.col = colKey; reportSort.dir = 'asc'; }
  renderReportsTable();
}

function renderReportsTable() {
  const table = $id('rep-table');
  if (!table) return;
  _renderActiveFilters();
  const rows = getReportRows();

  // Count + sum
  const cnt = $id('rep-count');
  if (cnt) cnt.textContent = `${rows.length} ${rows.length === 1 ? 'Eintrag' : 'Einträge'}`;
  const sumVal = rows.reduce((s, i) => {
    const c = REPORT_COLS.find(x => x.key === 'GeschaetzterPreis');
    return s + (c.get(i) || 0);
  }, 0);
  const sumEl = $id('rep-sum');
  if (sumEl) sumEl.textContent = sumVal > 0 ? `· Gesamtvolumen ${fmtEuro(sumVal)}` : '';

  const arrow = key => reportSort.col === key ? (reportSort.dir === 'asc' ? ' ▲' : ' ▼') : '';
  const head = `<thead><tr>${REPORT_COLS.map(c =>
    `<th class="rep-th${c.type==='num'?' rep-right':''}" onclick="reportSortBy('${c.key}')">${esc(c.label)}${arrow(c.key)}</th>`
  ).join('')}</tr></thead>`;

  // Render-Performance: nur die ersten REP_ROW_CAP Zeilen darstellen (Export bleibt vollständig)
  const REP_ROW_CAP = 300;
  const shown = rows.slice(0, REP_ROW_CAP);
  const rowHtml = i => {
    const id = i.id;
    return `<tr class="rep-row" onclick="navigate('detail','${id}')">${REPORT_COLS.map(c => {
      let v = c.get(i);
      if (c.key === 'Status') return `<td class="rep-td">${statusBadge(v)}</td>`;
      if (c.type === 'date')  v = v ? fmtDate(v) : '–';
      else if (c.key === 'GeschaetzterPreis') v = v != null ? fmtEuro(v) : '–';
      else if (c.type === 'num') v = v != null ? v.toLocaleString('de-DE') : '–';
      else v = (v == null || v === '') ? '–' : v;
      return `<td class="rep-td${c.type==='num'?' rep-right':''}">${esc(String(v))}</td>`;
    }).join('')}</tr>`;
  };
  const capNote = rows.length > REP_ROW_CAP
    ? `<tr><td class="rep-cap-note" colspan="${REPORT_COLS.length}">Erste ${REP_ROW_CAP} von ${rows.length} angezeigt – Filter eingrenzen oder „Als Excel exportieren" für die vollständige Liste.</td></tr>`
    : '';
  const body = rows.length === 0
    ? `<tbody><tr><td class="rep-empty" colspan="${REPORT_COLS.length}">Keine Einträge für die gewählten Filter.</td></tr></tbody>`
    : `<tbody>${shown.map(rowHtml).join('')}${capNote}</tbody>`;

  table.innerHTML = head + body;
}

async function exportReportsExcel() {
  const rows = getReportRows();
  if (!rows.length) { toast('Keine Daten zum Exportieren.', 'error'); return; }
  await ensureXLSX().catch(() => {}); // bei Fehlschlag greift der CSV-Fallback unten
  const header = REPORT_COLS.map(c => c.label);
  const aoa = [header];
  for (const it of rows) {
    aoa.push(REPORT_COLS.map(c => {
      let v = c.get(it);
      if (c.type === 'date' && v) v = fmtDate(v);
      return v == null ? '' : v;
    }));
  }
  const stamp = new Date().toISOString().slice(0,10);
  if (typeof XLSX !== 'undefined') {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = REPORT_COLS.map(c => ({ wch: Math.max(10, c.label.length + 2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Bedarfsanfragen');
    XLSX.writeFile(wb, `Bedarfsanfragen_Report_${stamp}.xlsx`);
    toast(`${rows.length} Einträge als Excel exportiert.`, 'success');
  } else {
    // Fallback: CSV
    const csv = aoa.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(';')).join('\r\n');
    downloadCSV(csv, `Bedarfsanfragen_Report_${stamp}.csv`);
    toast(`${rows.length} Einträge als CSV exportiert (Excel-Bibliothek nicht geladen).`, 'success');
  }
}

// ── SCHULUNG (BETA) ──────────────────────────────────────────────────────────
function initSchulung() {
  $id('view-schulung').innerHTML = `
    <div class="schulung-wrap">
      <div class="schulung-hero">
        <div class="schulung-badge">Beta</div>
        <h1>Willkommen im DIHAG Bedarfsanfragen-Tool</h1>
        <p>Diese Schulung erklärt Schritt für Schritt, wie Sie Bedarfsanfragen erstellen,
           verfolgen und genehmigen. Die Anwendung ist mit SharePoint verbunden – alle
           Eingaben werden zentral gespeichert und über Power Automate weitergeleitet.</p>
      </div>

      <div class="schulung-toc">
        <a href="#sch-1">1 · Überblick</a>
        <a href="#sch-2">2 · Neue Anfrage</a>
        <a href="#sch-3">3 · Sammelanfrage</a>
        <a href="#sch-4">4 · Status & Genehmigung</a>
        <a href="#sch-5">5 · Einkauf-Daten</a>
        <a href="#sch-6">6 · Reports</a>
        <a href="#sch-7">7 · Wiedervorlage/Favoriten</a>
        <a href="#sch-8">8 · Tabellen-Importer</a>
        <a href="#sch-9">9 · FAQ</a>
      </div>

      <section class="schulung-sec" id="sch-1">
        <h2>1 · Überblick</h2>
        <p>Die Navigation links führt zu den Hauptbereichen:</p>
        <ul>
          <li><b>Dashboard</b> – Liste aller Anfragen mit Status-Chips, Warengruppen- und Mitarbeiterfilter.</li>
          <li><b>Neue Anfrage</b> – Assistent für eine einzelne Position.</li>
          <li><b>Sammelanfrage</b> – Assistent für mehrere Positionen in einer Anfrage.</li>
          <li><b>Meine Anfragen</b> – nur Ihre eigenen Anfragen.</li>
          <li><b>Reports</b> – Auswertungen mit vielen Filtern und Excel-Export.</li>
          <li><b>Tabellen-Importer</b> – Artikelnummern- und Kostenstellentabellen aktualisieren.</li>
        </ul>
        <div class="schulung-tip">💡 Die Liste aktualisiert sich automatisch alle 20 Sekunden.
          Während Sie eine Anfrage ausfüllen, pausiert die Aktualisierung, damit nichts verloren geht.</div>
      </section>

      <section class="schulung-sec" id="sch-2">
        <h2>2 · Neue Anfrage erstellen</h2>
        <ol>
          <li><b>Bedarf:</b> Bezeichnung und Warengruppe (Pflicht). Geben Sie eine Artikelnummer
              ein – Bezeichnung und Warengruppe werden automatisch vorgeschlagen.</li>
          <li><b>Menge:</b> Menge, Mengeneinheit und „Benötigt bis"-Termin.</li>
          <li><b>Beschaffung:</b> Beschaffungsart, Lieferant(en), Kostenstelle und
              <b>Bestellvolumen in € (Pflicht)</b>. Anhand des Volumens zeigt das Tool den
              erforderlichen Genehmigungsweg an (z. B. Anzahl Angebote).</li>
          <li><b>Prüfen & Einreichen:</b> Zusammenfassung kontrollieren und absenden.</li>
        </ol>
        <div class="schulung-tip">💡 Über das Feld „Externe Artikelnummer" erfassen Sie bei
          Katalogware die Lieferanten-Nummer.</div>
      </section>

      <section class="schulung-sec" id="sch-3">
        <h2>3 · Sammelanfrage (mehrere Positionen)</h2>
        <p>Nutzen Sie die Sammelanfrage, wenn mehrere unterschiedliche Artikel zusammen
           angefragt werden.</p>
        <ol>
          <li><b>Positionen erfassen:</b> Pro Zeile Artikelnummer, Bezeichnung, Menge, ME,
              <b>Preis</b>, <b>Benötigt bis</b> und <b>Kostenstelle</b>. Mit „+ Position
              hinzufügen" beliebig erweitern.</li>
          <li>Das <b>Bestellvolumen</b> wird automatisch aus den Einzelpreisen summiert.</li>
          <li>Warengruppe, Priorität und Beschaffungsart gelten für die gesamte Anfrage.</li>
        </ol>
      </section>

      <section class="schulung-sec" id="sch-4">
        <h2>4 · Status & Genehmigung</h2>
        <p>Jede Anfrage durchläuft den folgenden Freigabeprozess. Die <b>Genehmigung beginnt in
           der App</b> (Schritte 1–3, Prüfung durch den Einkauf); ab „Genehmigt (Einkauf)"
           übernimmt der <b>Power-Automate-Workflow</b> die weiteren Stufen. Je nach Beschaffungsart
           und Bestellvolumen werden einzelne Stufen automatisch übersprungen – im „Verlauf" einer
           Anfrage erscheinen <b>nur die Stufen, die für genau diese Anfrage zutreffen</b>.</p>

        <div class="flow-steps">
          <div class="flow-step">
            <div class="flow-num">1</div>
            <div class="flow-body"><div class="flow-name">Eingereicht</div>
              <div class="flow-sub">Anfrage wurde erstellt.</div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step">
            <div class="flow-num">2</div>
            <div class="flow-body"><div class="flow-name">In Prüfung (Einkauf)</div>
              <div class="flow-sub">Der Einkauf prüft und entscheidet (Genehmigen / Ablehnen) – direkt in der App.</div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step">
            <div class="flow-num">3</div>
            <div class="flow-body"><div class="flow-name">Genehmigt (Einkauf)</div>
              <div class="flow-sub">Startet den automatischen Genehmigungs-Workflow in Power Automate.</div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step flow-cond">
            <div class="flow-num">4</div>
            <div class="flow-body"><div class="flow-name">In Prüfung (Controlling)</div>
              <div class="flow-sub"><span class="flow-tag">nur bei Dienstleistung</span></div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step flow-cond">
            <div class="flow-num">5</div>
            <div class="flow-body"><div class="flow-name">In Prüfung (Werkleitung)</div>
              <div class="flow-sub"><span class="flow-tag">ab 1.500 € Bestellvolumen</span></div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step flow-cond">
            <div class="flow-num">6</div>
            <div class="flow-body"><div class="flow-name">In Prüfung (strategischer Einkauf)</div>
              <div class="flow-sub"><span class="flow-tag">ab 10.000 € &amp; kein Lead-Buyer-Abschluss</span></div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step flow-ok">
            <div class="flow-num">✓</div>
            <div class="flow-body"><div class="flow-name">Freigegeben</div>
              <div class="flow-sub">Alle erforderlichen Stufen haben zugestimmt.</div></div>
          </div>
          <div class="flow-arrow">↓</div>
          <div class="flow-step flow-ok">
            <div class="flow-num">📦</div>
            <div class="flow-body"><div class="flow-name">Bestellt</div>
              <div class="flow-sub">Einkauf hat die Bestellnummer hinterlegt.</div></div>
          </div>
        </div>

        <p style="margin-top:14px"><b>Ablehnung:</b> Lehnt eine Stufe ab, wechselt die Anfrage
           sofort auf <b>Abgelehnt</b> und der Antragsteller wird per E-Mail informiert –
           die nachfolgenden Stufen entfallen.</p>
        <div class="schulung-tip">💡 Im „Verlauf" einer Anfrage sehen Sie nur die Stufen, die für
          diese Anfrage tatsächlich durchlaufen wurden – inklusive Genehmiger und Kommentar.</div>
        <div class="schulung-tip">📨 Sobald Ihre Genehmigung erforderlich ist, werden Sie
          <b>automatisch über Power Automate per E-Mail und Microsoft Teams</b> benachrichtigt.
          Sie müssen die App also nicht ständig offen halten – die Aufforderung kommt zu Ihnen.</div>
        <div class="schulung-tip">💡 In „Meine Anfragen" sehen Sie nur den Fortschritt – die
          Genehmigungsschaltflächen erscheinen dort nicht.</div>
      </section>

      <section class="schulung-sec" id="sch-5">
        <h2>5 · Einkauf-Daten eintragen</h2>
        <p>Bei Status <b>Freigegeben</b> öffnen Sie über „📦 Einkauf-Daten" das Formular für
           Bestellnummer, Lieferdatum, tatsächlichen Preis und Angebots-PDFs.</p>
        <div class="schulung-tip">⚠️ Ist eine <b>Bestellnummer (BE)</b> bereits bei einer anderen
          Anfrage hinterlegt, warnt das Tool sofort – so vermeiden Sie Doppelbestellungen.</div>
      </section>

      <section class="schulung-sec" id="sch-6">
        <h2>6 · Reports & Auswertungen</h2>
        <p>Im Reiter <b>Reports</b> filtern Sie alle Anfragen nach Status, Warengruppe,
           Mitarbeiter, Priorität, Beschaffungsart, Lieferant, Kostenstelle, Bestell-Nr.,
           Volumen und Zeiträumen – ähnlich wie in Excel.</p>
        <ul>
          <li>Spaltenüberschrift anklicken = sortieren.</li>
          <li>Zeile anklicken = Anfrage öffnen.</li>
          <li><b>„Als Excel exportieren"</b> lädt die aktuell gefilterte Liste als .xlsx herunter.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-7">
        <h2>7 · Wiedervorlage / Favoriten</h2>
        <p>Wiederkehrende Anfragen müssen Sie nicht neu tippen:</p>
        <ul>
          <li>In einer geöffneten Anfrage auf <b>„⭐ Als Wiedervorlage"</b> klicken.</li>
          <li>Die Anfrage wird als Favorit gespeichert (lokal in Ihrem Browser).</li>
          <li>Oben in <b>Neue Anfrage</b> bzw. <b>Sammelanfrage</b> erscheinen Ihre Favoriten –
              ein Klick füllt den Assistenten automatisch aus. Sammelanfragen landen dabei
              automatisch im richtigen Assistenten.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-8">
        <h2>8 · Tabellen-Importer</h2>
        <p>Aktualisieren Sie die Hinterlegungen für die Autovervollständigung:</p>
        <ul>
          <li><b>Artikelnummern</b> und <b>Kostenstellen</b> per CSV oder XLSX hochladen.</li>
          <li>Mit <b>„⬇ Export"</b> laden Sie die aktuelle Tabelle herunter, mit
              <b>„📄 Vorlage"</b> eine leere Beispieldatei.</li>
          <li>„↺ Zurücksetzen" stellt die Standardtabelle wieder her.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-9">
        <h2>9 · Häufige Fragen</h2>
        <p><b>Warum sehe ich keine Genehmigungs-Buttons?</b><br>
           Die Genehmigungs-Aufforderung wird Ihnen <b>per Power Automate automatisch per
           E-Mail und Microsoft Teams</b> zugestellt – Sie müssen also nicht selbst danach suchen.
           In der App erscheinen die Schaltflächen „Genehmigen / Ablehnen" zusätzlich im
           <b>Dashboard / „Alle Anfragen"</b>, jedoch nicht in „Meine Anfragen".</p>
        <p><b>Meine Änderung ist weg?</b><br>
           Prüfen Sie, ob die Anfrage noch im Status „In Prüfung (Einkauf)" ist – nur dann ist
           sie bearbeitbar.</p>
        <p><b>Wo werden Favoriten gespeichert?</b><br>
           Lokal in Ihrem Browser. Bei Browserwechsel oder Cache-Löschung sind sie nicht mehr da.</p>
        <div class="schulung-tip">📨 Problem oder Frage an die IT? Schreiben Sie einfach eine
          E-Mail an <a href="mailto:ticket@dihag.com"><b>ticket@dihag.com</b></a> – es wird
          <b>automatisch ein Ticket erstellt</b> und dem passenden Kollegen zugeordnet.</div>
      </section>
    </div>`;
}

// ── SAMMELANFRAGE (MULTI-POSITION WIZARD) ────────────────────────────────────

let multiPositions = [];   // [{artNr, extArtNr, bezeichnung, menge, me, termin, kostenstelle, preis}, ...]
let multiWizardData = {};  // step2, step3

function initMultiWizard() {
  multiPositions = [{ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' }];
  multiWizardData = {};
  showMultiStep(1);
  // Reset Allgemein fields
  const wg = $id('mf-Warengruppe'); if (wg) wg.value = '';
  const prioEl = $id('mf-Prioritaet');
  if (prioEl) {
    const prioOpts = [...prioEl.options];
    const normalOpt = prioOpts.find(o => /^(normal|standard)$/i.test(o.value.trim()));
    prioEl.value = normalOpt ? normalOpt.value : (prioOpts.find(o => o.value !== '')?.value || '');
  }
  const beschEl = $id('mf-Beschreibung'); if (beschEl) beschEl.value = '';
  const terminEl = $id('mf-Termin'); if (terminEl) terminEl.value = '';
  // Reset Beschaffung fields
  ['mf-Lieferant','mf-Lieferant2','mf-Lieferant3','mf-Lieferant4','mf-GeschaetzterPreis']
    .forEach(k => { const el = $id(k); if (el) el.value = ''; });
  const firstRadio = document.querySelector('input[name=mBeschaffungslogik]');
  if (firstRadio) firstRadio.checked = true;
  document.querySelectorAll('#m-beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  [2,3,4].forEach(n => {
    const grp = $id('m-lieferant-extra-' + n);
    if (grp) grp.style.display = 'none';
  });
  // Kostenstelle autocomplete
  // Reset submit button
  const sb = $id('btn-multi-submit');
  if (sb) { sb.disabled = false; sb.textContent = '✓ Sammelanfrage einreichen'; }

  // Ausstehende Wiedervorlage anwenden (vor renderMultiPositions, damit Positionen gesetzt sind)
  if (pendingPrefill && pendingPrefill.type === 'multi') {
    applyMultiPrefill(pendingPrefill);
    toast('Wiedervorlage übernommen – bitte prüfen und absenden.', 'success');
  }
  pendingPrefill = null;

  renderMultiPositions();
  renderFavBar('multi');
}

// Füllt den Sammel-Wizard mit Favoritendaten
function applyMultiPrefill(fav) {
  const data = fav.data || {};
  const setVal = (id, v) => { const el = $id(id); if (el && v != null && v !== '') el.value = v; };
  _setSelectFlexible('mf-Warengruppe', data.Warengruppe);
  _setSelectFlexible('mf-Prioritaet', data.Prioritaet);
  setVal('mf-Beschreibung', data.Beschreibung);
  setVal('mf-Lieferant', data.Lieferant);
  setVal('mf-Lieferant2', data.Lieferant2);
  setVal('mf-Lieferant3', data.Lieferant3);
  setVal('mf-Lieferant4', data.Lieferant4);
  setVal('mf-GeschaetzterPreis', data.GeschaetzterPreis);
  // Beschaffungslogik: Radio + Zusatz-Karten
  _applyBeschaffungPrefill(data.Beschaffungslogik, 'mBeschaffungslogik', 'm-beschaffungslogik-extra-cards');
  [2,3,4].forEach(n => {
    if (data['Lieferant' + n]) { const grp = $id('m-lieferant-extra-' + n); if (grp) grp.style.display = ''; }
  });
  // Positionen übernehmen
  if (Array.isArray(fav.positions) && fav.positions.length) {
    multiPositions = fav.positions.map(p => ({
      artNr:        p.Artikelnummer || '',
      extArtNr:     p.ExterneArtikelnummer || '',
      bezeichnung:  p.Bezeichnung || '',
      menge:        p.Menge != null ? String(p.Menge) : '',
      me:           p.ME || p.me || '',
      preis:        p.Preis != null ? String(p.Preis) : '',
      termin:       p.Termin ? String(p.Termin).slice(0,10) : '',
      kostenstelle: p.Kostenstelle || '',
    }));
  }
}

function showMultiStep(n) {
  [1,2,3,4].forEach(i => {
    const body = $id('mwstep-' + i);
    if (body) body.classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-mstep="${i}"]`);
    if (!s) return;
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function renderMultiPositions() {
  const container = $id('multi-positions-table');
  if (!container) return;
  if (multiPositions.length === 0) multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  const meOptions = ['','Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  const rows = multiPositions.map((pos, i) => `
    <div class="multi-pos-row" data-idx="${i}">
      <div class="multi-pos-num">${i + 1}</div>
      <div class="multi-pos-fields">
        <div class="multi-pos-field">
          <label>Artikelnummer</label>
          <div class="tid-ac-wrap">
            <input type="text" class="mpos-artnr" data-idx="${i}"
              value="${esc(pos.artNr)}"
              placeholder="z. B. 4001-00010"
              oninput="multiPosChange(${i},'artNr',this.value)"
              autocomplete="off"/>
          </div>
        </div>
        <div class="multi-pos-field">
          <label>Externe ArtNr. <span class="field-sub">Katalog</span></label>
          <input type="text" class="mpos-extartnr" data-idx="${i}"
            value="${esc(pos.extArtNr || '')}"
            placeholder="Lieferanten-Nr."
            oninput="multiPosChange(${i},'extArtNr',this.value)"
            autocomplete="off"/>
        </div>
        <div class="multi-pos-field" style="flex:2">
          <label>Bezeichnung <span class="req">*</span></label>
          <input type="text" class="mpos-bez" data-idx="${i}"
            value="${esc(pos.bezeichnung)}"
            placeholder="Artikelbezeichnung"
            oninput="multiPosChange(${i},'bezeichnung',this.value)"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 90px">
          <label>Menge <span class="req">*</span></label>
          <input type="number" class="mpos-menge" data-idx="${i}"
            value="${esc(pos.menge)}"
            placeholder="1" min="0.001" step="any"
            oninput="multiPosChange(${i},'menge',this.value); updateMultiTotal()"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 130px">
          <label>ME <span class="req">*</span></label>
          <select class="mpos-me" data-idx="${i}" onchange="multiPosChange(${i},'me',this.value)">
            ${meOptions.map(o => `<option value="${o}"${pos.me === o ? ' selected' : ''}>${o || '– wählen –'}</option>`).join('')}
          </select>
        </div>
        <div class="multi-pos-field" style="flex:0 0 110px">
          <label>Einzelpreis (€)</label>
          <input type="number" class="mpos-preis" data-idx="${i}"
            value="${esc(String(pos.preis || ''))}"
            placeholder="0.00" min="0" step="0.01"
            oninput="multiPosChange(${i},'preis',this.value); updateMultiTotal()"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 140px">
          <label>Benötigt bis</label>
          <input type="date" class="mpos-termin" data-idx="${i}"
            value="${esc(String(pos.termin || ''))}"
            oninput="multiPosChange(${i},'termin',this.value)"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 160px">
          <label>Kostenstelle</label>
          <div class="tid-ac-wrap">
            <input type="text" class="mpos-kostenstelle" data-idx="${i}"
              value="${esc(String(pos.kostenstelle || ''))}"
              placeholder="Nr. oder Bezeichnung…"
              oninput="multiPosChange(${i},'kostenstelle',this.value)"
              autocomplete="off"/>
          </div>
        </div>
      </div>
      ${multiPositions.length > 1
        ? `<button type="button" class="multi-pos-del" title="Position entfernen" onclick="removeMultiPosition(${i})">✕</button>`
        : '<div class="multi-pos-del-placeholder"></div>'}
    </div>`).join('');
  container.innerHTML = rows;
  // Init TID autocomplete for each artnr input
  container.querySelectorAll('.mpos-artnr').forEach(inp => {
    const idx = parseInt(inp.dataset.idx, 10);
    initTidAutocomplete(inp, key => {
      if (key === 'Artikelnummer') return inp;
      if (key === 'Title') {
        // Proxy: sets this row's Bezeichnung
        return {
          get value() { return multiPositions[idx]?.bezeichnung || ''; },
          set value(v) {
            if (multiPositions[idx]) {
              multiPositions[idx].bezeichnung = v;
              const bezEl = document.querySelector(`.mpos-bez[data-idx="${idx}"]`);
              if (bezEl) bezEl.value = v;
            }
          }
        };
      }
      if (key === 'Warengruppe') {
        // Fill global Warengruppe in step 2
        return $id('mf-Warengruppe');
      }
      return null;
    });
  });
  // Init Kostenstelle autocomplete for each row
  container.querySelectorAll('.mpos-kostenstelle').forEach(inp => {
    const idx = parseInt(inp.dataset.idx, 10);
    initKostenstAuto(inp);
    inp.addEventListener('change', e => { multiPosChange(idx, 'kostenstelle', e.target.value); });
  });
  updateMultiTotal();
}

// Positions-Gesamtbetrag = Einzelpreis × Menge (Menge fehlt/ungültig → als 1 gewertet)
function _lineSum(preis, menge) {
  const p = parseFloat(preis) || 0;
  const m = parseFloat(menge);
  return p * (m > 0 ? m : 1);
}
function updateMultiTotal() {
  const total = multiPositions.reduce((s, p) => s + _lineSum(p.preis, p.menge), 0);
  const el = $id('mf-GeschaetzterPreis');
  if (el) el.value = total > 0 ? total.toFixed(2) : '';
  const displayEl = $id('mf-total-display');
  if (displayEl) displayEl.textContent = total > 0 ? fmtEuro(total) : '–';
}

function multiPosChange(idx, field, val) {
  if (multiPositions[idx]) multiPositions[idx][field] = val;
}

function addMultiPosition() {
  multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  renderMultiPositions();
}

function removeMultiPosition(idx) {
  multiPositions.splice(idx, 1);
  if (multiPositions.length === 0) multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  renderMultiPositions();
}

function toggleMBLExtra(el) {
  el.classList.toggle('selected');
}

function addMLieferant(n) {
  const grp = $id('m-lieferant-extra-' + n);
  if (grp) grp.style.display = '';
  const prevBtn = $id('mbtn-add-lieferant-' + (n - 1));
  if (prevBtn) prevBtn.style.display = 'none';
}

function wMultiNext(step) {
  if (step === 1) {
    // Validate all positions
    for (let i = 0; i < multiPositions.length; i++) {
      const p = multiPositions[i];
      // sync values from DOM (in case oninput didn't fire on all)
      const bezEl   = document.querySelector(`.mpos-bez[data-idx="${i}"]`);
      const mengeEl = document.querySelector(`.mpos-menge[data-idx="${i}"]`);
      const meEl    = document.querySelector(`.mpos-me[data-idx="${i}"]`);
      const artEl    = document.querySelector(`.mpos-artnr[data-idx="${i}"]`);
      const extArtEl = document.querySelector(`.mpos-extartnr[data-idx="${i}"]`);
      if (bezEl)    p.bezeichnung = bezEl.value.trim();
      if (mengeEl)  p.menge = mengeEl.value;
      if (meEl)     p.me = meEl.value;
      if (artEl)    p.artNr = artEl.value.trim();
      if (extArtEl) p.extArtNr = extArtEl.value.trim();
      const terminEl = document.querySelector(`.mpos-termin[data-idx="${i}"]`);
      const kstEl    = document.querySelector(`.mpos-kostenstelle[data-idx="${i}"]`);
      const preisEl  = document.querySelector(`.mpos-preis[data-idx="${i}"]`);
      if (terminEl) p.termin      = terminEl.value;
      if (kstEl)    p.kostenstelle = kstEl.value.trim();
      if (preisEl)  p.preis       = preisEl.value;
      if (!p.bezeichnung) { toast(`Position ${i + 1}: Bitte Bezeichnung angeben.`, 'error'); return; }
      if (!p.menge || parseFloat(p.menge) <= 0) { toast(`Position ${i + 1}: Bitte gültige Menge eingeben.`, 'error'); return; }
      if (!p.me)   { toast(`Position ${i + 1}: Bitte Mengeneinheit wählen.`, 'error'); return; }
    }
  } else if (step === 2) {
    const wg = $id('mf-Warengruppe').value;
    if (!wg) { toast('Bitte Warengruppe wählen.', 'error'); return; }
    multiWizardData.step2 = {
      Warengruppe:  wg,
      Prioritaet:   $id('mf-Prioritaet').value,
      Beschreibung: $id('mf-Beschreibung').value.trim(),
    };
  } else if (step === 3) {
    const lieferant = $id('mf-Lieferant').value.trim();
    if (!lieferant) { toast('Bitte mindestens Lieferant 1 angeben.', 'error'); return; }
    const extraSelected = document.querySelector('#m-beschaffungslogik-extra-cards .check-card.selected');
    if (!extraSelected) { toast('Bitte unter „Zusätzlich kombinierbar" eine Option auswählen.', 'error'); return; }
    multiWizardData.step3 = {
      Beschaffungslogik: [
        document.querySelector('input[name=mBeschaffungslogik]:checked')?.value || '',
        ...[...document.querySelectorAll('#m-beschaffungslogik-extra-cards .check-card.selected')].map(c => c.dataset.value)
      ].filter(Boolean).join(', '),
      Lieferant:         $id('mf-Lieferant').value.trim(),
      Lieferant2:        $id('mf-Lieferant2').value.trim(),
      Lieferant3:        $id('mf-Lieferant3').value.trim(),
      Lieferant4:        $id('mf-Lieferant4').value.trim(),
      GeschaetzterPreis: $id('mf-GeschaetzterPreis').value ? parseFloat($id('mf-GeschaetzterPreis').value) : null,
    };
    buildMultiReview();
  }
  showMultiStep(step + 1);
}

function wMultiBack(step) {
  showMultiStep(step - 1);
}

function buildMultiReview() {
  const s2 = multiWizardData.step2 || {};
  const s3 = multiWizardData.step3 || {};
  const posRows = multiPositions.map((p, i) => `
    <tr>
      <td style="padding:4px 8px;color:#6b7280">${i + 1}</td>
      <td style="padding:4px 8px">${esc(p.artNr) || '–'}</td>
      <td style="padding:4px 8px;color:#9ca3af">${esc(p.extArtNr) || '–'}</td>
      <td style="padding:4px 8px">${esc(p.bezeichnung)}</td>
      <td style="padding:4px 8px;text-align:right">${esc(p.menge)}</td>
      <td style="padding:4px 8px">${esc(p.me)}</td>
      <td style="padding:4px 8px;text-align:right">${p.preis ? fmtEuro(p.preis) : '–'}</td>
      <td style="padding:4px 8px;text-align:right;font-weight:600">${p.preis ? fmtEuro(_lineSum(p.preis, p.menge)) : '–'}</td>
      <td style="padding:4px 8px">${p.termin || '–'}</td>
      <td style="padding:4px 8px">${esc(p.kostenstelle) || '–'}</td>
    </tr>`).join('');
  const liefs = [s3.Lieferant, s3.Lieferant2, s3.Lieferant3, s3.Lieferant4].filter(Boolean).join(', ');
  const preis = s3.GeschaetzterPreis != null ? new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(s3.GeschaetzterPreis) + ' €' : '–';
  $id('multi-review-content').innerHTML = `
    <div class="review-section">
      <h3 class="review-section-title">Positionen (${multiPositions.length})</h3>
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:.875rem">
          <thead><tr style="background:#f3f4f6;font-weight:600">
            <th style="padding:6px 8px;text-align:left">#</th>
            <th style="padding:6px 8px;text-align:left">ArtNr.</th>
            <th style="padding:6px 8px;text-align:left">Ext. ArtNr.</th>
            <th style="padding:6px 8px;text-align:left">Bezeichnung</th>
            <th style="padding:6px 8px;text-align:right">Menge</th>
            <th style="padding:6px 8px;text-align:left">ME</th>
            <th style="padding:6px 8px;text-align:right">Einzelpreis (€)</th>
            <th style="padding:6px 8px;text-align:right">Gesamt (€)</th>
            <th style="padding:6px 8px;text-align:left">Termin</th>
            <th style="padding:6px 8px;text-align:left">KST</th>
          </tr></thead>
          <tbody>${posRows}</tbody>
        </table>
      </div>
    </div>
    <div class="review-section" style="margin-top:16px">
      <h3 class="review-section-title">Allgemein</h3>
      <dl class="review-dl">
        <dt>Warengruppe</dt><dd>${esc(s2.Warengruppe)}</dd>
        <dt>Priorität</dt><dd>${esc(s2.Prioritaet)}</dd>
        <dt>Beschreibung</dt><dd>${esc(s2.Beschreibung) || '–'}</dd>
      </dl>
    </div>
    <div class="review-section" style="margin-top:16px">
      <h3 class="review-section-title">Beschaffung</h3>
      <dl class="review-dl">
        <dt>Beschaffungsart</dt><dd>${esc(s3.Beschaffungslogik)}</dd>
        <dt>Lieferant(en)</dt><dd>${esc(liefs) || '–'}</dd>
        <dt>Bestellvolumen</dt><dd>${preis}</dd>
      </dl>
    </div>`;
}

async function submitMultiRequest() {
  const btn = $id('btn-multi-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';
  try {
    const s2 = multiWizardData.step2 || {};
    const s3 = multiWizardData.step3 || {};
    const first = multiPositions[0] || {};

    // Auto-generate title from first position
    const autoTitle = first.bezeichnung
      ? (multiPositions.length > 1
          ? `Sammelanfrage: ${first.bezeichnung} +${multiPositions.length - 1} weitere`
          : `Sammelanfrage: ${first.bezeichnung}`)
      : 'Sammelanfrage';

    // Serialize all positions as JSON for the Positionen SP field
    const posJson = JSON.stringify(multiPositions.map((p, i) => ({
      Nr: i + 1,
      Artikelnummer: p.artNr,
      ExterneArtikelnummer: p.extArtNr || '',
      Bezeichnung: p.bezeichnung,
      Menge: p.menge,
      ME: p.me,
      Preis: p.preis ? parseFloat(p.preis) : null,
      Termin: p.termin || '',
      Kostenstelle: p.kostenstelle || '',
    })));

    const totalPreis = multiPositions.reduce((s,p) => s + _lineSum(p.preis, p.menge), 0);

    const rawData = {
      Title:             autoTitle,
      Artikelnummer:     first.artNr || '',
      Menge:             first.menge || '',
      Mengeneinheit:     first.me || '',
      Warengruppe:       s2.Warengruppe || '',
      Prioritaet:        s2.Prioritaet || '',
      Beschreibung:      s2.Beschreibung || '',
      Termin:            '',
      Beschaffungslogik: s3.Beschaffungslogik || '',
      Lieferant:         s3.Lieferant || '',
      Lieferant2:        s3.Lieferant2 || '',
      Lieferant3:        s3.Lieferant3 || '',
      Lieferant4:        s3.Lieferant4 || '',
      GeschaetzterPreis: totalPreis > 0 ? totalPreis : (s3.GeschaetzterPreis || null),
    };

    const allFields = buildFields(rawData, FORM_FIELDS);

    if (!siteId || !listId) await discoverSP();

    const listPath = `/sites/${siteId}/lists/${listId}`;

    btn.textContent = 'Anfrage wird angelegt…';
    const newItem = await gPost(`${listPath}/items`, { fields: { Title: allFields['Title'] } });
    const itemId  = newItem.id;

    // Patch remaining fields + Positionen (JSON multi-line field)
    const patchFields = { ...allFields };
    delete patchFields['Title'];

    // Add Positionen field if SP field exists
    const posField = resolvedFields['Positionen'] || 'Positionen';
    patchFields[posField] = posJson;

    if (Object.keys(patchFields).length) {
      btn.textContent = 'Felder werden gespeichert…';
      try {
        await patchRetry(`${listPath}/items/${itemId}/fields`, patchFields);
      } catch(patchErr) {
        console.warn('[submitMultiRequest] PATCH fehlgeschlagen:', patchErr.message);
        toast('Anfrage erstellt, aber einige Felder konnten nicht gespeichert werden.', 'error');
      }
    }

    toast(`Sammelanfrage mit ${multiPositions.length} Positionen eingereicht! Power Automate startet den Genehmigungsprozess.`, 'success');
    await loadItems(false);
    navigate('mine');
  } catch(e) {
    toast('Fehler: ' + e.message, 'error');
    btn.disabled = false;
    btn.textContent = '✓ Sammelanfrage einreichen';
  }
}

// ── CARD / ROW TEMPLATES ─────────────────────────────────────────────────────
function beschlShort(v) {
  if (!v) return '';
  if (/bestand/i.test(v))  return '📦 Bestand';
  if (/neu|nicht.bestand/i.test(v)) return '🆕 Neu';
  if (/dienst/i.test(v))   return '🔧 Dienstl.';
  return v.substring(0, 18);
}

function getApprovalSummary(item) {
  // Overall item status is the primary source of truth
  const sv = (getStatusVal(item) || '').trim();
  if (/^abgelehnt$/i.test(sv)) return `<span class="ica-no">✗ Abgelehnt</span>`;
  const vals = Object.entries(colByKey)
    .filter(([k,c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
    .map(([k]) => getField(item, k)).filter(Boolean);
  if (!vals.length) return '';
  if (vals.some(v => /abgelehnt|rejected/i.test(v)))
    return `<span class="ica-no">✗ Abgelehnt</span>`;
  const ok = vals.filter(v => /freigegeben|genehmigt|approved/i.test(v));
  if (ok.length) return `<span class="ica-ok">✓ ${ok.length}× Freigabe</span>`;
  return `<span class="ica-pending">⏳ In Prüfung</span>`;
}

function itemCard(item) {
  const title   = esc(cleanTitle(getField(item,'Title')));
  const status  = getStatusVal(item);
  const wg      = getField(item, resolvedFields['Warengruppe']    || 'Warengruppe')      || '';
  const preis   = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
  const prio    = getField(item, resolvedFields['Prioritaet']     || 'Prioritaet')       || '';
  const menge   = getField(item, resolvedFields['Menge']          || 'Menge')            || '';
  const me      = getField(item, resolvedFields['Mengeneinheit']  || 'Mengeneinheit')    || '';
  const beschl  = getField(item, resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik') || '';
  const liefant = getField(item, resolvedFields['Lieferant']      || 'Lieferant')        || '';
  const termin  = getField(item, resolvedFields['Termin']         || 'Termin')           || '';
  const ks      = getField(item, resolvedFields['Kostenstelle']   || 'Kostenstelle')     || '';
  const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
  const creator = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '';
  const sel     = panelItemId === String(item.id) ? ' selected' : '';

  return `
    <div class="item-card${sel}" data-id="${item.id}" onclick="openPanel('${item.id}')">
      <div class="ic-top">
        <div class="ic-title">${prioDot(prio)}${title}</div>
        <div class="ic-topright">
          <span class="ic-price">${preis ? fmtEuro(preis) : '–'}</span>
          ${statusBadge(status)}
        </div>
      </div>
      <div class="ic-tags">
        <span class="ic-tag ic-banf">#${item.id}</span>
        ${wg      ? `<span class="ic-tag ic-wg">${esc(wg)}</span>` : ''}
        ${beschl  ? `<span class="ic-tag">${beschlShort(beschl)}</span>` : ''}
        ${menge && me ? `<span class="ic-tag">⚖ ${esc(menge)} ${esc(me)}</span>` : ''}
        ${ks      ? `<span class="ic-tag">KST ${esc(ks)}</span>` : ''}
        ${liefant ? `<span class="ic-tag">🏭 ${esc(liefant)}</span>` : ''}
        ${termin  ? `<span class="ic-tag">📅 bis ${fmtDate(termin)}</span>` : ''}
      </div>
      <div class="ic-footer">
        <span class="ic-by">${creator ? `👤 ${esc(creator)}` : ''} ${created ? `· ${created}` : ''}</span>
      </div>
    </div>`;
}

function renderApprovalHighlight(item) {
  // If the overall status is "Abgelehnt", show that regardless of field values
  const sv = (getStatusVal(item) || '').trim();
  if (/^abgelehnt$/i.test(sv)) {
    return `<div class="cr-appr cr-appr-no"><span class="cr-appr-icon">✗</span><div><strong>Abgelehnt</strong></div></div>`;
  }
  const found = Object.entries(colByKey)
    .filter(([k,c]) => APPROVAL_RE.test(c.displayName||k) && !SYSTEM_FIELDS.has(k))
    .map(([k,c]) => ({ label: c.displayName||k, val: getField(item,k) }))
    .filter(c => c.val);
  if (!found.length) return '';

  const rejected = found.find(c => /abgelehnt|rejected/i.test(c.val));
  const approved = found.filter(c => /freigegeben|genehmigt|approved/i.test(c.val));
  const pending  = found.filter(c => !/abgelehnt|rejected|freigegeben|genehmigt|approved/i.test(c.val));

  if (rejected) {
    const comment = found.find(c => /kommentar|ablehn|grund/i.test(c.label) && c.val);
    return `<div class="cr-appr cr-appr-no">
      <span class="cr-appr-icon">✗</span>
      <div><strong>Abgelehnt</strong>${comment ? ` — ${esc(String(comment.val))}` : ''}</div>
    </div>`;
  }
  if (approved.length) {
    const stages = approved.map(c => esc(c.label)).join(', ');
    return `<div class="cr-appr cr-appr-ok">
      <span class="cr-appr-icon">✓</span>
      <div><strong>${approved.length}× Freigabe</strong> — ${stages}</div>
    </div>`;
  }
  if (pending.length) {
    const stages = pending.map(c => `${esc(c.label)}: ${esc(String(c.val))}`).join(' · ');
    return `<div class="cr-appr cr-appr-pending">
      <span class="cr-appr-icon">⏳</span>
      <div>${stages}</div>
    </div>`;
  }
  return '';
}

function compactRow(item) {
  const title   = getField(item,'Title') || '–';
  const status  = getStatusVal(item);
  const prio    = getField(item, resolvedFields['Prioritaet'] || 'Prioritaet') || '';
  const wg      = getField(item, resolvedFields['Warengruppe'] || 'Warengruppe') || '';
  const preis   = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
  const beschl  = getField(item, resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik') || '';
  const menge   = getField(item, resolvedFields['Menge'] || 'Menge') || '';
  const me      = getField(item, resolvedFields['Mengeneinheit'] || 'Mengeneinheit') || '';
  const termin  = getField(item, resolvedFields['Termin'] || 'Termin') || '';
  const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
  const creator = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '';
  const apprSum = getApprovalSummary(item);

  const tags = [
    wg     ? `<span class="cr2-tag cr2-wg">${esc(wg)}</span>` : '',
    beschl ? `<span class="cr2-tag">${beschlShort(beschl)}</span>` : '',
    menge && me ? `<span class="cr2-tag">⚖ ${esc(menge)} ${esc(me)}</span>` : '',
    termin ? `<span class="cr2-tag">📅 ${fmtDate(termin)}</span>` : '',
  ].filter(Boolean).join('');

  return `
    <div class="cr2" onclick="navigate('detail','${item.id}')">
      <div class="cr2-head">
        <div class="cr2-title">${prioDot(prio)}${esc(title)}</div>
        <div class="cr2-right">
          ${preis ? `<span class="cr2-price">${fmtEuro(preis)}</span>` : ''}
          ${statusBadge(status)}
        </div>
      </div>
      ${tags ? `<div class="cr2-tags">${tags}</div>` : ''}
      <div class="cr2-foot">
        <span class="cr2-meta">${creator ? `👤 ${esc(creator)}` : ''}${created ? ` · ${created}` : ''} · #${item.id}</span>
        ${apprSum}
      </div>
    </div>`;
}

// ── SPLIT PANEL ───────────────────────────────────────────────────────────────
async function openPanel(itemId) {
  if (!['mine','all','dashboard'].includes(currentView)) { navigate('detail', itemId); return; }
  panelItemId = String(itemId);
  const item  = allItems.find(i => String(i.id) === panelItemId);
  if (!item) return;

  // Auto-advance: Statuswechsel Eingereicht → In Prüfung (Einkauf) beim Öffnen.
  // Nur Einkäufer/Admin lösen ihn aus (canApprove). Auch bei EIGENEN Anfragen erlaubt –
  // ein Einkäufer darf seine eigene Anfrage in die Prüfung nehmen. Normale Antragsteller
  // sind durch canApprove() ohnehin ausgeschlossen.
  if (canApprove()) {
    const sv        = (getStatusVal(item) || 'Eingereicht').trim(); // leeres Feld = Eingereicht
    const statusCol = resolvedFields['Status'] || 'Status';
    let   advanced  = false;

    // Eingereicht ODER „Automatisch erstellter Entwurf" → In Prüfung (Einkauf).
    // Öffnet ein Einkäufer einen Auto-Entwurf, soll er wie eine normale Anfrage in die
    // Prüfung wandern (statt im grauen Entwurfsstatus zu verbleiben).
    if (/^eingereicht$/i.test(sv) || /automatisch erstellter entwurf/i.test(sv)) {
      const target = statusChoices.find(c => /pr[üu]fung/i.test(c) && /einkauf/i.test(c) && !/strategisch/i.test(c))
                  || 'In Prüfung (Einkauf)';
      try {
        await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [statusCol]: target });
        // Optimistisch allItems sofort aktualisieren (Graph-Cache umgehen)
        const cached = allItems.find(i => String(i.id) === String(itemId));
        if (cached?.fields) cached.fields[statusCol] = target;
        advanced = true;
      } catch(e) {
        const denied = /\b403\b|accessdenied|denied/i.test(e.message);
        toast('Status-Update fehlgeschlagen: ' + e.message
          + (denied ? ' – vermutlich fehlen Schreibrechte (Mitwirken) auf der SharePoint-Liste.' : ''), 'error');
      }
    }

    // Freigegeben → In Bestellung
    if (/^freigegeben$/i.test(sv)) {
      const inBestKey = statusChoices.find(c => /^in bestellung$/i.test(c.trim()));
      if (inBestKey) {
        try {
          await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [statusCol]: inBestKey });
          const cached = allItems.find(i => String(i.id) === String(itemId));
          if (cached?.fields) cached.fields[statusCol] = inBestKey;
          advanced = true;
        } catch(e) { console.warn('[openPanel] auto-advance Freigegeben failed:', e.message); }
      }
    }

    if (advanced) await loadItems(false);
  }

  // Always use the freshest item from allItems after potential reload
  const freshItem = allItems.find(i => String(i.id) === panelItemId) || item;

  $id(`panel-${currentView}-content`).innerHTML = renderPanel(freshItem);
  $id(`panel-${currentView}`).classList.remove('hidden');
  $id(`split-${currentView}`).classList.add('has-panel');

  document.querySelectorAll('.item-card').forEach(c =>
    c.classList.toggle('selected', c.dataset.id === panelItemId));

  bindPanelEvents(itemId);
}

function closePanel() {
  ['mine','all','dashboard'].forEach(v => {
    $id('panel-' + v)?.classList.add('hidden');
    $id('split-' + v)?.classList.remove('has-panel');
  });
  panelItemId = null;
  panelEditMode = false;
  _vhLoaded = {};
  document.querySelectorAll('.item-card').forEach(c => c.classList.remove('selected'));
}

function bindPanelEvents(itemId) {
  // IMPORTANT: scope all DOM queries to the CURRENT panel container.
  // Multiple views (mine, dashboard) each render renderPanel() and produce identical
  // IDs (panel-attach-body, panel-history-body, …). document.getElementById() would
  // return the first match — usually the wrong panel. Use scoped queries instead.
  const panelRoot = $id(`panel-${currentView}-content`);
  const pq = sel => panelRoot?.querySelector(sel);

  pq('#panel-close')?.addEventListener('click', async () => {
    await releaseLock(itemId);
    closePanel();
  });
  pq('#panel-edit')?.addEventListener('click', () => startEditMode(itemId));
  pq('#panel-order')?.addEventListener('click', () => openOrderModal(itemId));
  pq('#panel-fav')?.addEventListener('click', () => saveItemAsFavorite(itemId));

  // Genehmigungsaktion-Buttons
  pq('#btn-approve')?.addEventListener('click', () => doApprove(itemId));
  pq('#btn-reject')?.addEventListener('click', () => {
    pq('#aab-reject-form').style.display = '';
    pq('#btn-approve').style.display = 'none';
    pq('#btn-reject').style.display  = 'none';
  });
  pq('#btn-reject-cancel')?.addEventListener('click', () => {
    pq('#aab-reject-form').style.display = 'none';
    pq('#btn-approve').style.display = '';
    pq('#btn-reject').style.display  = '';
  });
  pq('#btn-reject-confirm')?.addEventListener('click', () =>
    doReject(itemId, pq('#aab-reject-comment')?.value?.trim() || ''));
  pq('#panel-save')?.addEventListener('click', () => savePanelEdits(itemId, panelRoot));
  pq('#panel-cancel')?.addEventListener('click', async () => {
    panelEditMode = false;
    await releaseLock(itemId);
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
  });

  // Artikelnummer autocomplete in edit mode
  const tidInput = pq('[data-ac="tid"]');
  if (tidInput) {
    initTidAutocomplete(tidInput, key => pq(`.pf-input[data-key="${key}"]`));
  }

  // Approvers are read directly from item.fields — no async history load needed.

  // Load attachments — scoped to this panel
  const attachEl = pq('#panel-attach-body');
  if (attachEl) {
    getSpToken()
      .then(tok => fetch(
        `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles?_=${Date.now()}`,
        { headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata',
            'Cache-Control': 'no-cache', Pragma: 'no-cache' } }
      ))
      .then(r => r.ok ? r.json() : { value: [] })
      .then(data => {
        const files = data.value || [];
        attachEl.innerHTML = files.length
          ? files.map(attachmentLink).join('')
          : '<span class="no-order">Keine Anhänge.</span>';
      })
      .catch(() => { attachEl.innerHTML = '<span class="no-order">Anhänge konnten nicht geladen werden.</span>'; });
  }

  // Kommentare aus SP-Listenfeld lesen (Mehrere Textzeilen, Format: "[Datum – Autor]: Text\n---\n")
  const commentsSection = pq('#panel-comments-section');
  const commentsEl      = pq('#panel-comments-body');

  function loadComments() {
    if (!commentsEl) return;
    if (commentsSection) commentsSection.style.display = '';
    const cur  = allItems.find(i => String(i.id) === String(itemId));
    const raw  = getField(cur, getKommentarCol());
    const entries = parseKommentare(raw);
    if (!entries.length) {
      commentsEl.innerHTML = '<span class="no-order">Noch keine Kommentare.</span>';
      return;
    }
    commentsEl.innerHTML = [...entries].reverse().map(c =>
      `<div class="pf-comment-row">
        <div class="pf-comment-header"><span class="pf-comment-author">💬 ${esc(c.header || '–')}</span></div>
        <div class="pf-comment-text">${esc(c.text)}</div>
      </div>`
    ).join('');
  }

  loadComments();

  // Send button — post new comment then reload
  pq('#panel-comment-send')?.addEventListener('click', async () => {
    const inp = pq('#panel-comment-input');
    const text = inp?.value?.trim();
    if (!text) return;
    const btn = pq('#panel-comment-send');
    btn.disabled = true; btn.textContent = '…';
    const ok = await postSpComment(itemId, text);
    btn.disabled = false; btn.textContent = '💬 Senden';
    if (ok) { if (inp) inp.value = ''; await loadComments(); }
    else toast('Kommentar konnte nicht gespeichert werden.', 'error');
  });

  // Ctrl+Enter shortcut in textarea
  pq('#panel-comment-input')?.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.ctrlKey) pq('#panel-comment-send')?.click();
  });

  pq('#panel-history')?.addEventListener('click', () => {
    const sec = pq('#panel-history-section');
    if (!sec) return;
    const isOpen = sec.style.display !== 'none';
    sec.style.display = isOpen ? 'none' : '';
    pq('#panel-history').textContent = isOpen ? '📋 Verlauf' : '📋 Verlauf ▲';
    if (!isOpen) loadVersionHistory(itemId, pq);
  });
}

let _vhLoaded = {};
async function loadVersionHistory(itemId, pq) {
  // pq: scoped query function from bindPanelEvents (avoids duplicate-ID collision)
  const el = pq ? pq('#panel-history-body') : $id('panel-history-body');
  if (!el) return;
  if (_vhLoaded[itemId]) return;
  _vhLoaded[itemId] = true;

  try {
    const data = await gGet(`/sites/${siteId}/lists/${listId}/items/${itemId}/versions?$expand=fields($select=*)&$top=50`);
    const vers = (data.value || []).sort((a,b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
    if (!vers.length) { el.innerHTML = '<div class="vh-empty">Keine Versionen gefunden.</div>'; return; }

    // Helper: SP Choice → {Value:...}, Lookup → {LookupValue:...}
    const spVal = v => {
      if (v === null || v === undefined) return '';
      if (typeof v === 'object') return String(v.Value ?? v.LookupValue ?? '');
      return String(v);
    };
    const fmtVH = (k, v) => {
      const s = spVal(v);
      if (!s) return '–';
      if (/preis/i.test(k))              return fmtEuro(parseFloat(s)) || s;
      if (/datum|termin|ben.?tigt/i.test(k)) return fmtDate(s);
      return s;
    };

    // Reverse map: any known SP column name → display label
    const labelMap = { Status:'Status' };
    for (const f of [...FORM_FIELDS, ...EINKAUF_FIELDS]) {
      labelMap[f.key] = f.label;
      const col = resolvedFields[f.key];
      if (col) labelMap[col] = f.label;
    }

    // Fields to always skip
    const skipKey = k =>
      SYSTEM_FIELDS.has(k) || k.endsWith('LookupId') ||
      k.startsWith('@') || k.startsWith('_') ||
      /^(id|Edit|Attachments|ContentType|AppAuthor|AppEditor|LinkTitle|ComplianceAsset)$/i.test(k);

    const statusKey = resolvedFields['Status'] || 'Status';

    const rows = vers.map((v, idx) => {
      const prev    = vers[idx + 1];
      const status  = spVal(v.fields?.[statusKey]);
      const by      = v.lastModifiedBy?.user?.displayName || v.lastModifiedBy?.user?.email || 'System';
      const dt      = fmtDateTime(v.lastModifiedDateTime);
      const isFirst = idx === vers.length - 1;

      const changes = [];
      if (isFirst) {
        changes.push('Anfrage erstellt');
      } else {
        const allKeys = new Set([
          ...Object.keys(v.fields || {}),
          ...Object.keys(prev?.fields || {}),
        ]);
        for (const k of allKeys) {
          if (skipKey(k)) continue;
          const cur = v.fields?.[k];
          const prv = prev?.fields?.[k];
          if (spVal(cur) === spVal(prv)) continue;
          const label = labelMap[k] || k;
          changes.push(`${esc(label)}: <em>${esc(fmtVH(k, prv))}</em> → <em>${esc(fmtVH(k, cur))}</em>`);
        }
        if (!changes.length) changes.push('Systemfelder aktualisiert');
      }

      const st  = STATUS_STYLES[(status||'').toLowerCase().trim()] || { bg:'#f3f4f6', color:'#6b7280' };
      const dot = status?.toLowerCase().includes('freigegeben') ? '#15803d'
                : status?.toLowerCase().includes('abgelehnt')   ? '#b91c1c'
                : status?.toLowerCase().includes('prüfung')     ? '#b45309'
                : '#6b7280';

      return `<div class="vh-item${isFirst?' vh-first':''}">
        <div class="vh-line"></div>
        <div class="vh-dot-wrap"><div class="vh-dot" style="background:${dot}"></div></div>
        <div class="vh-content">
          <div class="vh-meta"><span class="vh-by">👤 ${esc(by)}</span><span class="vh-time">${dt}</span></div>
          <div class="vh-changes">${changes.join(' · ')}</div>
          ${status ? `<span class="status-badge" style="background:${st.bg};color:${st.color};margin-top:4px;display:inline-block">${esc(status)}</span>` : ''}
        </div>
      </div>`;
    });

    el.innerHTML = `<div class="vh-timeline">${rows.join('')}</div>`;
  } catch(e) {
    el.innerHTML = `<div class="vh-error">Verlauf konnte nicht geladen werden.</div>`;
  }
}

function bindPanelEditEvents(itemId) {
  $id('panel-close')?.addEventListener('click', closePanel);
  $id('panel-save')?.addEventListener('click', () => saveEdits(itemId));
  $id('panel-cancel')?.addEventListener('click', () => {
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(item); bindPanelEvents(itemId); }
  });
}

async function savePanelEdits(itemId, panelRoot) {
  const data = {};
  (panelRoot || document).querySelectorAll('.pf-input[data-key]').forEach(inp => {
    data[inp.dataset.key] = inp.value;
  });
  const fields = buildFields(data, FORM_FIELDS);
  if (!Object.keys(fields).length) return;

  const btn = panelRoot?.querySelector('#panel-save') || $id('panel-save');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }

  const patch = { ...fields };
  const skipped = [];
  for (let i = 0; i < 15; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      await releaseLock(itemId); // clear edit lock on successful save
      panelEditMode = false;
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Gespeichert ✓', 'success');
      await loadItems(false);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='💾 Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
}

async function saveEdits(itemId) {
  const data = {};
  document.querySelectorAll('.pf-input[data-key]').forEach(inp => {
    data[inp.dataset.key] = inp.value;
  });
  const fields = buildFields(data, FORM_FIELDS);
  if (!Object.keys(fields).length) return;

  const btn = $id('panel-save');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }

  const patch = { ...fields };
  const skipped = [];
  for (let i = 0; i < 15; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Gespeichert ✓', 'success');
      await loadItems(false);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
}

// Build the inner HTML for the "Status & Genehmigung" approval section.
// ── GENEHMIGUNG IM DASHBOARD ─────────────────────────────────────────────────

async function doApprove(itemId) {
  if (!canApprove()) { toast('Keine Berechtigung zum Genehmigen.', 'error'); return; }
  const banner = document.querySelector('#approval-action-banner');
  if (banner) banner.innerHTML = '<div class="aab-hint">⏳ Wird gespeichert…</div>';

  const statusCol   = resolvedFields['Status']               || 'Status';
  const approverCol = resolvedFields['Entscheider_x002a_in'] || 'Entscheider_x002a_in';
  const commentCol  = resolvedFields['Genehmigungskommentar']|| 'Genehmigungskommentar';

  const approverName = account?.name || account?.username || '';
  const genehmigt    = 'Genehmigt (Einkauf)';

  // Find the exact SP choice string (case-sensitive)
  const choiceMatch = statusChoices.find(c => /^genehmigt.*einkauf/i.test(c.trim())) || genehmigt;

  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
      [statusCol]:   choiceMatch,
      [approverCol]: approverName,
    });

    // Update local cache
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item?.fields) {
      item.fields[statusCol]   = choiceMatch;
      item.fields[approverCol] = approverName;
    }

    toast('✅ Genehmigung erteilt.', 'success');
    await loadItems(false);
    // Re-render panel with updated item
    const updated = allItems.find(i => String(i.id) === String(itemId));
    if (updated) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(updated); bindPanelEvents(itemId); }
    }
  } catch(e) {
    toast('Fehler beim Genehmigen: ' + e.message, 'error');
    // Restore banner
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
    }
  }
}

async function doReject(itemId, comment) {
  if (!canApprove()) { toast('Keine Berechtigung zum Ablehnen.', 'error'); return; }
  const banner = document.querySelector('#approval-action-banner');
  if (banner) banner.innerHTML = '<div class="aab-hint">⏳ Wird gespeichert…</div>';

  const statusCol   = resolvedFields['Status']               || 'Status';
  const approverCol = resolvedFields['Entscheider_x002a_in'] || 'Entscheider_x002a_in';
  const commentCol  = resolvedFields['Genehmigungskommentar']|| 'Genehmigungskommentar';

  const approverName = account?.name || account?.username || '';
  const abgelehnt    = statusChoices.find(c => /^abgelehnt$/i.test(c.trim())) || 'Abgelehnt';

  const patch = {
    [statusCol]:   abgelehnt,
    [approverCol]: approverName,
  };
  if (comment && commentCol) patch[commentCol] = comment;

  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);

    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item?.fields) {
      item.fields[statusCol]   = abgelehnt;
      item.fields[approverCol] = approverName;
      if (comment) item.fields[commentCol] = comment;
    }

    toast('❌ Anfrage abgelehnt.', 'success');
    await loadItems(false);
    const updated = allItems.find(i => String(i.id) === String(itemId));
    if (updated) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(updated); bindPanelEvents(itemId); }
    }
  } catch(e) {
    toast('Fehler beim Ablehnen: ' + e.message, 'error');
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
    }
  }
}

// Extracted so it can be called both from renderPanel and from the async
// loadApproverHistory refresh without re-rendering the whole panel.
function buildApprovalInner(item, statusVal) {
  const sv = statusVal || getStatusVal(item) || 'Eingereicht';
  // statusTimeline reads all approver/comment fields directly from item.fields —
  // it handles all stages including "In Bestellung" upcoming/past logic.
  return `<div class="approval-stages">${statusTimeline(sv, item)}</div>`;
}

// Panel-positions in-place editing
let _panelPosData = []; // working copy while editing
function _syncPanelPosFromDOM() {
  const tbody = $id('pos-edit-tbody');
  if (!tbody) return;
  tbody.querySelectorAll('input.pos-edit-input, select.pos-edit-input').forEach(el => {
    const idx = parseInt(el.dataset.idx, 10);
    const field = el.dataset.field;
    if (_panelPosData[idx] !== undefined) _panelPosData[idx][field] = el.value;
  });
}
function panelPosAdd() {
  _syncPanelPosFromDOM();
  _panelPosData.push({ Nr: _panelPosData.length + 1, Artikelnummer:'', ExterneArtikelnummer:'', Bezeichnung:'', Menge:'', ME:'', Preis:'', Termin:'', Kostenstelle:'' });
  _rebuildPosEditTable();
}
function panelPosDelete(idx) {
  _syncPanelPosFromDOM();
  _panelPosData.splice(idx, 1);
  _panelPosData.forEach((p,i) => p.Nr = i + 1);
  _rebuildPosEditTable();
}
function _rebuildPosEditTable() {
  const tbody = $id('pos-edit-tbody');
  if (!tbody) return;
  const meOpts = ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  tbody.innerHTML = _panelPosData.map((p, idx) => {
    const meSelOpts = meOpts.map(o => `<option value="${o}"${(p.ME||'')=== o?' selected':''}>${o}</option>`).join('');
    return `<tr class="pos-edit-row" data-pos-idx="${idx}">
      <td class="pos-td pos-nr">${p.Nr}</td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Artikelnummer" value="${esc(String(p.Artikelnummer||''))}" placeholder="ArtNr."/></td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="ExterneArtikelnummer" value="${esc(String(p.ExterneArtikelnummer||''))}" placeholder="Ext.ArtNr."/></td>
      <td class="pos-td pos-bez"><input class="pos-edit-input" data-idx="${idx}" data-field="Bezeichnung" value="${esc(String(p.Bezeichnung||''))}" placeholder="Bezeichnung"/></td>
      <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Menge" value="${esc(String(p.Menge||''))}" placeholder="1" min="0.001" step="any"/></td>
      <td class="pos-td"><select class="pos-edit-input pos-edit-me" data-idx="${idx}" data-field="ME"><option value="">–</option>${meSelOpts}</select></td>
      <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Preis" value="${esc(String(p.Preis||''))}" placeholder="0,00" min="0" step="0.01"/></td>
      <td class="pos-td"><input class="pos-edit-input" type="date" data-idx="${idx}" data-field="Termin" value="${esc(String(p.Termin||''))}"/></td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Kostenstelle" value="${esc(String(p.Kostenstelle||''))}" placeholder="KST"/></td>
      <td class="pos-td"><button type="button" class="pos-del-btn" onclick="panelPosDelete(${idx})" title="Entfernen">✕</button></td>
    </tr>`;
  }).join('');
  // Update count header
  const title = document.querySelector('#pf-pos-section .pf-sec-title');
  if (title) title.textContent = `Positionen (${_panelPosData.length})`;
}
async function panelPosSave(itemId) {
  _syncPanelPosFromDOM();
  const posField = resolvedFields['Positionen'] || 'Positionen';
  const priceField = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
  const totalPrice = _panelPosData.reduce((s,p) => s + _lineSum(p.Preis, p.Menge), 0);
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
      [posField]: JSON.stringify(_panelPosData),
      [priceField]: totalPrice || null,
    });
    // Update allItems optimistically
    const cached = allItems.find(i => String(i.id) === String(itemId));
    if (cached?.fields) {
      cached.fields[posField] = JSON.stringify(_panelPosData);
      if (totalPrice) cached.fields[priceField] = totalPrice;
    }
    toast('Positionen gespeichert ✓', 'success');
    await loadItems(false);
  } catch(e) {
    toast('Speichern fehlgeschlagen: ' + e.message, 'error');
  }
}

function renderPanel(item, editMode = false) {
  const isMineView = currentView === 'mine';
  const statusVal  = getStatusVal(item) || 'Eingereicht';
  const createdBy  = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '–';
  const createdAt = item.createdDateTime ? fmtDate(item.createdDateTime) : '–';

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const dv  = v   => v ? String(v).slice(0,10) : '';  // YYYY-MM-DD for date inputs

  const choices = key => {
    const sp = resolvedFields[key] || key;
    return colByKey[sp]?.choice?.choices || null;
  };

  const WG_OPTS   = choices('Warengruppe')      || ['4001 – Schrott','4002 – Legierungen','4103 – Bleche','4116 – Kaufteile','4211 – Kaufteile Formerei (Kühlm. ..)','4304 – Strahlmittel','4305 – Modellbaumaterial','4306 – Formstoffkomponenten','4307 – Eingußsysteme','4308 – exotherme Kappen (Speiser)','4309 – Gießfilter','4310 – Kerne (bez. Kerne f. Gießerei)','4311 – metallurgische Zusatzstoffe','4312 – Hilfsstoffe Gießerei','4313 – Feuerfestmaterial Ofen/Pfannen','4314 – Kohleprodukte (Grafitelektr., Aufk...)','4317 – externer Modellbau','4318 – Befestigungsteile','4319 – WENDESCHNEIDPLATTE','4320 – Ersatzteile','4321 – Schleif- und Fräsmittel','4322 – Arbeitsschutzmittel','4323 – Schweißbedarf','4324 – Gemeinkostenmaterial','4325 – Technische Gase','4326 – Treibstoffe, Öle, Fette','4328 – Werkzeug','4329 – sonst. Hilfs- u. Betriebsstoffe','4333 – Kooperation','4334 – Reparatur','4335 – Serviceleistung','4336 – Entsorgung','4337 – Büromaterial','4338 – Mieten','4340 – Investitionen','4333 – Transporte','4335 – Leiharbeiter','2000 – Frachtkosten, sonstiges'];
  const PRIO_OPTS = choices('Prioritaet')       || ['Normal','Hoch','Dringend'];
  const ME_OPTS   = choices('Mengeneinheit')    || ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  const BL_OPTS   = choices('Beschaffungslogik')|| ['Bestandsmaterial (bestandsgeführt)','Neues Material (nicht-bestandsgeführt)','Dienstleistung'];

  const fRow = (fd, type, opts) => {
    const raw = gv(fd.key);
    const lbl = `<span class="pf-label">${esc(fd.label)}</span>`;
    if (editMode) {
      let inp;
      if (opts) {
        const optHtml = (fd.required ? [] : ['']).concat(opts)
          .map(o => `<option value="${esc(o)}"${String(raw)===o?' selected':''}>${esc(o||'–')}</option>`).join('');
        inp = `<select class="pf-input" data-key="${fd.key}">${optHtml}</select>`;
      } else if (type === 'textarea') {
        inp = `<textarea class="pf-input" data-key="${fd.key}" rows="2">${esc(String(raw))}</textarea>`;
      } else if (fd.key === 'Artikelnummer') {
        const val = esc(String(raw));
        const hit = TID_MAP_ACTIVE[String(raw).trim().toUpperCase()] || null;
        const hint = hit ? `<div class="tid-ac-confirm">✓ ${esc(hit.b)} · ${esc(hit.w)}</div>` : '';
        inp = `<div class="tid-ac-wrap">
          <input type="text" class="pf-input" data-key="Artikelnummer" data-ac="tid"
            value="${val}" autocomplete="off" placeholder="Nr. oder Bezeichnung…"/>
          <div class="tid-ac-dropdown" style="display:none"></div>
          ${hint}
        </div>`;
      } else {
        const val = type === 'date' ? dv(raw) : esc(String(raw));
        const isTsd = fd.key === 'Menge' || fd.key === 'Mindestlagermenge';
        const stepAttr = type === 'number' ? (isTsd ? ' step="0.001"' : ' step="0.01"') : '';
        inp = `<input type="${type}" class="pf-input" data-key="${fd.key}" value="${val}"${stepAttr}/>`;
      }
      return `<div class="pf-row">${lbl}${inp}</div>`;
    }
    if (!raw && raw !== 0) return '';
    let display = String(raw);
    if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') display = fmtEuro(raw);
    else if (fd.key === 'Termin' || fd.key === 'Lieferdatum') display = fmtDate(raw);
    else if (fd.key === 'Menge' || fd.key === 'Mindestlagermenge') {
      const n = parseFloat(raw);
      display = isNaN(n) ? String(raw) : n.toLocaleString('de-DE', { minimumFractionDigits: 0, maximumFractionDigits: 3 });
    }
    return `<div class="pf-row">${lbl}<span class="pf-val">${esc(display)}</span></div>`;
  };

  const preis = parseFloat(gv('GeschaetzterPreis')) || 0;
  const menge = parseFloat(gv('Menge')) || 1;
  const lbRaw     = gv('LeadBuyerAbschluss');
  const leadBuyer = lbRaw === true || /^(true|ja|yes|1)$/i.test(String(lbRaw || ''));
  const gesamtHint = !editMode && preis > 0 ? genehmigungsweg(preis, leadBuyer) : '';

  const orderNr  = gv('Bestellnummer');
  const lieferd  = gv('Lieferdatum');
  const tatPreis = gv('TatsaechlicherPreis');

  const isFreigegeben = /^freigegeben$/i.test((statusVal || '').trim());
  const orderBtn = isMineView ? '' : isFreigegeben
    ? `<button class="btn btn-outline btn-sm" id="panel-order">📦 Einkauf-Daten</button>`
    : `<button class="btn btn-outline btn-sm" id="panel-order" disabled title="Nur bei Status 'Freigegeben' möglich">📦 Einkauf-Daten</button>`;

  // Lock state for the edit button
  const lockField   = getLockField();
  const lockData    = parseLock(lockField ? (getField(item, lockField) || '') : '');
  const isLocked    = lockData && lockData.expiresAt > new Date();
  const me          = account?.name || account?.username || '';
  const lockedByMe  = isLocked && lockData.by === me;
  const lockBadge   = isLocked && !lockedByMe
    ? `<span class="panel-lock-badge">🔒 ${esc(lockData.by)} bearbeitet gerade</span>` : '';

  // Einkäufer/Admin dürfen jederzeit bearbeiten (z. B. Titel/Bezeichnung ändern, auch bei
  // Auto-Entwürfen). Alle anderen nur im Status „In Prüfung (Einkauf)".
  const canEdit = canApprove() || (/pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal));
  const editBtn = isMineView ? '' : (isLocked && !lockedByMe)
    ? `<button class="btn btn-outline btn-sm" disabled title="Gesperrt von ${esc(lockData.by)}">🔒 Gesperrt</button>`
    : canEdit
      ? `<button class="btn btn-outline btn-sm" id="panel-edit">✏️ Bearbeiten</button>`
      : `<button class="btn btn-outline btn-sm" disabled title="Nur bei Status 'In Prüfung (Einkauf)' möglich">✏️ Bearbeiten</button>`;

  const favBtn = `<button class="btn btn-outline btn-sm" id="panel-fav" title="Als Wiedervorlage / Favorit speichern">⭐ Als Wiedervorlage</button>`;
  const buttons = editMode
    ? `${orderBtn}
       <button class="btn btn-primary btn-sm" id="panel-save">💾 Speichern</button>
       <button class="btn btn-ghost btn-sm" id="panel-cancel">✕ Abbrechen</button>
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`
    : `${editBtn} ${orderBtn}
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>
       ${favBtn}`;

  const approvalInner = buildApprovalInner(item, statusVal);

  // Genehmigungsaktion-Banner: sichtbar für alle im Dashboard/All-View (nicht in eigener Mine-Ansicht)
  const needsApproval = !editMode && !isMineView && canApprove()
    && /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
  const approvalActionBanner = needsApproval ? `
    <div class="approval-action-banner" id="approval-action-banner">
      <div class="aab-title">⏳ Genehmigung ausstehend – Einkauf</div>
      <div class="aab-hint">Bitte prüfen Sie die Anfrage und erteilen Sie Ihre Entscheidung:</div>
      <div class="aab-buttons">
        <button class="btn btn-success btn-sm" id="btn-approve">✅ Genehmigen</button>
        <button class="btn btn-danger  btn-sm" id="btn-reject">❌ Ablehnen</button>
      </div>
      <div class="aab-reject-form" id="aab-reject-form" style="display:none">
        <textarea id="aab-reject-comment" class="pf-comment-textarea" rows="2"
          placeholder="Ablehnungsgrund (optional)…"></textarea>
        <div class="aab-buttons" style="margin-top:6px">
          <button class="btn btn-danger btn-sm" id="btn-reject-confirm">❌ Ablehnung bestätigen</button>
          <button class="btn btn-ghost  btn-sm" id="btn-reject-cancel">Abbrechen</button>
        </div>
      </div>
    </div>` : '';

  return `
    <div class="panel-hdr">
      <div class="panel-hdr-top">
        <div class="panel-meta">
          <span class="item-id">ID ${item.id}</span>
          ${statusWithApprover(item)}
          ${prioTag(gv('Prioritaet'))}
        </div>
        <button class="panel-close" id="panel-close" title="Schließen">✕</button>
      </div>
      <div class="panel-title">${esc(gv('Title') || '–')}</div>
      <div class="panel-byline">von ${esc(createdBy)} · ${createdAt}</div>
      ${lockBadge}
      <div class="panel-actions">${buttons}</div>
    </div>

    ${approvalActionBanner}
    <div class="panel-body">
      <div class="pf-section">
        <div class="pf-sec-title">Bedarf</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Title'),       'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Beschreibung'),'textarea')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Warengruppe'), 'text', WG_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Prioritaet'),  'text', PRIO_OPTS)}
        ${(() => { const fd = FORM_FIELDS.find(f=>f.key==='AutoBedarfDatum'); return (editMode || gv('AutoBedarfDatum')) ? fRow(fd, 'date') : ''; })()}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Mengengaben</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Menge'),             'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Mengeneinheit'),     'text', ME_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Mindestlagermenge'), 'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Termin'),            'date')}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Beschaffungsdetails</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Beschaffungslogik'),'text', BL_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Artikelnummer'),         'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='ExterneArtikelnummer'), 'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Lieferant'),        'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='GeschaetzterPreis'),'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Kostenstelle'),     'text')}
        ${gesamtHint ? `<div class="info-box info" style="margin-top:8px;font-size:.78rem">${gesamtHint}</div>` : ''}
      </div>
      ${(() => {
        const posRaw = gv('Positionen');
        if (!posRaw) return '';
        let positions;
        try { positions = JSON.parse(posRaw); } catch { return ''; }
        if (!Array.isArray(positions) || positions.length === 0) return '';

        _panelPosData = positions.map(p => ({...p}));

        const canEditPos = !editMode && !isMineView && /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
        const meOpts = ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];

        const rows = positions.map((p, idx) => {
          if (canEditPos) {
            const meSelOpts = meOpts.map(o => `<option value="${o}"${(p.ME||p.me||'')=== o?' selected':''}>${o}</option>`).join('');
            return `<tr class="pos-edit-row" data-pos-idx="${idx}">
              <td class="pos-td pos-nr">${p.Nr ?? idx+1}</td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Artikelnummer" value="${esc(String(p.Artikelnummer||''))}" placeholder="ArtNr."/></td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="ExterneArtikelnummer" value="${esc(String(p.ExterneArtikelnummer||''))}" placeholder="Ext.ArtNr."/></td>
              <td class="pos-td pos-bez"><input class="pos-edit-input" data-idx="${idx}" data-field="Bezeichnung" value="${esc(String(p.Bezeichnung||''))}" placeholder="Bezeichnung"/></td>
              <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Menge" value="${esc(String(p.Menge||''))}" placeholder="1" min="0.001" step="any"/></td>
              <td class="pos-td"><select class="pos-edit-input pos-edit-me" data-idx="${idx}" data-field="ME"><option value="">–</option>${meSelOpts}</select></td>
              <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Preis" value="${esc(String(p.Preis||''))}" placeholder="0,00" min="0" step="0.01"/></td>
              <td class="pos-td"><input class="pos-edit-input" type="date" data-idx="${idx}" data-field="Termin" value="${esc(String(p.Termin||''))}"/></td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Kostenstelle" value="${esc(String(p.Kostenstelle||''))}" placeholder="KST"/></td>
              <td class="pos-td"><button type="button" class="pos-del-btn" onclick="panelPosDelete(${idx})" title="Entfernen">✕</button></td>
            </tr>`;
          } else {
            return `<tr>
              <td class="pos-td pos-nr">${p.Nr ?? ''}</td>
              <td class="pos-td">${esc(String(p.Artikelnummer||'–'))}</td>
              <td class="pos-td">${esc(String(p.ExterneArtikelnummer||'–'))}</td>
              <td class="pos-td pos-bez">${esc(String(p.Bezeichnung||'–'))}</td>
              <td class="pos-td pos-right">${esc(String(p.Menge||'–'))}</td>
              <td class="pos-td">${esc(String(p.ME||p.me||'–'))}</td>
              <td class="pos-td pos-right">${p.Preis ? fmtEuro(p.Preis) : '–'}</td>
              <td class="pos-td">${p.Termin ? fmtDate(p.Termin) : '–'}</td>
              <td class="pos-td">${esc(String(p.Kostenstelle||'–'))}</td>
            </tr>`;
          }
        }).join('');

        const addBtn = canEditPos ? `<button type="button" class="btn btn-ghost btn-sm" style="margin-top:6px" onclick="panelPosAdd()">+ Position hinzufügen</button>` : '';
        const saveBtn = canEditPos ? `<div style="margin-top:8px;display:flex;gap:8px"><button class="btn btn-primary btn-sm" onclick="panelPosSave('${item.id}')">💾 Positionen speichern</button></div>` : '';

        return `
        <div class="pf-section" id="pf-pos-section">
          <div class="pf-sec-title">Positionen (${positions.length})</div>
          <div class="pos-table-wrap">
            <table class="pos-table" id="pos-edit-table">
              <thead><tr>
                <th class="pos-th pos-nr">#</th>
                <th class="pos-th">Artikel-Nr.</th>
                <th class="pos-th">Ext. Artikel-Nr.</th>
                <th class="pos-th pos-bez">Bezeichnung</th>
                <th class="pos-th pos-right">Menge</th>
                <th class="pos-th">ME</th>
                <th class="pos-th pos-right">Einzelpreis (€)</th>
                <th class="pos-th">Benötigt bis</th>
                <th class="pos-th">Kostenstelle</th>
                ${canEditPos ? '<th class="pos-th"></th>' : ''}
              </tr></thead>
              <tbody id="pos-edit-tbody">${rows}</tbody>
            </table>
          </div>
          ${addBtn}
          ${saveBtn}
        </div>`;
      })()}
      <div class="pf-section">
        <div class="pf-sec-title">Status &amp; Genehmigung</div>
        <div id="panel-approval-body">${approvalInner}</div>
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Bestellung (Einkauf)</div>
        ${orderNr  ? `<div class="pf-row"><span class="pf-label">Bestellnummer</span><span class="pf-val">${esc(orderNr)}</span></div>` : ''}
        ${lieferd  ? `<div class="pf-row"><span class="pf-label">Lieferdatum</span><span class="pf-val">${fmtDate(lieferd)}</span></div>` : ''}
        ${tatPreis ? `<div class="pf-row"><span class="pf-label">Tatsächl. Preis</span><span class="pf-val">${fmtEuro(tatPreis)}</span></div>` : ''}
        ${!orderNr && !lieferd && !tatPreis ? '<p class="no-order">Noch keine Bestelldaten.</p>' : ''}
      </div>
      <div class="pf-section" id="panel-comments-section">
        <div class="pf-sec-title">Kommentare</div>
        <div id="panel-comments-body"><div class="vh-loading">Lädt…</div></div>
        <div class="pf-comment-compose">
          <textarea id="panel-comment-input" class="pf-comment-textarea" rows="2" placeholder="Kommentar schreiben…"></textarea>
          <button class="btn btn-sm btn-primary" id="panel-comment-send">💬 Senden</button>
        </div>
      </div>
      <div class="pf-section" id="panel-attach-section">
        <div class="pf-sec-title">Anhänge</div>
        <div id="panel-attach-body"><div class="vh-loading">Lädt…</div></div>
      </div>
      <div class="pf-section vh-section" id="panel-history-section" style="display:none">
        <div class="pf-sec-title">Versionsverlauf</div>
        <div id="panel-history-body"><div class="vh-loading">Lädt…</div></div>
      </div>
    </div>`;
}

function emptyState(msg) {
  return `<div class="empty-state">
    <svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M5 4a3 3 0 00-3 3v6a3 3 0 003 3h10a3 3 0 003-3V7a3 3 0 00-3-3H5zm-1 9v-1h5v2H5a1 1 0 01-1-1zm7 1h4a1 1 0 001-1v-1h-5v2zm0-4h5V8h-5v2zM9 8H4v2h5V8z" clip-rule="evenodd"/></svg>
    <p>${msg}</p></div>`;
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
function closeModal() { $id('modal-overlay').classList.add('hidden'); }
document.addEventListener('DOMContentLoaded', () => {
  $id('modal-overlay')?.addEventListener('click', e => {
    if (e.target === $id('modal-overlay')) closeModal();
  });
});

// ── ATTACHMENT LINKS ─────────────────────────────────────────────────────────
// Cross-origin fetch to SharePoint is blocked by CORS.
// The user already has an AAD browser session → direct navigation works fine.
function spFullUrl(relUrl) {
  // Encode special chars that break URL parsing (# → %23, ? → %3F, space → %20)
  const encoded = relUrl.replace(/%/g, '%25')   // must be first to avoid double-encoding
                        .replace(/#/g, '%23')
                        .replace(/\?/g, '%3F')
                        .replace(/ /g,  '%20');
  // But if SP already percent-encoded some chars, we'd double-encode — safer: only fix bare #
  // Use a simpler targeted approach:
  const safe = relUrl.split('/').map(seg =>
    seg.replace(/#/g, '%23').replace(/\?/g, '%3F').replace(/ /g, '%20')
  ).join('/');
  return 'https://' + SP_SITE.split(':/')[0] + safe;
}

function getKommentarCol() {
  // Try resolved name first, then fallbacks
  const col = resolvedFields[KOMMENTAR_FIELD.key];
  if (col) return col;
  for (const t of [KOMMENTAR_FIELD.key, ...KOMMENTAR_FIELD.alsoTry]) {
    if (colByKey[t]) return t;
  }
  return KOMMENTAR_FIELD.key; // fallback — will fail gracefully
}

function parseKommentare(raw) {
  if (!raw) return [];
  return String(raw).split('\n---\n').map(s => s.trim()).filter(Boolean).map(entry => {
    const m = entry.match(/^\[(.+?)\]:\s*([\s\S]*)$/);
    if (m) return { header: m[1], text: m[2].trim() };
    return { header: '', text: entry };
  });
}

async function postSpComment(itemId, text) {
  if (!text?.trim()) return false;
  const col = getKommentarCol();
  const item = allItems.find(i => String(i.id) === String(itemId));
  const existing = String(getField(item, col) || '').trim();
  const author = account?.name || account?.username || 'Unbekannt';
  const now = new Date();
  const dt = now.toLocaleString('de-DE', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });
  const newEntry = `[${dt} – ${author}]: ${text.trim()}`;
  const combined = existing ? existing + '\n---\n' + newEntry : newEntry;
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [col]: combined });
    // Update local cache so re-render is instant
    if (item?.fields) item.fields[col] = combined;
    return true;
  } catch(e) {
    console.warn('[comments] Speichern fehlgeschlagen:', e.message);
    return false;
  }
}

// ── LOCK HELPERS ──────────────────────────────────────────────────────────────
function getLockField() {
  if (_lockField !== undefined) return _lockField;
  for (const n of LOCK_FIELD_CANDIDATES) { if (colByKey[n]) { _lockField = n; return n; } }
  _lockField = ''; return '';
}
function parseLock(raw) {
  if (!raw) return null;
  const i = String(raw).indexOf('|');
  if (i < 0) return null;
  const expiresAt = new Date(String(raw).slice(i + 1));
  return isNaN(expiresAt.getTime()) ? null : { by: String(raw).slice(0, i), expiresAt };
}
async function tryAcquireLock(itemId) {
  const field = getLockField();
  const item  = allItems.find(i => String(i.id) === String(itemId));
  const me    = account?.name || account?.username || 'Unbekannt';
  if (field && item) {
    const lock = parseLock(getField(item, field));
    if (lock && lock.expiresAt > new Date() && lock.by !== me) {
      const mins = Math.ceil((lock.expiresAt - new Date()) / 60000);
      toast(`Wird von „${lock.by}" bearbeitet – noch ${mins} Min. gesperrt.`, 'error');
      return false;
    }
    const until = new Date(Date.now() + 15 * 60 * 1000).toISOString();
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [field]: `${me}|${until}` });
      if (item.fields) item.fields[field] = `${me}|${until}`;
    } catch(e) { console.warn('[lock] acquire failed (field may not exist in SP):', e.message); }
  }
  if (lockTimers[itemId]) clearTimeout(lockTimers[itemId]);
  lockTimers[String(itemId)] = setTimeout(() => autoReleaseLock(itemId), 15 * 60 * 1000);
  return true;
}
async function releaseLock(itemId) {
  const key = String(itemId);
  if (lockTimers[key]) { clearTimeout(lockTimers[key]); delete lockTimers[key]; }
  const field = getLockField();
  if (!field) return;
  const item = allItems.find(i => String(i.id) === key);
  const me   = account?.name || account?.username || 'Unbekannt';
  const lock = parseLock(getField(item, field));
  if (!lock || lock.by !== me) return; // never release someone else's lock
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${key}/fields`, { [field]: '' });
    if (item?.fields) item.fields[field] = '';
  } catch(e) { console.warn('[lock] release failed:', e.message); }
}
async function autoReleaseLock(itemId) {
  const key = String(itemId);
  delete lockTimers[key];
  await releaseLock(key);
  if (String(panelItemId) === key) {
    const item      = allItems.find(i => String(i.id) === key);
    const panelRoot = $id(`panel-${currentView}-content`);
    if (item && panelRoot) {
      toast('Bearbeitungssperre abgelaufen (15 Min.) – Bearbeitungsmodus beendet.', 'info');
      panelRoot.innerHTML = renderPanel(item);
      bindPanelEvents(key);
    }
  }
}
async function startEditMode(itemId) {
  const ok = await tryAcquireLock(itemId);
  if (!ok) return;
  const item      = allItems.find(i => String(i.id) === String(itemId));
  const panelRoot = $id(`panel-${currentView}-content`);
  if (!item || !panelRoot) return;
  panelEditMode = true; // Auto-Refresh aussetzen, solange bearbeitet wird
  panelRoot.innerHTML = renderPanel(item, true);
  bindPanelEvents(itemId);
}

function openAttachment(relUrl) {
  // Legacy: open directly in new tab
  window.open(spFullUrl(relUrl), '_blank', 'noopener');
}

async function openPdfViewer(relUrl, fileName) {
  const fullUrl = spFullUrl(relUrl);
  const overlay = $id('pdf-viewer-modal');
  const iframe  = $id('pdf-viewer-iframe');
  const titleEl = $id('pdf-viewer-title');
  const extBtn  = $id('pdf-viewer-ext');
  if (!overlay) { window.open(fullUrl, '_blank', 'noopener'); return; }

  // Show modal immediately with loading state
  if (titleEl) titleEl.textContent = (fileName || 'Dokument') + ' – wird geladen…';
  if (extBtn)  extBtn.onclick = () => window.open(fullUrl, '_blank', 'noopener');
  iframe.src = 'about:blank';
  overlay.style.display = 'flex';

  // Fetch file content via SP REST API $value endpoint.
  // Using /_api/ URL instead of direct file URL: avoids X-Frame-Options on direct SP URLs
  // and benefits from CORS headers that SP sets for /_api/ routes with OAuth tokens.
  try {
    const tok = await getSpToken();
    // Decode any %23 → # to get the raw server-relative URL
    const rawRelUrl = decodeURIComponent(relUrl);

    // GetFileByServerRelativeUrl cannot handle '#' in filenames (SP treats it as URL
    // fragment regardless of encoding). Instead use the AttachmentFiles endpoint which
    // addresses files by name via a query-string alias — '#' round-trips correctly there.
    //   /Lists/Bedarfsanfrage/Attachments/{itemId}/{filename}
    const attachMatch = rawRelUrl.match(/\/Attachments\/(\d+)\/(.+)$/);
    let apiUrl;
    if (attachMatch) {
      const itemId  = attachMatch[1];
      const rawName = attachMatch[2]; // may contain #, spaces, ' etc.
      apiUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles(@f)/$value?@f='${encodeURIComponent(rawName.replace(/'/g, "''"))}'`;
    } else {
      // Fallback for non-attachment URLs
      const spSafe = rawRelUrl.replace(/#/g, '%23');
      apiUrl = `${SP_BASE}/_api/web/GetFileByServerRelativeUrl(@u)/$value?@u='${encodeURIComponent(spSafe.replace(/'/g, "''"))}'`;
    }
    const resp   = await fetch(apiUrl, { headers: { Authorization: 'Bearer ' + tok } });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const blob    = await resp.blob();
    const blobUrl = URL.createObjectURL(blob);
    // Revoke previous blob URL if any
    if (iframe._blobUrl) URL.revokeObjectURL(iframe._blobUrl);
    iframe._blobUrl = blobUrl;
    iframe.src      = blobUrl;
    if (titleEl) titleEl.textContent = fileName || 'Dokument';
  } catch(e) {
    console.warn('[openPdfViewer]', e);
    if (titleEl) titleEl.textContent = fileName || 'Dokument';
    // Fallback: plain message with external open button
    iframe.srcdoc = `<html><body style="font-family:sans-serif;display:flex;align-items:center;
      justify-content:center;height:100%;margin:0;background:#f8fafc;color:#374151">
      <div style="text-align:center;padding:32px">
        <p style="font-size:1rem;margin-bottom:20px">PDF kann nicht direkt angezeigt werden<br>(SP blockiert Einbettung).</p>
        <button onclick="parent.document.getElementById('pdf-viewer-ext').click()"
          style="padding:10px 24px;background:#1a56db;color:#fff;border:none;border-radius:8px;
          font-size:1rem;cursor:pointer">↗ In SharePoint öffnen</button>
      </div></body></html>`;
  }
}

// ── SETTINGS MODAL ───────────────────────────────────────────────────────────
function openSettings() {
  const email    = (account?.username || '').toLowerCase();
  const s        = getSettings(email);
  const adminMode = email === ADMIN_EMAIL;

  // ── Admin: Zugriffsverwaltung (zentral in SharePoint) ─────────────────────
  const accessRow = (upn, p) => `
    <div class="access-row">
      <span class="su-email" title="${esc(upn)}">${esc(upn)}</span>
      <label class="tgl-wrap" title="Dashboard">
        <input type="checkbox" ${p.dashboard ? 'checked' : ''}
          onchange="setAccess('${escJs(upn)}','dashboard',this.checked)"><span class="tgl"></span></label>
      <label class="tgl-wrap" title="Reports">
        <input type="checkbox" ${p.reports ? 'checked' : ''}
          onchange="setAccess('${escJs(upn)}','reports',this.checked)"><span class="tgl"></span></label>
      <label class="tgl-wrap" title="Tabellen-Importer">
        <input type="checkbox" ${p.importer ? 'checked' : ''}
          onchange="setAccess('${escJs(upn)}','importer',this.checked)"><span class="tgl"></span></label>
      <label class="tgl-wrap" title="Einkauf: Status setzen + genehmigen/ablehnen">
        <input type="checkbox" ${p.einkauf ? 'checked' : ''}
          onchange="setAccess('${escJs(upn)}','einkauf',this.checked)"><span class="tgl"></span></label>
      <button class="su-del" title="UPN entfernen" onclick="removeAccessUser('${escJs(upn)}')">✕</button>
    </div>`;

  const adminSection = adminMode ? `
    <hr class="modal-hr">
    <div class="settings-section-title">🔐 Zugriffsverwaltung (UPN)</div>
    <p class="su-hint">Lege fest, wer <b>Dashboard</b>, <b>Reports</b>, <b>Tabellen-Importer</b> sehen darf und wer
       als <b>Einkäufer</b> beim Öffnen den Status auf „In Prüfung (Einkauf)" setzen sowie
       genehmigen/ablehnen darf. Als Administrator hast du immer alle Rechte. Gilt zentral für alle Geräte.</p>
    <div class="su-add">
      <input type="email" id="access-new-upn" placeholder="einkauf@dihag.com" class="su-input">
      <button class="btn btn-sm btn-primary" onclick="addAccessUser()">+ Hinzufügen</button>
    </div>
    <div class="access-row access-header">
      <span class="su-email" style="font-size:.7rem;color:#6b7280">UPN / E-Mail</span>
      <span class="su-col-lbl" title="Dashboard">📊</span>
      <span class="su-col-lbl" title="Reports">📈</span>
      <span class="su-col-lbl" title="Tabellen-Importer">🗂</span>
      <span class="su-col-lbl" title="Einkauf (genehmigen)">🛒</span>
      <span style="width:22px"></span>
    </div>
    <div id="access-list">${
      Object.entries(accessUsers).sort((a,b)=>a[0].localeCompare(b[0]))
        .map(([u, p]) => accessRow(u, p)).join('') ||
      '<p class="su-empty">Noch keine UPNs freigegeben – nur der Administrator sieht diese Bereiche.</p>'
    }</div>` : '';

  // ── "Meine Einstellungen" sections ────────────────────────────────────────
  const roVal = (on, label) => `<span class="settings-val-ro">${on ? `✅ ${label}` : '—'}</span>`;

  // Auto-Refresh: admin can toggle for self; others see read-only
  const arRow = adminMode
    ? `<div class="settings-row">
        <span class="settings-label">⏱ Auto-Aktualisierung (alle 30s)</span>
        <label class="tgl-wrap">
          <input type="checkbox" ${s.autoRefresh && s.autoRefreshGranted ? 'checked' : ''}
            onchange="saveUserSettings('${esc(email)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked});this.checked?startAutoRefresh():stopAutoRefresh()">
          <span class="tgl"></span>
        </label>
       </div>`
    : `<div class="settings-row"><span class="settings-label">⏱ Auto-Aktualisierung</span>
        ${roVal(s.autoRefresh && s.autoRefreshGranted, 'Aktiviert')}</div>`;

  // Zugriffs-Übersicht für den aktuellen Nutzer (zentral gesteuert, nur lesbar)
  const featLabel = { dashboard:'Dashboard', reports:'Reports', importer:'Tabellen-Importer' };
  const myAreas = isAdmin()
    ? 'Alle Bereiche (Administrator)'
    : ([...GATED_FEATURES.filter(f => canAccess(f)).map(f => featLabel[f]),
        ...(canApprove() ? ['Einkauf (genehmigen)'] : [])].join(', ') || '— (Standardansicht)');
  const allRow = `<div class="settings-row">
      <span class="settings-label">🔐 Meine Bereiche</span>
      <span class="settings-val-ro">${esc(myAreas)}</span>
    </div>`;

  // Token-System entfällt – Zugriff wird zentral über die Zugriffsverwaltung gesteuert
  const syncBtn = '';

  $id('settings-body').innerHTML = `
    <div class="settings-section-title">⚙️ Meine Einstellungen</div>
    ${arRow}
    ${allRow}
    ${syncBtn}
    <div class="settings-row">
      <span class="settings-label">📄 Seitengröße beim Laden <span class="field-sub">es werden stets alle geladen</span></span>
      <input type="number" value="${s.pageSize}" min="50" max="200" step="10" class="su-num"
        onchange="saveUserSettings('${esc(email)}',{pageSize:parseInt(this.value)||100})">
    </div>
    <hr class="modal-hr">
    <div class="settings-section-title">🗂 Ansicht (Meine Anfragen)</div>
    <div class="settings-row">
      <span class="settings-label">Kompaktansicht</span>
      <label class="tgl-wrap">
        <input type="checkbox" ${s.compactView ? 'checked' : ''}
          onchange="saveUserSettings('${esc(email)}',{compactView:this.checked});filterView('mine')">
        <span class="tgl"></span>
      </label>
    </div>
    <div class="settings-row">
      <span class="settings-label">Abgeschlossene ausblenden</span>
      <label class="tgl-wrap">
        <input type="checkbox" ${s.hideCompleted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(email)}',{hideCompleted:this.checked});filterView('mine')">
        <span class="tgl"></span>
      </label>
    </div>
    <div class="settings-row">
      <span class="settings-label">Standard-Sortierung</span>
      <select class="su-select"
        onchange="saveUserSettings('${esc(email)}',{defaultSort:this.value});mineSortOrder=this.value;filterView('mine')">
        <option value="date-desc"  ${s.defaultSort==='date-desc'  ? 'selected':''}>Neueste zuerst</option>
        <option value="date-asc"   ${s.defaultSort==='date-asc'   ? 'selected':''}>Älteste zuerst</option>
        <option value="status"     ${s.defaultSort==='status'     ? 'selected':''}>Nach Status</option>
        <option value="price-desc" ${s.defaultSort==='price-desc' ? 'selected':''}>Preis absteigend</option>
        <option value="price-asc"  ${s.defaultSort==='price-asc'  ? 'selected':''}>Preis aufsteigend</option>
      </select>
    </div>
    ${adminSection}`;

  $id('settings-modal').classList.remove('hidden');
}

function closeSettings() { $id('settings-modal').classList.add('hidden'); }

function closePdfViewer() {
  const overlay = $id('pdf-viewer-modal');
  const iframe  = $id('pdf-viewer-iframe');
  if (overlay) overlay.style.display = 'none';
  if (iframe) {
    if (iframe._blobUrl) { URL.revokeObjectURL(iframe._blobUrl); iframe._blobUrl = null; }
    iframe.src    = 'about:blank';
    iframe.srcdoc = '';
  }
}

function attachmentLink(f) {
  const safeUrl = escJs(f.ServerRelativeUrl.replace(/#/g,'%23').replace(/\?/g,'%3F').replace(/ /g,'%20'));
  const name    = esc(f.FileName);
  return `<div class="attach-item">
    📎 <span class="attach-name">${name}</span>
    <span class="attach-actions">
      <button class="btn-attach-dl" onclick="openAttachment('${safeUrl}')" title="In SharePoint öffnen">↗ Öffnen</button>
    </span>
  </div>`;
}

// ── TOAST ─────────────────────────────────────────────────────────────────────
function toast(msg, type='info') {
  const icons = { success:'✅', error:'❌', info:'ℹ️' };
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.innerHTML = `<span>${icons[type]||'ℹ️'}</span><span>${esc(msg)}</span>`;
  $id('toast-c').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(()=>t.remove(),260); }, 4000);
}

// ── GLOBALE FEHLER-SICHTBARKEIT ──────────────────────────────────────────────
// Macht sonst stille (unerwartete) Fehler sichtbar – gedrosselt, um Spam zu vermeiden.
let _lastErrToast = 0;
function reportError(where, err) {
  const msg = err && err.message ? err.message : String(err || 'Unbekannter Fehler');
  console.error('[Fehler]', where, err);
  const now = Date.now();
  if (now - _lastErrToast > 4000) {
    _lastErrToast = now;
    try { toast('Unerwarteter Fehler: ' + msg, 'error'); } catch {}
  }
}
window.addEventListener('error', e => { if (e?.error) reportError('JS', e.error); });
window.addEventListener('unhandledrejection', e => reportError('Promise', e?.reason));

// ── HELPERS ──────────────────────────────────────────────────────────────────
const $id = id => document.getElementById(id);
const esc = s  => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

// Für Werte, die in einem einfach gequoteten JS-String INNERHALB eines
// HTML-Attributs landen, z. B. onclick="tuWas('${escJs(v)}')".
// esc() allein reicht dort NICHT: Der HTML-Parser dekodiert Entities, BEVOR der
// JS-Code kompiliert wird – aus &#39; würde wieder ' und der String bräche auf.
// Daher erst auf JS-Ebene maskieren (Backslash zuerst!), dann fürs Attribut escapen.
const escJs = s => esc(String(s||'').replace(/\\/g,'\\\\').replace(/'/g,"\\'"));

function getField(item, key) {
  if (!key || !item?.fields) return null;
  return item.fields[key] ?? null;
}

function cleanTitle(t) {
  if (!t) return '–';
  return t
    .replace(/^Einkauf\s*[-–]\s*BANF\s*#\d+\s*[-–]\s*/i, '')
    .replace(/\s*[-–]\s*[\d.,]+\s*€\s*$/i, '')
    .trim() || t;
}

// Show status badge + current approver (if a column matching APPROVER_COL_RE exists)
function statusWithApprover(item) {
  const approver = getApproverVal(item);
  const badge    = statusBadge(getStatusVal(item));
  if (!approver) return badge;
  return `${badge}<span class="status-approver" title="Aktueller Genehmiger">👤 ${esc(approver)}</span>`;
}

function statusBadge(s) {
  const label = s || 'Eingereicht';
  const sl = label.toLowerCase().trim();
  let st = STATUS_STYLES[sl];
  if (!st) {
    for (const [key, style] of Object.entries(STATUS_STYLES)) {
      if (sl.includes(key)) { st = style; break; }
    }
  }
  st = st || { bg:'#f3f4f6', color:'#374151' };
  let icon = '';
  if      (sl.includes('abgelehnt') || sl.includes('rejected'))               icon = '✗ ';
  else if (sl.includes('freigegeben') || sl.includes('bestellt') || sl.includes('erledigt')) icon = '✓ ';
  else if (sl.includes('prüfung') || sl.includes('bearbeitung'))               icon = '⏳ ';
  else if (sl.includes('eingereicht') || sl.includes('angefragt'))             icon = '📋 ';
  return `<span class="status-badge" style="background:${st.bg};color:${st.color}">${icon}${esc(label)}</span>`;
}

function prioTag(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const color = p.toLowerCase() === 'dringend' ? '#b91c1c' : '#b45309';
  const bg    = p.toLowerCase() === 'dringend' ? '#fef2f2' : '#fffbeb';
  return `<span class="status-badge" style="background:${bg};color:${color}">${esc(p)}</span>`;
}

function prioDot(p) {
  const pl = (p || '').toLowerCase();
  let c;
  if      (pl.includes('hoch') || pl.includes('dringend')) c = '#ef4444';
  else if (pl.includes('mittel'))                          c = '#f59e0b';
  else                                                     c = '#22c55e';
  return `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${c};margin-right:5px;vertical-align:middle;flex-shrink:0"></span>`;
}

function fmtEuro(v) {
  if (!v && v !== 0) return '';
  return new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(v);
}
function fmtDate(s) {
  if (!s) return '';
  return new Date(s).toLocaleDateString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric'});
}
function toLocalInputDate(s) {
  if (!s) return '';
  const d = new Date(s);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function fmtDateTime(s) {
  if (!s) return '';
  return new Date(s).toLocaleString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
function fmtRelative(s) {
  const d = Date.now() - new Date(s).getTime();
  if (d < 60000)     return 'gerade eben';
  if (d < 3600000)   return Math.floor(d/60000) + ' Min.';
  if (d < 86400000)  return Math.floor(d/3600000) + ' Std.';
  if (d < 2592000000)return Math.floor(d/86400000) + ' Tage';
  return fmtDate(s);
}
