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
  { key:'Title',             label:'Bezeichnung',                  step:1, required:true  },
  { key:'Beschreibung',      label:'Beschreibung',                 step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung','Grund'] },
  { key:'Warengruppe',       label:'Warengruppe',                  step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',        label:'Priorität',                    step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  // Step 2: Menge
  { key:'Menge',             label:'Menge',                        step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',     label:'Mengeneinheit',                step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge', label:'Mindestlagermenge',            step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',            label:'Benötigt bis',                 step:2, required:true, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis','Ben_x00f6_tigtbis'] },
  // Step 3: Beschaffung
  { key:'Artikelnummer',     label:'Artikelnummer',                 step:1, alsoTry:['MaterialNumber','ItemNumber','Artikelnummer_x002f_Nummernangab'] },
  { key:'Beschaffungslogik', label:'Beschaffungsart',              step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Lieferant',         label:'Lieferant 1',                  step:3, alsoTry:['Vendor','Supplier'] },
  { key:'Lieferant2',        label:'Lieferant 2 (Alternative)',    step:3, alsoTry:['Vendor2','Supplier2','Lieferant_2'] },
  { key:'Lieferant3',        label:'Lieferant 3 (Alternative)',    step:3, alsoTry:['Vendor3','Supplier3','Lieferant_3'] },
  { key:'Lieferant4',        label:'Lieferant 4 (Alternative)',    step:3, alsoTry:['Vendor4','Supplier4','Lieferant_4'] },
  { key:'GeschaetzterPreis',    label:'Bestellvolumen in €',          step:3, alsoTry:['EstimatedPrice','Preis','Price','Gesch_x00e4_tzterPreisnetto_x002'] },
  { key:'Kostenstelle',         label:'Kostenstelle',                 step:3, alsoTry:['CostCenter'] },
  { key:'LeadBuyerAbschluss',   label:'Lead-Buyer-Abschluss',         step:3, alsoTry:['LeadBuyer','LeadBuyerAbschlus'] },
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


// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp, account;
let siteId = null, listId = null;
let allItems = [];
let colByKey   = {};  // internal name → column definition (from SP)
let spUserMap  = {};  // SP user id (string) → display name (for Person-column LookupId resolution)
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let currentView = 'dashboard';
let prevView    = 'dashboard';
let wizardData  = {};
let wizardFilesArr = []; // Drag-and-drop file store for wizard attachments
let panelItemId = null;
// SP column may be misspelled "Stauts" in some tenants → always try both
const getStatusVal = item => getField(item,'Status') || getField(item,'Stauts') || '';

// ── USER SETTINGS ────────────────────────────────────────────────────────────
const SETTINGS_KEY = 'bedarfsanfrage_settings_v1';
const ADMIN_EMAIL  = 'administrator@dihag.com';

function getSettings(email) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  return Object.assign({
    autoRefresh:        false,
    autoRefreshGranted: false,
    pageSize:           100,
    compactView:        false,   // dense single-line cards in list views
    hideCompleted:      false,   // hide bestellt/erledigt/abgelehnt in Meine Anfragen
    defaultSort:        'date-desc', // default sort for Meine Anfragen
    canSeeDashboard:    false,   // admin-granted: show Dashboard tab (all items visible)
  }, all[(email||'').toLowerCase()] || {});
}
function saveUserSettings(email, patch, _skipSP = false) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  const key = (email||'').toLowerCase();
  all[key] = Object.assign(getSettings(email), patch);
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(all));
  // Persist to SP asynchronously so other users/devices pick up the change.
  if (!_skipSP) persistSpSettings().catch(()=>{});
  return all[key];
}
function getAllUserSettings() {
  return JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
}

// ── SP-BASED SETTINGS STORE ───────────────────────────────────────────────────
// SharePoint list/drive creation requires Manage.All permissions the app does
// not have. Instead we use a token-based grant system:
//   Admin  → generates a compact base64 token from current user grants
//          → copies it and sends to the user (e.g. via Teams/Email)
//   User   → pastes token in Settings modal → grants applied locally
// localStorage is the runtime cache; token is the out-of-band sync channel.
const SP_REST = 'https://' + SP_SITE.replace(':/', '/') + '/_api/web';
let _cfgItemId = null;

// SP REST: common JSON headers (Bearer token makes X-RequestDigest unnecessary)
function _spHdr(tok, extra = {}) {
  return Object.assign({
    Authorization: 'Bearer ' + tok,
    Accept:        'application/json;odata=nometadata',
    'Content-Type':'application/json;odata=verbose',
  }, extra);
}

// No-op: SP sync removed (insufficient permissions to create lists or write drives).
// Grants are distributed via copy-paste tokens (see openSettings / applyGrantToken).
async function loadSpSettings()    { /* token-based, no remote fetch needed */ }
async function persistSpSettings() { /* token-based, no remote write needed */ }

// ── AUTO-REFRESH ─────────────────────────────────────────────────────────────
// autoRefreshTimer: 1-second tick; arCountdown: seconds until next refresh
let autoRefreshTimer = null;
let arCountdown = 30;
let arPaused = false;   // user can pause without admin losing the feature-enable

function startAutoRefresh() {
  stopAutoRefresh();
  if (!account) return;
  if (!getSettings(account.username).autoRefresh) { updateARBtn(); return; }
  arCountdown = 30;
  arPaused    = false;
  autoRefreshTimer = setInterval(() => {
    if (arPaused) return;
    arCountdown--;
    if (arCountdown <= 0) {
      arCountdown = 30;
      loadItems(false);
    }
    updateARBtn();
  }, 1000);
  updateARBtn();
}
function stopAutoRefresh() {
  if (autoRefreshTimer) { clearInterval(autoRefreshTimer); autoRefreshTimer = null; }
  updateARBtn();
}
function updateARBtn() {
  const btn = $id('btn-autorefresh');
  if (!btn) return;
  // Feature visibility: only show button when admin has enabled autoRefresh for this user
  const featureOn = !!(account && getSettings(account.username).autoRefresh);
  btn.style.display = featureOn ? '' : 'none';
  if (!featureOn) return;
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
  // Toggle pause/resume only — does NOT change the admin-controlled autoRefresh setting
  if (!account) return;
  if (!getSettings(account.username).autoRefresh) return;
  if (!autoRefreshTimer) { startAutoRefresh(); return; }
  arPaused = !arPaused;
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
    // Load shared settings from SP (source of truth for admin grants).
    // Must happen before migration so granted flags are already in localStorage.
    await loadSpSettings();
    await loadItems(false);
    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';
    applyNavVisibility(); // re-apply after SP settings loaded (grants may have changed)
    // Log what the current user's effective settings are — visible in browser console
    if (account) {
      const _em = (account.username || '').toLowerCase();
      const _s  = getSettings(_em);
      console.log('[bootDone] User:', _em, '| canSeeDashboard:', _s.canSeeDashboard, '| autoRefresh:', _s.autoRefresh);
    }
    // Auto-refresh is admin-controlled. Clear any self-set value from old localStorage
    // if admin hasn't explicitly granted the feature (autoRefreshGranted flag).
    if (account) {
      const em = account.username.toLowerCase();
      const s  = getSettings(em);
      if (em !== ADMIN_EMAIL && s.autoRefresh && !s.autoRefreshGranted) {
        saveUserSettings(em, { autoRefresh: false }, true); // _skipSP: already synced
      }
    }
    startAutoRefresh();
    applyDashboardVisibility();
    // Set user info in sidebar
    const name = account?.name || account?.username || '?';
    $id('hdr-av').textContent   = name.split(' ').map(p=>p[0]||'').join('').substring(0,2).toUpperCase();
    $id('hdr-name').textContent = name;
    $id('hdr-mail').textContent = account?.username || '';
    const canDash = isAdmin() || getSettings((account?.username||'').toLowerCase()).canSeeDashboard;
    navigate(canDash ? 'dashboard' : 'mine');
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

async function getToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account})).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({scopes:SCOPES});
    }
    throw e;
  }
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
  const tok = await getToken();
  const r   = await fetch(API + path, {
    headers: { Authorization: 'Bearer ' + tok, 'Cache-Control': 'no-cache', Pragma: 'no-cache' }
  });
  if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
async function gPost(path, body) {
  const tok = await getToken();
  const url = API + path;
  console.log('[gPost] POST', url, JSON.stringify(body).slice(0, 300));
  const r   = await fetch(url, {
    method:'POST', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
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
  const tok = await getToken();
  const r   = await fetch(API + path, {
    method:'PATCH', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (r.status === 204) return {};
  if (!r.ok) throw new Error(`Graph PATCH ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json().catch(()=>({}));
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
    console.log('[discoverSP] Status-Werte synchronisiert:', choices);
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
    console.log('[discoverSP] spUserMap loaded:', Object.keys(spUserMap).length, 'users');
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
    const pageSize = account ? getSettings(account.username).pageSize : 100;
    const data = await gGet(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*)&$top=${pageSize}&$orderby=createdDateTime desc`
    );
    allItems = data.value || [];
    if (showToast) toast('Daten aktualisiert', 'success');
    // Re-render current view
    if (currentView === 'dashboard') renderDashboard();
    else if (currentView === 'mine')  renderList('mine');
    else if (currentView === 'all')   renderList('all');
    // Refresh open panel
    if (panelItemId && ['mine','all','dashboard'].includes(currentView)) {
      const pi = allItems.find(i => String(i.id) === panelItemId);
      if (pi) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(pi); bindPanelEvents(panelItemId); }
    }
  } catch(e) {
    toast('Fehler beim Laden: ' + e.message, 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ── ROUTING ──────────────────────────────────────────────────────────────────
// Dashboard title differs by role: admins see "Dashboard (Alle)", others see "Meine Anfragen"
function VIEW_TITLES(view) {
  const map = { new:'Neue Bedarfsanfrage', mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details' };
  if (view === 'dashboard') return isAdmin() ? 'Dashboard (Alle Anfragen)' : 'Meine Anfragen';
  return map[view] || view;
}

// Dashboard is visible to all logged-in users.
// Non-admins see only their own items; admins see all items.
function canSeeDashboard() { return !!account; }
function isAdmin() { return account?.username?.toLowerCase() === ADMIN_EMAIL; }

// Show/hide nav items based on role and admin-granted permissions.
function applyNavVisibility() {
  const email    = (account?.username || '').toLowerCase();
  const s        = getSettings(email);
  const canDash  = isAdmin() || s.canSeeDashboard;
  const dashNav  = document.querySelector('.nav-item[data-view="dashboard"]');
  if (dashNav) dashNav.style.display = canDash ? '' : 'none';
}
// Alias kept for any remaining call-sites
const applyDashboardVisibility = applyNavVisibility;

function navigate(view, id) {
  if (!canSeeDashboard() && view === 'dashboard') view = 'mine'; // should never happen now
  // Always close panels of all split views when navigating
  ['mine','all','dashboard'].forEach(v => {
    $id('panel-' + v)?.classList.add('hidden');
    $id('split-' + v)?.classList.remove('has-panel');
  });
  panelItemId = null;

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

  if      (view === 'dashboard') renderDashboard();
  else if (view === 'mine')      renderList('mine');
  else if (view === 'all')       renderList('all');
  else if (view === 'new')       initWizard();
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

  // Boot
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

  // Admin and canSeeDashboard users see all items; others see only their own.
  const _dEmail   = (account?.username||'').toLowerCase();
  const _canDash  = isAdmin() || getSettings(_dEmail).canSeeDashboard;
  const baseItems = _canDash ? allItems : myItems();

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
      onclick="setDashFilter('${s.replace(/'/g,"\\'")}')">
      ${icon ? `<span class="sc-icon">${icon}</span>` : ''}
      <span class="sc-label">${esc(labelText)}</span>
      <span class="sc-count" style="background:${st.bg};color:${st.color}">${n}</span>
    </button>`;
  }).join('');

  const vol = `<span class="sc-volume">${fmtEuro(volume)}<small>Volumen</small></span>`;
  el.innerHTML = allChip + statusChips + vol;
}

function setDashFilter(status) {
  dashStatusFilter = (dashStatusFilter === status) ? '' : status;
  renderStatusChips();
  filterDashboard();
}

function filterDashboard() {
  const search = ($id('search-dashboard')?.value || '').toLowerCase();
  const _fEmail  = (account?.username||'').toLowerCase();
  const _fCanAll = isAdmin() || getSettings(_fEmail).canSeeDashboard;
  let items = _fCanAll ? [...allItems] : [...myItems()];
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) || String(i.id||'').includes(search)
  );
  if (dashStatusFilter !== '') items = items.filter(i => (getStatusVal(i)||'') === dashStatusFilter);

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
    ? items.map(i => itemCard(i)).join('')
    : emptyState('Keine Anfragen für diesen Status.');
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
      onclick="setMineFilter('${s.replace(/'/g,"\\'")}')">
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
      ? items.map(i => itemCard(i)).join('')
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
  if (fd.key === 'Termin' || fd.key === 'Lieferdatum') return v ? fmtDate(v) : '';
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
  { label: 'In Prüfung (Werkleitung)',
    test: v => /werkleitung/i.test(v),
    approverField: 'Genehmiger3', commentField: 'Genehmigungskommentar3' },
  { label: 'In Prüfung (Controlling)',
    test: v => /controlling/i.test(v),
    approverField: 'Genehmiger2', commentField: 'Genehmigungskommentar2' },
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
  'In Prüfung (Werkleitung)',
  'In Prüfung (Controlling)',
  'In Prüfung (strategischer Einkauf)',
  'Freigegeben',
  'In Bestellung',
  'Bestellt',
];

// Farbe für einen Status anhand von Schlüsselwörtern bestimmen
function statusColorFor(val) {
  const v = (val || '').toLowerCase();
  if (/eingereicht/.test(v))            return { bg:'#fce7f3', color:'#be185d' };
  if (/werkleitung/.test(v))            return { bg:'#ffe4e6', color:'#be123c' };
  if (/prüf|pruef|prüf/.test(v))        return { bg:'#ccfbf1', color:'#0f766e' };
  if (/freigegeben/.test(v))            return { bg:'#f3e8ff', color:'#7e22ce' };
  if (/abgelehnt/.test(v))              return { bg:'#f3f4f6', color:'#374151' };
  if (/bestellt/.test(v))               return { bg:'#dbeafe', color:'#1d4ed8' };
  if (/erledigt|abgeschlossen/.test(v)) return { bg:'#f3f4f6', color:'#374151' };
  return { bg:'#f0f9ff', color:'#0369a1' };
}

// Data-driven status timeline.
// Reads approver/comment directly from item.fields — no version history needed.
// A stage is visible if: it IS the current status, OR its approverField is filled
// (meaning Power Automate put someone there, i.e. the stage was actually reached).
// "Eingereicht" is always shown.
function statusTimeline(statusVal, item) {
  const sv     = (statusVal || '').trim();
  const fields = item?.fields || item || {};
  const isRej       = /^abgelehnt$/i.test(sv);
  const isBestelltNow = /^bestellt$/i.test(sv);
  const TERMINAL_OK = /^(freigegeben|bestellt)$/i;
  const IN_BESTELLG = /^in bestellung$/i;

  const rows = TIMELINE_STAGES.map(d => {
    const isCurrent = d.test(sv);
    const approver  = resolvePersonField(fields, d.approverField);
    const comment   = d.commentField ? String(fields[d.commentField] || '').trim() : '';

    // ── Visibility ────────────────────────────────────────────────────────
    const isFreigegebenNow     = /^freigegeben$/i.test(sv);
    const isInBestellungFuture = d.label === 'In Bestellung' && isFreigegebenNow;
    const isInBestellungPast   = d.label === 'In Bestellung' && isBestelltNow;
    const visible = d.label === 'Eingereicht'  // always
      || isCurrent                             // active step
      || approver != null                      // field filled → stage was reached
      || isInBestellungPast                    // passed when now "Bestellt"
      || isInBestellungFuture;                 // upcoming when now "Freigegeben"
    if (!visible) return null;

    // ── Dot / class ───────────────────────────────────────────────────────
    let dot, cls;
    if (isCurrent) {
      if      (TERMINAL_OK.test(d.label)) { dot = '✓'; cls = 'ap-ok';      }
      else if (IN_BESTELLG.test(d.label)) { dot = '○'; cls = 'ap-circle';  }
      else if (isRej)                     { dot = '✗'; cls = 'ap-no';      }
      else                                { dot = '●'; cls = 'ap-pending'; }
    } else if (isInBestellungFuture) {
      dot = '○'; cls = 'ap-circle-future';    // upcoming, not yet reached
    } else {
      dot = '✓'; cls = 'ap-ok';               // past stage
    }

    // ── Approver + comment ────────────────────────────────────────────────
    let approverHtml = '';
    if (approver) {
      const commentHtml = comment
        ? `<div class="ap-inline-comment">💬 ${esc(comment)}</div>` : '';
      const pastCls = isCurrent ? '' : ' ap-approver-past';
      approverHtml = `<div class="ap-approver${pastCls}">👤 ${esc(approver)}${commentHtml}</div>`;
    }

    const isActive = isCurrent && !TERMINAL_OK.test(d.label) && !IN_BESTELLG.test(d.label) && !isRej;
    const bold = isActive ? ' style="font-weight:600"' : '';
    return `<div class="approval-stage"><div class="ap-dot ${cls}">${dot}</div>`
         + `<div class="ap-body"><div class="ap-stage-label"${bold}>${esc(d.label)}</div>${approverHtml}</div></div>`;
  }).filter(Boolean);

  return rows.join('');
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
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. BE252093"/>
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
        const fname = encodeURIComponent(file.name);
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
// Searches TID_MAP by number prefix OR name substring.
const TID_ENTRIES = Object.entries(TID_MAP); // pre-built for fast filtering

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
    dropdown.style.left  = r.left + 'px';
    dropdown.style.top   = (r.bottom + 2) + 'px';
    dropdown.style.width = r.width + 'px';
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
    const hit = TID_MAP[tid];
    if (!hit) return;
    inputEl.value = tid;
    const titleEl = getRelatedInput?.('Title');
    const wgEl    = getRelatedInput?.('Warengruppe');
    if (titleEl && !titleEl.dataset.manual) titleEl.value = hit.b;
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

  // Cleanup when input is removed from DOM (panel close/re-render)
  new MutationObserver(() => {
    if (!document.contains(inputEl)) { hide(); if (dropdown) dropdown.remove(); dropdown = null; }
  }).observe(document.body, { childList: true, subtree: true });
}

// ── WIZARD ────────────────────────────────────────────────────────────────────
function lookupTID(val) {
  const tid = (val || '').trim().toUpperCase();
  const matchEl = $id('tid-match');
  // Normalize: user might type "4001-00010" or "400100010" — try exact then padded
  const hit = TID_MAP[tid] || TID_MAP[tid.replace(/^(\d{4})(\d{5})$/, '$1-$2')] || null;
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

function initWizard() {
  wizardData = {};
  showStep(1);
  // Artikelnummer autocomplete in wizard
  const wArtNr = $id('f-Artikelnummer');
  if (wArtNr) initTidAutocomplete(wArtNr, key => $id('f-' + key));
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
   'Mindestlagermenge','Termin','Artikelnummer','Lieferant','Lieferant2','Lieferant3','Lieferant4',
   'GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  const prioEl = $id('f-Prioritaet'); if (prioEl) prioEl.value = 'Normal'; // Standardwert
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
}

function showStep(n) {
  [1,2,3,4].forEach(i => {
    $id('wstep-'+i).classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-step="${i}"]`);
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function wNext(step) {
  if (step === 1) {
    const title = $id('f-Title').value.trim();
    const wg    = $id('f-Warengruppe').value;
    if (!title) { toast('Bitte Bezeichnung angeben.', 'error'); return; }
    if (!wg)    { toast('Bitte Warengruppe wählen.', 'error'); return; }
    wizardData.step1 = {
      Title:        title,
      Beschreibung: $id('f-Beschreibung').value.trim(),
      Warengruppe:  wg,
      Prioritaet:   $id('f-Prioritaet').value,
      Artikelnummer: $id('f-Artikelnummer').value.trim(),
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
    wizardData.step3 = {
      Beschaffungslogik: [
        document.querySelector('input[name=Beschaffungslogik]:checked')?.value || '',
        ...[...document.querySelectorAll('#beschaffungslogik-extra-cards .check-card.selected')].map(c => c.dataset.value)
      ].filter(Boolean).join(', '),
      Lieferant:         $id('f-Lieferant').value.trim(),
      Lieferant2:        $id('f-Lieferant2').value.trim(),
      Lieferant3:        $id('f-Lieferant3').value.trim(),
      Lieferant4:        $id('f-Lieferant4').value.trim(),
      GeschaetzterPreis:  $id('f-GeschaetzterPreis').value ? parseFloat($id('f-GeschaetzterPreis').value) : null,
      Kostenstelle:       $id('f-Kostenstelle').value.trim(),
      LeadBuyerAbschluss: $id('f-LeadBuyerAbschluss')?.checked ?? false,
    };
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

function updatePreisHint() {
  const preis = parseFloat($id('f-GeschaetzterPreis').value) || 0;
  const lba   = $id('f-LeadBuyerAbschluss')?.checked ?? false;
  const hint  = $id('preis-route-hint');
  if (preis > 0) {
    let text = genehmigungsweg(preis);
    if (lba) text = text.replace(/ \(entfällt bei Lead-Buyer-Abschluss\)/g, ' ✓ entfällt');
    hint.textContent = text;
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

function genehmigungsweg(gesamt) {
  let stufe, freigabe, angebote;
  if (gesamt <= 500) {
    stufe    = 1;
    freigabe = 'Anforderer Fachabt. + Einkäufer Werk';
    angebote = 'kein Angebot nötig';
  } else if (gesamt <= 2500) {
    stufe    = 2;
    freigabe = 'Einkäufer Werk + Werkleiter';
    angebote = 'mind. 2 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 5000) {
    stufe    = 2;
    freigabe = 'Einkäufer Werk + Werkleiter';
    angebote = 'mind. 2–3 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 10000) {
    stufe    = 3;
    freigabe = 'Werkleiter + Einkaufsleitung Holding';
    angebote = 'mind. 3 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 50000) {
    stufe    = 3;
    freigabe = 'Werkleiter + Einkaufsleitung Holding';
    angebote = 'mind. 3–4 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else {
    stufe    = 4;
    freigabe = 'Einkaufsleitung Holding + GF kfm. Leitung Holding';
    angebote = 'Europ. Ausschreibung (mind. 5 Angebote, entfällt bei Lead-Buyer-Abschluss)';
  }
  return `Stufe ${stufe} · ${fmtEuro(gesamt)} · ${angebote} · Freigabe: ${freigabe}`;
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
        ['Bestellvolumen in €', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
        ['Kostenstelle', d.Kostenstelle],
      ])}
    </div>
    ${d.GeschaetzterPreis ? `<div class="info-box info" style="margin-top:12px">${genehmigungsweg(d.GeschaetzterPreis)}</div>` : ''}
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
const DATE_FIELDS   = new Set(['Termin','Lieferdatum']);
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
    console.log('[submitRequest] item created, id=', itemId);

    // ── Schritt 2: Restliche Felder per PATCH nachpflegen ──
    // SP Online has a propagation delay: the new item may not be addressable via PATCH
    // immediately after creation. Retry with generous delay (6 × 2 s = up to 12 s total).
    const patchFields = { ...allFields };
    delete patchFields['Title'];
    if (Object.keys(patchFields).length) {
      btn.textContent = 'Felder werden gespeichert…';
      console.log('[submitRequest] PATCH fields:', JSON.stringify(patchFields));
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

// ── CARD / ROW TEMPLATES ─────────────────────────────────────────────────────
function beschlShort(v) {
  if (!v) return '';
  if (/bestand/i.test(v))  return '📦 Bestand';
  if (/neu|nicht.bestand/i.test(v)) return '🆕 Neu';
  if (/dienst/i.test(v))   return '🔧 Dienstl.';
  return v.substring(0, 18);
}

function getApprovalSummary(item) {
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
  const title   = esc(getField(item,'Title') || '–');
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
        <div class="ic-title">${prioDot(prio)}${cleanTitle(getField(item,'Title'))}</div>
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
function openPanel(itemId) {
  if (!['mine','all','dashboard'].includes(currentView)) { navigate('detail', itemId); return; }
  panelItemId = String(itemId);
  const item  = allItems.find(i => String(i.id) === panelItemId);
  if (!item) return;

  $id(`panel-${currentView}-content`).innerHTML = renderPanel(item);
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

  pq('#panel-close')?.addEventListener('click', closePanel);
  pq('#panel-order')?.addEventListener('click', () => openOrderModal(itemId));
  pq('#panel-save')?.addEventListener('click', () => savePanelEdits(itemId, panelRoot));
  pq('#panel-cancel')?.addEventListener('click', () => {
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
// Extracted so it can be called both from renderPanel and from the async
// loadApproverHistory refresh without re-rendering the whole panel.
function buildApprovalInner(item, statusVal) {
  const sv = statusVal || getStatusVal(item) || 'Eingereicht';
  // statusTimeline reads all approver/comment fields directly from item.fields —
  // it handles all stages including "In Bestellung" upcoming/past logic.
  return `<div class="approval-stages">${statusTimeline(sv, item)}</div>`;
}

function renderPanel(item, editMode = false) {
  const isMineView = currentView === 'mine';
  const statusVal = getStatusVal(item) || 'Eingereicht';
  const isInPruefungEinkauf = !isMineView && /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
  if (isInPruefungEinkauf) editMode = true;
  const createdBy = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '–';
  const createdAt = item.createdDateTime ? fmtDate(item.createdDateTime) : '–';

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const dv  = v   => v ? String(v).slice(0,10) : '';  // YYYY-MM-DD for date inputs

  const choices = key => {
    const sp = resolvedFields[key] || key;
    return colByKey[sp]?.choice?.choices || null;
  };

  const WG_OPTS   = choices('Warengruppe')      || ['Schrott','Legierungen','Kaufteile','Kaufteile Formerei (Kühlm. ..)','Formstoffkomponenten','Eingußsysteme','exotherme Kappen (Speiser)','Gießfilter','Kerne (bez. Kerne f. Gießerei)','metallurgische Zusatzstoffe','Hilfsstoffe Gießerei','Feuerfestmaterial Ofen/Pfannen','Kohleprodukte (Grafitelektr., Aufk...)','Ersatzteile','Befestigungsteile','Schleif- und Fräsmittel','Strahlmittel','Werkzeug','Arbeitsschutzmittel','Schweißbedarf','Gemeinkostenmaterial','Technische Gase','Treibstoffe, Öle, Fette','sonst. Hilfs- u. Betriebsstoffe'];
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
        const hit = TID_MAP[String(raw).trim().toUpperCase()] || null;
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
  const gesamtHint = !editMode && preis > 0 ? genehmigungsweg(preis) : '';

  const orderNr  = gv('Bestellnummer');
  const lieferd  = gv('Lieferdatum');
  const tatPreis = gv('TatsaechlicherPreis');

  const isFreigegeben = /^freigegeben$/i.test((statusVal || '').trim());
  const orderBtn = isMineView ? '' : isFreigegeben
    ? `<button class="btn btn-outline btn-sm" id="panel-order">📦 Einkauf-Daten</button>`
    : `<button class="btn btn-outline btn-sm" id="panel-order" disabled title="Nur bei Status 'Freigegeben' möglich">📦 Einkauf-Daten</button>`;
  const buttons = editMode
    ? `${orderBtn}
       <button class="btn btn-primary btn-sm" id="panel-save">💾 Speichern</button>
       <button class="btn btn-ghost btn-sm" id="panel-cancel">✕ Abbrechen</button>
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`
    : `${orderBtn}
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`;

  const approvalInner = buildApprovalInner(item, statusVal);

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
      <div class="panel-actions">${buttons}</div>
    </div>

    <div class="panel-body">
      <div class="pf-section">
        <div class="pf-sec-title">Bedarf</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Title'),       'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Beschreibung'),'textarea')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Warengruppe'), 'text', WG_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Prioritaet'),  'text', PRIO_OPTS)}
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
        ${fRow(FORM_FIELDS.find(f=>f.key==='Artikelnummer'),    'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Lieferant'),        'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='GeschaetzterPreis'),'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Kostenstelle'),     'text')}
        ${gesamtHint ? `<div class="info-box info" style="margin-top:8px;font-size:.78rem">${gesamtHint}</div>` : ''}
      </div>
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
      const rawName = attachMatch[2]; // may contain #, spaces, etc.
      apiUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles(@f)/$value?@f='${encodeURIComponent(rawName)}'`;
    } else {
      // Fallback for non-attachment URLs
      const spSafe = rawRelUrl.replace(/#/g, '%23');
      apiUrl = `${SP_BASE}/_api/web/GetFileByServerRelativeUrl(@u)/$value?@u='${encodeURIComponent(spSafe)}'`;
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

  // ── Admin: per-user row ───────────────────────────────────────────────────
  const userRow = (em, us) => `
    <div class="su-row" id="su-row-${btoa(em).replace(/=/g,'')}">
      <span class="su-email" title="${esc(em)}">${esc(em)}</span>
      <label class="tgl-wrap" title="Auto-Refresh">
        <input type="checkbox" ${us.autoRefresh && us.autoRefreshGranted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(em)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked})">
        <span class="tgl"></span>
      </label>
      <label class="tgl-wrap" title="Dashboard-Zugriff (alle Anfragen)">
        <input type="checkbox" ${us.canSeeDashboard ? 'checked' : ''}
          onchange="saveUserSettings('${esc(em)}',{canSeeDashboard:this.checked})">
        <span class="tgl"></span>
      </label>
      <input type="number" value="${us.pageSize||100}" min="10" max="500" step="10"
        class="su-num" title="Max. Elemente"
        onchange="saveUserSettings('${esc(em)}',{pageSize:parseInt(this.value)||100})">
      <button class="btn btn-sm btn-outline su-token-btn" title="Grant-Token kopieren und an Benutzer senden"
        onclick="copyGrantToken('${esc(em)}')">📋 Token</button>
      <button class="su-del" title="Benutzer entfernen"
        onclick="deleteUserSetting('${esc(em)}')">✕</button>
    </div>`;

  const adminSection = adminMode ? `
    <hr class="modal-hr">
    <div class="settings-section-title">👥 Benutzer verwalten</div>
    <div class="su-add">
      <input type="email" id="su-new-email" placeholder="user@dihag.com" class="su-input">
      <button class="btn btn-sm btn-primary" onclick="addUserSetting()">+ Hinzufügen</button>
    </div>
    <div class="su-header-row">
      <span class="su-email" style="font-size:.7rem;color:#6b7280">E-Mail</span>
      <span class="su-col-lbl" title="Auto-Refresh (30s)">⏱</span>
      <span class="su-col-lbl" title="Dashboard-Zugriff">📊</span>
      <span class="su-col-lbl" title="Max. Elemente">Elem.</span>
      <span style="width:22px"></span>
    </div>
    <div id="su-list">${
      Object.entries(getAllUserSettings())
        .filter(([k]) => k !== email)
        .map(([k, v]) => userRow(k, v)).join('') ||
      '<p class="su-empty">Noch keine weiteren Benutzer konfiguriert.</p>'
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

  // canSeeDashboard: admin can toggle for self; others see read-only
  const allRow = adminMode
    ? `<div class="settings-row">
        <span class="settings-label">📊 Dashboard-Zugriff</span>
        <label class="tgl-wrap">
          <input type="checkbox" ${s.canSeeDashboard ? 'checked' : ''}
            onchange="saveUserSettings('${esc(email)}',{canSeeDashboard:this.checked});applyNavVisibility()">
          <span class="tgl"></span>
        </label>
       </div>`
    : `<div class="settings-row"><span class="settings-label">📊 Dashboard-Zugriff</span>
        ${roVal(s.canSeeDashboard, 'Freigegeben')}</div>`;

  // Token-import row for non-admin (admin generates token, sends it, user pastes here)
  const syncBtn = !adminMode ? `
    <div class="settings-row" style="margin-top:6px;gap:6px;flex-wrap:wrap;align-items:center">
      <span class="settings-label" style="font-size:.78rem;color:#6b7280">Grant-Token vom Admin:</span>
      <input id="grant-token-input" type="text" placeholder="Token hier einfügen…"
        class="su-input" style="flex:1;min-width:160px;font-size:.78rem">
      <button class="btn btn-sm btn-primary" onclick="applyGrantToken()">✅ Anwenden</button>
    </div>` : '';

  $id('settings-body').innerHTML = `
    <div class="settings-section-title">⚙️ Meine Einstellungen</div>
    ${arRow}
    ${allRow}
    ${syncBtn}
    <div class="settings-row">
      <span class="settings-label">📄 Elemente laden (max.)</span>
      <input type="number" value="${s.pageSize}" min="10" max="500" step="10" class="su-num"
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

// Admin: generate a base64 token for a specific user and copy to clipboard.
function copyGrantToken(targetEmail) {
  const settings = getAllUserSettings();
  const userCfg  = settings[(targetEmail||'').toLowerCase()];
  if (!userCfg) { toast('Benutzer nicht gefunden.', 'error'); return; }
  // Token encodes: {email, grants} — only grant fields, not personal preferences
  const grantFields = ['canSeeDashboard', 'autoRefresh', 'autoRefreshGranted'];
  const grants = {};
  grantFields.forEach(k => { if (userCfg[k] !== undefined) grants[k] = userCfg[k]; });
  const token = btoa(unescape(encodeURIComponent(JSON.stringify({ email: targetEmail.toLowerCase(), grants }))));
  navigator.clipboard.writeText(token).then(
    () => toast(`Token für ${targetEmail} kopiert. Bitte an den Benutzer senden.`, 'success'),
    () => {
      // Fallback: show in prompt
      prompt('Token kopieren:', token);
    }
  );
}

// User: paste a token from admin and apply the contained grants.
function applyGrantToken() {
  const raw = ($id('grant-token-input')?.value || '').trim();
  if (!raw) { toast('Bitte Token eingeben.', 'error'); return; }
  try {
    const decoded = JSON.parse(decodeURIComponent(escape(atob(raw))));
    if (!decoded?.email || !decoded?.grants) throw new Error('Ungültiges Format');
    const myEmail = (account?.username || '').toLowerCase();
    if (decoded.email !== myEmail) {
      toast(`Dieser Token ist für ${decoded.email}, nicht für dein Konto.`, 'error'); return;
    }
    saveUserSettings(myEmail, decoded.grants);
    applyNavVisibility();
    toast('Berechtigungen erfolgreich übernommen!', 'success');
    if (decoded.grants.canSeeDashboard) {
      setTimeout(() => location.reload(), 900);
    } else {
      openSettings();
    }
  } catch(e) {
    toast('Ungültiger Token — bitte erneut beim Admin anfragen.', 'error');
    console.warn('[applyGrantToken]', e.message);
  }
}

function addUserSetting() {
  const em = ($id('su-new-email')?.value || '').trim().toLowerCase();
  if (!em || !em.includes('@') || !em.includes('.')) {
    toast('Bitte gültige E-Mail-Adresse eingeben.', 'error'); return;
  }
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  if (all[em]) { toast(`${em} ist bereits vorhanden.`, 'info'); openSettings(); return; }
  saveUserSettings(em, { pageSize: 100 });
  toast(`${em} hinzugefügt.`, 'success');
  openSettings();
}

function deleteUserSetting(em) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  delete all[(em||'').toLowerCase()];
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(all));
  persistSpSettings().catch(() => {});
  toast(`${em} entfernt.`, 'info');
  openSettings();
}

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
  const safeUrl = esc(f.ServerRelativeUrl.replace(/#/g,'%23').replace(/\?/g,'%3F').replace(/ /g,'%20'));
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

// ── HELPERS ──────────────────────────────────────────────────────────────────
const $id = id => document.getElementById(id);
const esc = s  => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

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
