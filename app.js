'use strict';

// ── CONFIG ──────────────────────────────────────────────────────────────────
const CLIENT_ID = '75e627e8-2de0-4ec6-bec9-311757b89e08';
const TENANT_ID = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES    = ['User.Read', 'Sites.ReadWrite.All'];
const SP_SITE   = 'dihag.sharepoint.com:/sites/gruppe_shb';
const SP_LIST   = 'Bedarfsanfrage';
const API       = 'https://graph.microsoft.com/v1.0';

// ── BPMN-KONFORME FORMFELD-DEFINITION ───────────────────────────────────────
// key        = interner SP-Spaltenname (wird beim POST verwendet)
// alsoTry    = alternative interne Namen falls key nicht gefunden
// step       = Wizard-Schritt (1-3)
const FORM_FIELDS = [
  // Step 1: Bedarf
  { key:'Title',             label:'Bezeichnung',                  step:1, required:true  },
  { key:'Beschreibung',      label:'Beschreibung / Begründung',    step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung'] },
  { key:'Warengruppe',       label:'Warengruppe',                  step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',        label:'Priorität',                    step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  // Step 2: Menge
  { key:'Menge',             label:'Menge',                        step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',     label:'Mengeneinheit',                step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge', label:'Mindestlagermenge',            step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',            label:'Benötigt bis',                 step:2, required:true, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis'] },
  // Step 3: Beschaffung
  { key:'Artikelnummer',     label:'Artikelnummer / TID',           step:1, alsoTry:['MaterialNumber','ItemNumber'] },
  { key:'Beschaffungslogik', label:'Beschaffungsart',              step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Lieferant',         label:'Lieferant 1',                  step:3, alsoTry:['Vendor','Supplier'] },
  { key:'Lieferant2',        label:'Lieferant 2 (Alternative)',    step:3, alsoTry:['Vendor2','Supplier2','Lieferant_2'] },
  { key:'Lieferant3',        label:'Lieferant 3 (Alternative)',    step:3, alsoTry:['Vendor3','Supplier3','Lieferant_3'] },
  { key:'Lieferant4',        label:'Lieferant 4 (Alternative)',    step:3, alsoTry:['Vendor4','Supplier4','Lieferant_4'] },
  { key:'GeschaetzterPreis', label:'Geschätzter Preis netto (€)',  step:3, alsoTry:['EstimatedPrice','Preis','Price'] },
  { key:'Kostenstelle',      label:'Kostenstelle',                 step:3, alsoTry:['CostCenter'] },
];

// Felder die Einkauf nach der Einreichung befüllt
const EINKAUF_FIELDS = [
  { key:'Bestellnummer',  label:'Bestellnummer',        alsoTry:['OrderNumber','PO_Number'] },
  { key:'Lieferdatum',    label:'Lieferdatum',          alsoTry:['DeliveryDate'] },
  { key:'TatsaechlicherPreis', label:'Tatsächlicher Preis (€)', alsoTry:['ActualPrice','FinalPrice'] },
];

// Status-Werte → Darstellung (kommt von Power Automate)
// Reihenfolge wichtig: spezifischere zuerst (partielle Treffersuche)
const STATUS_STYLES = {
  'angefragt':      { bg:'#f0f9ff', color:'#0369a1' },
  'eingereicht':    { bg:'#eff6ff', color:'#1d4ed8' },
  'in prüfung':     { bg:'#fffbeb', color:'#b45309' },
  'freigegeben':    { bg:'#f0fdf4', color:'#15803d' },
  'abgelehnt':      { bg:'#fef2f2', color:'#b91c1c' },
  'bestellt':       { bg:'#faf5ff', color:'#7e22ce' },
  'erledigt':       { bg:'#f3f4f6', color:'#374151' },
  'in bearbeitung': { bg:'#fffbeb', color:'#b45309' },
  'offen':          { bg:'#eff6ff', color:'#1d4ed8' },
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
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let currentView = 'dashboard';
let prevView    = 'dashboard';
let wizardData  = {};
let panelItemId = null;
// SP column may be misspelled "Stauts" in some tenants → always try both
const getStatusVal = item => getField(item,'Status') || getField(item,'Stauts') || '';

// ── AUTH ────────────────────────────────────────────────────────────────────
async function initAuth() {
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth: { clientId:CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation:'localStorage', storeAuthStateInCookie:true }
  });
  await msalApp.initialize();
  await msalApp.handleRedirectPromise();
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) { account = accounts[0]; return true; }
  return false;
}

async function doLogin() {
  $id('boot-btn').style.display = 'none';
  $id('boot-sub').textContent = 'Anmeldung läuft…';
  $id('boot-spinner').style.display = 'block';
  try {
    const r = await msalApp.loginPopup({ scopes: SCOPES });
    account = r.account;
    bootDone();
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
    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';
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

async function getToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try   { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account})).accessToken; }
  catch { return (await msalApp.acquireTokenPopup ({scopes:SCOPES, account})).accessToken; }
}

// ── GRAPH API ────────────────────────────────────────────────────────────────
async function gGet(path) {
  const tok = await getToken();
  const r   = await fetch(API + path, { headers:{ Authorization:'Bearer '+tok } });
  if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
async function gPost(path, body) {
  const tok = await getToken();
  const r   = await fetch(API + path, {
    method:'POST', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (!r.ok) throw new Error(`Graph POST ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
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
  return null;
}

// ── LOAD ITEMS ───────────────────────────────────────────────────────────────
async function loadItems(showToast = true) {
  const btn = $id('btn-reload');
  if (btn) btn.disabled = true;
  try {
    const data = await gGet(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500&$orderby=createdDateTime desc`
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
const VIEW_TITLES = { dashboard:'Dashboard', new:'Neue Bedarfsanfrage',
  mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details' };

function navigate(view, id) {
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

  $id('page-title').textContent = VIEW_TITLES[view] || view;
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

function renderDashboard() {
  renderStatusChips();
  filterDashboard();
}

function renderStatusChips() {
  const el = $id('status-chips');
  if (!el) return;

  // Count per status value
  const counts = {};
  let total = 0;
  let volume = 0;
  for (const item of allItems) {
    const s = getStatusVal(item) || '';
    counts[s] = (counts[s] || 0) + 1;
    total++;
    volume += parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0;
  }

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
  let items = [...allItems];
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) || String(i.id||'').includes(search)
  );
  if (dashStatusFilter !== '') items = items.filter(i => (getStatusVal(i)||'') === dashStatusFilter);
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
function renderList(type) {
  // Populate status filter options once
  const sel = $id(`filter-${type}-status`);
  if (sel.options.length <= 1) {
    const statuses = [...new Set(allItems.map(i => getStatusVal(i)).filter(Boolean))];
    statuses.forEach(s => sel.add(new Option(s, s)));
  }
  filterView(type);
}

function filterView(type) {
  const search = ($id(`search-${type}`)?.value || '').toLowerCase();
  const status = $id(`filter-${type}-status`)?.value || '';
  const userName = (account?.name || account?.username || '').toLowerCase();

  let items = [...allItems];
  if (type === 'mine') {
    items = items.filter(i => {
      const author = (getField(i,'Author') || getField(i,'AuthorLookupId') || '').toString().toLowerCase();
      const creator = (i.createdBy?.user?.displayName || i.createdBy?.user?.email || '').toLowerCase();
      return author.includes(userName.split(' ')[0]) || creator.includes(userName.split(' ')[0]);
    });
  }
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) ||
    String(i.id||'').includes(search)
  );
  if (status) items = items.filter(i => (getStatusVal(i)||'') === status);

  const container = $id(`list-${type}`);
  container.innerHTML = items.length
    ? items.map(i => itemCard(i)).join('')
    : emptyState(type === 'mine' ? 'Sie haben noch keine Anfragen erstellt.' : 'Keine Anfragen gefunden.');
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
        ${isEinkauf ? `<button class="btn btn-sm btn-outline" id="btn-add-order">Bearbeiten</button>` : ''}
      </div>
      <div class="detail-card-body">
        ${orderNr  ? detailRow('Bestellnummer', orderNr) : '<p class="no-order">Noch keine Bestellnummer eingetragen.</p>'}
        ${lieferd  ? detailRow('Lieferdatum', fmtDate(lieferd)) : ''}
        ${tatPreis ? detailRow('Tatsächlicher Preis', fmtEuro(tatPreis)) : ''}
      </div>
    </div>`;
}

// Keywords to detect approval-related SP columns by displayName or internal name
const APPROVAL_RE   = /genehmig|freigab|entscheid|ablehn|kommentar.*genehm|genehm.*kommentar/i;
const STAGE_MAP = [
  { label: 'Einkauf',          re: /einkauf/i },
  { label: 'Verwaltung',       re: /verwaltung/i },
  { label: 'Geschäftsführung', re: /\bgf\b|geschäftsführ|geschaeftsfuehr/i },
];

function approvalStyle(val) {
  const v = (val || '').toLowerCase();
  if (/freigegeben|genehmigt|approved|ja\b/.test(v)) return { bg:'#f0fdf4', color:'#15803d', dot:'✓', cls:'ap-ok' };
  if (/abgelehnt|rejected|nein\b/.test(v))           return { bg:'#fef2f2', color:'#b91c1c', dot:'✗', cls:'ap-no' };
  return { bg:'#fffbeb', color:'#b45309', dot:'…', cls:'ap-pending' };
}

function renderApprovalCard(item) {
  const statusVal = getStatusVal(item) || 'Eingereicht';

  // Find all approval-related columns that have a value on this item
  const found = Object.entries(colByKey)
    .filter(([k, c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
    .map(([k, c]) => ({ key: k, label: c.displayName || k, val: getField(item, k) }))
    .filter(c => c.val !== null && c.val !== undefined && c.val !== '');

  // Group into stages; ungrouped = shown below without stage header
  const stages = STAGE_MAP.map(s => ({
    label: s.label,
    cols:  found.filter(c => s.re.test(c.label) || s.re.test(c.key)),
  })).filter(s => s.cols.length);
  const assigned = new Set(stages.flatMap(s => s.cols.map(c => c.key)));
  const ungrouped = found.filter(c => !assigned.has(c.key));

  const stagesHtml = stages.map(s => {
    // Pick the "decision" field (contains genehmig/entscheid/freigab/ablehn)
    const decisionCol = s.cols.find(c => /genehmig|entscheid|freigab|ablehn/i.test(c.label));
    const otherCols   = s.cols.filter(c => c !== decisionCol);
    const st = decisionCol ? approvalStyle(decisionCol.val) : { bg:'#f3f4f6', color:'#6b7280', dot:'○', cls:'ap-neutral' };
    return `
      <div class="approval-stage">
        <div class="ap-dot ${st.cls}">${st.dot}</div>
        <div class="ap-body">
          <div class="ap-stage-label">${esc(s.label)}</div>
          ${decisionCol ? `<span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(decisionCol.val))}</span>` : ''}
          ${otherCols.map(c => `<div class="ap-meta">${esc(c.label)}: ${esc(String(c.val))}</div>`).join('')}
        </div>
      </div>`;
  }).join('');

  const ungroupedHtml = ungrouped.map(c => {
    const st = approvalStyle(c.val);
    const isComment = /kommentar|ablehn|grund/i.test(c.label);
    if (isComment) return `<div class="ap-comment-box"><strong>${esc(c.label)}:</strong> ${esc(String(c.val))}</div>`;
    return `<div class="approval-stage">
      <div class="ap-dot ${st.cls}">${st.dot}</div>
      <div class="ap-body">
        <div class="ap-stage-label">${esc(c.label)}</div>
        <span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(c.val))}</span>
      </div>
    </div>`;
  }).join('');

  const noData = !stages.length && !ungrouped.length;

  return `
    <div class="detail-card">
      <div class="detail-card-header">Status &amp; Genehmigung</div>
      <div class="detail-card-body">
        <div class="ap-current-status">${statusBadge(statusVal)}</div>
        ${noData
          ? `<p class="ap-empty">Noch keine Genehmigungsdaten — Power Automate aktualisiert diesen Bereich automatisch.</p>`
          : `<div class="approval-stages">${stagesHtml}${ungroupedHtml}</div>`
        }
      </div>
    </div>`;
}

// ── EINKAUF ORDER MODAL ───────────────────────────────────────────────────────
function openOrderModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  const orderNr  = getField(item,'Bestellnummer')      || getField(item, resolvedFields['Bestellnummer'])      || '';
  const lieferd  = getField(item,'Lieferdatum')        || getField(item, resolvedFields['Lieferdatum'])        || '';
  const tatPreis = getField(item,'TatsaechlicherPreis')|| getField(item, resolvedFields['TatsaechlicherPreis'])|| '';

  $id('modal-title').textContent = 'Einkauf-Daten eintragen';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellnummer</label>
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. PO-2025-1234"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferdatum</label>
      <input type="date" id="m-delivery" value="${lieferd ? lieferd.slice(0,10) : ''}"/>
    </div>
    <div class="form-group">
      <label>Tatsächlicher Preis netto (€)</label>
      <input type="number" id="m-price" value="${tatPreis}" min="0" step="0.01" placeholder="0,00"/>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveOrderData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

async function saveOrderData(itemId) {
  const orderNr  = $id('m-ordernr').value.trim();
  const delivery = $id('m-delivery').value;
  const price    = $id('m-price').value;

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

  if (!Object.keys(patch).length) { closeModal(); return; }

  // Retry PATCH removing unrecognized fields
  const skipped = [];
  for (let i = 0; i < 10; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Einkauf-Daten gespeichert ✓', 'success');
      closeModal();
      await loadItems(false);
      renderDetail(itemId);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); return; }
      skipped.push(m[1]);
      delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele nicht erkannte Felder.', 'error');
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
  ['Title','Beschreibung','Warengruppe','Prioritaet','Menge','Mengeneinheit',
   'Mindestlagermenge','Termin','Artikelnummer','Lieferant','Lieferant2','Lieferant3','Lieferant4',
   'GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  const firstRadio = document.querySelector('input[name=Beschaffungslogik]');
  if (firstRadio) firstRadio.checked = true;
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
    wizardData.step3 = {
      Beschaffungslogik: document.querySelector('input[name=Beschaffungslogik]:checked')?.value || '',
      Lieferant:         $id('f-Lieferant').value.trim(),
      Lieferant2:        $id('f-Lieferant2').value.trim(),
      Lieferant3:        $id('f-Lieferant3').value.trim(),
      Lieferant4:        $id('f-Lieferant4').value.trim(),
      GeschaetzterPreis: $id('f-GeschaetzterPreis').value ? parseFloat($id('f-GeschaetzterPreis').value) : null,
      Kostenstelle:      $id('f-Kostenstelle').value.trim(),
    };
    buildReview();
  }
  showStep(step + 1);
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
  const hint  = $id('preis-route-hint');
  if (preis > 0) {
    hint.textContent = genehmigungsweg(preis);
    hint.style.display = 'block';
  } else {
    hint.style.display = 'none';
  }
}

function genehmigungsweg(gesamt) {
  let angebote, freigabe;
  if      (gesamt <   250) { angebote = 'kein Angebot nötig';                         freigabe = 'Einkauf'; }
  else if (gesamt <   750) { angebote = 'mind. 2 Angebote';                            freigabe = 'Einkauf'; }
  else if (gesamt <  1500) { angebote = 'mind. 2 Angebote';                            freigabe = 'Einkauf + Verwaltung'; }
  else if (gesamt < 10000) { angebote = 'mind. 3 Angebote';                            freigabe = 'Einkauf + Verwaltung'; }
  else if (gesamt < 50000) { angebote = 'mind. 3 Angebote';                            freigabe = 'Einkauf + Verwaltung + GF'; }
  else                     { angebote = 'Europ. Ausschreibung (mind. 5 Angebote)';     freigabe = 'Einkauf + Verwaltung + GF'; }
  return `Volumen: ${fmtEuro(gesamt)} · ${angebote} · Freigabe: ${freigabe}`;
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
        ['Gesch. Preis (netto)', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
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
    if (!spCol || val === null || val === undefined || val === '') continue;
    if (DATE_FIELDS.has(fd.key)) {
      const d = toSpDate(val, spCol);
      if (d) fields[spCol] = d;
    } else if (NUMBER_FIELDS.has(fd.key)) {
      const col = colByKey[spCol];
      const n   = parseFloat(val);
      if (!isNaN(n)) {
        // Send as number only if SP column is actually numeric; otherwise string
        fields[spCol] = col?.number ? n : String(n);
      }
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
      await gPost(path, { fields });
      return skipped;
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
}

async function submitRequest() {
  const btn = $id('btn-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';

  try {
    const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
    const fields = buildFields(d, FORM_FIELDS);

    if (!fields['Title']) { toast('Titel fehlt.', 'error'); btn.disabled=false; btn.textContent='✓ Anfrage einreichen'; return; }

    let skipped;
    try {
      skipped = await postRetry(`/sites/${siteId}/lists/${listId}/items`, fields);
    } catch(e404) {
      if (!e404.message.includes('404')) throw e404;
      // Site/list IDs may have changed — re-discover and retry once
      toast('SharePoint-IDs werden aktualisiert…', 'info');
      await discoverSP();
      const fields2 = buildFields(d, FORM_FIELDS);
      skipped = await postRetry(`/sites/${siteId}/lists/${listId}/items`, fields2);
    }

    if (skipped.length) {
      // Map SP column names back to human-readable field labels
      const labels = skipped.map(s => {
        const colName = s.replace(/\(\d+\)$/, '');
        const fd = FORM_FIELDS.find(f => resolvedFields[f.key] === colName || f.key === colName);
        return fd ? fd.label : colName;
      });
      toast(`Eingereicht. Folgende Felder wurden von SharePoint abgelehnt und nicht gespeichert: ${labels.join(', ')}`, 'info');
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
  const appr    = getApprovalSummary(item);

  return `
    <div class="item-card${sel}" data-id="${item.id}" onclick="openPanel('${item.id}')">
      <div class="ic-top">
        <div class="ic-title">${prioDot(prio)}${title}</div>
        <div class="ic-topright">
          ${preis ? `<span class="ic-price">${fmtEuro(preis)}</span>` : ''}
          ${statusBadge(status)}
        </div>
      </div>
      <div class="ic-tags">
        ${wg      ? `<span class="ic-tag ic-wg">${esc(wg)}</span>` : ''}
        ${beschl  ? `<span class="ic-tag">${beschlShort(beschl)}</span>` : ''}
        ${menge && me ? `<span class="ic-tag">⚖ ${esc(menge)} ${esc(me)}</span>` : ''}
        ${ks      ? `<span class="ic-tag">KST ${esc(ks)}</span>` : ''}
        ${liefant ? `<span class="ic-tag">🏭 ${esc(liefant)}</span>` : ''}
        ${termin  ? `<span class="ic-tag">📅 bis ${fmtDate(termin)}</span>` : ''}
      </div>
      <div class="ic-footer">
        <span class="ic-by">${creator ? `👤 ${esc(creator)}` : ''} ${created ? `· ${created}` : ''}</span>
        ${appr}
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
  document.querySelectorAll('.item-card').forEach(c => c.classList.remove('selected'));
}

function bindPanelEvents(itemId) {
  $id('panel-close')?.addEventListener('click', closePanel);
  $id('panel-edit')?.addEventListener('click', () => {
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (!item) return;
    $id(`panel-${currentView}-content`).innerHTML = renderPanel(item, true);
    bindPanelEditEvents(itemId);
  });
  $id('panel-order')?.addEventListener('click', () => openOrderModal(itemId));
}

function bindPanelEditEvents(itemId) {
  $id('panel-close')?.addEventListener('click', closePanel);
  $id('panel-save')?.addEventListener('click', () => saveEdits(itemId));
  $id('panel-cancel')?.addEventListener('click', () => {
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(item); bindPanelEvents(itemId); }
  });
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

function renderPanel(item, editMode = false) {
  const statusVal = getStatusVal(item) || 'Eingereicht';
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
      } else {
        const val = type === 'date' ? dv(raw) : esc(String(raw));
        inp = `<input type="${type}" class="pf-input" data-key="${fd.key}" value="${val}"/>`;
      }
      return `<div class="pf-row">${lbl}${inp}</div>`;
    }
    if (!raw && raw !== 0) return '';
    let display = String(raw);
    if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') display = fmtEuro(raw);
    else if (fd.key === 'Termin' || fd.key === 'Lieferdatum') display = fmtDate(raw);
    return `<div class="pf-row">${lbl}<span class="pf-val">${esc(display)}</span></div>`;
  };

  const preis = parseFloat(gv('GeschaetzterPreis')) || 0;
  const menge = parseFloat(gv('Menge')) || 1;
  const gesamtHint = !editMode && preis > 0 ? genehmigungsweg(preis) : '';

  const orderNr  = gv('Bestellnummer');
  const lieferd  = gv('Lieferdatum');
  const tatPreis = gv('TatsaechlicherPreis');

  const buttons = editMode
    ? `<button class="btn btn-primary btn-sm" id="panel-save">Speichern</button>
       <button class="btn btn-ghost btn-sm" id="panel-cancel">Abbrechen</button>`
    : `<button class="btn btn-outline btn-sm" id="panel-edit">✏ Bearbeiten</button>
       <button class="btn btn-outline btn-sm" id="panel-order">📦 Einkauf</button>`;

  // Approval inner HTML (reuse logic from renderApprovalCard but without the wrapping card)
  const approvalInner = (() => {
    const found = Object.entries(colByKey)
      .filter(([k,c]) => APPROVAL_RE.test(c.displayName||k) && !SYSTEM_FIELDS.has(k))
      .map(([k,c]) => ({ key:k, label:c.displayName||k, val:getField(item,k) }))
      .filter(c => c.val !== null && c.val !== undefined && c.val !== '');
    if (!found.length) return '<p class="ap-empty">Noch keine Genehmigungsdaten.</p>';
    const stages = STAGE_MAP.map(s => ({ label:s.label, cols:found.filter(c=>s.re.test(c.label)||s.re.test(c.key)) })).filter(s=>s.cols.length);
    const assigned = new Set(stages.flatMap(s=>s.cols.map(c=>c.key)));
    const extra = found.filter(c=>!assigned.has(c.key));
    const mkStage = s => {
      const dec = s.cols.find(c=>/genehmig|entscheid|freigab|ablehn/i.test(c.label));
      const others = s.cols.filter(c=>c!==dec);
      const st = dec ? approvalStyle(dec.val) : {bg:'#f3f4f6',color:'#6b7280',dot:'○',cls:'ap-neutral'};
      return `<div class="approval-stage"><div class="ap-dot ${st.cls}">${st.dot}</div><div class="ap-body">
        <div class="ap-stage-label">${esc(s.label)}</div>
        ${dec?`<span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(dec.val))}</span>`:''}
        ${others.map(c=>`<div class="ap-meta">${esc(c.label)}: ${esc(String(c.val))}</div>`).join('')}
      </div></div>`;
    };
    const mkExtra = c => {
      const st = approvalStyle(c.val);
      if (/kommentar|ablehn|grund/i.test(c.label)) return `<div class="ap-comment-box"><strong>${esc(c.label)}:</strong> ${esc(String(c.val))}</div>`;
      return `<div class="approval-stage"><div class="ap-dot ${st.cls}">${st.dot}</div><div class="ap-body"><div class="ap-stage-label">${esc(c.label)}</div><span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(c.val))}</span></div></div>`;
    };
    return stages.map(mkStage).join('') + extra.map(mkExtra).join('');
  })();

  return `
    <div class="panel-hdr">
      <div class="panel-hdr-top">
        <div class="panel-meta">
          <span class="item-id">ID ${item.id}</span>
          ${statusBadge(statusVal)}
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
        ${approvalInner}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Bestellung (Einkauf)</div>
        ${orderNr  ? `<div class="pf-row"><span class="pf-label">Bestellnummer</span><span class="pf-val">${esc(orderNr)}</span></div>` : ''}
        ${lieferd  ? `<div class="pf-row"><span class="pf-label">Lieferdatum</span><span class="pf-val">${fmtDate(lieferd)}</span></div>` : ''}
        ${tatPreis ? `<div class="pf-row"><span class="pf-label">Tatsächl. Preis</span><span class="pf-val">${fmtEuro(tatPreis)}</span></div>` : ''}
        ${!orderNr && !lieferd && !tatPreis ? '<p class="no-order">Noch keine Bestelldaten.</p>' : ''}
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

function statusBadge(s) {
  if (!s) return '';
  const sl = s.toLowerCase().trim();
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
  return `<span class="status-badge" style="background:${st.bg};color:${st.color}">${icon}${esc(s)}</span>`;
}

function prioTag(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const color = p.toLowerCase() === 'dringend' ? '#b91c1c' : '#b45309';
  const bg    = p.toLowerCase() === 'dringend' ? '#fef2f2' : '#fffbeb';
  return `<span class="status-badge" style="background:${bg};color:${color}">${esc(p)}</span>`;
}

function prioDot(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const c = p.toLowerCase() === 'dringend' ? '#ef4444' : '#f59e0b';
  return `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${c};margin-right:5px;vertical-align:middle"></span>`;
}

function fmtEuro(v) {
  if (!v && v !== 0) return '';
  return new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(v);
}
function fmtDate(s) {
  if (!s) return '';
  return new Date(s).toLocaleDateString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric'});
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
