// ─────────────────────────────────────────────────────────────────────────────
// Automatische Bedarfsanfrage-Entwürfe (App-only Cron, läuft via GitHub Actions)
//
// Für jede Anfrage, deren Feld „Datum zur Automatische Bedarfsanfrage" fällig ist
// (heute oder früher), wird eine Kopie als Entwurf mit Status
// „Automatisch erstellter Entwurf" angelegt und das Quelldatum geleert
// (verhindert erneutes Auslösen).
//
// Auth: app-only (Client Credentials). Benötigt Microsoft-Graph-Application-
// Permission „Sites.Selected" + per-Site-Schreibrecht auf gruppe_shb.
// Secrets (GitHub Actions): TENANT_ID, CLIENT_ID, CLIENT_SECRET.
// ─────────────────────────────────────────────────────────────────────────────

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const GRAPH                = 'https://graph.microsoft.com/v1.0';
const SITE_PATH            = 'dihag.sharepoint.com:/sites/gruppe_shb';
const LIST_NAME            = 'Bedarfsanfrage';
const DRAFT_STATUS         = 'Automatisch erstellter Entwurf';
const DATE_FIELD_DISPLAY   = 'Datum zur Automatische Bedarfsanfrage';

// Anzeigenamen-Teilstrings, die NICHT in den Entwurf kopiert werden
// (Status, Genehmiger/Entscheider, Einkauf-Daten, Bearbeitungssperre, Erinnerungsdatum):
const EXCLUDE_DISPLAY = [
  'status', 'stauts', 'genehmig', 'entscheid', 'bestellnummer', 'lieferdatum',
  'tatsächlich', 'tatsaechlich', 'datum zur automatisch', 'bearbeitet', 'sperre', 'lock',
];

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error('❌ Fehlende Secrets: TENANT_ID / CLIENT_ID / CLIENT_SECRET'); process.exit(1);
}

const norm = s => (s || '').toLowerCase().replace(/[^a-z0-9äöü]/g, '');

async function getToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials',
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body,
  });
  if (!r.ok) throw new Error('Token-Abruf fehlgeschlagen: ' + r.status + ' ' + await r.text());
  return (await r.json()).access_token;
}

let TOKEN;
async function g(method, path, body) {
  const r = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    method,
    headers: { Authorization: 'Bearer ' + TOKEN, 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (r.status === 204) return {};
  const txt = await r.text();
  if (!r.ok) throw new Error(`${method} ${path} → ${r.status}: ${txt}`);
  return txt ? JSON.parse(txt) : {};
}
async function gAll(path) {
  const out = []; let next = path;
  while (next) { const d = await g('GET', next); out.push(...(d.value || [])); next = d['@odata.nextLink'] || null; }
  return out;
}

(async () => {
  TOKEN = await getToken();

  const site = await g('GET', `/sites/${SITE_PATH}`);
  const list = await g('GET', `/sites/${site.id}/lists/${encodeURIComponent(LIST_NAME)}`);
  const cols = (await g('GET', `/sites/${site.id}/lists/${list.id}/columns`)).value || [];

  // Spalten auflösen
  const dateCol = cols.find(c => norm(c.displayName) === norm(DATE_FIELD_DISPLAY))
              || cols.find(c => { const d = norm(c.displayName); return d.includes('automatisch') && d.includes('datum'); });
  const statusCol = cols.find(c => ['stauts', 'status'].includes((c.name || '').toLowerCase()))
                 || cols.find(c => (c.displayName || '').toLowerCase() === 'status');
  if (!dateCol)   { console.error(`❌ Datumsspalte "${DATE_FIELD_DISPLAY}" nicht gefunden.`); process.exit(1); }
  if (!statusCol) { console.error('❌ Status-Spalte nicht gefunden.'); process.exit(1); }

  // Zu kopierende (schreibbare) Spalten – ohne Status/Datum/ausgeschlossene
  const copyCols = cols.filter(c => {
    if (c.readOnly || c.hidden) return false;
    if (c.name === dateCol.name || c.name === statusCol.name) return false;
    const dn = (c.displayName || '').toLowerCase();
    return !EXCLUDE_DISPLAY.some(x => dn.includes(x));
  });

  const items = await gAll(`/sites/${site.id}/lists/${list.id}/items?$expand=fields&$top=200`);
  const todayEnd = new Date(); todayEnd.setHours(23, 59, 59, 999);

  const due = items.filter(it => {
    const f = it.fields || {};
    const d = f[dateCol.name];
    if (!d) return false;
    const dd = new Date(d);
    if (isNaN(dd.getTime()) || dd > todayEnd) return false;
    const st = String(f[statusCol.name] || '').toLowerCase();
    return !/entwurf|automatisch erstellt/.test(st); // Entwürfe nicht erneut kopieren
  });

  console.log(`ℹ️  ${items.length} Elemente geladen, ${due.length} fällig.`);
  let created = 0;
  for (const src of due) {
    try {
      const f = src.fields || {};
      const newFields = {};
      for (const c of copyCols) {
        const v = f[c.name];
        if (v == null || v === '' || typeof v === 'object') continue; // Personen/Arrays/Objekte überspringen
        newFields[c.name] = v;
      }
      newFields.Title = 'Auto-Entwurf: ' + (f.Title || 'Bedarfsanfrage');
      newFields[statusCol.name] = DRAFT_STATUS;

      await g('POST', `/sites/${site.id}/lists/${list.id}/items`, { fields: newFields });
      await g('PATCH', `/sites/${site.id}/lists/${list.id}/items/${src.id}/fields`, { [dateCol.name]: null });
      created++;
      console.log(`✓ Entwurf aus #${src.id} erstellt.`);
    } catch (e) {
      console.error(`✗ #${src.id}: ${e.message}`);
    }
  }
  console.log(`✅ Fertig: ${created} Entwurf/Entwürfe erstellt.`);
})().catch(e => { console.error(e); process.exit(1); });
