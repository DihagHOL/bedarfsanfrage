// Prüft die Werke ohne Browser: Zuordnung über die Anmeldedomäne, Auswahl für Admin und
// Holding-Konten, eigene Site je Werk, kein Archiv und kein Tabellen-Importer bei ZAI,
// Direktlink mit Werk, Auto-Entwürfe für alle Werke.
// Aufruf: node tests/test-werke.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HIER = path.dirname(fileURLToPath(import.meta.url));
const lies = n => fs.readFileSync(path.join(HIER, '..', n), 'utf8');
let fehler = 0;
const pruefe = (ok, text) => { console.log(`  ${ok ? 'ok  ' : 'FEHL'}  ${text}`); if (!ok) fehler++; };
const app = lies('app.js'), html = lies('index.html'), cron = lies('scripts/auto-drafts.mjs');

function funktion(name) {
  let start = app.indexOf(`function ${name}(`);
  if (start < 0) throw new Error('Funktion fehlt: ' + name);
  if (app.slice(start - 6, start) === 'async ') start -= 6;
  const auf = app.indexOf('{', app.indexOf(')', start));
  let tiefe = 0;
  for (let i = auf; i < app.length; i++) {
    if (app[i] === '{') tiefe++;
    else if (app[i] === '}' && --tiefe === 0) return app.slice(start, i + 1);
  }
  throw new Error('Rumpf nicht gefunden: ' + name);
}

const von = app.indexOf('// ── WERKE'), bis = app.indexOf('// ── BPMN-KONFORME');
pruefe(von > 0 && bis > von, 'app.js enthält den Abschnitt Werke');

function neuerKontext({ upn, admin = false, zaiNenntMich = false, start = '', gespeichert = '' }) {
  const sitzung = new Map(start ? [['banf_startwerk', start]] : []);
  const lokal = new Map(gespeichert ? [['banf_werk', gespeichert]] : []);
  const k = vm.createContext({
    account: { username: upn, idTokenClaims: {} }, ACCESS_LIST_NAME: 'BANF_Konfiguration',
    TID_MAP_ACTIVE: { x: 1 }, TID_ENTRIES: [1], KOSTENST_DATA: [1],
    sessionStorage: { getItem: x => sitzung.get(x) ?? null, setItem: (x, v) => sitzung.set(x, v), removeItem: x => sitzung.delete(x) },
    localStorage: { getItem: x => lokal.get(x) ?? null, setItem: (x, v) => lokal.set(x, v) },
    abrufe: [], accessUsers: {}, adminRolle: admin, zaiNenntMich,
    JSON, Object, Array, Set, Map, String, Promise, Date, console,
  });
  vm.runInContext(app.slice(von, bis), k, { filename: 'app.js (Werke)' });
  vm.runInContext(`
    var isAdmin = () => adminRolle;
    var myUPN = () => account.username.toLowerCase();
    var _myIdentities = () => new Set([account.username.toLowerCase()]);
    var myAbteilungen = () => [];
    var GATED_FEATURES = ['dashboard', 'reports', 'importer', 'gemeinsame'];
    var gGet = async (p) => {
      abrufe.push(p);
      if (!p.includes('gruppe_zaigler') && !p.includes('site-zai')) throw new Error('Graph GET 403');
      if (p.endsWith('gruppe_zaigler')) return { id: 'site-zai' };
      if (p.includes('/lists?')) return { value: [{ id: 'cfg', displayName: 'BANF_Konfiguration' }] };
      const users = zaiNenntMich ? { [account.username.toLowerCase()]: { einkauf: true } } : { 'jemand@dihag-zaigler.com': {} };
      return { value: [{ fields: { Title: 'access', ConfigValue: JSON.stringify({ users }) } }] };
    };
    ${funktion('werkFestlegen')}
    ${funktion('werkNenntMich')}
    ${funktion('istArchiv')}
    ${funktion('canAccess')}
  `, k);
  return k;
}
const lauf = async (opt) => { const k = neuerKontext(opt); await vm.runInContext('werkFestlegen(Promise.resolve())', k); return k; };
const wert = (k, ausdruck) => vm.runInContext(ausdruck, k);

console.log('Zuordnung');
let k = neuerKontext({ upn: 'x@shb-guss.de' });
pruefe(wert(k, "werkVonDomain('kollegin@dihag-zaigler.com')") === 'zai', 'dihag-zaigler.com gehört zu ZAI');
pruefe(wert(k, "werkVonDomain('einkauf.test@shb-guss.de')") === 'shb' && wert(k, "werkVonDomain('holding.test@dihag.com')") === 'shb'
  && wert(k, "werkVonDomain('')") === 'shb', 'SHB, Holding und Unbekanntes bleiben bei SHB');
pruefe(wert(k, 'WERKE.zai.site') === 'dihag.sharepoint.com:/sites/gruppe_zaigler' && wert(k, 'WERKE.shb.site') === 'dihag.sharepoint.com:/sites/gruppe_shb',
  'jedes Werk hat seine eigene Site');

console.log('Auswahl');
k = await lauf({ upn: 'Kollegin.Test@dihag-zaigler.com' });
pruefe(wert(k, 'WERK.key') === 'zai' && wert(k, 'JSON.stringify(werkeVerfuegbar)') === '["zai"]', 'ZAI-Konto: nur ZAI');
pruefe(wert(k, 'SP_SITE') === 'dihag.sharepoint.com:/sites/gruppe_zaigler' && wert(k, 'SP_BASE') === 'https://dihag.sharepoint.com/sites/gruppe_zaigler',
  'ZAI liest und schreibt auf gruppe_zaigler, auch über SharePoint-REST (Anhänge)');
pruefe(wert(k, 'TID_ENTRIES.length + KOSTENST_DATA.length') === 0, 'ZAI: keine Artikel- und Kostenstellendaten aus SHB');
k = await lauf({ upn: 'einkauf.test@shb-guss.de', start: 'zai' });
pruefe(wert(k, 'WERK.key') === 'shb' && wert(k, 'abrufe.length') === 0, 'SHB-Konto: Link auf ZAI öffnet trotzdem SHB, ohne Abfrage bei ZAI');
k = await lauf({ upn: 'administrator@dihag.com', admin: true, gespeichert: 'zai' });
pruefe(wert(k, 'JSON.stringify(werkeVerfuegbar)') === '["shb","zai"]' && wert(k, 'WERK.key') === 'zai', 'Admin: alle Werke, zuletzt gewähltes bleibt');
k = await lauf({ upn: 'administrator@dihag.com', admin: true, gespeichert: 'zai', start: 'shb' });
pruefe(wert(k, 'WERK.key') === 'shb', 'Direktlink mit werk=shb geht vor das zuletzt gewählte Werk');
k = await lauf({ upn: 'holding.test@dihag.com', zaiNenntMich: true, start: 'zai' });
pruefe(wert(k, 'JSON.stringify(werkeVerfuegbar)') === '["shb","zai"]' && wert(k, 'WERK.key') === 'zai', 'Holding-Konto in der ZAI-Konfiguration: SHB und ZAI, Direktlink öffnet ZAI');
k = await lauf({ upn: 'holding.ohne@dihag.com', zaiNenntMich: false, start: 'zai' });
pruefe(wert(k, 'JSON.stringify(werkeVerfuegbar)') === '["shb"]' && wert(k, 'WERK.key') === 'shb', 'Holding-Konto ohne ZAI-Eintrag: nur SHB');

console.log('Archiv, Importer, Speicher');
k = await lauf({ upn: 'einkauf.test@shb-guss.de' });
pruefe(wert(k, "istArchiv({ createdDateTime: '2026-11-10T08:00:00Z' })") === true, 'SHB: Anfragen vor dem 16.11.2026 sind Archiv');
pruefe(wert(k, "lsWerk('DIHAG_FAVORITES')") === 'DIHAG_FAVORITES', 'SHB behält die bisherigen Speicherschlüssel');
k = await lauf({ upn: 'administrator@dihag.com', admin: true, gespeichert: 'zai' });
pruefe(wert(k, "istArchiv({ createdDateTime: '2026-11-10T08:00:00Z' })") === false, 'ZAI: kein Archiv');
pruefe(wert(k, "canAccess('importer')") === false && wert(k, "canAccess('reports')") === true, 'ZAI: kein Tabellen-Importer, auch nicht für den Admin');
pruefe(wert(k, "lsWerk('DIHAG_FAVORITES')") === 'DIHAG_FAVORITES_zai', 'ZAI: eigene Speicherschlüssel für Wiedervorlagen und Report-Filter');
k = await lauf({ upn: 'administrator@dihag.com', admin: true });
pruefe(wert(k, "canAccess('importer')") === true, 'SHB: Admin behält den Tabellen-Importer');

console.log('Einbindung');
pruefe(/await werkFestlegen\(identLauf\);\s*await discoverSP\(\);/.test(app), 'Start legt das Werk fest, bevor die Site gelesen wird');
pruefe(/if \(WERK\.stammdaten\) loadLookupDataSP\(\)/.test(app), 'Stammdaten nur in Werken, die sie haben');
pruefe(/if \(view === 'prozess' && !WERK\.prozess\) view = 'mine';/.test(app) && /setVis\('prozess', !!WERK\.prozess\)/.test(app), 'Prozess-Seite nur, wo es Modelle gibt');
pruefe(/q\.get\('werk'\)/.test(app) && /banf_startwerk/.test(app), 'Direktlink ?anfrage=…&werk=zai wird gemerkt');
pruefe(html.includes('id="werk-zeile"'), 'Seitenleiste zeigt das Werk');
pruefe(!/const SP_SITE\s*=/.test(app), 'keine feste Site mehr im Code');
pruefe(cron.includes("sites/gruppe_shb',      pflicht: true") && cron.includes("sites/gruppe_zaigler',  pflicht: false"), 'Auto-Entwürfe laufen für SHB (Pflicht) und ZAI');

console.log(fehler ? `\n${fehler} Prüfung(en) fehlgeschlagen.` : '\nAlle Prüfungen bestanden.');
process.exit(fehler ? 1 : 0);
