/**
 * Start der App: Laufen unabhängige Abrufe gleichzeitig, und wird das Verzeichnis
 * der Listen nur einmal gelesen?
 *
 * Anlass (16.11.2026, gemessen auf der Live-Seite): Der Start war eine Kette von elf
 * Abrufen nacheinander, gut 3,5 Sekunden bis zur ersten Ansicht. Vier Stellen lasen
 * dabei dasselbe Verzeichnis der Listen, jede mit einem eigenen Abruf.
 *
 *   node tests/test-start.mjs
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const app = fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8');

let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log('  ✓', m); } else { fail++; console.log('  ✗', m); } };

/* 1. Das Verzeichnis der Listen: der echte Code aus app.js, mit einem gezählten Abruf */
const von = app.indexOf('let _listenLauf = null');
const bis = app.indexOf('// Liste anlegen.');
ok(von > 0 && bis > von, 'Der Abschnitt mit dem Verzeichnis der Listen ist in app.js zu finden');
const bau = new Function('gGet', 'siteId', 'uhr', app.slice(von, bis).replace(/Date\.now\(\)/g, 'uhr()')
  + '\nreturn { _alleListen, _listenVergessen, _findListByName };');

let abrufe = 0, jetzt = 1000, scheitern = false;
const gGet = async () => {
  abrufe++;
  await new Promise(r => setTimeout(r, 5));
  if (scheitern) throw new Error('Netz weg');
  return { value: [{ id: 'a1', displayName: 'BANF_Konfiguration' }, { id: 'b2', displayName: ' Artikel ' }] };
};
const m = bau(gGet, 'SITE', () => jetzt);

const [k, a, f, x] = await Promise.all([
  m._findListByName('banf_konfiguration'), m._findListByName('Artikel'), m._findListByName('Favoriten'), m._alleListen(),
]);
ok(abrufe === 1, 'Vier Suchen gleichzeitig: ein Abruf statt vier');
ok(k === 'a1' && a === 'b2' && f === null && x.length === 2, 'Gefunden wird ohne Rücksicht auf Groß- und Kleinschreibung und Leerraum, Fehlendes ist null');
await m._findListByName('Artikel');
ok(abrufe === 1, 'Eine spätere Suche nutzt den gemerkten Stand');
m._listenVergessen();
await m._findListByName('Artikel');
ok(abrufe === 2, 'Nach dem Anlegen einer Liste wird neu gelesen');
jetzt += 61000;
await m._findListByName('Artikel');
ok(abrufe === 3, 'Nach einer Minute wird neu gelesen');
m._listenVergessen(); scheitern = true;
ok(await m._findListByName('Artikel') === null, 'Ein gescheiterter Abruf wirft nicht, die Suche meldet null');
scheitern = false;
ok(await m._findListByName('Artikel') === 'b2' && abrufe === 5, 'Der Fehler wird nicht gemerkt: Der nächste Versuch fragt neu');

/* 2. Verdrahtung in app.js */
ok(app.split('lists?$select=id,displayName').length === 2, 'Das Verzeichnis wird an genau einer Stelle abgerufen');
for (const name of ['_findConfigList', '_findFavList', '_findListByName']) {
  const stelle = app.indexOf('async function ' + name + '(');
  ok(stelle > 0 && app.slice(stelle, stelle + 400).includes('await _alleListen()'), name + ' nutzt das gemeinsame Verzeichnis');
}
for (const name of ['_createList', '_createFavList']) {
  const stelle = app.indexOf('async function ' + name + '(');
  const ende = app.indexOf('\n}', stelle);
  ok(stelle > 0 && app.slice(stelle, ende).includes('finally { _listenVergessen(); }'), name + ' verwirft das Verzeichnis, auch wenn das Anlegen scheitert');
}
const boot = app.slice(app.indexOf('async function bootDone()'), app.indexOf('function startansichtHolen()'));
ok(/const identLauf = loadMyIdentities\(\);\s*await discoverSP\(\);/.test(boot), 'Die Identitäten starten vor dem Erkunden der Site, nicht danach');
ok(boot.includes('await Promise.all([loadItems(false), identLauf, loadAccessConfig(), _spUserLauf]);'),
  'Anfragen, Zugriffssteuerung, Identitäten und Nutzerliste laufen gleichzeitig');
ok(!/await loadItems\(false\);\s*(\/\/[^\n]*\s*)*await loadMyIdentities\(\)/.test(boot), 'Die alte Kette ist weg');
const erkunden = app.slice(app.indexOf('async function discoverSP()'), app.indexOf('async function _ladeSpUserMap()'));
ok(erkunden.indexOf('_spUserLauf = _ladeSpUserMap();') > erkunden.indexOf('siteId = site.id;')
  && erkunden.indexOf('_spUserLauf = _ladeSpUserMap();') < erkunden.indexOf('lists/${encodeURIComponent(SP_LIST)}'),
  'Die Nutzerliste startet, sobald die Site bekannt ist, und wartet nicht auf Liste und Spalten');
ok(!erkunden.includes("lists('User Information List')"), 'discoverSP wartet nicht mehr selbst auf die Nutzerliste');
const nutzer = app.slice(app.indexOf('async function _ladeSpUserMap()'), app.indexOf('// Decode SP hex-encoded field names'));
ok(/try \{[\s\S]*\} catch\(e\) \{[\s\S]*\}\s*\}\s*$/.test(nutzer.trim() + '\n'), 'Die Nutzerliste wirft nie: Ein Fehler dort hält den Start nicht auf');

console.log(`\n${fail ? '✗' : '✓'} ${pass} grün, ${fail} rot`);
process.exit(fail ? 1 : 0);
