// Prüft den Produktivstart ohne Browser: Stichtag, Archiv-Kennzeichnung, „überfällig" nur
// für Anfragen nach dem Stichtag, die Trennlinie in den Listen und dass keine Marke des
// Testsystems mehr in der Oberfläche steht.
// Aufruf: node tests/test-archiv.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HIER = path.dirname(fileURLToPath(import.meta.url));
const lies = n => fs.readFileSync(path.join(HIER, '..', n), 'utf8');
let fehler = 0;
const pruefe = (ok, text) => { console.log(`  ${ok ? 'ok  ' : 'FEHL'}  ${text}`); if (!ok) fehler++; };

const app = lies('app.js');
const von = app.indexOf('// ── PRODUKTIVSTART UND ARCHIV');
const bis = app.indexOf('// Aktive Reservierung einer Anfrage');
pruefe(von > 0 && bis > von, 'app.js enthält den Abschnitt Produktivstart und Archiv');

// Feste Uhr: Freitag, 20.11.2026, 13 Uhr deutscher Zeit
const JETZT = Date.UTC(2026, 10, 20, 12, 0, 0);
const kontext = vm.createContext({
  Date: class extends Date { static now() { return JETZT; } },
  esc: s => String(s),
  getStatusVal: item => item.status || '',
});
vm.runInContext(app.slice(von, bis), kontext, { filename: 'app.js (Ausschnitt)' });
const A = vm.runInContext('({ PRODUKTIV_AB, PRODUKTIV_AB_TEXT, ARCHIV_TRENNER, istArchiv, isOverdue, archivTrennerVor, archivMarke })', kontext);
const anfrage = (iso, status = 'Freigegeben') => ({ createdDateTime: iso, status });

console.log('Stichtag');
pruefe(Date.parse(A.PRODUKTIV_AB) === Date.UTC(2026, 10, 15, 23, 0, 0), 'Produktivstart ist der 16.11.2026, 0 Uhr deutscher Zeit');
const [jahr, monat, tag] = A.PRODUKTIV_AB.slice(0, 10).split('-');
pruefe(A.PRODUKTIV_AB_TEXT === `${tag}.${monat}.${jahr}`, 'angezeigtes Datum passt zum Stichtag');
pruefe(A.istArchiv(anfrage('2026-11-15T22:59:59Z')) === true, 'Anfrage vom 15.11., 23:59 Uhr ist Archiv');
pruefe(A.istArchiv(anfrage('2026-11-15T23:00:00Z')) === false, 'Anfrage vom 16.11., 0:00 Uhr zählt zum Produktivbetrieb');
pruefe(A.istArchiv(anfrage('2026-11-16T06:00:10Z')) === false, 'Anfrage vom 16.11., 7 Uhr zählt zum Produktivbetrieb');
pruefe(A.istArchiv(anfrage('2026-06-03T08:00:00Z')) === true, 'Anfrage aus dem Juni ist Archiv');
pruefe(A.istArchiv({}) === false && A.istArchiv(null) === false, 'ohne Anlagedatum kein Archiv');

console.log('Überfällig');
pruefe(A.isOverdue(anfrage('2026-11-10T08:00:00Z')) === false, 'offene Archiv-Anfrage gilt nicht als überfällig');
pruefe(A.isOverdue(anfrage('2026-11-16T08:00:00Z')) === true, 'neue Anfrage, seit vier Tagen offen, ist überfällig');
pruefe(A.isOverdue(anfrage('2026-11-19T12:00:00Z')) === false, 'neue Anfrage, seit 24 Stunden offen, ist es nicht');
pruefe(A.isOverdue(anfrage('2026-11-16T08:00:00Z', 'Bestellt')) === false, 'bestellte Anfrage ist nie überfällig');
pruefe(A.isOverdue(anfrage('2026-11-16T08:00:00Z', '')) === true, 'neue Anfrage mit leerem Status (Eingereicht, noch nicht geöffnet) wird überfällig');
pruefe(A.isOverdue(anfrage('2026-11-16T08:00:00Z', 'Zurückgezogen')) === false, 'zurückgezogene Anfrage ist nicht mehr offen');
pruefe(A.isOverdue(anfrage('2026-11-16T08:00:00Z', 'Automatisch erstellter Entwurf')) === false, 'Entwurf ist nicht offen');

console.log('Trennlinie');
const neu = anfrage('2026-11-17T08:00:00Z'), alt = anfrage('2026-11-12T08:00:00Z');
pruefe(A.archivTrennerVor([neu, neu, alt, alt]) === 2, 'neueste zuerst: Linie vor der ersten Archiv-Anfrage');
pruefe(A.archivTrennerVor([alt, alt]) === 0, 'nur Archiv: Linie ganz oben');
pruefe(A.archivTrennerVor([neu, neu]) === -1 && A.archivTrennerVor([]) === -1, 'ohne Archiv keine Linie');
pruefe(A.archivTrennerVor([alt, neu]) === -1 && A.archivTrennerVor([neu, alt, neu]) === -1, 'gemischte Reihenfolge: keine Linie');
pruefe(A.ARCHIV_TRENNER.includes('Produktivstart 16.11.2026') && A.ARCHIV_TRENNER.includes('bearbeitbar'), 'Linie nennt den Produktivstart und dass Archiv bearbeitbar bleibt');
pruefe(A.archivMarke('ic-tag').includes('class="ic-tag marke-archiv"') && A.archivMarke('ic-tag').includes('Archiv'), 'Marke Archiv für Karten und Kopfzeilen');

console.log('Oberfläche');
const html = lies('index.html'), css = lies('style.css'), prozess = lies('prozess.js');
pruefe(!/testumgebung|env-test|boot-env|nav-beta/i.test(html), 'index.html ohne Testumgebung und Beta-Marke');
pruefe(!/[.]env-test|[.]boot-env|[.]nav-beta/.test(css), 'style.css ohne Stile des Testsystems');
pruefe(!app.includes('Schulung (Beta)') && !app.includes('schulung-badge">Beta'), 'Schulung ohne Beta');
pruefe(app.includes("${archiv ? ' is-archiv' : ''}") && app.includes("archivMarke('ic-tag')") && app.includes("archivMarke('cr2-tag')"), 'Karten und Zeilen tragen Klasse und Marke');
pruefe((app.match(/archivMarke[(]'status-badge'[)]/g) || []).length === 2, 'Detailansicht und Panel zeigen die Marke');
pruefe(app.includes("key:'archiv'") && app.includes("' rep-archiv'"), 'Reports mit Spalte Archiv und grauen Zeilen');
pruefe(css.includes('.item-card.is-archiv') && css.includes('.archiv-trenner') && css.includes('.rep-row.rep-archiv'), 'Stile für Archiv vorhanden');
const schulung = app.slice(app.indexOf('function initSchulung()'), app.indexOf('// ── SAMMELANFRAGE (MULTI-POSITION WIZARD)'));
pruefe(schulung.includes('Marke <b>Archiv</b>') && schulung.includes('Produktivbetrieb') && !schulung.includes('Zurücksetzen'), 'Schulung erklärt Archiv und Produktivbetrieb, ohne den alten Knopf Zurücksetzen');
pruefe(!schulung.includes(String.fromCharCode(8211)), 'Schulung ohne Gedankenstriche');
pruefe(prozess.includes(`Produktivstart am ${A.PRODUKTIV_AB_TEXT}`) && prozess.includes(`Seit dem ${A.PRODUKTIV_AB_TEXT}`), 'Prozess-Seite nennt denselben Stichtag');

console.log(fehler ? `${String.fromCharCode(10)}${fehler} Prüfung(en) fehlgeschlagen.` : `${String.fromCharCode(10)}Alle Prüfungen bestanden.`);
process.exit(fehler ? 1 : 0);
