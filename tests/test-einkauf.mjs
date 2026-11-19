// Prüft die Änderungen aus der Rückmeldung des Einkaufs vom 19.11.2026 ohne Browser:
// Zuteilung und „nur Zugeteiltes", Reservierung bis zum Abschluss, zusammengeführte Abteilungen,
// wiederkehrender Bedarf im letzten Schritt, Dashboard als Tabelle, ohne Bestellung abschließen,
// Lead-Buyer in der Einkaufsprüfung und die Beschriftung des Liefertermins.
// Aufruf: node tests/test-einkauf.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HIER = path.dirname(fileURLToPath(import.meta.url));
const lies = n => fs.readFileSync(path.join(HIER, '..', n), 'utf8');
let fehler = 0;
const pruefe = (ok, text) => { console.log(`  ${ok ? 'ok  ' : 'FEHL'}  ${text}`); if (!ok) fehler++; };
const app = lies('app.js'), html = lies('index.html'), css = lies('style.css');
const LF = String.fromCharCode(10), Q = String.fromCharCode(39);

// Eine Funktion aus app.js samt Rumpf herauslösen (Klammern zählen)
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

// Kontext mit Attrappen: zwei Einkäufer, einer davon nur mit Zugeteiltem
const STATUS = 'Stauts', LOCK = 'EditLock';
const kontext = vm.createContext({
  account: { username: 'azubi@firma.test', name: 'Azubi Test', idTokenClaims: {} }, LOCK_FIELD: LOCK,
  accessUsers: {
    'azubi@firma.test': { einkauf: true, dashboard: true, nurZugeteilt: true, abteilungen: ['Wareneingang'] },
    'chef@firma.test': { einkauf: true, dashboard: true, name: 'Chef Einkauf', abteilungen: ['Einkauf'] },
    'qs@firma.test': { abteilungen: ['QS'] },
  },
  abteilungen: ['Einkauf/Wareneingang', 'QS'],
  abteilungAlias: { Einkauf: 'Einkauf/Wareneingang', Wareneingang: 'Einkauf/Wareneingang' },
  resolvedFields: { Status: STATUS, Lieferant: 'Lieferant', GeschaetzterPreis: 'GeschaetzterPreis', Termin: 'Termin' },
  allItems: [], wizardData: {}, panelItemId: null, dashSortOrder: 'lieferant', LIST_ROW_CAP: 200,
  rolle: { admin: false, einkauf: true },
  esc: s => String(s ?? ''), cleanTitle: s => String(s ?? ''), prioDot: () => '',
  fmtEuro: v => Number(v).toFixed(2) + ' €', fmtDate: s => String(s).slice(0, 10),
  statusBadge: s => `<b>${s || 'Eingereicht'}</b>`, anzahl: (n, a, b) => `${n} ${n === 1 ? a : b}`,
  toastText: [], hakenAn: false, datumWert: '',
  Date, String, Number, JSON, Object, Array, Set, Map, console,
});
vm.runInContext(`
  var toast = (t) => toastText.push(t);
  var $id = (id) => id === 'f-Wiederkehrend' ? { checked: hakenAn } : id === 'f-AutoBedarfDatum' ? { value: datumWert } : null;
  var isAdmin = () => rolle.admin;
  var canAccess = f => f === 'einkauf' ? rolle.einkauf : false;
  var myWarengruppen = () => [];
  var wgCodeOf = v => String(v).split(' ')[0];
  var getLockField = () => LOCK_FIELD;
  var getField = (i, k) => i && i.fields ? i.fields[k] : undefined;
  var getStatusVal = i => getField(i, 'Stauts') || '';
  var _myIdentities = () => new Set([account.username.toLowerCase()]);
  var reservationOf = () => null; var istArchiv = i => !!i.archiv; var isOverdue = () => false;
`, kontext);
for (const n of ['isOffenStatus', 'parseLock', 'sperreAktiv', 'istMeineSperre', 'meinAnmeldename', 'nurZugeteilt',
  'istMirZugeteilt', 'canApproveItem', 'abteilungKanon', 'myAbteilungen', 'einkaeuferListe', 'wiederkehrendDatum',
  'lieferantVon', 'dashTabelleHtml']) vm.runInContext(funktion(n), kontext, { filename: 'app.js:' + n });
const K = kontext;
const anfrage = (id, status, sperre = '', extra = {}) =>
  ({ id: String(id), createdDateTime: '2026-11-18T08:00:00Z', fields: { [STATUS]: status, [LOCK]: sperre, ...extra }, ...extra });
const morgen = new Date(Date.now() + 864e5).toISOString(), gestern = new Date(Date.now() - 864e5).toISOString();

console.log('Sperre und Reservierung');
const neu = K.parseLock(`Azubi Test|${gestern}|R|AZUBI@firma.test`);
pruefe(neu.kind === 'reserve' && neu.upn === 'azubi@firma.test', 'Reservierung mit Anmeldename, kleingeschrieben');
pruefe(K.parseLock(`Name|${morgen}|R`).upn === '' && K.parseLock(`Name|${morgen}`).kind === 'edit', 'altes Format ohne Anmeldename und Edit-Lock bleiben lesbar');
pruefe(K.sperreAktiv(neu, anfrage(1, 'In Prüfung (Einkauf)')) === true, 'Reservierung gilt über die Ablaufzeit hinaus, solange die Anfrage offen ist');
pruefe(K.sperreAktiv(neu, anfrage(1, 'Freigegeben')) === true, 'auch bei Freigegeben bis zur Bestellung');
pruefe(K.sperreAktiv(neu, anfrage(1, 'Bestellt')) === false && K.sperreAktiv(neu, anfrage(1, 'Erledigt (ohne Bestellung)')) === false, 'nach Bestellung oder Abschluss ohne Bestellung endet sie');
pruefe(K.sperreAktiv(K.parseLock(`X|${gestern}`), anfrage(1, 'Eingereicht')) === false
  && K.sperreAktiv(K.parseLock(`X|${morgen}`), anfrage(1, 'Eingereicht')) === true, 'Edit-Lock endet weiter mit seiner Ablaufzeit');
pruefe(K.istMeineSperre(neu) === true && K.istMeineSperre(K.parseLock(`Chef|${morgen}|R|chef@firma.test`)) === false, 'eigene Sperre über den Anmeldenamen erkannt');
pruefe(K.istMeineSperre(K.parseLock(`Azubi Test|${morgen}|R`)) === true, 'ohne Anmeldename zählt der Name wie bisher');

console.log('Zuteilung und nur Zugeteiltes');
pruefe(K.nurZugeteilt() === true, 'Flag nurZugeteilt wird erkannt');
pruefe(K.canApproveItem(anfrage(2, 'In Prüfung (Einkauf)')) === false, 'nicht zugeteilte Anfrage: keine Bearbeitung');
pruefe(K.canApproveItem(anfrage(3, 'In Prüfung (Einkauf)', `Azubi Test|${morgen}|R|azubi@firma.test`)) === true, 'zugeteilte Anfrage: Bearbeitung erlaubt');
pruefe(K.canApproveItem(anfrage(4, 'In Prüfung (Einkauf)', `Chef Einkauf|${morgen}|R|chef@firma.test`)) === false, 'einem anderen zugeteilt: keine Bearbeitung');
kontext.account = { username: 'chef@firma.test', name: 'Chef Einkauf', idTokenClaims: {} };
pruefe(K.nurZugeteilt() === false && K.canApproveItem(anfrage(2, 'In Prüfung (Einkauf)')) === true, 'normaler Einkäufer bearbeitet weiter alles');
kontext.allItems = [{ createdBy: { user: { email: 'azubi@firma.test', displayName: 'Azubi aus Anfragen' } } }];
const liste = K.einkaeuferListe();
pruefe(liste.length === 2 && liste.map(e => e.name).join(',') === 'Azubi aus Anfragen,Chef Einkauf', 'Einkäuferliste: Namen aus Konfiguration oder Anfragen, sortiert, ohne Nicht-Einkäufer');

console.log('Zusammengeführte Abteilungen');
pruefe(K.abteilungKanon('Wareneingang') === 'Einkauf/Wareneingang' && K.abteilungKanon(' einkauf ') === 'Einkauf/Wareneingang', 'alte Namen landen bei der neuen Gruppe');
pruefe(K.abteilungKanon('QS') === 'QS' && K.abteilungKanon('Einkauf/Wareneingang') === 'Einkauf/Wareneingang' && K.abteilungKanon('Lager') === null, 'angelegte Namen bleiben, Unbekanntes ist null');
pruefe(JSON.stringify(K.myAbteilungen()) === '["Einkauf/Wareneingang"]', 'Zuordnung mit altem Namen zählt für die neue Gruppe');
pruefe(app.includes('zusammengefuehrt: abteilungAlias') && app.includes('cfg.zusammengefuehrt'), 'Laden und Speichern der Abteilungen behalten die Zuordnung');

console.log('Wiederkehrender Bedarf');
const inTagen = n => new Date(Date.now() + n * 864e5).toISOString().slice(0, 10);
kontext.hakenAn = false; pruefe(K.wiederkehrendDatum() === '', 'ohne Haken kein Datum');
kontext.hakenAn = true; kontext.datumWert = ''; pruefe(K.wiederkehrendDatum() === false, 'Haken ohne Datum wird gemeldet');
kontext.datumWert = inTagen(-3); pruefe(K.wiederkehrendDatum() === false, 'Datum in der Vergangenheit wird gemeldet');
kontext.datumWert = inTagen(20); kontext.wizardData = { step2: { Termin: inTagen(20) } };
pruefe(K.wiederkehrendDatum() === false && kontext.toastText.at(-1).includes('Liefertermin'), 'gleich dem Liefertermin wird abgewiesen');
kontext.wizardData = { step2: { Termin: inTagen(10) } }; pruefe(K.wiederkehrendDatum() === inTagen(20), 'späteres Datum wird übernommen');
const schritt1 = html.slice(html.indexOf('id="wstep-1"'), html.indexOf('id="wstep-2"'));
const schritt4 = html.slice(html.indexOf('id="wstep-4"'), html.indexOf('id="mwstep-1"'));
pruefe(!schritt1.includes('f-AutoBedarfDatum') && schritt4.includes('id="f-Wiederkehrend"') && schritt4.includes('id="f-AutoBedarfDatum"'), 'Datum steht nicht mehr in Schritt 1, sondern zum Ankreuzen im letzten Schritt');
pruefe(app.split(LF).some(l => l.includes("key:'AutoBedarfDatum'") && l.includes("'DatumzurAutomatischeBedarfsanfra'")), 'interner Spaltenname des Datums steht in alsoTry');
pruefe(/abgelehnt[|]zur/.test(lies('scripts/auto-drafts.mjs')), 'Auto-Entwurf-Lauf überspringt abgelehnte und zurückgezogene Anfragen');

console.log('Dashboard als Tabelle');
const tab = K.dashTabelleHtml([
  anfrage(10, 'In Prüfung (Einkauf)', '', { Lieferant: 'Würth', GeschaetzterPreis: 100 }),
  anfrage(11, 'Bestellt', '', { Lieferant: 'würth', GeschaetzterPreis: 50, archiv: true }),
  anfrage(12, 'Eingereicht', '', { Lieferant: '', GeschaetzterPreis: 0 }),
], true);
pruefe((tab.match(/class="dt-gruppe"/g) || []).length === 2, 'eine Kopfzeile je Lieferant, Groß- und Kleinschreibung egal');
pruefe(tab.includes('2 Anfragen, davon 1 offen') && tab.includes('150.00 €'), 'Kopfzeile mit Anzahl, offenen und Volumen');
pruefe(tab.includes('Ohne Lieferant') && (tab.match(/class="dt-zeile/g) || []).length === 3, 'Anfragen ohne Lieferant bekommen eine eigene Gruppe');
pruefe(tab.includes('dt-zeile is-archiv') && tab.includes('openPanel(' + Q + '11' + Q + ')'), 'Archiv grau, Klick öffnet das Panel');
pruefe(html.includes('value="lieferant"') && html.includes("setDashAnsicht('tabelle')") && css.includes('.dash-tabelle'), 'Sortierung nach Lieferant, Umschalter und Stile vorhanden');

console.log('Einkauf: ohne Bestellung, Lead-Buyer, Liefertermin');
pruefe(app.includes("'erledigt (ohne bestellung)'") && app.includes('function obStatusWahl') && app.includes('OB_PRAEFIX'), 'Status Erledigt (ohne Bestellung) mit Stil und Begründung im Kommentar');
pruefe(app.includes('if (/^erledigt/i.test(sv))'), 'Statusverlauf zeigt den Abschluss ohne Bestellung');
pruefe(K.isOffenStatus('Erledigt (ohne Bestellung)') === false && K.isOffenStatus('Freigegeben') === true, 'Erledigt zählt nicht als offen');
pruefe(['#btn-ob-ok', '#btn-zuteilen-ok', '#aab-leadbuyer'].every(s => app.includes(`pq('${s}')`)), 'Panel bindet Abschluss, Zuteilen und Lead-Buyer');
pruefe(app.includes('id="aab-leadbuyer"') && app.includes('data-key="LeadBuyerAbschluss"'), 'Lead-Buyer im Einkaufsbanner und beim Bearbeiten');
pruefe(app.includes("inp.type === 'checkbox' ? inp.checked : inp.value"), 'Speichern liest Häkchen als ja oder nein');
pruefe(/key:'Termin', +label:'Liefertermin'/.test(app) && html.includes('Liefertermin (spätestens)') && !html.includes('Benötigt bis'), 'Liefertermin klar beschriftet, „Benötigt bis" ist aus der Oberfläche raus');
pruefe(app.includes("q.get('anfrage')") && app.includes("navigate('detail', _startAnfrage)"), 'Direktlink ?anfrage= öffnet die Anfrage');
pruefe(app.includes("'nurZugeteilt',this.checked"), 'Einstellungen: Häkchen „nur Zugeteiltes" je Einkäufer');

console.log(fehler ? `${LF}${fehler} Prüfung(en) fehlgeschlagen.` : `${LF}Alle Prüfungen bestanden.`);
process.exit(fehler ? 1 : 0);
