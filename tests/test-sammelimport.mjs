// Prüft sammelimport.js ohne Browser: Kopfzeile des Vorschlags, Zuordnung der Werte,
// Bündelung zu Sammelanfragen, Fehlermeldungen, CSV-Leser und die erzeugte Excel-Vorlage.
// Aufruf: node tests/test-sammelimport.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HIER = path.dirname(fileURLToPath(import.meta.url));
const lies = n => fs.readFileSync(path.join(HIER, '..', n), 'utf8');
let fehler = 0;
const pruefe = (ok, text) => { console.log(`  ${ok ? 'ok  ' : 'FEHL'}  ${text}`); if (!ok) fehler++; };

// Warengruppen wie in app.js (aus dem Quelltext gelesen, damit die Liste nicht doppelt gepflegt wird)
const app = lies('app.js');
const wgQuelle = app.match(/const WG_FALLBACK = (\[[^\]]*\]);/);
pruefe(!!wgQuelle, 'app.js enthält WG_FALLBACK');
const kontext = vm.createContext({
  module: { exports: {} }, TextEncoder, TextDecoder, console,
  WG_FALLBACK: JSON.parse(wgQuelle[1].replace(/'/g, '"')),
  abteilungen: ['Instandhaltung', 'Modellbau', 'techn. Dienste', 'AV'],
  KOSTENST_DATA: [{ nr: '2011', label: 'Kostenstelle B' }, { nr: '1111', label: 'Kostenstelle A' }],
  colByKey: {}, resolvedFields: {},
});
vm.runInContext(lies('sammelimport.js'), kontext, { filename: 'sammelimport.js' });
const S = kontext.module.exports;

console.log('Kopfzeile');
const kopfVorschlag = ['Menge', 'VP', 'Hersteller', 'Art-Nr / Best-Nr', 'Typ-Nr', 'Bezeichnung / Beschreibung des Artikel',
  'Begründung des Bedarfes', 'Kostenstelle', 'Warengruppe', 'Abteilung', 'Beschaffungsart 1', 'Beschaffungsart 2', 'Lieferrant', 'Einzelpreis'];
const kopf = S.siKopfFinden([[], kopfVorschlag]);
pruefe(kopf && kopf.index === 1, 'findet die Kopfzeile auch unter einer Leerzeile');
pruefe(kopf && Object.keys(kopf.spalten).length === 14, 'erkennt alle 14 Spalten des Vorschlags (VP, Lieferrant, des Artikel)');
const kopfVorlage = S.siKopfFinden([S.SI_SPALTEN.map(s => s.titel)]);
pruefe(kopfVorlage && Object.keys(kopfVorlage.spalten).length === S.SI_SPALTEN.length, 'erkennt alle Spalten der eigenen Vorlage');

console.log('Werte');
const wg = v => S.siWarengruppe(v);
pruefe(wg('Reparatur').wert === '4334 – Reparatur', 'Warengruppe nur mit Namen');
pruefe(wg('4334').wert === '4334 – Reparatur', 'Warengruppe nur mit Code');
pruefe(wg('Transporte').wert === '4333 – Transporte', 'Warengruppe mit doppeltem Code über den Namen');
pruefe(/nicht eindeutig/.test(wg('4333').fehler || ''), '4333 allein ist nicht eindeutig');
pruefe(/gibt es nicht/.test(wg('Elektro').fehler || ''), 'unbekannte Warengruppe wird gemeldet');
pruefe(S.siAbteilung('Instandh.').wert === 'Instandhaltung', 'Abteilung als Kurzform');
pruefe(S.siAbteilung('Technische Dienste').wert === 'techn. Dienste', 'Abteilung ausgeschrieben');
pruefe(!!S.siAbteilung('Lager').fehler && S.siAbteilung('').wert === '', 'unbekannte Abteilung ist ein Fehler, leer ist erlaubt');
pruefe(S.siEinheit('Stck').wert === 'Stück' && S.siEinheit('VP').wert === 'VPE' && S.siEinheit('Rolle').wert === 'Rolle', 'Einheiten: Stck, VP, eigene');
pruefe(S.siArt1('Nicht-Bestandsgeführt').wert === S.SI_ART1[1] && S.siArt1('Bestandsgeführt').wert === S.SI_ART1[0], 'Beschaffungsart 1');
pruefe(S.siArt2('indirektes Material').wert === S.SI_ART2[1] && S.siArt2('Direktes Material').wert === S.SI_ART2[0]
  && S.siArt2('Dienstleistung').wert === S.SI_ART2[1], 'Beschaffungsart 2');
pruefe(S.siZahl('84,00 €') === 84 && S.siZahl('0,0267') === 0.0267 && S.siZahl('1.234,50') === 1234.5 && S.siZahl(7) === 7, 'Zahlen deutsch und aus Excel');
const serie = (Date.UTC(2026, 11, 31) - Date.UTC(1899, 11, 30)) / 864e5;
pruefe(S.siDatum('31.12.2026').wert === '2026-12-31' && S.siDatum(serie).wert === '2026-12-31' && S.siDatum('').wert === '', 'Datum als Text und als Excel-Zahl');
pruefe(!!S.siDatum('31.02.2026').fehler, 'unmögliches Datum wird gemeldet');

console.log('Auswerten');
const zeile = (menge, vp, bez, kst, preis, lief = 'Lieferant X') => [menge, vp, 'Hersteller A', 4099854292750, 'Typ 1', bez,
  'Halle 2', kst, 'Reparatur', 'Instandhaltung', 'Nicht-Bestandsgeführt', 'indirektes Material', lief, preis];
const e1 = S.siAuswerten([kopfVorschlag, zeile(10, 'Stck', 'Deckenleuchte', '000123', 84), zeile(1, 'Stck', 'Spiralkabel', 4711, 84),
  [], zeile(5, 'VPE', 'Aderendhülse (1 VPE = 100)', 4711, 0.0267)], 'vorschlag.xlsx');
pruefe(e1.fehler.length === 0, 'Beispiel wie im Vorschlag ohne Fehler');
pruefe(e1.gruppen.length === 1 && e1.positionen === 3, 'drei Positionen, eine Sammelanfrage, Leerzeile übersprungen');
const g = e1.gruppen[0] || { positionen: [] };
pruefe(g.warengruppe === '4334 – Reparatur' && g.abteilung === 'Instandhaltung' && g.lieferant === 'Lieferant X', 'Warengruppe, Abteilung, Lieferant der Gruppe');
pruefe(g.beschaffungslogik === 'Nicht-bestandsgeführtes Material, Indirektes Material / Dienstleistung', 'Beschaffungslogik wie im Assistenten');
pruefe(g.summe === 924.13, `Bestellvolumen 924,13 € (ist ${g.summe})`);
pruefe(g.positionen.map(p => p.me).join() === 'Stück,Stück,VPE', 'Einheiten der Positionen');
pruefe(g.positionen[0].kostenstelle === '000123' && g.positionen[1].kostenstelle === '4711', 'Kostenstelle behält führende Nullen');
pruefe(g.positionen[0].extArtNr === '4099854292750' && g.positionen[0].hersteller === 'Hersteller A' && g.positionen[0].typNr === 'Typ 1', 'Art-Nr, Hersteller, Typ-Nr');
pruefe(g.beschreibung === 'Halle 2', 'Beschreibung aus den Begründungen');

const e2 = S.siAuswerten([kopfVorschlag, zeile(1, 'Stck', 'A', '', 10), zeile(2, 'Stck', 'B', '', 20, 'Lieferant Y')], 'zwei.xlsx');
pruefe(e2.gruppen.length === 2, 'anderer Lieferant ergibt eine zweite Sammelanfrage');

const kaputt = [...zeile(0, 'Stck', '', '', 'abc', '')];
kaputt[8] = 'Elektro';
const e3 = S.siAuswerten([kopfVorschlag, zeile(1, 'Stck', 'Gut', '', 5), kaputt], 'fehler.xlsx');
const f3 = e3.fehler.join(' ');
pruefe(e3.fehler.length === 1 && /^Zeile 3:/.test(e3.fehler[0]), 'Fehler mit Zeilennummer wie in Excel');
pruefe(['Menge', 'Bezeichnung fehlt', 'Elektro', 'Lieferant fehlt', 'abc'].every(t => f3.includes(t)), 'meldet Menge, Bezeichnung, Warengruppe, Lieferant und Preis');
const e4 = S.siAuswerten([['Menge', 'Bezeichnung', 'Warengruppe']], 'wenig.csv');
pruefe(/Spalten fehlen/.test(e4.fehler[0] || ''), 'fehlende Pflichtspalten werden genannt');
pruefe(/Kopfzeile/.test(S.siAuswerten([['a', 'b']], 'x.csv').fehler[0] || ''), 'ohne Kopfzeile verständliche Meldung');

console.log('CSV');
const LF = String.fromCharCode(10), CR = String.fromCharCode(13), BOM = String.fromCharCode(65279);
const csv = BOM + kopfVorschlag.join(';') + CR + LF
  + '3;Stck;"Firma; mit Semikolon";123;;"Kabel ""rot""";Grund;000123;4334 – Reparatur;Instandhaltung;Bestandsmaterial (bestandsgeführt);Direktes Material;Lieferant X;12,50' + CR + LF;
const zeilenCsv = S.siCsvZeilen(csv);
pruefe(zeilenCsv.length === 2 && zeilenCsv[1][2] === 'Firma; mit Semikolon' && zeilenCsv[1][5] === 'Kabel "rot"', 'Anführungszeichen, Semikolon im Feld, BOM, CRLF');
const e5 = S.siAuswerten(zeilenCsv, 'liste.csv');
pruefe(e5.fehler.length === 0 && e5.gruppen[0]?.summe === 37.5, 'CSV-Zeile wird zur Position (3 × 12,50 €)');
pruefe(S.siVorlageCsvText().startsWith(BOM + 'Menge;Einheit (VP);Hersteller'), 'CSV-Vorlage mit BOM und Kopfzeile');

console.log('Excel-Vorlage');
const zip = S.siVorlageXlsx(new Date(2026, 10, 12));
const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
const endeAb = zip.length - 22;
pruefe(dv.getUint32(endeAb, true) === 0x06054b50, 'ZIP endet mit dem Endeintrag');
const anzahlDateien = dv.getUint16(endeAb + 10, true);
let p = dv.getUint32(endeAb + 16, true);
const teile = {};
for (let i = 0; i < anzahlDateien; i++) {
  const nameLaenge = dv.getUint16(p + 28, true), versatz = dv.getUint32(p + 42, true), crc = dv.getUint32(p + 16, true);
  const name = new TextDecoder().decode(zip.subarray(p + 46, p + 46 + nameLaenge));
  const groesse = dv.getUint32(versatz + 18, true), lokalName = dv.getUint16(versatz + 26, true);
  const daten = zip.subarray(versatz + 30 + lokalName, versatz + 30 + lokalName + groesse);
  teile[name] = { text: new TextDecoder().decode(daten), crcOk: S.siCrc32(daten) === crc };
  p += 46 + nameLaenge;
}
const erwartet = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
  'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml'];
pruefe(erwartet.every(n => teile[n]), 'alle Teile der Arbeitsmappe vorhanden');
pruefe(Object.values(teile).every(t => t.crcOk), 'Prüfsummen stimmen');
const blatt1 = teile['xl/worksheets/sheet1.xml']?.text || '';
pruefe(S.SI_SPALTEN.every(s => blatt1.includes(`<t xml:space="preserve">${s.titel.replace(/&/g, '&amp;')}</t>`)), 'Kopfzeile mit allen Spalten');
pruefe((blatt1.match(/<dataValidation /g) || []).length === 9 && blatt1.includes('<dataValidations count="9">'), 'neun Prüfungen und Auswahllisten');
pruefe(blatt1.includes('state="frozen"'), 'Kopfzeile fixiert');
const mappe = teile['xl/workbook.xml']?.text || '';
pruefe(['Liste_Warengruppe', 'Liste_Abteilung', 'Liste_Kostenstelle', 'Liste_Einheit'].every(n => mappe.includes(`name="${n}"`)), 'benannte Auswahllisten');
pruefe(mappe.includes('name="Listen" sheetId="3" state="hidden"'), 'Blatt Listen ist ausgeblendet');
const listen = teile['xl/worksheets/sheet3.xml']?.text || '';
pruefe(listen.includes('techn. Dienste') && listen.includes('1111 – Kostenstelle A'), 'Listen enthalten Abteilungen und Kostenstellen des Augenblicks');
pruefe(listen.indexOf('1111 – Kostenstelle A') < listen.indexOf('2011 – Kostenstelle B'), 'Kostenstellen nach Nummer sortiert');
pruefe((teile['xl/worksheets/sheet2.xml']?.text || '').includes('12.11.2026'), 'Hinweise nennen den Stand der Listen');
pruefe((teile['xl/worksheets/sheet2.xml']?.text || '').includes('„Übernehmen“'), 'Hinweise erklären das Übernehmen aus dem OneDrive');
const ausgewogen = t => (t.match(/</g) || []).length === (t.match(/>/g) || []).length;
pruefe(Object.values(teile).every(t => t.text.startsWith('<?xml') && ausgewogen(t.text)), 'jedes Teil beginnt mit XML-Kopf, spitze Klammern ausgeglichen');

console.log('OneDrive');
pruefe(S.siVorlageName(new Date(2026, 10, 12, 13, 5)) === 'Sammelanfrage 12.11.2026 13.05.xlsx', 'Dateiname mit Datum und Uhrzeit');
pruefe(S.siVorlageName(new Date(2026, 10, 12, 13, 5), 2) === 'Sammelanfrage 12.11.2026 13.05 (2).xlsx', 'zweite Vorlage derselben Minute');
const docUrl = 'https://firma-my.sharepoint.com/personal/x/_layouts/15/Doc.aspx?sourcedoc=%7BAB12%7D&file=Sammelanfrage%2012.11.2026.xlsx&action=default&mobileredirect=true';
pruefe(S.siBearbeitenUrl({ webUrl: docUrl }) === docUrl.replace('action=default', 'action=edit'), 'Excel im Browser öffnet zum Bearbeiten (action=edit)');
pruefe(S.siBearbeitenUrl({ webUrl: 'https://firma-my.sharepoint.com/personal/x/_layouts/15/Doc.aspx?sourcedoc=%7BAB12%7D' }).endsWith('&action=edit'), 'fehlendes action=edit wird ergänzt');
pruefe(S.siBearbeitenUrl({ webUrl: 'https://firma-my.sharepoint.com/personal/x/Documents/Bedarfsanfrage/a.xlsx' }).endsWith('/a.xlsx?web=1'), 'direkte Dateiadresse öffnet im Browser (web=1)');
pruefe(S.siBearbeitenUrl({ webUrl: 'javascript:alert(1)' }) === '' && S.siBearbeitenUrl({}) === '', 'nur https-Adressen');
pruefe(S.siDriveId("01ABC'); alert(1);//") === '01ABCalert1' && S.siDriveId('D4648F06C91D9D3D!54927') === 'D4648F06C91D9D3D!54927', 'Element-ID ohne fremde Zeichen');
pruefe(S.siStatus(new Error('Graph GET 404: x')) === 404 && S.siStatus(new Error('Graph POST 409: y')) === 409 && S.siStatus(new Error('x')) === 0, 'Statuscode aus der Fehlermeldung');
const jetzt = new Date(2026, 10, 12, 14, 0).getTime();
pruefe(S.siZeitText(new Date(jetzt - 20e3).toISOString(), jetzt) === 'gerade gespeichert'
  && S.siZeitText(new Date(jetzt - 5 * 60e3).toISOString(), jetzt) === 'gespeichert vor 5 Min.', 'Zeitangabe: gerade, vor Minuten');
pruefe(S.siZeitText(new Date(2026, 10, 12, 9, 7).toISOString(), jetzt) === 'gespeichert heute 09:07'
  && S.siZeitText(new Date(2026, 10, 10, 9, 7).toISOString(), jetzt) === 'gespeichert 10.11.2026 09:07', 'Zeitangabe: heute, älter');
const quelle = lies('sammelimport.js');
const knopf = quelle.slice(quelle.indexOf('async function siVorlageExcel()'), quelle.indexOf('async function siVorlagenLaden'));
pruefe(knopf.length > 100 && knopf.includes('siVorlageHochladen()') && !knopf.includes('siHerunterladen('), '„Excel-Vorlage öffnen“ legt im OneDrive an und lädt nichts herunter');
pruefe(quelle.includes("'@microsoft.graph.downloadUrl'") && !/[/]content['`]/.test(quelle.replace(':/content?', '')), 'Übernehmen liest über die Download-Adresse, nicht über /content');

console.log('Einbindung');
const html = lies('index.html');
pruefe(html.indexOf('src="sammelimport.js"') > html.indexOf('src="app.js"'), 'index.html lädt sammelimport.js nach app.js');
pruefe(html.includes('id="si-karte"') && html.includes('siVorlageExcel()') && html.includes('id="si-datei"'), 'Kasten mit Vorlage und Import im Reiter Sammelanfrage');
pruefe(html.includes('id="si-vorlagen"') && html.includes('id="si-hinweis"') && html.includes('id="si-btn-excel"'), 'Kasten mit Liste der Vorlagen und Hinweisfeld');
pruefe(/renderFavBar[(]'multi'[)];[^}]*siVorlagenLaden[(][)]/.test(app), 'Reiter Sammelanfrage lädt die Vorlagen aus dem OneDrive');
pruefe(/function sammelanfrageSpeichern/.test(app) && /sammelanfrageSpeichern\(multiPositions/.test(app), 'Assistent und Import legen über dieselbe Funktion an');
pruefe(/o\.Hersteller\s*=/.test(app) && /o\.TypNr\s*=/.test(app) && /o\.Begruendung\s*=/.test(app), 'Positionen speichern Hersteller, Typ-Nr und Begründung');

console.log(fehler ? `\n${fehler} Prüfung(en) fehlgeschlagen.` : '\nAlle Prüfungen bestanden.');
process.exit(fehler ? 1 : 0);
