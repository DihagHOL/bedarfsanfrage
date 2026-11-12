'use strict';
// ── SAMMELANFRAGE AUS EXCEL ODER CSV ─────────────────────────────────────────
// Vor allem die Instandhaltung erfasst viele Positionen lieber in Excel (Vorschlag
// „Elektr. Bedarfsmeldung“). Hier: Vorlage erzeugen (Excel mit Auswahllisten oder CSV),
// die ausgefüllte Datei lesen und prüfen, die Zeilen nach Warengruppe, Abteilung,
// Beschaffungsart und Lieferant zu Sammelanfragen bündeln, eine Vorschau zeigen und anlegen.
// Die Vorlage entsteht im Browser: Abteilungen und Kostenstellen kommen aus SharePoint
// und gehören nicht in das öffentliche Repository.
// Nutzt aus app.js: $id, esc, toast, fmtEuro, anzahl, ensureXLSX, abteilungen, KOSTENST_DATA,
// WG_FALLBACK, colByKey, resolvedFields, genehmigungsweg, positionFuerSp,
// sammelanfrageSpeichern, pendingPrefill, navigate, closeModal, loadItems.

const SI_TAB = String.fromCharCode(9);
const SI_LF  = String.fromCharCode(10);
const SI_CR  = String.fromCharCode(13);
const SI_BOM = String.fromCharCode(65279);
const SI_GS  = String.fromCharCode(8211);   // Gedankenstrich wie in „4334 – Reparatur“
const SI_NBSP = String.fromCharCode(160);
const SI_MAX_ZEILEN = 1000;                 // so weit reichen Auswahllisten und Formate der Vorlage
const SI_MAX_JSON = 60000;                  // Spalte Positionen: mehrzeiliger Text, Reserve zum Höchstwert

// Spalten der Vorlage. alias: normierte Kopfzeilen, die der Import erkennt (* = beginnt mit).
// Die Titel folgen dem Vorschlag der Instandhaltung, der Import liest auch dessen Schreibweise.
const SI_SPALTEN = [
  { key: 'menge',        titel: 'Menge',                                   pflicht: true,  breite: 9,  hinweis: 'Zahl, zum Beispiel 10 oder 2,5.', alias: ['menge', 'anzahl'] },
  { key: 'einheit',      titel: 'Einheit (VP)',                            pflicht: true,  breite: 13, hinweis: 'Aus der Liste. VPE heißt Verpackungseinheit.', alias: ['einheitvp', 'einheit', 'vp', 'vpe', 'me', 'mengeneinheit'] },
  { key: 'hersteller',   titel: 'Hersteller',                              pflicht: false, breite: 20, hinweis: '', alias: ['hersteller'] },
  { key: 'artnr',        titel: 'Art-Nr / Best-Nr',                        pflicht: false, breite: 22, hinweis: 'Nummer beim Hersteller.', alias: ['artnrbestnr', 'artnr', 'bestnr', 'artikelnummer', 'bestellnummer', 'externeartikelnummer'] },
  { key: 'typnr',        titel: 'Typ-Nr',                                  pflicht: false, breite: 28, hinweis: 'Typ beim Hersteller.', alias: ['typnr', 'typ', 'typnummer'] },
  { key: 'bezeichnung',  titel: 'Bezeichnung / Beschreibung des Artikels', pflicht: true,  breite: 44, hinweis: 'Was gebraucht wird.', alias: ['bezeichnung*', 'beschreibung*', 'artikel'] },
  { key: 'begruendung',  titel: 'Begründung des Bedarfs',                  pflicht: false, breite: 30, hinweis: 'Wofür, zum Beispiel Anlage oder Gebäude.', alias: ['begrundung*', 'begruendung*', 'grund'] },
  { key: 'kostenstelle', titel: 'Kostenstelle',                            pflicht: false, breite: 24, hinweis: 'Aus der Liste oder die Nummer eintragen.', alias: ['kostenstelle', 'kst'] },
  { key: 'warengruppe',  titel: 'Warengruppe',                             pflicht: true,  breite: 30, hinweis: 'Aus der Liste.', alias: ['warengruppe'] },
  { key: 'abteilung',    titel: 'Abteilung',                               pflicht: false, breite: 22, hinweis: 'Aus der Liste.', alias: ['abteilung'] },
  { key: 'art1',         titel: 'Beschaffungsart 1',                       pflicht: true,  breite: 32, hinweis: 'Bestandsmaterial oder nicht bestandsgeführt.', alias: ['beschaffungsart1', 'beschaffungsart'] },
  { key: 'art2',         titel: 'Beschaffungsart 2',                       pflicht: true,  breite: 34, hinweis: 'Direktes Material oder indirektes Material / Dienstleistung.', alias: ['beschaffungsart2'] },
  { key: 'lieferant',    titel: 'Lieferant',                               pflicht: true,  breite: 20, hinweis: '', alias: ['lieferant', 'lieferrant'] },
  { key: 'preis',        titel: 'Einzelpreis (€ netto)',                   pflicht: true,  breite: 16, hinweis: 'Preis je Einheit.', alias: ['einzelpreis*', 'preis'] },
  { key: 'termin',       titel: 'Benötigt bis',                            pflicht: false, breite: 14, hinweis: 'Datum, zum Beispiel 31.12.2026.', alias: ['benotigtbis', 'benoetigtbis', 'termin', 'liefertermin'] },
];
// Werte wie im Assistenten der Sammelanfrage (Auswahl und „Zusätzlich kombinierbar“)
const SI_ART1 = ['Bestandsmaterial (bestandsgeführt)', 'Nicht-bestandsgeführtes Material'];
const SI_ART2 = ['Direktes Material', 'Indirektes Material / Dienstleistung'];
const SI_EINHEITEN = ['Stück', 'VPE', 'Paar', 'kg', 'm', 'Liter', 'Anzahl', 'Lagereinheiten'];
const SI_EINHEIT_KURZ = {
  stk: 'Stück', stck: 'Stück', st: 'Stück', stuck: 'Stück', stueck: 'Stück',
  vp: 'VPE', vpe: 'VPE', verpackungseinheit: 'VPE',
  l: 'Liter', ltr: 'Liter', liter: 'Liter', kg: 'kg', m: 'm', meter: 'm', paar: 'Paar', pr: 'Paar',
  anzahl: 'Anzahl', le: 'Lagereinheiten', lagereinheit: 'Lagereinheiten', lagereinheiten: 'Lagereinheiten',
};

// ── Werte lesen ──────────────────────────────────────────────────────────────
function siNorm(v) {
  return String(v ?? '').toLowerCase()
    .replace(/ä/g, 'a').replace(/ö/g, 'o').replace(/ü/g, 'u').replace(/ß/g, 'ss')
    .replace(/[^a-z0-9]/g, '');
}
function siIso(d) {
  const z = n => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${z(d.getUTCMonth() + 1)}-${z(d.getUTCDate())}`;
}
function siText(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return siIso(v);
  return String(v).trim();
}
// Zahl aus Excel (Zahl) oder Text („84,00 €“, „0,0267“, „1.234,50“)
function siZahl(v) {
  if (typeof v === 'number') return v;
  let s = siText(v).split(SI_NBSP).join('').replace(/€|eur/gi, '').replace(/ /g, '');
  if (!s) return NaN;
  if (s.includes(',')) s = s.replace(/[.]/g, '').replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}
function siDatumPruefen(j, mo, t, roh) {
  const d = new Date(Date.UTC(j, mo - 1, t));
  if (d.getUTCFullYear() !== j || d.getUTCMonth() !== mo - 1 || d.getUTCDate() !== t) return { fehler: `Datum „${roh}“ gibt es nicht` };
  return { wert: siIso(d) };
}
// Datum aus Excel (Seriennummer) oder Text (TT.MM.JJJJ, JJJJ-MM-TT); leer ist erlaubt
function siDatum(v) {
  if (v === null || v === undefined || v === '') return { wert: '' };
  if (v instanceof Date) return { wert: siIso(v) };
  if (typeof v === 'number') {
    if (v < 20000 || v > 80000) return { fehler: `Datum „${v}“ ist ungültig` };
    return { wert: siIso(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5)) };
  }
  const s = String(v).trim();
  if (!s) return { wert: '' };
  let m = s.match(/^([0-9]{1,2})[.]([0-9]{1,2})[.]([0-9]{2,4})$/);
  if (m) return siDatumPruefen(m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]), Number(m[2]), Number(m[1]), s);
  m = s.match(/^([0-9]{4})-([0-9]{2})-([0-9]{2})/);
  if (m) return siDatumPruefen(Number(m[1]), Number(m[2]), Number(m[3]), s);
  return { fehler: `Datum „${s}“ nicht lesbar (TT.MM.JJJJ)` };
}
function siEinheit(v) {
  const s = siText(v);
  if (!s) return { fehler: 'Einheit fehlt' };
  return { wert: SI_EINHEIT_KURZ[siNorm(s)] || s };
}
function siWarengruppenListe() {
  const sp = (typeof resolvedFields !== 'undefined' && resolvedFields && resolvedFields.Warengruppe) || 'Warengruppe';
  const ch = (typeof colByKey !== 'undefined' && colByKey && colByKey[sp]?.choice?.choices) || [];
  return ch.length ? ch.slice() : WG_FALLBACK.slice();
}
function siWgCode(o) { return (String(o).match(/^[0-9]+/) || [''])[0]; }
function siWgName(o) { return String(o).replace(/^[0-9]+/, '').replace(/^[^a-zA-Z0-9äöüÄÖÜß]+/, '').trim(); }
// „Reparatur“, „4334“, „4334 Reparatur“ oder der volle Wert der Liste
function siWarengruppe(v) {
  const s = siText(v);
  if (!s) return { fehler: 'Warengruppe fehlt' };
  const liste = siWarengruppenListe();
  const exakt = liste.find(o => o.trim().toLowerCase() === s.toLowerCase());
  if (exakt) return { wert: exakt };
  const code = siWgCode(s), name = siNorm(siWgName(s));
  const kand = liste.filter(o => (!code || siWgCode(o) === code) && (!name || siNorm(siWgName(o)) === name));
  if (kand.length === 1) return { wert: kand[0] };
  if (kand.length > 1) return { fehler: `Warengruppe ${s} ist nicht eindeutig (${kand.map(siWgName).join(' oder ')})` };
  return { fehler: `Warengruppe „${s}“ gibt es nicht` };
}
// Abteilung aus der zentralen Liste; erkennt auch Kurzformen wie „Instandh.“
function siAbteilung(v) {
  const s = siText(v);
  if (!s) return { wert: '' };
  const liste = (typeof abteilungen !== 'undefined' && Array.isArray(abteilungen)) ? abteilungen : [];
  if (!liste.length) return { wert: s };
  const exakt = liste.find(a => a.toLowerCase() === s.toLowerCase());
  if (exakt) return { wert: exakt };
  const gleich = liste.filter(a => siNorm(a) === siNorm(s));
  if (gleich.length === 1) return { wert: gleich[0] };
  const woerter = x => String(x).toLowerCase().replace(/ä/g, 'a').replace(/ö/g, 'o').replace(/ü/g, 'u')
    .replace(/ß/g, 'ss').split(/[^a-z0-9]+/).filter(Boolean);
  const w = woerter(s);
  const passt = liste.filter(a => {
    const u = woerter(a);
    return u.length === w.length && u.every((x, i) => x.startsWith(w[i]) || w[i].startsWith(x));
  });
  if (passt.length === 1) return { wert: passt[0] };
  return { fehler: `Abteilung „${s}“ gibt es nicht` };
}
function siArt1(v) {
  const n = siNorm(v);
  if (!n) return { fehler: 'Beschaffungsart 1 fehlt' };
  if (n.includes('nicht') || n.startsWith('neu')) return { wert: SI_ART1[1] };
  if (n.includes('bestand')) return { wert: SI_ART1[0] };
  return { fehler: `Beschaffungsart 1 „${siText(v)}“ unbekannt` };
}
function siArt2(v) {
  const n = siNorm(v);
  if (!n) return { fehler: 'Beschaffungsart 2 fehlt' };
  if (n.includes('indirekt') || n.includes('dienst')) return { wert: SI_ART2[1] };
  if (n.includes('direkt')) return { wert: SI_ART2[0] };
  return { fehler: `Beschaffungsart 2 „${siText(v)}“ unbekannt` };
}

// ── Datei lesen ──────────────────────────────────────────────────────────────
// CSV mit Semikolon, Komma oder Tab; Felder in Anführungszeichen dürfen Trenner enthalten
function siCsvZeilen(text) {
  if (text.charCodeAt(0) === 65279) text = text.slice(1);
  const erste = text.split(SI_LF)[0] || '';
  const trenner = [';', SI_TAB, ','].map(t => [t, erste.split(t).length]).sort((a, b) => b[1] - a[1])[0][0];
  const zeilen = [];
  let zeile = [], feld = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') {
        if (text[i + 1] === '"') { feld += '"'; i++; } else inQ = false;
      } else feld += c;
    } else if (c === '"') inQ = true;
    else if (c === trenner) { zeile.push(feld); feld = ''; }
    else if (c === SI_LF || c === SI_CR) {
      if (c === SI_CR && text[i + 1] === SI_LF) i++;
      zeile.push(feld); zeilen.push(zeile); zeile = []; feld = '';
    } else feld += c;
  }
  if (feld !== '' || zeile.length) { zeile.push(feld); zeilen.push(zeile); }
  return zeilen;
}
// Kopfzeile in den ersten Zeilen suchen. Liefert { index, spalten: { key: Spaltennummer } }.
function siKopfFinden(zeilen) {
  for (let i = 0; i < Math.min(zeilen.length, 15); i++) {
    const spalten = {};
    (zeilen[i] || []).forEach((zelle, j) => {
      const n = siNorm(zelle);
      if (!n) return;
      const s = SI_SPALTEN.find(sp => spalten[sp.key] === undefined
        && sp.alias.some(a => a.endsWith('*') ? n.startsWith(a.slice(0, -1)) : n === a));
      if (s) spalten[s.key] = j;
    });
    if (['menge', 'bezeichnung', 'warengruppe'].every(k => spalten[k] !== undefined)) return { index: i, spalten };
  }
  return null;
}
async function siDateiLesen(file) {
  const endung = (file.name.split('.').pop() || '').toLowerCase();
  if (endung === 'csv' || endung === 'txt') {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let text;
    // Excel speichert „CSV (Trennzeichen-getrennt)“ in Windows-1252, „CSV UTF-8“ in UTF-8
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch { text = new TextDecoder('windows-1252').decode(bytes); }
    return siCsvZeilen(text);
  }
  if (['xlsx', 'xlsm', 'xls'].includes(endung)) {
    await ensureXLSX();
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const blatt = name => XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
    // Bevorzugt das Blatt der Vorlage, sonst das erste mit erkennbarer Kopfzeile
    const namen = [...wb.SheetNames].sort((a, b) => (b === 'Bedarfsmeldung') - (a === 'Bedarfsmeldung'));
    for (const name of namen) {
      const zeilen = blatt(name);
      if (siKopfFinden(zeilen)) return zeilen;
    }
    return blatt(wb.SheetNames[0]);
  }
  throw new Error('Bitte eine Excel-Datei (.xlsx, .xls) oder eine CSV-Datei wählen.');
}

// ── Prüfen und bündeln ───────────────────────────────────────────────────────
// Liefert { gruppen, fehler, positionen, dateiname }. Eine Gruppe wird eine Sammelanfrage.
function siAuswerten(zeilen, dateiname = '') {
  const ergebnis = { gruppen: [], fehler: [], positionen: 0, dateiname };
  const kopf = siKopfFinden(zeilen);
  if (!kopf) {
    ergebnis.fehler.push('Keine Kopfzeile gefunden. Bitte die Vorlage verwenden (Spalten wie Menge, Bezeichnung, Warengruppe, Lieferant).');
    return ergebnis;
  }
  const fehlt = SI_SPALTEN.filter(s => s.pflicht && kopf.spalten[s.key] === undefined).map(s => s.titel);
  if (fehlt.length) {
    ergebnis.fehler.push(`Diese Spalten fehlen: ${fehlt.join(', ')}.`);
    return ergebnis;
  }
  const gruppen = new Map();
  for (let i = kopf.index + 1; i < zeilen.length; i++) {
    const roh = zeilen[i] || [];
    const zelle = key => kopf.spalten[key] === undefined ? '' : roh[kopf.spalten[key]];
    if (SI_SPALTEN.every(s => siText(zelle(s.key)) === '')) continue;   // leere Zeile
    const f = [];
    const menge = siZahl(zelle('menge'));
    if (!(menge > 0)) f.push('Menge fehlt oder ist keine Zahl größer 0');
    const bezeichnung = siText(zelle('bezeichnung'));
    if (!bezeichnung) f.push('Bezeichnung fehlt');
    const einheit = siEinheit(zelle('einheit'));
    const wg = siWarengruppe(zelle('warengruppe'));
    const abt = siAbteilung(zelle('abteilung'));
    const art1 = siArt1(zelle('art1'));
    const art2 = siArt2(zelle('art2'));
    const termin = siDatum(zelle('termin'));
    [einheit, wg, abt, art1, art2, termin].forEach(x => { if (x.fehler) f.push(x.fehler); });
    const lieferant = siText(zelle('lieferant'));
    if (!lieferant) f.push('Lieferant fehlt');
    const preisText = siText(zelle('preis'));
    const preis = siZahl(zelle('preis'));
    if (!preisText) f.push('Einzelpreis fehlt');
    else if (!(preis >= 0)) f.push(`Einzelpreis „${preisText}“ ist keine Zahl`);
    if (f.length) { ergebnis.fehler.push(`Zeile ${i + 1}: ${f.join(', ')}.`); continue; }

    const position = {
      artNr: '', extArtNr: siText(zelle('artnr')), bezeichnung, menge: String(menge), me: einheit.wert,
      preis: String(preis), termin: termin.wert, kostenstelle: siText(zelle('kostenstelle')),
      hersteller: siText(zelle('hersteller')), typNr: siText(zelle('typnr')), begruendung: siText(zelle('begruendung')),
    };
    const beschaffungslogik = `${art1.wert}, ${art2.wert}`;
    const schluessel = [wg.wert, abt.wert, beschaffungslogik, lieferant.toLowerCase()].join('|');
    if (!gruppen.has(schluessel)) gruppen.set(schluessel, { warengruppe: wg.wert, abteilung: abt.wert, beschaffungslogik, lieferant, positionen: [] });
    gruppen.get(schluessel).positionen.push(position);
    ergebnis.positionen++;
  }
  ergebnis.gruppen = [...gruppen.values()].map(g => ({
    ...g,
    summe: Math.round(g.positionen.reduce((s, p) => s + Number(p.preis) * Number(p.menge), 0) * 100) / 100,
    beschreibung: [...new Set(g.positionen.map(p => p.begruendung).filter(Boolean))].join('; ').slice(0, 500),
  }));
  ergebnis.gruppen.forEach((g, i) => {
    if (JSON.stringify(g.positionen).length > SI_MAX_JSON)
      ergebnis.fehler.push(`Sammelanfrage ${i + 1} hat zu viele Positionen (${g.positionen.length}). Bitte die Datei teilen.`);
  });
  if (!ergebnis.positionen && !ergebnis.fehler.length) ergebnis.fehler.push('Die Datei enthält keine Positionen.');
  return ergebnis;
}

// ── Vorlage als Excel-Datei (ohne Bibliothek: SpreadsheetML in einem ZIP) ──────
const SI_XML_KOPF = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SI_NS_MAIN  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SI_NS_R     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SI_NS_PKG   = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SI_NS_TYPEN = 'http://schemas.openxmlformats.org/package/2006/content-types';

function siXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function siBuchstabe(i) {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
function siZelle(ref, wert, stil) {
  const s = stil ? ` s="${stil}"` : '';
  if (wert === null || wert === undefined || wert === '') return stil ? `<c r="${ref}"${s}/>` : '';
  if (typeof wert === 'number') return `<c r="${ref}"${s}><v>${wert}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${siXml(wert)}</t></is></c>`;
}
// zeilen: [{ werte: [...], stil, hoehe }], spalten: [{ breite, stil }]
function siBlattXml({ spalten = [], zeilen = [], fixiert = false, aktiv = false, pruefungen = [] }) {
  const ansicht = `<sheetViews><sheetView workbookViewId="0"${aktiv ? ' tabSelected="1"' : ''}>`
    + (fixiert ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' : '')
    + '</sheetView></sheetViews>';
  const cols = spalten.length
    ? `<cols>${spalten.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.breite}" customWidth="1"${c.stil ? ` style="${c.stil}"` : ''}/>`).join('')}</cols>`
    : '';
  const daten = zeilen.map((z, r) => {
    const zellen = z.werte.map((w, c) => siZelle(`${siBuchstabe(c)}${r + 1}`, w, z.stil || 0)).join('');
    return `<row r="${r + 1}"${z.hoehe ? ` ht="${z.hoehe}" customHeight="1"` : ''}>${zellen}</row>`;
  }).join('');
  const dv = pruefungen.length ? `<dataValidations count="${pruefungen.length}">${pruefungen.join('')}</dataValidations>` : '';
  return `${SI_XML_KOPF}<worksheet xmlns="${SI_NS_MAIN}" xmlns:r="${SI_NS_R}">${ansicht}<sheetFormatPr defaultRowHeight="15"/>${cols}`
    + `<sheetData>${daten}</sheetData>${dv}<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}
function siPruefung({ typ = 'list', operator = '', bereich, formel, art = 'stop', titel, text }) {
  return `<dataValidation type="${typ}"${operator ? ` operator="${operator}"` : ''} errorStyle="${art}" allowBlank="1"`
    + ` showInputMessage="1" showErrorMessage="1" errorTitle="${siXml(titel)}" error="${siXml(text)}" sqref="${bereich}">`
    + `<formula1>${siXml(formel)}</formula1></dataValidation>`;
}
// Stile: 1 Kopfzeile, 2 Text, 3 Euro, 4 Datum, 5 Überschrift, 6 Umbruch, 7 Tabellenkopf hell, 8 Tabellenzelle
function siStileXml() {
  const rahmen = '<left style="thin"><color rgb="FFB4C6DC"/></left><right style="thin"><color rgb="FFB4C6DC"/></right>'
    + '<top style="thin"><color rgb="FFB4C6DC"/></top><bottom style="thin"><color rgb="FFB4C6DC"/></bottom><diagonal/>';
  return `${SI_XML_KOPF}<styleSheet xmlns="${SI_NS_MAIN}">`
    + '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00 &quot;€&quot;"/><numFmt numFmtId="165" formatCode="dd.mm.yyyy"/></numFmts>'
    + '<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
    + '<font><b/><sz val="14"/><color rgb="FF1A2644"/><name val="Calibri"/><family val="2"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FF1A2644"/><name val="Calibri"/><family val="2"/></font></fonts>'
    + '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FF17509E"/><bgColor indexed="64"/></patternFill></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF6"/><bgColor indexed="64"/></patternFill></fill></fills>'
    + `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${rahmen}</border></borders>`
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="9">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
    + '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
    + '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
    + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
}
// Auswahllisten zum Zeitpunkt des Klicks: Warengruppen, zentrale Abteilungen, Kostenstellen
function siListen() {
  const abt = (typeof abteilungen !== 'undefined' && Array.isArray(abteilungen)) ? abteilungen.slice() : [];
  const kst = ((typeof KOSTENST_DATA !== 'undefined' && Array.isArray(KOSTENST_DATA)) ? KOSTENST_DATA : [])
    .slice()
    .sort((a, b) => (parseInt(a.nr, 10) || 0) - (parseInt(b.nr, 10) || 0) || String(a.nr).localeCompare(String(b.nr)))
    .map(e => e.label ? `${e.nr} ${SI_GS} ${e.label}` : String(e.nr));
  return [
    { titel: 'Warengruppe', name: 'Liste_Warengruppe', werte: siWarengruppenListe() },
    { titel: 'Abteilung', name: 'Liste_Abteilung', werte: abt },
    { titel: 'Beschaffungsart 1', name: 'Liste_Art1', werte: SI_ART1.slice() },
    { titel: 'Beschaffungsart 2', name: 'Liste_Art2', werte: SI_ART2.slice() },
    { titel: 'Einheit', name: 'Liste_Einheit', werte: SI_EINHEITEN.slice() },
    { titel: 'Kostenstelle', name: 'Liste_Kostenstelle', werte: kst },
  ];
}
function siVorlageXlsx(heute = new Date()) {
  const listen = siListen();
  const namen = listen.map((l, i) => l.werte.length
    ? [l.name, `Listen!$${siBuchstabe(i)}$2:$${siBuchstabe(i)}$${l.werte.length + 1}`] : null).filter(Boolean);
  const hat = name => namen.some(n => n[0] === name);
  const tiefe = Math.max(1, ...listen.map(l => l.werte.length));
  const listenZeilen = [{ werte: listen.map(l => l.titel), stil: 7 }];
  for (let r = 0; r < tiefe; r++) listenZeilen.push({ werte: listen.map(l => l.werte[r] ?? '') });

  const spalte = key => siBuchstabe(SI_SPALTEN.findIndex(s => s.key === key));
  const bereich = key => `${spalte(key)}2:${spalte(key)}${SI_MAX_ZEILEN}`;
  const pruefungen = [
    siPruefung({ typ: 'decimal', operator: 'greaterThan', bereich: bereich('menge'), formel: '0', titel: 'Menge', text: 'Bitte eine Zahl größer 0 eintragen.' }),
    hat('Liste_Einheit') && siPruefung({ bereich: bereich('einheit'), formel: 'Liste_Einheit', art: 'warning', titel: 'Einheit', text: 'Diese Einheit steht nicht in der Liste. Trotzdem übernehmen?' }),
    hat('Liste_Kostenstelle') && siPruefung({ bereich: bereich('kostenstelle'), formel: 'Liste_Kostenstelle', art: 'warning', titel: 'Kostenstelle', text: 'Diese Kostenstelle steht nicht in der Liste. Trotzdem übernehmen?' }),
    siPruefung({ bereich: bereich('warengruppe'), formel: 'Liste_Warengruppe', titel: 'Warengruppe', text: 'Bitte eine Warengruppe aus der Liste wählen.' }),
    hat('Liste_Abteilung') && siPruefung({ bereich: bereich('abteilung'), formel: 'Liste_Abteilung', titel: 'Abteilung', text: 'Bitte eine Abteilung aus der Liste wählen.' }),
    siPruefung({ bereich: bereich('art1'), formel: 'Liste_Art1', titel: 'Beschaffungsart 1', text: 'Bitte aus der Liste wählen.' }),
    siPruefung({ bereich: bereich('art2'), formel: 'Liste_Art2', titel: 'Beschaffungsart 2', text: 'Bitte aus der Liste wählen.' }),
    siPruefung({ typ: 'decimal', operator: 'greaterThanOrEqual', bereich: bereich('preis'), formel: '0', titel: 'Einzelpreis', text: 'Bitte den Preis je Einheit als Zahl eintragen.' }),
    siPruefung({ typ: 'date', operator: 'greaterThan', bereich: bereich('termin'), formel: '36526', titel: 'Benötigt bis', text: 'Bitte ein Datum eintragen, zum Beispiel 31.12.2026.' }),
  ].filter(Boolean);
  const spaltenStil = { artnr: 2, typnr: 2, kostenstelle: 2, preis: 3, termin: 4 };
  const blattDaten = siBlattXml({
    spalten: SI_SPALTEN.map(s => ({ breite: s.breite, stil: spaltenStil[s.key] || 0 })),
    zeilen: [{ werte: SI_SPALTEN.map(s => s.titel), stil: 1, hoehe: 32 }],
    fixiert: true, aktiv: true, pruefungen,
  });

  const z = n => String(n).padStart(2, '0');
  const datum = `${z(heute.getDate())}.${z(heute.getMonth() + 1)}.${heute.getFullYear()}`;
  const blattHinweise = siBlattXml({
    spalten: [{ breite: 40 }, { breite: 9 }, { breite: 70 }],
    zeilen: [
      { werte: ['Bedarfsmeldung per Excel'], stil: 5, hoehe: 22 },
      { werte: [''] },
      { werte: ['1. Im Blatt „Bedarfsmeldung“ jeden Artikel in eine eigene Zeile eintragen. Die Kopfzeile bitte nicht ändern.'] },
      { werte: ['2. Pflicht sind Menge, Einheit, Bezeichnung, Warengruppe, Beschaffungsart 1 und 2, Lieferant und Einzelpreis.'] },
      { werte: ['3. Warengruppe, Abteilung, Beschaffungsart, Einheit und Kostenstelle über die Auswahlliste in der Zelle wählen.'] },
      { werte: ['4. Zeilen mit gleicher Warengruppe, Abteilung, Beschaffungsart und gleichem Lieferanten werden eine Sammelanfrage.'] },
      { werte: ['5. Datei als Excel oder CSV speichern und in der App unter Sammelanfrage mit „Datei importieren“ hochladen.'] },
      { werte: ['6. Vor dem Anlegen zeigt die App eine Vorschau mit Positionen, Summen und Genehmigungsweg.'] },
      { werte: ['7. Angebote danach in der Anfrage unter „Anhänge“ anhängen.'] },
      { werte: [''] },
      { werte: [`Auswahllisten nach dem Stand vom ${datum}. Für neue Abteilungen oder Kostenstellen die Vorlage in der App neu öffnen.`] },
      { werte: [''] },
      { werte: ['Spalte', 'Pflicht', 'Hinweis'], stil: 7 },
      ...SI_SPALTEN.map(s => ({ werte: [s.titel, s.pflicht ? 'ja' : 'nein', s.hinweis], stil: 8 })),
    ],
  });
  const blattListen = siBlattXml({ spalten: listen.map(() => ({ breite: 36 })), zeilen: listenZeilen });

  const blaetter = ['Bedarfsmeldung', 'Hinweise', 'Listen'];
  const mappe = `${SI_XML_KOPF}<workbook xmlns="${SI_NS_MAIN}" xmlns:r="${SI_NS_R}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>`
    + blaetter.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}"${n === 'Listen' ? ' state="hidden"' : ''} r:id="rId${i + 1}"/>`).join('')
    + `</sheets>${namen.length ? `<definedNames>${namen.map(([n, ref]) => `<definedName name="${n}">${siXml(ref)}</definedName>`).join('')}</definedNames>` : ''}</workbook>`;
  const mappeRels = `${SI_XML_KOPF}<Relationships xmlns="${SI_NS_PKG}">`
    + blaetter.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${SI_NS_R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
    + `<Relationship Id="rId${blaetter.length + 1}" Type="${SI_NS_R}/styles" Target="styles.xml"/></Relationships>`;
  const typen = `${SI_XML_KOPF}<Types xmlns="${SI_NS_TYPEN}">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + blaetter.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
  const wurzelRels = `${SI_XML_KOPF}<Relationships xmlns="${SI_NS_PKG}"><Relationship Id="rId1" Type="${SI_NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  return siZip([
    { name: '[Content_Types].xml', daten: typen },
    { name: '_rels/.rels', daten: wurzelRels },
    { name: 'xl/workbook.xml', daten: mappe },
    { name: 'xl/_rels/workbook.xml.rels', daten: mappeRels },
    { name: 'xl/styles.xml', daten: siStileXml() },
    { name: 'xl/worksheets/sheet1.xml', daten: blattDaten },
    { name: 'xl/worksheets/sheet2.xml', daten: blattHinweise },
    { name: 'xl/worksheets/sheet3.xml', daten: blattListen },
  ]);
}
function siVorlageCsvText() {
  return SI_BOM + SI_SPALTEN.map(s => s.titel).join(';') + SI_CR + SI_LF;
}

// ZIP ohne Kompression (Methode 0). Reicht für die kleine Vorlage, Excel liest das.
const SI_CRC_TABELLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();
function siCrc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = SI_CRC_TABELLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function siZip(dateien) {
  const enc = new TextEncoder();
  const teile = [], verzeichnis = [];
  let versatz = 0;
  for (const d of dateien) {
    const name = enc.encode(d.name);
    const daten = typeof d.daten === 'string' ? enc.encode(d.daten) : d.daten;
    const crc = siCrc32(daten);
    const lokal = new DataView(new ArrayBuffer(30));
    lokal.setUint32(0, 0x04034b50, true); lokal.setUint16(4, 20, true); lokal.setUint16(6, 0x0800, true);
    lokal.setUint16(8, 0, true); lokal.setUint16(10, 0, true); lokal.setUint16(12, 0x21, true);
    lokal.setUint32(14, crc, true); lokal.setUint32(18, daten.length, true); lokal.setUint32(22, daten.length, true);
    lokal.setUint16(26, name.length, true); lokal.setUint16(28, 0, true);
    teile.push(new Uint8Array(lokal.buffer), name, daten);
    const z = new DataView(new ArrayBuffer(46));
    z.setUint32(0, 0x02014b50, true); z.setUint16(4, 20, true); z.setUint16(6, 20, true); z.setUint16(8, 0x0800, true);
    z.setUint16(10, 0, true); z.setUint16(12, 0, true); z.setUint16(14, 0x21, true); z.setUint32(16, crc, true);
    z.setUint32(20, daten.length, true); z.setUint32(24, daten.length, true); z.setUint16(28, name.length, true);
    z.setUint16(30, 0, true); z.setUint16(32, 0, true); z.setUint16(34, 0, true); z.setUint16(36, 0, true);
    z.setUint32(38, 0, true); z.setUint32(42, versatz, true);
    verzeichnis.push(new Uint8Array(z.buffer), name);
    versatz += 30 + name.length + daten.length;
  }
  const groesse = verzeichnis.reduce((s, x) => s + x.length, 0);
  const ende = new DataView(new ArrayBuffer(22));
  ende.setUint32(0, 0x06054b50, true); ende.setUint16(4, 0, true); ende.setUint16(6, 0, true);
  ende.setUint16(8, dateien.length, true); ende.setUint16(10, dateien.length, true);
  ende.setUint32(12, groesse, true); ende.setUint32(16, versatz, true); ende.setUint16(20, 0, true);
  const alle = [...teile, ...verzeichnis, new Uint8Array(ende.buffer)];
  const aus = new Uint8Array(alle.reduce((s, x) => s + x.length, 0));
  let p = 0;
  for (const x of alle) { aus.set(x, p); p += x.length; }
  return aus;
}

// ── Oberfläche im Reiter Sammelanfrage ───────────────────────────────────────
let siErgebnis = null;   // letzte ausgewertete Datei, bis zum Anlegen

function siHerunterladen(inhalt, dateiname, typ) {
  const url = URL.createObjectURL(new Blob([inhalt], { type: typ }));
  const a = document.createElement('a');
  a.href = url; a.download = dateiname;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
function siVorlageExcel() {
  try {
    siHerunterladen(siVorlageXlsx(), 'Bedarfsmeldung-Vorlage.xlsx',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    toast('Vorlage erstellt. Öffnen Sie sie aus dem Download in Excel.', 'success');
  } catch (e) { toast('Vorlage konnte nicht erstellt werden: ' + e.message, 'error'); }
}
function siVorlageCsv() {
  siHerunterladen(siVorlageCsvText(), 'Bedarfsmeldung-Vorlage.csv', 'text/csv;charset=utf-8');
}
function siDateiGewaehlt(input) {
  const datei = input.files?.[0];
  input.value = '';
  siImport(datei);
}
async function siImport(datei) {
  if (!datei) return;
  try {
    siErgebnis = siAuswerten(await siDateiLesen(datei), datei.name);
  } catch (e) {
    toast('Datei konnte nicht gelesen werden: ' + e.message, 'error');
    return;
  }
  siVorschau();
}
// Einzelpreise wie 0,0267 € (Aderendhülse je Stück) nicht auf 0,03 € runden
function siEuro(v) {
  return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 4 }).format(v);
}
function siGruppeHtml(g, i) {
  const zeilen = g.positionen.map((p, j) => {
    const zusatz = [p.hersteller, p.extArtNr ? `Art-Nr ${p.extArtNr}` : '', p.typNr ? `Typ ${p.typNr}` : '', p.begruendung, p.termin ? `bis ${fmtDate(p.termin)}` : '']
      .filter(Boolean).join(' · ');
    return `<tr><td>${j + 1}</td><td class="r">${esc(p.menge)}</td><td>${esc(p.me)}</td>
      <td>${esc(p.bezeichnung)}${zusatz ? `<div class="pos-extra">${esc(zusatz)}</div>` : ''}</td>
      <td>${esc(p.kostenstelle) || SI_GS}</td><td class="r">${siEuro(Number(p.preis))}</td>
      <td class="r">${fmtEuro(Number(p.preis) * Number(p.menge))}</td></tr>`;
  }).join('');
  const weg = genehmigungsweg(g.summe, false, g.beschaffungslogik, g.warengruppe);
  return `<div class="si-gruppe">
    <div class="si-kopf"><strong>Sammelanfrage ${i + 1}</strong>
      <span>${esc(g.warengruppe)}</span><span>${esc(g.abteilung || 'ohne Abteilung')}</span>
      <span>${esc(g.beschaffungslogik)}</span><span>Lieferant ${esc(g.lieferant)}</span></div>
    <div class="si-tab-wrap"><table class="si-tab"><thead><tr><th>#</th><th class="r">Menge</th><th>Einheit</th>
      <th>Bezeichnung</th><th>Kostenstelle</th><th class="r">Einzelpreis</th><th class="r">Summe</th></tr></thead>
      <tbody>${zeilen}</tbody></table></div>
    <div class="si-fuss"><span class="si-weg">${esc(weg)}</span>
      <button type="button" class="btn btn-outline btn-sm" onclick="siImAssistenten(${i})"
        title="Nur diese Sammelanfrage in den Assistenten übernehmen, etwa um Priorität oder Angebote zu ergänzen">✏️ Im Assistenten bearbeiten</button></div>
  </div>`;
}
function siVorschau() {
  const e = siErgebnis;
  if (!e) return;
  const n = e.gruppen.length;
  $id('modal-title').textContent = `Import prüfen: ${e.dateiname}`;
  const zusammenfassung = e.positionen
    ? `${anzahl(e.positionen, 'Position', 'Positionen')} ergeben ${anzahl(n, 'Sammelanfrage', 'Sammelanfragen')}.` : '';
  const angelegt = e.angelegt ? `<div class="info-box info">Bereits angelegt: ${anzahl(e.angelegt, 'Sammelanfrage', 'Sammelanfragen')}. Unten stehen nur die übrigen.</div>` : '';
  const fehler = e.fehler.length ? `<div class="info-box warn si-fehler"><strong>Die Datei enthält Fehler.</strong>
      Bitte in der Datei korrigieren und erneut importieren.
      <ul>${e.fehler.slice(0, 25).map(f => `<li>${esc(f)}</li>`).join('')}</ul>
      ${e.fehler.length > 25 ? `<div>Dazu ${e.fehler.length - 25} weitere.</div>` : ''}</div>` : '';
  $id('modal-body').innerHTML = `<div class="si-vorschau">${zusammenfassung ? `<p class="si-zus">${esc(zusammenfassung)}</p>` : ''}
    ${angelegt}${fehler}${e.fehler.length ? '' : e.gruppen.map(siGruppeHtml).join('')}</div>`;
  const ok = n > 0 && !e.fehler.length;
  $id('modal-footer').innerHTML = `<button class="btn btn-ghost" onclick="closeModal()">${ok ? 'Abbrechen' : 'Schließen'}</button>`
    + (ok ? `<button class="btn btn-success" id="si-anlegen" onclick="siAnlegen()">✓ ${n === 1 ? 'Sammelanfrage anlegen' : `${n} Sammelanfragen anlegen`}</button>` : '');
  document.querySelector('#modal-overlay .modal')?.classList.add('modal-wide');
  $id('modal-overlay').classList.remove('hidden');
}
function siSchritte(g) {
  return [
    { Warengruppe: g.warengruppe, Prioritaet: 'Normal', Beschreibung: g.beschreibung, Abteilung: g.abteilung },
    { Beschaffungslogik: g.beschaffungslogik, Lieferant: g.lieferant, Lieferant2: '', Lieferant3: '', Lieferant4: '',
      GeschaetzterPreis: g.summe, LeadBuyerAbschluss: false },
  ];
}
async function siAnlegen() {
  const e = siErgebnis;
  if (!e || e.fehler.length || !e.gruppen.length) return;
  const btn = $id('si-anlegen');
  if (btn) btn.disabled = true;
  const gesamt = e.gruppen.length;
  let fertig = 0, ohneFelder = 0;
  try {
    for (const g of e.gruppen) {
      if (btn) btn.textContent = `Sammelanfrage ${fertig + 1} von ${gesamt} wird angelegt…`;
      const [s2, s3] = siSchritte(g);
      const { felderOk } = await sammelanfrageSpeichern(g.positionen, s2, s3);
      if (!felderOk) ohneFelder++;
      fertig++;
    }
  } catch (err) {
    // Angelegte nicht noch einmal anlegen: nur die übrigen bleiben in der Vorschau
    e.angelegt = (e.angelegt || 0) + fertig;
    e.gruppen = e.gruppen.slice(fertig);
    toast(`${fertig} von ${gesamt} angelegt. Fehler: ${err.message}`, 'error');
    siVorschau();
    return;
  }
  const summe = (e.angelegt || 0) + fertig;
  siErgebnis = null;
  closeModal();
  toast(`${anzahl(summe, 'Sammelanfrage', 'Sammelanfragen')} aus ${e.dateiname} eingereicht. Als Nächstes prüft der Einkauf.`, 'success');
  if (ohneFelder) toast(`Bei ${anzahl(ohneFelder, 'Anfrage', 'Anfragen')} konnten nicht alle Felder gespeichert werden.`, 'error');
  await loadItems(false);
  navigate('mine');
}
// Eine Gruppe in den Assistenten übernehmen, etwa für Priorität, Lead-Buyer oder Angebote
function siImAssistenten(i) {
  const g = siErgebnis?.gruppen[i];
  if (!g) return;
  const [s2, s3] = siSchritte(g);
  pendingPrefill = {
    type: 'multi',
    hinweis: 'Import übernommen. Bitte prüfen und einreichen.',
    data: { ...s2, ...s3 },
    positions: g.positionen.map((p, j) => positionFuerSp(p, j)),
  };
  closeModal();
  navigate('multi');
}

// Datei auch per Drag-and-drop auf den Kasten
if (typeof document !== 'undefined') {
  document.addEventListener('DOMContentLoaded', () => {
    const karte = document.getElementById('si-karte');
    if (!karte) return;
    karte.addEventListener('dragover', e => { e.preventDefault(); karte.classList.add('drag-over'); });
    karte.addEventListener('dragleave', () => karte.classList.remove('drag-over'));
    karte.addEventListener('drop', e => {
      e.preventDefault();
      karte.classList.remove('drag-over');
      siImport(e.dataTransfer?.files?.[0]);
    });
  });
}

// Für die Tests unter Node
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    SI_SPALTEN, SI_ART1, SI_ART2, siNorm, siZahl, siText, siDatum, siEinheit, siWarengruppe, siAbteilung,
    siArt1, siArt2, siCsvZeilen, siKopfFinden, siAuswerten, siVorlageXlsx, siVorlageCsvText, siZip, siCrc32,
  };
}
