'use strict';
/**
 * BPMN-Generator: Bedarfsanfrage (DIHAG SHB)
 * ==========================================
 * Erzeugt die fünf Modelle der Bedarfsanfrage im **Hausschema des RMS**
 * (rms.dihag.de): den Hauptprozess, drei eingebundene Unterprozesse
 * (Anfrage erfassen, Anfrage prüfen und entscheiden, Freigabelauf) und den
 * Auto-Entwurf, der mit einer Uhrzeit beginnt und deshalb ein eigener
 * Prozess ist (R1: genau ein Auslöser).
 *
 * Die Regeln stehen im RMS (`js/prozessschema.js`, R1 bis R10) und gelten für
 * jedes Modell im Haus. `tests/test-prozess.mjs` prüft die erzeugten Dateien.
 *
 * **Führend sind die Modelle im RMS** (sites/ISMS, ISMS Dokumente,
 * Prozesse/SHB). Dieser Generator ist der Weg für den Erstimport und für
 * größere Umbauten. Die Prozess-Seite der App liest die Modelle live aus dem
 * RMS, sobald das Konto Zugriff hat, sonst die Kopie aus docs/bpmn.
 *
 * Der Motor (Raster, Kantenführung, BPMN-DI) ist wörtlich der aus
 * `e-rechnung/scripts/bpmn/generate-bpmn.js` (über `crm`). Zwei Layout-Engines
 * im Haus wären zwei Bildsprachen für dieselbe Sache.
 *
 * Aufruf:  node scripts/bpmn/generate-bpmn.js
 * Ausgabe: docs/bpmn/*.bpmn und docs/bpmn/modelle.json
 * Eingabe: docs/bpmn/rms-ids.json (Datei-Kennungen im RMS)
 *
 * Raster: c = Spalte (von links), r = Zeile innerhalb der Bahn.
 */

const fs = require('fs');
const path = require('path');

/* ── Raster & Maße ─────────────────────────────────────────────────────── */
const COL_W = 170, ROW_H = 120;
const TASK_W = 132, TASK_H = 84, EV = 36, GW = 50;
const POOL_X = 100, POOL_HEAD = 30, LANE_HEAD = 30, POOL_GAP = 70, TOP = 60, BLACKBOX_H = 70;
const CONTENT_X = POOL_X + POOL_HEAD + LANE_HEAD + 20;

/* ── Farben (identisch zur Legende auf der Prozess-Seite der App) ──────── */
const ACTOR = {
  extern: { lane: '#F6F7F9', fill: '#E9ECF0', stroke: '#5B6472' },  // Lieferant
  auto:   { lane: '#F3F8FE', fill: '#D8E8F8', stroke: '#17509E' },  // Power Automate, GitHub Actions
  api:    { lane: '#F8F5FD', fill: '#E6DDF7', stroke: '#5B3FA8' },  // (hier ungenutzt)
  archiv: { lane: '#F1FAF8', fill: '#D2EEE8', stroke: '#0F766E' },  // (hier ungenutzt)
  mensch: { lane: '#FFF8F1', fill: '#FFE3C8', stroke: '#C2410C' },  // Antragsteller, Einkauf, Genehmiger
  app:    { lane: '#F3F6FB', fill: '#DCE5F2', stroke: '#1A2644' },  // Bedarfsanfrage-App
  erp:    { lane: '#F7F7F6', fill: '#E6E6E4', stroke: '#424241' },  // (hier ungenutzt)
};
const TONE = {
  ok:   { fill: '#DDF3E4', stroke: '#1E7B3A' },
  warn: { fill: '#FFF1CC', stroke: '#A65F00' },
  err:  { fill: '#FDE2E1', stroke: '#B42318' },
  gate: { fill: '#FFF8DB', stroke: '#8A6100' },
  plan: { fill: '#F1F1F3', stroke: '#8A8F98' },
};

/* ── Elementarten (Hausschema: Aufgaben sind 👤 user, ⚙ service oder ✋ manual) ── */
const KIND = {
  start:       { tag: 'startEvent',  size: 'ev' },
  msgStart:    { tag: 'startEvent',  size: 'ev', def: 'message' },
  timerStart:  { tag: 'startEvent',  size: 'ev', def: 'timer' },
  end:         { tag: 'endEvent',    size: 'ev' },
  errEnd:      { tag: 'endEvent',    size: 'ev', def: 'error' },
  boundaryErr: { tag: 'boundaryEvent', size: 'ev', def: 'error' },
  service:     { tag: 'serviceTask', size: 'task' },
  user:        { tag: 'userTask',    size: 'task' },
  manual:      { tag: 'manualTask',  size: 'task' },
  call:        { tag: 'callActivity', size: 'task' },
  gw:          { tag: 'exclusiveGateway', size: 'gw' },
  gwInc:       { tag: 'inclusiveGateway', size: 'gw' },
  gwPar:       { tag: 'parallelGateway',  size: 'gw' },
};

const R = b => b.x + b.w, B = b => b.y + b.h, CX = b => b.x + b.w / 2, CY = b => b.y + b.h / 2;
const rnd = v => Math.round(v);
const isGw = n => KIND[n.k].size === 'gw';
const sizeOf = k => (KIND[k].size === 'ev' ? [EV, EV] : KIND[k].size === 'gw' ? [GW, GW] : [TASK_W, TASK_H]);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function colorOf(n, actor) {
  if (n.tone) return TONE[n.tone];
  if (n.k === 'start' || n.k === 'msgStart' || n.k === 'timerStart' || n.k === 'end') return TONE.ok;
  if (n.k === 'errEnd' || n.k === 'boundaryErr') return TONE.err;
  if (isGw(n)) return TONE.gate;
  const a = ACTOR[actor] || ACTOR.app;
  return { fill: a.fill, stroke: a.stroke };
}
const colorAttrs = c => (c
  ? ` bioc:stroke="${c.stroke}" bioc:fill="${c.fill}" color:background-color="${c.fill}" color:border-color="${c.stroke}"`
  : '');

/* Kurzschreibweise für Sequenzflüsse: F(von, nach, Beschriftung, Optionen) */
const F = (from, to, name, o) => ({ from, to, name: name || '', o: o || {} });

/* Die Prozess-Seite der App. `modell` wählt den Reiter bzw. das Modell. */
const ANSICHT = 'https://bedarfsanfrage.dihag-extern.com/?ansicht=prozess';

/* Ablage im RMS: Bibliothek „ISMS Dokumente" auf sites/ISMS, Ordner SHB. */
const RMS_ABLAGE = 'Prozesse/SHB';

/* Verknüpfte Regelwerke im RMS. Die Kennungen sind die Listen-IDs in
   `Richtlinien` auf sites/IT, dieselben, die `rms.dihag.de/?richtlinie=<id>`
   öffnet. R9 verlangt mindestens eine je Modell. Eine Einkaufs- oder
   Beschaffungsrichtlinie gibt es im RMS noch nicht; die Wertgrenzen stehen
   bisher nur in der App und im Flow. */
const POLICIES = {
  120: 'Konzernfachregelung Prozessmanagement',
  119: 'ISMS-Richtlinie Aufbewahrungsfristen und -pflichten',
  111: 'ISMS-Richtlinie Sichere Softwareentwicklung',
};

/* ═════════════════════════════════════════════════════════════════════════
 *  DIE MODELLE
 * ═════════════════════════════════════════════════════════════════════════ */

const HAUPT = {
  key: 'haupt', file: '01-bedarfsanfrage.bpmn', rmsDatei: 'Bedarfsanfrage.bpmn', tab: 'haupt', haupt: true,
  name: 'Bedarfsanfrage',
  kennung: 'Process_Bedarfsanfrage',
  policies: ['120', '119', '111'],
  neben: ['autoentwurf'],
  doku: 'Ablauf einer Bedarfsanfrage bei SHB von der Erfassung bis zur Bestellung. Der Antragsteller '
      + 'erfasst die Anfrage in der App, gespeichert wird sie in der SharePoint-Liste Bedarfsanfrage auf '
      + 'gruppe_shb. Solange der Einkauf nicht entschieden hat, kann er sie zurückziehen. Öffnet der '
      + 'zuständige Einkauf die Anfrage, setzt die App den Status auf In Prüfung (Einkauf). Der Einkauf '
      + 'genehmigt oder lehnt in der App ab. Nach der Genehmigung läuft der Freigabe-Flow in Power '
      + 'Automate. Ist die Anfrage freigegeben, bestellt der Einkauf im ERP und trägt Bestellnummer, '
      + 'Liefertermin und Preis ein. Dann steht sie auf Bestellt.',
  pools: [
    {
      id: 'P_Banf', name: 'DIHAG SHB · Bedarfsanfrage', main: true,
      lanes: [
        { id: 'LA', name: 'Antragsteller (Fachbereich)', actor: 'mensch' },
        { id: 'LB', name: 'Bedarfsanfrage-App', actor: 'app' },
        { id: 'LC', name: 'Einkauf', actor: 'mensch' },
        { id: 'LD', name: 'Power Automate', actor: 'auto' },
      ],
      nodes: [
        { id: 'A_start', k: 'start', n: 'Bedarf entsteht', lane: 'LA', c: 0, r: 0 },
        { id: 'A_erf', k: 'call', ref: 'erfassen', n: 'Anfrage erfassen und einreichen', lane: 'LA', c: 1, r: 0 },
        { id: 'A_gw', k: 'gw', n: 'Anfrage zurückziehen?', lane: 'LA', c: 2, r: 0 },
        { id: 'A_zur', k: 'user', n: 'Anfrage in „Meine Anfragen" zurückziehen', lane: 'LA', c: 3, r: 0 },
        { id: 'A_endZur', k: 'end', n: 'Anfrage zurückgezogen', lane: 'LA', c: 4, r: 0, tone: 'plan' },
        { id: 'C_open', k: 'user', n: 'Anfrage im Dashboard oder unter Gemeinsame Anfragen öffnen', lane: 'LC', c: 3, r: 0 },
        { id: 'B_adv', k: 'service', n: 'Status auf In Prüfung (Einkauf) setzen', lane: 'LB', c: 4, r: 0 },
        { id: 'C_pruef', k: 'call', ref: 'pruefen', n: 'Anfrage prüfen und entscheiden', lane: 'LC', c: 5, r: 0 },
        { id: 'C_gw', k: 'gw', n: 'Vom Einkauf genehmigt?', lane: 'LC', c: 6, r: 0 },
        { id: 'C_endAb', k: 'end', n: 'Vom Einkauf abgelehnt', lane: 'LC', c: 7, r: 1, tone: 'err' },
        { id: 'D_frei', k: 'call', ref: 'freigabe', n: 'Freigabelauf durchlaufen', lane: 'LD', c: 7, r: 0 },
        { id: 'D_gw', k: 'gw', n: 'Freigegeben?', lane: 'LD', c: 8, r: 0 },
        { id: 'D_endAb', k: 'end', n: 'In der Freigabe abgelehnt', lane: 'LD', c: 8, r: 1, tone: 'err' },
        { id: 'C_best', k: 'user', n: 'Bestellung im ERP auslösen', lane: 'LC', c: 9, r: 0 },
        { id: 'C_daten', k: 'user', n: 'Bestellnummer, Liefertermin und Preis eintragen', lane: 'LC', c: 10, r: 0 },
        { id: 'B_best', k: 'service', n: 'Bestellnummer prüfen und Status Bestellt setzen', lane: 'LB', c: 11, r: 0 },
        { id: 'B_end', k: 'end', n: 'Bedarf bestellt', lane: 'LB', c: 12, r: 0 },
      ],
      flows: [
        F('A_start', 'A_erf'),
        F('A_erf', 'A_gw'),
        F('A_gw', 'A_zur', 'ja'),
        F('A_gw', 'C_open', 'nein'),
        F('A_zur', 'A_endZur'),
        F('C_open', 'B_adv'),
        F('B_adv', 'C_pruef'),
        F('C_pruef', 'C_gw'),
        F('C_gw', 'D_frei', 'ja'),
        F('C_gw', 'C_endAb', 'nein'),
        F('D_frei', 'D_gw'),
        F('D_gw', 'C_best', 'ja'),
        F('D_gw', 'D_endAb', 'nein'),
        F('C_best', 'C_daten'),
        F('C_daten', 'B_best'),
        F('B_best', 'B_end'),
      ],
      annotations: [
        { id: 'An_zur', of: 'A_zur', lane: 'LA', c: 3, r: 1, w: 220, h: 62,
          text: 'Geht nur, solange der Einkauf nicht entschieden hat (Eingereicht oder In Prüfung Einkauf). Die App setzt den Status Zurückgezogen.' },
        { id: 'An_open', of: 'C_open', lane: 'LC', c: 3, r: 1, w: 200, h: 62,
          text: 'Keine Mail bei neuen Anfragen. Was über 48 Std offen ist, markiert die App rot.' },
        { id: 'An_adv', of: 'B_adv', lane: 'LB', c: 4, r: 1, w: 150, h: 62,
          text: 'Nur beim zuständigen Einkauf. Ein Auto-Entwurf wandert genauso in die Prüfung.' },
      ],
    },
    { id: 'P_Lief', name: 'Lieferant', actor: 'extern', blackbox: true },
  ],
  messages: [
    { from: 'C_best', to: 'P_Lief', name: 'Bestellung' },
  ],
};

const ERFASSEN = {
  key: 'erfassen', file: '01a-anfrage-erfassen.bpmn', rmsDatei: 'Bedarfsanfrage-Erfassen.bpmn', tab: 'haupt', von: 'haupt',
  name: 'Anfrage erfassen und einreichen',
  kennung: 'Process_Bedarfsanfrage_Erfassen',
  policies: ['120', '119'],
  doku: 'Der Antragsteller erfasst den Bedarf im Assistenten der App, als Einzelanfrage oder als '
      + 'Sammelanfrage mit mehreren Positionen. Favoriten und Wiedervorlagen füllen den Assistenten vor. '
      + 'Aus Menge und Einzelpreis rechnet die App das Bestellvolumen und zeigt den Genehmigungsweg nach '
      + 'Regelwerk. Ab 500 € verlangt das Regelwerk Angebote, beim Lead-Buyer-Abschluss entfällt die '
      + 'Pflicht. Mit dem Einreichen speichert die App die Anfrage mit dem Status Eingereicht.',
  pools: [
    {
      id: 'P_Erf', name: 'DIHAG SHB · Anfrage erfassen', main: true,
      lanes: [
        { id: 'LA', name: 'Antragsteller (Fachbereich)', actor: 'mensch' },
        { id: 'LB', name: 'Bedarfsanfrage-App', actor: 'app' },
      ],
      nodes: [
        { id: 'A_start', k: 'start', n: 'Bedarf soll angefragt werden', lane: 'LA', c: 0, r: 0, lp: 'above' },
        { id: 'A_gw1', k: 'gw', n: 'Mehrere Positionen?', lane: 'LA', c: 1, r: 0 },
        { id: 'A_einzel', k: 'user', n: 'Einzelanfrage erfassen', lane: 'LA', c: 2, r: 0 },
        { id: 'A_sammel', k: 'user', n: 'Sammelanfrage mit Positionen erfassen', lane: 'LA', c: 2, r: 1 },
        { id: 'A_besch', k: 'user', n: 'Beschaffung, Lieferant und Abteilung angeben', lane: 'LA', c: 3, r: 0 },
        { id: 'B_vol', k: 'service', n: 'Bestellvolumen und Genehmigungsweg anzeigen', lane: 'LB', c: 4, r: 0 },
        { id: 'A_gw2', k: 'gw', n: 'Angebote nötig?', lane: 'LA', c: 5, r: 0 },
        { id: 'A_ang', k: 'user', n: 'Vorhandene Angebote anhängen', lane: 'LA', c: 6, r: 1 },
        { id: 'A_send', k: 'user', n: 'Zusammenfassung prüfen und einreichen', lane: 'LA', c: 7, r: 0 },
        { id: 'B_save', k: 'service', n: 'Anfrage mit Status Eingereicht speichern', lane: 'LB', c: 8, r: 0 },
        { id: 'B_end', k: 'end', n: 'Anfrage eingereicht', lane: 'LB', c: 9, r: 0 },
      ],
      flows: [
        F('A_start', 'A_gw1'),
        F('A_gw1', 'A_einzel', 'nein'),
        F('A_gw1', 'A_sammel', 'ja'),
        F('A_einzel', 'A_besch'),
        F('A_sammel', 'A_besch'),
        F('A_besch', 'B_vol'),
        F('B_vol', 'A_gw2'),
        F('A_gw2', 'A_ang', 'ja'),
        F('A_gw2', 'A_send', 'nein'),
        F('A_ang', 'A_send'),
        F('A_send', 'B_save'),
        F('B_save', 'B_end'),
      ],
      annotations: [
        { id: 'An_fav', of: 'A_start', lane: 'LA', c: 0, r: 1, w: 150, h: 62,
          text: 'Favoriten und Wiedervorlagen füllen den Assistenten vor.' },
        { id: 'An_ang', of: 'B_vol', lane: 'LB', c: 4, r: 1, w: 200, h: 62,
          text: 'Der Genehmigungsweg nennt die Angebote nach Regelwerk: ab 500 € 2 bis 5, beim Lead-Buyer-Abschluss keine.' },
        { id: 'An_save', of: 'B_save', lane: 'LB', c: 8, r: 1, w: 200, h: 62,
          text: 'Ein Datum zur automatischen Bedarfsanfrage merkt die Anfrage für den Auto-Entwurf vor.' },
      ],
    },
  ],
  messages: [],
};

const PRUEFEN = {
  key: 'pruefen', file: '01b-anfrage-pruefen.bpmn', rmsDatei: 'Bedarfsanfrage-Einkaufsprüfung.bpmn', tab: 'haupt', von: 'haupt',
  name: 'Anfrage prüfen und entscheiden',
  kennung: 'Process_Bedarfsanfrage_Einkauf',
  policies: ['120', '119'],
  doku: 'Der Einkauf prüft die Anfrage in der App. Zuständig ist, wem die Warengruppe in der '
      + 'Zugriffssteuerung zugeordnet ist, ohne Zuordnung sind es alle Einkäufer. Fehlen Angebote nach '
      + 'Regelwerk, holt der Einkauf sie beim Lieferanten ein und hängt sie an. Danach genehmigt er oder '
      + 'lehnt mit Grund ab. Die App speichert Status und Entscheider. Eine Mail geht dabei nicht raus, '
      + 'der Antragsteller sieht das Ergebnis in Meine Anfragen.',
  pools: [
    {
      id: 'P_Pruef', name: 'DIHAG SHB · Anfrage prüfen und entscheiden', main: true,
      lanes: [
        { id: 'LA', name: 'Einkauf', actor: 'mensch' },
        { id: 'LB', name: 'Bedarfsanfrage-App', actor: 'app' },
      ],
      nodes: [
        { id: 'A_start', k: 'start', n: 'Anfrage steht auf In Prüfung (Einkauf)', lane: 'LA', c: 0, r: 0 },
        { id: 'A_res', k: 'user', n: 'Anfrage reservieren und Angaben prüfen', lane: 'LA', c: 1, r: 0 },
        { id: 'A_gw1', k: 'gw', n: 'Angebote vollständig?', lane: 'LA', c: 2, r: 0 },
        { id: 'A_ang', k: 'user', n: 'Angebote beim Lieferanten einholen und anhängen', lane: 'LA', c: 3, r: 1 },
        { id: 'A_bew', k: 'user', n: 'Bedarf, Preis und Lieferant bewerten', lane: 'LA', c: 4, r: 0 },
        { id: 'A_gw2', k: 'gw', n: 'Bedarf genehmigen?', lane: 'LA', c: 5, r: 0 },
        { id: 'A_grund', k: 'user', n: 'Ablehnungsgrund eintragen', lane: 'LA', c: 6, r: 0 },
        { id: 'B_ok', k: 'service', n: 'Status Genehmigt (Einkauf) speichern', lane: 'LB', c: 6, r: 0 },
        { id: 'B_endOk', k: 'end', n: 'Vom Einkauf genehmigt', lane: 'LB', c: 7, r: 0 },
        { id: 'B_ab', k: 'service', n: 'Status Abgelehnt mit Grund speichern', lane: 'LB', c: 7, r: 1 },
        { id: 'B_endAb', k: 'end', n: 'Vom Einkauf abgelehnt', lane: 'LB', c: 8, r: 1, tone: 'err' },
      ],
      flows: [
        F('A_start', 'A_res'),
        F('A_res', 'A_gw1'),
        F('A_gw1', 'A_bew', 'ja'),
        F('A_gw1', 'A_ang', 'nein'),
        F('A_ang', 'A_bew'),
        F('A_bew', 'A_gw2'),
        F('A_gw2', 'B_ok', 'ja'),
        F('A_gw2', 'A_grund', 'nein'),
        F('B_ok', 'B_endOk'),
        F('A_grund', 'B_ab'),
        F('B_ab', 'B_endAb'),
      ],
      annotations: [
        { id: 'An_zust', of: 'A_res', lane: 'LA', c: 1, r: 1, w: 150, h: 62,
          text: 'Zuständig ist, wem die Warengruppe zugeordnet ist. Reservieren hält die Anfrage 8 Std.' },
        { id: 'An_mail', of: 'A_grund', lane: 'LA', c: 6, r: 1, w: 150, h: 62,
          text: 'Keine Mail. Der Antragsteller sieht Status und Grund in Meine Anfragen.' },
      ],
    },
    { id: 'P_Lief', name: 'Lieferant', actor: 'extern', blackbox: true },
  ],
  messages: [
    { from: 'A_ang', to: 'P_Lief', name: 'Angebotsanfrage' },
  ],
};

const FREIGABE = {
  key: 'freigabe', file: '02-freigabelauf.bpmn', rmsDatei: 'Bedarfsanfrage-Freigabelauf.bpmn', tab: 'haupt', von: 'haupt',
  name: 'Freigabelauf (Power Automate)',
  kennung: 'Process_Bedarfsanfrage_Freigabelauf',
  policies: ['120', '119'],
  doku: 'Power-Automate-Flow „Bedarfsanfragen SHB". Er startet, sobald eine Anfrage auf Genehmigt '
      + '(Einkauf) steht. Das Controlling prüft nur Dienstleistungen, die Werkleitung ab 1.500 € '
      + 'Bestellvolumen, der strategische Einkauf ab 10.000 € und nur ohne Lead-Buyer-Abschluss. Jede '
      + 'Stufe bekommt eine Genehmigung in Teams und per Mail, die erste Antwort zählt. Bei Ablehnung '
      + 'setzt der Flow den Status Abgelehnt, sonst Freigegeben, und schreibt dem Antragsteller in beiden '
      + 'Fällen eine Mail.',
  pools: [
    {
      id: 'P_Frei', name: 'DIHAG SHB · Freigabelauf', main: true,
      lanes: [
        { id: 'LA', name: 'Power Automate (Flow Bedarfsanfragen SHB)', actor: 'auto' },
        { id: 'LB', name: 'Controlling', actor: 'mensch' },
        { id: 'LC', name: 'Werkleitung', actor: 'mensch' },
        { id: 'LD', name: 'Strategischer Einkauf', actor: 'mensch' },
      ],
      nodes: [
        { id: 'A_start', k: 'start', n: 'Status steht auf Genehmigt (Einkauf)', lane: 'LA', c: 0, r: 0 },
        { id: 'A_gw1', k: 'gw', n: 'Dienstleistung?', lane: 'LA', c: 1, r: 0 },
        { id: 'A_ctrl', k: 'service', n: 'Controlling um Freigabe bitten', lane: 'LA', c: 2, r: 0 },
        { id: 'B_pruef', k: 'user', n: 'Dienstleistung prüfen', lane: 'LB', c: 3, r: 0 },
        { id: 'B_gw', k: 'gw', n: 'Vom Controlling freigegeben?', lane: 'LB', c: 4, r: 0 },
        { id: 'A_gw2', k: 'gw', n: 'Ab 1.500 € Bestellvolumen?', lane: 'LA', c: 5, r: 0 },
        { id: 'A_wl', k: 'service', n: 'Werkleitung um Freigabe bitten', lane: 'LA', c: 6, r: 0 },
        { id: 'C_pruef', k: 'user', n: 'Bestellvolumen prüfen', lane: 'LC', c: 7, r: 0 },
        { id: 'C_gw', k: 'gw', n: 'Von der Werkleitung freigegeben?', lane: 'LC', c: 8, r: 0 },
        { id: 'A_gw3', k: 'gw', n: 'Ab 10.000 € und ohne Lead-Buyer-Abschluss?', lane: 'LA', c: 9, r: 0 },
        { id: 'A_se', k: 'service', n: 'Strategischen Einkauf um Freigabe bitten', lane: 'LA', c: 10, r: 0 },
        { id: 'D_pruef', k: 'user', n: 'Konzernrahmen prüfen', lane: 'LD', c: 11, r: 0 },
        { id: 'D_gw', k: 'gw', n: 'Vom strategischen Einkauf freigegeben?', lane: 'LD', c: 12, r: 0 },
        { id: 'A_frei', k: 'service', n: 'Freigeben und Antragsteller benachrichtigen', lane: 'LA', c: 13, r: 0 },
        { id: 'A_endOk', k: 'end', n: 'Freigegeben', lane: 'LA', c: 14, r: 0 },
        { id: 'A_ab', k: 'service', n: 'Ablehnen und Antragsteller benachrichtigen', lane: 'LA', c: 13, r: 1 },
        { id: 'A_endAb', k: 'end', n: 'Abgelehnt', lane: 'LA', c: 14, r: 1, tone: 'err' },
      ],
      flows: [
        F('A_start', 'A_gw1'),
        F('A_gw1', 'A_ctrl', 'ja'),
        F('A_gw1', 'A_gw2', 'nein', { via: 'above' }),
        F('A_ctrl', 'B_pruef'),
        F('B_pruef', 'B_gw'),
        F('B_gw', 'A_gw2', 'ja'),
        F('B_gw', 'A_ab', 'nein'),
        F('A_gw2', 'A_wl', 'ja'),
        F('A_gw2', 'A_frei', 'nein', { via: 'below' }),
        F('A_wl', 'C_pruef'),
        F('C_pruef', 'C_gw'),
        F('C_gw', 'A_gw3', 'ja'),
        F('C_gw', 'A_ab', 'nein'),
        F('A_gw3', 'A_se', 'ja'),
        F('A_gw3', 'A_frei', 'nein', { via: 'below' }),
        F('A_se', 'D_pruef'),
        F('D_pruef', 'D_gw'),
        F('D_gw', 'A_frei', 'ja'),
        F('D_gw', 'A_ab', 'nein'),
        F('A_frei', 'A_endOk'),
        F('A_ab', 'A_endAb'),
      ],
      annotations: [
        { id: 'An_appr', of: 'A_ctrl', lane: 'LA', c: 2, r: 1, w: 150, h: 76,
          text: 'Setzt den Status In Prüfung und schickt die Genehmigung in Teams und per Mail. Die erste Antwort zählt.' },
      ],
    },
  ],
  messages: [],
};

const AUTOENTWURF = {
  key: 'autoentwurf', file: '03-auto-entwurf.bpmn', rmsDatei: 'Bedarfsanfrage-AutoEntwurf.bpmn', tab: 'autoentwurf', haupt: true,
  name: 'Auto-Entwurf (Wiedervorlage)',
  kennung: 'Process_Bedarfsanfrage_AutoEntwurf',
  policies: ['120', '111'],
  doku: 'Für wiederkehrende Bedarfe. Hat eine Anfrage ein Datum zur automatischen Bedarfsanfrage, legt '
      + 'ein täglicher Lauf in GitHub Actions an diesem Tag eine Kopie als Entwurf an (Status '
      + 'Automatisch erstellter Entwurf) und leert das Datum an der Vorlage, damit kein zweiter Entwurf '
      + 'entsteht. Öffnet der Einkauf den Entwurf, wandert er wie eine eingereichte Anfrage in die Prüfung.',
  pools: [
    {
      id: 'P_Auto', name: 'DIHAG SHB · Auto-Entwurf', main: true,
      lanes: [
        { id: 'LA', name: 'Automatik (GitHub Actions)', actor: 'auto' },
      ],
      nodes: [
        { id: 'A_start', k: 'timerStart', n: 'Täglich um 6 Uhr', lane: 'LA', c: 0, r: 0, lp: 'above' },
        { id: 'A_suche', k: 'service', n: 'Fällige Wiedervorlagen suchen', lane: 'LA', c: 1, r: 0 },
        { id: 'A_gw', k: 'gw', n: 'Fällige Anfrage gefunden?', lane: 'LA', c: 2, r: 0 },
        { id: 'A_nichts', k: 'end', n: 'Nichts fällig', lane: 'LA', c: 3, r: 1, tone: 'plan' },
        { id: 'A_kopie', k: 'service', n: 'Kopie als Entwurf anlegen', lane: 'LA', c: 3, r: 0 },
        { id: 'A_leeren', k: 'service', n: 'Datum an der Vorlage leeren', lane: 'LA', c: 4, r: 0 },
        { id: 'A_end', k: 'end', n: 'Entwurf liegt bereit', lane: 'LA', c: 5, r: 0, lp: 'above' },
      ],
      flows: [
        F('A_start', 'A_suche'),
        F('A_suche', 'A_gw'),
        F('A_gw', 'A_kopie', 'ja'),
        F('A_gw', 'A_nichts', 'nein'),
        F('A_kopie', 'A_leeren'),
        F('A_leeren', 'A_end'),
      ],
      annotations: [
        { id: 'An_takt', of: 'A_start', lane: 'LA', c: 0, r: 1, w: 150, h: 62,
          text: '5 Uhr UTC, im Sommer also 7 Uhr. GitHub startet oft einige Minuten später.' },
        { id: 'An_weiter', of: 'A_end', lane: 'LA', c: 5, r: 1, w: 220, h: 62,
          text: 'Status Automatisch erstellter Entwurf. Öffnet der Einkauf ihn, wandert er in die Prüfung (Hauptprozess).' },
      ],
      groups: [
        { id: 'G_je', members: ['A_kopie', 'A_leeren'], label: 'je fällige Anfrage' },
      ],
    },
  ],
  messages: [],
};

const MODELLE = [HAUPT, ERFASSEN, PRUEFEN, FREIGABE, AUTOENTWURF];
const BY_KEY = Object.fromEntries(MODELLE.map(m => [m.key, m]));

/* ═════════════════════════════════════════════════════════════════════════
 *  LAYOUT
 * ═════════════════════════════════════════════════════════════════════════ */

function placeNode(n, cx, cy) {
  const [w, h] = sizeOf(n.k);
  n.b = { x: rnd(cx - w / 2), y: rnd(cy - h / 2), w, h };
}

/** Pools und Bahnen anordnen, Knoten platzieren. Geschlossene Pools (externe
 *  Partner) sind nur ein Band ohne Inhalt. */
function layoutMain(spec) {
  const N = {};
  let maxCol = 0;
  for (const p of spec.pools) {
    for (const n of p.nodes || []) if (n.c != null) maxCol = Math.max(maxCol, n.c);
    for (const a of p.annotations || []) maxCol = Math.max(maxCol, a.c + Math.ceil((a.w || 160) / COL_W) - 1);
  }
  const width = (CONTENT_X - POOL_X) + (maxCol + 1) * COL_W + 30;
  let y = TOP;
  for (const p of spec.pools) {
    p.x = POOL_X; p.y = y; p.w = width; p.laneGeo = [];
    if (p.blackbox) {
      p.h = BLACKBOX_H;
      y += p.h + POOL_GAP;
      continue;
    }
    for (const ld of p.lanes) {
      let rows = ld.rows || 1;
      for (const n of p.nodes) if (n.lane === ld.id && n.r != null) rows = Math.max(rows, n.r + 1);
      for (const a of p.annotations || []) if (a.lane === ld.id) rows = Math.max(rows, a.r + 1);
      const h = rows * ROW_H + 20;
      p.laneGeo.push({ ...ld, x: POOL_X + POOL_HEAD, y, w: width - POOL_HEAD, h });
      y += h;
    }
    p.h = y - p.y;
    y += POOL_GAP;
    const laneOf = o => p.laneGeo.find(l => l.id === o.lane);
    p.colLeft = c => CONTENT_X + c * COL_W;
    p.rowTop = (o, r) => laneOf(o).y + 10 + r * ROW_H;
    for (const n of p.nodes) {
      if (n.k === 'boundaryErr') continue;
      const lane = laneOf(n);
      if (!lane) throw new Error(`${spec.key}: Bahn "${n.lane}" für ${n.id} fehlt`);
      n.pool = p; n.actor = lane.actor;
      placeNode(n, CONTENT_X + n.c * COL_W + COL_W / 2, lane.y + 10 + n.r * ROW_H + ROW_H / 2);
      N[n.id] = n;
    }
    for (const n of p.nodes.filter(x => x.k === 'boundaryErr')) {
      const h = N[n.attachedTo];
      n.pool = p; n.actor = h.actor;
      placeNode(n, R(h.b) - 28, B(h.b));
      N[n.id] = n;
    }
  }
  return N;
}

function checkOverlaps(nodes, where) {
  const seen = new Map();
  for (const n of nodes) {
    if (n.k === 'boundaryErr') continue;
    const key = `${n.lane || ''}|${n.c}|${n.r}`;
    if (seen.has(key)) throw new Error(`${where}: ${n.id} liegt auf demselben Rasterplatz wie ${seen.get(key)}`);
    seen.set(key, n.id);
  }
}

/* ── Kanten führen (orthogonal) ───────────────────────────────────────── */

function route(f, N, inCnt, outCnt) {
  const s = N[f.from], t = N[f.to];
  if (!s || !t) throw new Error(`Fluss ${f.from} -> ${f.to}: Knoten fehlt`);
  const sb = s.b, tb = t.b, o = f.o;
  const sx = CX(sb), sy = CY(sb), tx = CX(tb), ty = CY(tb);
  // Übergang in ein tieferes Band: unten raus, durch die freie Zeile, oben rein
  // (rechts raus, damit Ein- und Ausgänge nicht dieselbe Kante teilen)
  if (o.viaRow != null) {
    const yv = t.pool.rowTop(t, o.viaRow) + ROW_H / 2, xr = R(sb) + 28;
    return [[R(sb), sy], [xr, sy], [xr, yv], [tx, yv], [tx, tb.y]];
  }
  // Senkrecht aus der Aufgabe in die Zeile des Ziels, dann waagerecht hinein
  if (o.out === 'vertical' && tx > sx) return [[sx, ty < sy ? sb.y : B(sb)], [sx, ty], [tb.x, ty]];
  if (o.via === 'above' || o.via === 'below') {
    const up = o.via === 'above';
    const g = o.gap || 28;
    const yv = up ? Math.min(sb.y, tb.y) - g : Math.max(B(sb), B(tb)) + g;
    return [[sx, up ? sb.y : B(sb)], [sx, yv], [tx, yv], [tx, up ? tb.y : B(tb)]];
  }
  if (o.via === 'left') return [[sb.x, sy], [tx, sy], [tx, ty < sy ? B(tb) : tb.y]];
  if (s.k === 'boundaryErr') return [[sx, B(sb)], [sx, ty], [tx < sx ? R(tb) : tb.x, ty]];
  // Von einer Entscheidung schräg nach unten rechts, oben in das Ziel
  if (o.in === 'top' && tx > sx && ty > sy) {
    const yv = tb.y - 30;
    return [[sx, B(sb)], [sx, yv], [tx, yv], [tx, tb.y]];
  }

  const sameRow = Math.abs(sy - ty) < 1, sameCol = Math.abs(sx - tx) < 1;
  if (sameRow) {
    if (tx > sx) return [[R(sb), sy], [tb.x, ty]];
    const yv = Math.max(B(sb), B(tb)) + 28;
    return [[sx, B(sb)], [sx, yv], [tx, yv], [tx, B(tb)]];
  }
  if (sameCol) return ty > sy ? [[sx, B(sb)], [tx, tb.y]] : [[sx, sb.y], [tx, B(tb)]];

  const split = isGw(s) && outCnt[s.id] > 1;
  const join = isGw(t) && inCnt[t.id] > 1;
  if (tx > sx) {
    if (split) return [[sx, ty < sy ? sb.y : B(sb)], [sx, ty], [tb.x, ty]];
    if (o.in || join) {
      const fromBelow = o.in ? o.in === 'bottom' : ty < sy;
      return [[R(sb), sy], [tx, sy], [tx, fromBelow ? B(tb) : tb.y]];
    }
    const mx = rnd((R(sb) + tb.x) / 2);
    return [[R(sb), sy], [mx, sy], [mx, ty], [tb.x, ty]];
  }
  // Rücksprung nach links in eine andere Zeile: außen herum
  const yv = ty > sy ? Math.max(B(sb), B(tb)) + 28 : Math.min(sb.y, tb.y) - 28;
  return ty > sy
    ? [[sx, B(sb)], [sx, yv], [tx, yv], [tx, B(tb)]]
    : [[sx, sb.y], [sx, yv], [tx, yv], [tx, tb.y]];
}

function labelLines(text, perLine) {
  return Math.max(1, Math.ceil(String(text).length / perLine));
}

function flowLabel(wp, name) {
  const lines = name.length > 16 ? 2 : 1;
  const w = Math.min(112, Math.max(24, rnd(name.length * 6.4 / lines) + 12));
  const h = lines === 2 ? 27 : 14;
  const [a, b] = wp;
  if (wp.length >= 3 && Math.abs(a[0] - b[0]) < 1 && Math.abs(wp[1][1] - wp[2][1]) < 1) {
    const c = wp[2];
    return { x: c[0] > b[0] ? b[0] + 8 : b[0] - w - 8, y: b[1] - h - 3, w, h };
  }
  if (Math.abs(a[1] - b[1]) < 1) return { x: b[0] > a[0] ? a[0] + 6 : a[0] - w - 6, y: a[1] - h - 3, w, h };
  return { x: a[0] + 6, y: rnd((a[1] + b[1]) / 2 - h / 2), w, h };
}

function nodeLabel(n, flows, N) {
  const b = n.b;
  if (!n.n || KIND[n.k].size === 'task') return null;
  if (isGw(n)) {
    const lines = labelLines(n.n, 16);
    const h = lines * 13 + 2, w = 96;
    let below = n.lp === 'below';
    if (!n.lp) below = flows.some(f => f.to === n.id && CY(N[f.from].b) < CY(b) - 1);
    return { x: rnd(CX(b) - 100), y: below ? B(b) + 4 : b.y - h - 4, w, h };
  }
  if (n.k === 'boundaryErr') return { x: b.x - 112, y: B(b) - 2, w: 108, h: 27 };
  const lines = labelLines(n.n, 20);
  const h = lines * 13 + 2;
  if (n.lp === 'above') return { x: rnd(CX(b) - 60), y: b.y - h - 5, w: 120, h };
  return { x: rnd(CX(b) - 60), y: B(b) + 5, w: 120, h };
}

/* ── Anmerkungen, Gruppen ─────────────────────────────────────────────── */

function placeAnnotation(a, host) {
  const w = a.w || 160, h = a.h || 56;
  a.b = { x: host.colLeft(a.c) + 10, y: rnd(host.rowTop(a, a.r) + (ROW_H - h) / 2), w, h };
}

function assocPath(nb, ab) {
  const cl = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const ncx = CX(nb), ncy = CY(nb);
  if (ab.y >= B(nb)) return [[ncx, B(nb)], [cl(ncx, ab.x + 12, R(ab) - 12), ab.y]];
  if (B(ab) <= nb.y) return [[ncx, nb.y], [cl(ncx, ab.x + 12, R(ab) - 12), B(ab)]];
  if (ab.x >= R(nb)) return [[R(nb), ncy], [ab.x, cl(ncy, ab.y + 8, B(ab) - 8)]];
  return [[nb.x, ncy], [R(ab), cl(ncy, ab.y + 8, B(ab) - 8)]];
}

function groupBounds(g, byId) {
  const bs = g.members.map(id => {
    const o = byId[id];
    if (!o || !o.b) throw new Error(`Gruppe ${g.id}: Mitglied ${id} fehlt`);
    return o.b;
  });
  const x = Math.min(...bs.map(b => b.x)) - 18, y = Math.min(...bs.map(b => b.y)) - 24;
  const r = Math.max(...bs.map(R)) + 18, bo = Math.max(...bs.map(B)) + 16;
  return { x, y, w: r - x, h: bo - y };
}

/* ═════════════════════════════════════════════════════════════════════════
 *  XML
 * ═════════════════════════════════════════════════════════════════════════ */

const fmtB = b => `<dc:Bounds x="${rnd(b.x)}" y="${rnd(b.y)}" width="${rnd(b.w)}" height="${rnd(b.h)}" />`;
const fmtWp = wp => wp.map(([x, y]) => `<di:waypoint x="${rnd(x)}" y="${rnd(y)}" />`).join('');
const labelXml = lb => (lb ? `<bpmndi:BPMNLabel>${fmtB(lb)}</bpmndi:BPMNLabel>` : '');

function counts(flows) {
  const inCnt = {}, outCnt = {};
  for (const f of flows) {
    outCnt[f.from] = (outCnt[f.from] || 0) + 1;
    inCnt[f.to] = (inCnt[f.to] || 0) + 1;
  }
  return { inCnt, outCnt };
}

function flowId(f) { return `Flow_${f.from}_${f.to}`; }

/** Titel eines Modells, wie das RMS ihn zeigt: der Dateiname ohne Endung. */
function rmsTitel(m) { return m.rmsDatei.replace(/\.bpmn$/i, ''); }

/** Text der Dokumentation einer Aufrufaktivität, wie das RMS ihn schreibt:
 *  Klartextzeile plus Marker mit der Datei-Kennung im RMS. */
function callDoku(n, rms) {
  const ziel = BY_KEY[n.ref];
  if (!ziel) throw new Error(`Aufruf ${n.id}: Modell "${n.ref}" unbekannt`);
  const zeilen = ['Unterprozess: ' + rmsTitel(ziel)];
  const itemId = rms.modelle && rms.modelle[n.ref];
  if (itemId) zeilen.push(`[[rms:modell=${itemId}]]`);
  return zeilen.join('\n');
}

/** Semantik eines Knotens. */
function nodeSem(n, flows, rms) {
  const K = KIND[n.k];
  const attrs = [`id="${n.id}"`];
  if (n.n) attrs.push(`name="${esc(n.n)}"`);
  if (n.k === 'boundaryErr') attrs.push(`attachedToRef="${n.attachedTo}"`, 'cancelActivity="true"');
  if (n.k === 'call') attrs.push(`calledElement="${BY_KEY[n.ref].kennung}"`);
  let body = '';
  if (n.k === 'call') body += `<bpmn:documentation>${esc(callDoku(n, rms))}</bpmn:documentation>`;
  for (const f of flows) if (f.to === n.id) body += `<bpmn:incoming>${flowId(f)}</bpmn:incoming>`;
  for (const f of flows) if (f.from === n.id) body += `<bpmn:outgoing>${flowId(f)}</bpmn:outgoing>`;
  if (K.def === 'message') body += `<bpmn:messageEventDefinition id="${n.id}_def" />`;
  if (K.def === 'timer') body += `<bpmn:timerEventDefinition id="${n.id}_def" />`;
  if (K.def === 'error') body += `<bpmn:errorEventDefinition id="${n.id}_def" errorRef="Error_Pruefung" />`;
  return `      <bpmn:${K.tag} ${attrs.join(' ')}>${body}</bpmn:${K.tag}>\n`;
}

function flowsSem(flows) {
  return flows.map(f => `      <bpmn:sequenceFlow id="${flowId(f)}" sourceRef="${f.from}" targetRef="${f.to}"${f.name ? ` name="${esc(f.name)}"` : ''} />\n`).join('');
}

function artifactsSem(annotations, groups) {
  let x = '';
  for (const a of annotations || []) {
    x += `      <bpmn:textAnnotation id="${a.id}"><bpmn:text>${esc(a.text)}</bpmn:text></bpmn:textAnnotation>\n`;
    x += `      <bpmn:association id="Assoc_${a.id}" associationDirection="None" sourceRef="${a.of}" targetRef="${a.id}" />\n`;
  }
  for (const g of groups || []) x += `      <bpmn:group id="${g.id}" categoryValueRef="CV_${g.id}" />\n`;
  return x;
}

/* Dokumentation des Prozesses. Bis zur Zeile „Anwender-Ansicht" schreibt sie
   dieser Generator. Was danach steht, kann das RMS selbst pflegen
   (Nebenprozesse, Regelwerke, Prozessmanagement, Kennzahlen); beim
   Zurückspielen ins RMS gilt dort die Fassung aus dem RMS. Die Marker
   schreibt dieser Generator so, wie das RMS es selbst tut, sonst gingen die
   Verknüpfungen beim nächsten Speichern im RMS verloren. */
function processDoku(spec, rms) {
  const zeilen = [spec.doku];
  if (spec.von) zeilen.push(`Eingebunden in: ${rmsTitel(BY_KEY[spec.von])}.`);
  zeilen.push(`Anwender-Ansicht: ${ANSICHT}&modell=${spec.key}`);
  const neben = (spec.neben || []).filter(k => rms.modelle && rms.modelle[k]);
  if (neben.length) {
    zeilen.push('Nebenprozesse: ' + neben.map(k => rmsTitel(BY_KEY[k])).join('; '));
    zeilen.push(`[[rms:neben=${neben.map(k => rms.modelle[k]).join(',')}]]`);
  }
  const ids = spec.policies || [];
  if (ids.length) {
    zeilen.push('Im Einklang mit den Richtlinien: ' + ids.map(i => POLICIES[i] || i).join('; '));
    zeilen.push(`[[rms:policies=${ids.join(',')}]]`);
  }
  return zeilen.join('\n');
}

/** DI-Formen und -Kanten für eine Menge Knoten/Flüsse/Anmerkungen. */
function planeDi(nodes, flows, N, annotations, groups, byId) {
  const { inCnt, outCnt } = counts(flows);
  let x = '';
  for (const n of nodes) {
    const c = colorOf(n, n.actor);
    const extra = isGw(n) ? ' isMarkerVisible="true"' : '';
    x += `      <bpmndi:BPMNShape id="${n.id}_di" bpmnElement="${n.id}"${extra}${colorAttrs(c)}>${fmtB(n.b)}${labelXml(nodeLabel(n, flows, N))}</bpmndi:BPMNShape>\n`;
  }
  for (const f of flows) {
    const wp = route(f, N, inCnt, outCnt);
    const lb = f.name ? flowLabel(wp, f.name) : null;
    x += `      <bpmndi:BPMNEdge id="${flowId(f)}_di" bpmnElement="${flowId(f)}">${fmtWp(wp)}${labelXml(lb)}</bpmndi:BPMNEdge>\n`;
  }
  for (const a of annotations || []) {
    const c = a.tone ? TONE[a.tone] : { fill: '#FFFFFF', stroke: '#6B7280' };
    x += `      <bpmndi:BPMNShape id="${a.id}_di" bpmnElement="${a.id}"${colorAttrs(c)}>${fmtB(a.b)}</bpmndi:BPMNShape>\n`;
    x += `      <bpmndi:BPMNEdge id="Assoc_${a.id}_di" bpmnElement="Assoc_${a.id}">${fmtWp(assocPath(N[a.of].b, a.b))}</bpmndi:BPMNEdge>\n`;
  }
  for (const g of groups || []) {
    const gb = groupBounds(g, byId);
    x += `      <bpmndi:BPMNShape id="${g.id}_di" bpmnElement="${g.id}" bioc:stroke="#17509E" color:border-color="#17509E">${fmtB(gb)}`
       + `${labelXml({ x: gb.x + 10, y: gb.y + 6, w: 260, h: 14 })}</bpmndi:BPMNShape>\n`;
  }
  return x;
}

function build(spec, rms) {
  const N = layoutMain(spec);
  const categories = [];
  let usesError = false;
  const inner = spec.pools.filter(p => !p.blackbox);
  if (inner.length !== 1) throw new Error(`${spec.key}: genau ein eigener Pool erwartet`);

  for (const p of inner) {
    checkOverlaps(p.nodes, `${spec.key}/${p.id}`);
    for (const a of p.annotations || []) placeAnnotation(a, p);
  }

  // Semantik
  let collab = `  <bpmn:collaboration id="Collab_${spec.key}">\n`;
  for (const p of spec.pools) {
    collab += p.blackbox
      ? `    <bpmn:participant id="${p.id}" name="${esc(p.name)}" />\n`
      : `    <bpmn:participant id="${p.id}" name="${esc(p.name)}" processRef="${spec.kennung}" />\n`;
  }
  for (const m of spec.messages) collab += `    <bpmn:messageFlow id="Msg_${m.from}_${m.to}" name="${esc(m.name)}" sourceRef="${m.from}" targetRef="${m.to}" />\n`;
  collab += '  </bpmn:collaboration>\n';

  let procs = '';
  for (const p of inner) {
    let x = `  <bpmn:process id="${spec.kennung}" name="${esc(spec.name)}" isExecutable="false">\n`;
    x += `    <bpmn:documentation>${esc(processDoku(spec, rms))}</bpmn:documentation>\n`;
    x += `    <bpmn:laneSet id="LaneSet_${p.id}">\n`;
    for (const l of p.lanes) {
      x += `      <bpmn:lane id="${l.id}" name="${esc(l.name)}">`;
      for (const n of p.nodes.filter(n => n.lane === l.id)) x += `<bpmn:flowNodeRef>${n.id}</bpmn:flowNodeRef>`;
      x += '</bpmn:lane>\n';
    }
    x += '    </bpmn:laneSet>\n';
    for (const n of p.nodes) {
      if (KIND[n.k].def === 'error') usesError = true;
      x += nodeSem(n, p.flows, rms);
    }
    x += flowsSem(p.flows);
    x += artifactsSem(p.annotations, p.groups);
    for (const g of p.groups || []) categories.push(g);
    x += '  </bpmn:process>\n';
    procs += x;
  }

  // DI
  let di = `  <bpmndi:BPMNDiagram id="Diagram_${spec.key}">\n    <bpmndi:BPMNPlane id="Plane_${spec.key}" bpmnElement="Collab_${spec.key}">\n`;
  const byId = { ...N };
  for (const p of inner) for (const a of p.annotations || []) byId[a.id] = a;
  for (const p of spec.pools) {
    const a = ACTOR[p.actor] || null;
    const pc = p.blackbox ? { fill: a.lane, stroke: a.stroke } : { fill: '#FFFFFF', stroke: '#1A2644' };
    di += `      <bpmndi:BPMNShape id="${p.id}_di" bpmnElement="${p.id}" isHorizontal="true"${colorAttrs(pc)}>${fmtB(p)}</bpmndi:BPMNShape>\n`;
    for (const l of p.laneGeo) {
      const ac = ACTOR[l.actor];
      di += `      <bpmndi:BPMNShape id="${l.id}_di" bpmnElement="${l.id}" isHorizontal="true"${colorAttrs({ fill: ac.lane, stroke: ac.stroke })}>${fmtB(l)}</bpmndi:BPMNShape>\n`;
    }
  }
  for (const p of inner) di += planeDi(p.nodes, p.flows, N, p.annotations, p.groups, byId);
  // Nachrichtenflüsse: Knoten oder geschlossener Pool auf beiden Seiten
  const poolById = Object.fromEntries(spec.pools.map(p => [p.id, p]));
  for (const m of spec.messages) {
    const sPool = poolById[m.from], tPool = poolById[m.to];
    const sNode = N[m.from], tNode = N[m.to];
    let wp, lx, ly;
    if (sPool) {   // aus einem geschlossenen Pool auf einen Knoten
      const tb = tNode.b, tx = CX(tb);
      const down = tNode.pool.y > sPool.y;
      wp = [[tx, down ? sPool.y + sPool.h : sPool.y], [tx, down ? tb.y : B(tb)]];
      lx = tx + 8; ly = rnd((wp[0][1] + wp[1][1]) / 2 - 14);
    } else {
      const sb = sNode.b, sx = CX(sb) + (m.dx || 0), sp = sNode.pool;
      const tp = tPool || tNode.pool;
      const down = tp.y > sp.y;
      const gapMid = down ? (sp.y + sp.h + tp.y) / 2 : (tp.y + tp.h + sp.y) / 2;
      if (tPool) wp = [[sx, down ? B(sb) : sb.y], [sx, down ? tPool.y : tPool.y + tPool.h]];
      else {
        const tb = tNode.b, tx = CX(tb);
        const y1 = down ? B(sb) : sb.y, y2 = down ? tb.y : B(tb);
        wp = Math.abs(sx - tx) < 1 ? [[sx, y1], [tx, y2]] : [[sx, y1], [sx, gapMid], [tx, gapMid], [tx, y2]];
      }
      lx = sx + 8; ly = rnd(gapMid - 14);
    }
    di += `      <bpmndi:BPMNEdge id="Msg_${m.from}_${m.to}_di" bpmnElement="Msg_${m.from}_${m.to}">${fmtWp(wp)}${labelXml({ x: lx, y: ly, w: 170, h: 27 })}</bpmndi:BPMNEdge>\n`;
  }
  di += '    </bpmndi:BPMNPlane>\n  </bpmndi:BPMNDiagram>\n';

  let defs = '';
  if (usesError || inner.some(p => p.nodes.some(n => n.k === 'boundaryErr')))
    defs += '  <bpmn:error id="Error_Pruefung" name="Prüfung nicht bestanden" errorCode="PRUEFUNG" />\n';
  for (const g of categories) defs += `  <bpmn:category id="Cat_${g.id}"><bpmn:categoryValue id="CV_${g.id}" value="${esc(g.label)}" /></bpmn:category>\n`;

  return '<?xml version="1.0" encoding="UTF-8"?>\n'
    + '<bpmn:definitions xmlns:bpmn="http://www.omg.org/spec/BPMN/20100524/MODEL" '
    + 'xmlns:bpmndi="http://www.omg.org/spec/BPMN/20100524/DI" xmlns:dc="http://www.omg.org/spec/DD/20100524/DC" '
    + 'xmlns:di="http://www.omg.org/spec/DD/20100524/DI" xmlns:bioc="http://bpmn.io/schema/bpmn/biocolor/1.0" '
    + 'xmlns:color="http://www.omg.org/spec/BPMN/non-normative/color/1.0" '
    + `id="Definitions_${spec.key}" targetNamespace="https://bedarfsanfrage.dihag-extern.com/bpmn" `
    + 'exporter="DIHAG BPMN-Generator (scripts/bpmn/generate-bpmn.js)" exporterVersion="2.0">\n'
    + defs + collab + procs + di
    + '</bpmn:definitions>\n';
}


/* ── Ausgabe ──────────────────────────────────────────────────────────── */
if (require.main === module) {
  const outDir = path.join(__dirname, '..', '..', 'docs', 'bpmn');
  fs.mkdirSync(outDir, { recursive: true });
  let rms = {};
  try { rms = JSON.parse(fs.readFileSync(path.join(outDir, 'rms-ids.json'), 'utf8')); } catch (e) { /* noch kein Import */ }
  for (const spec of MODELLE) {
    const xml = build(spec, rms);
    fs.writeFileSync(path.join(outDir, spec.file), xml, 'utf8');
    console.log(`geschrieben: docs/bpmn/${spec.file} (${(xml.length / 1024).toFixed(1)} KB)`);
  }
  // Modell-Liste für die Prozess-Seite: welche Datei, welche Kennung, wo im RMS
  const liste = {
    hinweis: 'Erzeugt von scripts/bpmn/generate-bpmn.js. Führend sind die Modelle im RMS.',
    rms: { url: 'https://rms.dihag.de', ablage: rms.ablage || RMS_ABLAGE, driveId: rms.driveId || '' },
    richtlinien: POLICIES,
    modelle: MODELLE.map(m => ({
      key: m.key, name: m.name, kennung: m.kennung, datei: 'docs/bpmn/' + m.file, rmsDatei: m.rmsDatei,
      tab: m.tab, haupt: !!m.haupt, von: m.von || '',
      unter: (m.pools.find(p => !p.blackbox).nodes || []).filter(n => n.k === 'call').map(n => n.ref),
      neben: m.neben || [],
      policies: m.policies || [],
      itemId: (rms.modelle && rms.modelle[m.key]) || '',
    })),
  };
  fs.writeFileSync(path.join(outDir, 'modelle.json'), JSON.stringify(liste, null, 2) + '\n', 'utf8');
  console.log('geschrieben: docs/bpmn/modelle.json');
}

module.exports = { MODELLE, build, processDoku };
