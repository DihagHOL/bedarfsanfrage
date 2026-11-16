/* Die Prozessmodelle der Bedarfsanfrage gegen das Hausschema des RMS.
   ---------------------------------------------------------------------
   Die Regeln stehen im RMS (`js/prozessschema.js`, R1 bis R10). Dieser Test
   prüft dieselben Regeln an den ERZEUGTEN Dateien, bevor sie ins RMS gehen.
   Liegt das RMS-Repository daneben, läuft zusätzlich dessen eigene Prüfung.

   Aufruf: node tests/test-prozess.mjs                                    */

import { readFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";

const wurzel = join(dirname(fileURLToPath(import.meta.url)), "..");
// Zeilenenden vereinheitlichen: Git wandelt sie unter Windows beim Auschecken um.
const lies = f => readFileSync(join(wurzel, f), "utf8").replace(/\r\n/g, "\n");
const require = createRequire(import.meta.url);

let fehler = 0;
const pruefe = (b, t) => { console.log(`  ${b ? "ok  " : "FEHL"}  ${t}`); if (!b) fehler++; };
const gleich = (a, b, t) => pruefe(JSON.stringify(a) === JSON.stringify(b),
  `${t}${JSON.stringify(a) === JSON.stringify(b) ? "" : `  (erwartet ${JSON.stringify(b)}, war ${JSON.stringify(a)})`}`);

const KNOTEN_RE = /<bpmn:(startEvent|endEvent|userTask|serviceTask|manualTask|task|scriptTask|sendTask|receiveTask|businessRuleTask|exclusiveGateway|parallelGateway|inclusiveGateway|intermediateCatchEvent|intermediateThrowEvent|subProcess|callActivity)\b([^>]*)>/g;
const attr = (roh, name) => (new RegExp(name + '="([^"]*)"').exec(roh || "") || [])[1] || "";

function lesen(xml) {
  const knoten = [];
  KNOTEN_RE.lastIndex = 0;
  let m;
  while ((m = KNOTEN_RE.exec(xml))) knoten.push({ typ: m[1], id: attr(m[2], "id"), name: attr(m[2], "name"), roh: m[2] });
  const fluesse = [...xml.matchAll(/<bpmn:sequenceFlow\b([^>]*)>/g)].map(x => ({
    id: attr(x[1], "id"), von: attr(x[1], "sourceRef"), nach: attr(x[1], "targetRef"), name: attr(x[1], "name") }));
  const bahnen = [...xml.matchAll(/<bpmn:lane\b([^>]*)>([\s\S]*?)<\/bpmn:lane>/g)].map(x => ({
    id: attr(x[1], "id"), name: attr(x[1], "name"),
    knoten: [...x[2].matchAll(/<bpmn:flowNodeRef>([^<]+)<\/bpmn:flowNodeRef>/g)].map(y => y[1]) }));
  const doku = (xml.match(/<bpmn:process\b[^>]*>\s*<bpmn:documentation>([\s\S]*?)<\/bpmn:documentation>/) || [])[1] || "";
  return { knoten, fluesse, bahnen, doku };
}

const liste = JSON.parse(lies("docs/bpmn/modelle.json"));
const ids = JSON.parse(lies("docs/bpmn/rms-ids.json"));

console.log("\nDie Modell-Liste beschreibt, was da ist");
{
  gleich(liste.modelle.map(m => m.key), ["haupt", "erfassen", "pruefen", "freigabe", "autoentwurf"],
    "fünf Modelle: Hauptprozess, drei Unterprozesse, Auto-Entwurf");
  pruefe(liste.rms.url === "https://rms.dihag.de", "sie zeigt auf das RMS, dort sind die Modelle führend");
  pruefe(liste.rms.ablage === "Prozesse/SHB", "Ablage im RMS ist SHB");
  pruefe(/^b!/.test(liste.rms.driveId || ""), "die Bibliothek im RMS ist bekannt (driveId)");
  for (const m of liste.modelle) {
    pruefe(existsSync(join(wurzel, m.datei)), `${m.key}: die Datei ${m.datei} liegt im Repository`);
    pruefe(/\.bpmn$/.test(m.rmsDatei || ""), `${m.key}: hat einen Dateinamen im RMS (${m.rmsDatei})`);
    gleich(m.itemId, (ids.modelle && ids.modelle[m.key]) || "", `${m.key}: Kennung im RMS wie in rms-ids.json`);
  }
  const kennungen = new Set(liste.modelle.map(m => m.kennung));
  gleich(kennungen.size, liste.modelle.length, "jede Prozess-Kennung kommt genau einmal vor");
  for (const m of liste.modelle) for (const u of m.unter)
    pruefe(liste.modelle.some(x => x.key === u), `${m.key} bindet „${u}" ein, und das Modell gibt es`);
  gleich(liste.modelle.filter(m => m.haupt).map(m => m.key), ["haupt", "autoentwurf"],
    "zwei Hauptprozesse: der Auto-Entwurf hat einen eigenen Auslöser (R1)");
}

console.log("\nDie Dateien sind der Stand des Generators");
{
  const gen = require(join(wurzel, "scripts/bpmn/generate-bpmn.js"));
  for (const spec of gen.MODELLE) {
    const datei = liste.modelle.find(m => m.key === spec.key).datei;
    pruefe(gen.build(spec, ids) === lies(datei), `${spec.key}: ${datei} ist frisch erzeugt (sonst: node scripts/bpmn/generate-bpmn.js)`);
  }
}

for (const modell of liste.modelle) {
  const xml = lies(modell.datei);
  const { knoten, fluesse, bahnen, doku } = lesen(xml);
  console.log(`\n${modell.name}  (${modell.datei})`);

  gleich(knoten.filter(k => k.typ === "startEvent").length, 1, "R1 genau ein Auslöser");
  const enden = knoten.filter(k => k.typ === "endEvent");
  pruefe(enden.length >= 1, "R2 mindestens ein Ergebnis");
  pruefe(enden.every(k => k.name.trim()), "R2 jedes Ergebnis ist benannt");
  pruefe(!knoten.some(k => k.typ === "task"), "R3 keine Aufgabe ohne Typ");
  pruefe(!knoten.some(k => ["scriptTask", "sendTask", "receiveTask", "businessRuleTask"].includes(k.typ)),
    "R3 nur die drei Aufgabenarten des Hausschemas");

  const zuordnung = new Map();
  for (const b of bahnen) for (const id of b.knoten) zuordnung.set(id, (zuordnung.get(id) || 0) + 1);
  gleich(knoten.filter(k => !zuordnung.has(k.id)).map(k => k.id), [], "R4 kein Knoten ohne Bahn");
  gleich([...zuordnung].filter(([, n]) => n > 1).map(([id]) => id), [], "R4 kein Knoten in zwei Bahnen");
  pruefe(bahnen.every(b => b.name.trim()), "R5 jede Bahn ist benannt");
  pruefe(!bahnen.some(b => /@/.test(b.name)), "R5 keine Mailadresse als Bahn");

  for (const t of knoten.filter(k => k.typ === "exclusiveGateway")) {
    const raus = fluesse.filter(f => f.von === t.id);
    pruefe(raus.length >= 2, `R6 „${t.name}" hat mindestens zwei Ausgänge`);
    pruefe(raus.every(f => f.name.trim()), `R6 jeder Ausgang von „${t.name}" ist beschriftet`);
    pruefe(/\?$/.test(t.name.trim()), `R6/R8 „${t.name}" ist eine Frage`);
  }

  const hatEin = new Set(fluesse.map(f => f.nach)), hatAus = new Set(fluesse.map(f => f.von));
  gleich(knoten.filter(k => k.typ !== "startEvent" && !hatEin.has(k.id)).map(k => k.id), [], "R7 jeder Knoten hat einen Eingang");
  gleich(knoten.filter(k => k.typ !== "endEvent" && !hatAus.has(k.id)).map(k => k.id), [], "R7 jeder Knoten hat einen Ausgang");

  for (const a of knoten.filter(k => /Task$/.test(k.typ))) {
    const nm = a.name.replace(/\s*\([^)]*\)\s*$/, "").trim();
    pruefe(/\b\w+e?n$/.test(nm), `R8 „${a.name}" endet auf einem Verb`);
  }

  const marker = (doku.match(/\[\[rms:policies=([^\]]+)\]\]/) || [])[1] || "";
  pruefe(marker.split(",").filter(Boolean).length >= 1, "R9 mindestens eine Richtlinie verknüpft");
  gleich(marker.split(",").filter(Boolean), modell.policies, "R9 Liste in der Datei und in modelle.json sind dieselbe");

  const aufrufe = knoten.filter(k => k.typ === "callActivity");
  gleich(aufrufe.length, modell.unter.length, "R10 so viele Aufrufe wie eingebundene Modelle");
  for (const a of aufrufe) {
    const ziel = liste.modelle.find(m => m.kennung === attr(a.roh, "calledElement"));
    pruefe(!!ziel, `R10 „${a.name}" zeigt auf ein Modell, das es gibt`);
    if (ziel && ziel.itemId) {
      const block = (xml.match(new RegExp(`<bpmn:callActivity id="${a.id}"[\\s\\S]*?</bpmn:callActivity>`)) || [""])[0];
      pruefe(block.includes(`[[rms:modell=${ziel.itemId}]]`), `R10 „${a.name}" trägt den RMS-Marker des Ziels`);
    }
  }
  pruefe(!knoten.some(k => k.typ === "subProcess"), "R10 kein ausgeschriebener Unterprozess");

  for (const k of knoten) pruefe(xml.includes(`bpmnElement="${k.id}"`), `Bild: „${k.name || k.id}" hat eine Form`);
  for (const f of fluesse) pruefe(xml.includes(`bpmnElement="${f.id}"`), `Bild: Kante ${f.id} hat einen Verlauf`);
  const formen = [...xml.matchAll(/<bpmndi:BPMNShape\b[^>]*bpmnElement="([^"]+)"/g)].map(x => x[1]);
  gleich(formen.length, new Set(formen).size, "Bild: keine doppelte Form");

  pruefe(doku.includes(`Anwender-Ansicht: https://bedarfsanfrage.dihag-extern.com/?ansicht=prozess&amp;modell=${modell.key}`),
    "die Dokumentation nennt die Anwender-Ansicht in der App");
  for (const n of modell.neben || []) {
    const id = ids.modelle && ids.modelle[n];
    if (id) pruefe(doku.includes(`[[rms:neben=${id}]]`), `Nebenprozess „${n}" ist mit seiner RMS-Kennung zugeordnet`);
  }
  pruefe(!/@[a-z0-9-]+\.[a-z]{2,}/i.test(xml), "keine Mailadressen im Modell (Repository ist öffentlich)");
}

console.log("\nPrüfung des RMS selbst (falls das Repository daneben liegt)");
{
  const pfad = join(wurzel, "..", "richtlinienmanagementsystem", "js", "prozessschema.js");
  if (!existsSync(pfad)) console.log("  --    RMS-Repository nicht gefunden, übersprungen");
  else {
    const { prozessSchemaPruefen } = require(pfad);
    for (const m of liste.modelle) {
      const r = prozessSchemaPruefen(lies(m.datei), { policyIds: m.policies });
      pruefe(!r.fehler.length && !r.hinweise.length,
        `${m.key}: ${r.fehler.length} Fehler, ${r.hinweise.length} Hinweise`
        + (r.fehler.length + r.hinweise.length ? "  " + r.fehler.concat(r.hinweise).map(f => f.regel + " " + f.text).join(" | ") : ""));
    }
  }
}

console.log("\nDie Seite zeigt, was die Liste nennt");
{
  const js = lies("prozess.js"), html = lies("index.html"), app = lies("app.js");
  pruefe(js.includes("docs/bpmn/modelle.json"), "prozess.js liest die Modell-Liste statt eigene Namen zu führen");
  pruefe(js.includes("m.datei") && js.includes(".haupt"), "prozess.js baut Reiter und Downloads aus der Liste");
  pruefe(js.includes("?modell=") && js.includes("?richtlinie="), "prozess.js verlinkt Modell und Regelwerk im RMS");
  pruefe(js.includes("@microsoft.graph.downloadUrl"), "prozess.js lädt aus dem RMS über die signierte Adresse, nicht über /content");
  pruefe(html.indexOf('src="prozess.js"') > html.indexOf('src="app.js"'), "index.html lädt prozess.js nach app.js");
  pruefe(!/function initProzess\b/.test(app) && /function initProzess\b/.test(js), "initProzess gibt es genau einmal, in prozess.js");
  pruefe(!/\blet plState\b/.test(app), "app.js deklariert plState nicht mehr (sonst Doppeldeklaration)");
  pruefe(/START_ANSICHTEN = \[[^\]]*'prozess'/.test(app), "der Direktlink ?ansicht=prozess ist erlaubt");
  for (const f of ["js/vendor/bpmn/bpmn-navigated-viewer.production.min.js", "js/vendor/bpmn/bpmn-js.css",
                   "js/vendor/bpmn/diagram-js.css", "js/vendor/bpmn/LICENSE-bpmn-js.txt"])
    pruefe(existsSync(join(wurzel, f)), `${f} liegt lokal (kein CDN)`);
}

console.log(fehler ? `\n${fehler} Prüfung(en) fehlgeschlagen.\n` : "\nAlle Prüfungen bestanden.\n");
process.exit(fehler ? 1 : 0);
