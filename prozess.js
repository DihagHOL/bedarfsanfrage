'use strict';
/* ── PROZESS-SEITE (Reiter „Prozess") ─────────────────────────────────────────
   Die Bedarfsanfrage als BPMN im Hausschema des RMS, aufgebaut wie die
   Prozess-Seite der E-Rechnung: Prozesslandkarte, Legende, je Hauptprozess ein
   Reiter mit Diagramm (bpmn-js, lokal unter js/vendor/bpmn) und Schritten.

   Führend sind die Modelle im RMS (sites/ISMS, ISMS Dokumente, Prozesse/SHB).
   Die Seite liest sie live von dort, wenn das Konto die Bibliothek lesen darf.
   Sonst zeigt sie die Kopie aus docs/bpmn und sagt das dazu. Welche Modelle es
   gibt, steht in docs/bpmn/modelle.json (erzeugt von scripts/bpmn/generate-bpmn.js),
   die Seite führt keine eigene Liste. Unterprozesse (⊞) sind eigene Modelle;
   ein Klick darauf öffnet sie, über die Kennung aus dem Marker [[rms:modell=…]].

   Nutzt aus app.js: $id, esc, fmtEuro, _authedFetch, API. */

const PZ = {
  liste: null,           // Inhalt von modelle.json
  modelle: {},           // key → Modell (mit cache: { xml, quelle, geaendert, von })
  rms: { url: 'https://rms.dihag.de', driveId: '', ablage: 'Prozesse/SHB' },
  rmsGesperrt: false,    // einmal 401/403/404 aus dem RMS: in dieser Sitzung nicht weiter fragen
  tabs: {},              // tab → { haupt, cur, calls, viewer, needsFit, loading }
  viewerLaden: null,
  start: null,           // { modell } aus dem Deep-Link, wird einmal verbraucht
};
const PZ_MODELL_MARKER = /\[\[rms:modell=([^\]]+)\]\]/;
const PZ_TAB_LABEL = { haupt: '📝 Bedarfsanfrage', autoentwurf: '🔁 Auto-Entwurf' };

/* Deep-Link aus dem RMS („Anwender-Ansicht"): ?ansicht=prozess&modell=<key>.
   app.js gibt das gewünschte Modell hier ab, bevor es den Reiter öffnet. */
function prozessStartModell(key) { PZ.start = key ? { modell: String(key) } : null; }

/* ── bpmn-js erst laden, wenn der Reiter geöffnet wird ── */
function pzViewerLaden() {
  if (window.BpmnJS) return Promise.resolve();
  if (PZ.viewerLaden) return PZ.viewerLaden;
  PZ.viewerLaden = new Promise((resolve, reject) => {
    for (const href of ['js/vendor/bpmn/diagram-js.css', 'js/vendor/bpmn/bpmn-js.css']) {
      const l = document.createElement('link');
      l.rel = 'stylesheet'; l.href = href;
      document.head.appendChild(l);
    }
    const s = document.createElement('script');
    s.src = 'js/vendor/bpmn/bpmn-navigated-viewer.production.min.js';
    s.onload = () => resolve();
    s.onerror = () => { PZ.viewerLaden = null; reject(new Error('Die Diagramm-Anzeige ließ sich nicht laden.')); };
    document.head.appendChild(s);
  });
  return PZ.viewerLaden;
}

/* ── Seite aufbauen ──
   Einmal je Sitzung. Beim nächsten Öffnen des Reiters bleiben Diagramme und
   Probelauf stehen, nur die Ansicht wird neu eingepasst. */
function initProzess() {
  const root = $id('view-prozess');
  if (root.querySelector('.pz-wrap') && PZ.liste) {
    const wunsch = PZ.start && PZ.modelle[PZ.start.modell];
    PZ.start = null;
    if (wunsch) pzZeigeTab(wunsch.tab, wunsch.haupt ? null : wunsch.key);
    else for (const tab of Object.keys(PZ.tabs)) if (PZ.tabs[tab].viewer) requestAnimationFrame(() => pzEinpassen(tab));
    return;
  }
  plState = null;
  PZ.tabs = {};
  root.innerHTML = `
    <div class="pz-wrap">
      <div class="schulung-hero">
        <div class="schulung-badge">Prozess</div>
        <h1>So läuft eine Bedarfsanfrage bei SHB</h1>
        <p>Von der Erfassung im Fachbereich über die Prüfung im Einkauf und die Freigabe in Power Automate
           bis zur Bestellung. Jede Farbe steht für einen Beteiligten, jede Raute für eine Entscheidung.
           Unterprozesse (⊞) öffnen Sie per Klick. Gepflegt werden die Modelle im RMS, diese Seite zeigt
           ihren aktuellen Stand.</p>
        <a class="hero-link" id="pz-rms-link" href="https://rms.dihag.de/" target="_blank" rel="noopener">Prozess im RMS öffnen ↗</a>
      </div>

      <div class="pz-card">
        <h3>🗺️ Prozesslandkarte</h3>
        <p class="pz-lead">Ein Klick auf eine Zeile öffnet das passende Diagramm weiter unten.</p>
        <div class="pz-map">
          <button type="button" class="pz-map-row" data-go="haupt">
            <span class="pz-map-ext"><b>👷</b>Fachbereich</span>
            <span class="pz-map-arrow">➜</span>
            <span class="pz-map-core pz-core-haupt"><strong>📝 Bedarfsanfrage</strong>
              <span>Erfassen in der App, Prüfung und Entscheidung im Einkauf, Freigabe in Power Automate, Bestellung im ERP</span></span>
            <span class="pz-map-arrow">➜</span>
            <span class="pz-map-ext"><b>🚚</b>Lieferant</span>
          </button>
          <button type="button" class="pz-map-row" data-go="autoentwurf">
            <span class="pz-map-ext"><b>📅</b>Datum zur Wiedervorlage</span>
            <span class="pz-map-arrow">➜</span>
            <span class="pz-map-core pz-core-auto"><strong>🔁 Auto-Entwurf</strong>
              <span>Ein täglicher Lauf legt fällige Wiedervorlagen als Entwurf an</span></span>
            <span class="pz-map-arrow">➜</span>
            <span class="pz-map-ext"><b>📝</b>Prüfung im Einkauf</span>
          </button>
          <div class="pz-map-base">🗄️ <strong>Überall dabei:</strong> die SharePoint-Liste Bedarfsanfrage auf gruppe_shb.
            Was jemand sieht und tun darf, regelt die Zugriffssteuerung: die Abteilung für Gemeinsame Anfragen,
            die Warengruppe für den Einkauf.</div>
        </div>
      </div>

      <div class="pz-card pz-legend">
        <div>
          <h4>Wer macht was (Farbe der Bahn)</h4>
          <div class="pz-chips">
            <span class="pz-chip a-mensch">Menschen: Antragsteller, Einkauf, Genehmiger</span>
            <span class="pz-chip a-app">Bedarfsanfrage-App</span>
            <span class="pz-chip a-auto">Power Automate und GitHub Actions</span>
            <span class="pz-chip a-extern">Extern: Lieferant</span>
          </div>
          <h4 style="margin-top:14px">Was ist das Ergebnis</h4>
          <div class="pz-chips">
            <span class="pz-chip t-ok">Start und gutes Ende</span>
            <span class="pz-chip t-gate">Entscheidung</span>
            <span class="pz-chip t-err">Abgelehnt</span>
            <span class="pz-chip t-plan">Zurückgezogen oder nichts zu tun</span>
          </div>
        </div>
        <div>
          <h4>Symbole</h4>
          <div class="pz-syms">${pzSymbole()}</div>
        </div>
      </div>

      <div id="pz-prozesse">
        <div class="pz-tabs" id="pz-tabs" role="tablist"></div>
        <div id="pz-panels"><div class="pz-panel"><p class="pz-lead">Modelle werden geladen…</p></div></div>
      </div>

      <div class="pz-card">
        <h3>▶ Probelauf</h3>
        <p class="pz-lead">Spielen Sie den Freigabeweg mit eigenen Werten durch. Je nach Bestellvolumen,
           Beschaffung und Warengruppe entfallen einzelne Stufen, genau wie im Flow.</p>
        <div class="pl-controls">
          <div class="pl-field">
            <label for="pl-preis">Bestellvolumen</label>
            <div class="pl-euro"><input type="number" id="pl-preis" min="0" step="100" value="3500"/><span>€</span></div>
          </div>
          <div class="pl-field">
            <label for="pl-art">Beschaffung</label>
            <select id="pl-art">
              <option value="material">Direktes Material, sonstige Warengruppe</option>
              <option value="indirekt">Indirektes Material / Dienstleistung</option>
              <option value="wg">Warengruppe 2000, 4333 bis 4338 oder 4340</option>
            </select>
          </div>
          <div class="pl-field pl-field-check">
            <label><input type="checkbox" id="pl-lead"/> Lead-Buyer-Abschluss</label>
          </div>
          <button class="btn btn-primary" onclick="startProbelauf()">▶ Probelauf starten</button>
        </div>
        <div id="pl-stage"></div>
      </div>

      <div class="pz-card">
        <h3>⬇️ Prozessmodelle herunterladen</h3>
        <p class="pz-lead">BPMN 2.0 mit Farben, gepflegt im <a href="https://rms.dihag.de/" target="_blank" rel="noopener">RMS</a>
           (Reiter Prozesse, Ablage SHB). Hier liegt die Kopie aus dem Repository, den angezeigten Stand speichert
           der Knopf „⬇ BPMN" am Diagramm. Öffnen lassen sich die Dateien im Camunda Modeler, auf bpmn.io oder im RMS.</p>
        <div class="pz-dl" id="pz-dl"></div>
        <p class="pz-note" id="pz-regelwerke"></p>
      </div>
    </div>`;

  $id('view-prozess').querySelectorAll('[data-go]').forEach(b => b.addEventListener('click', () => {
    pzZeigeTab(b.dataset.go);
    $id('pz-prozesse')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }));
  pzModelleLaden().catch(e => {
    const p = $id('pz-panels');
    if (p) p.innerHTML = `<div class="pz-panel"><div class="pz-error">Die Modell-Liste ließ sich nicht laden: ${esc(e.message)}</div></div>`;
  });
}

function pzSymbole() {
  const s = (svg, text) => `<div class="pz-sym"><svg viewBox="0 0 34 26" aria-hidden="true">${svg}</svg>${text}</div>`;
  return [
    s('<circle cx="17" cy="13" r="10" fill="#DDF3E4" stroke="#1E7B3A" stroke-width="1.5"/>', 'Start'),
    s('<circle cx="17" cy="13" r="10" fill="#DDF3E4" stroke="#1E7B3A" stroke-width="4"/>', 'Ende'),
    s('<circle cx="17" cy="13" r="10" fill="#DDF3E4" stroke="#1E7B3A" stroke-width="1.5"/><circle cx="17" cy="13" r="6.5" fill="#fff" stroke="#1E7B3A" stroke-width="1"/><path d="M17 9 V13 L20 15" fill="none" stroke="#1E7B3A" stroke-width="1.2"/>', 'Start zu einer Uhrzeit'),
    s('<rect x="3" y="3" width="28" height="20" rx="4" fill="#DCE5F2" stroke="#1A2644" stroke-width="1.5"/><circle cx="9" cy="9" r="3" fill="none" stroke="#1A2644" stroke-width="1.2"/>', 'Automatischer Schritt (Zahnrad)'),
    s('<rect x="3" y="3" width="28" height="20" rx="4" fill="#FFE3C8" stroke="#C2410C" stroke-width="1.5"/><circle cx="9" cy="8" r="2.2" fill="none" stroke="#C2410C" stroke-width="1.2"/><path d="M5.5 14 Q9 10 12.5 14" fill="none" stroke="#C2410C" stroke-width="1.2"/>', 'Schritt einer Person'),
    s('<path d="M17 2 L28 13 L17 24 L6 13 Z" fill="#FFF8DB" stroke="#8A6100" stroke-width="1.5"/><path d="M13 9 L21 17 M21 9 L13 17" stroke="#8A6100" stroke-width="1.8"/>', 'Entscheidung: genau ein Weg'),
    s('<rect x="3" y="3" width="28" height="20" rx="4" fill="#FFE3C8" stroke="#C2410C" stroke-width="3"/><rect x="13" y="14" width="8" height="7" fill="#fff" stroke="#C2410C" stroke-width="1"/><path d="M17 15.5 V19.5 M15 17.5 H19" stroke="#C2410C" stroke-width="1.1"/>', 'Unterprozess (⊞): Klick öffnet ihn'),
    s('<rect x="2" y="3" width="30" height="20" rx="4" fill="none" stroke="#17509E" stroke-width="1.4" stroke-dasharray="4 3"/>', 'Gruppe, z. B. „je fällige Anfrage"'),
    s('<path d="M3 13 H28" stroke="#374151" stroke-width="1.3" stroke-dasharray="4 3"/><circle cx="4" cy="13" r="2.5" fill="#fff" stroke="#374151"/><path d="M26 9 L31 13 L26 17 Z" fill="#fff" stroke="#374151"/>', 'Nachricht an den Lieferanten'),
  ].join('');
}

/* ── Modell-Liste lesen, Reiter und Downloads daraus bauen ── */
async function pzModelleLaden() {
  if (!PZ.liste) {
    const r = await fetch('docs/bpmn/modelle.json', { cache: 'no-cache' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    PZ.liste = await r.json();
    PZ.rms = Object.assign(PZ.rms, PZ.liste.rms || {});
    for (const m of PZ.liste.modelle) PZ.modelle[m.key] = m;
  }
  const haupt = PZ.liste.modelle.filter(m => m.haupt);
  for (const m of haupt) PZ.tabs[m.tab] = { haupt: m.key, cur: m.key, calls: [] };

  const rmsUrl = String(PZ.rms.url || 'https://rms.dihag.de').replace(/\/+$/, '');
  const hauptModell = PZ.modelle.haupt;
  const heroLink = $id('pz-rms-link');
  if (heroLink && hauptModell && hauptModell.itemId) heroLink.href = `${rmsUrl}/?modell=${encodeURIComponent(hauptModell.itemId)}`;

  $id('pz-tabs').innerHTML = haupt.map(m =>
    `<button type="button" class="pz-tab" data-tab="${esc(m.tab)}" role="tab">${esc(PZ_TAB_LABEL[m.tab] || m.name)}</button>`).join('');
  $id('pz-panels').innerHTML = haupt.map(m => pzPanelHtml(m)).join('');
  $id('pz-tabs').querySelectorAll('.pz-tab').forEach(b => b.addEventListener('click', () => pzZeigeTab(b.dataset.tab)));
  for (const m of haupt) pzBedienleiste(m.tab);

  // Downloads und Regelwerke aus derselben Liste, keine zweite Fassung der Namen
  $id('pz-dl').innerHTML = PZ.liste.modelle.map(m => {
    const zeichen = m.haupt ? (m.tab === 'autoentwurf' ? '🔁' : '📝') : '⊞';
    const rms = m.itemId ? ` · <a href="${rmsUrl}/?modell=${encodeURIComponent(m.itemId)}" target="_blank" rel="noopener">im RMS ↗</a>` : '';
    return `<div class="pz-dl-row"><a href="${esc(m.datei)}" download>${zeichen} ${esc(m.name)}</a><small>${esc(m.datei.split('/').pop())}${rms}</small></div>`;
  }).join('');
  const ids = [...new Set(PZ.liste.modelle.flatMap(m => m.policies || []))];
  const namen = PZ.liste.richtlinien || {};
  $id('pz-regelwerke').innerHTML = ids.length
    ? 'Verknüpfte Regelwerke im RMS: ' + ids.map(id =>
        `<a href="${rmsUrl}/?richtlinie=${encodeURIComponent(id)}" target="_blank" rel="noopener">${esc(namen[id] || 'Regelwerk ' + id)}</a>`).join(' · ')
    : '';

  // Start: Deep-Link, sonst der erste Hauptprozess
  let tab = haupt[0] && haupt[0].tab, modell = null;
  const wunsch = PZ.start && PZ.modelle[PZ.start.modell];
  PZ.start = null;
  if (wunsch) { tab = wunsch.tab; if (!wunsch.haupt) modell = wunsch.key; }
  if (tab) await pzZeigeTab(tab, modell);
}

function pzPanelHtml(m) {
  const t = esc(m.tab);
  const texte = PZ_PANEL[m.tab] || { titel: esc(m.name), lead: '', danach: '' };
  return `
    <section class="pz-panel" id="pz-panel-${t}" hidden>
      <div class="pz-head">
        <div>
          <h2>${texte.titel}</h2>
          <p class="pz-lead">${texte.lead}</p>
        </div>
        <div class="pz-actions" data-actions="${t}"></div>
      </div>
      <div class="pz-views" id="pz-views-${t}"></div>
      <div class="pz-box" id="pz-box-${t}">
        <div class="pz-canvas" id="pz-canvas-${t}"></div>
        <div class="pz-hint">Ziehen verschiebt · Strg + Mausrad zoomt · Klick auf ⊞ öffnet den Unterprozess</div>
      </div>
      ${texte.danach}
    </section>`;
}

/* Feste Texte je Reiter. Bewusst hier und nicht im Modell: Sie erklären die
   Bedienung der App, das Modell beschreibt den Ablauf. */
const pzChip = (cls, text) => `<span class="pz-chip ${cls}">${text}</span>`;
const PZ_PANEL = {
  haupt: {
    titel: '📝 Bedarfsanfrage',
    lead: 'Der Antragsteller erfasst den Bedarf in der App. Der zuständige Einkauf prüft und entscheidet in der App, '
        + 'danach holt Power Automate je nach Art und Betrag die Freigaben von Controlling, Werkleitung und '
        + 'strategischem Einkauf ein. Ist alles freigegeben, bestellt der Einkauf und trägt die Bestellnummer ein.',
    danach: `
      <div class="pz-steps-wrap">
        <div>
          <h4 class="pz-box-title">Schritt für Schritt</h4>
          <ol class="pz-steps">
            <li>${pzChip('a-mensch', 'Antragsteller')}erfasst die Anfrage im Assistenten, als Einzelanfrage oder als Sammelanfrage mit Positionen. Favoriten und Wiedervorlagen füllen ihn vor. Die App rechnet das Bestellvolumen aus und zeigt den Genehmigungsweg mit der Zahl der Angebote.</li>
            <li>${pzChip('a-app', 'App')}speichert die Anfrage mit dem Status Eingereicht in der SharePoint-Liste. Eine Mail an den Einkauf geht nicht raus.</li>
            <li>${pzChip('a-mensch', 'Antragsteller')}kann die Anfrage in „Meine Anfragen" zurückziehen, solange der Einkauf nicht entschieden hat. Sie steht dann auf Zurückgezogen und ist aus der Bearbeitung.</li>
            <li>${pzChip('a-mensch', 'Einkauf')}sieht neue Anfragen im Dashboard und unter Gemeinsame Anfragen. Was seit über 48 Stunden offen ist, markiert die App rot. Öffnet der zuständige Einkäufer die Anfrage, setzt die App den Status auf In Prüfung (Einkauf).</li>
            <li>${pzChip('a-mensch', 'Einkauf')}reserviert die Anfrage bei Bedarf für sich, holt fehlende Angebote ein, hängt sie an und genehmigt oder lehnt mit Grund ab. Bei einer Ablehnung geht keine Mail raus, der Antragsteller sieht sie in „Meine Anfragen".</li>
            <li>${pzChip('a-auto', 'Power Automate')}startet mit dem Status Genehmigt (Einkauf) den Freigabelauf und holt je nach Art und Betrag die Freigaben ein. Ob freigegeben oder abgelehnt, der Antragsteller bekommt vom Flow eine Mail.</li>
            <li>${pzChip('a-mensch', 'Einkauf')}bestellt die freigegebene Anfrage im ERP und trägt Bestellnummer, Liefertermin und tatsächlichen Preis ein. Steht die Bestellnummer schon an einer anderen Anfrage, warnt die App. Danach steht die Anfrage auf Bestellt.</li>
          </ol>
        </div>
        <div class="pz-side">
          <h4 class="pz-box-title">Wer gibt frei?</h4>
          <table class="pz-rules">
            <tr><td>${pzChip('t-gate', 'Einkauf')}</td><td>Immer, in der App. Zuständig ist, wem die Warengruppe zugeordnet ist. Ohne Zuordnung sind es alle Einkäufer.</td></tr>
            <tr><td>${pzChip('a-mensch', 'Controlling')}</td><td>Bei indirektem Material und den Warengruppen 2000, 4333 bis 4338 und 4340.</td></tr>
            <tr><td>${pzChip('a-mensch', 'Werkleitung')}</td><td>Ab 1.500 € Bestellvolumen.</td></tr>
            <tr><td>${pzChip('a-mensch', 'Strategischer Einkauf')}</td><td>Ab 10.000 €. Entfällt beim Lead-Buyer-Abschluss.</td></tr>
          </table>
          <p class="pz-note">Controlling, Werkleitung und strategischer Einkauf genehmigen in Teams oder per Mail, die erste Antwort zählt. Die Wertgrenzen stehen im Flow „Bedarfsanfragen SHB".</p>
        </div>
      </div>

      <div class="pz-steps-wrap">
        <div>
          <h4 class="pz-box-title">⊞ Anfrage erfassen und einreichen</h4>
          <ol class="pz-steps">
            <li>${pzChip('a-mensch', 'Antragsteller')}wählt Einzel- oder Sammelanfrage. Zur Artikelnummer schlägt die App Bezeichnung und Warengruppe aus dem Artikelstamm vor.</li>
            <li>${pzChip('a-mensch', 'Antragsteller')}gibt Beschaffungsart, Lieferanten, Kostenstelle und Abteilung an. Die Abteilung bestimmt, unter welcher Gruppe die Anfrage in Gemeinsame Anfragen erscheint.</li>
            <li>${pzChip('a-app', 'App')}summiert das Bestellvolumen aus Menge und Einzelpreis und zeigt den Genehmigungsweg.</li>
            <li>${pzChip('a-mensch', 'Antragsteller')}hängt vorhandene Angebote an, prüft die Zusammenfassung und reicht ein.</li>
          </ol>
        </div>
        <div class="pz-side">
          <h4 class="pz-box-title">Angebote nach Regelwerk</h4>
          <table class="pz-rules">
            <tr><td>${pzChip('t-ok', 'bis 500 €')}</td><td>Kein Angebot nötig.</td></tr>
            <tr><td>${pzChip('t-warn', 'über 500 €')}</td><td>2 Angebote.</td></tr>
            <tr><td>${pzChip('t-warn', 'über 2.500 €')}</td><td>3 Angebote.</td></tr>
            <tr><td>${pzChip('t-warn', 'über 10.000 €')}</td><td>4 Angebote.</td></tr>
            <tr><td>${pzChip('t-err', 'über 50.000 €')}</td><td>5 Angebote, europäische Ausschreibung.</td></tr>
            <tr><td>${pzChip('t-ok', 'Lead-Buyer')}</td><td>Die Angebotspflicht entfällt, ebenso die Freigabe des strategischen Einkaufs.</td></tr>
          </table>
          <p class="pz-note">Die App zeigt die geforderte Zahl beim Erfassen an, erzwingt sie aber nicht. Fehlende Angebote holt der Einkauf ein. Angebote lassen sich in jedem Status nachreichen: Anfrage öffnen und unter „Anhänge" ablegen.</p>
        </div>
      </div>

      <div class="pz-steps-wrap">
        <div>
          <h4 class="pz-box-title">⊞ Freigabelauf (Power Automate)</h4>
          <ol class="pz-steps">
            <li>${pzChip('a-auto', 'Power Automate')}startet, sobald eine Anfrage auf Genehmigt (Einkauf) steht.</li>
            <li>${pzChip('a-auto', 'Power Automate')}setzt bei indirektem Material und bei den Warengruppen 2000, 4333 bis 4338 und 4340 den Status In Prüfung (Controlling) und schickt dem Controlling die Genehmigung.</li>
            <li>${pzChip('a-auto', 'Power Automate')}setzt ab 1.500 € In Prüfung (Werkleitung), ab 10.000 € ohne Lead-Buyer-Abschluss danach In Prüfung (strategischer Einkauf), und schickt jeweils die Genehmigung.</li>
            <li>${pzChip('a-mensch', 'Genehmiger')}antworten in Teams oder per Mail. Lehnt eine Stufe ab, setzt der Flow Abgelehnt und die folgenden Stufen entfallen.</li>
            <li>${pzChip('a-auto', 'Power Automate')}setzt am Ende Freigegeben und schreibt dem Antragsteller eine Mail, bei einer Ablehnung ebenso.</li>
          </ol>
        </div>
        <div class="pz-side">
          <h4 class="pz-box-title">Gut zu wissen</h4>
          <table class="pz-rules">
            <tr><td>${pzChip('t-ok', 'Verlauf')}</td><td>Im Verlauf einer Anfrage stehen nur die Stufen, die für sie gelten.</td></tr>
            <tr><td>${pzChip('t-gate', 'Erste Antwort')}</td><td>Sind für eine Stufe mehrere Genehmiger hinterlegt, zählt die erste Antwort.</td></tr>
            <tr><td>${pzChip('t-err', 'Abgelehnt')}</td><td>Die Anfrage ist beendet. Für den Bedarf braucht es eine neue Anfrage.</td></tr>
            <tr><td>${pzChip('a-auto', 'Empfänger')}</td><td>Je Stufe fest im Flow hinterlegt. Änderungen macht die IT im Flow.</td></tr>
          </table>
        </div>
      </div>`,
  },
  autoentwurf: {
    titel: '🔁 Auto-Entwurf',
    lead: 'Für wiederkehrende Bedarfe. Wer beim Erfassen ein Datum zur automatischen Bedarfsanfrage setzt, '
        + 'bekommt an diesem Tag eine Kopie der Anfrage als Entwurf. Öffnet der Einkauf den Entwurf, läuft er '
        + 'wie eine eingereichte Anfrage weiter.',
    danach: `
      <div class="pz-steps-wrap">
        <div>
          <h4 class="pz-box-title">Schritt für Schritt</h4>
          <ol class="pz-steps">
            <li>${pzChip('a-mensch', 'Antragsteller')}setzt beim Erfassen ein Datum zur automatischen Bedarfsanfrage.</li>
            <li>${pzChip('a-auto', 'GitHub Actions')}sucht jeden Morgen gegen 6 Uhr (im Sommer 7 Uhr) die Anfragen, deren Datum heute oder früher ist.</li>
            <li>${pzChip('a-auto', 'GitHub Actions')}legt je fällige Anfrage eine Kopie mit dem Status Automatisch erstellter Entwurf an. Status, Entscheider und Bestelldaten kopiert er nicht.</li>
            <li>${pzChip('a-auto', 'GitHub Actions')}leert das Datum an der Vorlage, damit kein zweiter Entwurf entsteht.</li>
            <li>${pzChip('a-mensch', 'Einkauf')}findet den Entwurf im Dashboard. Öffnet er ihn, setzt die App den Status auf In Prüfung (Einkauf). Ab hier gilt der Hauptprozess.</li>
          </ol>
        </div>
        <div class="pz-side">
          <h4 class="pz-box-title">Gut zu wissen</h4>
          <table class="pz-rules">
            <tr><td>${pzChip('t-plan', 'Titel')}</td><td>Entwürfe beginnen mit „Auto-Entwurf:".</td></tr>
            <tr><td>${pzChip('a-auto', 'Zeitplan')}</td><td>GitHub startet geplante Läufe oft einige Minuten später. Hier ist das unkritisch.</td></tr>
            <tr><td>${pzChip('t-ok', 'Kein Doppel')}</td><td>Entwürfe werden nicht erneut kopiert, und das geleerte Datum verhindert einen zweiten Lauf.</td></tr>
          </table>
        </div>
      </div>`,
  },
};

/* ── Reiter, Bedienleiste, Ansicht-Knöpfe ── */
async function pzZeigeTab(tab, modell) {
  if (!PZ.tabs[tab]) return;
  document.querySelectorAll('#pz-tabs .pz-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('#pz-panels .pz-panel').forEach(p => { p.hidden = p.id !== 'pz-panel-' + tab; });
  await pzDiagrammLaden(tab);
  if (modell && PZ.tabs[tab].viewer) await pzZeigeModell(tab, modell);
}

function pzBedienleiste(tab) {
  const cfg = PZ.tabs[tab];
  const views = $id('pz-views-' + tab);
  views.addEventListener('click', e => {
    const b = e.target.closest('button[data-modell]');
    if (b) pzZeigeModell(tab, b.dataset.modell);
  });
  const acts = document.querySelector(`#view-prozess [data-actions="${tab}"]`);
  acts.innerHTML =
      `<button type="button" data-act="in" title="Vergrößern">＋</button>`
    + `<button type="button" data-act="out" title="Verkleinern">－</button>`
    + `<button type="button" data-act="fit" title="Ganzes Diagramm zeigen">⤢ Einpassen</button>`
    + `<button type="button" data-act="full" title="Vollbild">⛶ Vollbild</button>`
    + `<button type="button" data-act="svg" title="Aktuelle Ansicht als Bild speichern">🖼 SVG</button>`
    + `<button type="button" data-act="bpmn" class="dl" title="Angezeigtes Modell als BPMN-Datei speichern">⬇ BPMN</button>`;
  acts.addEventListener('click', e => {
    const b = e.target.closest('button[data-act]');
    if (!b || !cfg.viewer) return;
    const canvas = cfg.viewer.get('canvas');
    const act = b.dataset.act;
    if (act === 'in') canvas.zoom(canvas.zoom() * 1.25);
    if (act === 'out') canvas.zoom(canvas.zoom() / 1.25);
    if (act === 'fit') pzEinpassen(tab);
    if (act === 'full') {
      const box = $id('pz-box-' + tab);
      if (document.fullscreenElement) document.exitFullscreen(); else if (box.requestFullscreen) box.requestFullscreen();
    }
    if (act === 'svg') pzSpeichernSvg(tab);
    if (act === 'bpmn') pzSpeichernBpmn(tab);
  });
}

function pzViews(tab) {
  const cfg = PZ.tabs[tab];
  const views = $id('pz-views-' + tab);
  if (!views) return;
  views.innerHTML = '<span>Ansicht:</span>'
    + `<button type="button" data-modell="${esc(cfg.haupt)}">Gesamtprozess</button>`
    + cfg.calls.map(c => `<button type="button" data-modell="${esc(c.key)}">⊞ ${esc(c.name)}</button>`).join('')
    + `<span class="pz-src" id="pz-src-${esc(tab)}"></span>`;
  views.querySelectorAll('button[data-modell]').forEach(b => b.classList.toggle('active', b.dataset.modell === cfg.cur));
  const m = PZ.modelle[cfg.cur], c = m && m.cache;
  const src = $id('pz-src-' + tab);
  if (!c || !src) return;
  const rmsUrl = String(PZ.rms.url || 'https://rms.dihag.de').replace(/\/+$/, '');
  const link = m.itemId
    ? ` · <a href="${rmsUrl}/?modell=${encodeURIComponent(m.itemId)}" target="_blank" rel="noopener">im RMS öffnen ↗</a>` : '';
  if (c.quelle === 'rms') {
    const d = c.geaendert ? new Date(c.geaendert).toLocaleDateString('de-DE') : '';
    src.innerHTML = `<b class="pz-src-live">● live aus dem RMS</b>${d ? ' · geändert ' + esc(d) : ''}${c.von ? ' von ' + esc(c.von) : ''}${link}`;
  } else {
    src.innerHTML = `<b class="pz-src-repo">● Stand aus dem Repository</b> (kein Zugriff auf das RMS)${link}`;
  }
}

/* ── Modelle laden: RMS zuerst, Repository als Rückfall ── */
async function pzHoleModell(key) {
  const m = PZ.modelle[key];
  if (!m) throw new Error('Modell unbekannt: ' + key);
  if (m.cache) return m.cache;
  if (PZ.rms.driveId && m.itemId && !PZ.rmsGesperrt && typeof _authedFetch === 'function') {
    try {
      // Erst die Metadaten, dann der Inhalt über die vorab signierte Adresse und ohne
      // Anmeldetoken. /content leitet auf einen anderen Host um, das scheitert im Browser.
      const r = await _authedFetch(`${API}/drives/${PZ.rms.driveId}/items/${m.itemId}`);
      if (r.status === 401 || r.status === 403 || r.status === 404) PZ.rmsGesperrt = true;
      if (r.ok) {
        const meta = await r.json();
        const url = meta['@microsoft.graph.downloadUrl'];
        if (url) {
          const x = await fetch(url);
          if (!x.ok) throw new Error('HTTP ' + x.status);
          m.cache = { xml: await x.text(), quelle: 'rms', geaendert: meta.lastModifiedDateTime,
            von: (meta.lastModifiedBy && meta.lastModifiedBy.user && meta.lastModifiedBy.user.displayName) || '' };
          return m.cache;
        }
      }
    } catch (e) {
      console.warn('[Prozess] RMS nicht erreichbar, nehme die Kopie aus dem Repository:', m.name, e.message);
    }
  }
  if (!m.datei) throw new Error('Kein Zugriff auf das RMS, und im Repository gibt es keine Kopie.');
  const res = await fetch(m.datei, { cache: 'no-cache' });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  m.cache = { xml: await res.text(), quelle: 'repo' };
  return m.cache;
}

/** Ziel einer Aufrufaktivität: Datei-Kennung aus dem RMS-Marker, sonst die Prozess-Kennung. */
function pzZielVonAufruf(bo) {
  const doku = (bo.documentation || []).map(d => d.text || '').join('\n');
  const mk = doku.match(PZ_MODELL_MARKER);
  const itemId = mk ? mk[1].trim() : '';
  let m = Object.values(PZ.modelle).find(x => (itemId && x.itemId === itemId)
    || (bo.calledElement && x.kennung === bo.calledElement));
  if (!m && itemId) {
    // Im RMS neu eingebunden und noch nicht in der Modell-Liste: direkt von dort lesen.
    m = PZ.modelle['rms:' + itemId] = { key: 'rms:' + itemId, name: bo.name || 'Unterprozess', itemId, datei: '', kennung: bo.calledElement || '' };
  }
  return m ? m.key : null;
}

function pzWennSichtbar(el) {
  return new Promise(resolve => {
    if (el.clientWidth && el.clientHeight) return resolve();
    if (!window.ResizeObserver) return setTimeout(resolve, 300);
    const ro = new ResizeObserver(() => { if (el.clientWidth && el.clientHeight) { ro.disconnect(); resolve(); } });
    ro.observe(el);
  });
}

async function pzDiagrammLaden(tab) {
  const cfg = PZ.tabs[tab];
  if (!cfg) return;
  if (cfg.viewer) { requestAnimationFrame(() => pzEinpassen(tab)); return; }
  if (cfg.loading) return cfg.loading;
  const canvasEl = $id('pz-canvas-' + tab);
  cfg.loading = (async () => {
    try {
      await pzViewerLaden();
    } catch (e) {
      canvasEl.innerHTML = `<div class="pz-error">${esc(e.message)}</div>`;
      return;
    }
    await pzWennSichtbar(canvasEl);   // erst mit echter Größe starten, sonst zoomt der Viewer ins Leere
    if (!canvasEl.isConnected) return; // Reiter inzwischen verlassen
    const viewer = new BpmnJS({ container: canvasEl });
    cfg.viewer = viewer;
    cfg.needsFit = true;
    viewer.get('eventBus').on('element.click', e => {
      const bo = e.element && e.element.businessObject;
      if (!bo || bo.$type !== 'bpmn:CallActivity') return;
      const ziel = pzZielVonAufruf(bo);
      if (ziel) pzZeigeModell(tab, ziel);
    });
    if (window.ResizeObserver) {
      new ResizeObserver(() => { viewer.get('canvas').resized(); if (cfg.needsFit) pzEinpassen(tab); }).observe(canvasEl);
    }
    await pzZeigeModell(tab, cfg.haupt);
  })();
  await cfg.loading;
  cfg.loading = null;
}

async function pzZeigeModell(tab, modellKey) {
  const cfg = PZ.tabs[tab];
  if (!cfg || !cfg.viewer) return;
  const canvasEl = $id('pz-canvas-' + tab);
  canvasEl.querySelectorAll('.pz-error').forEach(x => x.remove());
  try {
    const mod = await pzHoleModell(modellKey);
    const { warnings } = await cfg.viewer.importXML(mod.xml);
    if (warnings && warnings.length) console.warn('[Prozess] BPMN-Hinweise (' + modellKey + '):', warnings);
  } catch (e) {
    canvasEl.insertAdjacentHTML('beforeend', `<div class="pz-error">Diagramm konnte nicht geladen werden: ${esc(e.message)}</div>`);
    return;
  }
  cfg.cur = modellKey;
  const reg = cfg.viewer.get('elementRegistry'), canvas = cfg.viewer.get('canvas');
  const calls = [];
  reg.filter(el => el.businessObject && el.businessObject.$type === 'bpmn:CallActivity').forEach(el => {
    canvas.addMarker(el.id, 'pz-call');
    const ziel = pzZielVonAufruf(el.businessObject);
    if (ziel && !calls.some(c => c.key === ziel)) calls.push({ key: ziel, name: el.businessObject.name || PZ.modelle[ziel].name });
  });
  if (modellKey === cfg.haupt) cfg.calls = calls;
  pzViews(tab);
  cfg.needsFit = true;
  pzEinpassen(tab);
}

// Einpassen nur mit echter Größe (versteckter Reiter hat 0 px), sonst später nachholen.
function pzEinpassen(tab) {
  const cfg = PZ.tabs[tab];
  if (!cfg || !cfg.viewer) return;
  const el = $id('pz-canvas-' + tab);
  if (!el || !el.clientWidth || !el.clientHeight) { cfg.needsFit = true; return; }
  const canvas = cfg.viewer.get('canvas');
  canvas.resized();
  try { canvas.zoom('fit-viewport', 'auto'); cfg.needsFit = false; }
  catch (e) { cfg.needsFit = true; }
}

function pzDateiName(tab, endung) {
  const m = PZ.modelle[PZ.tabs[tab].cur];
  const basis = m && m.datei ? m.datei.split('/').pop().replace(/\.bpmn$/, '')
    : String((m && m.name) || 'prozess').toLowerCase().replace(/[^a-z0-9äöüß]+/g, '-');
  return basis + endung;
}

function pzSpeichern(inhalt, typ, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([inhalt], { type: typ }));
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function pzSpeichernSvg(tab) {
  const { svg } = await PZ.tabs[tab].viewer.saveSVG();
  pzSpeichern(svg, 'image/svg+xml', pzDateiName(tab, '.svg'));
}

function pzSpeichernBpmn(tab) {
  const m = PZ.modelle[PZ.tabs[tab].cur];
  if (m && m.cache) pzSpeichern(m.cache.xml, 'application/xml', pzDateiName(tab, '.bpmn'));
}

document.addEventListener('fullscreenchange', () => {
  for (const tab of Object.keys(PZ.tabs)) if (PZ.tabs[tab].viewer) setTimeout(() => pzEinpassen(tab), 60);
});

/* ── Probelauf: den Freigabeweg mit eigenen Werten durchspielen ──────────────
   Dieselben Regeln wie im Flow „Bedarfsanfragen SHB" und in der Status-Timeline:
   Controlling bei indirektem Material und den Warengruppen 2000, 4333 bis 4338, 4340,
   Werkleitung ab 1.500 €, strategischer Einkauf ab 10.000 € und ohne Lead-Buyer-Abschluss. */
let plState = null;
const PL_ART = { material: 'direktes Material', indirekt: 'indirektes Material', wg: 'Controlling-Warengruppe' };

function roleBadge(role) {
  const map = {
    'Antragsteller': 'rb-fach', 'App': 'rb-system', 'Einkauf': 'rb-einkauf', 'Power Automate': 'rb-system',
    'Controlling': 'rb-controlling', 'Werkleitung': 'rb-werk', 'Strategischer Einkauf': 'rb-strat',
  };
  return `<span class="role-badge ${map[role] || 'rb-system'}">${esc(role)}</span>`;
}

function plBuildStages(preis, art, lead) {
  const controlling = art !== 'material';
  return [
    { role: 'Antragsteller', name: 'Bedarfsanfrage erfassen', kind: 'task', applies: true,
      desc: 'Der Antragsteller beschreibt den Bedarf im Assistenten und reicht die Anfrage ein.' },
    { role: 'App', name: 'Anfrage speichern (Eingereicht)', kind: 'auto', applies: true,
      desc: 'Die App speichert die Anfrage in SharePoint. Der Einkauf sieht sie im Dashboard und unter Gemeinsame Anfragen.' },
    { role: 'Einkauf', name: 'In Prüfung (Einkauf)', kind: 'decision', applies: true,
      desc: 'Der zuständige Einkauf öffnet die Anfrage, holt fehlende Angebote ein und entscheidet in der App.' },
    { role: 'Power Automate', name: 'Genehmigt (Einkauf): Freigabelauf startet', kind: 'auto', applies: true,
      desc: 'Mit der Genehmigung des Einkaufs startet der Flow „Bedarfsanfragen SHB".' },
    { role: 'Controlling', name: 'Freigabe Controlling', kind: 'decision', applies: controlling,
      cond: 'bei indirektem Material und Warengruppen 2000, 4333 bis 4338, 4340', desc: 'Das Controlling gibt die Anfrage frei.' },
    { role: 'Werkleitung', name: 'Freigabe Werkleitung', kind: 'decision', applies: preis >= 1500,
      cond: 'ab 1.500 € Bestellvolumen', desc: 'Die Werkleitung gibt das Bestellvolumen frei.' },
    { role: 'Strategischer Einkauf', name: 'Freigabe strategischer Einkauf', kind: 'decision',
      applies: preis >= 10000 && !lead, cond: 'ab 10.000 € und ohne Lead-Buyer-Abschluss',
      desc: 'Der strategische Einkauf gibt den Konzernrahmen frei.' },
    { role: 'Power Automate', name: 'Freigegeben', kind: 'auto', applies: true,
      desc: 'Alle nötigen Stufen haben zugestimmt. Der Antragsteller bekommt eine Mail.' },
    { role: 'Einkauf', name: 'Bestellung auslösen (Bestellt)', kind: 'task', applies: true,
      desc: 'Der Einkauf bestellt im ERP und trägt die Bestellnummer ein. Die App prüft sie auf Dubletten und setzt Bestellt.' },
  ];
}

function plNextApplicable(from) {
  for (let j = from; j < plState.stages.length; j++) if (plState.stages[j].applies) return j;
  return -1;
}

function startProbelauf() {
  const preis = Math.max(0, parseFloat(String($id('pl-preis').value).replace(',', '.')) || 0);
  const art = $id('pl-art').value;
  const lead = $id('pl-lead').checked;
  plState = { stages: plBuildStages(preis, art, lead), idx: 0, rejected: null, done: false, preis, art, lead };
  plState.idx = plNextApplicable(0);
  renderProbelauf();
}

function plAdvance() {
  if (!plState || plState.rejected !== null || plState.done) return;
  const n = plNextApplicable(plState.idx + 1);
  if (n === -1) { plState.done = true; plState.idx = plState.stages.length; }
  else plState.idx = n;
  renderProbelauf();
}

function plReject() {
  if (!plState || plState.done || plState.rejected !== null) return;
  if (plState.stages[plState.idx]?.kind !== 'decision') return;
  plState.rejected = plState.idx;
  renderProbelauf();
}

function plReset() {
  plState = null;
  const el = $id('pl-stage');
  if (el) el.innerHTML = '';
}

function renderProbelauf() {
  if (!plState) return;
  const S = plState;
  const decisions = S.stages.filter(s => s.applies && s.kind === 'decision').length;
  const skipped = S.stages.filter(s => !s.applies).map(s => s.role);

  const steps = S.stages.map((st, i) => {
    let status;
    if (!st.applies) status = 'skip';
    else if (S.rejected !== null) status = i < S.rejected ? 'done' : (i === S.rejected ? 'reject' : 'cancel');
    else status = i < S.idx ? 'done' : (i === S.idx ? 'current' : 'pending');
    const mark = status === 'done' ? '✓' : status === 'reject' ? '✗'
      : status === 'skip' ? '–' : status === 'cancel' ? '·' : (i + 1);
    const note = status === 'skip' ? ' <span class="pl-skip-note">entfällt für diese Werte</span>' : '';
    const tag = st.cond ? ` <span class="flow-tag">${esc(st.cond)}</span>` : '';
    return `
      <div class="pl-step pl-${status}">
        <div class="pl-mark">${mark}</div>
        <div class="pl-body">
          <div class="pl-head">${roleBadge(st.role)} <span class="pl-name">${esc(st.name)}</span>${note}</div>
          <div class="pl-desc">${esc(st.desc)}${tag}</div>
        </div>
      </div>`;
  }).join('<div class="flow-arrow">↓</div>');

  let controls;
  if (S.rejected !== null) {
    const st = S.stages[S.rejected];
    const folge = st.role === 'Einkauf'
      ? 'Die Anfrage steht auf Abgelehnt. Eine Mail geht nicht raus, der Antragsteller sieht den Grund in „Meine Anfragen".'
      : 'Power Automate setzt Abgelehnt und schreibt dem Antragsteller eine Mail. Die folgenden Stufen entfallen.';
    controls = `
      <div class="pl-result pl-result-rej"><b>✗ Abgelehnt</b> durch ${esc(st.role)}. ${folge}</div>
      <div class="pl-actions"><button class="btn btn-ghost" onclick="plReset()">Zurücksetzen</button></div>`;
  } else if (S.done) {
    controls = `
      <div class="pl-result pl-result-ok"><b>📦 Bestellt.</b> Durchlaufen: ${decisions} Freigabestufe${decisions === 1 ? '' : 'n'}
        bei ${fmtEuro(S.preis)}, ${esc(PL_ART[S.art] || S.art)}${S.lead ? ', Lead-Buyer-Abschluss' : ''}.</div>
      <div class="pl-actions"><button class="btn btn-ghost" onclick="plReset()">Nochmal</button></div>`;
  } else {
    const st = S.stages[S.idx];
    const hint = st.kind === 'decision' ? 'diese Rolle entscheidet.'
      : st.kind === 'auto' ? 'läuft automatisch.' : 'Aufgabe der Rolle.';
    const btns = st.kind === 'decision'
      ? `<button class="btn btn-success" onclick="plAdvance()">✓ Genehmigen</button>
         <button class="btn btn-danger" onclick="plReject()">✗ Ablehnen</button>`
      : `<button class="btn btn-primary" onclick="plAdvance()">Weiter →</button>`;
    controls = `
      <div class="pl-current-hint">Aktuell: <b>${esc(st.name)}</b> · ${hint}</div>
      <div class="pl-actions">${btns} <button class="btn btn-ghost" onclick="plReset()">Zurücksetzen</button></div>`;
  }

  const summary = `<div class="pl-summary">${fmtEuro(S.preis)} · ${esc(PL_ART[S.art] || S.art)}${S.lead ? ' · Lead-Buyer' : ''} · ${decisions} Freigabestufe${decisions === 1 ? '' : 'n'}${skipped.length ? ' · entfällt: ' + skipped.map(esc).join(', ') : ''}</div>`;
  $id('pl-stage').innerHTML = summary + `<div class="pl-steps">${steps}</div>` + controls;
}
