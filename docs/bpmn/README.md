# Prozessmodelle (BPMN 2.0)

Die Bedarfsanfrage als BPMN 2.0 im Hausschema des RMS: zwei Hauptprozesse und drei
Unterprozesse, die per Aufrufaktivität (⊞) eingebunden sind.

**Führend sind die Modelle im RMS** (rms.dihag.de, Reiter Prozesse, Ablage SHB; Bibliothek
„ISMS Dokumente" auf sites/ISMS, Ordner Prozesse/SHB). Der Reiter „Prozess" der App liest sie
live von dort, sobald das angemeldete Konto die Bibliothek lesen darf. Wer im RMS ein Modell
ändert, ändert es damit auch in der App, und weil die Unterprozesse eingebunden statt kopiert
sind, auch im Hauptprozess. Ohne Zugriff zeigt die App die Kopie aus diesem Ordner und sagt
das dazu. „im RMS öffnen" führt über `?modell=<Kennung>` direkt in die Ansicht des Modells.

| Datei | Modell im RMS | Bindet ein |
|-------|---------------|------------|
| `01-bedarfsanfrage.bpmn` | Bedarfsanfrage | Anfrage erfassen, Anfrage prüfen und entscheiden, Freigabelauf; Nebenprozess Auto-Entwurf |
| `01a-anfrage-erfassen.bpmn` | Bedarfsanfrage-Erfassen | |
| `01b-anfrage-pruefen.bpmn` | Bedarfsanfrage-Einkaufsprüfung | |
| `02-freigabelauf.bpmn` | Bedarfsanfrage-Freigabelauf | |
| `03-auto-entwurf.bpmn` | Bedarfsanfrage-AutoEntwurf | |

Der Auto-Entwurf ist ein eigener Prozess, weil er mit einer Uhrzeit beginnt und nicht mit
einem Bedarf (R1: genau ein Auslöser). Im RMS hängt er als Nebenprozess am Hauptprozess.

`modelle.json` ist die Liste, aus der die App liest: Datei, Prozess-Kennung, Dateiname und
Kennung im RMS (`itemId`) und die Bibliothek (`driveId`). `rms-ids.json` hält die Kennungen
nach dem Import fest. Die Kennung einer Datei ist `01` plus Base32 aus dem Präfix der
Bibliothek (`69fcd7f4`) und der SharePoint-Kennung (UniqueId, Bytes in Windows-Reihenfolge).

## Was die Modelle abbilden

Stand 11.11.2026, abgeglichen mit dem Code der App, dem Flow „Bedarfsanfragen SHB" und dem
Auto-Entwurf-Lauf in GitHub Actions:

- Neue Anfragen lösen keine Mail aus. Der Flow „Bedarfsanfragen SHB Einkauf" ist seit
  04.11.2026 aus, der Einkauf entscheidet in der App.
- Öffnet der zuständige Einkauf eine eingereichte Anfrage oder einen Auto-Entwurf, setzt die
  App den Status auf In Prüfung (Einkauf). Zuständig ist, wem die Warengruppe zugeordnet ist.
- Bis zur Entscheidung des Einkaufs kann der Antragsteller die Anfrage zurückziehen.
- Lehnt der Einkauf ab, geht keine Mail raus. Freigabe und Ablehnung im Flow melden sich beim
  Antragsteller per Mail.
- Freigabelauf: Controlling nur bei Dienstleistung, Werkleitung ab 1.500 €, strategischer
  Einkauf ab 10.000 € und ohne Lead-Buyer-Abschluss. Approvals vom Typ „erste Antwort".

## Hausschema

Alle fünf Modelle bestehen die Prüfung des RMS (`prozessSchemaPruefen`) ohne Befund. Dazu
prüft `tests/test-prozess.mjs` die erzeugten Dateien (R1 bis R10, Formen und Kanten, frisch
erzeugt, Marker, keine Mailadressen) und die Anbindung der Seite:

```bash
node tests/test-prozess.mjs
```

Verknüpfte Regelwerke: 120 Konzernfachregelung Prozessmanagement, 119 Aufbewahrungsfristen
und -pflichten, 111 Sichere Softwareentwicklung. Eine Einkaufs- oder Beschaffungsrichtlinie
gibt es im RMS noch nicht. Die Wertgrenzen stehen bisher nur in der App und im Flow.

## Farben

| Farbe | Beteiligter |
|-------|-------------|
| Orange | Menschen: Antragsteller, Einkauf, Controlling, Werkleitung, strategischer Einkauf |
| Dunkelblau | Bedarfsanfrage-App |
| Blau | Power Automate und GitHub Actions |
| Hellgrau | extern: Lieferant |

Ergebnisfarben: grün = Start und gutes Ende, gelb = Entscheidung, rot = abgelehnt,
grau = zurückgezogen oder nichts zu tun.

## Ändern

Im Alltag im RMS: Modell öffnen, bearbeiten, speichern. Die App zeigt den neuen Stand beim
nächsten Laden.

Für größere Umbauten gibt es den Generator `scripts/bpmn/generate-bpmn.js` (Motor aus der
E-Rechnung: Knoten im Raster, Flüsse, Notizen). Er schreibt die fünf Dateien und
`modelle.json` neu und setzt die Marker aus `rms-ids.json`:

```bash
node scripts/bpmn/generate-bpmn.js
```

Danach müssen die Dateien wieder ins RMS (gleicher Name in Prozesse/SHB ergibt eine neue
Version derselben Datei, die Kennungen bleiben). Vorsicht: Das überschreibt Änderungen, die
seit dem letzten Import im RMS gemacht wurden. Deshalb vorher die aktuelle Fassung lesen und
mit der letzten eigenen vergleichen. Was das RMS selbst pflegt, steht in der Dokumentation des
Prozesses hinter der Zeile „Anwender-Ansicht" (Nebenprozesse, Regelwerke, Prozessmanagement,
Kennzahlen) und muss beim Zurückspielen aus der RMS-Fassung übernommen werden.

Letzte Übernahme: siehe Session-Log.
