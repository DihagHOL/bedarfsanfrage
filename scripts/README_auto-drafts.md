# Automatische Bedarfsanfrage-Entwürfe – Einrichtung

Ein geplanter **GitHub-Actions-Cron** (`.github/workflows/auto-drafts.yml`) führt täglich
`scripts/auto-drafts.mjs` aus. Das Skript meldet sich **app-only** an Microsoft Graph an,
sucht Anfragen mit fälligem Feld **„Datum zur Automatische Bedarfsanfrage"** und legt je
einen Entwurf (Status **„Automatisch erstellter Entwurf"**) an; danach wird das Quelldatum
geleert.

Damit das läuft, ist **einmalig** ein app-only-Zugang nötig (Azure-Admin):

---

## 1. App-Registrierung anlegen
1. **Azure-Portal** → *Microsoft Entra ID* → *App-Registrierungen* → **Neue Registrierung**.
2. Name z. B. `BANF Auto-Entwürfe`, *Unterstützte Kontotypen*: „Nur dieses Verzeichnis".
   → **Registrieren**.
3. Auf der Übersichtsseite notieren:
   - **Anwendungs-(Client-)ID** → später `CLIENT_ID`
   - **Verzeichnis-(Mandanten-)ID** → später `TENANT_ID`
     (für dieses Tenant: `fdb70646-023a-403b-a4b9-1f474a935123`)

## 2. Client-Secret erstellen
1. In der App → *Zertifikate & Geheimnisse* → **Neuer geheimer Clientschlüssel**.
2. Ablauf z. B. 24 Monate → **Hinzufügen**.
3. Den **Wert** (nicht die „Geheimnis-ID") **sofort kopieren** → später `CLIENT_SECRET`.
   (Wird nur einmal angezeigt.)

## 3. Graph-Berechtigung „Sites.Selected" + Admin-Consent
1. In der App → *API-Berechtigungen* → **Berechtigung hinzufügen** → *Microsoft Graph*
   → **Anwendungsberechtigungen** → suche **`Sites.Selected`** → hinzufügen.
2. **„Administratorzustimmung erteilen"** klicken (Häkchen muss grün werden).

## 4. App Schreibrecht NUR auf die Site `gruppe_shb` geben
`Sites.Selected` gibt zunächst **keinen** Zugriff – die Site muss explizit freigegeben
werden. Das macht ein Admin per **Graph Explorer** (https://aka.ms/ge), angemeldet mit
einem Admin-Konto:

**a) Site-ID holen** (GET):
```
https://graph.microsoft.com/v1.0/sites/dihag.sharepoint.com:/sites/gruppe_shb
```
→ aus der Antwort das Feld `id` kopieren (Form: `dihag.sharepoint.com,<guid>,<guid>`).

**b) Schreibrecht für die App setzen** (POST):
```
POST https://graph.microsoft.com/v1.0/sites/<SITE-ID>/permissions
Content-Type: application/json

{
  "roles": ["write"],
  "grantedToIdentities": [
    { "application": { "id": "<CLIENT_ID>", "displayName": "BANF Auto-Entwürfe" } }
  ]
}
```
(Benötigt beim ausführenden Admin die Berechtigung `Sites.FullControl.All` – im Graph
Explorer ggf. unter „Modify permissions" zustimmen.)

## 5. GitHub-Secrets hinterlegen
Repo **dfedorov12/bedarfsanfrage** → *Settings* → *Secrets and variables* → *Actions*
→ **New repository secret**, jeweils anlegen:

| Name            | Wert                                   |
|-----------------|----------------------------------------|
| `TENANT_ID`     | Verzeichnis-(Mandanten-)ID aus Schritt 1 |
| `CLIENT_ID`     | Anwendungs-(Client-)ID aus Schritt 1     |
| `CLIENT_SECRET` | Secret-**Wert** aus Schritt 2            |

## 6. SharePoint-Spalte sicherstellen
In der Liste **Bedarfsanfrage** muss die Spalte **„Datum zur Automatische Bedarfsanfrage"**
(Typ *Datum*) existieren und der Status **„Automatisch erstellter Entwurf"** als Auswahlwert
in der Status-Spalte vorhanden sein (beides bereits vorhanden).

## 7. Testen
- Repo → Tab **Actions** → Workflow *„Automatische Bedarfsanfrage-Entwürfe"* →
  **Run workflow** (manueller Lauf).
- Lege zur Probe in einer Testanfrage das Datum auf **heute** und starte den Workflow.
- Im Lauf-Log siehst du `… fällig` und `✓ Entwurf aus #… erstellt`.
- Danach läuft er automatisch **täglich um 05:00 UTC**.

---

### Anpassen
- **Uhrzeit/Intervall:** `cron`-Ausdruck in `.github/workflows/auto-drafts.yml`.
- **Logik:** `scripts/auto-drafts.mjs` ist reines Node – Felder, Filter, Status etc. frei
  änderbar.

### Gut zu wissen
- Geplante GitHub-Workflows **pausieren nach 60 Tagen ohne Repo-Aktivität** (ein Commit
  reaktiviert sie) und können sich um einige Minuten verzögern.
- Das Secret läuft ab (Schritt 2) → rechtzeitig erneuern und in GitHub aktualisieren.
- Die **App-seitige** Auto-Erstellung ist deaktiviert (in `app.js` auskommentiert), damit
  keine doppelten Entwürfe entstehen.
