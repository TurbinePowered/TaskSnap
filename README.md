# TaskSnap-WebApp

Statische PWA fürs iPhone: Notizen und Bilder in `OneDrive/TaskSnap/Eingang` ablegen, Aufgaben aus `tasks/`
anzeigen und abhaken. Kein Build, kein Server – nur Dateien, gehostet auf GitHub Pages.
Bedienung am iPhone: [../IPHONE.md](../IPHONE.md), Hintergründe: [../KONZEPT.md](../KONZEPT.md).

| Datei | Inhalt |
|---|---|
| `config.js` | Client-ID der App-Registrierung, Datenordner im OneDrive |
| `app.js` | Anmeldung (MSAL), Graph-Zugriff, Upload-Warteschlange (IndexedDB), Anzeige |
| `sw.js` | Service Worker: App-Hülle offline, Netz zuerst |
| `vendor/msal-browser.min.js` | `@azure/msal-browser` 5.20.0 (MIT, siehe `msal-browser.LICENSE`) |
| `serve.py` | lokaler Testserver |

## Lokal testen

```powershell
python web\serve.py
```

- `http://localhost:5173/?demo` – Oberfläche mit Beispieldaten, ohne Anmeldung
- `http://localhost:5173/` – echt, mit Anmeldung (Redirect-URI `http://localhost:5173/` muss registriert sein).
  Achtung: arbeitet auf dem echten OneDrive-Ordner.

## App-Registrierung

Entra (NewTec-Verzeichnis) → App-Registrierungen → **TaskSnap**, Client-ID `b24fe469-2239-4a29-a9de-9e2288dea4cb`.

- Unterstützte Kontotypen: **Nur persönliche Microsoft-Konten**
- Authentifizierung → Plattform **Single-Page-Anwendung**, Umleitungs-URIs:
  `https://turbinepowered.github.io/TaskSnap/` und `http://localhost:5173/`
- API-Berechtigungen: Microsoft Graph, delegiert, `Files.ReadWrite` (+ `User.Read`)

## Veröffentlichen (GitHub Pages)

Einmalig: auf GitHub das **öffentliche** Repo `TurbinePowered/TaskSnap` (angelegt am 2026-10-12), dann im Klon `bellmann-tools`:

```powershell
git remote add github https://github.com/TurbinePowered/TaskSnap.git
```

GitHub → Repo → Settings → Pages → *Deploy from a branch*, Branch `main`, Ordner `/ (root)`.

Jedes Mal (nach dem Commit im GitLab-Repo):

```powershell
git subtree split --prefix TaskSnap/web -b tasksnap-web
git push github tasksnap-web:main
```

Nach 1–2 Minuten ist die neue Fassung online; die App lädt sie beim nächsten Öffnen (Netz zuerst).
