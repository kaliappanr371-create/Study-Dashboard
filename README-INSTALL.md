# Study Dashboard – installable Windows app (PWA)

## Install (5 minutes)
1. Unzip this folder somewhere permanent (e.g. `C:\Apps\Study_Dashboard_PWA`).
2. Double-click **serve.bat** (needs Python 3). No Python? Right-click **serve.ps1** → Run with PowerShell.
   Keep that window open; the app is served at **http://localhost:8765** (always this address – your data is stored per address).
3. Open http://localhost:8765 in **Edge or Chrome**. Wait ~2 seconds (first load saves the app for offline use).
4. Install: Edge → ⋯ menu → Apps → Install this site as an app. Chrome → install icon in the address bar. Or Data & sync → *App & offline* → Install app.
5. After install the app opens from the Start menu in its own window and **works with the server stopped and no internet**.
   (The server is only needed for the first install and for updates.)

## Move your existing data (do this once, BEFORE closing the old file)
1. Open your old HTML file → Data & sync → **Backup & restore → Download backup (.json)**.
2. Open the installed app → Data & sync → Backup & restore → **Restore** → choose that .json.
   (Browser storage is per address, so file:// data cannot appear automatically.) Then relink Excel files if you used Live link.

## Test checklist
- Data & sync → *App & offline* shows "Works offline: Yes".
- Stop the server window, close and reopen the app from Start → it still opens and your data is there.
- Import an .xlsx and export one with the server stopped.
- Upgrade: replace the folder files with a newer build, start the server, open the app → "A new version is ready" → Update. Progress is untouched; one pre-update copy is kept automatically.

## Files
- index.html – page shell, links manifest/icons, loads the app.
- assets/app.js, assets/style.css – the original app code (patched, see below). assets/seed.js – empty (blank template, no personal data).
- manifest.webmanifest – name, standalone display, icons (incl. maskable), theme colours, 3 shortcuts.
- sw.js – service worker: precaches the app, cache-first offline, offline navigation fallback, user-approved updates, old caches removed.
- icons/ – 192/512/maskable/apple/favicon.
- serve.py / serve.bat / serve.ps1 – local server on fixed port 8765.

## What changed in the app code
- Data safety: unreadable/newer-version data is parked under a recovery key, never overwritten; one snapshot kept before each app update; migration hook for future schema changes; persistent storage requested.
- Timers: the always-on 500 ms and 20 s timers are gone. A timer now exists only while a focus session or an unfired reminder exists (30 s when the window is hidden); Excel polling only while linked files exist and the window is visible.
- Styling: removed backdrop blur and a hover transition; reduced-motion respected.
- New "App & offline" card (install, offline status, storage); `?tab=` shortcuts; theme-colour follows theme.
