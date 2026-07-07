# Trimble Connect Extension Template

React + TypeScript + Vite Template für Trimble Connect Extensions.

## Setup

```bash
npm install
npm run dev
```

## Deployment (Vercel)

```bash
git init
git add .
git commit -m "initial commit"
# → Vercel mit GitHub verknüpfen
```

## Konfiguration

1. `public/manifest.json` anpassen:
   - `id` → neue UUID generieren
   - `name` → Extension-Name
   - `url` → Vercel URL

2. `src/hooks/useApi.ts`:
   - `setMenu` URL anpassen

3. `src/App.tsx`:
   - ViewerApp und PanelApp befüllen

## TC API Docs

- Viewer API: https://components.connect.trimble.com/trimble-connect-workspace-api/index.html
- Project API: https://components.connect.trimble.com/trimble-connect-project-workspace-api/docs/index.html

## Wichtige Erkenntnisse

- getObjectProperties: max 10 IDs pro Aufruf
- setSelection: plain Array [123, 456] — kein wrapped Object
- Viewer-Kontext erkennen: viewer.onCameraChanged feuert NUR im 3D Viewer
- Access Token: api.extension.requestPermission("accesstoken")
- Projekt ID: api.project.getProject() → {id: "..."}

## Ordnerstruktur

```
src/
  components/    — UI Komponenten
  hooks/
    useApi.ts    — TC Workspace API Hook (Kern)
  types/
    index.ts     — TypeScript Typen
  App.tsx        — Haupt-App (Viewer vs Panel)
  index.css      — Design System (TC-Farben + Klassen)
  main.tsx       — Entry Point
public/
  manifest.json  — TC Extension Manifest
  icon.svg       — Extension Icon
vercel.json      — Vercel Konfiguration
```
