# Trimble Connect Team & Berechtigungsmanager

React + TypeScript + Vite Web Extension fuer Trimble Connect. Die Extension liest eine
CDE-Berechtigungsmatrix aus Excel, erstellt fehlende Trimble-Gruppen und bereitet
Folder-Permissions als Dry-Run vor, bevor sie angewendet werden.

## Aktueller Stand

- Workspace API Verbindung ueber `trimble-connect-workspace-api`
- Access Token via `extension.requestPermission("accesstoken")`
- Projekt-ID via `project.getProject()`
- Trimble REST API v2 ueber konfigurierbare Region
- Teams in der UI, technisch Trimble `Groups`
- Excel-Parser fuer Matrixwerte `V`, `L`, `K`, `Vollzugriff`, `Lesezugriff`, `Kein Zugriff`
- Dry-Run fuer `FULL_ACCESS`, `READ`, `NO_ACCESS`
- Anwenden per `PATCH /folders/fs/{folderId}/permissions`

## Wichtige Korrekturen zur ersten Voreinstellung

- Die verlinkte Project Workspace API ist Legacy; die App nutzt die aktuelle
  `trimble-connect-workspace-api`.
- Trimble REST nutzt laut SwaggerHub `Groups`, nicht `/projects/{projectId}/teams`.
- Gruppen werden ueber `GET /groups?projectId=...` und `POST /groups` verwaltet.
- Folder-Permissions erwarten ACLs mit `READ`, `FULL_ACCESS`, `NO_ACCESS`.
- Gruppen-Actor-IDs muessen in ACLs als `tc-groups:<groupId>` gesendet werden.
- Die Excel-Datei nutzt ausgeschriebene Werte, nicht nur `V/L/K`.
- Die API-Basis ist regional, z.B. EU: `https://app21.connect.trimble.com/tc/api/2.0`.
- Lokale Entwicklung laeuft gemaess Konzept auf `http://localhost:3000`.
- Versuch, `"type": "panel"` und `"extensionType": ["project"]` selbst im
  Manifest zu setzen, war falsch: Trimble Connect hat die URL dann als
  "not a valid extension" abgelehnt. `extensionType` wird von Trimble Connect
  selbst gesetzt, je nachdem ob man die Manifest-URL unter Projekt- oder unter
  3D-Viewer-Einstellungen eintraegt — es ist kein Feld, das der Manifest-Autor
  liefert. Das Manifest bleibt daher bei den dokumentierten Feldern
  `url`, `icon`, `title`, `description` (ohne eigenes `type`/`extensionType`).
  Ob die Extension im Projekt-Explorer oder im 3D-Viewer erscheint, wird
  ausschliesslich dadurch bestimmt, unter welcher Einstellungsseite man die
  Manifest-URL registriert (siehe "Registrierung in Trimble Connect").

## Entwicklung

```bash
npm install
npm run dev
```

Falls PowerShell `npm.ps1` blockiert:

```bash
npm.cmd run dev
npm.cmd run build
```

## Registrierung in Trimble Connect

1. Extension unter **Projekt > Einstellungen > Extensions** registrieren
   (nicht unter 3D-Viewer > Einstellungen > Extensions — das erzeugt eine
   3D-Viewer-Erweiterung statt einer Projekt-Extension im Explorer).
2. Falls die Extension zuvor faelschlich im 3D-Viewer registriert wurde: dort
   entfernen und stattdessen unter Projekt > Einstellungen > Extensions mit
   der Manifest-URL neu hinzufuegen, damit sie im linken Navigationsmenü und
   im mittleren/rechten Bereich der Projektseite erscheint.
3. Lokale URL fuer Entwicklung: `http://localhost:3000`.
4. Produktions-URL im Manifest und in Trimble Connect auf die Vercel-URL setzen.
5. In der App die richtige API-Region waehlen, bei Schweizer/EU-Projekten meist `EU`.

## Projektstruktur

```text
src/
  api/
    trimble.ts        REST Client fuer Groups, Projects, Folders, Permissions
  excel/
    parser.ts         Excel-Matrix Parser
  hooks/
    useApi.ts         Workspace API Verbindung und Token
  permissions/
    planner.ts        Dry-Run Planung fuer Folder ACLs
  types/
    index.ts          Gemeinsame TypeScript Typen
  utils/
    text.ts           Normalisierung fuer Namen und Pfade
  App.tsx             UI und Workflow
  index.css           Layout und Styling
```

## Naechster fachlicher Test

1. App in Trimble Connect oeffnen.
2. Excel-Matrix laden.
3. Gruppen laden.
4. Fehlende Gruppen erstellen.
5. Ordner scannen.
6. Dry-Run pruefen.
7. Erst danach `Aenderungen wirklich anwenden` aktivieren.
