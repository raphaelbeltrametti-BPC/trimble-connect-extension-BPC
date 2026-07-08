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

1. Extension als Project Extension registrieren.
2. Lokale URL fuer Entwicklung: `http://localhost:3000`.
3. Produktions-URL im Manifest und in Trimble Connect auf die Vercel-URL setzen.
4. In der App die richtige API-Region waehlen, bei Schweizer/EU-Projekten meist `EU`.

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
