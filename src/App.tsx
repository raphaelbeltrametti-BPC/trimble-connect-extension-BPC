import { useEffect, useMemo, useRef, useState } from "react";
import { API_BASE_URLS, TrimbleApiError, TrimbleClient } from "./api/trimble";
import { useApi } from "./hooks/useApi";
import { buildPermissionPlan } from "./permissions/planner";
import { buildFolderCreationPlan } from "./permissions/folderPlanner";
import type {
  ApiRegion,
  FolderCreationStep,
  LogEntry,
  PermissionMatrix,
  PermissionPlanItem,
  ProgressState,
  TCFolder,
  TCGroup,
  TCProject,
  WorkbookModel,
} from "./types";
import { normalizeLookup } from "./utils/text";
import { loadStoredWorkbook, saveStoredWorkbook } from "./utils/workbookStorage";
import "./index.css";

type TabId = "teams" | "folders" | "permissions" | "log";

export default function App() {
  const { connected, connectionError, accessToken, project, hostName, refreshAccessToken } = useApi();
  const [activeTab, setActiveTab] = useState<TabId>("teams");
  const [region, setRegion] = useState<ApiRegion>("eu");
  const [regionTouched, setRegionTouched] = useState(false);
  const [customBaseUrl, setCustomBaseUrl] = useState(API_BASE_URLS.eu);
  const [manualProjectId, setManualProjectId] = useState("");
  const [manualToken, setManualToken] = useState("");
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [groups, setGroups] = useState<TCGroup[]>([]);
  const [folders, setFolders] = useState<TCFolder[]>([]);
  const [projectDetails, setProjectDetails] = useState<TCProject | null>(null);
  const [newGroupName, setNewGroupName] = useState("");
  const [plan, setPlan] = useState<PermissionPlanItem[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [busy, setBusy] = useState("");
  const [progress, setProgress] = useState<ProgressState>({ current: 0, total: 0, label: "" });
  const [allowApply, setAllowApply] = useState(false);
  const [allowDeleteGroups, setAllowDeleteGroups] = useState(false);
  const [targetParentId, setTargetParentId] = useState("");
  const [targetParentFilter, setTargetParentFilter] = useState("");
  const [folderPlan, setFolderPlan] = useState<FolderCreationStep[]>([]);
  const [allowCreateFolders, setAllowCreateFolders] = useState(false);

  useEffect(() => {
    if (regionTouched || !project?.location) return;
    const guessed = guessRegion(project.location);
    if (guessed) setRegion(guessed);
  }, [project?.location, regionTouched]);

  const baseUrl = region === "custom" ? customBaseUrl : API_BASE_URLS[region];
  const effectiveProject = project ?? (manualProjectId.trim()
    ? { id: manualProjectId.trim(), name: "Manuelles Projekt" }
    : null);

  const restoredWorkbookRef = useRef(false);

  useEffect(() => {
    if (restoredWorkbookRef.current || workbook || !effectiveProject?.id) return;
    restoredWorkbookRef.current = true;

    const stored = loadStoredWorkbook(effectiveProject.id);
    if (stored) {
      setWorkbook(stored);
      setSelectedSheet(stored.sheetNames[0] ?? "");
      addLog(
        "info",
        `Zuletzt geladene Excel-Matrix wiederhergestellt: ${stored.fileName}.`,
        `${stored.sheetNames.length} Phasen, ${stored.teamNames.length} Teams.`
      );
    }
  }, [effectiveProject?.id, workbook]);

  useEffect(() => {
    if (!workbook || !effectiveProject?.id) return;
    const result = saveStoredWorkbook(effectiveProject.id, workbook);
    if (!result.ok) {
      addLog(
        "warning",
        "Matrix konnte nicht lokal gespeichert werden (ueberlebt keinen Hardreload).",
        result.error
      );
    }
  }, [workbook, effectiveProject?.id]);

  const client = useMemo(
    () =>
      new TrimbleClient({
        baseUrl,
        delayMs: 250,
        getToken: async () => {
          if (manualToken.trim()) return manualToken.trim();
          if (accessToken) return accessToken;
          return refreshAccessToken();
        },
      }),
    [accessToken, baseUrl, manualToken, refreshAccessToken]
  );

  const [accessState, setAccessState] = useState<"checking" | "admin" | "denied">("checking");
  const [accessRole, setAccessRole] = useState("");
  const [accessError, setAccessError] = useState("");

  useEffect(() => {
    if (!effectiveProject?.id) {
      setAccessState("checking");
      return;
    }

    let cancelled = false;
    setAccessState("checking");
    setAccessError("");

    client
      .getProject(effectiveProject.id)
      .then((details) => {
        if (cancelled) return;
        setAccessRole(details.role ?? "");
        setAccessState(details.role === "ADMIN" ? "admin" : "denied");
      })
      .catch((error) => {
        if (cancelled) return;
        setAccessRole("");
        setAccessError(formatError(error));
        setAccessState("denied");
      });

    return () => {
      cancelled = true;
    };
  }, [client, effectiveProject?.id]);

  const currentMatrix = useMemo(
    () => workbook?.matrices.find((matrix) => matrix.sheetName === selectedSheet) ?? null,
    [selectedSheet, workbook]
  );

  const effectiveTargetParentId = targetParentId || projectDetails?.rootId || "";

  const targetParentCandidates = useMemo(() => {
    const query = normalizeLookup(targetParentFilter);
    if (!query) return [];
    return folders
      .filter((folder) => normalizeLookup(folder.name).includes(query) || normalizeLookup(folder.path).includes(query))
      .slice(0, 50);
  }, [folders, targetParentFilter]);

  const targetParentLabel = useMemo(() => {
    if (targetParentId) {
      const match = folders.find((folder) => folder.id === targetParentId);
      return match ? match.path : targetParentId;
    }
    return projectDetails?.rootId ? "Projekt-Root (Standard)" : "";
  }, [folders, projectDetails?.rootId, targetParentId]);

  const missingTeamNames = useMemo(() => {
    if (!workbook) return [];
    const existing = new Set(groups.map((group) => normalizeLookup(group.name)));
    return workbook.teamNames.filter((teamName) => !existing.has(normalizeLookup(teamName)));
  }, [groups, workbook]);

  const unusedGroups = useMemo(() => {
    if (!workbook) return [];
    const wanted = new Set(workbook.teamNames.map((name) => normalizeLookup(name)));
    return groups.filter((group) => !wanted.has(normalizeLookup(group.name)));
  }, [groups, workbook]);

  const planStats = useMemo(() => {
    const ready = plan.filter((item) => item.status === "ready").length;
    const missingFolders = plan.filter((item) => item.status === "missing-folder").length;
    const missingGroups = plan.filter((item) => item.status === "missing-groups").length;
    return { ready, missingFolders, missingGroups };
  }, [plan]);

  function addLog(level: LogEntry["level"], message: string, detail?: string) {
    setLogs((current) => [
      {
        id: `${Date.now()}-${Math.random()}`,
        level,
        message,
        detail,
        createdAt: new Date().toLocaleTimeString("de-CH", { hour: "2-digit", minute: "2-digit", second: "2-digit" }),
      },
      ...current,
    ].slice(0, 250));
  }

  async function handleWorkbook(file: File) {
    setBusy("excel");
    setPlan([]);
    try {
      const { parseWorkbookFile } = await import("./excel/parser");
      const parsed = await parseWorkbookFile(file);
      setWorkbook(parsed);
      setSelectedSheet(parsed.sheetNames[0] ?? "");
      addLog("success", `Excel geladen: ${parsed.sheetNames.length} Phasen, ${parsed.teamNames.length} Teams.`);
    } catch (error) {
      addLog("error", "Excel konnte nicht gelesen werden.", formatError(error));
    } finally {
      setBusy("");
    }
  }

  async function loadGroups() {
    const tcProject = getProjectOrLog(effectiveProject, addLog);
    if (!tcProject) return [];

    setBusy("groups");
    try {
      const loaded = await client.listGroups(tcProject.id);
      setGroups(loaded);
      addLog("success", `${loaded.length} Trimble-Gruppen geladen.`);
      return loaded;
    } catch (error) {
      addLog("error", "Gruppen konnten nicht geladen werden.", formatError(error));
      return [];
    } finally {
      setBusy("");
    }
  }

  async function createSingleGroup() {
    const name = newGroupName.trim();
    if (!name) return;

    const tcProject = getProjectOrLog(effectiveProject, addLog);
    if (!tcProject) return;

    setBusy("create-group");
    try {
      const group = await client.createGroup(tcProject.id, name);
      setGroups((current) => uniqueGroups([...current, group]));
      setNewGroupName("");
      addLog("success", `Gruppe erstellt: ${group.name}.`);
    } catch (error) {
      addLog("error", `Gruppe konnte nicht erstellt werden: ${name}.`, formatError(error));
    } finally {
      setBusy("");
    }
  }

  async function importMissingGroups() {
    if (!workbook) {
      addLog("warning", "Bitte zuerst die Berechtigungsmatrix laden.");
      return;
    }

    const tcProject = getProjectOrLog(effectiveProject, addLog);
    if (!tcProject) return;

    setBusy("import-groups");

    try {
      let existingGroups = groups.length > 0 ? groups : await client.listGroups(tcProject.id);
      const existingNames = new Set(existingGroups.map((group) => normalizeLookup(group.name)));
      const namesToCreate = workbook.teamNames.filter((name) => !existingNames.has(normalizeLookup(name)));
      let created = 0;

      setProgress({ current: 0, total: namesToCreate.length, label: "Teams erstellen" });

      for (let index = 0; index < namesToCreate.length; index += 1) {
        const name = namesToCreate[index];
        try {
          const group = await client.createGroup(tcProject.id, name);
          existingGroups = uniqueGroups([...existingGroups, group]);
          existingNames.add(normalizeLookup(group.name));
          created += 1;
          addLog("success", `Gruppe erstellt: ${group.name}.`);
        } catch (error) {
          addLog("error", `Gruppe konnte nicht erstellt werden: ${name}.`, formatError(error));
        }
        setProgress({ current: index + 1, total: namesToCreate.length, label: "Teams erstellen" });
      }

      setGroups(existingGroups);
      addLog("info", `Team-Import fertig. Erstellt: ${created}, Fehler: ${namesToCreate.length - created}.`);
    } catch (error) {
      addLog("error", "Team-Import abgebrochen.", formatError(error));
    } finally {
      setBusy("");
      setProgress({ current: 0, total: 0, label: "" });
    }
  }

  async function deleteUnusedGroups() {
    if (!allowDeleteGroups) {
      addLog("warning", "Loeschen blockiert: Checkbox fuer Gruppen-Loeschung ist nicht gesetzt.");
      return;
    }

    if (unusedGroups.length === 0) {
      addLog("warning", "Keine ungenutzten Gruppen zum Loeschen.");
      return;
    }

    setBusy("delete-groups");
    setProgress({ current: 0, total: unusedGroups.length, label: "Gruppen loeschen" });

    let deleted = 0;
    for (let index = 0; index < unusedGroups.length; index += 1) {
      const group = unusedGroups[index];
      try {
        await client.deleteGroup(group.id);
        deleted += 1;
        addLog("success", `Gruppe geloescht: ${group.name}.`);
      } catch (error) {
        addLog("error", `Gruppe konnte nicht geloescht werden: ${group.name}.`, formatError(error));
      }
      setProgress({ current: index + 1, total: unusedGroups.length, label: "Gruppen loeschen" });
    }

    setGroups((current) => current.filter((group) => !unusedGroups.some((removed) => removed.id === group.id)));
    setAllowDeleteGroups(false);
    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog("info", `Gruppen-Loeschung fertig. Geloescht: ${deleted}, Fehler: ${unusedGroups.length - deleted}.`);
  }

  async function scanFolders() {
    const tcProject = getProjectOrLog(effectiveProject, addLog);
    if (!tcProject) return [];

    setBusy("folders");
    setPlan([]);

    try {
      setProgress({ current: 0, total: 1, label: "Projekt laden" });
      const details = await client.getProject(tcProject.id);
      setProjectDetails(details);

      if (!details.rootId) {
        throw new Error("Das Projekt liefert keine rootId. Bitte API-Region und Projekt-ID pruefen.");
      }

      addLog("info", `Root-Ordner-ID: ${details.rootId}`, `Basis-URL: ${baseUrl}`);

      setProgress({ current: 0, total: 0, label: "Ordner scannen (0 gefunden)" });
      const tree = await client.listFolderTree(details.rootId, {
        onProgress: (scanned, found) => {
          setProgress({ current: scanned, total: scanned, label: `Ordner scannen (${found} gefunden)` });
        },
      });
      setFolders(tree);
      addLog("success", `${tree.length} Projektordner gescannt.`);

      if (tree.length === 0) {
        try {
          const debug = await client.fetchRawDebug(`/folders/${encodeURIComponent(details.rootId)}/items`);
          addLog(
            "warning",
            `Diagnose: Root-Ordner liefert 0 Elemente (HTTP ${debug.status}).`,
            `URL: ${debug.url}\n\n${debug.body.slice(0, 2000)}`
          );
        } catch (debugError) {
          addLog("warning", "Diagnose-Aufruf fehlgeschlagen.", formatError(debugError));
        }
      }

      return tree;
    } catch (error) {
      addLog("error", "Projektordner konnten nicht gescannt werden.", formatError(error));
      return [];
    } finally {
      setBusy("");
      setProgress({ current: 0, total: 0, label: "" });
    }
  }

  function planMissingFolders() {
    if (!currentMatrix) {
      addLog("warning", "Bitte zuerst eine Phase aus der Matrix waehlen.");
      return;
    }
    if (!effectiveTargetParentId) {
      addLog("warning", "Kein Zielordner bekannt. Bitte zuerst Ordner scannen oder manuell waehlen.");
      return;
    }

    const nextPlan = buildFolderCreationPlan(currentMatrix, folders, effectiveTargetParentId);
    setFolderPlan(nextPlan);

    const toCreate = nextPlan.filter((step) => step.status === "create").length;
    addLog("info", `Ordner-Plan erstellt: ${toCreate} von ${nextPlan.length} Ordnern fehlen unter "${targetParentLabel}".`);
  }

  async function createMissingFolders() {
    if (!allowCreateFolders) {
      addLog("warning", "Erstellen blockiert: Checkbox fuer Ordner-Erstellung ist nicht gesetzt.");
      return;
    }
    if (!currentMatrix || folderPlan.length === 0) {
      addLog("warning", "Bitte zuerst einen Ordner-Plan erstellen.");
      return;
    }
    if (!effectiveTargetParentId) return;

    setBusy("create-folders");
    const toCreate = folderPlan.filter((step) => step.status === "create").length;
    setProgress({ current: 0, total: toCreate, label: "Ordner erstellen" });

    const resolvedIdByDepth = new Map<number, string>();
    let created = 0;
    let failed = 0;
    let progressDone = 0;
    const newFolders: TCFolder[] = [];

    for (const step of folderPlan) {
      const parentId = step.depth === 0 ? effectiveTargetParentId : resolvedIdByDepth.get(step.depth - 1);

      if (!parentId) {
        addLog("error", `Uebersprungen: ${step.path} (uebergeordneter Ordner wurde nicht erstellt).`);
        failed += 1;
        continue;
      }

      if (step.status === "existing" && step.folderId) {
        resolvedIdByDepth.set(step.depth, step.folderId);
        continue;
      }

      try {
        const folder = await client.createFolder(parentId, step.name);
        const parentPath = newFolders.find((f) => f.id === parentId)?.path
          ?? folders.find((f) => f.id === parentId)?.path
          ?? "";
        const withPath: TCFolder = { ...folder, path: parentPath ? `${parentPath}/${folder.name}` : folder.name, depth: step.depth };
        newFolders.push(withPath);
        resolvedIdByDepth.set(step.depth, folder.id);
        created += 1;
        addLog("success", `Ordner erstellt: ${step.path}.`);
      } catch (error) {
        failed += 1;
        addLog("error", `Ordner konnte nicht erstellt werden: ${step.path}.`, formatError(error));
      }

      progressDone += 1;
      setProgress({ current: progressDone, total: toCreate, label: "Ordner erstellen" });
    }

    if (newFolders.length > 0) {
      setFolders((current) => [...current, ...newFolders]);
    }

    setAllowCreateFolders(false);
    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog("info", `Ordner-Erstellung fertig. Erstellt: ${created}, Fehler: ${failed}.`);
  }

  async function createDryRun() {
    if (!currentMatrix) {
      addLog("warning", "Bitte zuerst eine Phase aus der Matrix waehlen.");
      return;
    }

    setBusy("dry-run");
    try {
      const activeGroups = groups.length > 0 ? groups : await loadGroups();
      const activeFolders = folders.length > 0 ? folders : await scanFolders();
      const nextPlan = buildPermissionPlan(currentMatrix, activeGroups, activeFolders);
      setPlan(nextPlan);

      const ready = nextPlan.filter((item) => item.status === "ready").length;
      const missingFolders = nextPlan.filter((item) => item.status === "missing-folder").length;
      const missingGroups = nextPlan.filter((item) => item.status === "missing-groups").length;
      addLog("info", `Dry-Run erstellt: ${ready} bereit, ${missingFolders} Ordner fehlen, ${missingGroups} mit fehlenden Gruppen.`);
    } finally {
      setBusy("");
    }
  }

  async function applyPermissions() {
    if (!allowApply) {
      addLog("warning", "Aktivieren blockiert: Checkbox fuer Aenderungen ist nicht gesetzt.");
      return;
    }

    const readyItems = plan.filter((item) => item.status === "ready" && item.folderId);
    if (readyItems.length === 0) {
      addLog("warning", "Keine bereiten Ordner im Dry-Run.");
      return;
    }

    setBusy("apply");
    setProgress({ current: 0, total: readyItems.length, label: "Berechtigungen anwenden" });

    let applied = 0;
    for (let index = 0; index < readyItems.length; index += 1) {
      const item = readyItems[index];
      const oversized = (["READ", "FULL_ACCESS"] as const).filter((level) => item.acl[level].length > 100);

      if (oversized.length > 0) {
        addLog(
          "warning",
          `Uebersprungen: ${item.folderPath} hat mehr als 100 Gruppen bei ${oversized.join(", ")} (API-Limit).`,
          "Diese Zeile muss manuell in mehreren Schritten in Trimble Connect gepflegt werden."
        );
        setProgress({ current: index + 1, total: readyItems.length, label: "Berechtigungen anwenden" });
        continue;
      }

      try {
        await client.updateFolderPermissions(item.folderId!, item.acl, false);
        applied += 1;
      } catch (error) {
        addLog("error", `Berechtigungen fehlgeschlagen: ${item.folderPath}.`, formatError(error));
      }
      setProgress({ current: index + 1, total: readyItems.length, label: "Berechtigungen anwenden" });
    }

    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog("success", `${applied} Ordner aktualisiert.`);
  }

  const isBusy = busy.length > 0;

  if (accessState !== "admin") {
    return (
      <div className="app-shell">
        <header className="topbar">
          <div>
            <div className="eyebrow">Trimble Connect Web Extension</div>
            <h1>Team & Berechtigungsmanager</h1>
          </div>
        </header>
        <div className="access-gate">
          {accessState === "checking" ? (
            <p>Berechtigung wird geprueft ...</p>
          ) : (
            <>
              <h2>Keine Berechtigung</h2>
              <p>Diese Extension steht nur Projekt-Administrator:innen zur Verfuegung.</p>
              {(accessRole || accessError) && (
                <p className="access-detail">
                  {accessRole && `Erkannte Rolle: ${accessRole}`}
                  {accessError && `Pruefung fehlgeschlagen: ${accessError}`}
                </p>
              )}
            </>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <div className="eyebrow">Trimble Connect Web Extension</div>
          <h1>Team & Berechtigungsmanager</h1>
        </div>
        <div className="status-stack">
          <span className={connected ? "status ok" : "status muted"}>{connected ? "Workspace verbunden" : "Standalone"}</span>
          <span className={accessToken || manualToken ? "status ok" : "status warn"}>{accessToken || manualToken ? "Token bereit" : "Token fehlt"}</span>
          <a
            className="help-button"
            href="/hilfe.html"
            target="_blank"
            rel="noopener noreferrer"
            title="Anleitung oeffnen"
            aria-label="Anleitung oeffnen"
          >
            ?
          </a>
        </div>
      </header>

      <section className="context-bar">
        <div className="context-item wide">
          <span>Projekt</span>
          <strong>{effectiveProject?.name ?? "Nicht verbunden"}</strong>
          {effectiveProject?.id && <code>{effectiveProject.id}</code>}
        </div>
        <label className="field compact">
          <span>Region</span>
          <select
            value={region}
            onChange={(event) => {
              setRegion(event.target.value as ApiRegion);
              setRegionTouched(true);
            }}
          >
            <option value="eu">EU</option>
            <option value="us">US</option>
            <option value="uk">UK</option>
            <option value="ap">AP</option>
            <option value="ap2">AP2</option>
            <option value="custom">Custom</option>
          </select>
        </label>
        {region === "custom" && (
          <label className="field grow">
            <span>API Base URL</span>
            <input value={customBaseUrl} onChange={(event) => setCustomBaseUrl(event.target.value)} />
          </label>
        )}
        {!project && (
          <>
            <label className="field grow">
              <span>Projekt-ID</span>
              <input value={manualProjectId} onChange={(event) => setManualProjectId(event.target.value)} />
            </label>
            <label className="field grow">
              <span>Access Token</span>
              <input type="password" value={manualToken} onChange={(event) => setManualToken(event.target.value)} />
            </label>
          </>
        )}
      </section>

      {connectionError && <div className="banner warning">Workspace API: {connectionError}</div>}
      {hostName && <div className="banner info">Host: {hostName}</div>}

      <nav className="tabs" aria-label="Hauptbereiche">
        <button className={activeTab === "teams" ? "active" : ""} onClick={() => setActiveTab("teams")}>Teams</button>
        <button className={activeTab === "folders" ? "active" : ""} onClick={() => setActiveTab("folders")}>Ordner erstellen</button>
        <button className={activeTab === "permissions" ? "active" : ""} onClick={() => setActiveTab("permissions")}>Berechtigungen</button>
        <button className={activeTab === "log" ? "active" : ""} onClick={() => setActiveTab("log")}>Protokoll</button>
        <button
          className="tab-action"
          onClick={async () => {
            const { downloadMatrixTemplate } = await import("./excel/template");
            downloadMatrixTemplate();
            addLog("info", "Matrix-Vorlage heruntergeladen.");
          }}
        >
          Vorlage exportieren
        </button>
      </nav>

      <main className="content">
        {activeTab === "teams" && (
          <TeamsTab
            busy={isBusy}
            workbook={workbook}
            groups={groups}
            missingTeamNames={missingTeamNames}
            unusedGroups={unusedGroups}
            newGroupName={newGroupName}
            setNewGroupName={setNewGroupName}
            allowDeleteGroups={allowDeleteGroups}
            setAllowDeleteGroups={setAllowDeleteGroups}
            onWorkbook={handleWorkbook}
            onLoadGroups={loadGroups}
            onCreateGroup={createSingleGroup}
            onImportGroups={importMissingGroups}
            onDeleteUnusedGroups={deleteUnusedGroups}
          />
        )}

        {activeTab === "folders" && (
          <FoldersTab
            busy={isBusy}
            workbook={workbook}
            selectedSheet={selectedSheet}
            setSelectedSheet={setSelectedSheet}
            currentMatrix={currentMatrix}
            folders={folders}
            projectDetails={projectDetails}
            progress={progress}
            onWorkbook={handleWorkbook}
            onScanFolders={scanFolders}
            targetParentFilter={targetParentFilter}
            setTargetParentFilter={setTargetParentFilter}
            targetParentCandidates={targetParentCandidates}
            targetParentId={targetParentId}
            setTargetParentId={setTargetParentId}
            targetParentLabel={targetParentLabel}
            folderPlan={folderPlan}
            allowCreateFolders={allowCreateFolders}
            setAllowCreateFolders={setAllowCreateFolders}
            onPlanFolders={planMissingFolders}
            onCreateFolders={createMissingFolders}
          />
        )}

        {activeTab === "permissions" && (
          <PermissionsTab
            busy={isBusy}
            workbook={workbook}
            selectedSheet={selectedSheet}
            setSelectedSheet={setSelectedSheet}
            currentMatrix={currentMatrix}
            folders={folders}
            groups={groups}
            projectDetails={projectDetails}
            plan={plan}
            planStats={planStats}
            progress={progress}
            allowApply={allowApply}
            setAllowApply={setAllowApply}
            onWorkbook={handleWorkbook}
            onLoadGroups={loadGroups}
            onScanFolders={scanFolders}
            onDryRun={createDryRun}
            onApply={applyPermissions}
          />
        )}

        {activeTab === "log" && <LogPanel logs={logs} />}
      </main>

      {progress.total > 0 && (
        <footer className="progress-footer">
          <span>{progress.label}</span>
          <progress value={progress.current} max={progress.total} />
          <strong>{progress.current}/{progress.total}</strong>
        </footer>
      )}
    </div>
  );
}

interface TeamsTabProps {
  busy: boolean;
  workbook: WorkbookModel | null;
  groups: TCGroup[];
  missingTeamNames: string[];
  unusedGroups: TCGroup[];
  newGroupName: string;
  setNewGroupName: (value: string) => void;
  allowDeleteGroups: boolean;
  setAllowDeleteGroups: (value: boolean) => void;
  onWorkbook: (file: File) => void;
  onLoadGroups: () => void;
  onCreateGroup: () => void;
  onImportGroups: () => void;
  onDeleteUnusedGroups: () => void;
}

function TeamsTab(props: TeamsTabProps) {
  return (
    <div className="panel-grid">
      <section className="panel">
        <h2>Matrix</h2>
        <FileInput disabled={props.busy} onFile={props.onWorkbook} />
        <MetricGrid
          metrics={[
            ["Phasen", props.workbook?.sheetNames.length ?? 0],
            ["Matrix-Teams", props.workbook?.teamNames.length ?? 0],
            ["Fehlend", props.missingTeamNames.length],
          ]}
        />
        {props.missingTeamNames.length > 0 && (
          <div className="compact-list">
            {props.missingTeamNames.slice(0, 24).map((name) => <span key={name}>{name}</span>)}
            {props.missingTeamNames.length > 24 && <span>+{props.missingTeamNames.length - 24}</span>}
          </div>
        )}
      </section>

      <section className="panel">
        <h2>Gruppen</h2>
        <div className="inline-actions">
          <button onClick={props.onLoadGroups} disabled={props.busy}>Laden</button>
          <button className="primary" onClick={props.onImportGroups} disabled={props.busy || !props.workbook}>
            Fehlende erstellen
          </button>
        </div>
        <div className="input-row">
          <input
            value={props.newGroupName}
            onChange={(event) => props.setNewGroupName(event.target.value)}
            placeholder="Teamname"
          />
          <button onClick={props.onCreateGroup} disabled={props.busy || !props.newGroupName.trim()}>Erstellen</button>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Name</th>
                <th>Mitglieder</th>
              </tr>
            </thead>
            <tbody>
              {props.groups.map((group) => (
                <tr key={group.id}>
                  <td>{group.name}</td>
                  <td>{group.usersCount ?? "-"}</td>
                </tr>
              ))}
              {props.groups.length === 0 && (
                <tr>
                  <td colSpan={2} className="empty-cell">Keine Gruppen geladen</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel span-2">
        <h2>Ungenutzte Gruppen</h2>
        <p>
          Gruppen, die in der geladenen Excel-Matrix (ueber alle Phasen) namentlich nicht mehr
          vorkommen. Loeschen betrifft die Gruppe global in Trimble Connect, auch wenn sie
          ausserhalb dieser Matrix noch irgendwo Berechtigungen haelt.
        </p>
        <MetricGrid metrics={[["Ungenutzt", props.unusedGroups.length]]} />
        {props.unusedGroups.length > 0 && (
          <div className="compact-list">
            {props.unusedGroups.slice(0, 24).map((group) => <span key={group.id}>{group.name}</span>)}
            {props.unusedGroups.length > 24 && <span>+{props.unusedGroups.length - 24}</span>}
          </div>
        )}
        <div className="apply-row">
          <label className="check">
            <input
              type="checkbox"
              checked={props.allowDeleteGroups}
              onChange={(event) => props.setAllowDeleteGroups(event.target.checked)}
            />
            <span>Loeschen wirklich durchfuehren</span>
          </label>
          <button
            className="danger"
            onClick={props.onDeleteUnusedGroups}
            disabled={props.busy || !props.allowDeleteGroups || props.unusedGroups.length === 0}
          >
            Ungenutzte Gruppen loeschen
          </button>
        </div>
      </section>
    </div>
  );
}

interface FoldersTabProps {
  busy: boolean;
  workbook: WorkbookModel | null;
  selectedSheet: string;
  setSelectedSheet: (sheet: string) => void;
  currentMatrix: PermissionMatrix | null;
  folders: TCFolder[];
  projectDetails: TCProject | null;
  progress: ProgressState;
  onWorkbook: (file: File) => void;
  onScanFolders: () => void;
  targetParentFilter: string;
  setTargetParentFilter: (value: string) => void;
  targetParentCandidates: TCFolder[];
  targetParentId: string;
  setTargetParentId: (value: string) => void;
  targetParentLabel: string;
  folderPlan: FolderCreationStep[];
  allowCreateFolders: boolean;
  setAllowCreateFolders: (value: boolean) => void;
  onPlanFolders: () => void;
  onCreateFolders: () => void;
}

function FoldersTab(props: FoldersTabProps) {
  const toCreate = props.folderPlan.filter((step) => step.status === "create").length;

  return (
    <div className="panel-grid wide">
      <section className="panel">
        <h2>Phase</h2>
        <FileInput disabled={props.busy} onFile={props.onWorkbook} />
        <label className="field">
          <span>Tabellenblatt</span>
          <select value={props.selectedSheet} onChange={(event) => props.setSelectedSheet(event.target.value)}>
            {props.workbook?.sheetNames.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
          </select>
        </label>
        <MetricGrid
          metrics={[
            ["Ordner in Matrix", props.currentMatrix?.stats.folders ?? 0],
            ["Gescannt", props.folders.length],
          ]}
        />
        <div className="inline-actions">
          <button onClick={props.onScanFolders} disabled={props.busy}>Ordner scannen</button>
        </div>
      </section>

      <section className="panel">
        <h2>Zielordner</h2>
        <p className="hint">
          Wohin der Phasenordner (Blattname <strong>{props.currentMatrix?.sheetName ?? "–"}</strong>) kommt,
          falls er noch fehlt.
        </p>
        <label className="field">
          <span>Ordner suchen</span>
          <input
            value={props.targetParentFilter}
            onChange={(event) => props.setTargetParentFilter(event.target.value)}
            placeholder="Suche (leer = Projekt-Root)"
          />
        </label>
        {props.targetParentCandidates.length > 0 && (
          <div className="table-wrap" style={{ maxHeight: 160 }}>
            <table>
              <tbody>
                {props.targetParentCandidates.map((folder) => (
                  <tr
                    key={folder.id}
                    className={`selectable${folder.id === props.targetParentId ? " selected" : ""}`}
                    onClick={() => props.setTargetParentId(folder.id)}
                  >
                    <td>{folder.path}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <p className="hint">Aktueller Zielordner: <strong>{props.targetParentLabel || "unbekannt (zuerst Ordner scannen)"}</strong></p>
      </section>

      <section className="panel span-2">
        <div className="panel-title-row">
          <h2>Ordner-Plan</h2>
          <div className="summary-pills">
            <span className="pill warn">{toCreate} werden erstellt</span>
            <span className="pill ok">{props.folderPlan.length - toCreate} vorhanden</span>
          </div>
        </div>
        <div className="inline-actions">
          <button className="primary" onClick={props.onPlanFolders} disabled={props.busy || !props.currentMatrix}>
            Ordner-Plan erstellen
          </button>
        </div>
        <div className="apply-row">
          <label className="check">
            <input
              type="checkbox"
              checked={props.allowCreateFolders}
              onChange={(event) => props.setAllowCreateFolders(event.target.checked)}
            />
            <span>Ordner wirklich erstellen</span>
          </label>
          <button
            className="danger"
            onClick={props.onCreateFolders}
            disabled={props.busy || !props.allowCreateFolders || toCreate === 0}
          >
            Ordner erstellen
          </button>
        </div>
        <div className="table-wrap plan-table">
          <table>
            <thead>
              <tr><th>Status</th><th>Pfad</th></tr>
            </thead>
            <tbody>
              {props.folderPlan.map((step) => (
                <tr key={`${step.rowNumber}-${step.path}`}>
                  <td><span className={`state ${step.status === "create" ? "missing-folder" : "ready"}`}>{step.status === "create" ? "Wird erstellt" : "Vorhanden"}</span></td>
                  <td>{step.path}</td>
                </tr>
              ))}
              {props.folderPlan.length === 0 && (
                <tr>
                  <td colSpan={2} className="empty-cell">Noch kein Ordner-Plan</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

interface PermissionsTabProps {
  busy: boolean;
  workbook: WorkbookModel | null;
  selectedSheet: string;
  setSelectedSheet: (sheet: string) => void;
  currentMatrix: PermissionMatrix | null;
  folders: TCFolder[];
  groups: TCGroup[];
  projectDetails: TCProject | null;
  plan: PermissionPlanItem[];
  planStats: { ready: number; missingFolders: number; missingGroups: number };
  progress: ProgressState;
  allowApply: boolean;
  setAllowApply: (value: boolean) => void;
  onWorkbook: (file: File) => void;
  onLoadGroups: () => void;
  onScanFolders: () => void;
  onDryRun: () => void;
  onApply: () => void;
}

function PermissionsTab(props: PermissionsTabProps) {
  return (
    <div className="panel-grid wide">
      <section className="panel">
        <h2>Phase</h2>
        <FileInput disabled={props.busy} onFile={props.onWorkbook} />
        <label className="field">
          <span>Tabellenblatt</span>
          <select value={props.selectedSheet} onChange={(event) => props.setSelectedSheet(event.target.value)}>
            {props.workbook?.sheetNames.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
          </select>
        </label>
        <MetricGrid
          metrics={[
            ["Ordner", props.currentMatrix?.stats.folders ?? 0],
            ["Teams", props.currentMatrix?.stats.teams ?? 0],
            ["Voll", props.currentMatrix?.stats.fullAccess ?? 0],
            ["Lesen", props.currentMatrix?.stats.read ?? 0],
            ["Kein", props.currentMatrix?.stats.noAccess ?? 0],
          ]}
        />
      </section>

      <section className="panel">
        <h2>Projektabgleich</h2>
        <MetricGrid
          metrics={[
            ["Gruppen", props.groups.length],
            ["Ordner", props.folders.length],
            ["Root", props.projectDetails?.rootId ? 1 : 0],
          ]}
        />
        <div className="inline-actions">
          <button onClick={props.onLoadGroups} disabled={props.busy}>Gruppen laden</button>
          <button onClick={props.onScanFolders} disabled={props.busy}>Ordner scannen</button>
          <button className="primary" onClick={props.onDryRun} disabled={props.busy || !props.currentMatrix}>Dry-Run</button>
        </div>
      </section>

      <section className="panel span-2">
        <div className="panel-title-row">
          <h2>Dry-Run</h2>
          <div className="summary-pills">
            <span className="pill ok">{props.planStats.ready} bereit</span>
            <span className="pill warn">{props.planStats.missingFolders} Ordner fehlen</span>
            <span className="pill warn">{props.planStats.missingGroups} Gruppen fehlen</span>
          </div>
        </div>
        <div className="apply-row">
          <label className="check">
            <input
              type="checkbox"
              checked={props.allowApply}
              onChange={(event) => props.setAllowApply(event.target.checked)}
            />
            <span>Aenderungen wirklich anwenden</span>
          </label>
          <button className="danger" onClick={props.onApply} disabled={props.busy || !props.allowApply || props.planStats.ready === 0}>
            Berechtigungen anwenden
          </button>
        </div>
        <div className="table-wrap plan-table">
          <table>
            <thead>
              <tr>
                <th>Status</th>
                <th>Ordnerpfad</th>
                <th>Voll</th>
                <th>Lesen</th>
                <th>Kein</th>
                <th>Fehlende Gruppen</th>
              </tr>
            </thead>
            <tbody>
              {props.plan.slice(0, 120).map((item) => (
                <tr key={`${item.rowNumber}-${item.folderPath}`}>
                  <td><span className={`state ${item.status}`}>{statusLabel(item.status)}</span></td>
                  <td>{item.folderPath}</td>
                  <td>{item.counts.fullAccess}</td>
                  <td>{item.counts.read}</td>
                  <td>{item.counts.noAccess}</td>
                  <td>{item.missingGroups.slice(0, 3).join(", ")}{item.missingGroups.length > 3 ? " ..." : ""}</td>
                </tr>
              ))}
              {props.plan.length === 0 && (
                <tr>
                  <td colSpan={6} className="empty-cell">Noch kein Dry-Run</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function LogPanel({ logs }: { logs: LogEntry[] }) {
  return (
    <section className="panel log-panel">
      <h2>Protokoll</h2>
      <div className="log-list">
        {logs.map((entry) => (
          <article key={entry.id} className={`log-entry ${entry.level}`}>
            <time>{entry.createdAt}</time>
            <strong>{entry.message}</strong>
            {entry.detail && <pre>{entry.detail}</pre>}
          </article>
        ))}
        {logs.length === 0 && <div className="empty-cell">Noch keine Eintraege</div>}
      </div>
    </section>
  );
}

function FileInput({ disabled, onFile }: { disabled: boolean; onFile: (file: File) => void }) {
  return (
    <label className="file-drop">
      <input
        type="file"
        accept=".xlsx,.xls"
        disabled={disabled}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
          event.currentTarget.value = "";
        }}
      />
      <span>Excel-Matrix laden</span>
    </label>
  );
}

function MetricGrid({ metrics }: { metrics: Array<[string, number | string]> }) {
  return (
    <div className="metric-grid">
      {metrics.map(([label, value]) => (
        <div key={label} className="metric">
          <span>{label}</span>
          <strong>{value}</strong>
        </div>
      ))}
    </div>
  );
}

function requireProject(project: TCProject | null): TCProject {
  if (!project?.id) {
    throw new Error("Kein Projekt verbunden. Bitte in Trimble Connect oeffnen oder Projekt-ID manuell setzen.");
  }
  return project;
}

function getProjectOrLog(
  project: TCProject | null,
  addLog: (level: LogEntry["level"], message: string, detail?: string) => void
): TCProject | null {
  try {
    return requireProject(project);
  } catch (error) {
    addLog("error", "Kein Projekt verbunden.", formatError(error));
    return null;
  }
}

function formatError(error: unknown): string {
  if (error instanceof TrimbleApiError) {
    return [error.message, error.body].filter(Boolean).join("\n").slice(0, 1200);
  }
  if (error instanceof Error) return error.message;
  return String(error);
}

function uniqueGroups(groups: TCGroup[]): TCGroup[] {
  const byName = new Map<string, TCGroup>();
  groups.forEach((group) => byName.set(normalizeLookup(group.name), group));
  return Array.from(byName.values()).sort((a, b) => a.name.localeCompare(b.name, "de"));
}

function statusLabel(status: PermissionPlanItem["status"]): string {
  if (status === "ready") return "Bereit";
  if (status === "missing-folder") return "Ordner fehlt";
  return "Gruppe fehlt";
}

function guessRegion(location: string): ApiRegion | null {
  const value = normalizeLookup(location);
  if (value.includes("europe") || value.includes("eu")) return "eu";
  if (value.includes("uk") || value.includes("united kingdom")) return "uk";
  if (value.includes("asia") || value.includes("apac")) return "ap";
  if (value.includes("northamerica") || value.includes("us")) return "us";
  return null;
}
