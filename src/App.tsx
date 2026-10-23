import { useEffect, useMemo, useRef, useState } from "react";
import { API_BASE_URLS, TrimbleApiError, TrimbleClient } from "./api/trimble";
import { useApi } from "./hooks/useApi";
import { buildPermissionPlan } from "./permissions/planner";
import { buildFolderCreationPlan } from "./permissions/folderPlanner";
import type { ExportSheet } from "./excel/export";
import type {
  ApiRegion,
  FolderCreationStep,
  LogEntry,
  PermissionMatrix,
  PermissionPlanItem,
  ProgressState,
  TCFolder,
  TCGroup,
  TCUser,
  TCProject,
  WorkbookModel,
} from "./types";
import { normalizeLookup } from "./utils/text";
import { loadAppliedMap, loadStoredWorkbook, saveAppliedMap, saveStoredWorkbook } from "./utils/workbookStorage";
import "./index.css";

type TabId = "teams" | "folders" | "permissions" | "log" | "access";

const HISTORY_LIMIT = 20;

interface HistorySnapshot {
  workbook: WorkbookModel | null;
  selectedSheet: string;
  groups: TCGroup[];
  deleteMarks: Record<string, boolean>;
  folders: TCFolder[];
  folderPlan: FolderCreationStep[];
  plan: PermissionPlanItem[];
  targetParentId: string;
}

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
  const [deleteMarks, setDeleteMarks] = useState<Record<string, boolean>>({});
  const [folders, setFolders] = useState<TCFolder[]>([]);
  const [projectDetails, setProjectDetails] = useState<TCProject | null>(null);
  const [newGroupName, setNewGroupName] = useState("");
  const [plan, setPlan] = useState<PermissionPlanItem[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [busy, setBusy] = useState("");
  const [progress, setProgress] = useState<ProgressState>({ current: 0, total: 0, label: "" });
  const [allowApply, setAllowApply] = useState(false);
  const [allowPartial, setAllowPartial] = useState(true);
  const [targetParentId, setTargetParentId] = useState("");
  const [targetParentFilter, setTargetParentFilter] = useState("");
  const [folderPlan, setFolderPlan] = useState<FolderCreationStep[]>([]);
  const [allowCreateFolders, setAllowCreateFolders] = useState(false);
  const [extraMarks, setExtraMarks] = useState<Record<string, boolean>>({});
  const [allSheets, setAllSheets] = useState(false);
  const [excludedSheets, setExcludedSheets] = useState<Record<string, boolean>>({});
  const cancelRef = useRef(false);

  const snapshot: HistorySnapshot = { workbook, selectedSheet, groups, deleteMarks, folders, folderPlan, plan, targetParentId };
  const lastSnapshotRef = useRef<HistorySnapshot>(snapshot);
  const pastRef = useRef<HistorySnapshot[]>([]);
  const futureRef = useRef<HistorySnapshot[]>([]);
  const skipHistoryRef = useRef(false);
  const [historySizes, setHistorySizes] = useState({ past: 0, future: 0 });

  // Jede Aenderung (auch eine ganze Operation mit mehreren Teilschritten) wird zu einem Undo-Schritt.
  useEffect(() => {
    if (busy) return;
    const last = lastSnapshotRef.current;
    const changed = (Object.keys(snapshot) as Array<keyof HistorySnapshot>).some((key) => snapshot[key] !== last[key]);
    if (!changed) return;
    if (skipHistoryRef.current) {
      skipHistoryRef.current = false;
    } else {
      pastRef.current = [...pastRef.current, last].slice(-HISTORY_LIMIT);
      futureRef.current = [];
      setHistorySizes({ past: pastRef.current.length, future: 0 });
    }
    lastSnapshotRef.current = snapshot;
  });

  function applySnapshot(target: HistorySnapshot) {
    skipHistoryRef.current = true;
    setWorkbook(target.workbook);
    setSelectedSheet(target.selectedSheet);
    setGroups(target.groups);
    setDeleteMarks(target.deleteMarks);
    setFolders(target.folders);
    setFolderPlan(target.folderPlan);
    setPlan(target.plan);
    setTargetParentId(target.targetParentId);
  }

  function undo() {
    const previous = pastRef.current[pastRef.current.length - 1];
    if (!previous || busy) return;
    pastRef.current = pastRef.current.slice(0, -1);
    futureRef.current = [...futureRef.current, lastSnapshotRef.current].slice(-HISTORY_LIMIT);
    setHistorySizes({ past: pastRef.current.length, future: futureRef.current.length });
    applySnapshot(previous);
  }

  function redo() {
    const next = futureRef.current[futureRef.current.length - 1];
    if (!next || busy) return;
    futureRef.current = futureRef.current.slice(0, -1);
    pastRef.current = [...pastRef.current, lastSnapshotRef.current].slice(-HISTORY_LIMIT);
    setHistorySizes({ past: pastRef.current.length, future: futureRef.current.length });
    applySnapshot(next);
  }

  function toggleMark(groupId: string) {
    setDeleteMarks((current) => ({ ...current, [groupId]: !markedIds.has(groupId) }));
  }

  function keepAllGroups() {
    setDeleteMarks(Object.fromEntries(groups.map((group) => [group.id, false])));
  }

  async function commitMarkedGroups() {
    const toDelete = groups.filter((group) => markedIds.has(group.id));
    if (toDelete.length === 0) return;
    const names = toDelete.map((group) => group.name).join("\n");
    if (!window.confirm(`${toDelete.length} Team(s) in Trimble Connect endgueltig loeschen?\n\n${names}`)) return;

    setBusy("delete-groups");
    setProgress({ current: 0, total: toDelete.length, label: "Teams loeschen" });

    const deletedIds = new Set<string>();
    for (let index = 0; index < toDelete.length; index += 1) {
      const group = toDelete[index];
      try {
        await client.deleteGroup(group.id);
        deletedIds.add(group.id);
        addLog("success", `Gruppe geloescht: ${group.name}.`);
      } catch (error) {
        addLog("error", `Gruppe konnte nicht geloescht werden: ${group.name}.`, formatError(error));
      }
      setProgress({ current: index + 1, total: toDelete.length, label: "Teams loeschen" });
    }

    setGroups((current) => current.filter((group) => !deletedIds.has(group.id)));
    setDeleteMarks({});
    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog("info", `Team-Loeschung fertig. Geloescht: ${deletedIds.size}, Fehler: ${toDelete.length - deletedIds.size}.`);
  }

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

  // Lange Laeufe (alle Blaetter) ueberdauern die Token-Laufzeit: das Token wird bei jeder Anfrage
  // aus dem Ref gelesen, bei einem 401 neu angefordert.
  const accessTokenRef = useRef(accessToken);
  accessTokenRef.current = accessToken;

  const client = useMemo(
    () =>
      new TrimbleClient({
        baseUrl,
        delayMs: 250,
        getToken: async () => {
          if (manualToken.trim()) return manualToken.trim();
          if (accessTokenRef.current) return accessTokenRef.current;
          return refreshAccessToken();
        },
        refreshToken: async () => {
          if (manualToken.trim()) return manualToken.trim();
          return refreshAccessToken();
        },
      }),
    [baseUrl, manualToken, refreshAccessToken]
  );

  const [accessState, setAccessState] = useState<"checking" | "admin" | "denied">("checking");
  const [accessRole, setAccessRole] = useState("");
  const [accessError, setAccessError] = useState("");
  const [projectAdmins, setProjectAdmins] = useState<TCUser[]>([]);
  const [allowedUsers, setAllowedUsers] = useState<TCUser[]>([]);
  const [currentUser, setCurrentUser] = useState<TCUser | null>(null);
  const [removedNotes, setRemovedNotes] = useState<{ name: string; email: string; at: string }[]>([]);
  const [accessListWarning, setAccessListWarning] = useState("");

  const removedKey = `tc-access-removed:v1:${effectiveProject?.id ?? ""}`;

  // Liest Admins und freigegebene Nutzer neu ein. Mit autoRemove werden freigegebene Nutzer, die
  // nicht mehr Projekt-Administrator sind, aus der Liste entfernt und als Hinweis vermerkt.
  // Ergebnis "denied": Der aktuelle Nutzer steht nicht auf der (nicht leeren) Liste.
  async function refreshAccessList(autoRemove: boolean): Promise<"allowed" | "denied"> {
    const projectId = effectiveProject?.id;
    if (!projectId) return "allowed";

    try {
      const [me, users, access] = await Promise.all([
        client.getCurrentUser(),
        client.listProjectUsers(projectId),
        client.getAccessGroup(projectId),
      ]);
      const admins = users.filter((user) => user.role?.includes("ADMIN"));
      let members = access.members;

      if (autoRemove && users.length > 0) {
        const adminIds = new Set(admins.map((user) => user.id));
        const demoted = members.filter((member) => !adminIds.has(member.id));
        const removed: { name: string; email: string; at: string }[] = [];
        for (const member of demoted) {
          try {
            await client.removeAccessUser(projectId, member.id);
            removed.push({ name: member.name, email: member.email, at: new Date().toISOString() });
            members = members.filter((entry) => entry.id !== member.id);
          } catch (error) {
            addLog("warning", `Zugriff von ${member.name} konnte nicht entfernt werden.`, formatError(error));
          }
        }
        if (removed.length > 0) {
          let previous: typeof removed = [];
          try {
            previous = JSON.parse(window.localStorage.getItem(removedKey) ?? "[]");
          } catch {
            previous = [];
          }
          const merged = [...removed, ...previous].slice(0, 20);
          try {
            window.localStorage.setItem(removedKey, JSON.stringify(merged));
          } catch {
            // Ignore.
          }
          setRemovedNotes(merged);
          addLog("warning", `Zugriff automatisch entfernt (nicht mehr Admin): ${removed.map((entry) => entry.name).join(", ")}.`);
        }
      }

      setCurrentUser(me);
      setProjectAdmins(admins);
      setAllowedUsers(members);
      setAccessListWarning("");
      return members.length > 0 && !members.some((member) => member.id === me.id) ? "denied" : "allowed";
    } catch (error) {
      // Nicht lesbar: Admin-Rolle genuegt, damit sich niemand aussperrt.
      setAccessListWarning(`Zugriffsliste nicht lesbar: ${formatError(error)}`);
      return "allowed";
    }
  }

  useEffect(() => {
    try {
      setRemovedNotes(JSON.parse(window.localStorage.getItem(removedKey) ?? "[]"));
    } catch {
      setRemovedNotes([]);
    }
  }, [removedKey]);

  async function addAccessUser(userId: string) {
    const projectId = effectiveProject?.id;
    if (!projectId) return;
    setBusy("access");
    try {
      // Erste Freigabe: den aktuellen Nutzer mit aufnehmen, damit er sich nicht aussperrt.
      if (allowedUsers.length === 0 && currentUser && currentUser.id !== userId) {
        await client.addAccessUser(projectId, currentUser.id);
      }
      await client.addAccessUser(projectId, userId);
      await refreshAccessList(false);
      addLog("success", "Administrator fuer die App freigegeben.");
    } catch (error) {
      addLog("error", "Freigabe fehlgeschlagen.", formatError(error));
    } finally {
      setBusy("");
    }
  }

  async function removeAccessUser(userId: string) {
    const projectId = effectiveProject?.id;
    if (!projectId) return;
    setBusy("access");
    try {
      await client.removeAccessUser(projectId, userId);
      await refreshAccessList(false);
      addLog("success", "Freigabe entfernt.");
    } catch (error) {
      addLog("error", "Freigabe konnte nicht entfernt werden.", formatError(error));
    } finally {
      setBusy("");
    }
  }

  async function reloadAccess() {
    setBusy("access");
    try {
      await refreshAccessList(true);
    } finally {
      setBusy("");
    }
  }

  useEffect(() => {
    if (!effectiveProject?.id) {
      setAccessState("checking");
      return;
    }

    let cancelled = false;
    setAccessState("checking");
    setAccessError("");

    client
      .getProjectRole(effectiveProject.id)
      .then(async (role) => {
        if (cancelled) return;
        setAccessRole(role ?? "");
        if (!role?.includes("ADMIN")) {
          setAccessState("denied");
          return;
        }
        const listed = await refreshAccessList(true);
        if (cancelled) return;
        setAccessState(listed === "denied" ? "denied" : "admin");
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

  const activeMatrices = useMemo(() => {
    if (!workbook) return [];
    if (allSheets) return workbook.matrices.filter((matrix) => !excludedSheets[matrix.sheetName]);
    return currentMatrix ? [currentMatrix] : [];
  }, [allSheets, currentMatrix, excludedSheets, workbook]);

  const scopeStats = useMemo(
    () => ({
      sheets: activeMatrices.length,
      folders: activeMatrices.reduce((sum, matrix) => sum + matrix.stats.folders, 0),
      teams: new Set(activeMatrices.flatMap((matrix) => matrix.teams)).size,
      fullAccess: activeMatrices.reduce((sum, matrix) => sum + matrix.stats.fullAccess, 0),
      read: activeMatrices.reduce((sum, matrix) => sum + matrix.stats.read, 0),
      noAccess: activeMatrices.reduce((sum, matrix) => sum + matrix.stats.noAccess, 0),
    }),
    [activeMatrices]
  );

  // Umfang gewechselt: bestehende Plaene gehoeren zum alten Umfang und werden verworfen.
  function changeScope(nextAll: boolean, nextExcluded: Record<string, boolean>) {
    setAllSheets(nextAll);
    setExcludedSheets(nextExcluded);
    setFolderPlan([]);
    setPlan([]);
    setExtraMarks({});
    setAllowCreateFolders(false);
    setAllowApply(false);
  }

  // Blattnamen, die in Trimble auf denselben Ordnernamen fallen wuerden (z.B. "Planung" und "Plänung").
  function hasAmbiguousSheetNames(matrices: PermissionMatrix[]): boolean {
    const seen = new Map<string, string>();
    for (const matrix of matrices) {
      const key = normalizeLookup(matrix.sheetName);
      const other = seen.get(key);
      if (other) {
        addLog("error", `Blattnamen nicht eindeutig: "${other}" und "${matrix.sheetName}" waeren derselbe Ordner in Trimble.`, "Bitte eines der Blaetter abwaehlen oder umbenennen.");
        return true;
      }
      seen.set(key, matrix.sheetName);
    }
    return false;
  }

  function buildFolderPlans(matrices: PermissionMatrix[], tree: TCFolder[], parentId: string): FolderCreationStep[] {
    return matrices.flatMap((matrix) => buildFolderCreationPlan(matrix, tree, parentId));
  }

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

  // Teams, die in der Matrix fehlen, sind automatisch zum Loeschen vorgemerkt; per Klick uebersteuerbar.
  const markedIds = useMemo(() => {
    const unused = new Set(unusedGroups.map((group) => group.id));
    return new Set(
      groups.filter((group) => (deleteMarks[group.id] ?? unused.has(group.id))).map((group) => group.id)
    );
  }, [deleteMarks, groups, unusedGroups]);

  // Ordner unterhalb des Phasenordners, die nicht in der Matrix vorkommen (nur oberste Ebene je Zweig).
  // Standardmaessig zum Entfernen vorgemerkt; pro Zeile uebersteuerbar.
  const extraFolders = useMemo(() => {
    // Nur unterhalb der Phasenordner der geplanten Blaetter; alles andere wird nie angefasst.
    const rootPaths = folderPlan
      .filter((step) => step.depth === 0 && step.folderId)
      .map((step) => folders.find((folder) => folder.id === step.folderId)?.path)
      .filter((path): path is string => !!path);
    if (rootPaths.length === 0) return [];
    const matched = new Set(folderPlan.map((step) => step.folderId).filter(Boolean) as string[]);
    return folders.filter(
      (folder) =>
        !matched.has(folder.id) &&
        !!folder.parentId &&
        matched.has(folder.parentId) &&
        rootPaths.some((rootPath) => folder.path.startsWith(`${rootPath}/`))
    );
  }, [folderPlan, folders]);

  const markedExtraIds = useMemo(
    () => new Set(extraFolders.filter((folder) => extraMarks[folder.id] ?? true).map((folder) => folder.id)),
    [extraFolders, extraMarks]
  );

  async function removeExtraFolders() {
    const toDelete = extraFolders.filter((folder) => markedExtraIds.has(folder.id));
    if (toDelete.length === 0 || activeMatrices.length === 0 || !effectiveTargetParentId) return;
    const names = toDelete.map((folder) => folder.path).join("\n");
    if (!window.confirm(`${toDelete.length} Ordner inkl. Inhalt in Trimble Connect endgueltig loeschen?\n\n${names}`)) return;

    setBusy("delete-folders");
    cancelRef.current = false;
    setProgress({ current: 0, total: toDelete.length, label: "Ordner entfernen" });
    let deleted = 0;
    let attempted = 0;
    for (let index = 0; index < toDelete.length; index += 1) {
      if (cancelRef.current) {
        addLog("warning", `Abgebrochen: ${toDelete.length - index} Ordner wurden nicht entfernt.`);
        break;
      }
      const folder = toDelete[index];
      attempted += 1;
      try {
        await client.deleteFolder(folder.id);
        deleted += 1;
        addLog("success", `Ordner entfernt: ${folder.path}.`);
      } catch (error) {
        addLog("error", `Ordner konnte nicht entfernt werden: ${folder.path}.`, formatError(error));
      }
      setProgress({ current: index + 1, total: toDelete.length, label: "Ordner entfernen" });
    }
    setExtraMarks({});
    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog("info", `Ordner-Entfernung fertig. Entfernt: ${deleted}, Fehler: ${attempted - deleted}.`);

    const tree = await scanFolders();
    if (tree.length > 0) setFolderPlan(buildFolderPlans(activeMatrices, tree, effectiveTargetParentId));
  }

  const planStats = useMemo(() => {
    const done = plan.filter((item) => item.appliedAt).length;
    const pending = plan.filter((item) => !item.appliedAt);
    const ready = pending.filter((item) => item.status === "ready").length;
    const missingFolders = pending.filter((item) => item.status === "missing-folder").length;
    const missingGroups = pending.filter((item) => item.status === "missing-groups").length;
    return { ready, done, missingFolders, missingGroups };
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
    setFolderPlan([]);
    setExcludedSheets({});
    setDeleteMarks({});
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
      setDeleteMarks({});
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
    if (activeMatrices.length === 0) {
      addLog("warning", allSheets ? "Keine Tabellenblaetter ausgewaehlt." : "Bitte zuerst eine Phase aus der Matrix waehlen.");
      return;
    }
    if (!effectiveTargetParentId) {
      addLog("warning", "Kein Zielordner bekannt. Bitte zuerst Ordner scannen oder manuell waehlen.");
      return;
    }
    if (hasAmbiguousSheetNames(activeMatrices)) return;
    if (folders.length === 0) {
      addLog("warning", "Es wurden noch keine Ordner gescannt. Bitte zuerst Ordner scannen, sonst wuerden alle Ordner als fehlend gelten.");
      return;
    }

    const nextPlan = buildFolderPlans(activeMatrices, folders, effectiveTargetParentId);
    setFolderPlan(nextPlan);
    setExtraMarks({});

    const toCreate = nextPlan.filter((step) => step.status === "create").length;
    if (activeMatrices.length > 1) {
      activeMatrices.forEach((matrix) => {
        const steps = nextPlan.filter((step) => step.sheetName === matrix.sheetName);
        addLog("info", `${matrix.sheetName}: ${steps.filter((step) => step.status === "create").length} von ${steps.length} Ordnern fehlen.`);
      });
    }
    addLog(
      "info",
      `Ordner-Plan erstellt (${activeMatrices.length} Blatt/Blaetter): ${toCreate} von ${nextPlan.length} Ordnern fehlen unter "${targetParentLabel}".`
    );
  }

  async function createMissingFolders() {
    if (!allowCreateFolders) {
      addLog("warning", "Erstellen blockiert: Checkbox fuer Ordner-Erstellung ist nicht gesetzt.");
      return;
    }
    if (activeMatrices.length === 0 || folderPlan.length === 0) {
      addLog("warning", "Bitte zuerst einen Ordner-Plan erstellen.");
      return;
    }
    if (!effectiveTargetParentId) return;

    setBusy("create-folders");
    cancelRef.current = false;
    let cancelled = false;
    const toCreate = folderPlan.filter((step) => step.status === "create").length;
    setProgress({ current: 0, total: toCreate, label: "Ordner erstellen" });

    const resolvedIdByDepth = new Map<number, string>();
    let created = 0;
    let failed = 0;
    let progressDone = 0;
    const newFolders: TCFolder[] = [];

    for (const step of folderPlan) {
      if (cancelRef.current) {
        cancelled = true;
        break;
      }

      // Eintraege gleicher oder tieferer Ebene gehoeren zu frueheren Geschwistern/Blaettern und duerfen
      // nie als Eltern eines spaeteren Ordners dienen (sonst landet er im falschen Ordner).
      for (const depth of Array.from(resolvedIdByDepth.keys())) {
        if (depth >= step.depth) resolvedIdByDepth.delete(depth);
      }

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
        if (error instanceof TrimbleApiError && error.status === 409) {
          // Name belegt: den vorhandenen Ordner uebernehmen, statt abzubrechen.
          try {
            const match = await client.findChildByName(parentId, step.name);
            if (match?.isFolder) {
              const parentPath = newFolders.find((f) => f.id === parentId)?.path
                ?? folders.find((f) => f.id === parentId)?.path
                ?? "";
              newFolders.push({
                id: match.id,
                name: match.name,
                type: "FOLDER",
                parentId,
                path: parentPath ? `${parentPath}/${match.name}` : match.name,
                depth: step.depth,
              });
              resolvedIdByDepth.set(step.depth, match.id);
              addLog("info", `Ordner existierte bereits (beim Scan nicht erfasst), uebernommen: ${step.path}.`);
            } else {
              failed += 1;
              addLog(
                "error",
                `Name bereits belegt: ${step.path}.`,
                match
                  ? "Unter diesem Namen existiert eine Datei, kein Ordner. Bitte in Trimble Connect umbenennen oder loeschen."
                  : "Der Name ist belegt, aber der Eintrag ist fuer dieses Konto nicht sichtbar (z.B. fehlende Leserechte oder Papierkorb). In Trimble Connect pruefen."
              );
            }
          } catch (lookupError) {
            failed += 1;
            addLog("error", `Ordner konnte nicht erstellt werden: ${step.path}.`, formatError(error) + "\n" + formatError(lookupError));
          }
        } else {
          failed += 1;
          addLog("error", `Ordner konnte nicht erstellt werden: ${step.path}.`, formatError(error));
        }
      }

      progressDone += 1;
      setProgress({ current: progressDone, total: toCreate, label: "Ordner erstellen" });
    }

    if (newFolders.length > 0) {
      setFolders((current) => [...current, ...newFolders]);
      // Plan neu berechnen, damit erstellte Ordner als "Vorhanden" erscheinen und nicht erneut erstellt werden.
      setFolderPlan(buildFolderPlans(activeMatrices, [...folders, ...newFolders], effectiveTargetParentId));
    }

    setAllowCreateFolders(false);
    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    if (cancelled) addLog("warning", "Ordner-Erstellung abgebrochen. Bereits erstellte Ordner bleiben bestehen.");
    addLog("info", `Ordner-Erstellung fertig. Erstellt: ${created}, Fehler: ${failed}.`);

    // Ordnerbaum automatisch neu einlesen und Plan daraus aktualisieren; verbleibende Luecken melden.
    const tree = await scanFolders();
    if (tree.length > 0) {
      const verified = buildFolderPlans(activeMatrices, tree, effectiveTargetParentId);
      setFolderPlan(verified);
      const stillMissing = verified.filter((step) => step.status === "create");
      if (stillMissing.length > 0) {
        addLog(
          "warning",
          `Kontrolle: ${stillMissing.length} Ordner fehlen nach dem Neu-Scan weiterhin.`,
          stillMissing.slice(0, 30).map((step) => step.path).join("\n")
        );
      } else {
        addLog("success", `Kontrolle: Alle ${verified.length} Ordner sind in Trimble vorhanden.`);
      }
    }
    if (failed > 0) setActiveTab("log");
  }

  async function exportCurrentMatrix() {
    if (!workbook) {
      addLog("warning", "Bitte zuerst die Berechtigungsmatrix laden (dient als Struktur fuer den Export).");
      return;
    }

    setBusy("export");
    try {
      const activeGroups = groups.length > 0 ? groups : await loadGroups();
      const activeFolders = folders.length > 0 ? folders : await scanFolders();
      const stripPrefix = (id: string) => id.replace(/^[a-z-]+:/i, "").toLowerCase();
      const groupIdByTeam = new Map(activeGroups.map((group) => [normalizeLookup(group.name), stripPrefix(group.id)]));

      const plans = workbook.matrices.map((matrix) => ({
        matrix,
        items: buildPermissionPlan(matrix, activeGroups, activeFolders, { targetParentId: effectiveTargetParentId }),
      }));
      const total = plans.reduce((sum, entry) => sum + entry.items.filter((item) => item.folderId).length, 0);
      let done = 0;
      let failed = 0;
      let sampleLogged = false;
      setProgress({ current: 0, total, label: "Berechtigungen auslesen" });

      const sheets: ExportSheet[] = [];
      for (const { matrix, items } of plans) {
        const rows: ExportSheet["rows"] = [];

        for (let index = 0; index < matrix.rows.length; index += 1) {
          const row = matrix.rows[index];
          const item = items[index];
          let values = matrix.teams.map(() => "");

          if (item?.folderId) {
            try {
              const acl = await client.getFolderPermissions(item.folderId);
              const full = new Set(acl.FULL_ACCESS.map(stripPrefix));
              const read = new Set(acl.READ.map(stripPrefix));
              if (!sampleLogged) {
                sampleLogged = true;
                addLog(
                  "info",
                  `Export-Diagnose: ${row.folderPath}: ${full.size} Voll, ${read.size} Lesen, ${groupIdByTeam.size} Gruppen bekannt.`,
                  `Antwort: ${JSON.stringify(acl.raw).slice(0, 600)} | Beispiel-Gruppen-ID: ${Array.from(groupIdByTeam.values())[0] ?? "-"}`
                );
              }
              values = matrix.teams.map((team) => {
                const groupId = groupIdByTeam.get(normalizeLookup(team));
                if (!groupId) return "";
                return full.has(groupId) ? "V" : read.has(groupId) ? "L" : "";
              });
            } catch (error) {
              failed += 1;
              addLog("error", `Berechtigungen nicht lesbar: ${row.folderPath}.`, formatError(error));
            }
            done += 1;
            setProgress({ current: done, total, label: "Berechtigungen auslesen" });
          }

          rows.push({ depth: row.depth, folderName: row.folderName, values });
        }

        sheets.push({ name: matrix.sheetName, teams: matrix.teams, rows });
      }

      const { downloadCurrentMatrix } = await import("./excel/export");
      const stamp = new Date().toISOString().slice(0, 10);
      downloadCurrentMatrix(sheets, `Berechtigungsmatrix-Ist-${stamp}.xlsx`);
      addLog(
        failed > 0 ? "warning" : "success",
        `Aktuelle Matrix exportiert (${sheets.length} Phasen, ${total} Ordner ausgelesen, ${failed} Fehler).`,
        "Ordner ohne Treffer in Trimble und Teams ohne Gruppe sind leer."
      );
    } catch (error) {
      addLog("error", "Export fehlgeschlagen.", formatError(error));
    } finally {
      setBusy("");
      setProgress({ current: 0, total: 0, label: "" });
    }
  }

  const normalizeActorId = (id: string) => id.replace(/^[a-z-]+:/i, "").toLowerCase();

  // Stimmt der Ist-Stand eines Ordners mit den Matrix-Teams ueberein? Zusaetzliche Teams/Benutzer zaehlen nicht.
  function aclMatches(item: PermissionPlanItem, current: { READ: string[]; FULL_ACCESS: string[] }): boolean {
    const full = new Set(current.FULL_ACCESS.map(normalizeActorId));
    const read = new Set(current.READ.map(normalizeActorId));
    return (
      item.acl.FULL_ACCESS.every((id) => full.has(normalizeActorId(id))) &&
      item.acl.READ.every((id) => read.has(normalizeActorId(id)) && !full.has(normalizeActorId(id))) &&
      item.acl.NO_ACCESS.every((id) => !full.has(normalizeActorId(id)) && !read.has(normalizeActorId(id)))
    );
  }

  async function createDryRun() {
    if (activeMatrices.length === 0) {
      addLog("warning", allSheets ? "Keine Tabellenblaetter ausgewaehlt." : "Bitte zuerst eine Phase aus der Matrix waehlen.");
      return;
    }
    if (hasAmbiguousSheetNames(activeMatrices)) return;

    setBusy("dry-run");
    cancelRef.current = false;
    try {
      const activeGroups = groups.length > 0 ? groups : await loadGroups();
      const activeFolders = folders.length > 0 ? folders : await scanFolders();
      setBusy("dry-run");

      const options = { targetParentId: effectiveTargetParentId, strict: allSheets };
      const nextPlan = activeMatrices.flatMap((matrix) => buildPermissionPlan(matrix, activeGroups, activeFolders, options));

      // Zwei Zeilen mit unterschiedlichen Rechten auf demselben Ordner: nicht anwenden, sondern melden.
      const byFolder = new Map<string, PermissionPlanItem[]>();
      nextPlan.forEach((item) => {
        if (item.folderId) byFolder.set(item.folderId, [...(byFolder.get(item.folderId) ?? []), item]);
      });
      const signature = (item: PermissionPlanItem) =>
        JSON.stringify([[...item.acl.FULL_ACCESS].sort(), [...item.acl.READ].sort(), [...item.acl.NO_ACCESS].sort()]);
      const conflicts: string[] = [];
      byFolder.forEach((items) => {
        if (items.length > 1 && new Set(items.map(signature)).size > 1) {
          items.forEach((item) => {
            item.applyError = "Mehrere Matrix-Zeilen mit unterschiedlichen Rechten zeigen auf denselben Ordner.";
            item.folderId = undefined;
            item.status = "missing-folder";
          });
          conflicts.push(items.map((item) => item.folderPath).join("  <->  "));
        }
      });
      if (conflicts.length > 0) {
        addLog("error", `${conflicts.length} Konflikt(e): mehrere Matrix-Zeilen zeigen auf denselben Trimble-Ordner und werden ausgelassen.`, conflicts.slice(0, 30).join("\n"));
      }
      setPlan(nextPlan);

      // Abgleich mit dem Ist-Stand in Trimble: stimmen die Berechtigungen schon, gilt die Zeile als erledigt.
      const projectId = effectiveProject?.id;
      const candidates = nextPlan.filter((item) => item.folderId && item.status !== "missing-folder");
      if (projectId && candidates.length > 0) {
        const appliedMap = loadAppliedMap(projectId);
        let matched = 0;
        let checked = 0;
        setProgress({ current: 0, total: candidates.length, label: "Ist-Stand abgleichen" });

        for (const item of candidates) {
          if (cancelRef.current) {
            addLog("warning", `Abgleich abgebrochen: ${candidates.length - checked} Ordner wurden nicht geprueft (gelten als noch offen).`);
            break;
          }
          try {
            const current = await client.getFolderPermissions(item.folderId!);
            if (aclMatches(item, current)) {
              const appliedAt = appliedMap[item.folderId!] ?? new Date().toISOString();
              appliedMap[item.folderId!] = appliedAt;
              matched += 1;
              markPlanItem(item.rowNumber, item.folderPath, { appliedAt });
            } else if (appliedMap[item.folderId!]) {
              delete appliedMap[item.folderId!];
            }
          } catch (error) {
            addLog("warning", `Ist-Stand nicht lesbar: ${item.folderPath}.`, formatError(error));
          }
          checked += 1;
          setProgress({ current: checked, total: candidates.length, label: "Ist-Stand abgleichen" });
        }
        saveAppliedMap(projectId, appliedMap);
        addLog("info", `Abgleich: ${matched} von ${candidates.length} Ordnern stimmen bereits mit der Matrix ueberein (erledigt).`);
      }

      if (activeMatrices.length > 1) {
        activeMatrices.forEach((matrix) => {
          const items = nextPlan.filter((item) => item.sheetName === matrix.sheetName);
          addLog(
            "info",
            `${matrix.sheetName}: ${items.filter((item) => item.status === "ready").length} bereit, ${items.filter((item) => item.status === "missing-folder").length} Ordner fehlen, ${items.filter((item) => item.status === "missing-groups").length} mit fehlenden Gruppen.`
          );
        });
      }
      const ready = nextPlan.filter((item) => item.status === "ready").length;
      const missingFolders = nextPlan.filter((item) => item.status === "missing-folder").length;
      const missingGroups = nextPlan.filter((item) => item.status === "missing-groups").length;
      addLog("info", `Dry-Run erstellt (${activeMatrices.length} Blatt/Blaetter): ${ready} bereit, ${missingFolders} Ordner fehlen, ${missingGroups} mit fehlenden Gruppen.`);
    } finally {
      setBusy("");
      setProgress({ current: 0, total: 0, label: "" });
    }
  }

  function markPlanItem(rowNumber: number, folderPath: string, patch: Partial<PermissionPlanItem>) {
    setPlan((current) =>
      current.map((entry) =>
        entry.rowNumber === rowNumber && entry.folderPath === folderPath ? { ...entry, ...patch } : entry
      )
    );
  }

  async function applyPermissions() {
    if (!allowApply) {
      addLog("warning", "Aktivieren blockiert: Checkbox fuer Aenderungen ist nicht gesetzt.");
      return;
    }

    const readyItems = plan.filter(
      (item) => !item.appliedAt && item.folderId && (item.status === "ready" || (allowPartial && item.status === "missing-groups"))
    );
    if (readyItems.length === 0) {
      addLog("warning", "Keine bereiten Ordner im Dry-Run.");
      return;
    }

    const sheetCount = new Set(readyItems.map((item) => item.sheetName)).size;
    if (
      sheetCount > 1 &&
      !window.confirm(`Berechtigungen auf ${readyItems.length} Ordner in ${sheetCount} Tabellenblaettern anwenden?\n\nDer Vorgang kann lange dauern und laesst sich im Protokoll nachvollziehen. Mit "Abbrechen" im Fortschrittsbalken kann er jederzeit gestoppt werden.`)
    ) {
      return;
    }

    setBusy("apply");
    cancelRef.current = false;
    setProgress({ current: 0, total: readyItems.length, label: "Berechtigungen anwenden" });

    let applied = 0;
    let failed = 0;
    let processed = 0;
    const appliedItems: PermissionPlanItem[] = [];
    for (let index = 0; index < readyItems.length; index += 1) {
      if (cancelRef.current) {
        addLog("warning", `Abgebrochen: ${readyItems.length - index} Ordner wurden nicht bearbeitet. Bereits angewendete bleiben bestehen.`);
        break;
      }
      processed += 1;
      const item = readyItems[index];
      const oversized = (["READ", "FULL_ACCESS"] as const).filter((level) => item.acl[level].length > 100);

      if (oversized.length > 0) {
        const message = `Mehr als 100 Gruppen bei ${oversized.join(", ")} (API-Limit)`;
        addLog(
          "warning",
          `Uebersprungen: ${item.folderPath} hat mehr als 100 Gruppen bei ${oversized.join(", ")} (API-Limit).`,
          "Diese Zeile muss manuell in mehreren Schritten in Trimble Connect gepflegt werden."
        );
        failed += 1;
        markPlanItem(item.rowNumber, item.folderPath, { applyError: message });
        setProgress({ current: index + 1, total: readyItems.length, label: "Berechtigungen anwenden" });
        continue;
      }

      try {
        // Bestehende Eintraege, die nicht in der Matrix vorkommen (zusaetzliche Teams, Benutzer),
        // bleiben erhalten - nur Matrix-Teams werden ueberschrieben.
        const managed = new Set([...item.acl.READ, ...item.acl.FULL_ACCESS, ...item.acl.NO_ACCESS]);
        const current = await client.getFolderPermissions(item.folderId!);
        const extraFull = current.FULL_ACCESS.filter((id) => !managed.has(id));
        const extraRead = current.READ.filter((id) => !managed.has(id) && !extraFull.includes(id));
        const merged = {
          ...item.acl,
          FULL_ACCESS: [...item.acl.FULL_ACCESS, ...extraFull],
          READ: [...item.acl.READ, ...extraRead],
        };
        if (merged.READ.length > 100 || merged.FULL_ACCESS.length > 100) {
          throw new Error("Mehr als 100 Eintraege inkl. zusaetzlicher Teams (API-Limit)");
        }
        await client.updateFolderPermissions(item.folderId!, merged, false);
        applied += 1;
        appliedItems.push(item);
        const appliedAt = new Date().toISOString();
        if (effectiveProject?.id) {
          saveAppliedMap(effectiveProject.id, { ...loadAppliedMap(effectiveProject.id), [item.folderId!]: appliedAt });
        }
        markPlanItem(item.rowNumber, item.folderPath, { appliedAt, applyError: undefined });
      } catch (error) {
        failed += 1;
        addLog("error", `Berechtigungen fehlgeschlagen: ${item.folderPath}.`, formatError(error));
        markPlanItem(item.rowNumber, item.folderPath, { applyError: formatError(error) });
      }
      setProgress({ current: index + 1, total: readyItems.length, label: "Berechtigungen anwenden" });
    }

    // Kontrolle: angewendete Ordner zurueckliesen und mit der Matrix vergleichen.
    let verifyFailed = 0;
    if (appliedItems.length > 0 && !cancelRef.current) {
      setProgress({ current: 0, total: appliedItems.length, label: "Ergebnis kontrollieren" });
      const appliedMap = effectiveProject?.id ? loadAppliedMap(effectiveProject.id) : {};
      for (let index = 0; index < appliedItems.length; index += 1) {
        if (cancelRef.current) {
          addLog("warning", "Kontrolle abgebrochen.");
          break;
        }
        const item = appliedItems[index];
        try {
          const current = await client.getFolderPermissions(item.folderId!);
          if (!aclMatches(item, current)) {
            verifyFailed += 1;
            delete appliedMap[item.folderId!];
            markPlanItem(item.rowNumber, item.folderPath, { appliedAt: undefined, applyError: "Kontrolle: Berechtigungen in Trimble weichen von der Matrix ab." });
            addLog("error", `Kontrolle fehlgeschlagen: ${item.folderPath} weicht nach dem Anwenden von der Matrix ab.`);
          }
        } catch (error) {
          addLog("warning", `Kontrolle nicht moeglich: ${item.folderPath}.`, formatError(error));
        }
        setProgress({ current: index + 1, total: appliedItems.length, label: "Ergebnis kontrollieren" });
      }
      if (effectiveProject?.id) saveAppliedMap(effectiveProject.id, appliedMap);
    }

    setBusy("");
    setProgress({ current: 0, total: 0, label: "" });
    addLog(
      failed + verifyFailed > 0 ? "warning" : "success",
      `${applied} Ordner aktualisiert, ${failed} Fehler, ${verifyFailed} bei der Kontrolle abweichend, ${readyItems.length - processed} nicht bearbeitet.`
    );
    if (failed + verifyFailed > 0) setActiveTab("log");
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
              <h2>Berechtigung fehlt für diese Applikation</h2>
              {accessError &&<p className="access-detail">Pruefung fehlgeschlagen: {accessError}</p>}
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
        <button className={activeTab === "access" ? "active" : ""} onClick={() => setActiveTab("access")}>App-Zugriff</button>
        <span className="history-actions">
          <button
            onClick={undo}
            disabled={isBusy || historySizes.past === 0}
            title={`Letzte Aenderung zurueck (${historySizes.past} moeglich, max. ${HISTORY_LIMIT}). Betrifft nur die Ansicht, nicht bereits in Trimble ausgefuehrte Aenderungen.`}
          >
            ↶ Zurück
          </button>
          <button
            onClick={redo}
            disabled={isBusy || historySizes.future === 0}
            title={`Aenderung wiederherstellen (${historySizes.future} moeglich)`}
          >
            Vorwärts ↷
          </button>
        </span>
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
            newGroupName={newGroupName}
            setNewGroupName={setNewGroupName}
            onWorkbook={handleWorkbook}
            onLoadGroups={loadGroups}
            onCreateGroup={createSingleGroup}
            onImportGroups={importMissingGroups}
            markedIds={markedIds}
            onToggleMark={toggleMark}
            onCommitMarked={commitMarkedGroups}
            onKeepAll={keepAllGroups}
          />
        )}

        {activeTab === "folders" && (
          <FoldersTab
            busy={isBusy}
            workbook={workbook}
            selectedSheet={selectedSheet}
            setSelectedSheet={setSelectedSheet}
            currentMatrix={currentMatrix}
            allSheets={allSheets}
            excludedSheets={excludedSheets}
            onScopeChange={changeScope}
            scopeStats={scopeStats}
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
            extraFolders={extraFolders}
            markedExtraIds={markedExtraIds}
            onToggleExtra={(id) => setExtraMarks((current) => ({ ...current, [id]: !markedExtraIds.has(id) }))}
            onRemoveExtra={removeExtraFolders}
          />
        )}

        {activeTab === "permissions" && (
          <PermissionsTab
            busy={isBusy}
            workbook={workbook}
            selectedSheet={selectedSheet}
            setSelectedSheet={setSelectedSheet}
            currentMatrix={currentMatrix}
            allSheets={allSheets}
            excludedSheets={excludedSheets}
            onScopeChange={changeScope}
            scopeStats={scopeStats}
            folders={folders}
            groups={groups}
            projectDetails={projectDetails}
            plan={plan}
            planStats={planStats}
            progress={progress}
            allowApply={allowApply}
            allowPartial={allowPartial}
            setAllowPartial={setAllowPartial}
            setAllowApply={setAllowApply}
            onWorkbook={handleWorkbook}
            onLoadGroups={loadGroups}
            onScanFolders={scanFolders}
            onDryRun={createDryRun}
            onApply={applyPermissions}
            onExport={exportCurrentMatrix}
          />
        )}

        {activeTab === "log" && <LogPanel logs={logs} />}

        {activeTab === "access" && (
          <AccessTab
            busy={isBusy}
            admins={projectAdmins}
            allowed={allowedUsers}
            currentUserId={currentUser?.id}
            removedNotes={removedNotes}
            warning={accessListWarning}
            onReload={reloadAccess}
            onAdd={addAccessUser}
            onRemove={removeAccessUser}
          />
        )}
      </main>

      {progress.total > 0 && (
        <footer className="progress-footer">
          <span>{progress.label}</span>
          <progress value={progress.current} max={progress.total} />
          <strong>{progress.current}/{progress.total}</strong>
          {isBusy && busy !== "folders" && (
            <button onClick={() => { cancelRef.current = true; }} title="Laufenden Vorgang nach dem aktuellen Schritt stoppen. Bereits Erledigtes bleibt bestehen.">
              Abbrechen
            </button>
          )}
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
  newGroupName: string;
  setNewGroupName: (value: string) => void;
  onWorkbook: (file: File) => void;
  onLoadGroups: () => void;
  onCreateGroup: () => void;
  onImportGroups: () => void;
  markedIds: Set<string>;
  onToggleMark: (groupId: string) => void;
  onCommitMarked: () => void;
  onKeepAll: () => void;
}

interface AccessTabProps {
  busy: boolean;
  admins: TCUser[];
  allowed: TCUser[];
  currentUserId?: string;
  removedNotes: { name: string; email: string; at: string }[];
  warning: string;
  onReload: () => void;
  onAdd: (userId: string) => void;
  onRemove: (userId: string) => void;
}

function AccessTab(props: AccessTabProps) {
  const [selected, setSelected] = useState("");
  const allowedIds = new Set(props.allowed.map((user) => user.id));
  const candidates = props.admins.filter((user) => !allowedIds.has(user.id));

  return (
    <div className="panel-grid">
      <section className="panel">
        <h2>App-Zugriff</h2>
        <p className="hint">
          Nur hier aufgeführte Projekt-Administratoren dürfen diese Applikation verwenden. Alle anderen sehen
          „Berechtigung fehlt für diese Applikation“. Solange die Liste leer ist, haben alle Administratoren Zugriff.
          Administratoren, die zurückgestuft wurden, werden beim Öffnen automatisch entfernt.
        </p>
        {props.warning && <div className="banner warning">{props.warning}</div>}
        <div className="input-row">
          <select value={selected} onChange={(event) => setSelected(event.target.value)} disabled={props.busy}>
            <option value="">Administrator wählen ...</option>
            {candidates.map((user) => (
              <option key={user.id} value={user.id}>{user.name}{user.email && user.email !== user.name ? ` (${user.email})` : ""}</option>
            ))}
          </select>
          <button
            className="primary"
            disabled={props.busy || !selected}
            onClick={() => {
              props.onAdd(selected);
              setSelected("");
            }}
          >
            Hinzufügen
          </button>
          <button onClick={props.onReload} disabled={props.busy} title="Aktuelle Administratoren neu laden">Aktualisieren</button>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Name</th>
                <th>E-Mail</th>
                <th aria-label="Entfernen"></th>
              </tr>
            </thead>
            <tbody>
              {props.allowed.map((user) => (
                <tr key={user.id}>
                  <td>{user.name}{user.id === props.currentUserId && <span className="delete-note"> (du)</span>}</td>
                  <td>{user.email}</td>
                  <td className="row-remove">
                    <button
                      className="icon-button"
                      onClick={() => props.onRemove(user.id)}
                      disabled={props.busy || user.id === props.currentUserId}
                      title={user.id === props.currentUserId ? "Eigener Zugriff kann nicht entfernt werden" : "Freigabe entfernen"}
                      aria-label={`${user.name} entfernen`}
                    >
                      ×
                    </button>
                  </td>
                </tr>
              ))}
              {props.allowed.length === 0 && (
                <tr>
                  <td colSpan={3} className="empty-cell">Keine Auswahl – alle Projekt-Administratoren haben Zugriff</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel">
        <h2>Automatisch entfernt</h2>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Benutzer</th>
                <th>Entfernt am</th>
              </tr>
            </thead>
            <tbody>
              {props.removedNotes.map((note, index) => (
                <tr key={`${note.email}-${note.at}-${index}`}>
                  <td>{note.name}{note.email && note.email !== note.name ? ` (${note.email})` : ""}</td>
                  <td>{formatDateTime(note.at)} – nicht mehr Administrator</td>
                </tr>
              ))}
              {props.removedNotes.length === 0 && (
                <tr>
                  <td colSpan={2} className="empty-cell">Keine automatischen Entfernungen</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
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
          <button
            className="danger"
            onClick={props.onCommitMarked}
            disabled={props.busy || props.markedIds.size === 0}
            title="Rot markierte Teams in Trimble Connect loeschen"
          >
            Übernehmen ({props.markedIds.size})
          </button>
          {props.markedIds.size > 0 && (
            <button onClick={props.onKeepAll} disabled={props.busy}>Alle behalten</button>
          )}
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
                <th aria-label="Entfernen"></th>
              </tr>
            </thead>
            <tbody>
              {[...props.groups]
                .sort((a, b) => Number(props.markedIds.has(b.id)) - Number(props.markedIds.has(a.id)))
                .map((group) => {
                const marked = props.markedIds.has(group.id);
                return (
                  <tr key={group.id} className={marked ? "marked-delete" : undefined}>
                    <td>
                      {group.name}
                      {marked && <span className="delete-note"> (wird gelöscht)</span>}
                    </td>
                    <td>{group.usersCount ?? "-"}</td>
                    <td className="row-remove">
                      <button
                        className="icon-button"
                        onClick={() => props.onToggleMark(group.id)}
                        disabled={props.busy}
                        title={marked ? "Loeschung abbrechen" : "Team zum Loeschen vormerken (erst mit Übernehmen wird es in Trimble geloescht)"}
                        aria-label={marked ? `${group.name} behalten` : `${group.name} entfernen`}
                      >
                        {marked ? "↶" : "×"}
                      </button>
                    </td>
                  </tr>
                );
              })}
              {props.groups.length === 0 && (
                <tr>
                  <td colSpan={3} className="empty-cell">Keine Gruppen geladen</td>
                </tr>
              )}
            </tbody>
          </table>
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
  allSheets: boolean;
  excludedSheets: Record<string, boolean>;
  onScopeChange: (allSheets: boolean, excludedSheets: Record<string, boolean>) => void;
  scopeStats: ScopeStats;
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
  extraFolders: TCFolder[];
  markedExtraIds: Set<string>;
  onToggleExtra: (id: string) => void;
  onRemoveExtra: () => void;
}

function FoldersTab(props: FoldersTabProps) {
  const toCreate = props.folderPlan.filter((step) => step.status === "create").length;
  const [folderSort, setFolderSort] = useState<{ key: "status" | "path"; dir: 1 | -1 } | null>(null);
  const sortedPlan = folderSort
    ? [...props.folderPlan].sort((a, b) =>
        folderSort.dir *
        (folderSort.key === "status"
          ? a.status.localeCompare(b.status)
          : a.path.localeCompare(b.path, "de"))
      )
    : props.folderPlan;
  const toggleSort = (key: "status" | "path") =>
    setFolderSort((prev) => (prev?.key === key ? { key, dir: prev.dir === 1 ? -1 : 1 } : { key, dir: 1 }));
  const sortMark = (key: "status" | "path") => (folderSort?.key === key ? (folderSort.dir === 1 ? " ▲" : " ▼") : "");

  return (
    <div className="panel-grid wide">
      <section className="panel">
        <h2>Phase</h2>
        <FileInput disabled={props.busy} onFile={props.onWorkbook} />
        <SheetScope
          workbook={props.workbook}
          selectedSheet={props.selectedSheet}
          setSelectedSheet={props.setSelectedSheet}
          allSheets={props.allSheets}
          excludedSheets={props.excludedSheets}
          onScopeChange={props.onScopeChange}
          disabled={props.busy}
        />
        <MetricGrid
          metrics={[
            ...(props.allSheets ? ([["Blätter", props.scopeStats.sheets]] as Array<[string, number]>) : []),
            ["Ordner in Matrix", props.scopeStats.folders],
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
          Wohin {props.allSheets ? "die Phasenordner (je Blattname)" : <>der Phasenordner (Blattname <strong>{props.currentMatrix?.sheetName ?? "–"}</strong>)</>} kommt,
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

      {props.extraFolders.length > 0 && (
        <section className="panel span-2">
          <div className="panel-title-row">
            <h2>Nicht in der Matrix</h2>
            <div className="summary-pills">
              <span className="pill warn">{props.markedExtraIds.size} werden entfernt</span>
              <span className="pill ok">{props.extraFolders.length - props.markedExtraIds.size} bleiben</span>
            </div>
          </div>
          <p className="hint">
            Diese Ordner existieren in Trimble, kommen aber nicht in der Matrix vor. Haken pro Zeile entfernen = Ordner bleibt.
            Beim Entfernen wird auch der Inhalt geloescht.
          </p>
          <div className="apply-row">
            <button
              className="danger"
              onClick={props.onRemoveExtra}
              disabled={props.busy || props.markedExtraIds.size === 0}
            >
              Markierte Ordner entfernen ({props.markedExtraIds.size})
            </button>
          </div>
          <div className="table-wrap plan-table">
            <table>
              <thead>
                <tr>
                  <th>Status</th>
                  <th>Pfad</th>
                  <th>Entfernen</th>
                </tr>
              </thead>
              <tbody>
                {[...props.extraFolders]
                  .sort((a, b) => a.path.localeCompare(b.path, "de"))
                  .map((folder) => {
                    const marked = props.markedExtraIds.has(folder.id);
                    return (
                      <tr key={folder.id} className={marked ? "marked-delete" : undefined}>
                        <td>
                          <span className={`state ${marked ? "missing-folder" : "ready"}`}>{marked ? "Wird entfernt" : "Bleibt"}</span>
                        </td>
                        <td>{folder.path}</td>
                        <td>
                          <input
                            type="checkbox"
                            checked={marked}
                            disabled={props.busy}
                            onChange={() => props.onToggleExtra(folder.id)}
                            aria-label={`${folder.path} entfernen`}
                          />
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <section className="panel span-2">
        <div className="panel-title-row">
          <h2>Ordner-Plan</h2>
          <div className="summary-pills">
            <span className="pill warn">{toCreate} werden erstellt</span>
            <span className="pill ok">{props.folderPlan.length - toCreate} vorhanden</span>
          </div>
        </div>
        <div className="inline-actions">
          <button className="primary" onClick={props.onPlanFolders} disabled={props.busy || props.scopeStats.sheets === 0}>
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
              <tr>
                <th style={{ cursor: "pointer" }} onClick={() => toggleSort("status")}>Status{sortMark("status")}</th>
                <th style={{ cursor: "pointer" }} onClick={() => toggleSort("path")}>Pfad{sortMark("path")}</th>
              </tr>
            </thead>
            <tbody>
              {sortedPlan.map((step) => (
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
  allSheets: boolean;
  excludedSheets: Record<string, boolean>;
  onScopeChange: (allSheets: boolean, excludedSheets: Record<string, boolean>) => void;
  scopeStats: ScopeStats;
  folders: TCFolder[];
  groups: TCGroup[];
  projectDetails: TCProject | null;
  plan: PermissionPlanItem[];
  planStats: { ready: number; done: number; missingFolders: number; missingGroups: number };
  progress: ProgressState;
  allowApply: boolean;
  allowPartial: boolean;
  setAllowPartial: (value: boolean) => void;
  setAllowApply: (value: boolean) => void;
  onWorkbook: (file: File) => void;
  onLoadGroups: () => void;
  onScanFolders: () => void;
  onDryRun: () => void;
  onApply: () => void;
  onExport: () => void;
}

function PermissionsTab(props: PermissionsTabProps) {
  const missingGroupNames = Array.from(new Set(props.plan.flatMap((item) => item.missingGroups))).sort((a, b) =>
    a.localeCompare(b, "de")
  );
  const applicableCount = props.plan.filter(
    (item) => !item.appliedAt && item.folderId && (item.status === "ready" || (props.allowPartial && item.status === "missing-groups"))
  ).length;

  return (
    <div className="panel-grid wide">
      <section className="panel">
        <h2>Phase</h2>
        <FileInput disabled={props.busy} onFile={props.onWorkbook} />
        <SheetScope
          workbook={props.workbook}
          selectedSheet={props.selectedSheet}
          setSelectedSheet={props.setSelectedSheet}
          allSheets={props.allSheets}
          excludedSheets={props.excludedSheets}
          onScopeChange={props.onScopeChange}
          disabled={props.busy}
        />
        <MetricGrid
          metrics={[
            ...(props.allSheets ? ([["Blätter", props.scopeStats.sheets]] as Array<[string, number]>) : []),
            ["Ordner", props.scopeStats.folders],
            ["Teams", props.scopeStats.teams],
            ["Voll", props.scopeStats.fullAccess],
            ["Lesen", props.scopeStats.read],
            ["Kein", props.scopeStats.noAccess],
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
          <button className="primary" onClick={props.onDryRun} disabled={props.busy || props.scopeStats.sheets === 0}>Dry-Run</button>
          <button
            onClick={props.onExport}
            disabled={props.busy || !props.workbook}
            title="Aktuelle Berechtigungen aus Trimble als Matrix (Excel) exportieren"
          >
            Aktuelle Matrix exportieren
          </button>
        </div>
      </section>

      <section className="panel span-2">
        <div className="panel-title-row">
          <h2>Dry-Run</h2>
          <div className="summary-pills">
            <span className="pill ok">{props.planStats.done} erledigt</span>
            <span className="pill ok">{props.planStats.ready} bereit</span>
            <span className="pill warn">{props.planStats.missingFolders} Ordner fehlen</span>
            <span className="pill warn">{props.planStats.missingGroups} Gruppen fehlen</span>
          </div>
        </div>
        {props.plan.length > 0 && (
          missingGroupNames.length === 0 && props.planStats.missingFolders === 0 ? (
            <div className="banner info">Alle Ordner und Gruppen sind vorhanden. Die Berechtigungen koennen angewendet werden.</div>
          ) : (
            <div className="banner warning">
              {props.planStats.missingFolders > 0 && (
                <div>{props.planStats.missingFolders} Ordner fehlen in Trimble und werden beim Anwenden ausgelassen (optional im Register "Ordner erstellen" anlegen).</div>
              )}
              {missingGroupNames.length > 0 && (
                <div>
                  {missingGroupNames.length} Teams fehlen in Trimble (im Register "Teams" erstellen):{" "}
                  {missingGroupNames.slice(0, 8).join(", ")}{missingGroupNames.length > 8 ? " ..." : ""}
                </div>
              )}
            </div>
          )
        )}
        <div className="apply-row">
          <label className="check">
            <input
              type="checkbox"
              checked={props.allowApply}
              onChange={(event) => props.setAllowApply(event.target.checked)}
            />
            <span>Aenderungen wirklich anwenden</span>
          </label>
          <label className="check">
            <input
              type="checkbox"
              checked={props.allowPartial}
              onChange={(event) => props.setAllowPartial(event.target.checked)}
            />
            <span>Auch Zeilen mit fehlenden Gruppen anwenden (fehlende Teams werden ausgelassen)</span>
          </label>
          <button className="danger" onClick={props.onApply} disabled={props.busy || !props.allowApply || applicableCount === 0}>
            Berechtigungen anwenden
          </button>
        </div>
        {props.plan.length > PLAN_DISPLAY_LIMIT && (
          <div className="banner info">Angezeigt werden die ersten {PLAN_DISPLAY_LIMIT} von {props.plan.length} Zeilen; verarbeitet werden alle.</div>
        )}
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
              {props.plan.slice(0, PLAN_DISPLAY_LIMIT).map((item) => {
                const display = planRowDisplay(item);
                return (
                  <tr key={`${item.rowNumber}-${item.folderPath}`}>
                    <td>
                      <span className={`state ${display.className}`} title={item.applyError}>
                        {display.label}
                      </span>
                    </td>
                    <td>{item.folderPath}</td>
                    <td>{item.counts.fullAccess}</td>
                    <td>{item.counts.read}</td>
                    <td>{item.counts.noAccess}</td>
                    <td title={item.missingGroups.join(", ")}>{item.missingGroups.slice(0, 3).join(", ")}{item.missingGroups.length > 3 ? ` ... (+${item.missingGroups.length - 3})` : ""}</td>
                  </tr>
                );
              })}
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

const PLAN_DISPLAY_LIMIT = 3000;

interface ScopeStats {
  sheets: number;
  folders: number;
  teams: number;
  fullAccess: number;
  read: number;
  noAccess: number;
}

interface SheetScopeProps {
  workbook: WorkbookModel | null;
  selectedSheet: string;
  setSelectedSheet: (sheet: string) => void;
  allSheets: boolean;
  excludedSheets: Record<string, boolean>;
  onScopeChange: (allSheets: boolean, excludedSheets: Record<string, boolean>) => void;
  disabled: boolean;
}

function SheetScope(props: SheetScopeProps) {
  const sheets = props.workbook?.sheetNames ?? [];
  const includedCount = sheets.filter((sheet) => !props.excludedSheets[sheet]).length;

  return (
    <>
      <label className="field">
        <span>Tabellenblatt</span>
        <select
          value={props.selectedSheet}
          onChange={(event) => props.setSelectedSheet(event.target.value)}
          disabled={props.allSheets}
        >
          {sheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
        </select>
      </label>
      <label className="check">
        <input
          type="checkbox"
          checked={props.allSheets}
          disabled={props.disabled || sheets.length === 0}
          onChange={(event) => props.onScopeChange(event.target.checked, props.excludedSheets)}
        />
        <span>Alle Tabellenblätter ({includedCount} von {sheets.length} gewählt)</span>
      </label>
      {props.allSheets && (
        <div className="sheet-list">
          {sheets.map((sheet) => (
            <label key={sheet} className="check">
              <input
                type="checkbox"
                checked={!props.excludedSheets[sheet]}
                disabled={props.disabled}
                onChange={() => props.onScopeChange(true, { ...props.excludedSheets, [sheet]: !props.excludedSheets[sheet] })}
              />
              <span>{sheet}</span>
            </label>
          ))}
        </div>
      )}
    </>
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

function formatDateTime(iso: string): string {
  const date = new Date(iso);
  const datePart = date.toLocaleDateString("de-CH");
  const timePart = date.toLocaleTimeString("de-CH", { hour: "2-digit", minute: "2-digit" });
  return `${datePart} ${timePart}`;
}

function planRowDisplay(item: PermissionPlanItem): { className: string; label: string } {
  if (item.appliedAt) return { className: "applied", label: `Erledigt (${formatDateTime(item.appliedAt)})` };
  if (item.applyError) return { className: "apply-failed", label: "Fehlgeschlagen" };
  return { className: item.status, label: statusLabel(item.status) };
}

function guessRegion(location: string): ApiRegion | null {
  const value = normalizeLookup(location);
  if (value.includes("europe") || value.includes("eu")) return "eu";
  if (value.includes("uk") || value.includes("united kingdom")) return "uk";
  if (value.includes("asia") || value.includes("apac")) return "ap";
  if (value.includes("northamerica") || value.includes("us")) return "us";
  return null;
}
