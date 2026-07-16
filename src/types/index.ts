export type ApiRegion = "us" | "eu" | "uk" | "ap" | "ap2" | "custom";

export type PermissionLevel = "FULL_ACCESS" | "READ" | "NO_ACCESS";

export interface TCProject {
  id: string;
  name: string;
  location?: string;
  rootId?: string;
  /** The caller's project membership role, e.g. "ADMIN", "USER" (uppercase, per GET /projects/{id}). */
  role?: string;
}

export interface TCGroup {
  id: string;
  name: string;
  projectId?: string;
  usersCount?: number;
}

export interface TCFolder {
  id: string;
  name: string;
  type: string;
  parentId?: string;
  path: string;
  depth: number;
}

export interface MatrixRow {
  rowNumber: number;
  folderName: string;
  relativePath: string;
  folderPath: string;
  depth: number;
  permissions: Record<string, PermissionLevel>;
  rawValues: Record<string, string>;
}

export interface PermissionMatrix {
  sheetName: string;
  teams: string[];
  rows: MatrixRow[];
  stats: {
    folders: number;
    teams: number;
    fullAccess: number;
    read: number;
    noAccess: number;
    unknownValues: number;
  };
}

export interface WorkbookModel {
  fileName: string;
  sheetNames: string[];
  matrices: PermissionMatrix[];
  teamNames: string[];
}

export interface FolderAcl {
  READ: string[];
  FULL_ACCESS: string[];
  NO_ACCESS: string[];
}

export interface FolderCreationStep {
  rowNumber: number;
  depth: number;
  name: string;
  path: string;
  status: "existing" | "create";
  folderId?: string;
}

export interface PermissionPlanItem {
  rowNumber: number;
  folderPath: string;
  folderName: string;
  folderId?: string;
  status: "ready" | "missing-folder" | "missing-groups";
  acl: FolderAcl;
  missingGroups: string[];
  counts: {
    fullAccess: number;
    read: number;
    noAccess: number;
  };
}

export interface LogEntry {
  id: string;
  level: "info" | "success" | "warning" | "error";
  message: string;
  detail?: string;
  createdAt: string;
}

export interface ProgressState {
  current: number;
  total: number;
  label: string;
}
