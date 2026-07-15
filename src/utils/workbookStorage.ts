import type { WorkbookModel } from "../types";

const STORAGE_PREFIX = "tc-permission-matrix:v1:";

function storageKey(projectId: string): string {
  return `${STORAGE_PREFIX}${projectId}`;
}

export function loadStoredWorkbook(projectId: string): WorkbookModel | null {
  try {
    const raw = window.localStorage.getItem(storageKey(projectId));
    if (!raw) return null;
    return JSON.parse(raw) as WorkbookModel;
  } catch {
    return null;
  }
}

/**
 * rawValues duplicates the (team name -> original cell text) map on every single matrix row and
 * is never read back anywhere in the app - only `permissions` (the parsed level) is used. For a
 * workbook with many phases/teams that duplication alone can push the serialized size past the
 * browser's localStorage quota (usually 5-10 MB/origin), so it's stripped before persisting.
 */
function stripForStorage(workbook: WorkbookModel): WorkbookModel {
  return {
    ...workbook,
    matrices: workbook.matrices.map((matrix) => ({
      ...matrix,
      rows: matrix.rows.map((row) => ({ ...row, rawValues: {} })),
    })),
  };
}

export function saveStoredWorkbook(projectId: string, workbook: WorkbookModel): { ok: true } | { ok: false; error: string } {
  try {
    window.localStorage.setItem(storageKey(projectId), JSON.stringify(stripForStorage(workbook)));
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

export function clearStoredWorkbook(projectId: string): void {
  try {
    window.localStorage.removeItem(storageKey(projectId));
  } catch {
    // Ignore.
  }
}
