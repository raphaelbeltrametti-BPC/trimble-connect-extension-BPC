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

export function saveStoredWorkbook(projectId: string, workbook: WorkbookModel): void {
  try {
    window.localStorage.setItem(storageKey(projectId), JSON.stringify(workbook));
  } catch {
    // Storage full or unavailable (e.g. private browsing) - the matrix simply won't survive a reload this time.
  }
}

export function clearStoredWorkbook(projectId: string): void {
  try {
    window.localStorage.removeItem(storageKey(projectId));
  } catch {
    // Ignore.
  }
}
