import * as XLSX from "xlsx";
import type { MatrixRow, PermissionLevel, PermissionMatrix, WorkbookModel } from "../types";
import { leadingSpaceCount, normalizeLookup, normalizeWhitespace, uniqueSorted } from "../utils/text";

const NO_ACCESS_VALUES = new Set(["", "k", "kein zugriff", "keine", "none", "no access", "no_access"]);
const READ_VALUES = new Set(["l", "lesezugriff", "read", "read only", "read_only"]);
const FULL_VALUES = new Set([
  "v",
  "vollzugriff",
  "full",
  "full access",
  "full_access",
  "read_write",
  "readwrite",
  "rw",
]);

export function permissionFromCell(value: unknown): { permission: PermissionLevel; known: boolean } {
  const normalized = normalizeLookup(value).replace(/-/g, " ");

  if (FULL_VALUES.has(normalized)) {
    return { permission: "FULL_ACCESS", known: true };
  }

  if (READ_VALUES.has(normalized)) {
    return { permission: "READ", known: true };
  }

  if (NO_ACCESS_VALUES.has(normalized)) {
    return { permission: "NO_ACCESS", known: true };
  }

  return { permission: "NO_ACCESS", known: false };
}

export async function parseWorkbookFile(file: File): Promise<WorkbookModel> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const matrices = workbook.SheetNames.map((sheetName) => parseSheet(sheetName, workbook.Sheets[sheetName]))
    .filter((matrix): matrix is PermissionMatrix => matrix !== null);

  return {
    fileName: file.name,
    sheetNames: matrices.map((matrix) => matrix.sheetName),
    matrices,
    teamNames: uniqueSorted(matrices.flatMap((matrix) => matrix.teams)),
  };
}

function parseSheet(sheetName: string, sheet: XLSX.WorkSheet | undefined): PermissionMatrix | null {
  if (!sheet) return null;

  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
    raw: false,
  });

  if (rows.length < 2) return null;

  const header = rows[0].map(normalizeWhitespace);
  const teams = header.slice(1).filter(Boolean);
  if (teams.length === 0) return null;

  const pathStack: string[] = [];
  const matrixRows: MatrixRow[] = [];
  let fullAccess = 0;
  let read = 0;
  let noAccess = 0;
  let unknownValues = 0;

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const rawFolder = String(rows[rowIndex][0] ?? "");
    const folderName = normalizeWhitespace(rawFolder);
    if (!folderName) continue;

    const depth = Math.max(0, Math.floor(leadingSpaceCount(rawFolder) / 2));
    pathStack[depth] = folderName;
    pathStack.length = depth + 1;

    const relativePath = pathStack.join("/");
    const folderPath = `${sheetName}/${relativePath}`;
    const permissions: Record<string, PermissionLevel> = {};
    const rawValues: Record<string, string> = {};

    teams.forEach((team, offset) => {
      const raw = normalizeWhitespace(rows[rowIndex][offset + 1]);
      const parsed = permissionFromCell(raw);
      permissions[team] = parsed.permission;
      rawValues[team] = raw;

      if (!parsed.known) unknownValues += 1;
      if (parsed.permission === "FULL_ACCESS") fullAccess += 1;
      if (parsed.permission === "READ") read += 1;
      if (parsed.permission === "NO_ACCESS") noAccess += 1;
    });

    matrixRows.push({
      rowNumber: rowIndex + 1,
      folderName,
      relativePath,
      folderPath,
      depth,
      permissions,
      rawValues,
    });
  }

  return {
    sheetName,
    teams,
    rows: matrixRows,
    stats: {
      folders: matrixRows.length,
      teams: teams.length,
      fullAccess,
      read,
      noAccess,
      unknownValues,
    },
  };
}
