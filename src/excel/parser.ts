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

function gcd(a: number, b: number): number {
  return b === 0 ? a : gcd(b, a % b);
}

/**
 * The indentation width isn't fixed at 2 spaces per level in practice - files produced by pasting
 * from Word/other tools, or by repeated tab expansion, commonly end up with 4, 6 or other widths
 * per level. Hardcoding "2" then overcounts depth (e.g. 6 spaces / 2 = depth 3 instead of 1),
 * producing garbled paths with empty segments that never match any real folder. Instead, derive
 * the sheet's actual per-level width from the greatest common divisor of all observed non-zero
 * leading-space counts, so a single first-level row (any consistent width) resolves to depth 1.
 */
function detectIndentUnit(rows: string[][]): number {
  let unit = 0;
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const rawFolder = String(rows[rowIndex][0] ?? "");
    if (!normalizeWhitespace(rawFolder)) continue;

    const count = leadingSpaceCount(rawFolder);
    if (count > 0) unit = gcd(unit, count);
  }

  return unit || 2;
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

  const indentUnit = detectIndentUnit(rows);
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

    const rawDepth = Math.max(0, Math.round(leadingSpaceCount(rawFolder) / indentUnit));

    // The first data row conventionally restates the phase/sheet name itself (its own permissions
    // row for the top folder). Treating it like a normal child would prefix the sheet name onto
    // itself (e.g. "31-VORPROJEKT/31-VORPROJEKT"), a path that can never exist. Its depth is set to
    // -1 (distinct from any real row) so consumers can recognize it refers to the phase root itself
    // rather than a nested folder.
    const isPhaseRootRow = matrixRows.length === 0 && rawDepth === 0 && normalizeLookup(folderName) === normalizeLookup(sheetName);

    pathStack[rawDepth] = folderName;
    pathStack.length = rawDepth + 1;

    const relativePath = isPhaseRootRow ? "" : pathStack.join("/");
    const folderPath = isPhaseRootRow ? sheetName : `${sheetName}/${relativePath}`;
    const depth = isPhaseRootRow ? -1 : rawDepth;
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
