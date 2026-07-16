import type { FolderCreationStep, PermissionMatrix, TCFolder } from "../types";
import { normalizeLookup } from "../utils/text";

/**
 * Plans which folders a phase (Excel sheet) needs, in creation order: the sheet name becomes the
 * phase's top folder under `targetParentId`, then every matrix row nests under its depth-1
 * predecessor - mirroring exactly how the Excel parser builds `relativePath` from indentation, so
 * a row's parent is always the immediately preceding row at depth-1 (already resolved by the time
 * we reach it, since rows are processed top to bottom in file order).
 *
 * This only reads already-scanned folders; it never calls the API. Existing folders are matched
 * by (parentId, normalized name) so the plan reflects exactly what applyFolderCreationPlan()
 * would need to create.
 */
export function buildFolderCreationPlan(
  matrix: PermissionMatrix,
  folders: TCFolder[],
  targetParentId: string
): FolderCreationStep[] {
  const childrenByParent = new Map<string, TCFolder[]>();
  folders.forEach((folder) => {
    if (!folder.parentId) return;
    childrenByParent.set(folder.parentId, [...(childrenByParent.get(folder.parentId) ?? []), folder]);
  });

  const findChild = (parentId: string, name: string): TCFolder | undefined =>
    (childrenByParent.get(parentId) ?? []).find((folder) => normalizeLookup(folder.name) === normalizeLookup(name));

  const steps: FolderCreationStep[] = [];
  const resolvedIdByDepth = new Map<number, string>();

  const phaseRoot = findChild(targetParentId, matrix.sheetName);
  steps.push({
    rowNumber: 0,
    depth: 0,
    name: matrix.sheetName,
    path: matrix.sheetName,
    status: phaseRoot ? "existing" : "create",
    folderId: phaseRoot?.id,
  });
  if (phaseRoot) resolvedIdByDepth.set(-1, phaseRoot.id);

  matrix.rows.forEach((row) => {
    const parentId = row.depth === 0 ? resolvedIdByDepth.get(-1) : resolvedIdByDepth.get(row.depth - 1);
    const existing = parentId ? findChild(parentId, row.folderName) : undefined;

    steps.push({
      rowNumber: row.rowNumber,
      depth: row.depth + 1,
      name: row.folderName,
      path: row.folderPath,
      status: existing ? "existing" : "create",
      folderId: existing?.id,
    });

    if (existing) resolvedIdByDepth.set(row.depth, existing.id);
    else resolvedIdByDepth.delete(row.depth);
  });

  return steps;
}
