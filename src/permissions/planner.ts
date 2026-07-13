import type { FolderAcl, PermissionMatrix, PermissionPlanItem, TCFolder, TCGroup } from "../types";
import { actorIdForGroup, normalizeLookup, normalizePath } from "../utils/text";

export function buildPermissionPlan(
  matrix: PermissionMatrix,
  groups: TCGroup[],
  folders: TCFolder[]
): PermissionPlanItem[] {
  const groupByName = new Map(groups.map((group) => [normalizeLookup(group.name), group]));
  const foldersByPath = new Map(folders.map((folder) => [normalizePath(folder.path), folder]));
  const foldersBySuffix = buildSuffixIndex(folders);

  return matrix.rows.map((row) => {
    const folder = findFolder(row.folderPath, row.relativePath, foldersByPath, foldersBySuffix);
    const acl: FolderAcl = { READ: [], FULL_ACCESS: [], NO_ACCESS: [] };
    const missingGroups: string[] = [];

    matrix.teams.forEach((teamName) => {
      const group = groupByName.get(normalizeLookup(teamName));
      if (!group) {
        missingGroups.push(teamName);
        return;
      }

      acl[row.permissions[teamName]].push(actorIdForGroup(group.id));
    });

    return {
      rowNumber: row.rowNumber,
      folderPath: row.folderPath,
      folderName: row.folderName,
      folderId: folder?.id,
      status: !folder ? "missing-folder" : missingGroups.length > 0 ? "missing-groups" : "ready",
      acl,
      missingGroups,
      counts: {
        fullAccess: acl.FULL_ACCESS.length,
        read: acl.READ.length,
        noAccess: acl.NO_ACCESS.length,
      },
    };
  });
}

/**
 * Maps every trailing sub-path ("suffix") of each folder's full path to the folders that end
 * with it, e.g. "a/b/c" contributes keys "a/b/c", "b/c" and "c". This lets rows from the Excel
 * matrix match real Trimble folders even when the matrix's assumed root (the sheet/phase name)
 * is nested deeper in the actual project than the matrix author expected.
 */
function buildSuffixIndex(folders: TCFolder[]): Map<string, TCFolder[]> {
  const index = new Map<string, TCFolder[]>();

  folders.forEach((folder) => {
    const segments = normalizePath(folder.path).split("/").filter(Boolean);
    for (let start = 0; start < segments.length; start += 1) {
      const suffix = segments.slice(start).join("/");
      index.set(suffix, [...(index.get(suffix) ?? []), folder]);
    }
  });

  return index;
}

function findFolder(
  folderPath: string,
  relativePath: string,
  foldersByPath: Map<string, TCFolder>,
  foldersBySuffix: Map<string, TCFolder[]>
): TCFolder | undefined {
  const direct = foldersByPath.get(normalizePath(folderPath));
  if (direct) return direct;

  const withoutPhase = foldersByPath.get(normalizePath(relativePath));
  if (withoutPhase) return withoutPhase;

  return findBySuffix(relativePath, foldersBySuffix);
}

/**
 * Tries the full relative path first, then progressively drops the outermost (leftmost)
 * segment, returning the first length with exactly one match. Since a shorter suffix can only
 * match the same folders as a longer one plus more, ambiguity never resolves by shortening
 * further, so we stop as soon as more than one folder matches.
 */
function findBySuffix(relativePath: string, foldersBySuffix: Map<string, TCFolder[]>): TCFolder | undefined {
  const segments = normalizePath(relativePath).split("/").filter(Boolean);

  for (let start = 0; start < segments.length; start += 1) {
    const suffix = segments.slice(start).join("/");
    const matches = foldersBySuffix.get(suffix) ?? [];
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) return undefined;
  }

  return undefined;
}
