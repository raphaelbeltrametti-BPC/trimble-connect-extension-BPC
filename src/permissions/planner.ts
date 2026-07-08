import type { FolderAcl, PermissionMatrix, PermissionPlanItem, TCFolder, TCGroup } from "../types";
import { actorIdForGroup, normalizeLookup, normalizePath } from "../utils/text";

export function buildPermissionPlan(
  matrix: PermissionMatrix,
  groups: TCGroup[],
  folders: TCFolder[]
): PermissionPlanItem[] {
  const groupByName = new Map(groups.map((group) => [normalizeLookup(group.name), group]));
  const foldersByPath = new Map(folders.map((folder) => [normalizePath(folder.path), folder]));
  const foldersByName = new Map<string, TCFolder[]>();

  folders.forEach((folder) => {
    const key = normalizeLookup(folder.name);
    foldersByName.set(key, [...(foldersByName.get(key) ?? []), folder]);
  });

  return matrix.rows.map((row) => {
    const folder = findFolder(row.folderPath, row.relativePath, row.folderName, foldersByPath, foldersByName);
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

function findFolder(
  folderPath: string,
  relativePath: string,
  folderName: string,
  foldersByPath: Map<string, TCFolder>,
  foldersByName: Map<string, TCFolder[]>
): TCFolder | undefined {
  const direct = foldersByPath.get(normalizePath(folderPath));
  if (direct) return direct;

  const withoutPhase = foldersByPath.get(normalizePath(relativePath));
  if (withoutPhase) return withoutPhase;

  const byName = foldersByName.get(normalizeLookup(folderName)) ?? [];
  return byName.length === 1 ? byName[0] : undefined;
}
