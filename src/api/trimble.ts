import type { ApiRegion, FolderAcl, TCFolder, TCGroup, TCProject } from "../types";

export const API_BASE_URLS: Record<Exclude<ApiRegion, "custom">, string> = {
  us: "https://app.connect.trimble.com/tc/api/2.0",
  eu: "https://app21.connect.trimble.com/tc/api/2.0",
  uk: "https://app22.connect.trimble.com/tc/api/2.0",
  ap: "https://app31.connect.trimble.com/tc/api/2.0",
  ap2: "https://app32.connect.trimble.com/tc/api/2.0",
};

interface TrimbleClientOptions {
  baseUrl: string;
  getToken: () => Promise<string>;
  delayMs?: number;
}

interface RawFolderItem {
  id?: string;
  nm?: string;
  name?: string;
  tp?: string;
  type?: string;
  pid?: string;
  parentId?: string;
}

interface FolderItemsResponse {
  items?: RawFolderItem[];
  next?: string;
  links?: {
    next?: {
      href?: string;
    };
  };
}

export class TrimbleApiError extends Error {
  readonly status: number;
  readonly body: string;

  constructor(status: number, message: string, body: string) {
    super(message);
    this.name = "TrimbleApiError";
    this.status = status;
    this.body = body;
  }
}

export class TrimbleClient {
  private readonly baseUrl: string;
  private readonly getToken: () => Promise<string>;
  private readonly delayMs: number;
  private lastRequestAt = 0;

  constructor(options: TrimbleClientOptions) {
    this.baseUrl = options.baseUrl.replace(/\/$/, "");
    this.getToken = options.getToken;
    this.delayMs = options.delayMs ?? 250;
  }

  async getProject(projectId: string): Promise<TCProject> {
    const project = await this.request<any>(`/projects/${encodeURIComponent(projectId)}`);

    return {
      id: String(project.id ?? projectId),
      name: String(project.name ?? ""),
      location: project.location,
      rootId: project.rootId,
    };
  }

  async listGroups(projectId: string): Promise<TCGroup[]> {
    const groups = await this.request<any[]>(`/groups?projectId=${encodeURIComponent(projectId)}`);

    return (Array.isArray(groups) ? groups : []).map((group) => ({
      id: String(group.id ?? ""),
      name: String(group.name ?? ""),
      projectId: group.projectId,
      usersCount: typeof group.usersCount === "number" ? group.usersCount : undefined,
    })).filter((group) => group.id && group.name);
  }

  async createGroup(projectId: string, name: string): Promise<TCGroup> {
    const group = await this.request<any>("/groups", {
      method: "POST",
      body: JSON.stringify({ projectId, name }),
    });

    return {
      id: String(group.id ?? ""),
      name: String(group.name ?? name),
      projectId: group.projectId ?? projectId,
      usersCount: typeof group.usersCount === "number" ? group.usersCount : undefined,
    };
  }

  async deleteGroup(groupId: string): Promise<void> {
    await this.request(`/groups/${encodeURIComponent(groupId)}`, { method: "DELETE" });
  }

  async listFolderTree(rootFolderId: string, maxDepth = 20): Promise<TCFolder[]> {
    const folders: TCFolder[] = [];
    const queue: Array<{ id: string; path: string; depth: number }> = [{ id: rootFolderId, path: "", depth: 0 }];

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current || current.depth > maxDepth) continue;

      const children = await this.listFolderItems(current.id);
      for (const child of children) {
        if (!isFolder(child)) continue;

        const name = String(child.nm ?? child.name ?? "");
        const id = String(child.id ?? "");
        if (!id || !name) continue;

        const path = current.path ? `${current.path}/${name}` : name;
        const folder: TCFolder = {
          id,
          name,
          type: String(child.tp ?? child.type ?? "FOLDER"),
          parentId: child.pid ?? child.parentId,
          path,
          depth: current.depth,
        };

        folders.push(folder);
        queue.push({ id, path, depth: current.depth + 1 });
      }
    }

    return folders;
  }

  async updateFolderPermissions(folderId: string, acl: FolderAcl, inheritance = false): Promise<void> {
    await this.request(`/folders/fs/${encodeURIComponent(folderId)}/permissions`, {
      method: "PATCH",
      body: JSON.stringify({ acl, inheritance }),
    });
  }

  private async listFolderItems(folderId: string): Promise<RawFolderItem[]> {
    const items: RawFolderItem[] = [];
    let next: string | undefined = `/folders/${encodeURIComponent(folderId)}/items`;

    while (next) {
      const response: FolderItemsResponse = await this.request<FolderItemsResponse>(next);
      items.push(...(response.items ?? []));
      next = response.next ?? response.links?.next?.href;
    }

    return items;
  }

  private async request<T>(pathOrUrl: string, init: RequestInit = {}): Promise<T> {
    await this.waitForSlot();

    const token = await this.getToken();
    if (!token) {
      throw new TrimbleApiError(401, "Kein Access Token vorhanden.", "");
    }

    const url = pathOrUrl.startsWith("http") ? pathOrUrl : `${this.baseUrl}${pathOrUrl}`;
    const response = await fetch(url, {
      ...init,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
        ...(init.headers ?? {}),
      },
    });

    const body = await response.text();
    if (!response.ok) {
      throw new TrimbleApiError(response.status, `${response.status} ${response.statusText}`, body);
    }

    if (!body) return undefined as T;

    try {
      return JSON.parse(body) as T;
    } catch {
      return body as T;
    }
  }

  private async waitForSlot(): Promise<void> {
    const elapsed = Date.now() - this.lastRequestAt;
    const waitMs = Math.max(0, this.delayMs - elapsed);
    if (waitMs > 0) {
      await new Promise((resolve) => window.setTimeout(resolve, waitMs));
    }
    this.lastRequestAt = Date.now();
  }
}

function isFolder(item: RawFolderItem): boolean {
  const type = String(item.tp ?? item.type ?? "").toUpperCase();
  return type === "FOLDER" || type === "DIR" || type === "DIRECTORY";
}
