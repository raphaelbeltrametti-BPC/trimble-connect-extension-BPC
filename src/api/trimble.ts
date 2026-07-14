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

  /**
   * Performs one raw request and returns status/body verbatim (even on non-2xx), so the caller
   * can surface exactly what the Trimble API sent back (e.g. in the UI log) without needing
   * browser devtools access to the extension's iframe.
   */
  async fetchRawDebug(pathOrUrl: string): Promise<{ url: string; status: number; body: string }> {
    await this.waitForSlot();
    const token = await this.getToken();
    const url = pathOrUrl.startsWith("http") ? pathOrUrl : `${this.baseUrl}${pathOrUrl}`;
    const response = await fetch(url, {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
    });
    const body = await response.text();
    return { url, status: response.status, body };
  }

  /**
   * Scans the folder tree level by level, firing up to `concurrency` sibling requests at once
   * per level instead of one request at a time. With trees of several hundred folders spread
   * across many phases, a fully sequential scan can take minutes; batching cuts that roughly by
   * the concurrency factor while still respecting the client's per-request pacing.
   */
  async listFolderTree(
    rootFolderId: string,
    options: { maxDepth?: number; concurrency?: number; onProgress?: (scanned: number, found: number) => void } = {}
  ): Promise<TCFolder[]> {
    const { maxDepth = 20, concurrency = 5, onProgress } = options;
    const folders: TCFolder[] = [];
    let currentLevel: Array<{ id: string; path: string; depth: number }> = [
      { id: rootFolderId, path: "", depth: 0 },
    ];
    let scanned = 0;

    while (currentLevel.length > 0) {
      const nextLevel: typeof currentLevel = [];

      for (let i = 0; i < currentLevel.length; i += concurrency) {
        const batch = currentLevel.slice(i, i + concurrency);
        const batchChildren = await Promise.all(batch.map((node) => this.listFolderItems(node.id)));

        batch.forEach((node, index) => {
          scanned += 1;
          if (node.depth > maxDepth) return;

          for (const child of batchChildren[index]) {
            if (!isFolder(child)) continue;

            const name = String(child.nm ?? child.name ?? "");
            const id = String(child.id ?? "");
            if (!id || !name) continue;

            const path = node.path ? `${node.path}/${name}` : name;
            folders.push({
              id,
              name,
              type: String(child.tp ?? child.type ?? "FOLDER"),
              parentId: child.pid ?? child.parentId,
              path,
              depth: node.depth,
            });
            nextLevel.push({ id, path, depth: node.depth + 1 });
          }
        });

        onProgress?.(scanned, folders.length);
      }

      currentLevel = nextLevel;
    }

    return folders;
  }

  async updateFolderPermissions(folderId: string, acl: FolderAcl, inheritance = false): Promise<void> {
    await this.request(`/folders/fs/${encodeURIComponent(folderId)}/permissions`, {
      method: "PATCH",
      body: JSON.stringify({ acl, inheritance }),
    });
  }

  /**
   * The `/folders/{id}/items` endpoint returns a plain JSON array of items, not an object with
   * an `items` property. Pagination (if any) comes via an RFC5988 `Link: <url>; rel="next"`
   * response header rather than a body field, so this uses requestRaw() to see headers too.
   */
  private async listFolderItems(folderId: string): Promise<RawFolderItem[]> {
    const items: RawFolderItem[] = [];
    let next: string | undefined = `/folders/${encodeURIComponent(folderId)}/items`;

    while (next) {
      const { body, headers } = await this.requestRaw(next);
      const parsed: FolderItemsResponse | RawFolderItem[] = body ? JSON.parse(body) : [];
      const pageItems = Array.isArray(parsed) ? parsed : parsed.items ?? [];
      items.push(...pageItems);

      next = parseNextLink(headers) ?? (Array.isArray(parsed) ? undefined : parsed.next ?? parsed.links?.next?.href);
    }

    return items;
  }

  private async request<T>(pathOrUrl: string, init: RequestInit = {}): Promise<T> {
    const { body } = await this.requestRaw(pathOrUrl, init);
    if (!body) return undefined as T;

    try {
      return JSON.parse(body) as T;
    } catch {
      return body as T;
    }
  }

  private async requestRaw(pathOrUrl: string, init: RequestInit = {}): Promise<{ body: string; headers: Headers }> {
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

    return { body, headers: response.headers };
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

function parseNextLink(headers: Headers): string | undefined {
  const link = headers.get("Link") ?? headers.get("link");
  if (!link) return undefined;

  const nextEntry = link.split(",").map((part) => part.trim()).find((part) => /rel="?next"?/i.test(part));
  const match = nextEntry?.match(/^<([^>]+)>/);
  return match?.[1];
}
