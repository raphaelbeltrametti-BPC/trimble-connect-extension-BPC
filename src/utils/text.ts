export function normalizeWhitespace(value: unknown): string {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

export function normalizeLookup(value: unknown): string {
  return normalizeWhitespace(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

export function normalizePath(path: string): string {
  return path
    .split("/")
    .map((part) => normalizeLookup(part))
    .filter(Boolean)
    .join("/");
}

export function leadingSpaceCount(value: string): number {
  const match = value.match(/^\s*/);
  return match ? match[0].length : 0;
}

export function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values.map(normalizeWhitespace).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "de")
  );
}

export function actorIdForGroup(groupId: string): string {
  return groupId.startsWith("tc-groups:") ? groupId : `tc-groups:${groupId}`;
}
