/**
 * useApi.ts — Trimble Connect Workspace API Hook
 * Docs Viewer:  https://components.connect.trimble.com/trimble-connect-workspace-api/index.html
 * Docs Project: https://components.connect.trimble.com/trimble-connect-project-workspace-api/docs/index.html
 */
import { useState, useEffect, useRef } from "react";
import * as WorkspaceAPI from "trimble-connect-workspace-api";
import type { TCModell, TCProjekt, ViewerState } from "../types";

export function parseObjectIds(rohe: any): number[] {
  if (!Array.isArray(rohe)) return [];
  const ids: number[] = [];
  for (const item of rohe) {
    if (Array.isArray(item?.objects)) {
      for (const o of item.objects) { const n = Number(o?.id ?? o); if (!isNaN(n)) ids.push(n); }
    } else if (typeof item === "number") {
      ids.push(item);
    } else if (item?.id != null) {
      const n = Number(item.id); if (!isNaN(n)) ids.push(n);
    }
  }
  return ids;
}

export function parseSelectionIds(data: any): number[] {
  if (!data) return [];
  const outer = Array.isArray(data) ? data : Array.isArray(data?.data) ? data.data : [];
  const ids: number[] = [];
  for (const item of outer) {
    if (Array.isArray(item?.objectRuntimeIds)) {
      for (const id of item.objectRuntimeIds) { const n = Number(id); if (!isNaN(n) && n >= 0) ids.push(n); }
      continue;
    }
    if (typeof item === "number") { ids.push(item); continue; }
    if (item != null && typeof item === "object") {
      for (const k of ["id", "entityId", "runtimeId", "objectRuntimeId"]) {
        if (item[k] != null) { const n = Number(item[k]); if (!isNaN(n)) { ids.push(n); break; } }
      }
    }
  }
  return ids;
}

export function useApi() {
  const [api, setApi] = useState<any>(null);
  const [connected, setConnected] = useState(false);
  const [isViewerContext, setIsViewerContext] = useState(false);
  const [accessToken, setAccessToken] = useState("");
  const [projekt, setProjekt] = useState<TCProjekt | null>(null);
  const [viewerState, setViewerState] = useState<ViewerState>({ selektion: [], aktivesModellId: "", modelle: [] });
  const apiRef = useRef<any>(null);
  const viewerBestaetigt = useRef(false);

  async function ladeModelle(instance: any): Promise<TCModell[]> {
    try {
      const res = await instance.viewer.getModels();
      const arr = Array.isArray(res) ? res : [];
      return arr.map((m: any) => ({ id: m.modelId || m.id || "", name: m.name || m.fileName || "Modell" })).filter((m: TCModell) => m.id);
    } catch { return []; }
  }

  useEffect(() => {
    async function connect() {
      try {
        const instance = await WorkspaceAPI.connect(
          window.parent,
          async (event: string, data: any) => {
            console.log("TC:", event, JSON.stringify(data)?.slice(0, 80));
            if (["viewer.onCameraChanged","viewer.onSelectionChanged","viewer.onModelStateChanged"].includes(event)) {
              if (!viewerBestaetigt.current) { viewerBestaetigt.current = true; setIsViewerContext(true); }
            }
            if (event === "viewer.onSelectionChanged") {
              setViewerState(prev => ({ ...prev, selektion: parseSelectionIds(data) }));
            }
            if (["viewer.onModelLoaded","viewer.onModelsLoaded","viewer.onModelAdded"].includes(event)) {
              const modelle = await ladeModelle(apiRef.current);
              setViewerState(prev => ({ ...prev, modelle, aktivesModellId: modelle.length > 0 ? (modelle.find(m => m.id === prev.aktivesModellId) ? prev.aktivesModellId : modelle[0].id) : prev.aktivesModellId }));
            }
            if (event === "extension.accessToken") {
              const token = (data as any)?.data || data;
              if (typeof token === "string" && token.length > 10) setAccessToken(token);
            }
          },
          30000
        );
        apiRef.current = instance;

        try { await instance.ui.setMenu({ title: "Meine Extension", icon: "https://deine-vercel-url.vercel.app/icon.svg", command: "open" }); } catch {}
        try { const t = await instance.extension.requestPermission("accesstoken"); if (typeof t === "string" && t.length > 10) setAccessToken(t); } catch {}
        try {
          const proj = await instance.project.getProject() as any;
          setProjekt({ id: proj?.id || "", name: proj?.name || "" });
        } catch {
          try { const proj = await (instance.project as any).getCurrentProject() as any; setProjekt({ id: proj?.id || "", name: proj?.name || "" }); } catch {}
        }

        const modelle = await ladeModelle(instance);
        if (modelle.length > 0) {
          viewerBestaetigt.current = true;
          setIsViewerContext(true);
          setViewerState(prev => ({ ...prev, modelle, aktivesModellId: modelle[0].id }));
        }

        setApi(instance);
        setConnected(true);
      } catch (err) { console.error("TC connect:", err); setConnected(false); }
    }
    connect();
  }, []);

  async function setSelection(ids: number[]) { try { await api?.viewer.setSelection(ids); } catch {} }
  async function clearSelection() { await setSelection([]); }
  async function getObjects(modelId: string): Promise<number[]> {
    try { return parseObjectIds(await api.viewer.getObjects(modelId)); } catch { return []; }
  }
  async function getProperties(modelId: string, ids: number[], batch = 10): Promise<any[]> {
    const all: any[] = [];
    for (let i = 0; i < ids.length; i += batch) {
      try { const r = await api.viewer.getObjectProperties(modelId, ids.slice(i, i + batch)); if (Array.isArray(r)) all.push(...r); } catch {}
    }
    return all;
  }

  return { api, connected, isViewerContext, accessToken, projekt, viewerState, setViewerState, setSelection, clearSelection, getObjects, getProperties };
}
