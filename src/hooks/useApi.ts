import { useCallback, useEffect, useRef, useState } from "react";
import * as WorkspaceApiPackage from "trimble-connect-workspace-api";
import type { WorkspaceAPI as TrimbleWorkspaceAPI } from "trimble-connect-workspace-api";
import type { TCProject } from "../types";

export function useApi() {
  const [api, setApi] = useState<TrimbleWorkspaceAPI | null>(null);
  const [connected, setConnected] = useState(false);
  const [connectionError, setConnectionError] = useState("");
  const [accessToken, setAccessToken] = useState("");
  const [project, setProject] = useState<TCProject | null>(null);
  const [hostName, setHostName] = useState("");
  const apiRef = useRef<TrimbleWorkspaceAPI | null>(null);
  const tokenRef = useRef("");
  const didConnect = useRef(false);

  const refreshAccessToken = useCallback(async () => {
    if (!apiRef.current) return tokenRef.current;

    const token = await apiRef.current.extension.requestPermission("accesstoken");
    if (typeof token === "string" && token.length > 0) {
      tokenRef.current = token;
      setAccessToken(token);
    }

    return tokenRef.current;
  }, []);

  useEffect(() => {
    if (didConnect.current) return;
    didConnect.current = true;

    let cancelled = false;

    async function connect() {
      try {
        const instance = await WorkspaceApiPackage.connect(
          window.parent,
          (event: string, payload: any) => {
            if (event === "extension.accessToken") {
              const token = typeof payload?.data === "string" ? payload.data : String(payload ?? "");
              if (token) {
                tokenRef.current = token;
                setAccessToken(token);
              }
            }
          },
          15000
        );

        if (cancelled) return;

        apiRef.current = instance;
        setApi(instance);
        setConnected(true);
        setConnectionError("");

        try {
          const host = await instance.extension.getHost();
          setHostName(host.name);
        } catch {
          setHostName("");
        }

        try {
          await instance.ui.setMenu({
            title: "Team & Permission Manager",
            icon: `${window.location.origin}/icon.svg`,
            command: "open",
          });
        } catch {
          // Menu configuration is optional and may be unavailable in some host contexts.
        }

        try {
          await refreshAccessToken();
        } catch {
          // The UI exposes the missing token state and lets the user retry.
        }

        const currentProject = await instance.project.getProject();
        setProject({
          id: currentProject.id,
          name: currentProject.name ?? currentProject.id,
          location: currentProject.location,
        });
      } catch (error) {
        if (cancelled) return;
        setConnected(false);
        setConnectionError(error instanceof Error ? error.message : "Workspace API Verbindung fehlgeschlagen.");
      }
    }

    connect();

    return () => {
      cancelled = true;
    };
  }, [refreshAccessToken]);

  return {
    api,
    connected,
    connectionError,
    accessToken,
    project,
    hostName,
    refreshAccessToken,
  };
}
