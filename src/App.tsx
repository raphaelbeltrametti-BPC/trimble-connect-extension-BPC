/**
 * App.tsx — Trimble Connect Extension Template
 *
 * Zeigt automatisch die richtige UI:
 *   - Im 3D Viewer: ViewerApp
 *   - Im Projektbereich / Explorer: PanelApp
 *
 * Für "nur Viewer"-Extensions: PanelApp entfernen und
 * isViewerContext-Logik weglassen.
 */

import { useApi } from "./hooks/useApi";
import "./index.css";

export default function App() {
  const {
    api,
    connected,
    isViewerContext,
    accessToken,
    projekt,
    viewerState,
    setSelection,
    clearSelection,
    getObjects,
    getProperties,
  } = useApi();

  // ── 3D VIEWER ─────────────────────────────────────────────
  if (isViewerContext) {
    return (
      <div className="app viewer-app">
        <header className="tc-header">
          <div className="tc-header-left">
            <div className="tc-logo">TC</div>
            <span className="tc-header-title">Meine Extension</span>
          </div>
          <div className="tc-header-right">
            <span className={`tc-dot ${connected ? "on" : "off"}`} />
          </div>
        </header>

        <div className="tc-content">
          <div className="tc-info-banner">
            ✅ 3D Viewer verbunden
            {viewerState.modelle.length > 0 && (
              <div>Modelle: {viewerState.modelle.map(m => m.name).join(", ")}</div>
            )}
            {viewerState.selektion.length > 0 && (
              <div>{viewerState.selektion.length} Objekte ausgewählt</div>
            )}
          </div>

          {/* Hier Viewer-spezifische UI einfügen */}
          <div className="tc-placeholder">
            <div className="tc-placeholder-icon">🏗️</div>
            <div className="tc-placeholder-title">Viewer-Bereich</div>
            <div className="tc-placeholder-sub">
              Hier kommt die Extension-UI für den 3D Viewer
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── PROJEKTPANEL / EXPLORER ────────────────────────────────
  return (
    <div className="app panel-app">
      <header className="tc-header">
        <div className="tc-header-left">
          <div className="tc-logo">TC</div>
          <span className="tc-header-title">Meine Extension</span>
        </div>
        <div className="tc-header-right">
          <span className={`tc-dot ${connected ? "on" : "off"}`} />
        </div>
      </header>

      <div className="tc-content">
        {projekt && (
          <div className="tc-info-banner">
            📁 Projekt: {projekt.name}
          </div>
        )}

        {/* Hier Panel-spezifische UI einfügen */}
        <div className="tc-placeholder">
          <div className="tc-placeholder-icon">📋</div>
          <div className="tc-placeholder-title">Projektbereich</div>
          <div className="tc-placeholder-sub">
            Hier kommt die Extension-UI für den Projektbereich
          </div>
        </div>
      </div>
    </div>
  );
}
