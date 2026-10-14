import React, { useState } from "react";
import ImportModal from "./ImportModal";
import ExportModal from "./ExportModal";
import { exportWorkbookToExcel } from "@/lib/export/excel";
import {
  exportActiveSheetToCSV,
  saveServerDatasetCsvToDesktop,
} from "@/lib/export/csv";

interface SpreadsheetToolbarProps {
  univerAPI: any;
  activeFileName: string;
  activeFileFormat: "xlsx" | "csv";
  onFileImported: (
    filename: string,
    format: "xlsx" | "csv",
    metadata?: any,
  ) => void;
  onNewSheet?: () => void;
  onOpenJumpToRow?: () => void;
  datasetRowCount?: number;
  datasetId?: string | null;
  onOpenQueryModal?: () => void;
  activeQuery?: {
    queryId: string;
    matchedRows: number;
    description?: string;
    queryDef?: any;
  } | null;
  onClearQuery?: () => void;
  isQueryLoading?: boolean;
}

const SpreadsheetToolbar = React.memo(function SpreadsheetToolbar({
  univerAPI,
  activeFileName,
  activeFileFormat,
  onFileImported,
  onNewSheet,
  onOpenJumpToRow,
  datasetRowCount,
  datasetId,
  onOpenQueryModal,
  activeQuery,
  onClearQuery,
  isQueryLoading,
}: SpreadsheetToolbarProps) {
  const [showImport, setShowImport] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [toastMessage, setToastMessage] = useState<{
    text: string;
    isError?: boolean;
  } | null>(null);

  const showToast = (text: string, isError = false) => {
    setToastMessage({ text, isError });
    setTimeout(() => {
      setToastMessage(null);
    }, 5000);
  };

  const handleToggleFilter = async () => {
    if (!univerAPI) {
      showToast("Mesin spreadsheet belum siap.", true);
      return;
    }
    try {
      await univerAPI.executeCommand("sheet.command.smart-toggle-filter");
    } catch (err: any) {
      console.warn("Filter toggle error:", err);
      showToast(err?.message || "Gagal mengubah status filter.", true);
    }
  };

  const handleClearFilter = async () => {
    if (!univerAPI) {
      showToast("Mesin spreadsheet belum siap.", true);
      return;
    }
    try {
      await univerAPI.executeCommand("sheet.command.clear-filter-criteria");
      showToast("✓ Filter berhasil dibersihkan");
    } catch (err: any) {
      console.warn("Clear filter error:", err);
    }
  };

  return (
    <>
      <header
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          height: "52px",
          padding: "0 16px",
          background: "#0f172a",
          borderBottom: "1px solid #1e293b",
          zIndex: 30,
          position: "relative",
          flexShrink: 0,
        }}
      >
        {/* Left: Brand Identity & Active File Info */}
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          {/* Brand Logo */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <div
              style={{
                width: "30px",
                height: "30px",
                borderRadius: "8px",
                background: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                boxShadow: "0 2px 6px rgba(16, 185, 129, 0.3)",
              }}
            >
              <svg
                width="18"
                height="18"
                viewBox="0 0 24 24"
                fill="none"
                stroke="#ffffff"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="3" width="18" height="18" rx="2" />
                <line x1="3" y1="9" x2="21" y2="9" />
                <line x1="3" y1="15" x2="21" y2="15" />
                <line x1="9" y1="3" x2="9" y2="21" />
                <line x1="15" y1="3" x2="15" y2="21" />
              </svg>
            </div>
            <span
              style={{
                color: "#f8fafc",
                fontSize: "14px",
                fontWeight: "700",
                letterSpacing: "-0.02em",
              }}
            >
              Redbox Sheet
            </span>
          </div>

          <div
            style={{ width: "1px", height: "20px", background: "#334155" }}
          />

          {/* Active File Pill */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "8px",
              padding: "4px 10px",
              borderRadius: "6px",
              background: "rgba(30, 41, 59, 0.8)",
              border: "1px solid #334155",
            }}
          >
            <span
              style={{
                padding: "2px 6px",
                borderRadius: "4px",
                background:
                  activeFileFormat === "xlsx"
                    ? "rgba(16, 185, 129, 0.15)"
                    : "rgba(56, 189, 248, 0.15)",
                color: activeFileFormat === "xlsx" ? "#34d399" : "#38bdf8",
                fontSize: "11px",
                fontWeight: "700",
                textTransform: "uppercase",
                letterSpacing: "0.05em",
              }}
            >
              {activeFileFormat}
            </span>
            <span
              style={{
                color: "#e2e8f0",
                fontSize: "13px",
                fontWeight: "500",
                maxWidth: "220px",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
              title={activeFileName}
            >
              {activeFileName}
            </span>
            {datasetRowCount && datasetRowCount > 0 ? (
              <span
                style={{
                  padding: "2px 8px",
                  borderRadius: "12px",
                  background: "rgba(99, 102, 241, 0.15)",
                  color: "#a5b4fc",
                  fontSize: "11px",
                  border: "1px solid rgba(99, 102, 241, 0.3)",
                  fontWeight: "600",
                }}
              >
                ⚡ {datasetRowCount.toLocaleString()} baris
              </span>
            ) : null}
          </div>
        </div>

        {/* Center: Tools, Query & Filters */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {/* Active Filter Pill if query is applied */}
          {activeQuery && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: "8px",
                background: "rgba(56, 189, 248, 0.15)",
                border: "1px solid rgba(56, 189, 248, 0.35)",
                padding: "4px 10px",
                borderRadius: "8px",
              }}
            >
              <span
                style={{
                  fontSize: "12px",
                  color: "#38bdf8",
                  fontWeight: "700",
                }}
              >
                🔍 Filter: {activeQuery.matchedRows.toLocaleString()} baris
              </span>
              {onClearQuery && (
                <button
                  onClick={onClearQuery}
                  style={clearFilterBtnStyle}
                  title="Hapus filter/sort server dan kembali ke dataset penuh"
                >
                  ✕ Reset
                </button>
              )}
            </div>
          )}

          {isQueryLoading && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: "6px",
                background: "rgba(251, 191, 36, 0.12)",
                border: "1px solid rgba(251, 191, 36, 0.3)",
                padding: "4px 10px",
                borderRadius: "8px",
              }}
            >
              <span
                style={{
                  fontSize: "12px",
                  animation: "spin-slow 1s infinite linear",
                }}
              >
                ⚡
              </span>
              <span
                style={{
                  fontSize: "12px",
                  color: "#fbbf24",
                  fontWeight: "600",
                }}
              >
                Memproses query lazy load...
              </span>
            </div>
          )}

          {/* Filter & Sort Server Button */}
          {onOpenQueryModal && (
            <button
              onClick={onOpenQueryModal}
              style={{
                ...headerBtnStyle,
                background: activeQuery
                  ? "rgba(56, 189, 248, 0.25)"
                  : "rgba(30, 41, 59, 0.9)",
                borderColor: activeQuery ? "#38bdf8" : "#334155",
                color: activeQuery ? "#38bdf8" : "#f8fafc",
                display: "flex",
                alignItems: "center",
                gap: "6px",
              }}
              title="Buka dialog filter dan sort server (Lazy Streaming)"
            >
              <span style={{ color: "#38bdf8" }}>⚡</span>
              <span>Filter & Sort Server</span>
            </button>
          )}

          {/* Lompat ke Baris (Ctrl+G) Button */}
          {onOpenJumpToRow && (
            <button
              onClick={onOpenJumpToRow}
              style={{
                ...headerBtnStyle,
                display: "flex",
                alignItems: "center",
                gap: "6px",
              }}
              title="Lompat ke baris spesifik di spreadsheet / dataset besar (Ctrl+G)"
            >
              <span style={{ color: "#818cf8" }}>🔍</span>
              <span>Lompat (Ctrl+G)</span>
            </button>
          )}

          {/* Sheet level native filter toggle */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "4px",
              background: "rgba(30, 41, 59, 0.6)",
              padding: "2px 4px",
              borderRadius: "6px",
              border: "1px solid #334155",
            }}
          >
            <button
              onClick={handleToggleFilter}
              style={headerGhostBtnStyle}
              title="Toggle Filter rentang terpilih pada sheet (Ctrl+Shift+L)"
            >
              Filter Sheet
            </button>
            <button
              onClick={handleClearFilter}
              style={headerGhostBtnStyle}
              title="Hapus kriteria filter sheet aktif"
            >
              Clear
            </button>
          </div>
        </div>

        {/* Right: Actions */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {onNewSheet && (
            <button
              onClick={onNewSheet}
              style={newSheetBtnStyle}
              title="Kosongkan dan buat spreadsheet baru yang bersih"
            >
              📄 Lembar Baru
            </button>
          )}

          <button
            onClick={() => setShowImport(true)}
            style={importBtnStyle}
            title="Impor file Excel (.xlsx) atau CSV"
          >
            📥 Impor File
          </button>

          <button
            onClick={() => setShowExport(true)}
            style={exportBtnStyle}
            title="Ekspor spreadsheet ke format Excel (.xlsx) atau CSV"
          >
            📤 Ekspor & Save
          </button>
        </div>
      </header>

      {toastMessage && (
        <div
          style={{
            position: "absolute",
            top: "60px",
            right: "20px",
            zIndex: 100,
            padding: "10px 16px",
            borderRadius: "8px",
            background: toastMessage.isError
              ? "rgba(127, 29, 29, 0.95)"
              : "rgba(6, 78, 59, 0.95)",
            color: toastMessage.isError ? "#fca5a5" : "#6ee7b7",
            fontSize: "13px",
            fontWeight: "600",
            border: toastMessage.isError
              ? "1px solid #991b1b"
              : "1px solid #047857",
            boxShadow: "0 8px 24px rgba(0, 0, 0, 0.5)",
            backdropFilter: "blur(6px)",
          }}
        >
          {toastMessage.text}
        </div>
      )}

      {showImport && (
        <ImportModal
          univerAPI={univerAPI}
          onClose={() => setShowImport(false)}
          onFileImported={onFileImported}
        />
      )}

      {showExport && (
        <ExportModal
          univerAPI={univerAPI}
          onClose={() => setShowExport(false)}
          activeFileName={activeFileName}
          activeFileFormat={activeFileFormat}
          datasetId={datasetId}
          datasetRowCount={datasetRowCount}
          activeQuery={activeQuery}
        />
      )}
    </>
  );
});

export default SpreadsheetToolbar;

const headerBtnStyle: React.CSSProperties = {
  height: "32px",
  padding: "0 12px",
  background: "#1e293b",
  color: "#cbd5e1",
  border: "1px solid #334155",
  borderRadius: "6px",
  cursor: "pointer",
  fontWeight: "600",
  fontSize: "12px",
  display: "inline-flex",
  alignItems: "center",
  gap: "5px",
  transition: "all 0.15s ease",
};

const headerGhostBtnStyle: React.CSSProperties = {
  height: "32px",
  padding: "0 10px",
  background: "transparent",
  color: "#94a3b8",
  border: "none",
  borderRadius: "6px",
  cursor: "pointer",
  fontWeight: "500",
  fontSize: "12px",
  display: "inline-flex",
  alignItems: "center",
  gap: "4px",
  transition: "all 0.15s ease",
};

const clearFilterBtnStyle: React.CSSProperties = {
  height: "26px",
  padding: "0 8px",
  background: "rgba(239, 68, 68, 0.2)",
  color: "#fca5a5",
  border: "1px solid rgba(239, 68, 68, 0.4)",
  borderRadius: "5px",
  cursor: "pointer",
  fontWeight: "600",
  fontSize: "11px",
  display: "inline-flex",
  alignItems: "center",
  gap: "3px",
  transition: "all 0.15s ease",
};

const newSheetBtnStyle: React.CSSProperties = {
  height: "34px",
  padding: "0 12px",
  background: "rgba(51, 65, 85, 0.4)",
  color: "#e2e8f0",
  border: "1px solid #334155",
  borderRadius: "7px",
  cursor: "pointer",
  fontWeight: "600",
  fontSize: "12px",
  display: "inline-flex",
  alignItems: "center",
  gap: "5px",
  transition: "all 0.15s ease",
};

const importBtnStyle: React.CSSProperties = {
  height: "34px",
  padding: "0 14px",
  background: "#1e293b",
  color: "#f8fafc",
  border: "1px solid #475569",
  borderRadius: "7px",
  cursor: "pointer",
  fontWeight: "600",
  fontSize: "13px",
  display: "inline-flex",
  alignItems: "center",
  gap: "6px",
  transition: "all 0.15s ease",
  boxShadow: "0 1px 3px rgba(0, 0, 0, 0.2)",
};

const exportBtnStyle: React.CSSProperties = {
  height: "34px",
  padding: "0 16px",
  background: "linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)",
  color: "#ffffff",
  border: "none",
  borderRadius: "7px",
  cursor: "pointer",
  fontWeight: "600",
  fontSize: "13px",
  display: "inline-flex",
  alignItems: "center",
  gap: "6px",
  transition: "all 0.15s ease",
  boxShadow: "0 2px 8px rgba(37, 99, 235, 0.35)",
};
