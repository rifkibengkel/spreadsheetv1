import React, { useState } from "react";
import { exportWorkbookToExcel } from "@/lib/export/excel";
import {
  exportActiveSheetToCSV,
  triggerServerDatasetCsvDownload,
  saveServerDatasetCsvToDesktop,
} from "@/lib/export/csv";

interface ExportModalProps {
  univerAPI: any;
  onClose: () => void;
  activeFileName?: string;
  activeFileFormat?: "xlsx" | "csv";
  datasetId?: string | null;
  datasetRowCount?: number;
  activeQuery?: {
    queryId: string;
    matchedRows: number;
    description?: string;
  } | null;
}

const ROWS_PER_LOGICAL_SHEET = 4_000_000;

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024)
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

function sanitizeExportFilename(
  rawInput: string,
  format: "xlsx" | "csv",
): string {
  let clean = (rawInput || "").trim();

  // Strip illegal Windows filesystem characters and control characters
  clean = clean.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");

  // Strip trailing spaces and dots (invalid on Windows filesystem)
  clean = clean.replace(/[. ]+$/, "");

  // Strip existing extension matching target format (case-insensitive)
  const extRegex = new RegExp(`\\.${format}$`, "i");
  clean = clean.replace(extRegex, "");

  if (!clean || clean.trim() === "") {
    clean = `workbook-export-${Date.now()}`;
  }

  return `${clean}.${format}`;
}

export default function ExportModal({
  univerAPI,
  onClose,
  activeFileName,
  activeFileFormat = "xlsx",
  datasetId,
  datasetRowCount,
  activeQuery,
}: ExportModalProps) {
  const [format, setFormat] = useState<"xlsx" | "csv">(activeFileFormat);
  const [filenameInput, setFilenameInput] = useState(
    () => activeFileName || `workbook-export.${activeFileFormat}`,
  );
  const [loading, setLoading] = useState(false);
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState("");
  const [error, setError] = useState("");
  const [successMessage, setSuccessMessage] = useState("");

  // Determine active virtual sheet controller (Server or OPFS) and partition information
  const serverController =
    typeof window !== "undefined"
      ? (window as any).__datasetViewportController
      : null;
  const opfsController =
    typeof window !== "undefined"
      ? (window as any).__viewportController
      : null;
  const controller = serverController || opfsController;
  const effectiveDatasetId = datasetId || (serverController as any)?.datasetId || null;
  const totalRows =
    datasetRowCount ||
    (serverController?.getTotalRows ? serverController.getTotalRows() : 0) ||
    opfsController?.getMetadata?.()?.rowCount ||
    0;
  const activeSheetIndex = controller?.getActiveSheetIndex
    ? controller.getActiveSheetIndex()
    : 0;

  const isVirtualDataset = !!(serverController || opfsController || effectiveDatasetId || totalRows > 0);
  const hasVirtualSheets = totalRows > ROWS_PER_LOGICAL_SHEET;
  const virtualSheetCount = hasVirtualSheets
    ? Math.max(1, Math.ceil(totalRows / ROWS_PER_LOGICAL_SHEET))
    : 1;

  // Compute rows for each virtual sheet (or use controller partitions if available)
  const sheetPartitions: Array<{
    index: number;
    name: string;
    rowCount: number;
  }> = [];
  const existingPartitions: any[] = controller?.getPartitions?.() || [];
  if (existingPartitions.length > 0) {
    existingPartitions.forEach((part: any) => {
      sheetPartitions.push({
        index: part.sheetIndex,
        name: part.name || `Sheet ${part.sheetIndex + 1}`,
        rowCount: part.rowCount,
      });
    });
  } else {
    for (let s = 0; s < virtualSheetCount; s++) {
      const sStart = s * ROWS_PER_LOGICAL_SHEET;
      const sEnd = Math.min(totalRows - 1, (s + 1) * ROWS_PER_LOGICAL_SHEET - 1);
      const rCount = Math.max(0, sEnd - sStart + 1);
      sheetPartitions.push({
        index: s,
        name: `Sheet ${s + 1}`,
        rowCount: rCount,
      });
    }
  }

  const [csvScopeSelection, setCsvScopeSelection] = useState<string>("current");

  // Update extension if user changes format manually
  const handleFormatChange = (newFormat: "xlsx" | "csv") => {
    setFormat(newFormat);
    const baseName =
      filenameInput.substring(0, filenameInput.lastIndexOf(".")) ||
      filenameInput;
    setFilenameInput(`${baseName}.${newFormat}`);
  };

  const resolveScopeAndSheetIndex = (): {
    scope: "sheet" | "all";
    sheetIndex: number;
  } => {
    if (csvScopeSelection === "all") {
      return { scope: "all", sheetIndex: 0 };
    }
    if (csvScopeSelection === "current") {
      return { scope: "sheet", sheetIndex: activeSheetIndex };
    }
    if (csvScopeSelection.startsWith("sheet-")) {
      const idx = parseInt(csvScopeSelection.split("-")[1], 10);
      return { scope: "sheet", sheetIndex: isNaN(idx) ? 0 : idx };
    }
    return { scope: "sheet", sheetIndex: activeSheetIndex };
  };

  const handleDownload = async () => {
    if (!univerAPI && !effectiveDatasetId) return;
    setLoading(true);
    setError("");
    setSuccessMessage("");
    setProgressPercent(0);
    setProgressMessage("Menyiapkan file untuk diekspor...");

    const targetName = sanitizeExportFilename(filenameInput, format);

    try {
      // 1. Server-backed Dataset CSV Streaming Export (Zero JS heap overhead)
      if (format === "csv" && effectiveDatasetId) {
        const { scope, sheetIndex } = resolveScopeAndSheetIndex();
        setProgressPercent(25);
        setProgressMessage("Menyiapkan streaming data CSV dari server...");

        await new Promise((r) => setTimeout(r, 200));
        setProgressPercent(75);
        setProgressMessage("Memulai pengunduhan file CSV...");

        triggerServerDatasetCsvDownload({
          datasetId: effectiveDatasetId,
          scope,
          sheetIndex,
          queryId: activeQuery?.queryId,
          filename: targetName,
        });

        setProgressPercent(100);
        setProgressMessage("✓ Pengunduhan berhasil dimulai di browser!");
        setTimeout(onClose, 1200);
        return;
      }

      // 2. XLSX or In-Memory Sheet Export via Web Worker Blob
      let blob: Blob;
      if (format === "xlsx") {
        blob = await exportWorkbookToExcel(univerAPI, (percent, msg) => {
          setProgressPercent(percent);
          setProgressMessage(msg);
        });
      } else {
        const { scope, sheetIndex } = resolveScopeAndSheetIndex();
        blob = await exportActiveSheetToCSV(
          univerAPI,
          (percent, msg) => {
            setProgressPercent(percent);
            setProgressMessage(msg);
          },
          { scope, sheetIndex }
        );
      }

      setProgressPercent(95);
      setProgressMessage("Menyiapkan file untuk diunduh...");

      const mimeType =
        format === "xlsx"
          ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          : "text/csv;charset=utf-8;";
      const file = new File([blob], targetName, {
        type: mimeType,
        lastModified: Date.now(),
      });
      const url = URL.createObjectURL(file);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;
      a.download = targetName;
      document.body.appendChild(a);
      a.click();

      setTimeout(() => {
        try {
          if (document.body.contains(a)) {
            document.body.removeChild(a);
          }
          URL.revokeObjectURL(url);
        } catch {}
      }, 120000);

      setProgressPercent(100);
      setProgressMessage("✓ Pengunduhan selesai!");
      setTimeout(onClose, 1000);
    } catch (e: any) {
      console.error("[EXPORT MODAL] Export failed:", e);
      setError(e.message || "Ekspor gagal.");
      setLoading(false);
    }
  };

  const handleSaveToDesktop = async () => {
    if (!univerAPI && !effectiveDatasetId) return;
    setLoading(true);
    setError("");
    setSuccessMessage("");
    setProgressPercent(0);
    setProgressMessage("Menyiapkan proses simpan ke Desktop...");

    const targetName = sanitizeExportFilename(filenameInput, format);

    try {
      // 1. Server-backed Dataset direct filesystem save (Fast, zero 50MB HTTP upload limit)
      if (format === "csv" && effectiveDatasetId) {
        const { scope, sheetIndex } = resolveScopeAndSheetIndex();
        setProgressPercent(20);
        setProgressMessage("Menghubungi server untuk ekspor dataset...");

        const result = await saveServerDatasetCsvToDesktop({
          datasetId: effectiveDatasetId,
          scope,
          sheetIndex,
          queryId: activeQuery?.queryId,
          filename: targetName,
        });

        setProgressPercent(100);
        setProgressMessage("✓ Berhasil disimpan ke Desktop!");
        setSuccessMessage(
          `✓ Disimpan ke Desktop: ${result.filename} (${result.rowCount.toLocaleString()} baris)`,
        );
        setLoading(false);
        setTimeout(onClose, 1600);
        return;
      }

      // 2. XLSX or In-Memory Sheet via /api/desktop/save
      let blob: Blob;
      if (format === "xlsx") {
        blob = await exportWorkbookToExcel(univerAPI, (percent, msg) => {
          setProgressPercent(percent);
          setProgressMessage(msg);
        });
      } else {
        const { scope, sheetIndex } = resolveScopeAndSheetIndex();
        blob = await exportActiveSheetToCSV(
          univerAPI,
          (percent, msg) => {
            setProgressPercent(percent);
            setProgressMessage(msg);
          },
          { scope, sheetIndex }
        );
      }

      setProgressPercent(90);
      setProgressMessage("Mengirim file ke Desktop...");

      const formData = new FormData();
      formData.append("file", blob, targetName);
      formData.append("filename", targetName);

      const res = await fetch("/api/desktop/save", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();
      if (!res.ok || data.error) {
        throw new Error(data.error || "Gagal menyimpan ke Desktop.");
      }

      setProgressPercent(100);
      setProgressMessage("✓ Berhasil disimpan ke Desktop!");
      setSuccessMessage(
        `✓ Berhasil disimpan ke Desktop: ${data.filePath || targetName}`,
      );
      setLoading(false);
      setTimeout(onClose, 1500);
    } catch (e: any) {
      console.error("[EXPORT MODAL] Save to Desktop failed:", e);
      setError(e.message || "Gagal menyimpan ke Desktop.");
      setLoading(false);
    }
  };

  const activePartitionInfo =
    sheetPartitions[activeSheetIndex] || sheetPartitions[0];

  return (
    <div id="export-modal-overlay" style={overlayStyle}>
      <div style={modalStyle}>
        <h2 style={{ margin: "0 0 8px 0", fontSize: "20px" }}>
          Ekspor & Simpan Data
        </h2>
        <p
          style={{
            color: "#94a3b8",
            fontSize: "13px",
            marginBottom: "16px",
            lineHeight: "1.4",
          }}
        >
          Simpan data spreadsheet ke Desktop atau unduh langsung sebagai file
          Excel / CSV.
        </p>

        <div style={{ marginBottom: "14px" }}>
          <label
            style={{
              display: "block",
              marginBottom: "6px",
              fontSize: "13px",
              fontWeight: "bold",
            }}
          >
            Nama File:
          </label>
          <input
            type="text"
            value={filenameInput}
            onChange={(e) => setFilenameInput(e.target.value)}
            disabled={loading}
            style={inputStyle}
          />
        </div>

        <div style={{ marginBottom: "16px" }}>
          <label
            style={{
              display: "block",
              marginBottom: "6px",
              fontSize: "13px",
              fontWeight: "bold",
            }}
          >
            Format File:
          </label>
          <select
            id="export-format-select"
            value={format}
            onChange={(e) => handleFormatChange(e.target.value as any)}
            style={selectStyle}
            disabled={loading}
          >
            <option value="xlsx">Excel Workbook (.xlsx)</option>
            <option value="csv">CSV Spreadsheet (.csv)</option>
          </select>
        </div>

        {/* CSV Scope Selection for Partitioned/Virtual Datasets */}
        {format === "csv" && isVirtualDataset && sheetPartitions.length > 0 && (
          <div
            style={{
              marginBottom: "16px",
              background: "#0f172a",
              padding: "12px",
              borderRadius: "8px",
              border: "1px solid #334155",
            }}
          >
            <label
              style={{
                display: "block",
                marginBottom: "6px",
                fontSize: "13px",
                fontWeight: "bold",
                color: "#38bdf8",
              }}
            >
              Cakupan Ekspor CSV:
            </label>
            <select
              id="export-csv-scope-select"
              value={csvScopeSelection}
              onChange={(e) => setCsvScopeSelection(e.target.value)}
              style={selectStyle}
              disabled={loading}
            >
              <option value="current">
                Sheet Aktif Saat Ini ({activePartitionInfo?.name || "Sheet 1"} —{" "}
                {activePartitionInfo?.rowCount.toLocaleString() || "4,000,000"}{" "}
                baris)
              </option>
              {hasVirtualSheets &&
                sheetPartitions.map((part) => (
                  <option key={part.index} value={`sheet-${part.index}`}>
                    {part.name} Saja ({part.rowCount.toLocaleString()} baris)
                  </option>
                ))}
              <option value="all">
                Seluruh Dataset ({totalRows.toLocaleString()} baris)
              </option>
            </select>
            {activeQuery && (
              <p
                style={{
                  color: "#a78bfa",
                  fontSize: "11px",
                  margin: "6px 0 0 0",
                }}
              >
                ℹ Filter aktif: {activeQuery.matchedRows.toLocaleString()} baris
                cocok
              </p>
            )}
          </div>
        )}

        {loading && (
          <div
            style={{
              marginBottom: "16px",
              background: "#0f172a",
              padding: "12px",
              borderRadius: "8px",
              border: "1px solid #334155",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "6px",
              }}
            >
              <span
                style={{
                  color: "#38bdf8",
                  fontSize: "12px",
                  fontWeight: "bold",
                }}
              >
                Proses Ekspor / Simpan File...
              </span>
              <span
                id="export-progress-percent"
                style={{
                  color: "#38bdf8",
                  fontSize: "12px",
                  fontWeight: "bold",
                }}
              >
                {progressPercent}%
              </span>
            </div>
            <div
              style={{
                width: "100%",
                height: "8px",
                background: "#1e293b",
                borderRadius: "4px",
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  height: "100%",
                  width: `${progressPercent}%`,
                  background:
                    "linear-gradient(90deg, #38bdf8 0%, #818cf8 100%)",
                  transition: "width 0.2s ease-out",
                }}
              />
            </div>
            <p
              id="export-status-message"
              style={{
                color: "#cbd5e1",
                fontSize: "11px",
                marginTop: "6px",
                margin: "6px 0 0 0",
                whiteSpace: "nowrap",
                overflow: "hidden",
                textOverflow: "ellipsis",
              }}
            >
              {progressMessage || "Memproses data..."}
            </p>
          </div>
        )}

        {error && (
          <p
            style={{ color: "#ef4444", marginBottom: "12px", fontSize: "13px" }}
          >
            {error}
          </p>
        )}
        {successMessage && (
          <p
            style={{
              color: "#10b981",
              marginBottom: "12px",
              fontWeight: "bold",
              fontSize: "13px",
            }}
          >
            {successMessage}
          </p>
        )}

        <div
          style={{
            display: "flex",
            gap: "10px",
            justifyContent: "flex-end",
            marginTop: "20px",
            flexWrap: "wrap",
          }}
        >
          <button
            id="export-cancel-btn"
            onClick={onClose}
            disabled={loading}
            style={cancelBtnStyle}
          >
            Batal
          </button>
          <button
            id="export-download-btn"
            onClick={handleDownload}
            disabled={loading}
            style={exportBtnStyle}
          >
            📥 Unduh File
          </button>
          {/* <button id="export-desktop-btn" onClick={handleSaveToDesktop} disabled={loading} style={desktopBtnStyle}>
            💾 Simpan ke Desktop
          </button> */}
        </div>
      </div>
    </div>
  );
}

const overlayStyle: React.CSSProperties = {
  position: "fixed",
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  background: "rgba(0, 0, 0, 0.75)",
  backdropFilter: "blur(6px)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  zIndex: 1000,
};

const modalStyle: React.CSSProperties = {
  background: "#0f172a",
  padding: "24px",
  borderRadius: "12px",
  width: "480px",
  maxWidth: "92vw",
  boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.7)",
  color: "#f8fafc",
  border: "1px solid #1e293b",
};

const selectStyle: React.CSSProperties = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: "7px",
  background: "#1e293b",
  border: "1px solid #334155",
  color: "#f8fafc",
  fontSize: "13px",
  outline: "none",
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: "7px",
  background: "#1e293b",
  border: "1px solid #334155",
  color: "#f8fafc",
  boxSizing: "border-box" as const,
  fontSize: "13px",
  outline: "none",
};

const cancelBtnStyle: React.CSSProperties = {
  padding: "9px 16px",
  borderRadius: "7px",
  border: "1px solid #334155",
  background: "transparent",
  color: "#cbd5e1",
  cursor: "pointer",
  fontSize: "13px",
  fontWeight: "600",
  transition: "all 0.15s ease",
};

const exportBtnStyle: React.CSSProperties = {
  padding: "9px 18px",
  borderRadius: "7px",
  border: "none",
  background: "linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)",
  color: "#ffffff",
  fontWeight: "600",
  cursor: "pointer",
  fontSize: "13px",
  boxShadow: "0 2px 8px rgba(37, 99, 235, 0.35)",
  transition: "all 0.15s ease",
};

const desktopBtnStyle: React.CSSProperties = {
  padding: "9px 16px",
  borderRadius: "7px",
  border: "none",
  background: "#0284c7",
  color: "#ffffff",
  fontWeight: "600",
  cursor: "pointer",
  fontSize: "13px",
  transition: "all 0.15s ease",
};
