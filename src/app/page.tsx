"use client";

import { useState, useRef, useCallback } from "react";
import {
  FileSpreadsheet,
  UploadCloud,
  Eraser,
  Download,
  CheckCircle2,
  ArrowLeft,
  RefreshCw,
  AlertTriangle,
  X,
  Lock,
  ImageIcon,
} from "lucide-react";
import AnimatedShaderBackground from "@/components/ui/animated-shader-background";

type AppState = "idle" | "uploading" | "selecting" | "processing" | "done";

interface UploadResult {
  sheets: string[];
  fileId: string;
  fileName: string;
  drawingCounts: Record<string, number>;
}

const STEPS: { label: string }[] = [
  { label: "Upload" },
  { label: "Select" },
  { label: "Download" },
];

function getStepStatus(
  appState: AppState,
  stepIdx: number
): "completed" | "active" | "" {
  const map: Record<AppState, number> = {
    idle: 0,
    uploading: 0,
    selecting: 1,
    processing: 1,
    done: 2,
  };
  const cur = map[appState];
  if (stepIdx < cur) return "completed";
  if (stepIdx === cur) return "active";
  return "";
}

export default function Home() {
  const [state, setState] = useState<AppState>("idle");
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [drawingCounts, setDrawingCounts] = useState<Record<string, number>>({});
  const [fileId, setFileId] = useState("");
  const [fileName, setFileName] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");
  const [downloadName, setDownloadName] = useState("");
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    const name = file.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".xlsm")) {
      setError("Only .xlsx and .xlsm files are supported.");
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      setError("File size must be under 50 MB.");
      return;
    }
    setError("");
    setState("uploading");
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch("/api/sheets", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to process file");
      const result = data as UploadResult;
      setSheets(result.sheets);
      setFileId(result.fileId);
      setFileName(result.fileName);
      setDrawingCounts(result.drawingCounts ?? {});
      setSelectedSheets(new Set());
      setState("selecting");
    } catch (err: any) {
      setError(err.message || "Something went wrong");
      setState("idle");
    }
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
  }, []);

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const toggleSheet = useCallback((name: string) => {
    setSelectedSheets((prev) => {
      const next = new Set(prev);
      next.has(name) ? next.delete(name) : next.add(name);
      return next;
    });
  }, []);

  const handleSheetKeyDown = useCallback(
    (e: React.KeyboardEvent, name: string) => {
      if (e.key === " " || e.key === "Enter") {
        e.preventDefault();
        toggleSheet(name);
      }
    },
    [toggleSheet]
  );

  const selectAll = useCallback(
    () => setSelectedSheets(new Set(sheets)),
    [sheets]
  );
  const selectNone = useCallback(() => setSelectedSheets(new Set()), []);

  const handleClean = useCallback(async () => {
    if (selectedSheets.size === 0) {
      setError("Select at least one sheet to clean.");
      return;
    }
    setError("");
    setState("processing");
    try {
      const res = await fetch("/api/clean", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileId,
          selectedSheets: Array.from(selectedSheets),
          fileName,
        }),
      });
      if (!res.ok) {
        const data = await res.json();
        throw new Error(data.error || "Failed to clean file");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const cleaned = fileName
        ? fileName.replace(/\.xlsx$/i, "_cleaned.xlsx")
        : "cleaned.xlsx";
      setDownloadUrl(url);
      setDownloadName(cleaned);
      setState("done");
    } catch (err: any) {
      setError(err.message || "Something went wrong");
      setState("selecting");
    }
  }, [fileId, selectedSheets, fileName]);

  const handleReset = useCallback(() => {
    setState("idle");
    setSheets([]);
    setSelectedSheets(new Set());
    setDrawingCounts({});
    setFileId("");
    setFileName("");
    setDownloadUrl("");
    setDownloadName("");
    setError("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

  return (
    <>
      <AnimatedShaderBackground />

      <main className="app-container">
        {/* Header */}
        <header className="app-header">
          <div className="app-logo">
            <FileSpreadsheet size={24} strokeWidth={1.5} />
          </div>
          <h1 className="app-title">Excel Cleaner</h1>
          <p className="app-subtitle">
            Remove images, shapes, and drawing objects from your worksheets.
          </p>
        </header>

        {/* Step indicator */}
        <div className="steps" role="list" aria-label="Progress">
          {STEPS.map(({ label }, i) => {
            const status = getStepStatus(state, i);
            return (
              <div key={label} style={{ display: "contents" }}>
                <div className="step-item" role="listitem">
                  <div
                    className={`step-circle ${status}`}
                    aria-label={`Step ${i + 1}: ${label}`}
                  >
                    <span className="step-num">{i + 1}</span>
                  </div>
                  <span className="step-label">{label}</span>
                </div>
                {i < STEPS.length - 1 && (
                  <div
                    className={`step-line ${status === "completed" ? "completed" : ""}`}
                    aria-hidden="true"
                  />
                )}
              </div>
            );
          })}
        </div>

        {/* Error */}
        {error && (
          <div className="error-banner" role="alert">
            <AlertTriangle size={15} color="#f87171" strokeWidth={2} style={{ flexShrink: 0 }} />
            <span className="error-text">{error}</span>
            <button
              className="error-dismiss"
              onClick={() => setError("")}
              aria-label="Dismiss"
            >
              <X size={13} />
            </button>
          </div>
        )}

        {/* Card */}
        <div className="card">

          {/* Step 1: Upload */}
          {(state === "idle" || state === "uploading") && (
            <>
              {state === "uploading" ? (
                <div className="processing">
                  <div className="spinner" />
                  <p className="processing-text">Reading workbook</p>
                  <p className="processing-sub">Extracting sheet names…</p>
                </div>
              ) : (
                <div
                  className={`upload-zone${dragging ? " dragging" : ""}`}
                  onDrop={handleDrop}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onClick={() => fileInputRef.current?.click()}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ")
                      fileInputRef.current?.click();
                  }}
                  role="button"
                  tabIndex={0}
                  aria-label="Upload Excel file"
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xlsm"
                    className="upload-input"
                    onChange={handleFileInput}
                    aria-label="File input"
                  />
                  <div className="upload-icon-area">
                    <UploadCloud size={24} strokeWidth={1.5} />
                  </div>
                  <p className="upload-main-text">
                    Drop a file, or <span>browse</span>
                  </p>
                  <p className="upload-sub-text">Up to 50 MB</p>
                  <div className="upload-formats">
                    <span className="format-chip">.xlsx</span>
                    <span className="format-chip">.xlsm</span>
                  </div>
                </div>
              )}
            </>
          )}

          {/* Step 2: Select sheets */}
          {state === "selecting" && (
            <div className="sheet-section">
              <div className="section-header">
                <h2 className="section-title">Select sheets to clean</h2>
                <div className="file-badge" title={fileName}>
                  <span className="file-badge-name">{fileName}</span>
                </div>
              </div>

              <p className="section-meta">
                {sheets.length} sheet{sheets.length !== 1 ? "s" : ""} found
                &nbsp;·&nbsp;
                {selectedSheets.size} selected
              </p>

              <div className="select-controls">
                <button className="select-btn" onClick={selectAll}>
                  All
                </button>
                <button className="select-btn" onClick={selectNone}>
                  None
                </button>
              </div>

              <div className="sheet-list" role="group" aria-label="Worksheets">
                {sheets.map((name, index) => (
                  <div
                    key={name}
                    className={`sheet-item${selectedSheets.has(name) ? " selected" : ""}`}
                    onClick={() => toggleSheet(name)}
                    onKeyDown={(e) => handleSheetKeyDown(e, name)}
                    role="checkbox"
                    aria-checked={selectedSheets.has(name)}
                    tabIndex={0}
                  >
                    <div className="sheet-checkbox">
                      {selectedSheets.has(name) && (
                        <svg
                          width="10"
                          height="10"
                          viewBox="0 0 10 10"
                          fill="none"
                        >
                          <path
                            d="M1.5 5.5L3.5 7.5L8.5 2.5"
                            stroke="currentColor"
                            strokeWidth="1.75"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          />
                        </svg>
                      )}
                    </div>
                    <span className="sheet-name" title={name}>
                      {name}
                    </span>
                    <div className="sheet-meta">
                      {drawingCounts[name] > 0 ? (
                        <span className="drawing-count">
                          <ImageIcon size={11} strokeWidth={2} />
                          {drawingCounts[name]}
                        </span>
                      ) : (
                        <span className="drawing-count drawing-count--clean">clean</span>
                      )}
                      <span className="sheet-index">{index + 1}</span>
                    </div>
                  </div>
                ))}
              </div>

              <button
                className="btn-primary"
                onClick={handleClean}
                disabled={selectedSheets.size === 0}
                id="clean-button"
              >
                <Eraser size={16} strokeWidth={2} />
                Clean {selectedSheets.size} sheet
                {selectedSheets.size !== 1 ? "s" : ""}
              </button>

              <button className="btn-secondary" onClick={handleReset}>
                <ArrowLeft size={14} strokeWidth={2} />
                Upload a different file
              </button>
            </div>
          )}

          {/* Processing */}
          {state === "processing" && (
            <div className="processing">
              <div className="spinner" />
              <p className="processing-text">Cleaning workbook</p>
              <p className="processing-sub">
                Removing drawings from {selectedSheets.size} sheet
                {selectedSheets.size !== 1 ? "s" : ""}…
              </p>
            </div>
          )}

          {/* Done */}
          {state === "done" && (
            <div className="done-section">
              <div className="done-check">
                <CheckCircle2 size={28} strokeWidth={1.5} />
              </div>

              <h2 className="done-title">Done</h2>
              <p className="done-desc">
                Images and shapes removed from the selected sheets.
              </p>

              <div className="done-stats">
                {selectedSheets.size} sheet{selectedSheets.size !== 1 ? "s" : ""} cleaned
              </div>

              <a
                className="download-btn"
                href={downloadUrl}
                download={downloadName}
                id="download-button"
              >
                <Download size={16} strokeWidth={2} />
                Download {downloadName}
              </a>

              <button className="btn-secondary" onClick={handleReset}>
                <RefreshCw size={13} strokeWidth={2} />
                Clean another file
              </button>
            </div>
          )}
        </div>

        {/* Footer */}
        <footer className="app-footer">
          <Lock size={11} strokeWidth={2} />
          Files are deleted from the server immediately after cleaning.
        </footer>
      </main>
    </>
  );
}
