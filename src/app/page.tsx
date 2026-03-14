"use client";

import { useState, useRef, useCallback } from "react";

type AppState = "idle" | "uploading" | "selecting" | "processing" | "done";

interface UploadResult {
  sheets: string[];
  fileId: string;
  fileName: string;
}

export default function Home() {
  const [state, setState] = useState<AppState>("idle");
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [fileId, setFileId] = useState<string>("");
  const [fileName, setFileName] = useState<string>("");
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [downloadName, setDownloadName] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    const ext = file.name.toLowerCase();
    if (!ext.endsWith(".xlsx") && !ext.endsWith(".xlsm")) {
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

      const res = await fetch("/api/sheets", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || "Failed to process file");
      }

      const result = data as UploadResult;
      setSheets(result.sheets);
      setFileId(result.fileId);
      setFileName(result.fileName);
      setSelectedSheets(new Set(result.sheets)); // select all by default
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
      if (next.has(name)) {
        next.delete(name);
      } else {
        next.add(name);
      }
      return next;
    });
  }, []);

  const selectAll = useCallback(() => {
    setSelectedSheets(new Set(sheets));
  }, [sheets]);

  const selectNone = useCallback(() => {
    setSelectedSheets(new Set());
  }, []);

  const handleClean = useCallback(async () => {
    if (selectedSheets.size === 0) {
      setError("Please select at least one sheet to clean.");
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

      const cleanedName = fileName
        ? fileName.replace(/\.xlsx$/i, "_cleaned.xlsx")
        : "cleaned.xlsx";

      setDownloadUrl(url);
      setDownloadName(cleanedName);
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
    setFileId("");
    setFileName("");
    setDownloadUrl("");
    setDownloadName("");
    setError("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }, []);

  const getStepState = (step: number) => {
    const steps: Record<AppState, number> = {
      idle: 0,
      uploading: 0,
      selecting: 1,
      processing: 1,
      done: 2,
    };
    const current = steps[state];
    if (step < current) return "completed";
    if (step === current) return "active";
    return "";
  };

  return (
    <>
      {/* Background orbs */}
      <div className="bg-decoration">
        <div className="bg-orb bg-orb-1" />
        <div className="bg-orb bg-orb-2" />
        <div className="bg-orb bg-orb-3" />
      </div>

      <main className="app-container">
        {/* Header */}
        <header className="app-header">
          <div className="app-logo">📄</div>
          <h1 className="app-title">Excel Cleaner</h1>
          <p className="app-subtitle">
            Remove images, shapes, and drawing objects from your Excel
            worksheets
          </p>
        </header>

        {/* Step Indicator */}
        <div className="steps">
          <div className={`step-dot ${getStepState(0)}`} />
          <div
            className={`step-connector ${getStepState(0) === "completed" ? "completed" : ""}`}
          />
          <div className={`step-dot ${getStepState(1)}`} />
          <div
            className={`step-connector ${getStepState(1) === "completed" ? "completed" : ""}`}
          />
          <div className={`step-dot ${getStepState(2)}`} />
        </div>

        {/* Error Banner */}
        {error && (
          <div className="error-banner">
            <span>⚠️</span>
            <span className="error-text">{error}</span>
            <button
              className="error-dismiss"
              onClick={() => setError("")}
              aria-label="Dismiss error"
            >
              ✕
            </button>
          </div>
        )}

        {/* Main Card */}
        <div className="card">
          {/* Step 1: Upload */}
          {(state === "idle" || state === "uploading") && (
            <div
              className={`upload-zone ${dragging ? "dragging" : ""}`}
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onClick={() => fileInputRef.current?.click()}
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
              {state === "uploading" ? (
                <div className="processing">
                  <div className="spinner" />
                  <p className="processing-text">Reading workbook…</p>
                </div>
              ) : (
                <>
                  <span className="upload-icon">📤</span>
                  <p className="upload-text">
                    Drop your .xlsx or .xlsm file here or click to browse
                  </p>
                  <p className="upload-hint">Supports files up to 50 MB</p>
                </>
              )}
            </div>
          )}

          {/* Step 2: Select Sheets */}
          {state === "selecting" && (
            <div className="sheet-section">
              <div className="section-header">
                <h2 className="section-title">Select Sheets to Clean</h2>
                <div className="file-badge">
                  <span>📎</span>
                  <span>{fileName}</span>
                </div>
              </div>

              <div className="select-controls">
                <button className="select-btn" onClick={selectAll}>
                  Select All
                </button>
                <button className="select-btn" onClick={selectNone}>
                  Select None
                </button>
              </div>

              <div className="sheet-list">
                {sheets.map((name, index) => (
                  <div
                    key={name}
                    className={`sheet-item ${selectedSheets.has(name) ? "selected" : ""}`}
                    onClick={() => toggleSheet(name)}
                    role="checkbox"
                    aria-checked={selectedSheets.has(name)}
                    tabIndex={0}
                  >
                    <div className="sheet-checkbox">
                      {selectedSheets.has(name) && "✓"}
                    </div>
                    <span className="sheet-name">{name}</span>
                    <span className="sheet-index">Sheet {index + 1}</span>
                  </div>
                ))}
              </div>

              <button
                className="btn-primary"
                onClick={handleClean}
                disabled={selectedSheets.size === 0}
                id="clean-button"
              >
                🧹 Clean {selectedSheets.size} Sheet
                {selectedSheets.size !== 1 ? "s" : ""}
              </button>

              <button className="btn-secondary" onClick={handleReset}>
                ← Upload a different file
              </button>
            </div>
          )}

          {/* Step 2.5: Processing */}
          {state === "processing" && (
            <div className="processing">
              <div className="spinner" />
              <p className="processing-text">Cleaning your workbook…</p>
              <p className="processing-sub">
                Removing images and shapes from {selectedSheets.size} sheet
                {selectedSheets.size !== 1 ? "s" : ""}
              </p>
            </div>
          )}

          {/* Step 3: Done */}
          {state === "done" && (
            <div className="done-section">
              <span className="done-icon">✅</span>
              <h2 className="done-title">Workbook Cleaned!</h2>
              <p className="done-desc">
                All images, shapes, and drawing objects have been removed from
                the selected sheets.
              </p>

              <a
                className="download-btn"
                href={downloadUrl}
                download={downloadName}
                id="download-button"
              >
                ⬇️ Download {downloadName}
              </a>

              <button className="btn-secondary" onClick={handleReset}>
                🔄 Clean another file
              </button>
            </div>
          )}
        </div>

        {/* Footer */}
        <footer className="app-footer">
          Files are processed on the server and automatically deleted after
          cleaning.
        </footer>
      </main>
    </>
  );
}
