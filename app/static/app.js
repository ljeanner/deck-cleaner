/**
 * app.js
 * ------
 * Vanilla JavaScript frontend logic for Deck Cleaner.
 *
 * Handles:
 *   - Drag-and-drop / file-picker for .pptx selection
 *   - File validation (extension, zero-size check)
 *   - POST /optimize → display result summary
 *   - GET  /download/{filename} → trigger browser download
 *   - Section visibility management (upload / loading / result / error)
 */

(function () {
  "use strict";

  // ---------------------------------------------------------------------------
  // DOM references
  // ---------------------------------------------------------------------------

  const dropZone     = document.getElementById("drop-zone");
  const fileInput    = document.getElementById("file-input");
  const fileInfo     = document.getElementById("file-info");
  const fileNameEl   = document.getElementById("file-name");
  const fileSizeEl   = document.getElementById("file-size");
  const btnClear     = document.getElementById("btn-clear");
  const btnOptimize  = document.getElementById("btn-optimize");

  const sectionUpload  = document.getElementById("section-upload");
  const sectionLoading = document.getElementById("section-loading");
  const sectionResult  = document.getElementById("section-result");
  const sectionError   = document.getElementById("section-error");

  const statOriginal  = document.getElementById("stat-original");
  const statOptimized = document.getElementById("stat-optimized");
  const statLayouts   = document.getElementById("stat-layouts");
  const statMasters   = document.getElementById("stat-masters");
  const statSavings   = document.getElementById("stat-savings");

  const btnDownload = document.getElementById("btn-download");
  const btnReset    = document.getElementById("btn-reset");
  const btnRetry    = document.getElementById("btn-retry");
  const errorMsg    = document.getElementById("error-message");

  // ---------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------

  /** @type {File|null} */
  let selectedFile = null;

  /** @type {string|null} – filename returned by the backend */
  let outputFilename = null;

  // ---------------------------------------------------------------------------
  // Utility helpers
  // ---------------------------------------------------------------------------

  /**
   * Format a byte count into a human-readable string.
   * @param {number} bytes
   * @returns {string}
   */
  function formatBytes(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
  }

  /**
   * Show only the specified section; hide all others.
   * @param {"upload"|"loading"|"result"|"error"} name
   */
  function showSection(name) {
    sectionUpload.classList.toggle("hidden",  name !== "upload");
    sectionLoading.classList.toggle("hidden", name !== "loading");
    sectionResult.classList.toggle("hidden",  name !== "result");
    sectionError.classList.toggle("hidden",   name !== "error");
  }

  /**
   * Apply a selected file to the UI state.
   * @param {File} file
   */
  function applyFile(file) {
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      showError("Only .pptx files are supported. Please select a valid PowerPoint file.");
      return;
    }
    if (file.size === 0) {
      showError("The selected file is empty.");
      return;
    }

    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = "(" + formatBytes(file.size) + ")";
    fileInfo.classList.remove("hidden");
    btnOptimize.disabled = false;
  }

  /**
   * Reset file selection back to the initial state.
   */
  function clearFile() {
    selectedFile = null;
    outputFilename = null;
    fileInput.value = "";
    fileInfo.classList.add("hidden");
    btnOptimize.disabled = true;
  }

  /**
   * Display the error section with the given message.
   * @param {string} message
   */
  function showError(message) {
    errorMsg.textContent = message;
    showSection("error");
  }

  // ---------------------------------------------------------------------------
  // File selection – drop zone
  // ---------------------------------------------------------------------------

  dropZone.addEventListener("click", () => fileInput.click());

  dropZone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      applyFile(files[0]);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files && fileInput.files.length > 0) {
      applyFile(fileInput.files[0]);
    }
  });

  // ---------------------------------------------------------------------------
  // Clear button
  // ---------------------------------------------------------------------------

  btnClear.addEventListener("click", () => {
    clearFile();
    showSection("upload");
  });

  // ---------------------------------------------------------------------------
  // Optimize button → POST /optimize
  // ---------------------------------------------------------------------------

  btnOptimize.addEventListener("click", async () => {
    if (!selectedFile) return;

    showSection("loading");

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/optimize", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let detail = "An error occurred during optimization.";
        try {
          const json = await response.json();
          detail = json.detail || detail;
        } catch (_) {
          // ignore JSON parse error
        }
        showError(detail);
        return;
      }

      const data = await response.json();
      outputFilename = data.output_filename;

      // Populate stats
      statOriginal.textContent  = formatBytes(data.original_size);
      statOptimized.textContent = formatBytes(data.optimized_size);
      statLayouts.textContent   = data.removed_layouts;
      statMasters.textContent   = data.removed_masters;

      const saved = data.original_size - data.optimized_size;
      const pct   = data.original_size > 0
        ? ((saved / data.original_size) * 100).toFixed(1)
        : 0;
      statSavings.textContent = formatBytes(Math.max(0, saved)) + " (" + pct + "%)";

      showSection("result");
    } catch (err) {
      showError("Network error: could not reach the server. Is it running?");
      console.error(err);
    }
  });

  // ---------------------------------------------------------------------------
  // Download button → GET /download/{filename}
  // ---------------------------------------------------------------------------

  btnDownload.addEventListener("click", () => {
    if (!outputFilename) return;
    // Trigger a browser download by navigating to the download endpoint.
    const link = document.createElement("a");
    link.href = "/download/" + encodeURIComponent(outputFilename);
    link.download = outputFilename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });

  // ---------------------------------------------------------------------------
  // Reset / retry buttons
  // ---------------------------------------------------------------------------

  btnReset.addEventListener("click", () => {
    clearFile();
    showSection("upload");
  });

  btnRetry.addEventListener("click", () => {
    showSection("upload");
  });

})();
