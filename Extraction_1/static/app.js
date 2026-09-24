(() => {
  "use strict";

  const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
  const PDF_EXTENSION = /\.pdf$/i;
  const THEME_KEY = "puc-theme";
  const EMPTY_STATS = {
    filesTotal: 0,
    filesDone: 0,
    pagesScanned: 0,
    urlsFound: 0,
    headerUrls: 0,
    footerUrls: 0,
    qrUrls: 0,
    imageUrls: 0,
    workingUrls: 0,
    failedUrls: 0,
    urlsChecked: 0,
    phase: "extract",
  };
  const STATUS = {
    idle: { label: "Ready to capture" },
    ready: { label: "Files queued" },
    uploading: { label: "Uploading", live: true },
    processing: { label: "In progress", live: true },
    completed: { label: "Complete" },
    error: { label: "Needs attention", error: true },
  };
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  const state = {
    phase: "idle",
    theme: document.documentElement.dataset.theme === "dark" ? "dark" : "light",
    files: [],
    jobId: null,
    snapshot: null,
    errorMessage: "Something went wrong while reading the PDF files.",
    busy: false,
    downloading: false,
    downloadStarted: false,
    pollTimer: null,
    numberTimers: new Map(),
    dragDepth: 0,
    dragging: false,
    previewFilter: "all",
  };

  const els = {
    app: document.getElementById("app"),
    brandLogo: document.getElementById("brand-logo"),
    themeToggle: document.getElementById("theme-toggle"),
    iconMoon: document.getElementById("icon-moon"),
    iconSun: document.getElementById("icon-sun"),
    statusPill: document.getElementById("status-pill"),
    statusLabel: document.getElementById("status-label"),
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    uploadBtn: document.getElementById("upload-btn"),
    fileList: document.getElementById("file-list"),
    addMoreBtn: document.getElementById("add-more-btn"),
    clearFilesBtn: document.getElementById("clear-files-btn"),
    processBtn: document.getElementById("process-btn"),
    processMessage: document.getElementById("process-message"),
    progressPercent: document.getElementById("progress-percent"),
    progressTrack: document.getElementById("progress-track"),
    progressFill: document.getElementById("progress-fill"),
    urlRail: document.getElementById("url-rail"),
    downloadBtn: document.getElementById("download-btn"),
    anotherFileBtn: document.getElementById("another-file-btn"),
    downloadStatus: document.getElementById("download-status"),
    resultTime: document.getElementById("result-time"),
    previewBody: document.getElementById("preview-body"),
    previewNote: document.getElementById("preview-note"),
    filterRow: document.getElementById("filter-row"),
    phaseTrack: document.getElementById("phase-track"),
    errorMessage: document.getElementById("error-message"),
    retryBtn: document.getElementById("retry-btn"),
    chooseAnotherBtn: document.getElementById("choose-another-btn"),
    screens: {
      idle: document.getElementById("screen-idle"),
      ready: document.getElementById("screen-ready"),
      processing: document.getElementById("screen-processing"),
      completed: document.getElementById("screen-completed"),
      error: document.getElementById("screen-error"),
    },
  };

  function apiUrl(path) {
    const prefix = window.location.pathname.replace(/\/index\.html$/, "").replace(/\/$/, "");
    return prefix ? `${prefix}${path}` : path;
  }

  const endpoints = {
    upload: () => apiUrl("/upload"),
    process: () => apiUrl("/process"),
    status: (jobId) => apiUrl(`/status/${encodeURIComponent(jobId)}`),
    download: (jobId) => apiUrl(`/download/${encodeURIComponent(jobId)}`),
  };

  function isPdfFile(file) {
    return PDF_EXTENSION.test(file.name) || file.type === "application/pdf";
  }

  function formatBytes(bytes) {
    if (bytes <= 0) return "0 B";
    const units = ["B", "KB", "MB", "GB"];
    const index = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
    const value = bytes / 1024 ** index;
    return `${value >= 10 || index === 0 ? Math.round(value) : value.toFixed(1)} ${units[index]}`;
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function toUserMessage(error, fallback) {
    const raw = error instanceof Error ? error.message : fallback;
    if (/failed to fetch|networkerror/i.test(raw)) {
      return "Cannot reach the server. Start it with python launch.py and keep it running.";
    }
    return raw || fallback;
  }

  async function parseError(response) {
    try {
      const data = await response.json();
      const detail = typeof data.detail === "string" ? data.detail : undefined;
      return new Error(data.error || data.message || detail || `Request failed (${response.status})`);
    } catch {
      return new Error(`Request failed (${response.status})`);
    }
  }

  async function uploadFiles(files) {
    const body = new FormData();
    files.forEach((file) => body.append("files", file));
    const response = await fetch(endpoints.upload(), { method: "POST", body });
    if (!response.ok) throw await parseError(response);
    return response.json();
  }

  async function processJob(jobId) {
    const response = await fetch(endpoints.process(), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: jobId }),
    });
    if (!response.ok) throw await parseError(response);
    return response.json();
  }

  async function getProcessingStatus(jobId) {
    const response = await fetch(endpoints.status(jobId));
    if (!response.ok) throw await parseError(response);
    return response.json();
  }

  async function downloadProcessedFile(jobId, fallbackName) {
    const response = await fetch(endpoints.download(jobId));
    if (!response.ok) throw await parseError(response);
    const header = response.headers.get("Content-Disposition");
    const match = header && header.match(/filename\*?=(?:UTF-8''|"?)([^";]+)/i);
    const filename =
      match && match[1]
        ? decodeURIComponent(match[1].replace(/"/g, ""))
        : fallbackName || "extracted_urls.xlsx";
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function animateNumber(el, value) {
    const next = Number(value) || 0;
    const previous = Number(el.dataset.current || 0);
    el.dataset.current = String(next);
    if (reducedMotion || previous === next) {
      el.textContent = next.toLocaleString();
      return;
    }
    const existing = state.numberTimers.get(el);
    if (existing) window.cancelAnimationFrame(existing);
    const started = performance.now();
    const duration = 450;
    const tick = (now) => {
      const t = Math.min(1, (now - started) / duration);
      const eased = 1 - (1 - t) * (1 - t);
      el.textContent = Math.round(previous + (next - previous) * eased).toLocaleString();
      if (t < 1) {
        state.numberTimers.set(el, window.requestAnimationFrame(tick));
      } else {
        state.numberTimers.delete(el);
      }
    };
    state.numberTimers.set(el, window.requestAnimationFrame(tick));
  }

  function stopPolling() {
    if (state.pollTimer) {
      window.clearInterval(state.pollTimer);
      state.pollTimer = null;
    }
  }

  function applyTheme() {
    document.documentElement.dataset.theme = state.theme;
    try {
      localStorage.setItem(THEME_KEY, state.theme);
    } catch {
      /* ignore */
    }
    const dark = state.theme === "dark";
    if (els.brandLogo) {
      els.brandLogo.src = dark ? "logo-dark.svg" : "logo.svg";
    }
    if (els.iconMoon) els.iconMoon.classList.toggle("is-hidden", dark);
    if (els.iconSun) els.iconSun.classList.toggle("is-hidden", !dark);
    if (els.themeToggle) {
      const label = dark ? "Switch to light mode" : "Switch to dark mode";
      els.themeToggle.setAttribute("aria-label", label);
      els.themeToggle.title = label;
    }
  }

  function setStatus(kind) {
    const meta = STATUS[kind] || STATUS.idle;
    if (els.statusLabel) els.statusLabel.textContent = meta.label;
    if (els.statusPill) {
      els.statusPill.classList.toggle("is-live", Boolean(meta.live));
      els.statusPill.classList.toggle("is-error", Boolean(meta.error));
    }
  }

  function showScreen(name) {
    Object.entries(els.screens).forEach(([key, node]) => {
      if (!node) return;
      const active = key === name;
      node.classList.toggle("is-active", active);
      node.hidden = !active;
    });
  }

  function renderFileList() {
    if (!els.fileList) return;
    els.fileList.innerHTML = state.files
      .map(
        (file, index) => `<li>
          <span>${escapeHtml(file.name)}</span>
          <em>${formatBytes(file.size)}</em>
          <button type="button" class="btn btn--ghost" data-remove="${index}">Remove</button>
        </li>`,
      )
      .join("");
  }

  function setPhase(phase) {
    state.phase = phase;
    if (phase === "idle") {
      showScreen("idle");
      setStatus("idle");
    } else if (phase === "ready") {
      showScreen("ready");
      setStatus("ready");
      renderFileList();
    } else if (phase === "processing") {
      showScreen("processing");
      setStatus("processing");
      renderProcessing();
    } else if (phase === "completed") {
      showScreen("completed");
      setStatus("completed");
      renderCompleted();
    } else if (phase === "error") {
      showScreen("error");
      setStatus("error");
      if (els.errorMessage) els.errorMessage.textContent = state.errorMessage;
    }
  }

  function renderProcessing() {
    const snap = state.snapshot || {};
    const progress = Math.max(0, Math.min(100, snap.progress || 0));
    const stats = snap.stats || EMPTY_STATS;
    if (els.processMessage) {
      els.processMessage.textContent = snap.message || "Opening files...";
    }
    if (els.progressPercent) els.progressPercent.textContent = `${progress}%`;
    if (els.progressFill) els.progressFill.style.width = `${progress}%`;
    if (els.progressTrack) {
      els.progressTrack.setAttribute("aria-valuenow", String(progress));
      els.progressTrack.setAttribute("aria-label", `Processing ${progress} percent`);
    }
    const phase = stats.phase || (progress >= 72 ? "verify" : "extract");
    if (els.phaseTrack) {
      els.phaseTrack.querySelectorAll("[data-phase]").forEach((node) => {
        const name = node.dataset.phase;
        node.classList.toggle("is-active", name === phase);
        node.classList.toggle(
          "is-done",
          (phase === "verify" && name === "extract") || (phase === "done" && name !== "done"),
        );
      });
    }
    document.querySelectorAll("[data-stat]").forEach((el) => {
      animateNumber(el, stats[el.dataset.stat] || 0);
    });
    document.querySelectorAll("[data-stat-total]").forEach((el) => {
      el.textContent = `of ${stats[el.dataset.statTotal] || 0}`;
    });
    document.querySelectorAll("[data-stat-pair]").forEach((el) => {
      const images = stats[el.dataset.statPair] || 0;
      el.textContent = images ? `+ ${images} image` : "";
    });
    if (els.urlRail) {
      els.urlRail.innerHTML = (snap.recentUrls || [])
        .map((item) => {
          const status = item.status
            ? `<span class="badge ${item.status === "Working" ? "badge--ok" : "badge--bad"}">${escapeHtml(item.status)}</span>`
            : `<span class="badge badge--muted">${escapeHtml(item.source || item.location || "body")}</span>`;
          return `<div class="url-chip">
            <b>p.${escapeHtml(item.page)}</b>
            <span>${escapeHtml(item.url)}</span>
            ${status}
          </div>`;
        })
        .join("");
    }
  }

  function matchesFilter(row, filter) {
    if (filter === "working") return row.status === "Working";
    if (filter === "failed") return String(row.status || "").toLowerCase().includes("not working");
    if (filter === "qr") return row.source === "qr";
    if (filter === "image") return row.source === "image" || row.source === "logo";
    if (filter === "header") return row.location === "header";
    if (filter === "footer") return row.location === "footer";
    return true;
  }

  function renderCompleted() {
    const results = (state.snapshot && state.snapshot.results) || {};
    document.querySelectorAll("[data-result]").forEach((el) => {
      animateNumber(el, results[el.dataset.result] || 0);
    });
    if (els.resultTime) {
      els.resultTime.textContent = `${Number(results.processingTimeSeconds || 0).toFixed(1)}s`;
    }
    if (els.filterRow) {
      els.filterRow.querySelectorAll("[data-filter]").forEach((chip) => {
        chip.classList.toggle("is-active", chip.dataset.filter === state.previewFilter);
      });
    }
    const preview = results.preview || [];
    const visible = preview.filter((row) => matchesFilter(row, state.previewFilter));
    if (els.previewBody) {
      els.previewBody.innerHTML = visible
        .map((row, index) => {
          const working = row.status === "Working";
          const badge = row.status
            ? `<span class="badge ${working ? "badge--ok" : "badge--bad"}">${escapeHtml(row.status)}</span>`
            : "";
          return `<tr style="animation-delay: ${index * 30}ms">
            <td>${escapeHtml(row.fileName)}</td>
            <td>${escapeHtml(row.fileDate || "")}</td>
            <td>${escapeHtml(row.last4 || "")}</td>
            <td>${escapeHtml(row.clientId || "")}</td>
            <td>${escapeHtml(row.formType || "")}</td>
            <td>${escapeHtml(row.pageNumber)}</td>
            <td><span class="badge badge--muted">${escapeHtml(row.source || "text")}</span></td>
            <td class="url-cell">${escapeHtml(row.url)}</td>
            <td>${badge}</td>
          </tr>`;
        })
        .join("");
    }
    const extra = Math.max(0, (results.totalUrls || 0) - preview.length);
    const filterNote =
      state.previewFilter === "all" ? "" : ` Showing ${visible.length} matching row${visible.length === 1 ? "" : "s"}.`;
    if (els.previewNote) {
      els.previewNote.textContent = extra
        ? `${extra} more URL${extra === 1 ? "" : "s"} are in the Excel download.${filterNote}`
        : preview.length
          ? `Captured URLs are shown above and in the Excel file.${filterNote}`
          : "No URLs were found in the uploaded PDFs.";
    }
    if (els.downloadBtn) els.downloadBtn.disabled = state.downloading;
    if (els.downloadStatus) {
      els.downloadStatus.textContent = state.downloadStarted ? "Download started." : "\u00a0";
    }
  }

  function addFiles(fileList) {
    const incoming = Array.from(fileList || []);
    const errors = [];
    incoming.forEach((file) => {
      if (!isPdfFile(file)) {
        errors.push(`${file.name} is not a PDF.`);
        return;
      }
      if (file.size > MAX_UPLOAD_BYTES) {
        errors.push(`${file.name} is larger than 50 MB.`);
        return;
      }
      const exists = state.files.some((item) => item.name === file.name && item.size === file.size);
      if (!exists) state.files.push(file);
    });
    if (state.files.length) {
      setPhase("ready");
    }
    if (errors.length) {
      state.errorMessage = errors.join(" ");
      if (!state.files.length) setPhase("error");
    }
  }

  function reset() {
    stopPolling();
    state.files = [];
    state.jobId = null;
    state.snapshot = null;
    state.busy = false;
    state.downloading = false;
    state.downloadStarted = false;
    state.dragging = false;
    state.dragDepth = 0;
    state.previewFilter = "all";
    if (els.fileInput) els.fileInput.value = "";
    setPhase("idle");
  }

  async function start() {
    if (!state.files.length || state.busy) return;
    state.busy = true;
    state.snapshot = { progress: 2, message: "Opening files...", stats: EMPTY_STATS, recentUrls: [] };
    setPhase("processing");
    try {
      const uploaded = await uploadFiles(state.files);
      state.jobId = uploaded.job_id;
      await processJob(state.jobId);
      stopPolling();
      state.pollTimer = window.setInterval(() => {
        void pollStatus();
      }, 400);
      await pollStatus();
    } catch (error) {
      stopPolling();
      state.busy = false;
      state.errorMessage = toUserMessage(error, "Unable to process files.");
      setPhase("error");
    }
  }

  async function pollStatus() {
    if (!state.jobId) return;
    try {
      const snap = await getProcessingStatus(state.jobId);
      state.snapshot = snap;
      if (snap.status === "completed") {
        stopPolling();
        state.busy = false;
        setPhase("completed");
        return;
      }
      if (snap.status === "failed") {
        stopPolling();
        state.busy = false;
        state.errorMessage = snap.error || snap.message || "Unable to process files.";
        setPhase("error");
        return;
      }
      renderProcessing();
    } catch (error) {
      stopPolling();
      state.busy = false;
      state.errorMessage = toUserMessage(error, "Lost connection while processing.");
      setPhase("error");
    }
  }

  function bindEvents() {
    els.themeToggle?.addEventListener("click", () => {
      state.theme = state.theme === "dark" ? "light" : "dark";
      applyTheme();
    });
    els.uploadBtn?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      els.fileInput?.click();
    });
    els.addMoreBtn?.addEventListener("click", () => els.fileInput?.click());
    els.clearFilesBtn?.addEventListener("click", reset);
    els.fileInput?.addEventListener("change", () => {
      addFiles(els.fileInput.files);
      els.fileInput.value = "";
    });
    els.fileList?.addEventListener("click", (event) => {
      const button = event.target.closest("[data-remove]");
      if (!button) return;
      const index = Number(button.dataset.remove);
      state.files.splice(index, 1);
      if (state.files.length) renderFileList();
      else reset();
    });
    els.processBtn?.addEventListener("click", () => void start());
    els.downloadBtn?.addEventListener("click", async () => {
      if (!state.jobId || state.downloading) return;
      state.downloading = true;
      renderCompleted();
      try {
        const name = state.snapshot?.results?.downloadFilename;
        await downloadProcessedFile(state.jobId, name);
        state.downloadStarted = true;
      } catch (error) {
        state.errorMessage = toUserMessage(error, "The Excel file could not be downloaded.");
        setPhase("error");
      } finally {
        state.downloading = false;
        if (state.phase === "completed") renderCompleted();
      }
    });
    els.anotherFileBtn?.addEventListener("click", reset);
    els.retryBtn?.addEventListener("click", () => void start());
    els.chooseAnotherBtn?.addEventListener("click", reset);
    els.filterRow?.addEventListener("click", (event) => {
      const chip = event.target.closest("[data-filter]");
      if (!chip) return;
      state.previewFilter = chip.dataset.filter || "all";
      if (state.phase === "completed") renderCompleted();
    });

    const zone = els.dropzone;
    if (zone) {
      zone.addEventListener("click", (event) => {
        if (event.target.closest("button")) return;
        els.fileInput?.click();
      });
      zone.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          els.fileInput?.click();
        }
      });
    }

    const onDragEnter = (event) => {
      event.preventDefault();
      state.dragDepth += 1;
      state.dragging = true;
      els.dropzone?.classList.add("is-dragging");
    };
    const onDragOver = (event) => {
      event.preventDefault();
    };
    const onDragLeave = (event) => {
      event.preventDefault();
      state.dragDepth = Math.max(0, state.dragDepth - 1);
      if (!state.dragDepth) {
        state.dragging = false;
        els.dropzone?.classList.remove("is-dragging");
      }
    };
    const onDrop = (event) => {
      event.preventDefault();
      state.dragDepth = 0;
      state.dragging = false;
      els.dropzone?.classList.remove("is-dragging");
      addFiles(event.dataTransfer?.files);
    };
    document.addEventListener("dragenter", onDragEnter);
    document.addEventListener("dragover", onDragOver);
    document.addEventListener("dragleave", onDragLeave);
    document.addEventListener("drop", onDrop);
  }

  applyTheme();
  bindEvents();
  setPhase("idle");
})();
