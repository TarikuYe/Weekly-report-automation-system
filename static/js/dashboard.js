/* =========================================================
   Weekly Report Dashboard — JavaScript
   Handles: navigation, API calls, Chart.js charts,
            SSE log streaming, file upload, settings.
   ========================================================= */

"use strict";

// ── Chart.js global defaults ──────────────────────────────
Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
Chart.defaults.font.size   = 12;
Chart.defaults.color       = "#6B7280";
Chart.defaults.plugins.legend.labels.boxWidth = 12;
Chart.defaults.plugins.legend.labels.padding  = 16;

const ACCENT      = "#2D5BFF";
const ACCENT_DARK = "#1B2A5C";
const SUCCESS     = "#1B8A5A";
const WARNING     = "#B7791F";
const DANGER      = "#C0392B";

const CHART_PALETTE = [
  "#2D5BFF","#10B981","#F59E0B","#8B5CF6","#EC4899",
  "#06B6D4","#EF4444","#84CC16","#F97316","#6366F1"
];

// ── State ─────────────────────────────────────────────────
const state = {
  pipelineRunning: false,
  lastData: null,
  charts: {},
};

// ── Utility helpers ───────────────────────────────────────
const $  = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

async function api(url, opts = {}) {
  const res = await fetch(url, { headers: { "Content-Type": "application/json" }, ...opts });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
  return res.json();
}

function showToast(msg, type = "info") {
  const container = $("#toastContainer");
  const icons = { success: "check-circle", error: "alert-circle", info: "info", warning: "alert-triangle" };
  const div = document.createElement("div");
  div.className = `toast toast-${type}`;
  div.innerHTML = `<i data-feather="${icons[type] || "info"}"></i><span>${msg}</span>`;
  container.appendChild(div);
  feather.replace({ "aria-hidden": "true" });
  setTimeout(() => {
    div.classList.add("removing");
    setTimeout(() => div.remove(), 300);
  }, 3800);
}

function formatNum(n, decimals = 1) {
  if (n == null || n === "" || isNaN(n)) return "—";
  return Number(n).toLocaleString("en", { maximumFractionDigits: decimals });
}

function statusBadge(val) {
  if (!val || val === "") return "";
  const v = String(val).toLowerCase();
  let cls = "badge-muted";
  if (v.includes("complet")) cls = "badge-success";
  else if (v.includes("progress") || v.includes("ongoing")) cls = "badge-warning";
  else if (v.includes("block") || v.includes("delay")) cls = "badge-danger";
  return `<span class="badge ${cls}">${val}</span>`;
}

// ── Navigation ────────────────────────────────────────────
function initNav() {
  const sectionTitles = {
    overview:  "Overview",
    analytics: "Analytics",
    pipeline:  "Pipeline",
    files:     "Files",
    settings:  "Settings",
  };

  $$(".nav-item").forEach(link => {
    link.addEventListener("click", e => {
      e.preventDefault();
      const target = link.dataset.section;
      activateSection(target);
      $$(".nav-item").forEach(l => l.classList.remove("active"));
      link.classList.add("active");
      $("#topbarTitle").textContent = sectionTitles[target] || target;
    });
  });
}

function activateSection(name) {
  $$(".page-section").forEach(s => s.classList.remove("active"));
  const section = $(`#section-${name}`);
  if (section) section.classList.add("active");
}

// ── Sidebar collapse ──────────────────────────────────────
function initSidebar() {
  $("#sidebarToggle").addEventListener("click", () => {
    document.body.classList.toggle("sidebar-collapsed");
    $("#sidebar").classList.toggle("collapsed");
  });
}

// ── Dashboard data load ───────────────────────────────────
async function loadDashboard() {
  try {
    const data = await api("/api/data/summary");
    state.lastData = data;

    if (!data.available) {
      const msg = data.error ? `Error: ${data.error}` : "No report found. Run the pipeline first.";
      showNoDataState(msg);
      return;
    }

    updateKPIs(data.kpis || {});
    renderDeptChart(data.dept_hours);
    renderStatusChart(data.status_dist);
    renderProjectChart(data.project_hours);
    renderTrendChart(data.hours_trend);
    renderRecentTable(data.recent_rows || []);
    renderDeptTable(data.dept_table || []);

    // Last updated timestamp
    const now = new Date().toLocaleTimeString("en", { hour: "2-digit", minute: "2-digit" });
    const today = new Date().toLocaleDateString("en", { weekday: "short", day: "numeric", month: "short" });
    $("#lastUpdated").textContent = `${today} · ${now}`;

  } catch (err) {
    console.error("loadDashboard:", err);
    showToast("Failed to load dashboard data", "error");
  }
}

function showNoDataState(msg) {
  $$(".kpi-value").forEach(el => el.textContent = "—");
  ["recentTableBody", "deptTableBody", "filesTableBody"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<tr><td colspan="10" class="empty-state">${msg}</td></tr>`;
  });
}

// ── KPIs ──────────────────────────────────────────────────
function updateKPIs(kpis) {
  const setKPI = (id, val, suffix = "") => {
    const el = document.getElementById(id);
    if (el) el.textContent = (val != null && val !== "") ? `${formatNum(val, 0)}${suffix}` : "—";
  };
  setKPI("kpiRows",       kpis.total_rows, "");
  setKPI("kpiHours",      kpis.total_hours, " hrs");
  setKPI("kpiDepts",      kpis.dept_count, "");
  setKPI("kpiStaff",      kpis.staff_count, "");
  setKPI("kpiAvg",        kpis.avg_hours_per_person, " hrs");
  setKPI("kpiCompletion", kpis.completion_pct, "%");
}

// ── Charts ────────────────────────────────────────────────
function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
    delete state.charts[id];
  }
}

function renderDeptChart(data) {
  destroyChart("dept");
  if (!data || !data.labels.length) return;

  const ctx = $("#deptChart").getContext("2d");
  state.charts["dept"] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: data.labels,
      datasets: [{
        label: "Hours Logged",
        data: data.values,
        backgroundColor: CHART_PALETTE.map(c => c + "CC"),
        borderColor: CHART_PALETTE,
        borderWidth: 1.5,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y.toLocaleString()} hrs` } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxRotation: 30 } },
        y: {
          grid: { color: "#E3E6EE" },
          ticks: { callback: v => v.toLocaleString() },
          beginAtZero: true,
        }
      }
    }
  });
}

function renderStatusChart(data) {
  destroyChart("status");
  if (!data || !data.labels.length) return;

  const ctx = $("#statusChart").getContext("2d");
  state.charts["status"] = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: data.labels,
      datasets: [{
        data: data.values,
        backgroundColor: CHART_PALETTE,
        borderWidth: 2,
        borderColor: "#fff",
        hoverOffset: 6,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "65%",
      plugins: {
        legend: {
          position: "bottom",
          labels: { padding: 12, usePointStyle: true, pointStyleWidth: 8 }
        },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString()}` } }
      }
    }
  });
}

function renderProjectChart(data) {
  destroyChart("project");
  if (!data || !data.labels.length) return;

  const ctx = $("#projectChart").getContext("2d");
  state.charts["project"] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: data.labels,
      datasets: [{
        label: "Hours",
        data: data.values,
        backgroundColor: ACCENT + "BB",
        borderColor: ACCENT,
        borderWidth: 1.5,
        borderRadius: 5,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.x.toLocaleString()} hrs` } }
      },
      scales: {
        x: {
          grid: { color: "#E3E6EE" },
          ticks: { callback: v => v.toLocaleString() },
          beginAtZero: true,
        },
        y: { grid: { display: false }, ticks: { font: { size: 11 } } }
      }
    }
  });
}

function renderTrendChart(data) {
  destroyChart("trend");
  if (!data || !data.labels.length) return;

  const ctx = $("#trendChart").getContext("2d");
  state.charts["trend"] = new Chart(ctx, {
    type: "line",
    data: {
      labels: data.labels,
      datasets: [{
        label: "Hours Logged",
        data: data.values,
        borderColor: ACCENT,
        backgroundColor: "rgba(45,91,255,.08)",
        borderWidth: 2.5,
        pointRadius: 4,
        pointBackgroundColor: ACCENT,
        tension: 0.4,
        fill: true,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y.toLocaleString()} hrs` } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxRotation: 30 } },
        y: {
          grid: { color: "#E3E6EE" },
          ticks: { callback: v => v.toLocaleString() },
          beginAtZero: true,
        }
      }
    }
  });
}

// ── Tables ────────────────────────────────────────────────
function renderRecentTable(rows) {
  const head = $("#recentTableHead");
  const body = $("#recentTableBody");
  const badge = $("#recentBadge");

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="10" class="empty-state">No activity data found.</td></tr>`;
    badge.textContent = "0 rows";
    head.innerHTML = "";
    return;
  }

  const cols = Object.keys(rows[0]);
  head.innerHTML = cols.map(c => `<th>${c}</th>`).join("");
  body.innerHTML = rows.map(row => `
    <tr>${cols.map(c => {
      const val = row[c] ?? "";
      if (c === "Status") return `<td>${statusBadge(val)}</td>`;
      return `<td>${val}</td>`;
    }).join("")}</tr>
  `).join("");
  badge.textContent = `${rows.length} rows`;
}

function renderDeptTable(rows) {
  const body = $("#deptTableBody");
  const badge = $("#deptTableBadge");

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="5" class="empty-state">No department data.</td></tr>`;
    badge.textContent = "0";
    return;
  }

  body.innerHTML = rows.map((r, i) => {
    const hours = r["Total Hours"] ?? r["Hours Worked"] ?? "—";
    const staff = r["Staff"] ?? r["Staff Name"] ?? "—";
    const recs  = r["Rows"] ?? "—";
    const avg   = (hours && recs && recs !== "—" && hours !== "—")
      ? (parseFloat(hours) / parseInt(recs)).toFixed(1)
      : "—";
    return `
      <tr>
        <td><strong>${r["Department"] ?? "—"}</strong></td>
        <td>${formatNum(hours)} hrs</td>
        <td>${staff}</td>
        <td>${recs}</td>
        <td>${avg !== "NaN" ? avg : "—"}</td>
      </tr>
    `;
  }).join("");
  badge.textContent = `${rows.length} depts`;
}

// ── Pipeline ──────────────────────────────────────────────
async function runPipeline() {
  if (state.pipelineRunning) return;

  state.pipelineRunning = true;
  setPipelineUI("running");
  showToast("Pipeline started…", "info");

  try {
    await api("/api/pipeline/run", { method: "POST" });
    pollPipelineStatus();
  } catch (err) {
    state.pipelineRunning = false;
    setPipelineUI("error", err.message);
    showToast(`Failed to start pipeline: ${err.message}`, "error");
  }
}

function pollPipelineStatus() {
  const interval = setInterval(async () => {
    try {
      const s = await api("/api/status");
      if (!s.running) {
        clearInterval(interval);
        state.pipelineRunning = false;

        if (s.success) {
          setPipelineUI("success", `Completed in ${s.elapsed}s`);
          showToast("Pipeline completed successfully!", "success");
          await loadDashboard();
          await loadFiles();
        } else {
          setPipelineUI("error", s.error || "Unknown error");
          showToast(`Pipeline failed: ${s.error}`, "error");
        }
      }
    } catch (err) {
      clearInterval(interval);
      state.pipelineRunning = false;
      setPipelineUI("error", err.message);
    }
  }, 1500);
}

function setPipelineUI(status, subText = "") {
  const dot        = $("#bigStatusDot");
  const label      = $("#bigStatusLabel");
  const sub        = $("#bigStatusSub");
  const progressW  = $("#progressWrap");
  const runBtn     = $("#runPipelineBtn");
  const topRunBtn  = $("#topbarRunBtn");
  const miniDot    = $("#miniDot");
  const miniText   = $("#miniStatusText");
  const badge      = $("#pipelineBadge");
  const meta       = $("#pipelineMeta");

  const configs = {
    idle:    { label: "Ready",     sub: "Click Run Pipeline to start", dotClass: "idle",    btnDisabled: false },
    running: { label: "Running…",  sub: "Merging department files",    dotClass: "running", btnDisabled: true  },
    success: { label: "Completed", sub: subText || "All files merged", dotClass: "success", btnDisabled: false },
    error:   { label: "Failed",    sub: subText || "An error occurred",dotClass: "error",   btnDisabled: false },
  };

  const cfg = configs[status] || configs.idle;

  // Remove all state classes, add new one
  dot.className = `big-status-dot ${cfg.dotClass}`;
  miniDot.className = `status-dot ${cfg.dotClass}`;
  label.textContent = cfg.label;
  sub.textContent = cfg.sub;
  miniText.textContent = cfg.label;

  progressW.style.display = status === "running" ? "block" : "none";
  meta.style.display = (status === "success" || status === "error") ? "flex" : "none";

  if (status === "success" || status === "error") {
    $("#metaStatus").textContent   = status === "success" ? "Success" : "Failed";
    $("#metaDuration").textContent = subText || "—";
    $("#metaStatus").style.color   = status === "success" ? "var(--success)" : "var(--danger)";
  }

  [runBtn, topRunBtn].forEach(btn => {
    if (btn) btn.disabled = cfg.btnDisabled;
  });

  badge.style.display = status === "running" ? "inline" : "none";
}

// ── SSE Log streaming ─────────────────────────────────────
function initLogStream() {
  const console_ = $("#logConsole");

  function appendLine(entry) {
    const div = document.createElement("div");
    div.className = "log-line";
    div.innerHTML = `
      <span class="log-ts">${entry.ts}</span>
      <span class="log-level ${entry.level}">${entry.level}</span>
      <span class="log-msg">${escapeHtml(entry.msg)}</span>
    `;
    console_.appendChild(div);
    console_.scrollTop = console_.scrollHeight;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  const evtSource = new EventSource("/api/logs/stream");
  evtSource.onmessage = e => {
    try {
      appendLine(JSON.parse(e.data));
    } catch (_) {}
  };
  evtSource.onerror = () => {
    // Reconnect automatically — browser SSE handles this
  };

  $("#clearLogBtn").addEventListener("click", () => {
    console_.innerHTML = "";
  });
}

// ── Files section ─────────────────────────────────────────
async function loadFiles() {
  try {
    const files = await api("/api/files/input");
    const body  = $("#filesTableBody");
    const badge = $("#fileCountBadge");

    badge.textContent = `${files.length} file${files.length !== 1 ? "s" : ""}`;

    if (!files.length) {
      body.innerHTML = `<tr><td colspan="4" class="empty-state">No .xlsx files in input folder yet.</td></tr>`;
      return;
    }

    body.innerHTML = files.map(f => `
      <tr>
        <td><strong>${f.name}</strong></td>
        <td>${f.size_kb} KB</td>
        <td>${f.modified}</td>
        <td style="text-align:center">
          <button class="btn btn-danger btn-sm" onclick="deleteFile('${escapeAttr(f.name)}')">
            <i data-feather="trash-2"></i> Delete
          </button>
        </td>
      </tr>
    `).join("");
    feather.replace({ "aria-hidden": "true" });
  } catch (err) {
    showToast("Could not load file list", "error");
  }
}

function escapeAttr(s) {
  return String(s).replace(/'/g, "\\'").replace(/"/g, "&quot;");
}

async function deleteFile(name) {
  if (!confirm(`Delete "${name}"?`)) return;
  try {
    await api(`/api/files/delete/${encodeURIComponent(name)}`, { method: "DELETE" });
    showToast(`Deleted ${name}`, "success");
    loadFiles();
  } catch (err) {
    showToast(`Delete failed: ${err.message}`, "error");
  }
}

function initDropZone() {
  const zone  = $("#dropZone");
  const input = $("#fileInput");

  zone.addEventListener("click", () => input.click());
  zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    uploadFiles(e.dataTransfer.files);
  });
  input.addEventListener("change", () => uploadFiles(input.files));
}

async function uploadFiles(fileList) {
  const feedback = $("#uploadFeedback");
  if (!fileList.length) return;

  const form = new FormData();
  let count = 0;
  for (const f of fileList) {
    if (f.name.toLowerCase().endsWith(".xlsx")) {
      form.append("files", f);
      count++;
    }
  }
  if (!count) {
    showToast("Only .xlsx files are accepted", "warning");
    return;
  }

  feedback.style.display = "block";
  feedback.className = "upload-feedback";
  feedback.textContent = `Uploading ${count} file(s)…`;

  try {
    const res = await fetch("/api/files/upload", { method: "POST", body: form });
    const data = await res.json();

    if (data.uploaded.length) {
      feedback.className = "upload-feedback success";
      feedback.textContent = `Uploaded: ${data.uploaded.join(", ")}`;
      showToast(`${data.uploaded.length} file(s) uploaded`, "success");
      loadFiles();
    }
    if (data.errors.length) {
      feedback.className = "upload-feedback error";
      feedback.textContent = `Errors: ${data.errors.join("; ")}`;
      showToast("Some files had upload errors", "error");
    }
    setTimeout(() => { feedback.style.display = "none"; }, 4000);
  } catch (err) {
    feedback.className = "upload-feedback error";
    feedback.textContent = `Upload failed: ${err.message}`;
    showToast("Upload failed", "error");
  }

  // Reset file input so the same file can be re-uploaded
  $("#fileInput").value = "";
}

// ── Settings ──────────────────────────────────────────────
async function initSettings() {
  try {
    const cfg = await api("/api/config");
    if (cfg.input_dir)  $("#inputDirInput").value  = cfg.input_dir;
    if (cfg.output_file) $("#outputFileInput").value = cfg.output_file;
  } catch (_) {}

  $("#settingsForm").addEventListener("submit", async e => {
    e.preventDefault();
    const feedback = $("#saveFeedback");
    const payload = {
      input_dir:   $("#inputDirInput").value.trim(),
      output_file: $("#outputFileInput").value.trim(),
    };
    try {
      await api("/api/config", { method: "POST", body: JSON.stringify(payload) });
      feedback.textContent = "✓ Saved";
      feedback.className = "save-feedback ok";
      feedback.style.display = "inline-block";
      showToast("Settings saved", "success");
      setTimeout(() => { feedback.style.display = "none"; }, 3000);
    } catch (err) {
      feedback.textContent = `Error: ${err.message}`;
      feedback.className = "save-feedback err";
      feedback.style.display = "inline-block";
      showToast("Failed to save settings", "error");
    }
  });
}

// ── Clear Data (Danger Zone) ──────────────────────────────
function initClearData() {
  const modal = $("#clearModal");
  const modalList = $("#clearModalList");
  const confirmBtn = $("#clearModalConfirm");
  const cancelBtn = $("#clearModalCancel");
  
  let currentMode = "output-only"; // "output-only" or "all"

  function showClearModal(mode) {
    currentMode = mode;
    
    // Build the list of what will be deleted
    const items = [
      "Master weekly report (output/master_weekly_report.xlsx)",
      "All department master exports (output/dept_masters/*.xlsx)"
    ];
    
    if (mode === "all") {
      items.push("All archived input batches from the root archive folder");
      items.push("All archived input batches from every department folder");
    }
    
    modalList.innerHTML = items.map(item => `<li>${item}</li>`).join("");
    modal.style.display = "flex";
    feather.replace({ "aria-hidden": "true" });
  }

  function hideClearModal() {
    modal.style.display = "none";
  }

  async function executeClear() {
    hideClearModal();
    const includeArchives = currentMode === "all";
    const url = `/api/data/clear?include_archives=${includeArchives}`;
    
    try {
      showToast(`Deleting ${currentMode === "all" ? "all data" : "output data"}...`, "info");
      const result = await api(url, { method: "DELETE" });
      
      if (result.ok) {
        const msg = `Cleared ${result.deleted_files} file(s) and ${result.deleted_folders} folder(s)`;
        showToast(msg, "success");
        
        // Reset dashboard UI to empty state
        showNoDataState("No data available. Run the pipeline to generate a new report.");
        
        // Destroy all charts
        Object.keys(state.charts).forEach(key => destroyChart(key));
        
        // Clear last data
        state.lastData = null;
        
        // Refresh file list
        await loadFiles();
      } else {
        const msg = result.errors && result.errors.length 
          ? `Partial success: ${result.errors.join("; ")}`
          : "Some files could not be deleted";
        showToast(msg, "warning");
      }
    } catch (err) {
      showToast(`Failed to clear data: ${err.message}`, "error");
    }
  }

  // Wire up all three clear buttons (topbar + two in danger zone)
  $("#topbarClearBtn").addEventListener("click", () => showClearModal("output-only"));
  $("#clearDataBtn").addEventListener("click", () => showClearModal("output-only"));
  $("#clearAllBtn").addEventListener("click", () => showClearModal("all"));
  
  // Modal controls
  confirmBtn.addEventListener("click", executeClear);
  cancelBtn.addEventListener("click", hideClearModal);
  
  // Close modal on backdrop click
  modal.addEventListener("click", e => {
    if (e.target === modal) hideClearModal();
  });
  
  // Close modal on Escape key
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && modal.style.display === "flex") {
      hideClearModal();
    }
  });
}

// ── Pipeline run buttons (wired globally) ─────────────────
function initPipelineButtons() {
  $("#runPipelineBtn").addEventListener("click", runPipeline);
  $("#topbarRunBtn").addEventListener("click", () => {
    // Switch to pipeline section and run
    activateSection("pipeline");
    $$(".nav-item").forEach(l => l.classList.remove("active"));
    $(".nav-item[data-section='pipeline']").classList.add("active");
    $("#topbarTitle").textContent = "Pipeline";
    runPipeline();
  });
}

// ── Refresh button ────────────────────────────────────────
function initRefreshButton() {
  const btn = $("#refreshBtn");
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    const icon = btn.querySelector("i");
    icon.style.animation = "spin 0.8s linear infinite";
    try {
      await loadDashboard();
      await loadFiles();
      showToast("Dashboard refreshed", "info");
    } finally {
      btn.disabled = false;
      icon.style.animation = "";
    }
  });
}

// Add spin keyframe dynamically
(function addSpinStyle() {
  const style = document.createElement("style");
  style.textContent = `@keyframes spin { to { transform: rotate(360deg); } }`;
  document.head.appendChild(style);
})();

// ── Boot ──────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  // Replace feather icons
  feather.replace({ "aria-hidden": "true" });

  initSidebar();
  initNav();
  initPipelineButtons();
  initLogStream();
  initDropZone();
  initRefreshButton();
  initClearData();
  await initSettings();
  await loadDashboard();
  await loadFiles();

  // Auto-refresh KPIs every 60 seconds
  setInterval(loadDashboard, 60_000);
  // Poll pipeline status when running
  setInterval(async () => {
    if (!state.pipelineRunning) return;
    try {
      const s = await api("/api/status");
      if (!s.running && state.pipelineRunning) {
        state.pipelineRunning = false;
        if (s.success) {
          setPipelineUI("success");
          showToast("Pipeline finished", "success");
          await loadDashboard();
        } else {
          setPipelineUI("error", s.error);
        }
      }
    } catch (_) {}
  }, 2000);
});
