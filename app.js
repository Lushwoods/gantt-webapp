// -------------------- State --------------------
let startDate = new Date("2024-07-01");
let endDate = new Date("2029-05-01");
let tasks = [];

// collapsed node indexes (in tasks array)
const collapsed = new Set();

// -------------------- Utils --------------------
function cleanDateStr(s) {
  if (!s) return "";
  return String(s).trim().replace(/\*/g, "");
}

function parseP6Date(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value)) return value;

  const s = cleanDateStr(value);
  if (!s) return null;

  const m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m) {
    const dd = m[1].padStart(2, "0");
    const mon = m[2];
    let yy = m[3];
    if (yy.length === 2) {
      const n = Number(yy);
      yy = String(n < 70 ? 2000 + n : 1900 + n);
    }
    const d = new Date(`${dd} ${mon} ${yy}`);
    return isNaN(d) ? null : d;
  }

  const d2 = new Date(s);
  return isNaN(d2) ? null : d2;
}

function toISODate(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function daysBetween(a, b) {
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function getPosition(dateStr) {
  const date = new Date(dateStr);
  const diff = date - startDate;
  const totalDiff = endDate - startDate;
  return (diff / totalDiff) * 100;
}

function leadingIndentCount(str) {
  if (!str) return 0;
  const s = String(str);
  let count = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    const code = s.charCodeAt(i);
    if (ch === " " || ch === "\t" || code === 160) count++;
    else break;
  }
  return count;
}

function indentToLevel(indentCount) {
  return Math.floor(indentCount / 2);
}

function looksLikeActivityId(s) {
  if (!s) return false;
  const t = String(s).trim();
  return /^P\d+_[A-Z]{2,}[_-]?\d+/.test(t);
}

function colorFromName(name) {
  const n = (name || "").toLowerCase();
  if (n.includes("piling")) return "bg-blue-700";
  if (n.includes("excav") || n.includes("erss") || n.includes("strut")) return "bg-amber-600";
  if (n.includes("basement") || n.includes("raft") || n.includes("slab")) return "bg-orange-500";
  if (n.includes("tower") || n.includes("podium") || n.includes("structure")) return "bg-emerald-600";
  if (n.includes("dts") || n.includes("viaduct") || n.includes("track")) return "bg-pink-600";
  if (n.includes("facade") || n.includes("fit") || n.includes("mep")) return "bg-indigo-600";
  if (n.includes("as-built") || n.includes("statutory")) return "bg-purple-500";
  return "bg-slate-500";
}

// -------------------- Timeline header --------------------
function buildTimelineHeader() {
  const header = document.getElementById("timelineHeader");
  if (!header) return;
  header.innerHTML = "";

  const months = [{ label: "Jan" }, { label: "Jul" }];
  const yStart = startDate.getFullYear();
  const yEnd = endDate.getFullYear();

  for (let y = yStart; y <= yEnd; y++) {
    months.forEach((m) => {
      const div = document.createElement("div");
      div.className = "timeline-cell";
      div.innerText = `${m.label} ${y}`;
      header.appendChild(div);
    });
  }
}

// -------------------- Collapse logic --------------------
function isHiddenByCollapse(idx) {
  const myLevel = tasks[idx].level ?? 0;

  for (let i = idx - 1; i >= 0; i--) {
    if (!tasks[i]) continue;
    const lvl = tasks[i].level ?? 0;

    // Only ancestors matter
    if (lvl < myLevel) {
      if (collapsed.has(i)) return true;
      myLevel = lvl; // walk up
    }
  }
  return false;
}

function toggleCollapse(idx) {
  if (collapsed.has(idx)) collapsed.delete(idx);
  else collapsed.add(idx);
}

// -------------------- Render gantt rows --------------------
function buildGanttRows() {
  const container = document.getElementById("ganttRows");
  if (!container) return;
  container.innerHTML = "";

  let hiddenCount = 0;

  tasks.forEach((task, idx) => {
    if (isHiddenByCollapse(idx)) {
      hiddenCount++;
      return;
    }

    const level = typeof task.level === "number" ? task.level : 0;
    const padLeft = 12 + level * 18;

    const isGroup = task.type === "header" || task.type === "sub-header";
    const isCollapsed = collapsed.has(idx);

    // ---------- Label ----------
    const label = document.createElement("div");
    label.className = `label-col border-b ${isGroup ? (task.type === "header" ? "section-header" : "sub-header") : ""}`;

    // Make group rows clickable (collapse/expand)
    if (isGroup) label.classList.add("clickable-header");

    label.style.paddingLeft = `${padLeft}px`;

    if (isGroup) {
      const caret = document.createElement("span");
      caret.className = "caret";
      caret.textContent = isCollapsed ? "▶" : "▼";
      label.appendChild(caret);
    }

    const text = document.createElement("span");
    text.textContent = task.name;
    label.appendChild(text);

    if (isGroup) {
      label.addEventListener("click", () => {
        toggleCollapse(idx);
        buildGanttRows();
        // keep header static
      });
    }

    container.appendChild(label);

    // ---------- Bar container ----------
    const barCont = document.createElement("div");
    barCont.className = `bar-container ${isGroup ? (task.type === "header" ? "section-header" : "sub-header") : ""} border-b`;

    // Summary bars for group rows (if start/end present)
    if (isGroup && task.start && task.end) {
      const startPos = getPosition(task.start);
      const endPos = getPosition(task.end);
      const width = Math.max(0.5, endPos - startPos);

      const sbar = document.createElement("div");
      sbar.className = "summary-bar";
      sbar.style.left = `${startPos}%`;
      sbar.style.width = `${width}%`;
      sbar.title = `${task.name} (summary)`;
      barCont.appendChild(sbar);
    }

    // Milestone
    if (!isGroup && task.type === "milestone") {
      const pos = getPosition(task.start);
      const diamond = document.createElement("div");
      diamond.className = "milestone-diamond";
      diamond.style.left = `calc(${pos}% - 4px)`;
      barCont.appendChild(diamond);

      const dateLabel = document.createElement("span");
      dateLabel.className = "absolute text-[7px] text-red-600 font-black";
      dateLabel.style.left = `calc(${pos}% + 10px)`;
      dateLabel.innerText = new Date(task.start).toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "2-digit",
      });
      barCont.appendChild(dateLabel);
    }

    // Activity bar
    if (!isGroup && task.type === "activity") {
      const startPos = getPosition(task.start);
      const endPos = getPosition(task.end);
      const width = Math.max(0.5, endPos - startPos);

      const bar = document.createElement("div");
      bar.className = `gantt-bar ${task.color || "bg-slate-500"}`;
      bar.style.left = `${startPos}%`;
      bar.style.width = `${width}%`;

      const dStart = new Date(task.start);
      const dEnd = new Date(task.end);
      const days = daysBetween(dStart, dEnd);

      bar.innerText = width > 2 ? `${days}d` : "";
      bar.title = `${task.name}: ${days} days`;
      barCont.appendChild(bar);
    }

    container.appendChild(barCont);
  });

  const status = document.getElementById("status");
  if (status && hiddenCount > 0) {
    status.textContent = `${status.textContent} (${hiddenCount} hidden by collapse)`;
  }
}

function renderAll() {
  buildTimelineHeader();
  buildGanttRows();
}

// -------------------- Excel -> tasks --------------------
function inferRowType(activityIdRaw, startD, finishD, durationRaw, indentLevel) {
  const idTrim = String(activityIdRaw ?? "").trim();
  const durNum = durationRaw === "" || durationRaw == null ? null : Number(durationRaw);

  // Milestone
  if (startD && (durNum === 0 || !finishD)) return "milestone";

  // Activity
  if (looksLikeActivityId(idTrim) && startD && finishD) return "activity";

  // Group/WBS
  if (!looksLikeActivityId(idTrim)) return indentLevel <= 0 ? "header" : "sub-header";

  // fallback
  if (startD && finishD) return "activity";
  return indentLevel <= 0 ? "header" : "sub-header";
}

function buildTasksFromRows(rows) {
  const newTasks = [];
  let minD = null;
  let maxD = null;

  rows.forEach((r) => {
    const rawId = r["Activity ID"] ?? r["Activity Id"] ?? r["ActivityID"] ?? "";
    const rawName = r["Activity Name"] ?? r["Activity"] ?? r["Name"] ?? "";
    const rawStart = r["Start"] ?? r["Start Date"] ?? "";
    const rawFinish = r["Finish"] ?? r["Finish Date"] ?? "";
    const rawDur =
      r["Original Duration"] ??
      r["Planned Duration"] ??
      r["Planned Dur."] ??
      r["Duration"] ??
      "";

    const indentCount = leadingIndentCount(rawId) || leadingIndentCount(rawName);
    const level = indentToLevel(indentCount);

    const idTrim = String(rawId || "").trim();
    const nameTrim = String(rawName || "").trim();

    if (!idTrim && !nameTrim && !rawStart && !rawFinish) return;

    const startD = parseP6Date(rawStart);
    const finishD = parseP6Date(rawFinish);

    const type = inferRowType(rawId, startD, finishD, rawDur, level);
    const title = nameTrim || idTrim;

    // Group rows: keep summary dates if present so we can draw summary bars
    if (type === "header" || type === "sub-header") {
      const obj = { name: title, type, level };

      if (startD) obj.start = toISODate(startD);
      if (finishD) obj.end = toISODate(finishD);

      newTasks.push(obj);

      // chart bounds also consider summary dates
      if (startD) minD = !minD || startD < minD ? startD : minD;
      if (finishD) maxD = !maxD || finishD > maxD ? finishD : maxD;

      return;
    }

    if (type === "milestone") {
      if (startD) {
        newTasks.push({ name: title, type: "milestone", start: toISODate(startD), level });
        minD = !minD || startD < minD ? startD : minD;
        maxD = !maxD || startD > maxD ? startD : maxD;
      }
      return;
    }

    if (startD && finishD) {
      newTasks.push({
        name: title,
        type: "activity",
        start: toISODate(startD),
        end: toISODate(finishD),
        level,
        color: colorFromName(title),
      });

      minD = !minD || startD < minD ? startD : minD;
      maxD = !maxD || finishD > maxD ? finishD : maxD;
    }
  });

  // Auto chart range
  if (minD && maxD) {
    const padStart = new Date(minD);
    padStart.setMonth(padStart.getMonth() - 1);

    const padEnd = new Date(maxD);
    padEnd.setMonth(padEnd.getMonth() + 1);

    startDate = padStart;
    endDate = padEnd;
  }

  return newTasks;
}

// -------------------- Wire up UI --------------------
document.addEventListener("DOMContentLoaded", () => {
  const status = document.getElementById("status");
  const fileInput = document.getElementById("excelFile");
  const loadBtn = document.getElementById("loadBtn");
  const expandAllBtn = document.getElementById("expandAllBtn");

  expandAllBtn.addEventListener("click", () => {
    collapsed.clear();
    if (status) status.textContent = "Expanded all.";
    buildGanttRows();
  });

  loadBtn.addEventListener("click", async () => {
    const f = fileInput.files?.[0];
    if (!f) {
      status.textContent = "Please choose an Excel file first.";
      return;
    }

    status.textContent = "Reading Excel…";
    collapsed.clear(); // reset collapse on new load

    const data = await f.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    tasks = buildTasksFromRows(rows);

    status.textContent = `Loaded ${tasks.length} rows. Rendering…`;
    renderAll();
    status.textContent = `Done. Loaded ${tasks.length} rows. Click headers to collapse.`;
  });

  renderAll();
});
