// app.js
// -------------------- State --------------------
let startDate = new Date("2024-07-01");
let endDate   = new Date("2029-05-01");

// Project “Month 1” anchor (Aug 2024 = M1)
const projectM1 = new Date("2024-08-01");

// Cap max display to 2031
const CAP_MAX = new Date("2031-12-01");

let tasks = [];               // built from excel
const collapsed = new Set();  // collapsed group row indexes

// -------------------- Utils --------------------
function normHeader(s){
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g," ")
    .replace(/[_-]+/g," ")
    .replace(/[^\w\s]/g,"");
}

function pickKey(row, candidates){
  const keys = Object.keys(row);
  const norm = {};
  keys.forEach(k => norm[k] = normHeader(k));

  for (const c of candidates){
    const cN = normHeader(c);
    const exact = keys.find(k => norm[k] === cN);
    if (exact) return exact;
  }
  for (const c of candidates){
    const cN = normHeader(c);
    const partial = keys.find(k => norm[k].includes(cN));
    if (partial) return partial;
  }
  return null;
}

function detectColumns(firstRow){
  const idKey = pickKey(firstRow, ["Activity ID","Activity Id","ActivityID","ID"]);
  const nameKey = pickKey(firstRow, ["Activity Name","Activity","Name","Task Name","Description"]);
  const startKey = pickKey(firstRow, ["Start","Start Date","Planned Start","Baseline Start","Early Start"]);
  const finishKey = pickKey(firstRow, ["Finish","Finish Date","Planned Finish","Baseline Finish","Early Finish","End","End Date"]);
  const durKey = pickKey(firstRow, ["Original Duration","Planned Duration","Duration","Remaining Duration"]);
  return { idKey, nameKey, startKey, finishKey, durKey };
}

function parseP6Date(value) {
  if (value == null || value === "") return null;

  if (value instanceof Date && !isNaN(value)) return value;

  // Excel serial number
  if (typeof value === "number" && isFinite(value)) {
    const epoch = Date.UTC(1899, 11, 30);
    const ms = epoch + Math.round(value * 86400000);
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }

  const s = String(value).trim().replace(/\*/g, "");
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

function fmtDisplayDate(d){
  if (!d) return "—";
  const dd = String(d.getDate()).padStart(2,"0");
  const mon = d.toLocaleString("en-GB", { month: "short" }).toUpperCase();
  const yyyy = d.getFullYear();
  return `${dd} ${mon} ${yyyy}`;
}

function daysBetween(a, b) {
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
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

// --- Legend/category rules ---
const LEGEND_RULES = [
  { key: "piling", color: "bg-blue-700" },

  { key: "erss", color: "bg-amber-600" },
  { key: "excav", color: "bg-amber-600" },
  { key: "strut", color: "bg-amber-600" },
  { key: "diaphragm", color: "bg-amber-600" },

  { key: "concrete", color: "bg-orange-500" },
  { key: "basement", color: "bg-orange-500" },
  { key: "raft", color: "bg-orange-500" },
  { key: "slab", color: "bg-orange-500" },

  { key: "structure", color: "bg-emerald-600" },
  { key: "superstructure", color: "bg-emerald-600" },
  { key: "tower", color: "bg-emerald-600" },
  { key: "podium", color: "bg-emerald-600" },

  { key: "dts", color: "bg-pink-600" },
  { key: "viaduct", color: "bg-pink-600" },
  { key: "track", color: "bg-pink-600" },

  { key: "fit-out", color: "bg-indigo-600" },
  { key: "fit out", color: "bg-indigo-600" },
  { key: "facade", color: "bg-indigo-600" },
  { key: "mep", color: "bg-indigo-600" },
  { key: "testing", color: "bg-indigo-600" },
  { key: "commission", color: "bg-indigo-600" },

  { key: "as-built", color: "bg-purple-500" },
  { key: "as built", color: "bg-purple-500" },
  { key: "statutory", color: "bg-purple-500" },
  { key: "top", color: "bg-purple-500" },
];

function legendColorFromText(txt){
  const t = (txt || "").toLowerCase();
  for (const r of LEGEND_RULES) {
    if (t.includes(r.key)) return r.color;
  }
  return null;
}

function colorFromName(name) {
  return legendColorFromText(name) || "bg-slate-500";
}

function monthDiff(a, b) {
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
}

function getMonthColPx() {
  const cssVal = getComputedStyle(document.documentElement)
    .getPropertyValue("--month-col")
    .trim();
  return Number(cssVal.replace("px", "")) || 140;
}

function setMonthColPx(px){
  const v = Math.max(10, Math.min(280, Math.round(px)));
  document.documentElement.style.setProperty("--month-col", `${v}px`);

  const zoom = document.getElementById("zoomSlider");
  const zv = document.getElementById("zoomValue");
  if (zoom) zoom.value = String(v);
  if (zv) zv.textContent = `${v}px/mo`;
}

function timelineWidthPx() {
  const startM = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endM   = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  const cols = monthDiff(startM, endM) + 1;
  return cols * getMonthColPx();
}

function dateToXpx(dateISO) {
  const d = new Date(dateISO);

  const startM = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const dM     = new Date(d.getFullYear(), d.getMonth(), 1);

  const colsFromStart = monthDiff(startM, dM);
  const colPx = getMonthColPx();

  const monthStart = new Date(d.getFullYear(), d.getMonth(), 1);
  const monthEnd   = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  const frac = (d - monthStart) / (monthEnd - monthStart);

  return (colsFromStart + frac) * colPx;
}

// -------------------- Collapse logic --------------------
function isHiddenByCollapse(idx) {
  let myLevel = tasks[idx].level ?? 0;

  for (let i = idx - 1; i >= 0; i--) {
    if (!tasks[i]) continue;
    const lvl = tasks[i].level ?? 0;

    if (lvl < myLevel) {
      if (collapsed.has(i)) return true;
      myLevel = lvl;
    }
  }
  return false;
}

function toggleCollapse(idx) {
  if (collapsed.has(idx)) collapsed.delete(idx);
  else collapsed.add(idx);
}

// -------------------- Colors from hierarchy (groups above) --------------------
function colorFromHierarchy(idx, fallbackName){
  let myLevel = tasks[idx]?.level ?? 0;

  for (let i = idx - 1; i >= 0; i--) {
    const row = tasks[i];
    if (!row) continue;

    const lvl = row.level ?? 0;
    const isGroup = row.type === "header" || row.type === "sub-header";

    if (isGroup && lvl < myLevel) {
      const c = legendColorFromText(row.name);
      if (c) return c;
      myLevel = lvl;
    }
  }
  return colorFromName(fallbackName);
}

function applyHierarchyColors(){
  for (let i = 0; i < tasks.length; i++) {
    if (tasks[i].type === "header" || tasks[i].type === "sub-header") {
      tasks[i].groupColor = legendColorFromText(tasks[i].name) || null;
    }
  }
  for (let i = 0; i < tasks.length; i++) {
    if (tasks[i].type === "activity") {
      tasks[i].color = colorFromHierarchy(i, tasks[i].name);
    }
  }
}

// -------------------- Timeline Header (3 rows) --------------------
function buildTimelineHeader() {
  const header = document.getElementById("timelineHeader");
  if (!header) return;

  header.innerHTML = "";

  const startM = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endM   = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  const totalCols = monthDiff(startM, endM) + 1;
  const colPx = getMonthColPx();

  const yearRow = document.createElement("div");
  yearRow.className = "timeline-header-row";

  const monthRow = document.createElement("div");
  monthRow.className = "timeline-header-row";

  const mRow = document.createElement("div");
  mRow.className = "timeline-header-row";

  const months = [];
  for (let i = 0; i < totalCols; i++) {
    months.push(new Date(startM.getFullYear(), startM.getMonth() + i, 1));
  }

  // Year spans
  let i = 0;
  while (i < months.length) {
    const y = months[i].getFullYear();
    let j = i;
    while (j < months.length && months[j].getFullYear() === y) j++;

    const span = j - i;
    const cell = document.createElement("div");
    cell.className = "th-cell";
    cell.style.width = `calc(${span} * ${colPx}px)`;
    cell.style.minWidth = `calc(${span} * ${colPx}px)`;
    cell.textContent = y;
    yearRow.appendChild(cell);

    i = j;
  }

  const fmtMonth = new Intl.DateTimeFormat("en", { month: "short" });

  months.forEach((d) => {
    const mCell = document.createElement("div");
    mCell.className = "th-cell small";
    mCell.textContent = fmtMonth.format(d).toUpperCase();
    monthRow.appendChild(mCell);

    const mIndex = monthDiff(projectM1, d) + 1; // Aug24 => 1
    const mmCell = document.createElement("div");
    mmCell.className = "th-cell mrow";
    mmCell.textContent = `M${mIndex}`;
    mRow.appendChild(mmCell);
  });

  header.appendChild(yearRow);
  header.appendChild(monthRow);
  header.appendChild(mRow);
}

// -------------------- Render rows (labels + bars) --------------------
function buildRows() {
  const labelCont = document.getElementById("labelRows");
  const barCont   = document.getElementById("barRows");
  const status    = document.getElementById("status");

  if (!labelCont || !barCont) return;

  labelCont.innerHTML = "";
  barCont.innerHTML = "";

  const totalWidth = timelineWidthPx();
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

    // LABEL ROW
    const labelRow = document.createElement("div");
    labelRow.className = "row label-row";

    if (isGroup) {
      labelRow.classList.add("group", "clickable");
      if (task.type === "header") labelRow.classList.add("header");
    }

    labelRow.style.paddingLeft = `${padLeft}px`;

    const caret = document.createElement("span");
    caret.className = "caret";
    caret.textContent = isGroup ? (isCollapsed ? "▶" : "▼") : "";
    labelRow.appendChild(caret);

    const text = document.createElement("span");
    text.className = "label-text";
    text.textContent = task.name;
    text.title = task.name;
    labelRow.appendChild(text);

    if (isGroup) {
      labelRow.addEventListener("click", () => {
        toggleCollapse(idx);
        buildRows();
      });
    }

    labelCont.appendChild(labelRow);

    // BAR ROW
    const barRow = document.createElement("div");
    barRow.className = "row bar-row";
    barRow.style.width = `${totalWidth}px`;

    // Summary bar for groups
    if (isGroup && task.start && task.end) {
      const leftPx = dateToXpx(task.start);
      const rightPx = dateToXpx(task.end);
      const w = Math.max(2, rightPx - leftPx);

      const sbar = document.createElement("div");
      sbar.className = "summary-bar";
      sbar.style.left = `${leftPx}px`;
      sbar.style.width = `${w}px`;

      const gc = task.groupColor || legendColorFromText(task.name);
      if (gc) {
        sbar.classList.add(gc);
        sbar.style.opacity = "0.25";
      }

      sbar.title = `${task.name} (summary)`;
      barRow.appendChild(sbar);
    }

    // Milestone
    if (!isGroup && task.type === "milestone" && task.start) {
      const x = dateToXpx(task.start);

      const diamond = document.createElement("div");
      diamond.className = "milestone-diamond";
      diamond.style.left = `${x - 6}px`;
      barRow.appendChild(diamond);

      const ms = document.createElement("span");
      ms.className = "ms-label";
      ms.style.left = `${x + 14}px`;
      ms.textContent = new Date(task.start).toLocaleDateString("en-GB", {
        day: "2-digit", month: "short", year: "2-digit",
      });
      barRow.appendChild(ms);
    }

    // Activity bar
    if (!isGroup && task.type === "activity" && task.start && task.end) {
      const leftPx = dateToXpx(task.start);
      const rightPx = dateToXpx(task.end);
      const wPx = Math.max(2, rightPx - leftPx);

      const bar = document.createElement("div");
      bar.className = `gantt-bar ${task.color || "bg-slate-500"}`;
      bar.style.left = `${leftPx}px`;
      bar.style.width = `${wPx}px`;

      const dStart = new Date(task.start);
      const dEnd = new Date(task.end);
      const days = daysBetween(dStart, dEnd);

      bar.textContent = wPx > 80 ? `${days}d` : "";
      bar.title = `${task.name}: ${days} days`;
      barRow.appendChild(bar);
    }

    barCont.appendChild(barRow);
  });

  if (status) {
    status.textContent =
      `Done. Loaded ${tasks.length} rows. Click group rows to collapse.` +
      (hiddenCount ? ` (${hiddenCount} hidden)` : "");
  }
}

function renderAll() {
  buildTimelineHeader();
  buildRows();
}

// -------------------- Excel -> tasks --------------------
function inferRowType(activityIdRaw, startD, finishD, durationRaw, indentLevel) {
  const idTrim = String(activityIdRaw ?? "").trim();
  const durNum = durationRaw === "" || durationRaw == null ? null : Number(durationRaw);

  if (startD && (durNum === 0 || !finishD)) return "milestone";
  if (looksLikeActivityId(idTrim) && startD && finishD) return "activity";
  if (!looksLikeActivityId(idTrim)) return indentLevel <= 0 ? "header" : "sub-header";

  if (startD && finishD) return "activity";
  return indentLevel <= 0 ? "header" : "sub-header";
}

function buildTasksFromRows(rows) {
  const newTasks = [];
  let minD = null;
  let maxD = null;

  const cols = detectColumns(rows[0] || {});
  const { idKey, nameKey, startKey, finishKey, durKey } = cols;

  if (!nameKey || (!startKey && !finishKey)) {
    throw new Error("Could not detect required columns. Need at least: Activity Name, Start, Finish.");
  }

  rows.forEach((r) => {
    const rawId = idKey ? r[idKey] : "";
    const rawName = nameKey ? r[nameKey] : "";
    const rawStart = startKey ? r[startKey] : "";
    const rawFinish = finishKey ? r[finishKey] : "";
    const rawDur = durKey ? r[durKey] : "";

    const indentCount = leadingIndentCount(rawId) || leadingIndentCount(rawName);
    const level = indentToLevel(indentCount);

    const idTrim = String(rawId || "").trim();
    const nameTrim = String(rawName || "").trim();
    if (!idTrim && !nameTrim && !rawStart && !rawFinish) return;

    const startD = parseP6Date(rawStart);
    const finishD = parseP6Date(rawFinish);

    const type = inferRowType(rawId, startD, finishD, rawDur, level);
    const title = nameTrim || idTrim;

    if (type === "header" || type === "sub-header") {
      const obj = { name: title, type, level };
      if (startD) obj.start = toISODate(startD);
      if (finishD) obj.end = toISODate(finishD);

      newTasks.push(obj);

      if (startD) minD = !minD || startD < minD ? startD : minD;
      if (finishD) maxD = !maxD || finishD > maxD ? finishD : maxD;
      return;
    }

    if (type === "milestone") {
      const d = startD || finishD;
      if (d) {
        newTasks.push({ name: title, type: "milestone", start: toISODate(d), level });
        minD = !minD || d < minD ? d : minD;
        maxD = !maxD || d > maxD ? d : maxD;
      }
      return;
    }

    if (startD && finishD) {
      newTasks.push({
        name: title,
        type: "activity",
        start: toISODate(startD),
        end: toISODate(finishD),
        level
      });

      minD = !minD || startD < minD ? startD : minD;
      maxD = !maxD || finishD > maxD ? finishD : maxD;
    }
  });

  if (minD && maxD) {
    const padStart = new Date(minD);
    padStart.setMonth(padStart.getMonth() - 2);

    const padEnd = new Date(maxD);
    padEnd.setMonth(padEnd.getMonth() + 2);

    startDate = padStart;
    endDate = padEnd > CAP_MAX ? CAP_MAX : padEnd;
  }

  // Update header stats (Commencement / Completion / Duration)
  const commEl = document.getElementById("commDate");
  const compEl = document.getElementById("compDate");
  const durEl  = document.getElementById("totalDur");

  if (commEl) commEl.textContent = fmtDisplayDate(minD);
  if (compEl) compEl.textContent = fmtDisplayDate(maxD);
  if (durEl && minD && maxD) durEl.textContent = `${daysBetween(minD, maxD)} DAYS`;

  return newTasks;
}

// -------------------- Resizable splitter (P6-like) --------------------
function setupSplitter() {
  const splitter = document.getElementById("splitter");
  const leftPane = document.getElementById("leftPane");
  if (!splitter || !leftPane) return;

  let dragging = false;

  splitter.addEventListener("mousedown", (e) => {
    dragging = true;
    document.body.style.userSelect = "none";
    e.preventDefault();
  });

  window.addEventListener("mousemove", (e) => {
    if (!dragging) return;
    const min = 240;
    const max = Math.min(window.innerWidth * 0.7, 900);
    const w = Math.max(min, Math.min(max, e.clientX));
    leftPane.style.width = `${w}px`;
  });

  window.addEventListener("mouseup", () => {
    if (!dragging) return;
    dragging = false;
    document.body.style.userSelect = "";
  });
}

// -------------------- Timeline header drag-to-zoom --------------------
function setupHeaderDragZoom(){
  const header = document.getElementById("timelineHeader");
  if (!header) return;

  let isDown = false;
  let startX = 0;
  let startW = 140;

  header.addEventListener("mousedown", (e) => {
    isDown = true;
    startX = e.clientX;
    startW = getMonthColPx();
    document.body.style.userSelect = "none";
  });

  window.addEventListener("mousemove", (e) => {
    if (!isDown) return;
    const dx = e.clientX - startX;
    const next = startW + dx * 0.7;
    setMonthColPx(next);
    renderAll();
  });

  window.addEventListener("mouseup", () => {
    if (!isDown) return;
    isDown = false;
    document.body.style.userSelect = "";
  });
}

// -------------------- Fit To Screen --------------------
function fitToScreen(){
  const pane = document.getElementById("timelinePane");
  if (!pane) return;

  const startM = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endM   = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  const cols = monthDiff(startM, endM) + 1;
  if (!cols) return;

  const avail = pane.clientWidth;
  const px = Math.floor(avail / cols);

  setMonthColPx(px);
  renderAll();
  pane.scrollLeft = 0;
}

// -------------------- Wire up UI --------------------
document.addEventListener("DOMContentLoaded", () => {
  const status = document.getElementById("status");
  const fileInput = document.getElementById("excelFile");
  const loadBtn = document.getElementById("loadBtn");
  const expandAllBtn = document.getElementById("expandAllBtn");
  const fitBtn = document.getElementById("fitBtn");
  const zoomSlider = document.getElementById("zoomSlider");

  const timelinePane = document.getElementById("timelinePane");
  const labelRows = document.getElementById("labelRows");

  setupSplitter();
  setupHeaderDragZoom();

  if (zoomSlider) {
    zoomSlider.addEventListener("input", () => {
      setMonthColPx(Number(zoomSlider.value));
      renderAll();
    });
  }

  if (fitBtn) fitBtn.addEventListener("click", fitToScreen);

  // Sync vertical scrolling between left labels and right timeline
  if (timelinePane && labelRows) {
    let lock = false;
    timelinePane.addEventListener("scroll", () => {
      if (lock) return;
      lock = true;
      labelRows.scrollTop = timelinePane.scrollTop;
      lock = false;
    });
    labelRows.addEventListener("scroll", () => {
      if (lock) return;
      lock = true;
      timelinePane.scrollTop = labelRows.scrollTop;
      lock = false;
    });
  }

  if (expandAllBtn) {
    expandAllBtn.addEventListener("click", () => {
      collapsed.clear();
      if (status) status.textContent = "Expanded all.";
      buildRows();
    });
  }

  if (loadBtn) {
    loadBtn.addEventListener("click", async () => {
      const f = fileInput.files?.[0];
      if (!f) {
        if (status) status.textContent = "Please choose an Excel file first.";
        return;
      }

      if (status) status.textContent = "Reading Excel…";
      collapsed.clear();

      try {
        const data = await f.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array", cellDates: true });

        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (!rows || rows.length === 0) throw new Error("Sheet is empty.");

        tasks = buildTasksFromRows(rows);

        // assign legend colors by hierarchy + group colors for summary bars
        applyHierarchyColors();

        if (status) status.textContent = `Loaded ${tasks.length} rows. Rendering…`;
        renderAll();

        // Auto scroll near Aug-24 so you see M1
        if (timelinePane) {
          const px = dateToXpx(toISODate(projectM1));
          timelinePane.scrollLeft = Math.max(0, px - 250);
        }

        if (status) status.textContent = `Done. Loaded ${tasks.length} rows. Click group rows to collapse.`;
      } catch (e) {
        console.error(e);
        if (status) status.textContent = `Error loading Excel: ${e.message || e}`;
      }
    });
  }

  // initial render
  renderAll();
});