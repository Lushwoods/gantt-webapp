// app.js
// -------------------- State --------------------
let startDate = new Date("2024-07-01");
let endDate   = new Date("2029-05-01");

// Cap max display to 2031
const CAP_MAX = new Date("2031-12-01");

let tasks = [];               // built from excel
const collapsed = new Set();  // collapsed group row indexes

// Project Month Anchor (for Month row: first visible month = M1 per upload)
let projectMonthAnchor = null; // Date (YYYY-MM-01) set from detected project start

// -------------------- High-level view state --------------------
let viewMode = "detailed";      // "detailed" | "high"
let highLevelCutoff = 2;        // default depth
let showMilestonesInHigh = true;
let savedCollapsed = null;      // remembers collapse state when toggling back to detailed

// -------------------- Dynamic legend (auto from file) --------------------
const AUTO_PALETTE = [
  "bg-blue-700","bg-amber-600","bg-orange-500","bg-emerald-600",
  "bg-pink-600","bg-indigo-600","bg-purple-500","bg-teal-600",
  "bg-rose-600","bg-cyan-600","bg-lime-600","bg-violet-600",
  "bg-sky-600","bg-fuchsia-600","bg-green-700","bg-yellow-600"
];

let legendMode = "auto";     // "auto" | "keywords"
let legendFieldKey = null;   // detected column name from the file
let legendMap = new Map();   // value -> tailwind bg class

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

function detectLegendColumn(firstRow){
  const candidates = [
    "Phase","Discipline","Area","Category","Trade","Package",
    "Workstream","Zone","Location","System","Stage",
    "Activity Code","Activity Codes","Activity Code -"
  ];

  const keys = Object.keys(firstRow || {});
  const norm = {};
  keys.forEach(k => norm[k] = normHeader(k));

  // Exact match first
  for (const c of candidates){
    const cN = normHeader(c);
    const exact = keys.find(k => norm[k] === cN);
    if (exact) return exact;
  }

  // Partial match, e.g. "Activity Code - Discipline"
  for (const c of candidates){
    const cN = normHeader(c);
    const partial = keys.find(k => norm[k].includes(cN));
    if (partial) return partial;
  }

  return null;
}

function buildLegendMapFromRows(rows){
  legendMap.clear();
  legendFieldKey = detectLegendColumn(rows?.[0] || {});
  if (!legendFieldKey) return;

  const vals = [];
  for (const r of rows) {
    const v = String(r[legendFieldKey] ?? "").trim();
    if (v) vals.push(v);
  }

  const unique = [...new Set(vals)];
  unique.forEach((val, i) => {
    legendMap.set(val, AUTO_PALETTE[i % AUTO_PALETTE.length]);
  });
}

function renderLegend(){
  const itemsEl = document.getElementById("legendItems");
  if (!itemsEl) return;

  itemsEl.innerHTML = "";

  // If keywords mode, show your hardcoded categories
  if (legendMode === "keywords" || legendMap.size === 0) {
    const fallback = [
      ["Piling", "bg-blue-700"],
      ["ERSS/Excavation", "bg-amber-600"],
      ["Concrete/Basement", "bg-orange-500"],
      ["Structure", "bg-emerald-600"],
      ["DTS Viaduct", "bg-pink-600"],
      ["Fit-out/Facade", "bg-indigo-600"],
      ["As-Builts", "bg-purple-500"],
      ["Default", "bg-slate-500"],
    ];

    fallback.forEach(([label, color]) => {
      const div = document.createElement("div");
      div.className = "legend-item";
      div.innerHTML = `<span class="dot ${color}"></span> ${label}`;
      itemsEl.appendChild(div);
    });

    return;
  }

  // Auto legend from file
  // Sort alphabetically for stability
  const entries = [...legendMap.entries()].sort((a,b) => a[0].localeCompare(b[0]));

  // If too many, cap display
  const MAX_SHOW = 18;
  const shown = entries.slice(0, MAX_SHOW);

  shown.forEach(([label, color]) => {
    const div = document.createElement("div");
    div.className = "legend-item";
    div.innerHTML = `<span class="dot ${color}"></span> ${escapeHtml(label)}`;
    itemsEl.appendChild(div);
  });

  if (entries.length > MAX_SHOW) {
    const more = document.createElement("div");
    more.className = "legend-item";
    more.style.color = "#64748b";
    more.textContent = `+${entries.length - MAX_SHOW} more`;
    itemsEl.appendChild(more);
  }
}

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
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

// --- Keyword legend/category rules (fallback) ---
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

// -------------------- High-level helpers --------------------
function isGroupTask(task) {
  return task.type === "header" || task.type === "sub-header";
}

function isAllowedByViewMode(task) {
  if (viewMode === "detailed") return true;

  if (isGroupTask(task)) return (task.level ?? 0) <= highLevelCutoff;
  if (showMilestonesInHigh && task.type === "milestone") return true;
  return (task.level ?? 0) <= highLevelCutoff + 1;
}

function applyHighLevelCollapse() {
  collapsed.clear();
  for (let idx = 0; idx < tasks.length; idx++) {
    const t = tasks[idx];
    if (!t) continue;
    if (isGroupTask(t) && (t.level ?? 0) >= highLevelCutoff) {
      collapsed.add(idx);
    }
  }
}

function enableHighLevel() {
  viewMode = "high";
  savedCollapsed = new Set(collapsed);
  applyHighLevelCollapse();
}

function disableHighLevel() {
  viewMode = "detailed";
  if (savedCollapsed) {
    collapsed.clear();
    for (const idx of savedCollapsed) collapsed.add(idx);
  }
}

function buildHighLevelControls(hostEl) {
  if (!hostEl) return;
  if (document.getElementById("btnHighLevel")) return;

  const wrap = document.createElement("div");
  wrap.style.display = "inline-flex";
  wrap.style.alignItems = "center";
  wrap.style.gap = "10px";
  wrap.style.marginLeft = "10px";
  wrap.style.flexWrap = "wrap";

  const btn = document.createElement("button");
  btn.id = "btnHighLevel";
  btn.type = "button";
  btn.textContent = "High Level: OFF";
  btn.style.padding = "8px 12px";
  btn.style.borderRadius = "10px";
  btn.style.border = "1px solid #cbd5e1";
  btn.style.background = "white";
  btn.style.color = "#334155";
  btn.style.cursor = "pointer";
  btn.style.boxShadow = "0 1px 2px rgba(0,0,0,0.06)";

  btn.addEventListener("click", () => {
    if (viewMode === "detailed") {
      enableHighLevel();
      btn.textContent = "High Level: ON";
    } else {
      disableHighLevel();
      btn.textContent = "High Level: OFF";
    }
    renderAll();
  });

  const sel = document.createElement("select");
  sel.id = "highLevelSelect";
  sel.style.padding = "8px 10px";
  sel.style.borderRadius = "10px";
  sel.style.border = "1px solid #cbd5e1";
  sel.style.background = "white";
  sel.style.color = "#334155";
  sel.style.cursor = "pointer";

  [1, 2, 3, 4].forEach((lvl) => {
    const opt = document.createElement("option");
    opt.value = String(lvl);
    opt.textContent = `Level ${lvl}`;
    if (lvl === highLevelCutoff) opt.selected = true;
    sel.appendChild(opt);
  });

  sel.addEventListener("change", (e) => {
    highLevelCutoff = parseInt(e.target.value, 10) || 2;
    if (viewMode === "high") {
      applyHighLevelCollapse();
      renderAll();
    }
  });

  const lab = document.createElement("label");
  lab.style.display = "inline-flex";
  lab.style.alignItems = "center";
  lab.style.gap = "6px";
  lab.style.fontSize = "13px";
  lab.style.color = "#475569";
  lab.style.userSelect = "none";

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = showMilestonesInHigh;
  cb.addEventListener("change", (e) => {
    showMilestonesInHigh = !!e.target.checked;
    if (viewMode === "high") renderAll();
  });

  const txt = document.createElement("span");
  txt.textContent = "Show milestones";

  lab.appendChild(cb);
  lab.appendChild(txt);

  wrap.appendChild(btn);
  wrap.appendChild(sel);
  wrap.appendChild(lab);

  hostEl.appendChild(wrap);
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
      // Prefer group's own auto legend if present
      if (legendMode === "auto" && row.legendValue && legendMap.has(row.legendValue)) {
        return legendMap.get(row.legendValue);
      }

      // Otherwise keyword from group name
      const c = legendColorFromText(row.name);
      if (c) return c;

      myLevel = lvl;
    }
  }

  // fallback keyword
  return colorFromName(fallbackName);
}

function applyHierarchyColors(){
  // groupColor used for summary bars
  for (let i = 0; i < tasks.length; i++) {
    if (tasks[i].type === "header" || tasks[i].type === "sub-header") {
      // Prefer auto legend mapping if available
      if (legendMode === "auto" && tasks[i].legendValue && legendMap.has(tasks[i].legendValue)) {
        tasks[i].groupColor = legendMap.get(tasks[i].legendValue);
      } else {
        tasks[i].groupColor = legendColorFromText(tasks[i].name) || null;
      }
    }
  }

  for (let i = 0; i < tasks.length; i++) {
    if (tasks[i].type === "activity") {
      // Priority: explicit legendValue -> map -> inherit -> keyword fallback
      const v = tasks[i].legendValue;
      if (legendMode === "auto" && v && legendMap.has(v)) {
        tasks[i].color = legendMap.get(v);
      } else {
        tasks[i].color = colorFromHierarchy(i, tasks[i].name);
      }
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
  const anchor = projectMonthAnchor || startM; // B2: project-based M1 anchor

  months.forEach((d) => {
    const mCell = document.createElement("div");
    mCell.className = "th-cell small";
    mCell.textContent = fmtMonth.format(d).toUpperCase();
    monthRow.appendChild(mCell);

    const mIndex = monthDiff(anchor, d) + 1; // project start month => M1
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
    // View mode filter first
    if (!isAllowedByViewMode(task)) {
      hiddenCount++;
      return;
    }

    // Collapse filter
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

      const sTxt = fmtDisplayDate(new Date(task.start));
      const eTxt = fmtDisplayDate(new Date(task.end));
      sbar.title = `${task.name} (summary)\nStart: ${sTxt}\nFinish: ${eTxt}`;

      barRow.appendChild(sbar);
    }

    // Milestone
    if (!isGroup && task.type === "milestone" && task.start) {
      const x = dateToXpx(task.start);

      const diamond = document.createElement("div");
      diamond.className = "milestone-diamond";
      diamond.style.left = `${x - 6}px`;
      diamond.title = `${task.name}\nDate: ${fmtDisplayDate(new Date(task.start))}`;
      barRow.appendChild(diamond);

      const ms = document.createElement("span");
      ms.className = "ms-label";
      ms.style.left = `${x + 14}px`;
      ms.textContent = new Date(task.start).toLocaleDateString("en-GB", {
        day: "2-digit", month: "short", year: "2-digit",
      });
      ms.title = diamond.title;
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

      const sTxt = fmtDisplayDate(dStart);
      const eTxt = fmtDisplayDate(dEnd);
      bar.title = `${task.name}\nStart: ${sTxt}\nFinish: ${eTxt}\nDuration: ${days} days`;

      barRow.appendChild(bar);
    }

    barCont.appendChild(barRow);
  });

  if (status) {
    const modeTxt = viewMode === "high" ? `High-level (L${highLevelCutoff})` : "Detailed";
    status.textContent =
      `Done. Loaded ${tasks.length} rows. View: ${modeTxt}. Click group rows to collapse.` +
      (hiddenCount ? ` (${hiddenCount} hidden)` : "");
  }
}

function renderAll() {
  buildTimelineHeader();
  buildRows();
  renderLegend();
}

// -------------------- Excel -> tasks --------------------
function inferRowType(activityIdRaw, startD, finishD, durationRaw, indentLevel) {
  const idTrim = String(activityIdRaw ?? "").trim();
  const durNum = (durationRaw === "" || durationRaw == null) ? null : Number(durationRaw);

  // Milestone if has start and either duration is 0 or finish missing
  if (startD && (durNum === 0 || !finishD)) return "milestone";

  // Typical activity
  if (looksLikeActivityId(idTrim) && startD && finishD) return "activity";

  // WBS/group rows (no activity id pattern)
  if (!looksLikeActivityId(idTrim)) return indentLevel <= 0 ? "header" : "sub-header";

  // Fallbacks
  if (startD && finishD) return "activity";
  return indentLevel <= 0 ? "header" : "sub-header";
}

function deriveGroupSummariesInPlace(list){
  for (let i = 0; i < list.length; i++) {
    const t = list[i];
    if (!t) continue;
    if (!(t.type === "header" || t.type === "sub-header")) continue;

    const myLevel = t.level ?? 0;
    let min = null, max = null;

    for (let j = i + 1; j < list.length; j++) {
      const c = list[j];
      if (!c) continue;
      const lvl = c.level ?? 0;
      if (lvl <= myLevel) break;

      const s = c.start ? new Date(c.start) : null;
      const e = c.end ? new Date(c.end) : (c.start ? new Date(c.start) : null);

      if (s && !isNaN(s)) min = !min || s < min ? s : min;
      if (e && !isNaN(e)) max = !max || e > max ? e : max;
    }

    if (min && max) {
      t.start = toISODate(min);
      t.end   = toISODate(max);
    }
  }
}

function buildTasksFromRows(rows) {
  const newTasks = [];

  const cols = detectColumns(rows[0] || {});
  const { idKey, nameKey, startKey, finishKey, durKey } = cols;

  if (!nameKey || (!startKey && !finishKey)) {
    throw new Error("Could not detect required columns. Need at least: Activity Name, Start, Finish.");
  }

  // Build legend mapping from file (for auto mode)
  buildLegendMapFromRows(rows);

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

    const legendVal = legendFieldKey ? String(r[legendFieldKey] ?? "").trim() : "";

    if (type === "header" || type === "sub-header") {
      const obj = { name: title, type, level };
      if (legendVal) obj.legendValue = legendVal;

      if (startD) obj.start = toISODate(startD);
      if (finishD) obj.end = toISODate(finishD);

      newTasks.push(obj);
      return;
    }

    if (type === "milestone") {
      const d = startD || finishD;
      if (d) {
        const obj = { name: title, type: "milestone", start: toISODate(d), level };
        if (legendVal) obj.legendValue = legendVal;
        newTasks.push(obj);
      }
      return;
    }

    if (startD && finishD) {
      const obj = {
        name: title,
        type: "activity",
        start: toISODate(startD),
        end: toISODate(finishD),
        level
      };
      if (legendVal) obj.legendValue = legendVal;
      newTasks.push(obj);
    }
  });

  // Derive group summary dates from children (so summary bars + min/max always work)
  deriveGroupSummariesInPlace(newTasks);

  // Compute project min/max from all tasks
  let minD = null;
  let maxD = null;
  for (const t of newTasks) {
    const s = t.start ? new Date(t.start) : null;
    const e = t.end ? new Date(t.end) : (t.start ? new Date(t.start) : null);
    if (s && !isNaN(s)) minD = !minD || s < minD ? s : minD;
    if (e && !isNaN(e)) maxD = !maxD || e > maxD ? e : maxD;
  }

  // Update timeline range + month anchor (B2)
  if (minD && maxD) {
    projectMonthAnchor = new Date(minD.getFullYear(), minD.getMonth(), 1); // project start month = M1

    const padStart = new Date(minD);
    padStart.setMonth(padStart.getMonth() - 2);

    const padEnd = new Date(maxD);
    padEnd.setMonth(padEnd.getMonth() + 2);

    startDate = padStart;
    endDate = padEnd > CAP_MAX ? CAP_MAX : padEnd;
  } else {
    projectMonthAnchor = null;
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

  const legendSelect = document.getElementById("legendSelect");
  if (legendSelect) {
    legendSelect.addEventListener("change", (e) => {
      legendMode = e.target.value === "keywords" ? "keywords" : "auto";
      applyHierarchyColors();
      renderAll();
    });
  }

  setupSplitter();
  setupHeaderDragZoom();

  const controlsHost = document.getElementById("controls") || (loadBtn ? loadBtn.parentElement : null) || document.body;
  buildHighLevelControls(controlsHost);

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
      savedCollapsed = null;

      try {
        const data = await f.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array", cellDates: true });

        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (!rows || rows.length === 0) throw new Error("Sheet is empty.");

        tasks = buildTasksFromRows(rows);

        // Apply colors (auto legend if available)
        applyHierarchyColors();

        // If currently in high-level mode, keep deterministic collapse
        if (viewMode === "high") applyHighLevelCollapse();

        if (status) status.textContent = `Loaded ${tasks.length} rows. Rendering…`;
        renderAll();

        // Auto scroll to project start month (B2)
        if (timelinePane && projectMonthAnchor) {
          const anchorISO = toISODate(projectMonthAnchor);
          const px = dateToXpx(anchorISO);
          timelinePane.scrollLeft = Math.max(0, px - 250);
        } else if (timelinePane) {
          timelinePane.scrollLeft = 0;
        }

        const modeTxt = viewMode === "high" ? `High-level (L${highLevelCutoff})` : "Detailed";
        if (status) status.textContent = `Done. Loaded ${tasks.length} rows. View: ${modeTxt}. Click group rows to collapse.`;
      } catch (e) {
        console.error(e);
        if (status) status.textContent = `Error loading Excel: ${e.message || e}`;
      }
    });
  }

  // initial render
  renderAll();
});