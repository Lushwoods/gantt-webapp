// Timeline bounds
const startDate = new Date("2024-07-01");
const endDate = new Date("2029-05-01");

// Your tasks data (same as you provided)
const tasks = [
  { name: "CONTRACTUAL KEY DATES", type: "header" },
  { name: "P1 CN 21432: Project Commencement", start: "2024-08-19", type: "milestone" },
  { name: "P1 CN 21470: [M18] L1 Structure for MK", start: "2026-02-19", type: "milestone" },
  { name: "P1 CN 21980: [M18] L1 Marine Deck for ICON", start: "2026-02-19", type: "milestone" },
  { name: "P1 CN 21990: [M30] DTS Viaduct & Track", start: "2027-02-19", type: "milestone" },
  { name: "Final Project Completion", start: "2029-02-19", type: "milestone" },

  { name: "FOUNDATION: PILING WORKS", type: "sub-header" },
  { name: "P1 Bored Piling - Zone 1 (ICON Area)", start: "2024-08-19", end: "2025-05-30", color: "bg-blue-700" },
  { name: "P1 Bored Piling - Zone 2 (MK Area)", start: "2024-10-01", end: "2025-08-15", color: "bg-blue-600" },
  { name: "Marine Piles [D01] Waterfront", start: "2024-12-01", end: "2026-06-07", color: "bg-blue-500" },
  { name: "Bored Piles [EBO] Zone", start: "2025-02-15", end: "2026-05-20", color: "bg-blue-400" },

  { name: "ERSS & NEW BASEMENT EXCAVATION", type: "sub-header" },
  { name: "Site Clearance & Guide Wall Construction", start: "2024-09-01", end: "2024-12-15", color: "bg-amber-800" },
  { name: "Diaphragm Wall / ERSS Install", start: "2025-01-01", end: "2025-06-15", color: "bg-amber-700" },
  { name: "Excavation Stage 1 & Strut L1", start: "2025-03-01", end: "2025-05-30", color: "bg-amber-600" },
  { name: "Excavation Stage 2 & Strut L2", start: "2025-06-01", end: "2025-08-15", color: "bg-amber-500" },
  { name: "Final Formation Level & Blinding", start: "2025-08-16", end: "2025-09-30", color: "bg-amber-400" },

  { name: "BASEMENT STRUCTURAL SUB-STRUCTURE", type: "sub-header" },
  { name: "Raft Foundation & Water-proofing", start: "2025-09-01", end: "2025-11-15", color: "bg-orange-700" },
  { name: "B2 Vertical Elements (Walls/Cols)", start: "2025-10-15", end: "2026-01-10", color: "bg-orange-600" },
  { name: "B1 Structural Slab (Suspended)", start: "2025-11-20", end: "2026-02-15", color: "bg-orange-500" },

  { name: "MK PODIUM & ICON TOWER (SUPERSTRUCTURE)", type: "sub-header" },
  { name: "L1 Transfer Plate & Deck Support", start: "2026-02-19", end: "2026-04-30", color: "bg-emerald-700" },
  { name: "Podium (MK) L2 - L5 Structure", start: "2026-04-01", end: "2026-09-15", color: "bg-emerald-600" },
  { name: "ICON Tower Typical (L6 - L12)", start: "2026-08-01", end: "2027-04-30", color: "bg-emerald-500" },
  { name: "Tower Crown & Mechanical Floors", start: "2027-05-01", end: "2027-09-15", color: "bg-emerald-400" },

  { name: "DTS VIADUCT & TRACK WORKS", type: "sub-header" },
  { name: "DTS Foundation & Substructure", start: "2025-11-01", end: "2026-07-30", color: "bg-pink-700" },
  { name: "DTS Piers & Cross-head Casting", start: "2026-06-01", end: "2026-10-15", color: "bg-pink-600" },
  { name: "DTS Gantry Launching / Segmental", start: "2026-09-15", end: "2027-01-15", color: "bg-pink-500" },
  { name: "Track Installation & Power Rail", start: "2026-12-01", end: "2027-02-19", color: "bg-pink-400" },

  { name: "BUILDING WORKS & ARCHITECTURAL", type: "sub-header" },
  { name: "Building Envelope (Glass Facade)", start: "2027-01-01", end: "2028-02-28", color: "bg-indigo-600" },
  { name: "Internal Architectural Finishes", start: "2027-05-01", end: "2028-09-30", color: "bg-indigo-500" },
  { name: "MEP Testing & Commissioning", start: "2028-06-01", end: "2029-01-15", color: "bg-indigo-400" },

  { name: "COMPLIANCE & STATUTORY", type: "sub-header" },
  { name: "Piling As-Built Submissions [PH1]", start: "2025-12-15", end: "2026-03-30", color: "bg-purple-600" },
  { name: "P1 CN 31220: Structural As-Built [PH1]", start: "2026-02-13", end: "2026-04-28", color: "bg-purple-500" },
  { name: "P1 CN 31240: Structural As-Built [PH2]", start: "2027-12-17", end: "2028-02-29", color: "bg-purple-500" },
  { name: "Final Statutory Inspection & TOP", start: "2028-11-01", end: "2029-02-19", color: "bg-slate-400" },
];

function getPosition(dateStr) {
  const date = new Date(dateStr);
  const diff = date - startDate;
  const totalDiff = endDate - startDate;
  return (diff / totalDiff) * 100;
}

function buildTimelineHeader() {
  const header = document.getElementById("timelineHeader");
  if (!header) return;

  header.innerHTML = "";

  const months = ["Jan", "Jul"];
  const years = [2024, 2025, 2026, 2027, 2028, 2029];

  years.forEach((year) => {
    months.forEach((month) => {
      const div = document.createElement("div");
      div.className = "timeline-cell";
      div.innerText = `${month} ${year}`;
      header.appendChild(div);
    });
  });
}

function buildGanttRows() {
  const container = document.getElementById("ganttRows");
  if (!container) return;

  container.innerHTML = "";

  tasks.forEach((task) => {
    // Headers / sub-headers
    if (task.type === "header" || task.type === "sub-header") {
      const label = document.createElement("div");
      label.className = `label-col ${task.type === "header" ? "section-header" : "sub-header"} border-b py-2`;
      label.innerText = task.name;
      container.appendChild(label);

      const barCont = document.createElement("div");
      barCont.className = `bar-container ${task.type === "header" ? "section-header" : "sub-header"} border-b`;
      container.appendChild(barCont);
      return;
    }

    // Normal label
    const label = document.createElement("div");
    label.className = "label-col border-b";
    label.innerText = task.name;
    container.appendChild(label);

    // Bar area
    const barCont = document.createElement("div");
    barCont.className = "bar-container";

    if (task.type === "milestone") {
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
    } else {
      const startPos = getPosition(task.start);
      const endPos = getPosition(task.end);
      const width = Math.max(0.5, endPos - startPos);

      const bar = document.createElement("div");
      bar.className = `gantt-bar ${task.color || "bg-slate-500"}`;
      bar.style.left = `${startPos}%`;
      bar.style.width = `${width}%`;

      const days = Math.round((new Date(task.end) - new Date(task.start)) / (1000 * 60 * 60 * 24));
      bar.innerText = width > 2 ? `${days}d` : "";
      bar.title = `${task.name}: ${days} days`;

      barCont.appendChild(bar);
    }

    container.appendChild(barCont);
  });
}

// Run
document.addEventListener("DOMContentLoaded", () => {
  buildTimelineHeader();
  buildGanttRows();
});
