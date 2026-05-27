const validGradePoints = {
  O: 10,
  E: 9,
  A: 8,
  B: 7,
  C: 6,
  D: 5,
  R: 10,
  F: 2,
  M: 0,
  S: 0,
};

let workbookData = [];
let currentReportData = null;
let isReportGenerated = false;
let currentZoomLevel = 1.0;

const GOOGLE_SCRIPT_URL =
  typeof ENV !== "undefined" ? ENV.GOOGLE_SCRIPT_URL : "";

/* PREVENT BROWSER ZOOM */
document.addEventListener(
  "wheel",
  function (e) {
    if (e.ctrlKey) e.preventDefault();
  },
  { passive: false },
);
document.addEventListener(
  "keydown",
  function (e) {
    if (
      e.ctrlKey &&
      (e.key === "+" || e.key === "-" || e.key === "0" || e.key === "=")
    )
      e.preventDefault();
  },
  { passive: false },
);

/* CONFETTI LOGIC */
function fireConfetti() {
  if (!window.confetti) {
    const script = document.createElement("script");
    script.src =
      "https://cdn.jsdelivr.net/npm/canvas-confetti@1.6.0/dist/confetti.browser.min.js";
    script.onload = () => doConfettiBlast();
    document.head.appendChild(script);
  } else {
    doConfettiBlast();
  }
}
function doConfettiBlast() {
  const duration = 3 * 1000;
  const end = Date.now() + duration;
  (function frame() {
    confetti({
      particleCount: 6,
      angle: 60,
      spread: 55,
      origin: { x: 0, y: 1 },
      colors: ["#407BFF", "#22c55e", "#ffcf40", "#ef4444", "#a855f7"],
      startVelocity: 45,
    });
    confetti({
      particleCount: 6,
      angle: 120,
      spread: 55,
      origin: { x: 1, y: 1 },
      colors: ["#407BFF", "#22c55e", "#ffcf40", "#ef4444", "#a855f7"],
      startVelocity: 45,
    });
    if (Date.now() < end) requestAnimationFrame(frame);
  })();
}

/* NAVBAR SCROLL */
window.addEventListener("scroll", () => {
  const nav = document.getElementById("main-navbar");
  if (window.scrollY > 20) nav.classList.add("scrolled");
  else nav.classList.remove("scrolled");
});

/* TAB SCROLL LOGIC */
function scrollTabs(amount) {
  const container = document.getElementById("main-tab-container");
  if (container) container.scrollBy({ left: amount, behavior: "smooth" });
}

window.requestCloseInternalModal = function () {
  const confirmModal = document.getElementById("custom-confirm-modal");
  if (confirmModal) {
    confirmModal.style.display = "flex";
    confirmModal.style.opacity = "1";
    confirmModal.style.visibility = "visible";
    confirmModal.style.pointerEvents = "auto";
    confirmModal.classList.add("open");
  }
};
window.cancelCloseModal = function () {
  const confirmModal = document.getElementById("custom-confirm-modal");
  if (confirmModal) {
    confirmModal.style.display = "none";
    confirmModal.style.opacity = "0";
    confirmModal.style.visibility = "hidden";
    confirmModal.classList.remove("open");
  }
};
window.executeCloseInternalModal = function () {
  const confirmModal = document.getElementById("custom-confirm-modal");
  const resultModal = document.getElementById("internal-result-modal");
  if (confirmModal) {
    confirmModal.style.display = "none";
    confirmModal.classList.remove("open");
  }
  if (resultModal) resultModal.style.display = "none";
  document.body.style.overflow = "";
};

/* RESET UI */
function resetUI() {
  document.body.style.overflow = "";
  const semSelect = document.getElementById("semester-number");
  if (semSelect) {
    semSelect.value = "";
    semSelect.disabled = false;
  }
  const semLockIcon = document.getElementById("sem-lock-icon");
  if (semLockIcon) semLockIcon.style.display = "none";
  const regNoInput = document.getElementById("regno-input");
  if (regNoInput) regNoInput.value = "";
  const fileInput = document.getElementById("excel-file");
  if (fileInput) {
    fileInput.value = "";
    fileInput.disabled = false;
  }
  const fileNameDisplay = document.getElementById("file-name-display");
  if (fileNameDisplay) {
    fileNameDisplay.innerText = "Drag & drop or click to browse";
    fileNameDisplay.style.color = "";
  }
  const fileContentUI = document.getElementById("file-content-ui");
  if (fileContentUI) fileContentUI.classList.remove("locked-ui");
  const fileLockIcon = document.getElementById("file-lock-icon");
  if (fileLockIcon) fileLockIcon.style.display = "none";
  const internalFileInput = document.getElementById("excel-file-internal");
  if (internalFileInput) {
    internalFileInput.value = "";
    document.getElementById("file-name-display-internal").innerText =
      "Drag & drop or click to browse";
  }
  const regNoInternal = document.getElementById("regno-input-internal");
  if (regNoInternal) regNoInternal.value = "";
  document.getElementById("internal-result-modal").style.display = "none";
  document.getElementById("custom-confirm-modal").style.display = "none";
  workbookData = [];
  const reportOutput = document.getElementById("report-output");
  if (reportOutput) reportOutput.innerHTML = "";
  const downloadActions = document.getElementById("download-actions");
  if (downloadActions) downloadActions.style.display = "none";
  const calcBtn = document.getElementById("calculate-btn");
  if (calcBtn) {
    calcBtn.disabled = false;
    calcBtn.innerHTML = "Generate Report";
    calcBtn.style.cursor = "pointer";
  }
  isReportGenerated = false;
  currentReportData = null;
  currentZoomLevel = 1.0;
  document
    .querySelectorAll(".error-msg")
    .forEach((el) => (el.style.display = "none"));

  const oldTable = document.getElementById("leaderboard-standalone-section");
  if (oldTable) oldTable.remove();
}
window.addEventListener("load", resetUI);

/* CUSTOM ZOOM & PAN */
function applySheetZoom() {
  const sheet = document.getElementById("grade-sheet");
  const container = document.getElementById("grade-sheet-target");
  if (!sheet || !container) return;
  const rawHeight = sheet.offsetHeight;
  sheet.style.transform = `scale(${currentZoomLevel})`;
  container.style.width = `${794 * currentZoomLevel}px`;
  container.style.height = `${rawHeight * currentZoomLevel}px`;
  const zoomLabel = document.getElementById("zoom-level-label");
  if (zoomLabel) zoomLabel.innerText = Math.round(currentZoomLevel * 100) + "%";
}
function changeZoom(step) {
  currentZoomLevel += step;
  if (currentZoomLevel < 0.2) currentZoomLevel = 0.2;
  if (currentZoomLevel > 3.0) currentZoomLevel = 3.0;
  applySheetZoom();
}
function fitToScreen() {
  const wrapper = document.getElementById("report-scroll-wrapper");
  if (!wrapper) return;
  const availableWidth = wrapper.clientWidth - 40;
  if (availableWidth > 0 && availableWidth < 794)
    currentZoomLevel = availableWidth / 794;
  else currentZoomLevel = 1.0;
  applySheetZoom();
}
window.addEventListener("resize", fitToScreen);

function initDragToScroll() {
  const slider = document.getElementById("report-scroll-wrapper");
  if (!slider) return;
  let isDown = false,
    startX,
    startY,
    scrollLeft,
    scrollTop;
  slider.addEventListener("mousedown", (e) => {
    isDown = true;
    slider.style.cursor = "grabbing";
    startX = e.pageX - slider.offsetLeft;
    startY = e.pageY - slider.offsetTop;
    scrollLeft = slider.scrollLeft;
    scrollTop = slider.scrollTop;
  });
  slider.addEventListener("mouseleave", () => {
    isDown = false;
    slider.style.cursor = "grab";
  });
  slider.addEventListener("mouseup", () => {
    isDown = false;
    slider.style.cursor = "grab";
  });
  slider.addEventListener("mousemove", (e) => {
    if (!isDown) return;
    e.preventDefault();
    const x = e.pageX - slider.offsetLeft;
    const y = e.pageY - slider.offsetTop;
    slider.scrollLeft = scrollLeft - (x - startX) * 1.5;
    slider.scrollTop = scrollTop - (y - startY) * 1.5;
  });
}

/* POPUPS & MENUS */
function customAlert(msg) {
  document.getElementById("alert-msg").innerText = msg;
  document.getElementById("custom-alert").classList.add("open");
}
function closeCustomAlert() {
  document.getElementById("custom-alert").classList.remove("open");
}
function openModal() {
  document.getElementById("formula-modal").classList.add("open");
}
function closeModal() {
  document.getElementById("formula-modal").classList.remove("open");
}
function openExcelModal() {
  document.getElementById("excel-modal").classList.add("open");
}
function closeExcelModal() {
  document.getElementById("excel-modal").classList.remove("open");
}
function toggleMenu() {
  const nav = document.getElementById("nav-menu");
  const icon = document.getElementById("menu-icon");
  nav.classList.toggle("active");
  if (nav.classList.contains("active"))
    icon.classList.replace("ri-menu-line", "ri-close-line");
  else icon.classList.replace("ri-close-line", "ri-menu-line");
}
function closeMenu() {
  document.getElementById("nav-menu").classList.remove("active");
  document
    .getElementById("menu-icon")
    .classList.replace("ri-close-line", "ri-menu-line");
}

/* SWITCH TAB */
function switchTab(tabId) {
  const sections = ["sgpa-section", "cgpa-section", "internal-section"];
  sections.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.style.display = "none";
  });
  const tabs = ["tab-sgpa", "tab-cgpa", "tab-internal"];
  tabs.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("active");
  });
  const targetSection = document.getElementById(tabId + "-section");
  if (targetSection) targetSection.style.display = "block";
  const targetTab = document.getElementById("tab-" + tabId);
  if (targetTab) targetTab.classList.add("active");
  const navLinks = ["nav-sgpa", "nav-cgpa", "nav-internal"];
  navLinks.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("active");
  });
  const activeNavLink = document.getElementById("nav-" + tabId);
  if (activeNavLink) activeNavLink.classList.add("active");
  const tabWrapper = document.querySelector(".tab-container");
  if (tabWrapper)
    tabWrapper.scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ============================================================
   TOP PERFORMERS — CLASS RANKINGS
   ============================================================ */

const DEVELOPER_REG = "230301120327";
const EXCLUDED_STUDENT_REG = "230301120504";

(function injectTopPerformerStyles() {
  if (document.getElementById("tp-styles")) return;
  const style = document.createElement("style");
  style.id = "tp-styles";
  style.textContent = `
    @keyframes tp-expand { from { opacity: 0; } to { opacity: 1; } }
    
    /* Animation for Developer Row Highlight */
    @keyframes devHighlight {
      0% { background-color: rgba(168, 85, 247, 0.05); }
      50% { background-color: rgba(168, 85, 247, 0.2); }
      100% { background-color: rgba(168, 85, 247, 0.05); }
    }

    .tp-master-container, .tp-master-container * {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
    }
    .tp-standalone-section {
      width: 100%; max-width: 860px; margin: 40px auto; padding: 0 15px; box-sizing: border-box;
    }
    .tp-wrapper {
      background: #0c0c0e; border-radius: 14px; border: 1px solid #1c1c20; overflow: hidden; box-shadow: 0 4px 32px rgba(0,0,0,0.6);
    }
    .tp-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 20px 24px 16px; border-bottom: 1px solid #1c1c20; background: linear-gradient(180deg, #111114 0%, #0c0c0e 100%);
    }
    .tp-title-group { display: flex; align-items: center; gap: 14px; }
    .tp-icon-wrap { width: 69px; height: 38px; background: #18181b; border-radius: 10px; border: 1px solid #27272a; display: flex; align-items: center; justify-content: center; }
    .tp-title { font-size: 15.5px; font-weight: 700; color: #f4f4f5; margin: 0; line-height: 1; }
    .tp-subtitle { font-size: 12px; color: #71717a; margin: 4px 0 0; }
    .tp-count-badge { background: #111114; border: 1px solid #27272a; border-radius: 20px; padding: 4px 22px; font-size: 12px; color: #71717a; font-weight: 600; }

    /* Strict Alignment Fixes - Mobile Responsive */
    .tp-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table.tp-table { width: 100%; min-width: 550px; border-collapse: collapse; table-layout: fixed; }
    table.tp-table colgroup col:nth-child(1) { width: 80px; }  /* Rank */
    table.tp-table colgroup col:nth-child(2) { width: auto; }  /* Name */
    table.tp-table colgroup col:nth-child(3) { width: 140px; } /* Reg No */
    table.tp-table colgroup col:nth-child(4) { width: 100px; } /* SGPA */

    table.tp-table thead th {
      padding: 12px 14px; font-size: 11px; font-weight: 700;
      letter-spacing: 0.9px; text-transform: uppercase; color: #52525b;
      border-bottom: 1px solid #1c1c20; background: #0e0e11;
      white-space: nowrap; /* Fixes overlapping/stacking headers on mobile */
    }
    
    /* Guaranteed straight lines */
    table.tp-table th, table.tp-table td { vertical-align: middle; padding: 14px 10px; }
    table.tp-table th.th-rank, table.tp-table td.tp-td-rank { text-align: center; }
    table.tp-table th.th-name, table.tp-table td.tp-td-name { text-align: left; }
    table.tp-table th.th-reg, table.tp-table td.tp-td-reg { text-align: center; }
    table.tp-table th.th-score, table.tp-table td.tp-td-score { text-align: center; }

    .tp-rank-container { display: flex; align-items: center; justify-content: center; gap: 8px; width: 100%; }
    .tp-rank-number { width: 18px; text-align: right; font-size: 14px; font-weight: 800; display: inline-block; }
    .tp-rank-number.gold { color: #facc15; }
    .tp-rank-number.silver { color: #a1a1aa; }
    .tp-rank-number.bronze { color: #d97706; }
    .tp-rank-number.normal { color: #71717a; }
    .tp-rank-icon { width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; border-radius: 8px; flex-shrink: 0; }
    .tp-rank-gold { background:rgba(234,179,8,0.07); border:1px solid rgba(234,179,8,0.22); }
    .tp-rank-silver { background:rgba(161,161,170,0.06); border:1px solid rgba(161,161,170,0.18); }
    .tp-rank-bronze { background:rgba(180,83,9,0.07); border:1px solid rgba(180,83,9,0.2); }
    .tp-rank-placeholder { width: 28px; display: inline-block; }

    /* Rows */
    .tp-row-gold { background: linear-gradient(90deg, rgba(20,16,4,0.85) 0%, rgba(12,12,14,0.85) 60%); border-bottom: 1px solid rgba(234,179,8,0.08); position: relative; }
    .tp-row-silver { background:rgba(12,12,14,0.8); border-bottom:1px solid rgba(161,161,170,0.06); position:relative; }
    .tp-row-bronze { background:rgba(12,12,14,0.8); border-bottom:1px solid rgba(180,83,9,0.07); position:relative; }
    .tp-row-normal { border-bottom: 1px solid #111114; }
    
    /* Developer Row Styles */
    .tp-row-dev {
      animation: devHighlight 2s infinite ease-in-out !important;
      border-top: 1px solid rgba(168, 85, 247, 0.3) !important;
      border-bottom: 1px solid rgba(168, 85, 247, 0.3) !important;
    }
    .tp-name-dev { color: #e9d5ff !important; font-weight: 800 !important; text-shadow: 0 0 8px rgba(168,85,247,0.4); }
    .tp-score-dev { color: #d8b4fe !important; font-weight: 800 !important; }
    .dev-badge {
      display: inline-block; background: linear-gradient(90deg, #9333ea, #a855f7);
      color: #fff; font-size: 9px; font-weight: 800; padding: 2px 5px;
      border-radius: 4px; margin-left: 6px; vertical-align: middle;
      text-transform: uppercase; letter-spacing: 0.5px;
    }

    .tp-name { font-size:13.5px; font-weight:600; color:#d4d4d8; display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .tp-name-gold { color:#fef08a; font-weight: 700; }
    .tp-name-silver { color:#e4e4e7; font-weight: 700; }
    .tp-name-bronze { color:#fed7aa; font-weight: 700; }
    .tp-reg { font-size:12px; color:#71717a; font-weight: 500;}

    .tp-score { font-size:14px; font-weight:700; }
    .tp-score-gold { color:#eab308; }
    .tp-score-silver { color:#a1a1aa; }
    .tp-score-bronze { color:#c2670a; }
    .tp-score-normal { color:#71717a; }

    .tp-extra-rows { display:none; }
    .tp-extra-rows.visible { display:table-row-group; animation:tp-expand 0.3s ease; }

    .tp-footer { padding:14px 20px; border-top:1px solid #1c1c20; text-align:center; background:#0a0a0c; }
    .tp-toggle-btn {
      display:inline-flex; align-items:center; gap:7px; padding:8px 18px; background:#111114;
      border:1px solid #27272a; border-radius:7px; color:#71717a; font-size:12.5px; font-weight:600;
      cursor:pointer; transition:all 0.18s;
    }
    .tp-toggle-btn:hover { background:#18181b; border-color:#3f3f46; color:#a1a1aa; }
    .tp-toggle-btn svg { transition:transform 0.3s; flex-shrink:0; }
    .tp-toggle-btn.expanded svg.tp-chevron { transform:rotate(180deg); }
    .tp-no-more { font-size:12px; color:#52525b; font-weight:500; margin: 0; }

    .tp-section-divider { display:flex; align-items:center; gap:14px; max-width:860px; margin: 0 auto 20px auto; padding:0 2px; }
    .tp-section-divider-line { flex:1; height:1px; background:#1c1c20; }
    .tp-section-divider-label { font-size:11px; font-weight:700; letter-spacing:1.3px; text-transform:uppercase; color:#52525b; }
  `;
  document.head.appendChild(style);
})();

/* SVG icons */
const SVG_TROPHY_TP = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M12 15C8.686 15 6 12.314 6 9V3H18V9C18 12.314 15.314 15 12 15Z" stroke="#71717a" stroke-width="1.7" stroke-linejoin="round"/><path d="M6 5H3C3 5 3 10 6 10" stroke="#71717a" stroke-width="1.7" stroke-linecap="round"/><path d="M18 5H21C21 5 21 10 18 10" stroke="#71717a" stroke-width="1.7" stroke-linecap="round"/><path d="M12 15V19" stroke="#71717a" stroke-width="1.7" stroke-linecap="round"/><path d="M8 21H16" stroke="#71717a" stroke-width="1.7" stroke-linecap="round"/></svg>`;
const SVG_USERS_TP = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M17 20C17 18.343 14.761 17 12 17C9.239 17 7 18.343 7 20" stroke="#71717a" stroke-width="1.8" stroke-linecap="round"/><circle cx="12" cy="11" r="3" stroke="#71717a" stroke-width="1.8"/><path d="M21 20C21 18.667 19.667 17.667 18 17.333" stroke="#71717a" stroke-width="1.8" stroke-linecap="round"/><circle cx="18" cy="11.5" r="2.5" stroke="#71717a" stroke-width="1.8"/><path d="M3 20C3 18.667 4.333 17.667 6 17.333" stroke="#71717a" stroke-width="1.8" stroke-linecap="round"/><circle cx="6" cy="11.5" r="2.5" stroke="#71717a" stroke-width="1.8"/></svg>`;
const SVG_CHEVRON_TP = `<svg class="tp-chevron" width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M6 9L12 15L18 9" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const SVG_MEDAL_GOLD_TP = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="14" r="6.5" stroke="#eab308" stroke-width="1.7"/><path d="M8.5 3.5L12 7L15.5 3.5" stroke="#eab308" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M12 11.5V14.5" stroke="#eab308" stroke-width="1.7" stroke-linecap="round"/><path d="M10.5 13H13.5" stroke="#eab308" stroke-width="1.7" stroke-linecap="round"/></svg>`;
const SVG_MEDAL_SILVER_TP = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="14" r="6.5" stroke="#a1a1aa" stroke-width="1.7"/><path d="M8.5 3.5L12 7L15.5 3.5" stroke="#a1a1aa" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M10 12C10 12 10 11 12 11C14 11 14 12.5 12 13.5C10 14.5 10 16 10 16H14" stroke="#a1a1aa" stroke-width="1.5" stroke-linecap="round"/></svg>`;
const SVG_MEDAL_BRONZE_TP = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="14" r="6.5" stroke="#c2670a" stroke-width="1.7"/><path d="M8.5 3.5L12 7L15.5 3.5" stroke="#c2670a" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/><path d="M10 11H13C14.1 11 14.5 12 13.5 12.8C14.6 13.4 14.5 16 12.5 16H10" stroke="#c2670a" stroke-width="1.5" stroke-linecap="round"/><path d="M10 11V16" stroke="#c2670a" stroke-width="1.5" stroke-linecap="round"/></svg>`;

/* Row Builder */
function buildTPRow(student, rank) {
  let rowClass, nameClass, scoreClass, rankContent;
  const isDev = student.regNo === DEVELOPER_REG;

  if (rank === 1) {
    rowClass = "tp-row-gold";
    nameClass = "tp-name-gold";
    scoreClass = "tp-score-gold";
    rankContent = `<div class="tp-rank-container"><span class="tp-rank-number gold">1</span><span class="tp-rank-icon tp-rank-gold">${SVG_MEDAL_GOLD_TP}</span></div>`;
  } else if (rank === 2) {
    rowClass = "tp-row-silver";
    nameClass = "tp-name-silver";
    scoreClass = "tp-score-silver";
    rankContent = `<div class="tp-rank-container"><span class="tp-rank-number silver">2</span><span class="tp-rank-icon tp-rank-silver">${SVG_MEDAL_SILVER_TP}</span></div>`;
  } else if (rank === 3) {
    rowClass = "tp-row-bronze";
    nameClass = "tp-name-bronze";
    scoreClass = "tp-score-bronze";
    rankContent = `<div class="tp-rank-container"><span class="tp-rank-number bronze">3</span><span class="tp-rank-icon tp-rank-bronze">${SVG_MEDAL_BRONZE_TP}</span></div>`;
  } else {
    rowClass = "tp-row-normal";
    nameClass = "";
    scoreClass = "tp-score-normal";
    rankContent = `<div class="tp-rank-container"><span class="tp-rank-number normal">${rank}</span><span class="tp-rank-placeholder"></span></div>`;
  }

  if (isDev) {
    rowClass += " tp-row-dev";
    nameClass = "tp-name-dev";
    scoreClass = "tp-score-dev";
  }

  const displayName = `${student.name}${isDev ? ' <span class="dev-badge">DEVELOPER</span>' : ""}`;

  return `
    <tr class="${rowClass}">
      <td class="tp-td-rank">${rankContent}</td>
      <td class="tp-td-name"><span class="tp-name ${nameClass}">${displayName}</span></td>
      <td class="tp-td-reg"><span class="tp-reg">${student.regNo}</span></td>
      <td class="tp-td-score tp-score ${scoreClass}">${student.score}</td>
    </tr>`;
}

/* Render the list strictly for top 100 with dynamic expand */
function renderTop10UI(title, rawData) {
  // Clear old table if it exists
  const oldTable = document.getElementById("leaderboard-standalone-section");
  if (oldTable) oldTable.remove();

  /* Only run logic if Report is actually generated */
  if (!isReportGenerated) return;

  let filteredData = [...rawData].filter(
    (s) => s.regNo !== EXCLUDED_STUDENT_REG,
  );
  let sorted = filteredData.sort((a, b) => {
    const diff = parseFloat(b.score) - parseFloat(a.score);
    return diff !== 0 ? diff : a.name.localeCompare(b.name);
  });

  // Cut strictly to top 100
  const top100 = sorted.slice(0, 100);

  // Store Global State for toggling
  window.tpTotalCount = top100.length;
  window.tpVisibleStage = 1;

  let html1to10 = "";
  let html11to50 = "";
  let html51to100 = "";

  top100.forEach((student, i) => {
    const rank = i + 1;
    const rowHTML = buildTPRow(student, rank);
    if (rank <= 10) html1to10 += rowHTML;
    else if (rank <= 50) html11to50 += rowHTML;
    else html51to100 += rowHTML;
  });

  const hasExtra = top100.length > 10;
  const extraCount = top100.length - 10;

  const footerHTML = hasExtra
    ? `<div class="tp-footer"><button class="tp-toggle-btn" id="tp-toggle-btn" onclick="toggleTPExtra()">${SVG_USERS_TP}<span id="tp-btn-label">View rest ${extraCount}</span>${SVG_CHEVRON_TP}</button></div>`
    : `<div class="tp-footer"><p class="tp-no-more">All ${top100.length} performers shown</p></div>`;

  let leaderboardSection = document.createElement("section");
  const semInput = document.getElementById("semester-number");
  const sem = semInput.value;
  leaderboardSection.id = "leaderboard-standalone-section";
  leaderboardSection.className = "tp-master-container tp-standalone-section";
  leaderboardSection.innerHTML = `
    <div class="tp-section-divider">
      <div class="tp-section-divider-line"></div>
      <span class="tp-section-divider-label">Class Rankings</span>
      <div class="tp-section-divider-line"></div>
    </div>
    <div class="tp-wrapper">
      <div class="tp-header">
        <div class="tp-title-group">
          <div class="tp-icon-wrap">${SVG_TROPHY_TP}</div>
          <div>
            <p class="tp-title">Top Performers</p>
            <p class="tp-subtitle">Ranked by Semester GPA &mdash; of Semester ${sem}</p>
          </div>
        </div>
        <div class="tp-count-badge">${SVG_USERS_TP}&nbsp;<em>${top100.length}</em> ranked</div>
      </div>
      <div class="tp-table-wrap">
        <table class="tp-table">
          <colgroup>
            <col>
            <col>
            <col>
            <col>
          </colgroup>
          <thead>
            <tr>
              <th class="th-rank">Rank</th>
              <th class="th-name">Student Name</th>
              <th class="th-reg">Regd. No</th>
              <th class="th-score">SGPA</th>
            </tr>
          </thead>
          <tbody id="tp-body-1">${html1to10}</tbody>
          <tbody class="tp-extra-rows" id="tp-body-2">${html11to50}</tbody>
          <tbody class="tp-extra-rows" id="tp-body-3">${html51to100}</tbody>
        </table>
      </div>
      ${footerHTML}
    </div>
  `;

  // Explicitly Place strictly above Premium Features Section
  let premiumFeaturesNode =
    document.getElementById("premium-features") ||
    document.querySelector(".premium-features");

  // Advanced text-based lookup if exact ID/Class is missing in HTML
  if (!premiumFeaturesNode) {
    const headings = document.querySelectorAll("h1, h2, h3, h4, h5, h6");
    for (let el of headings) {
      if (el.textContent && el.textContent.includes("Premium Features")) {
        premiumFeaturesNode = el.closest("section") || el.parentNode;
        break;
      }
    }
  }

  if (premiumFeaturesNode && premiumFeaturesNode.parentNode) {
    premiumFeaturesNode.parentNode.insertBefore(
      leaderboardSection,
      premiumFeaturesNode,
    );
  } else {
    // Fallback: Place at the bottom if Premium Features cannot be found
    document.body.appendChild(leaderboardSection);
  }
}

window.toggleTPExtra = function () {
  const b2 = document.getElementById("tp-body-2");
  const b3 = document.getElementById("tp-body-3");
  const btn = document.getElementById("tp-toggle-btn");
  const label = document.getElementById("tp-btn-label");
  if (!btn) return;

  const total = window.tpTotalCount || 100;

  if (window.tpVisibleStage === 1) {
    // Expand to 50
    if (b2) b2.classList.add("visible");
    window.tpVisibleStage = 2;
    let remain = total - 50;
    if (remain > 0) {
      label.textContent = `View rest ${remain}`;
      btn.classList.remove("expanded");
    } else {
      label.textContent = "Show less";
      btn.classList.add("expanded");
    }
  } else if (window.tpVisibleStage === 2 && total > 50) {
    // Expand up to 100
    if (b3) b3.classList.add("visible");
    window.tpVisibleStage = 3;
    label.textContent = "Show less";
    btn.classList.add("expanded");
  } else {
    // Collapse back to 10
    if (b2) b2.classList.remove("visible");
    if (b3) b3.classList.remove("visible");
    window.tpVisibleStage = 1;
    label.textContent = `View rest ${total - 10}`;
    btn.classList.remove("expanded");
  }
};

/* TOP SGPA CALCULATOR PROCESSOR */
function generateTop10SGPAList() {
  if (!workbookData || workbookData.length === 0) return;
  const students = {};
  workbookData.forEach((row) => {
    const regNo = String(row["Reg_No"] || "").trim();
    if (!regNo || regNo === "undefined") return;
    const name = row["Name"] || "Unknown";
    const grade = String(row["Grade"]).trim().toUpperCase();
    const credit = parseCredit(row["Credits"]);
    const points =
      validGradePoints[grade] !== undefined ? validGradePoints[grade] : 0;
    if (!students[regNo])
      students[regNo] = { name, totalCredits: 0, totalWeightedPoints: 0 };
    students[regNo].totalCredits += credit;
    students[regNo].totalWeightedPoints += points * credit;
  });
  const results = [];
  for (const regNo in students) {
    const s = students[regNo];
    if (s.totalCredits > 0)
      results.push({
        regNo,
        name: s.name,
        score: (s.totalWeightedPoints / s.totalCredits).toFixed(2),
      });
  }
  renderTop10UI("Top Performers", results);
}

/* EXCEL PARSING */
document.getElementById("excel-file").addEventListener("change", function (e) {
  const fileName = e.target.files[0]
    ? e.target.files[0].name
    : "Click to upload .xlsx file";
  document.getElementById("file-name-display").innerText = fileName;
  if (e.target.files[0]) {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        workbookData = [];
        wb.SheetNames.forEach((sheetName) => {
          const sheet = wb.Sheets[sheetName];
          const rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          const formatted = rawData.map((row) => {
            let newRow = {};
            for (let key in row) newRow[key.trim()] = row[key];
            return newRow;
          });
          workbookData = workbookData.concat(formatted);
        });
        /* Do NOT call generateTop10SGPAList() here. Only call it when report is generated */
      } catch (error) {
        customAlert(
          "Failed to read the Excel file. Please ensure it's a valid .xlsx file.",
        );
      }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
  }
});

document.getElementById("regno-input").addEventListener("input", function () {
  if (isReportGenerated) {
    const btn = document.getElementById("calculate-btn");
    btn.disabled = false;
    btn.innerHTML = "Generate Report";
    btn.style.cursor = "pointer";
  }
});

/* GENERATE REPORT (SGPA) */
document.getElementById("calculate-btn").addEventListener("click", function () {
  document
    .querySelectorAll(".error-msg")
    .forEach((el) => (el.style.display = "none"));
  const regNoInput = document.getElementById("regno-input");
  const semInput = document.getElementById("semester-number");
  const regNo = regNoInput.value.trim();
  const sem = semInput.value;
  const reportDiv = document.getElementById("report-output");
  let hasError = false;

  if (workbookData.length === 0) {
    customAlert("Please upload your Excel result file first.");
    return;
  }
  if (!sem) {
    document.getElementById("sem-error").innerText = "Required";
    document.getElementById("sem-error").style.display = "block";
    hasError = true;
  }
  if (!regNo || regNo.length < 5) {
    document.getElementById("reg-error").innerText = "Required";
    document.getElementById("reg-error").style.display = "block";
    hasError = true;
  }
  if (hasError) return;

  const studentRows = workbookData.filter(
    (row) => String(row["Reg_No"]).trim() === regNo,
  );
  if (studentRows.length === 0) {
    document.getElementById("reg-error").innerText = "Not found in excel.";
    document.getElementById("reg-error").style.display = "block";
    return;
  }

  let totalWeightedPoints = 0,
    totalCredits = 0,
    creditsCleared = 0;
  const studentName = studentRows[0]["Name"] || "Unknown Student";
  let batch = regNo.length >= 2 ? "20" + regNo.substring(0, 2) : "N/A";
  let subjectsArray = [];
  let actualBacklogs = [];

  const rowsHTML = studentRows
    .map((row, i) => {
      const grade = String(row["Grade"]).trim().toUpperCase();
      const credit = parseCredit(row["Credits"]);
      const subject = row["Subject_Name"] || "Unknown";
      const type = row["Type"] || "PP";
      subjectsArray.push({ name: subject, grade });
      let points =
        validGradePoints[grade] !== undefined ? validGradePoints[grade] : 0;
      totalCredits += credit;
      if (["F", "M", "S"].includes(grade)) {
        if (!actualBacklogs.includes(subject)) actualBacklogs.push(subject);
      }
      if (!["F", "S", "M"].includes(grade)) creditsCleared += credit;
      totalWeightedPoints += points * credit;
      return `<tr><td>${i + 1}</td><td>${row["Subject_Code"] || ""}</td><td>${subject}</td><td>${type}</td><td>${credit}</td><td>${grade}</td></tr>`;
    })
    .join("");

  const sgpa =
    totalCredits > 0 ? (totalWeightedPoints / totalCredits).toFixed(2) : "0.00";
  currentReportData = {
    studentName,
    regNo,
    batch,
    sem,
    sgpa,
    creditsCleared,
    totalCredits,
    backlogs: actualBacklogs,
    subjects: subjectsArray,
  };

  const currentDate = new Date();
  const dateString = currentDate
    .toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    })
    .replace(/ /g, "-");
  const timeString = currentDate.toLocaleString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

  let bannerHTML = "";
  const hasBacklogs = actualBacklogs.length > 0;
  const isOutstanding = parseFloat(sgpa) >= 9.0 && !hasBacklogs;
  if (isOutstanding) {
    bannerHTML = `<div class="report-status-banner status-outstanding"><div class="banner-icon"><img src="https://cdn-icons-png.flaticon.com/512/3176/3176294.png" style="width: 36px; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.2));" alt="Medal"></div><div class="banner-content"><h4>Outstanding Performance! 🏆</h4><p>Incredible job! You achieved a stellar SGPA of ${sgpa}. Keep up the excellent work!</p></div></div>`;
  } else if (hasBacklogs) {
    bannerHTML = `<div class="report-status-banner status-warning"><div class="banner-icon"><i class="ri-error-warning-fill"></i></div><div class="banner-content"><h4>Action Required: Pending Subjects</h4><p>You have pending backlogs (${actualBacklogs.join(", ")}). Please prepare well and clear them in upcoming exams.</p></div></div>`;
  } else {
    bannerHTML = `<div class="report-status-banner status-clear"><div class="banner-icon"><i class="ri-verified-badge-fill"></i></div><div class="banner-content"><h4>All Clear! 🎉</h4><p>Congratulations! You have successfully cleared all subjects for this semester.</p></div></div>`;
  }

  reportDiv.innerHTML = `${bannerHTML}
    <div id="report-scroll-wrapper" class="report-scroll-wrapper">
      <div id="grade-sheet-target" class="grade-sheet-target">
        <div id="grade-sheet" class="grade-sheet">
          <div class="sheet-top-header"><div>${timeString}</div><div>GradeFlow - Streamlining your academic journey</div></div>
          <div class="sheet-logos"><img src="Assets/cutm.png" alt="Logo" class="sheet-logo-img" onerror="this.src='Assets/cutm_text.jpg'"></div>
          <div class="sheet-titles"><h1>Centurion University of Technology and Management</h1><h3>Jatni, Khurda, Odisha</h3><h2>Semester Grade Sheet</h2></div>
          <div class="student-info-grid">
            <div class="info-row"><span class="lbl">Student Regd. No</span> <span class="val">: ${regNo}</span></div>
            <div class="info-row"><span class="lbl">Student Name</span> <span class="val">: ${studentName.toUpperCase()}</span></div>
            <div class="info-row"><span class="lbl">Batch</span> <span class="val">: ${batch}</span></div>
            <div class="info-row"><span class="lbl">Semester</span> <span class="val">: Sem ${sem}</span></div>
          </div>
          <table class="result-table"><thead><tr><th>SL.NO</th><th>SUB.CODE</th><th>SUBJECT</th><th>TYPE</th><th>CREDIT</th><th>GRADE</th></tr></thead><tbody>${rowsHTML}</tbody></table>
          <div class="summary-row" style="margin-top: 80px;"><div>Total Credits : ${totalCredits}</div><div>Credits Cleared : ${creditsCleared}</div><div>SGPA : ${sgpa}</div></div>
          <div class="signature-row"><div>Date : ${dateString}</div><div>Dean, Examinations</div></div>
        </div>
      </div>
    </div>
    <div style="text-align: center; width: 100%;"><div class="inline-zoom-controls"><button onclick="changeZoom(-0.1)">-</button><span id="zoom-level-label">100%</span><button onclick="changeZoom(0.1)">+</button></div></div>`;

  document.getElementById("download-actions").style.display = "flex";
  semInput.disabled = true;
  document.getElementById("sem-lock-icon").style.display = "inline-block";
  document.getElementById("excel-file").disabled = true;
  document.getElementById("file-content-ui").classList.add("locked-ui");
  document.getElementById("file-lock-icon").style.display = "inline-block";

  const calcBtn = document.getElementById("calculate-btn");
  calcBtn.disabled = true;
  calcBtn.innerHTML =
    '<i class="ri-verified-badge-fill" style="color: #3b82f6;"></i> Report Generated';
  calcBtn.style.cursor = "not-allowed";

  // Triggers Top Performers render
  isReportGenerated = true;

  if (GOOGLE_SCRIPT_URL) {
    const formData = new URLSearchParams();
    formData.append("date", dateString);
    formData.append("time", timeString);
    formData.append("regNo", regNo);
    formData.append("name", studentName);
    fetch(GOOGLE_SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      body: formData,
    }).catch((e) => console.log("Log error"));
  }

  setTimeout(() => {
    fitToScreen();
    initDragToScroll();
    reportDiv.scrollIntoView({ behavior: "smooth", block: "start" });
    if (isOutstanding) fireConfetti();

    // Generates the rankings UI only after clicking Generate Report
    generateTop10SGPAList();
  }, 50);
});

/* WHATSAPP SHARE */
let promptCallback = null;
const waInput = document.getElementById("prompt-input");
const waSendBtn = document.getElementById("wa-send-btn");

document.getElementById("whatsapp-btn").addEventListener("click", () => {
  if (!currentReportData) return;
  waInput.value = "";
  waSendBtn.disabled = true;
  waSendBtn.classList.add("disabled-btn");
  waSendBtn.style.background = "#1f1f1f";
  waSendBtn.style.color = "#555";
  promptCallback = (num) => {
    const d = currentReportData;
    let text = `◆ Semester Results Update ◆\n\nName: ${d.studentName}\nRegd No: ${d.regNo}\nBatch: ${d.batch}\nSemester: ${d.sem}\n\n◆ Performance Summary ◆\nTotal Credits: ${d.totalCredits}\nCredits Cleared: ${d.creditsCleared}\nSGPA: ${d.sgpa}\n\n◆ Course Grades ◆\n`;
    d.subjects.forEach((sub) => {
      text += `- ${sub.name} : ${sub.grade}\n`;
    });
    if (d.backlogs.length > 0)
      text += `\n[!] Backlogs Note: ${d.backlogs.join(", ")}\n\n`;
    else text += `\n[*] All clear! Excellent performance!\n\n`;
    text += `Please click the link below for detailed subject-wise grades.\nhttps://cutm-sgpa-cgpa-calculator-pro.vercel.app`;
    window.open(
      `https://wa.me/${num}?text=${encodeURIComponent(text)}`,
      "_blank",
    );
  };
  document.getElementById("custom-prompt").classList.add("open");
});
function closeCustomPrompt() {
  document.getElementById("custom-prompt").classList.remove("open");
}
waInput.addEventListener("input", (e) => {
  const val = e.target.value.trim();
  if (val.length >= 10 && !val.includes("+")) {
    waSendBtn.disabled = false;
    waSendBtn.classList.remove("disabled-btn");
    waSendBtn.style.background = "#25D366";
    waSendBtn.style.color = "#000";
  } else {
    waSendBtn.disabled = true;
    waSendBtn.classList.add("disabled-btn");
    waSendBtn.style.background = "#1f1f1f";
    waSendBtn.style.color = "#555";
  }
});
waSendBtn.addEventListener("click", () => {
  const val = waInput.value.trim();
  if (val.length >= 10) {
    closeCustomPrompt();
    if (promptCallback) promptCallback(val);
  }
});

/* CGPA CALCULATOR */
function addCgpaRow() {
  const div = document.createElement("div");
  div.className = "cgpa-row";
  div.innerHTML = `<div class="input-with-icon"><i class="ri-hashtag"></i><input type="number" class="cgpa-sgpa" placeholder="SGPA" step="0.01"></div><div class="input-with-icon"><i class="ri-coin-line"></i><input type="number" class="cgpa-credit" placeholder="Credits" step="0.5"></div>`;
  document.getElementById("cgpa-entries").appendChild(div);
}
function calculateCGPA() {
  const sgpas = document.querySelectorAll(".cgpa-sgpa");
  const credits = document.querySelectorAll(".cgpa-credit");
  const errorSpan = document.getElementById("cgpa-error");
  errorSpan.style.display = "none";
  let num = 0,
    den = 0,
    hasError = false,
    isEmpty = true;
  sgpas.forEach((inp, i) => {
    const s = parseFloat(inp.value);
    const c = parseFloat(credits[i].value);
    if (inp.value !== "" || credits[i].value !== "") isEmpty = false;
    if (!isNaN(s) && !isNaN(c)) {
      if (s > 10 || s < 0) hasError = true;
      else {
        num += s * c;
        den += c;
      }
    }
  });
  if (isEmpty) {
    errorSpan.innerText = "Fill at least one row.";
    errorSpan.style.display = "block";
    return;
  }
  if (hasError) {
    errorSpan.innerText = "SGPA must be between 0 and 10.";
    errorSpan.style.display = "block";
    return;
  }
  if (den === 0) {
    errorSpan.innerText = "Total credits cannot be zero.";
    errorSpan.style.display = "block";
    return;
  }
  document.getElementById("cgpa-result-value").innerText = (num / den).toFixed(
    2,
  );
}
function parseCredit(val) {
  if (!val && val !== 0) return 0;
  return val
    .toString()
    .split("+")
    .reduce((a, c) => a + parseFloat(c || 0), 0);
}

/* EXPORTS */
document.getElementById("download-btn").addEventListener("click", () => {
  const sheet = document.getElementById("grade-sheet");
  sheet.style.transform = "none";
  html2canvas(sheet, { scale: 2 }).then((canvas) => {
    const pdf = new window.jspdf.jsPDF("p", "mm", "a4");
    const imgWidth = 210;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    pdf.addImage(
      canvas.toDataURL("image/png"),
      "PNG",
      0,
      0,
      imgWidth,
      imgHeight,
    );
    pdf.save("GradeSheet.pdf");
    applySheetZoom();
  });
});
document.getElementById("download-photo-btn").addEventListener("click", () => {
  const sheet = document.getElementById("grade-sheet");
  sheet.style.transform = "none";
  html2canvas(sheet, { scale: 3 }).then((canvas) => {
    const a = document.createElement("a");
    a.download = "GradeSheet.jpg";
    a.href = canvas.toDataURL("image/jpeg");
    a.click();
    applySheetZoom();
  });
});

/* =======================================================
   INTERNAL MARKS ENGINE
======================================================= */
document
  .getElementById("excel-file-internal")
  .addEventListener("change", function (e) {
    const fileName = e.target.files[0]
      ? e.target.files[0].name
      : "Upload internal marks file";
    document.getElementById("file-name-display-internal").innerText = fileName;
  });
document
  .getElementById("regno-input-internal")
  .addEventListener("keypress", function (e) {
    if (e.key === "Enter") {
      e.preventDefault();
      document.getElementById("calc-internal-btn").click();
    }
  });
document
  .getElementById("calc-internal-btn")
  .addEventListener("click", function () {
    const fileInput = document.getElementById("excel-file-internal");
    const regNo = document.getElementById("regno-input-internal").value.trim();
    const errSpan = document.getElementById("reg-error-internal");
    const btn = this;
    errSpan.style.display = "none";
    if (!fileInput.files || fileInput.files.length === 0) {
      customAlert("Please upload your Internal Marks Excel file first.");
      return;
    }
    if (!regNo || regNo.length < 5) {
      errSpan.innerText = "Registration number required";
      errSpan.style.display = "block";
      return;
    }
    const originalText = btn.innerHTML;
    btn.innerHTML =
      '<i class="ri-loader-4-line ri-spin"></i> Extracting Data...';
    btn.disabled = true;
    btn.style.opacity = "0.8";

    setTimeout(() => {
      const file = fileInput.files[0];
      const reader = new FileReader();
      reader.onload = function (evt) {
        try {
          const data = new Uint8Array(evt.target.result);
          const wb = XLSX.read(data, { type: "array" });
          let targetReg = regNo.toLowerCase().replace(/[^a-z0-9]/g, "");
          let correctRawRows = null;
          let foundSheetName = "UNKNOWN BRANCH";

          for (let i = 0; i < wb.SheetNames.length; i++) {
            const sheetName = wb.SheetNames[i];
            const sheet = wb.Sheets[sheetName];
            const rawRows = XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              defval: "",
            });
            let isStudentInSheet = rawRows.some(
              (row) =>
                row &&
                row.some(
                  (cell) =>
                    String(cell)
                      .toLowerCase()
                      .replace(/[^a-z0-9]/g, "") === targetReg,
                ),
            );
            if (isStudentInSheet) {
              correctRawRows = rawRows;
              foundSheetName = sheetName;
              break;
            }
          }
          if (correctRawRows) {
            processInternalMarks(correctRawRows, regNo, foundSheetName);
          } else {
            const fallbackSheetName = wb.SheetNames[0];
            const fallbackSheet = wb.Sheets[fallbackSheetName];
            const fallbackRows = XLSX.utils.sheet_to_json(fallbackSheet, {
              header: 1,
              defval: "",
            });
            processInternalMarks(fallbackRows, regNo, fallbackSheetName);
          }
        } catch (err) {
          console.error(err);
          customAlert(
            "Failed to parse the file. Please ensure it's a valid Excel/CSV file.",
          );
        } finally {
          btn.innerHTML = originalText;
          btn.disabled = false;
          btn.style.opacity = "1";
        }
      };
      reader.readAsArrayBuffer(file);
    }, 800);
  });

function processInternalMarks(rawRows, regNo, branchName = "") {
  let headerRowIdx = -1,
    rollColIdx = -1;
  for (let r = 0; r < rawRows.length && r < 30; r++) {
    if (!rawRows[r]) continue;
    for (let c = 0; c < rawRows[r].length; c++) {
      let val = String(rawRows[r][c])
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
      if (
        val === "rollno" ||
        val === "registrationno" ||
        val === "regno" ||
        val === "regdno"
      ) {
        headerRowIdx = r;
        rollColIdx = c;
        break;
      }
    }
    if (headerRowIdx !== -1) break;
  }
  if (headerRowIdx === -1) {
    customAlert(
      "Invalid Internal Marks file. Could not find Registration/Roll No column.",
    );
    return;
  }

  let studentRowIdx = -1;
  let targetReg = regNo.toLowerCase().replace(/[^a-z0-9]/g, "");
  for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
    if (!rawRows[r]) continue;
    let cellVal = String(rawRows[r][rollColIdx])
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
    if (cellVal === targetReg) {
      studentRowIdx = r;
      break;
    }
  }
  if (studentRowIdx === -1) {
    document.getElementById("reg-error-internal").innerText =
      "Student registration number not found in this file.";
    document.getElementById("reg-error-internal").style.display = "block";
    return;
  }

  let subCompRowIdx = headerRowIdx,
    maxObtCount = 0;
  for (let r = Math.max(0, headerRowIdx - 2); r <= headerRowIdx + 2; r++) {
    if (!rawRows[r]) continue;
    let obtCount = rawRows[r].filter(
      (c) =>
        String(c).toLowerCase().includes("obtain") ||
        String(c).toLowerCase().includes("max"),
    ).length;
    if (obtCount > maxObtCount) {
      maxObtCount = obtCount;
      subCompRowIdx = r;
    }
  }
  let compRowIdx = subCompRowIdx - 1;
  let subjRowIdx = subCompRowIdx - 2;
  let maxCols = rawRows.reduce((max, r) => Math.max(max, r ? r.length : 0), 0);
  let subjectsList = [];
  let curSubj = null,
    curComp = null;
  let getCell = (r, c) =>
    rawRows[r] && rawRows[r][c] !== undefined
      ? String(rawRows[r][c]).trim()
      : "";

  for (let c = rollColIdx + 1; c < maxCols; c++) {
    let rawSubj = getCell(subjRowIdx, c);
    let rawComp = getCell(compRowIdx, c);
    let rawSubC = getCell(subCompRowIdx, c).toLowerCase();
    let sVal = getCell(studentRowIdx, c);
    if (sVal === "") sVal = "NA";
    if (rawSubj && rawSubj.toUpperCase() !== "NA") {
      if (!curSubj || curSubj.rawName !== rawSubj) {
        let name = rawSubj,
          code = "Unknown",
          type = "N/A";
        let m1 = rawSubj.match(/(.*?)\s*-\s*\((.*?)\)\s*\((.*?)\s*-/);
        let m2 = rawSubj.match(/(.*?)\s*-\s*\((.*?)\)/);
        if (m1) {
          name = m1[1].trim();
          code = m1[2].trim().toUpperCase();
          type = m1[3].trim().toUpperCase();
        } else if (m2) {
          name = m2[1].trim();
          code = m2[2].trim().toUpperCase();
        }
        curSubj = { rawName: rawSubj, name, code, type, components: [] };
        subjectsList.push(curSubj);
        curComp = null;
      }
    }
    if (rawComp && rawComp.toUpperCase() !== "NA") {
      if (
        !rawComp.toLowerCase().includes("obtain") &&
        !rawComp.toLowerCase().includes("max") &&
        !rawComp.toLowerCase().includes("round")
      ) {
        if (!curComp || curComp.name !== rawComp.toUpperCase()) {
          curComp = {
            name: rawComp.toUpperCase(),
            actObt: "NA",
            actMax: "NA",
            rndObt: "NA",
            rndMax: "NA",
            lastType: "act",
          };
          if (curSubj) curSubj.components.push(curComp);
        }
      }
    }
    if (curSubj && curComp && rawSubC) {
      if (rawSubC.includes("round")) {
        curComp.rndObt = sVal;
        curComp.lastType = "rnd";
      } else if (rawSubC.includes("obtain")) {
        curComp.actObt = sVal;
        curComp.lastType = "act";
      } else if (rawSubC.includes("max")) {
        if (curComp.lastType === "rnd") curComp.rndMax = sVal;
        else {
          curComp.actMax = sVal;
          if (curComp.rndMax === "NA") curComp.rndMax = sVal;
        }
      }
    }
  }

  let validSubjects = [];
  let uniqueHeadersSet = new Set();
  subjectsList.forEach((sub) => {
    let validComps = [];
    sub.components.forEach((comp) => {
      if (comp.rndMax === "NA" && comp.actMax !== "NA")
        comp.rndMax = comp.actMax;
      if (comp.actMax === "NA" && comp.rndMax !== "NA")
        comp.actMax = comp.rndMax;
      if (
        comp.actObt !== "NA" ||
        comp.rndObt !== "NA" ||
        comp.actMax !== "NA"
      ) {
        validComps.push(comp);
        if (
          comp.name !== "TOTAL" &&
          comp.name !== "TOTAL:" &&
          comp.name !== "INTERNAL"
        )
          uniqueHeadersSet.add(comp.name);
      }
    });
    if (validComps.length > 0) {
      sub.components = validComps;
      validSubjects.push(sub);
    }
  });

  if (validSubjects.length === 0) {
    customAlert(
      "No internal marks generated. Ensure the Excel follows the official subject/component format.",
    );
    return;
  }
  let dynamicHeaders = Array.from(uniqueHeadersSet);
  let hasTotal = validSubjects.some((sub) =>
    sub.components.some(
      (c) => c.name.includes("TOTAL") || c.name.includes("INTERNAL"),
    ),
  );

  let theadHTML = `<tr>
    <th rowspan="2" style="min-width:40px;padding:12px 6px;text-align:center;font-size:13px;color:#a1a1aa;border-right:1px solid #222;border-bottom:2px solid #333;position:sticky;top:-1px;background:#0a0a0a;z-index:60;">#</th>
    <th rowspan="2" style="min-width:220px;padding:12px 12px;text-align:left;font-size:13px;color:#a1a1aa;border-right:2px solid #333;border-bottom:2px solid #333;position:sticky;top:-1px;background:#0a0a0a;z-index:60;">SUBJECT DETAILS</th>`;
  dynamicHeaders.forEach((h) => {
    theadHTML += `<th colspan="2" style="height:42px;box-sizing:border-box;min-width:140px;padding:10px 8px;text-align:center;font-size:12px;color:#a1a1aa;border-right:2px solid #333;border-bottom:1px solid #333;position:sticky;top:-1px;background:#0a0a0a;z-index:60;letter-spacing:0.5px;">${h}</th>`;
  });
  if (hasTotal)
    theadHTML += `<th rowspan="2" style="min-width:110px;padding:12px 10px;text-align:center;font-size:14px;color:#34d399;font-weight:800;border-left:2px solid #10b981;border-bottom:2px solid #10b981;position:sticky;top:-1px;background:#0a0a0a;z-index:60;">TOTAL SCORE</th>`;
  theadHTML += `</tr><tr>`;
  dynamicHeaders.forEach(() => {
    theadHTML += `<th style="min-width:70px;padding:10px 6px 8px 6px;text-align:center;font-size:11px;color:#60a5fa;border-right:1px dashed #333;border-bottom:2px solid #333;position:sticky;top:38px;background:#0a0a0a;z-index:50;">OBTAINED</th>`;
    theadHTML += `<th style="min-width:70px;padding:10px 6px 8px 6px;text-align:center;font-size:11px;color:#c084fc;border-right:2px solid #333;border-bottom:2px solid #333;position:sticky;top:38px;background:#0a0a0a;z-index:50;">ROUND OFF</th>`;
  });
  theadHTML += `</tr>`;

  let tbodyHTML = validSubjects
    .map((sub, idx) => {
      let row = `<tr style="transition:background 0.2s;" onmouseover="this.style.background='#111'" onmouseout="this.style.background='transparent'">
      <td style="padding:12px 6px;text-align:center;font-weight:600;color:#64748b;border-right:1px solid #222;border-bottom:1px solid #222;font-size:13px;">${idx + 1}</td>
      <td style="padding:12px 12px;border-right:2px solid #333;border-bottom:1px solid #222;white-space:normal;word-break:break-word;">
        <div style="font-weight:700;color:#f8fafc;font-size:14px;line-height:1.4;margin-bottom:8px;">${sub.name}</div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          <span style="background:#1e293b;padding:4px 8px;border-radius:4px;font-size:11px;color:#cbd5e1;font-weight:700;border:1px solid #334155;white-space:nowrap;">${sub.code}</span>
          <span style="background:rgba(168,85,247,0.15);padding:4px 8px;border-radius:4px;font-size:11px;color:#d8b4fe;font-weight:800;border:1px solid rgba(168,85,247,0.3);white-space:nowrap;">${sub.type}</span>
        </div>
      </td>`;
      dynamicHeaders.forEach((h) => {
        let comp = sub.components.find((c) => c.name === h);
        if (comp) {
          let actMaxStr =
            comp.actMax !== "NA"
              ? `<span style="font-size:11px;font-weight:600;opacity:0.7;margin-left:2px;line-height:1;">/${comp.actMax}</span>`
              : "";
          let rndMaxStr =
            comp.rndMax !== "NA"
              ? `<span style="font-size:11px;font-weight:600;opacity:0.7;margin-left:2px;line-height:1;">/${comp.rndMax}</span>`
              : "";
          let actStr =
            comp.actObt !== "NA" && comp.actObt !== ""
              ? `<div style="background:rgba(59,130,246,0.08);border:1px solid rgba(59,130,246,0.3);padding:6px 10px;border-radius:6px;display:inline-flex;justify-content:center;align-items:baseline;min-width:55px;"><strong style="color:#60a5fa;font-size:14px;font-weight:800;line-height:1;">${comp.actObt}</strong>${actMaxStr}</div>`
              : `<span style="color:#475569;font-weight:700;font-size:14px;">-</span>`;
          let rndStr =
            comp.rndObt !== "NA" && comp.rndObt !== ""
              ? `<div style="background:rgba(168,85,247,0.08);border:1px solid rgba(168,85,247,0.3);padding:6px 10px;border-radius:6px;display:inline-flex;justify-content:center;align-items:baseline;min-width:55px;"><strong style="color:#c084fc;font-size:14px;font-weight:800;line-height:1;">${comp.rndObt}</strong>${rndMaxStr}</div>`
              : `<span style="color:#475569;font-weight:700;font-size:14px;">-</span>`;
          row += `<td style="padding:10px 8px;text-align:center;border-right:1px dashed #222;border-bottom:1px solid #222;">${actStr}</td>`;
          row += `<td style="padding:10px 8px;text-align:center;border-right:2px solid #333;border-bottom:1px solid #222;">${rndStr}</td>`;
        } else {
          row += `<td style="padding:10px 8px;text-align:center;color:#475569;border-right:1px dashed #222;border-bottom:1px solid #222;font-size:14px;">-</td>`;
          row += `<td style="padding:10px 8px;text-align:center;color:#475569;border-right:2px solid #333;border-bottom:1px solid #222;font-size:14px;">-</td>`;
        }
      });
      if (hasTotal) {
        let totComp = sub.components.find(
          (c) => c.name.includes("TOTAL") || c.name.includes("INTERNAL"),
        );
        if (totComp) {
          let tObt =
            totComp.rndObt !== "NA" && totComp.rndObt !== ""
              ? totComp.rndObt
              : totComp.actObt;
          let tMax =
            totComp.rndMax !== "NA" && totComp.rndMax !== ""
              ? totComp.rndMax
              : totComp.actMax;
          let maxStr =
            tMax !== "NA"
              ? `<span style="font-size:12px;font-weight:600;opacity:0.8;margin-left:3px;line-height:1;">/${tMax}</span>`
              : "";
          let finalVal =
            tObt === "NA" || tObt === ""
              ? `<span style="color:#475569;font-weight:700;font-size:16px;">-</span>`
              : `<div style="background:rgba(16,185,129,0.1);border:1px solid rgba(16,185,129,0.4);padding:8px 16px;border-radius:8px;display:inline-flex;justify-content:center;align-items:baseline;min-width:70px;"><strong style="color:#34d399;font-size:16px;font-weight:800;line-height:1;">${tObt}</strong>${maxStr}</div>`;
          row += `<td style="padding:12px 10px;text-align:center;background:rgba(16,185,129,0.03);border-left:2px solid rgba(16,185,129,0.3);border-bottom:1px solid #222;">${finalVal}</td>`;
        } else {
          row += `<td style="padding:12px 10px;text-align:center;color:#475569;background:rgba(16,185,129,0.03);border-left:2px solid rgba(16,185,129,0.3);border-bottom:1px solid #222;font-size:16px;">-</td>`;
        }
      }
      row += `</tr>`;
      return row;
    })
    .join("");

  let studentNameFallback =
    rawRows[studentRowIdx][1] !== undefined
      ? String(rawRows[studentRowIdx][1]).toUpperCase()
      : "UNKNOWN";
  let branchBadgeHtml = branchName
    ? `<div style="flex:1;min-width:140px;"><span style="font-size:12px;color:#94a3b8;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">Branch / Program</span><div style="font-size:20px;font-weight:800;color:#60a5fa;margin-top:4px;">${branchName.toUpperCase()}</div></div>`
    : "";

  let modalBody = document.getElementById("internal-modal-body");
  modalBody.innerHTML = `
    <div style="margin-bottom:16px;padding:16px 20px;background:#111;border-radius:10px;border:1px solid #222;display:flex;flex-wrap:wrap;gap:20px;flex-shrink:0;">
      <div style="flex:1;min-width:220px;"><span style="font-size:12px;color:#94a3b8;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">Student Name</span><div style="font-size:20px;font-weight:800;color:#fff;margin-top:4px;">${studentNameFallback}</div></div>
      <div style="flex:1;min-width:180px;"><span style="font-size:12px;color:#94a3b8;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">Registration No</span><div style="font-size:20px;font-weight:800;color:#fff;margin-top:4px;">${regNo}</div></div>
      ${branchBadgeHtml}
    </div>
    <div class="responsive-matrix-wrapper">
      <table style="width:100%;white-space:nowrap;border-collapse:separate;border-spacing:0;text-align:left;background:#0a0a0a;">
        <thead>${theadHTML}</thead>
        <tbody>${tbodyHTML}</tbody>
      </table>
    </div>`;
  const modalOverlay = document.getElementById("internal-result-modal");
  if (modalOverlay) {
    document.body.style.overflow = "hidden";
    modalOverlay.style.display = "flex";
  }
}
