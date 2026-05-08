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
    ) {
      e.preventDefault();
    }
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
  if (container) {
    container.scrollBy({ left: amount, behavior: "smooth" });
  }
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

/* RESET UI ON PAGE LOAD */
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
}
window.addEventListener("load", resetUI);

/* CUSTOM ZOOM & PAN LOGIC */
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
  let isDown = false;
  let startX, startY, scrollLeft, scrollTop;

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
    const walkX = (x - startX) * 1.5;
    const walkY = (y - startY) * 1.5;
    slider.scrollLeft = scrollLeft - walkX;
    slider.scrollTop = scrollTop - walkY;
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

/* SWITCH TAB FUNCTION */
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

/* EXCEL PARSING (SGPA) - MULTI SHEET SUPPORT ADDED */
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

        // Loop through all sheets for SGPA calculation
        wb.SheetNames.forEach((sheetName) => {
          const sheet = wb.Sheets[sheetName];
          const rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          const formatted = rawData.map((row) => {
            let newRow = {};
            for (let key in row) {
              newRow[key.trim()] = row[key];
            }
            return newRow;
          });
          workbookData = workbookData.concat(formatted);
        });
      } catch (error) {
        customAlert(
          "Failed to read the Excel file. Please ensure it's a valid .xlsx file.",
        );
      }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
  }
});

document.getElementById("regno-input").addEventListener("input", function (e) {
  if (isReportGenerated) {
    const btn = document.getElementById("calculate-btn");
    btn.disabled = false;
    btn.innerHTML = "Generate Report";
    btn.style.cursor = "pointer";
  }
});

/* GENERATE REPORT (SGPA) */
document.getElementById("calculate-btn").addEventListener("click", function () {
  const errElements = document.querySelectorAll(".error-msg");
  errElements.forEach((el) => (el.style.display = "none"));

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

      subjectsArray.push({ name: subject, grade: grade });
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
  let hasBacklogs = actualBacklogs.length > 0;
  let isOutstanding = parseFloat(sgpa) >= 9.0 && !hasBacklogs;

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
   ULTIMATE INTERNAL MARKS ENGINE (WITH LOADING ANIMATION & MULTI-SHEET)
======================================================= */

document
  .getElementById("excel-file-internal")
  .addEventListener("change", function (e) {
    const fileName = e.target.files[0]
      ? e.target.files[0].name
      : "Upload internal marks file";
    document.getElementById("file-name-display-internal").innerText = fileName;
  });

/* ---> Added "Enter" Key Listener for Internal Marks Registration Input <--- */
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

    // PREMIUM PROCESSING ANIMATION
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

          // Loop through ALL sheets to find the student's specific branch sheet
          for (let i = 0; i < wb.SheetNames.length; i++) {
            const sheetName = wb.SheetNames[i];
            const sheet = wb.Sheets[sheetName];
            const rawRows = XLSX.utils.sheet_to_json(sheet, {
              header: 1,
              defval: "",
            });

            // Check if the student's registration number exists anywhere in this sheet
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
              break; // Found the student, stop searching sheets
            }
          }

          // If found in a specific sheet, process those rows. Otherwise, fallback to the first sheet to trigger the default "not found" error
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
  // 1. Find Header Row robustly
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

  // 2. Find Student Row robustly
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

  // 3. Find structural rows dynamically based on keywords
  let subCompRowIdx = headerRowIdx;
  let maxObtCount = 0;

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

  // 4. Smart Left-to-Right Extraction Engine
  let subjectsList = [];
  let curSubj = null;
  let curComp = null;

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
        if (curComp.lastType === "rnd") {
          curComp.rndMax = sVal;
        } else {
          curComp.actMax = sVal;
          if (curComp.rndMax === "NA") curComp.rndMax = sVal;
        }
      }
    }
  }

  // 5. Cleanup Empty/Invalid Blocks
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
        ) {
          uniqueHeadersSet.add(comp.name);
        }
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

  // 6. Build Compact, Single-Screen Matrix Layout (PERFECT STICKY HEADER OVERLAP FIX)
  let dynamicHeaders = Array.from(uniqueHeadersSet);
  let hasTotal = validSubjects.some((sub) =>
    sub.components.some(
      (c) => c.name.includes("TOTAL") || c.name.includes("INTERNAL"),
    ),
  );

  let theadHTML = `<tr>
        <th rowspan="2" style="min-width: 40px; padding: 12px 6px; text-align: center; font-size: 13px; color: #a1a1aa; border-right: 1px solid #222; border-bottom: 2px solid #333; position: sticky; top: -1px; background: #0a0a0a; z-index: 60;">#</th>
        <th rowspan="2" style="min-width: 220px; padding: 12px 12px; text-align: left; font-size: 13px; color: #a1a1aa; border-right: 2px solid #333; border-bottom: 2px solid #333; position: sticky; top: -1px; background: #0a0a0a; z-index: 60;">SUBJECT DETAILS</th>`;

  dynamicHeaders.forEach((h) => {
    theadHTML += `<th colspan="2" style="height: 42px; box-sizing: border-box; min-width: 140px; padding: 10px 8px; text-align: center; font-size: 12px; color: #a1a1aa; border-right: 2px solid #333; border-bottom: 1px solid #333; position: sticky; top: -1px; background: #0a0a0a; z-index: 60; letter-spacing: 0.5px;">${h}</th>`;
  });

  if (hasTotal) {
    theadHTML += `<th rowspan="2" style="min-width: 110px; padding: 12px 10px; text-align: center; font-size: 14px; color: #34d399; font-weight: 800; border-left: 2px solid #10b981; border-bottom: 2px solid #10b981; position: sticky; top: -1px; background: #0a0a0a; z-index: 60;">TOTAL SCORE</th>`;
  }
  theadHTML += `</tr><tr>`;

  dynamicHeaders.forEach((h) => {
    theadHTML += `<th style="min-width: 70px; padding: 10px 6px 8px 6px; text-align: center; font-size: 11px; color: #60a5fa; border-right: 1px dashed #333; border-bottom: 2px solid #333; position: sticky; top: 38px; background: #0a0a0a; z-index: 50;">OBTAINED</th>`;
    theadHTML += `<th style="min-width: 70px; padding: 10px 6px 8px 6px; text-align: center; font-size: 11px; color: #c084fc; border-right: 2px solid #333; border-bottom: 2px solid #333; position: sticky; top: 38px; background: #0a0a0a; z-index: 50;">ROUND OFF</th>`;
  });
  theadHTML += `</tr>`;

  let tbodyHTML = validSubjects
    .map((sub, idx) => {
      let row = `<tr style="transition: background 0.2s;" onmouseover="this.style.background='#111'" onmouseout="this.style.background='transparent'">
            <td style="padding: 12px 6px; text-align: center; font-weight: 600; color: #64748b; border-right: 1px solid #222; border-bottom: 1px solid #222; font-size: 13px;">${idx + 1}</td>
            <td style="padding: 12px 12px; border-right: 2px solid #333; border-bottom: 1px solid #222; white-space: normal; word-break: break-word;">
                <div style="font-weight: 700; color: #f8fafc; font-size: 14px; line-height: 1.4; margin-bottom: 8px;">${sub.name}</div>
                <div style="display: flex; gap: 6px; flex-wrap: wrap;">
                    <span style="background: #1e293b; padding: 4px 8px; border-radius: 4px; font-size: 11px; color: #cbd5e1; font-weight: 700; border: 1px solid #334155; white-space: nowrap;">${sub.code}</span>
                    <span style="background: rgba(168, 85, 247, 0.15); padding: 4px 8px; border-radius: 4px; font-size: 11px; color: #d8b4fe; font-weight: 800; border: 1px solid rgba(168, 85, 247, 0.3); white-space: nowrap;">${sub.type}</span>
                </div>
            </td>`;

      dynamicHeaders.forEach((h) => {
        let comp = sub.components.find((c) => c.name === h);
        if (comp) {
          let actMaxStr =
            comp.actMax !== "NA"
              ? `<span style="font-size:11px; font-weight:600; opacity:0.7; margin-left: 2px; line-height: 1;">/${comp.actMax}</span>`
              : "";
          let rndMaxStr =
            comp.rndMax !== "NA"
              ? `<span style="font-size:11px; font-weight:600; opacity:0.7; margin-left: 2px; line-height: 1;">/${comp.rndMax}</span>`
              : "";

          let actStr =
            comp.actObt !== "NA" && comp.actObt !== ""
              ? `<div style="background: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.3); padding: 6px 10px; border-radius: 6px; display: inline-flex; justify-content: center; align-items: baseline; min-width: 55px;"><strong style="color: #60a5fa; font-size: 14px; font-weight: 800; line-height: 1;">${comp.actObt}</strong>${actMaxStr}</div>`
              : `<span style="color:#475569; font-weight:700; font-size: 14px;">-</span>`;

          let rndStr =
            comp.rndObt !== "NA" && comp.rndObt !== ""
              ? `<div style="background: rgba(168, 85, 247, 0.08); border: 1px solid rgba(168, 85, 247, 0.3); padding: 6px 10px; border-radius: 6px; display: inline-flex; justify-content: center; align-items: baseline; min-width: 55px;"><strong style="color: #c084fc; font-size: 14px; font-weight: 800; line-height: 1;">${comp.rndObt}</strong>${rndMaxStr}</div>`
              : `<span style="color:#475569; font-weight:700; font-size: 14px;">-</span>`;

          row += `<td style="padding: 10px 8px; text-align: center; border-right: 1px dashed #222; border-bottom: 1px solid #222;">${actStr}</td>`;
          row += `<td style="padding: 10px 8px; text-align: center; border-right: 2px solid #333; border-bottom: 1px solid #222;">${rndStr}</td>`;
        } else {
          row += `<td style="padding: 10px 8px; text-align: center; color: #475569; border-right: 1px dashed #222; border-bottom: 1px solid #222; font-size: 14px;">-</td>`;
          row += `<td style="padding: 10px 8px; text-align: center; color: #475569; border-right: 2px solid #333; border-bottom: 1px solid #222; font-size: 14px;">-</td>`;
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
              ? `<span style="font-size:12px; font-weight:600; opacity:0.8; margin-left: 3px; line-height: 1;">/${tMax}</span>`
              : "";

          let finalVal =
            tObt === "NA" || tObt === ""
              ? `<span style="color:#475569; font-weight:700; font-size: 16px;">-</span>`
              : `<div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.4); padding: 8px 16px; border-radius: 8px; display: inline-flex; justify-content: center; align-items: baseline; min-width: 70px;"><strong style="color: #34d399; font-size: 16px; font-weight: 800; line-height: 1;">${tObt}</strong>${maxStr}</div>`;

          row += `<td style="padding: 12px 10px; text-align: center; background: rgba(16, 185, 129, 0.03); border-left: 2px solid rgba(16, 185, 129, 0.3); border-bottom: 1px solid #222;">${finalVal}</td>`;
        } else {
          row += `<td style="padding: 12px 10px; text-align: center; color: #475569; background: rgba(16, 185, 129, 0.03); border-left: 2px solid rgba(16, 185, 129, 0.3); border-bottom: 1px solid #222; font-size: 16px;">-</td>`;
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
    ? `
      <div style="flex: 1; min-width: 140px;">
          <span style="font-size: 12px; color: #94a3b8; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Branch / Program</span>
          <div style="font-size: 20px; font-weight: 800; color: #60a5fa; margin-top: 4px;">${branchName.toUpperCase()}</div>
      </div>
    `
    : "";

  let modalBody = document.getElementById("internal-modal-body");
  modalBody.innerHTML = `
        <div style="margin-bottom: 16px; padding: 16px 20px; background: #111; border-radius: 10px; border: 1px solid #222; display: flex; flex-wrap: wrap; gap: 20px; flex-shrink: 0;">
            <div style="flex: 1; min-width: 220px;">
                <span style="font-size: 12px; color: #94a3b8; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Student Name</span>
                <div style="font-size: 20px; font-weight: 800; color: #fff; margin-top: 4px;">${studentNameFallback}</div>
            </div>
            <div style="flex: 1; min-width: 180px;">
                <span style="font-size: 12px; color: #94a3b8; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Registration No</span>
                <div style="font-size: 20px; font-weight: 800; color: #fff; margin-top: 4px;">${regNo}</div>
            </div>
            ${branchBadgeHtml}
        </div>
        <div class="responsive-matrix-wrapper">
            <table style="width: 100%; white-space: nowrap; border-collapse: separate; border-spacing: 0; text-align: left; background: #0a0a0a;">
                <thead>${theadHTML}</thead>
                <tbody>${tbodyHTML}</tbody>
            </table>
        </div>
    `;

  let modalOverlay = document.getElementById("internal-result-modal");
  if (modalOverlay) {
    document.body.style.overflow = "hidden";
    modalOverlay.style.display = "flex";
  }
}
