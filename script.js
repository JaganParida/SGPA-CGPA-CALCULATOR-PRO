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

// Safely check if ENV is defined
const GOOGLE_SCRIPT_URL =
  typeof ENV !== "undefined" ? ENV.GOOGLE_SCRIPT_URL : "";

/* PREVENT BROWSER ZOOM (KEYBOARD & MOUSE)  */
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

    if (Date.now() < end) {
      requestAnimationFrame(frame);
    }
  })();
}

/* NAVBAR SCROLL */
window.addEventListener("scroll", () => {
  const nav = document.getElementById("main-navbar");
  if (window.scrollY > 20) nav.classList.add("scrolled");
  else nav.classList.remove("scrolled");
});

/* RESET UI ON PAGE LOAD */
function resetUI() {
  const semSelect = document.getElementById("semester-number");
  semSelect.value = "";
  semSelect.disabled = false;
  document.getElementById("sem-lock-icon").style.display = "none";

  document.getElementById("regno-input").value = "";

  const fileInput = document.getElementById("excel-file");
  fileInput.value = "";
  fileInput.disabled = false;
  document.getElementById("file-name-display").innerText =
    "Drag & drop or click to browse";
  document.getElementById("file-name-display").style.color = "";
  document.getElementById("file-content-ui").classList.remove("locked-ui");
  document.getElementById("file-lock-icon").style.display = "none";

  workbookData = [];

  document.getElementById("report-output").innerHTML = "";
  document.getElementById("download-actions").style.display = "none";

  const calcBtn = document.getElementById("calculate-btn");
  calcBtn.disabled = false;
  calcBtn.innerHTML = "Generate Report";
  calcBtn.style.cursor = "pointer";

  isReportGenerated = false;
  currentReportData = null;
  currentZoomLevel = 1.0;

  document
    .querySelectorAll(".error-msg")
    .forEach((el) => (el.style.display = "none"));
}

window.addEventListener("load", resetUI);

/*  CUSTOM ZOOM & PAN LOGIC  */
function applySheetZoom() {
  const sheet = document.getElementById("grade-sheet");
  const container = document.getElementById("grade-sheet-target");
  if (!sheet || !container) return;

  const rawHeight = sheet.offsetHeight;

  sheet.style.transform = `scale(${currentZoomLevel})`;

  container.style.width = `${794 * currentZoomLevel}px`;
  container.style.height = `${rawHeight * currentZoomLevel}px`;

  const zoomLabel = document.getElementById("zoom-level-label");
  if (zoomLabel) {
    zoomLabel.innerText = Math.round(currentZoomLevel * 100) + "%";
  }
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

  if (availableWidth > 0 && availableWidth < 794) {
    currentZoomLevel = availableWidth / 794;
  } else {
    currentZoomLevel = 1.0;
  }
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

/*  POPUPS & MENUS  */
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

/*  UPDATED SWITCH TAB FUNCTION  */
function switchTab(tabId) {
  // 1. Hide both sections and reset tab buttons
  document.getElementById("sgpa-section").style.display = "none";
  document.getElementById("cgpa-section").style.display = "none";
  document.getElementById("tab-sgpa").classList.remove("active");
  document.getElementById("tab-cgpa").classList.remove("active");

  // 2. Show the targeted section and activate its button
  document.getElementById(tabId + "-section").style.display = "block";
  document.getElementById("tab-" + tabId).classList.add("active");

  // 3. Sync the Active class for the Navbar links
  const navSgpa = document.getElementById("nav-sgpa");
  const navCgpa = document.getElementById("nav-cgpa");
  if (navSgpa) navSgpa.classList.remove("active");
  if (navCgpa) navCgpa.classList.remove("active");

  const activeNavLink = document.getElementById("nav-" + tabId);
  if (activeNavLink) activeNavLink.classList.add("active");

  // 4. Smooth scroll to the Tab section
  const tabWrapper = document.querySelector(".tab-wrapper");
  if (tabWrapper) {
    tabWrapper.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

/*  EXCEL PARSING  */
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
        const sheet = wb.Sheets[wb.SheetNames[0]];

        const rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        workbookData = rawData.map((row) => {
          let newRow = {};
          for (let key in row) {
            newRow[key.trim()] = row[key];
          }
          return newRow;
        });
      } catch (error) {
        console.error("Excel Parsing Error: ", error);
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

/*  GENERATE REPORT  */
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

      // Logic: Total credits strictly sums all displayed credits regardless of grade
      totalCredits += credit;

      // Logic: Track backlogs if F, S, or M
      if (["F", "M", "S"].includes(grade)) {
        if (!actualBacklogs.includes(subject)) {
          actualBacklogs.push(subject);
        }
      }

      // Logic: Credits Cleared strictly ignores F, S, and M
      if (!["F", "S", "M"].includes(grade)) {
        creditsCleared += credit;
      }

      // Logic: SGPA Calculation (points * credit) / totalCredits
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

  // SMART GLASSY BANNER LOGIC
  let bannerHTML = "";
  let hasBacklogs = actualBacklogs.length > 0;
  let isOutstanding = parseFloat(sgpa) >= 9.0 && !hasBacklogs;

  if (isOutstanding) {
    bannerHTML = `
        <div class="report-status-banner status-outstanding">
            <div class="banner-icon"><img src="https://cdn-icons-png.flaticon.com/512/3176/3176294.png" style="width: 36px; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.2));" alt="Medal"></div>
            <div class="banner-content">
                <h4>Outstanding Performance! 🏆</h4>
                <p>Incredible job! You achieved a stellar SGPA of ${sgpa}. Keep up the excellent work!</p>
            </div>
        </div>
      `;
  } else if (hasBacklogs) {
    bannerHTML = `
        <div class="report-status-banner status-warning">
            <div class="banner-icon"><i class="ri-error-warning-fill"></i></div>
            <div class="banner-content">
                <h4>Action Required: Pending Subjects</h4>
                <p>You have pending backlogs (${actualBacklogs.join(", ")}). Please prepare well and clear them in upcoming exams.</p>
            </div>
        </div>
      `;
  } else {
    bannerHTML = `
        <div class="report-status-banner status-clear">
            <div class="banner-icon"><i class="ri-verified-badge-fill"></i></div>
            <div class="banner-content">
                <h4>All Clear! 🎉</h4>
                <p>Congratulations! You have successfully cleared all subjects for this semester.</p>
            </div>
        </div>
      `;
  }

  reportDiv.innerHTML = `
        ${bannerHTML}
        
        <div id="report-scroll-wrapper" class="report-scroll-wrapper">
            <div id="grade-sheet-target" class="grade-sheet-target">
                <div id="grade-sheet" class="grade-sheet">
                    <div class="sheet-top-header">
                        <div>${timeString}</div>
                        <div>GradeFlow - Streamlining your academic journey</div>
                    </div>
                    <div class="sheet-logos"><img src="Assets/cutm.png" alt="Logo" class="sheet-logo-img" onerror="this.src='Assets/cutm_text.jpg'"></div>
                    <div class="sheet-titles">
                        <h1>Centurion University of Technology and Management</h1>
                        <h3>Jatni, Khurda, Odisha</h3>
                        <h2>Semester Grade Sheet</h2>
                    </div>
                    <div class="student-info-grid">
                        <div class="info-row"><span class="lbl">Student Regd. No</span> <span class="val">: ${regNo}</span></div>
                        <div class="info-row"><span class="lbl">Student Name</span> <span class="val">: ${studentName.toUpperCase()}</span></div>
                        <div class="info-row"><span class="lbl">Batch</span> <span class="val">: ${batch}</span></div>
                        <div class="info-row"><span class="lbl">Semester</span> <span class="val">: Sem ${sem}</span></div>
                    </div>
                    <table class="result-table">
                        <thead><tr><th>SL.NO</th><th>SUB.CODE</th><th>SUBJECT</th><th>TYPE</th><th>CREDIT</th><th>GRADE</th></tr></thead>
                        <tbody>${rowsHTML}</tbody>
                    </table>
                    <div class="summary-row" style="margin-top: 80px;">
                        <div>Total Credits : ${totalCredits}</div>
                        <div>Credits Cleared : ${creditsCleared}</div>
                        <div>SGPA : ${sgpa}</div>
                    </div>
                    <div class="signature-row">
                        <div>Date : ${dateString}</div>
                        <div>Dean, Examinations</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div style="text-align: center; width: 100%;">
            <div class="inline-zoom-controls">
                <button onclick="changeZoom(-0.1)">-</button>
                <span id="zoom-level-label">100%</span>
                <button onclick="changeZoom(0.1)">+</button>
            </div>
        </div>`;

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
    // We use URLSearchParams so Google Script can read it easily
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
    if (isOutstanding) {
      fireConfetti();
    }
  }, 50);
});

/*  WHATSAPP SHARE  */
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

/*  CGPA CALCULATOR  */
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

/*  PERFECT PDF/IMAGE EXPORTS  */
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
