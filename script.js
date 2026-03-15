const validGradePoints = {
  O: 10,
  E: 9,
  A: 8,
  B: 7,
  C: 6,
  D: 5,
  R: 10,
  F: 0,
  M: 0,
  S: 0,
};
let workbookData = [];
let currentReportData = null;
let isReportGenerated = false;

// Keep your Google Apps Script URL
const GOOGLE_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbwNjwC0SNrXPd5IcsjwkCvy_gJEUbPf5gVkEbIgzDWLDv1q1M2-lJkEVDd9bZmHGR3daA/exec";

/* ================= NAVBAR SCROLL ================= */
window.addEventListener("scroll", () => {
  const nav = document.getElementById("main-navbar");
  if (window.scrollY > 20) nav.classList.add("scrolled");
  else nav.classList.remove("scrolled");
});

/* ================= RESET UI ON PAGE LOAD ================= */
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

  isReportGenerated = false;
  currentReportData = null;

  document
    .querySelectorAll(".error-msg")
    .forEach((el) => (el.style.display = "none"));
}

window.addEventListener("load", resetUI);

/* ================= PERFECT A4 SCALE LOGIC ================= */
function adjustSheetScale() {
  const sheet = document.getElementById("grade-sheet");
  const container = document.getElementById("grade-sheet-target");
  const wrapper = document.getElementById("report-output");
  if (!sheet || !container || !wrapper) return;

  // Reset to natural state
  sheet.style.transform = "none";
  container.style.width = "794px";
  container.style.height = "auto";

  // Calculate if mobile screen is smaller than A4
  const wrapperWidth = wrapper.clientWidth - 20; // 10px buffer left and right
  if (wrapperWidth > 0 && wrapperWidth < 794) {
    const scale = wrapperWidth / 794;
    sheet.style.transform = `scale(${scale})`;
    sheet.style.transformOrigin = `top left`;

    // Resize the container exactingly to avoid phantom blank spaces
    container.style.width = `${794 * scale}px`;
    container.style.height = `${sheet.offsetHeight * scale}px`;
  }
}
window.addEventListener("resize", adjustSheetScale);

/* ================= POPUPS & MENUS ================= */
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

function switchTab(tabId) {
  document.getElementById("sgpa-section").style.display = "none";
  document.getElementById("cgpa-section").style.display = "none";
  document.getElementById("tab-sgpa").classList.remove("active");
  document.getElementById("tab-cgpa").classList.remove("active");
  document.getElementById(tabId + "-section").style.display = "block";
  document.getElementById("tab-" + tabId).classList.add("active");
}

/* ================= EXCEL PARSING ================= */
document.getElementById("excel-file").addEventListener("change", function (e) {
  const fileName = e.target.files[0]
    ? e.target.files[0].name
    : "Click to upload .xlsx file";
  document.getElementById("file-name-display").innerText = fileName;
  if (e.target.files[0]) {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      workbookData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    };
    reader.readAsArrayBuffer(e.target.files[0]);
  }
});

document.getElementById("regno-input").addEventListener("input", function (e) {
  if (isReportGenerated) {
    const btn = document.getElementById("calculate-btn");
    btn.disabled = false;
    btn.innerHTML = "Generate Report";
  }
});

/* ================= GENERATE REPORT ================= */
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

  let totalPoints = 0,
    totalCredits = 0,
    creditsCleared = 0;
  const studentName = studentRows[0]["Name"] || "Unknown Student";
  let batch = regNo.length >= 2 ? "20" + regNo.substring(0, 2) : "N/A";
  let subjectsArray = [];

  const rowsHTML = studentRows
    .map((row, i) => {
      const grade = String(row["Grade"]).trim().toUpperCase();
      const credit = parseCredit(row["Credits"]);
      const subject = row["Subject_Name"] || "Unknown";
      const type = row["Type"] || "PP";
      subjectsArray.push({ name: subject, grade: grade });
      let points =
        validGradePoints[grade] !== undefined ? validGradePoints[grade] : 0;
      if (grade !== "S") {
        totalPoints += points * credit;
        totalCredits += credit;
        if (!["F", "M"].includes(grade)) creditsCleared += credit;
      } else {
        creditsCleared += credit;
      }
      return `<tr><td>${i + 1}</td><td>${row["Subject_Code"] || ""}</td><td>${subject}</td><td>${type}</td><td>${credit}</td><td>${grade}</td></tr>`;
    })
    .join("");

  const sgpa =
    totalCredits > 0 ? (totalPoints / totalCredits).toFixed(2) : "0.00";
  const actualBacklogs = studentRows
    .filter((row) =>
      ["F", "M"].includes(String(row["Grade"]).trim().toUpperCase()),
    )
    .map((r) => r["Subject_Name"]);

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

  // Only show backlogs notice if any, else keep it ultra-clean like image 3
  const feedback =
    actualBacklogs.length > 0
      ? `<div class="feedback-box" style="border-color:#ef4444; background:#fef2f2; color:#b91c1c;"><strong>Backlogs:</strong> ${actualBacklogs.join(", ")}</div>`
      : ``;

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

  // EXACT REPRODUCTION OF IMAGE 3 LAYOUT
  reportDiv.innerHTML = `
        <div id="grade-sheet-target" style="position: relative;">
            <div id="grade-sheet" class="grade-sheet">
                <div class="sheet-top-header">
                    <div>${timeString}</div>
                    <div>iCloudEMS - We honor great education and great educationists</div>
                </div>
                <div class="sheet-logos"><img src="cutm.png" alt="Logo" class="sheet-logo-img" onerror="this.src='cutm_text.jpg'"></div>
                <div class="sheet-titles">
                    <h1>Centurion University of Technology and Management</h1>
                    <h3>School of Engineering & Technology, Bhubaneswar</h3>
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
                ${feedback}
                <div class="summary-row">
                    <div>Total Credits : ${totalCredits}</div>
                    <div>Credits Cleared : ${creditsCleared}</div>
                    <div>SGPA : ${sgpa}</div>
                    <div>CGPA : N/A</div>
                </div>
                <div class="signature-row">
                    <div>Date : ${dateString}</div>
                    <div>Dean, Examinations</div>
                </div>
            </div>
        </div>`;

  document.getElementById("download-actions").style.display = "flex";

  // LOCK STATE
  semInput.disabled = true;
  document.getElementById("sem-lock-icon").style.display = "inline-block";
  document.getElementById("excel-file").disabled = true;
  document.getElementById("file-content-ui").classList.add("locked-ui");
  document.getElementById("file-lock-icon").style.display = "inline-block";

  const calcBtn = document.getElementById("calculate-btn");
  calcBtn.disabled = true;
  calcBtn.innerHTML = "Report Generated ✓";
  isReportGenerated = true;

  // GOOGLE SHEETS LOGGING
  if (GOOGLE_SCRIPT_URL !== "YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE") {
    const formData = new FormData();
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

  // Auto trigger the scale down after rendering so it fits the mobile view instantly
  setTimeout(() => {
    adjustSheetScale();
    reportDiv.scrollIntoView({ behavior: "smooth" });
  }, 50);
});

/* ================= WHATSAPP SHARE ================= */
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
    text += `Please click the link below for detailed subject-wise grades.\nhttps://cutm-calc.web.app`;
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

/* ================= CGPA CALCULATOR ================= */
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
  if (!val) return 0;
  return val
    .toString()
    .split("+")
    .reduce((a, c) => a + parseFloat(c || 0), 0);
}

/* ================= PERFECT PDF/IMAGE EXPORTS ================= */
document.getElementById("download-btn").addEventListener("click", () => {
  const sheet = document.getElementById("grade-sheet");

  // Temporarily reset scale to extract Full High-Res Desktop Render
  const currentTransform = sheet.style.transform;
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

    // Restore mobile scale UI immediately after snapping photo
    adjustSheetScale();
  });
});

document.getElementById("download-photo-btn").addEventListener("click", () => {
  const sheet = document.getElementById("grade-sheet");

  // Temporarily reset scale to extract Full High-Res Desktop Render
  const currentTransform = sheet.style.transform;
  sheet.style.transform = "none";

  html2canvas(sheet, { scale: 3 }).then((canvas) => {
    const a = document.createElement("a");
    a.download = "GradeSheet.jpg";
    a.href = canvas.toDataURL("image/jpeg");
    a.click();

    // Restore mobile scale UI immediately after snapping photo
    adjustSheetScale();
  });
});
