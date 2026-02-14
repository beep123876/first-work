const STORAGE_KEY = "attendanceDashboardData_v1";

const state = {
  records: [],
  employees: new Map(),
};

const excelFileInput = document.getElementById("excelFile");
const employeeSelect = document.getElementById("employeeSelect");
const saveBtn = document.getElementById("saveBtn");
const clearBtn = document.getElementById("clearBtn");
const empNoEl = document.getElementById("empNo");
const overtimeEl = document.getElementById("overtime");
const vacationEl = document.getElementById("vacation");
const earlyLeaveEl = document.getElementById("earlyLeave");
const sickLeaveEl = document.getElementById("sickLeave");
const detailRows = document.getElementById("detailRows");

function parseRecord(row) {
  const name = String(row["이름"] ?? row["성명"] ?? "").trim();
  const employeeId = String(row["사원번호"] ?? row["사번"] ?? "").trim();
  const rawDate = row["날짜"] ?? row["일자"];
  const attendanceType = String(row["근태유형"] ?? row["유형"] ?? "").trim();
  const overtimeHours = Number(row["시간외(시간)"] ?? row["시간외"] ?? 0) || 0;

  const date = convertExcelDate(rawDate);
  if (!name || !employeeId || !date) {
    return null;
  }

  return {
    name,
    employeeId,
    date: date.toISOString().slice(0, 10),
    attendanceType,
    overtimeHours,
  };
}

function convertExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value)) return value;
  if (typeof value === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = value * 24 * 60 * 60 * 1000;
    return new Date(epoch.getTime() + ms);
  }
  if (typeof value === "string" && value.trim()) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed)) return parsed;
  }
  return null;
}

function rebuildEmployeeMap(records) {
  state.employees = new Map();
  records.forEach((r) => {
    if (!state.employees.has(r.name)) {
      state.employees.set(r.name, r.employeeId);
    }
  });
}

function loadSavedData() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    state.records = Array.isArray(parsed) ? parsed : [];
    rebuildEmployeeMap(state.records);
    populateEmployees();
  } catch {
    state.records = [];
  }
}

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.records));
  alert("데이터를 저장했습니다.");
}

function populateEmployees() {
  employeeSelect.innerHTML = '<option value="">직원을 선택하세요</option>';
  [...state.employees.keys()].sort((a, b) => a.localeCompare(b, "ko")).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    employeeSelect.appendChild(opt);
  });
}

function getMonthlySummary(name) {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();

  const filtered = state.records.filter((r) => {
    if (r.name !== name) return false;
    const d = new Date(r.date);
    return d.getFullYear() === year && d.getMonth() === month;
  });

  const summary = {
    overtime: 0,
    vacation: 0,
    earlyLeave: 0,
    sickLeave: 0,
  };

  filtered.forEach((r) => {
    summary.overtime += Number(r.overtimeHours) || 0;
    if (r.attendanceType === "휴가") summary.vacation += 1;
    if (r.attendanceType === "조퇴") summary.earlyLeave += 1;
    if (r.attendanceType === "병가") summary.sickLeave += 1;
  });

  return { summary, filtered };
}

function renderDetails(records) {
  if (!records.length) {
    detailRows.innerHTML = '<tr><td colspan="3" class="empty">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  detailRows.innerHTML = records
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((r) => `<tr><td>${r.date}</td><td>${r.attendanceType || "-"}</td><td>${r.overtimeHours}</td></tr>`)
    .join("");
}

function updateDashboard() {
  const name = employeeSelect.value;
  if (!name) {
    empNoEl.textContent = "-";
    overtimeEl.textContent = "0";
    vacationEl.textContent = "0";
    earlyLeaveEl.textContent = "0";
    sickLeaveEl.textContent = "0";
    renderDetails([]);
    return;
  }

  empNoEl.textContent = state.employees.get(name) || "-";
  const { summary, filtered } = getMonthlySummary(name);
  overtimeEl.textContent = String(summary.overtime);
  vacationEl.textContent = String(summary.vacation);
  earlyLeaveEl.textContent = String(summary.earlyLeave);
  sickLeaveEl.textContent = String(summary.sickLeave);
  renderDetails(filtered);
}

excelFileInput.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = e.target?.result;
    const workbook = XLSX.read(data, { type: "binary" });
    const firstSheet = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });

    const parsedRecords = rows.map(parseRecord).filter(Boolean);
    state.records = parsedRecords;
    rebuildEmployeeMap(parsedRecords);
    populateEmployees();
    updateDashboard();
    alert(`엑셀 데이터 ${parsedRecords.length}건을 불러왔습니다. 저장 버튼을 눌러 유지하세요.`);
  };
  reader.readAsBinaryString(file);
});

employeeSelect.addEventListener("change", updateDashboard);
saveBtn.addEventListener("click", saveData);
clearBtn.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  state.records = [];
  state.employees.clear();
  populateEmployees();
  updateDashboard();
  alert("저장된 데이터를 초기화했습니다.");
});

loadSavedData();
updateDashboard();
