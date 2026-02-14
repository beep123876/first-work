const STORAGE_KEY = "attendanceDashboardData_v2";
const DRIVE_URL_KEY = "attendanceDashboardDriveUrl_v1";

const state = {
  records: [],
  employeesByMonth: new Map(),
  months: [],
};

const driveUrlInput = document.getElementById("driveUrl");
const loadDriveBtn = document.getElementById("loadDriveBtn");
const saveUrlBtn = document.getElementById("saveUrlBtn");
const clearBtn = document.getElementById("clearBtn");
const monthSelect = document.getElementById("monthSelect");
const employeeSelect = document.getElementById("employeeSelect");

const empNoEl = document.getElementById("empNo");
const overtimeEl = document.getElementById("overtime");
const vacationEl = document.getElementById("vacation");
const sickLeaveEl = document.getElementById("sickLeave");
const earlyLeaveEl = document.getElementById("earlyLeave");
const domesticTripEl = document.getElementById("domesticTrip");
const internationalTripEl = document.getElementById("internationalTrip");
const localTripEl = document.getElementById("localTrip");
const outsideTripEl = document.getElementById("outsideTrip");
const maternityLeaveEl = document.getElementById("maternityLeave");
const pregnancyCheckupEl = document.getElementById("pregnancyCheckup");
const compThisYearEl = document.getElementById("compThisYear");
const compPrevYearEl = document.getElementById("compPrevYear");
const detailRows = document.getElementById("detailRows");

function convertDriveUrl(url) {
  const trimmed = url.trim();
  const fileMatch = trimmed.match(/\/file\/d\/([^/]+)/);
  if (fileMatch) return `https://drive.google.com/uc?export=download&id=${fileMatch[1]}`;

  const queryMatch = trimmed.match(/[?&]id=([^&]+)/);
  if (queryMatch) return `https://drive.google.com/uc?export=download&id=${queryMatch[1]}`;

  return trimmed;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value)) return value;
  if (typeof value === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + value * 24 * 60 * 60 * 1000);
  }
  if (typeof value === "string") {
    const normalized = value.replace(/\(.*?\)/g, "").trim();
    if (!normalized) return null;
    const parsed = new Date(normalized);
    if (!Number.isNaN(parsed)) return parsed;
  }
  return null;
}

function parseHourText(text) {
  if (!text) return 0;
  const cleaned = String(text).trim();
  const hhmm = cleaned.match(/(\d{1,2}):(\d{1,2})/);
  if (hhmm) return Number(hhmm[1]) + Number(hhmm[2]) / 60;
  const numberOnly = cleaned.match(/\d+(\.\d+)?/);
  return numberOnly ? Number(numberOnly[0]) : 0;
}

function extractTypeList(cellText) {
  if (!cellText) return [];
  const text = String(cellText);
  const matches = [...text.matchAll(/종별\s*[:：]\s*([^\n\r]+)/g)].map((m) => m[1].trim());
  if (matches.length) return matches;
  return text.trim() ? [text.trim()] : [];
}

function classifyType(rawType, bucket) {
  const type = rawType.replace(/\s+/g, "");
  if (bucket === "earlyLeave") return { category: "조퇴", subType: "조기퇴근" };
  if (bucket === "overtime") return { category: "시간외", subType: "시간외근무" };

  if (type.includes("병가")) return { category: "병가", subType: rawType };
  if (type.includes("산전후휴가")) return { category: "산전후휴가", subType: rawType };
  if (type.includes("임산부정기검진")) return { category: "임산부정기검진", subType: rawType };

  if (type.includes("관내출장")) return { category: "출장", subType: "관내출장" };
  if (type.includes("관외출장")) return { category: "출장", subType: "관외출장" };
  if (type.includes("국외출장")) return { category: "출장", subType: "국외출장" };
  if (type.includes("국내출장")) return { category: "출장", subType: "국내출장" };
  if (type.includes("출장")) return { category: "출장", subType: rawType };

  if (type.includes("전년도") && type.includes("대체휴무")) {
    return { category: "휴가", subType: "전년도대체휴무" };
  }
  if (!type.includes("전년도") && type.includes("대체휴무")) {
    return { category: "휴가", subType: "당해대체휴무" };
  }
  if (type.includes("휴가")) return { category: "휴가", subType: rawType };

  return { category: "기타", subType: rawType };
}

function parseRow(row, sheetName) {
  const name = String(row["성명"] ?? row["이름"] ?? "").trim();
  if (!name) return [];

  const employeeId = String(row["사원번호"] ?? row["사번"] ?? "-").trim() || "-";
  const date = parseDate(row["날짜"] ?? row["일자"]);
  const dateIso = date ? date.toISOString().slice(0, 10) : "-";

  const results = [];

  const overtimeCell = row["시간외관리"] ?? row["시간외(시간)"] ?? row["시간외"];
  const overtimeHours = parseHourText(String(overtimeCell ?? "").match(/총시간\s*[:：]\s*([^\n\r]+)/)?.[1] ?? overtimeCell);
  if (overtimeHours > 0) {
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "시간외", subType: "시간외근무", overtimeHours });
  }

  const earlyLeaveCell = row["조기퇴근"] ?? row["조퇴"];
  if (parseHourText(earlyLeaveCell) > 0) {
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "조퇴", subType: "조기퇴근", overtimeHours: 0 });
  }

  const leaveTypes = extractTypeList(row["휴가관리"] ?? row["근태유형"] ?? row["유형"]);
  leaveTypes.forEach((t) => {
    const c = classifyType(t, "leave");
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: c.category, subType: c.subType, overtimeHours: 0 });
  });

  const tripTypes = extractTypeList(row["출장관리"]);
  tripTypes.forEach((t) => {
    const c = classifyType(t, "trip");
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: c.category, subType: c.subType, overtimeHours: 0 });
  });

  return results;
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const allRecords = [];

  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    rows.forEach((row) => {
      allRecords.push(...parseRow(row, sheetName));
    });
  });

  state.records = allRecords;
  state.months = workbook.SheetNames.slice();
  rebuildEmployeeMap();
}

function rebuildEmployeeMap() {
  state.employeesByMonth = new Map();
  state.records.forEach((r) => {
    if (!state.employeesByMonth.has(r.month)) state.employeesByMonth.set(r.month, new Map());
    const monthMap = state.employeesByMonth.get(r.month);
    if (!monthMap.has(r.name)) monthMap.set(r.name, r.employeeId);
  });
}

function sortMonths(months) {
  return months.slice().sort((a, b) => {
    const am = Number((a.match(/\d+/) || [999])[0]);
    const bm = Number((b.match(/\d+/) || [999])[0]);
    return am - bm || a.localeCompare(b, "ko");
  });
}

function populateMonths() {
  monthSelect.innerHTML = '<option value="">월을 선택하세요</option>';
  sortMonths(state.months).forEach((month) => {
    const option = document.createElement("option");
    option.value = month;
    option.textContent = month;
    monthSelect.appendChild(option);
  });
}

function populateEmployees() {
  const month = monthSelect.value;
  const monthMap = state.employeesByMonth.get(month) || new Map();
  employeeSelect.innerHTML = '<option value="">직원을 선택하세요</option>';
  [...monthMap.keys()].sort((a, b) => a.localeCompare(b, "ko")).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    employeeSelect.appendChild(opt);
  });
}

function summaryBase() {
  return {
    overtime: 0,
    vacation: 0,
    sickLeave: 0,
    earlyLeave: 0,
    domesticTrip: 0,
    internationalTrip: 0,
    localTrip: 0,
    outsideTrip: 0,
    maternityLeave: 0,
    pregnancyCheckup: 0,
    compThisYear: 0,
    compPrevYear: 0,
  };
}

function getSelectedRecords() {
  const month = monthSelect.value;
  const name = employeeSelect.value;
  if (!month || !name) return [];
  return state.records.filter((r) => r.month === month && r.name === name);
}

function buildSummary(records) {
  const summary = summaryBase();
  records.forEach((r) => {
    if (r.category === "시간외") summary.overtime += r.overtimeHours;
    if (r.category === "휴가") summary.vacation += 1;
    if (r.category === "병가") summary.sickLeave += 1;
    if (r.category === "조퇴") summary.earlyLeave += 1;
    if (r.category === "산전후휴가") summary.maternityLeave += 1;
    if (r.category === "임산부정기검진") summary.pregnancyCheckup += 1;

    if (r.subType === "당해대체휴무") summary.compThisYear += 1;
    if (r.subType === "전년도대체휴무") summary.compPrevYear += 1;

    if (r.subType === "국내출장") summary.domesticTrip += 1;
    if (r.subType === "국외출장") summary.internationalTrip += 1;
    if (r.subType === "관내출장") summary.localTrip += 1;
    if (r.subType === "관외출장") summary.outsideTrip += 1;
  });
  return summary;
}

function renderDetails(records) {
  if (!records.length) {
    detailRows.innerHTML = '<tr><td colspan="4" class="empty">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  detailRows.innerHTML = records
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((r) => `<tr><td>${r.date}</td><td>${r.category}</td><td>${r.subType}</td><td>${r.overtimeHours || 0}</td></tr>`)
    .join("");
}

function updateDashboard() {
  const month = monthSelect.value;
  const name = employeeSelect.value;
  const employeeId = state.employeesByMonth.get(month)?.get(name) || "-";
  const records = getSelectedRecords();
  const summary = buildSummary(records);

  empNoEl.textContent = employeeId;
  overtimeEl.textContent = summary.overtime.toFixed(1);
  vacationEl.textContent = String(summary.vacation);
  sickLeaveEl.textContent = String(summary.sickLeave);
  earlyLeaveEl.textContent = String(summary.earlyLeave);
  domesticTripEl.textContent = String(summary.domesticTrip);
  internationalTripEl.textContent = String(summary.internationalTrip);
  localTripEl.textContent = String(summary.localTrip);
  outsideTripEl.textContent = String(summary.outsideTrip);
  maternityLeaveEl.textContent = String(summary.maternityLeave);
  pregnancyCheckupEl.textContent = String(summary.pregnancyCheckup);
  compThisYearEl.textContent = String(summary.compThisYear);
  compPrevYearEl.textContent = String(summary.compPrevYear);

  renderDetails(records);
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ records: state.records, months: state.months }));
}

function loadSavedState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  const savedUrl = localStorage.getItem(DRIVE_URL_KEY);
  if (savedUrl) driveUrlInput.value = savedUrl;
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    state.records = Array.isArray(parsed.records) ? parsed.records : [];
    state.months = Array.isArray(parsed.months) ? parsed.months : [];
    rebuildEmployeeMap();
    populateMonths();
  } catch {
    state.records = [];
    state.months = [];
  }
}

async function loadFromDrive() {
  const sourceUrl = driveUrlInput.value.trim();
  if (!sourceUrl) {
    alert("구글드라이브 링크를 입력하세요.");
    return;
  }

  const downloadUrl = convertDriveUrl(sourceUrl);
  const response = await fetch(downloadUrl);
  if (!response.ok) throw new Error(`다운로드 실패: ${response.status}`);

  const buffer = await response.arrayBuffer();
  parseWorkbook(buffer);
  saveState();
  localStorage.setItem(DRIVE_URL_KEY, sourceUrl);
  populateMonths();
  monthSelect.value = sortMonths(state.months)[0] || "";
  populateEmployees();
  updateDashboard();
}

loadDriveBtn.addEventListener("click", async () => {
  try {
    await loadFromDrive();
    alert(`드라이브 파일을 반영했습니다. 전체 ${state.records.length}건`);
  } catch (error) {
    alert(`불러오기에 실패했습니다. 파일 공유 권한을 '링크가 있는 모든 사용자(보기)'로 설정했는지 확인하세요.\n${error.message}`);
  }
});

saveUrlBtn.addEventListener("click", () => {
  localStorage.setItem(DRIVE_URL_KEY, driveUrlInput.value.trim());
  alert("구글드라이브 링크를 저장했습니다.");
});

clearBtn.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(DRIVE_URL_KEY);
  state.records = [];
  state.months = [];
  state.employeesByMonth.clear();
  driveUrlInput.value = "";
  populateMonths();
  populateEmployees();
  updateDashboard();
  alert("저장된 데이터/링크를 초기화했습니다.");
});

monthSelect.addEventListener("change", () => {
  populateEmployees();
  updateDashboard();
});
employeeSelect.addEventListener("change", updateDashboard);

loadSavedState();
populateEmployees();
updateDashboard();
