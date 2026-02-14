const STORAGE_KEY = "attendanceDashboardData_v3";
const DEFAULT_DRIVE_XLSX_URL = "https://docs.google.com/spreadsheets/d/1grIcwPHx4XanTASz9UGmANC8L6bNAIMdH5D2h6wP73Q/export?format=xlsx";
const HOURS_PER_DAY = 8;

const state = {
  records: [],
  employeesByMonth: new Map(),
  months: [],
};

const loadDriveBtn = document.getElementById("loadDriveBtn");
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

  if (trimmed.includes("docs.google.com/spreadsheets/d/") && trimmed.includes("/edit")) {
    const fileId = trimmed.match(/\/d\/([^/]+)/)?.[1];
    if (fileId) return `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  }

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

function parseHourValue(value) {
  if (typeof value === "number") return value;
  if (!value) return 0;
  const text = String(value).trim();
  const hhmm = text.match(/(\d{1,2})\s*:\s*(\d{1,2})/);
  if (hhmm) return Number(hhmm[1]) + Number(hhmm[2]) / 60;
  const n = text.match(/\d+(\.\d+)?/);
  return n ? Number(n[0]) : 0;
}

function parseLeaveDurationHours(cellText) {
  const text = String(cellText ?? "");
  const days = Number(text.match(/(\d+)\s*일/)?.[1] ?? 0);
  const hours = Number(text.match(/(\d+)\s*시간/)?.[1] ?? 0);
  const mins = Number(text.match(/(\d+)\s*분/)?.[1] ?? 0);
  const total = days * HOURS_PER_DAY + hours + mins / 60;
  return total > 0 ? total : 0;
}

function parseOvertimeHours(cellText) {
  if (typeof cellText === "number") return cellText;
  const text = String(cellText ?? "");
  if (!text.trim()) return 0;

  const labeled = text.match(/(?:총시간|총시간외|시간외시간|실근무)\s*[:：]\s*([^\n\r]+)/);
  if (labeled) return parseHourValue(labeled[1]);

  if (text.includes("신청시각") || text.includes("실근무") || text.includes("종별")) {
    return 0;
  }

  return parseHourValue(text);
}

function extractTypeList(cellText) {
  if (!cellText) return [];
  const text = String(cellText);
  const matches = [...text.matchAll(/종별\s*[:：]\s*([^\n\r]+)/g)].map((m) => m[1].trim());
  if (matches.length) return matches;
  return text.trim() ? [text.trim()] : [];
}

function classifyType(rawType) {
  const type = rawType.replace(/\s+/g, "");

  if (type.includes("병가")) return { category: "병가", subType: rawType };
  if (type.includes("산전후휴가")) return { category: "산전후휴가", subType: rawType };
  if (type.includes("임산부정기검진")) return { category: "임산부정기검진", subType: rawType };

  if (type.includes("관내출장")) return { category: "출장", subType: "관내출장" };
  if (type.includes("관외출장")) return { category: "출장", subType: "관외출장" };
  if (type.includes("국외출장")) return { category: "출장", subType: "국외출장" };
  if (type.includes("국내출장")) return { category: "출장", subType: "국내출장" };
  if (type.includes("출장")) return { category: "출장", subType: rawType };

  if (type.includes("전년도") && type.includes("대체휴무")) return { category: "휴가", subType: "전년도대체휴무" };
  if (!type.includes("전년도") && type.includes("대체휴무")) return { category: "휴가", subType: "당해대체휴무" };
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

  const overtimeHours = parseOvertimeHours(row["시간외관리"] ?? row["시간외(시간)"] ?? row["시간외"]);
  if (overtimeHours > 0) {
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "시간외", subType: "시간외근무", overtimeHours, durationHours: 0 });
  }

  const earlyLeaveHours = parseHourValue(row["조기퇴근"] ?? row["조퇴"]);
  if (earlyLeaveHours > 0) {
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "조퇴", subType: "조기퇴근", overtimeHours: 0, durationHours: earlyLeaveHours });
  }

  const leaveCell = row["휴가관리"] ?? row["근태유형"] ?? row["유형"];
  const leaveDurationHours = parseLeaveDurationHours(leaveCell);
  const leaveTypes = extractTypeList(leaveCell);
  leaveTypes.forEach((t) => {
    const c = classifyType(t);
    results.push({
      month: sheetName,
      name,
      employeeId,
      date: dateIso,
      category: c.category,
      subType: c.subType,
      overtimeHours: 0,
      durationHours: c.category === "휴가" || c.category === "병가" || c.category === "산전후휴가" || c.category === "임산부정기검진" ? leaveDurationHours : 0,
    });
  });

  const tripTypes = extractTypeList(row["출장관리"]);
  tripTypes.forEach((t) => {
    const c = classifyType(t);
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: c.category, subType: c.subType, overtimeHours: 0, durationHours: 0 });
  });

  return results;
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const allRecords = [];

  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    rows.forEach((row) => allRecords.push(...parseRow(row, sheetName)));
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
    vacationHours: 0,
    sickLeaveHours: 0,
    earlyLeaveHours: 0,
    domesticTrip: 0,
    internationalTrip: 0,
    localTrip: 0,
    outsideTrip: 0,
    maternityLeaveHours: 0,
    pregnancyCheckupHours: 0,
    compThisYearHours: 0,
    compPrevYearHours: 0,
  };
}

function toDay(hours) {
  return hours / HOURS_PER_DAY;
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
    if (r.category === "휴가") summary.vacationHours += r.durationHours;
    if (r.category === "병가") summary.sickLeaveHours += r.durationHours;
    if (r.category === "조퇴") summary.earlyLeaveHours += r.durationHours;
    if (r.category === "산전후휴가") summary.maternityLeaveHours += r.durationHours;
    if (r.category === "임산부정기검진") summary.pregnancyCheckupHours += r.durationHours;

    if (r.subType === "당해대체휴무") summary.compThisYearHours += r.durationHours;
    if (r.subType === "전년도대체휴무") summary.compPrevYearHours += r.durationHours;

    if (r.subType === "국내출장") summary.domesticTrip += 1;
    if (r.subType === "국외출장") summary.internationalTrip += 1;
    if (r.subType === "관내출장") summary.localTrip += 1;
    if (r.subType === "관외출장") summary.outsideTrip += 1;
  });
  return summary;
}

function formatDays(hours) {
  return toDay(hours).toFixed(2);
}

function renderDetails(records) {
  if (!records.length) {
    detailRows.innerHTML = '<tr><td colspan="4" class="empty">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  detailRows.innerHTML = records
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
    .map(
      (r) =>
        `<tr><td>${r.date}</td><td>${r.category}</td><td>${r.subType}</td><td>${r.overtimeHours > 0 ? r.overtimeHours.toFixed(2) : "-"}</td></tr>`,
    )
    .join("");
}

function updateDashboard() {
  const month = monthSelect.value;
  const name = employeeSelect.value;
  const employeeId = state.employeesByMonth.get(month)?.get(name) || "-";
  const records = getSelectedRecords();
  const summary = buildSummary(records);

  empNoEl.textContent = employeeId;
  overtimeEl.textContent = summary.overtime.toFixed(2);
  vacationEl.textContent = formatDays(summary.vacationHours);
  sickLeaveEl.textContent = formatDays(summary.sickLeaveHours);
  earlyLeaveEl.textContent = formatDays(summary.earlyLeaveHours);
  domesticTripEl.textContent = String(summary.domesticTrip);
  internationalTripEl.textContent = String(summary.internationalTrip);
  localTripEl.textContent = String(summary.localTrip);
  outsideTripEl.textContent = String(summary.outsideTrip);
  maternityLeaveEl.textContent = formatDays(summary.maternityLeaveHours);
  pregnancyCheckupEl.textContent = formatDays(summary.pregnancyCheckupHours);
  compThisYearEl.textContent = formatDays(summary.compThisYearHours);
  compPrevYearEl.textContent = formatDays(summary.compPrevYearHours);

  renderDetails(records);
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ records: state.records, months: state.months }));
}

function loadSavedState() {
  const raw = localStorage.getItem(STORAGE_KEY);
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

async function fetchWorkbookBuffer(downloadUrl) {
  const targets = [
    downloadUrl,
    `https://cors.isomorphic-git.org/${downloadUrl}`,
    `https://corsproxy.io/?${encodeURIComponent(downloadUrl)}`,
  ];

  let lastError = null;
  for (const url of targets) {
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await response.arrayBuffer();
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`엑셀 다운로드 실패 (${lastError?.message ?? "알 수 없는 오류"})`);
}

async function loadFromDrive() {
  const downloadUrl = convertDriveUrl(DEFAULT_DRIVE_XLSX_URL);
  const buffer = await fetchWorkbookBuffer(downloadUrl);
  parseWorkbook(buffer);
  saveState();
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
    alert(`불러오기에 실패했습니다.\n1) 시트 공유 권한(링크 보기 가능)\n2) 사내망에서 외부 주소/프록시 접근 허용 여부\n3) 가능하면 로컬/사내 서버(http) 실행\n을 확인하세요.\n\n상세 오류: ${error.message}`);
  }
});

clearBtn.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  state.records = [];
  state.months = [];
  state.employeesByMonth.clear();
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

if (!state.records.length) {
  loadFromDrive().catch((error) => {
    alert(`초기 데이터 로딩에 실패했습니다.\n네트워크 정책 또는 공유권한 문제일 수 있습니다.\n상세 오류: ${error.message}`);
  });
}
