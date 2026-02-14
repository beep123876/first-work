const STORAGE_KEY = "attendanceDashboardData_v5";
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
const sickLeaveEl = document.getElementById("sickLeave");
const earlyLeaveEl = document.getElementById("earlyLeave");
const domesticTripEl = document.getElementById("domesticTrip");
const internationalTripEl = document.getElementById("internationalTrip");
const localTripEl = document.getElementById("localTrip");
const outsideTripEl = document.getElementById("outsideTrip");
const maternityLeaveEl = document.getElementById("maternityLeave");
const pregnancyCheckupEl = document.getElementById("pregnancyCheckup");
const dayOffEl = document.getElementById("dayOff");
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

function parseDurationFromText(text) {
  const src = String(text ?? "");
  const days = Number(src.match(/(\d+)\s*일/)?.[1] ?? 0);
  const hours = Number(src.match(/(\d+)\s*시간/)?.[1] ?? 0);
  const mins = Number(src.match(/(\d+)\s*분/)?.[1] ?? 0);
  return days * HOURS_PER_DAY + hours + mins / 60;
}

function formatHoursLabel(hours) {
  const totalMin = Math.round((Number(hours) || 0) * 60);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  if (h > 0 && m > 0) return `${h}시간 ${m}분`;
  if (h > 0) return `${h}시간`;
  return `${m}분`;
}

function parseOvertimeHours(cellText) {
  if (typeof cellText === "number") return cellText;
  const text = String(cellText ?? "");
  if (!text.trim()) return 0;

  const labeled = text.match(/(?:총시간|총시간외|시간외시간|실근무)\s*[:：]\s*([^\n\r]+)/);
  if (labeled) return parseHourValue(labeled[1]);
  if (text.includes("신청시각") || text.includes("실근무") || text.includes("종별")) return 0;
  return parseHourValue(text);
}

function parseDateTimeFromLine(line) {
  const m = line.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]));
}

function calcTripHours(fromLine, toLine) {
  const from = parseDateTimeFromLine(fromLine);
  const to = parseDateTimeFromLine(toLine);
  if (!from || !to) return 0;
  const diff = (to.getTime() - from.getTime()) / (1000 * 60 * 60);
  return diff > 0 ? diff : 0;
}

function classifyLeaveType(typeRaw, durationHours) {
  const type = typeRaw.replace(/\s+/g, "");
  if (type.includes("병가")) return { category: "병가", subType: typeRaw };
  if (type.includes("산전후휴가")) return { category: "산전후휴가", subType: typeRaw };
  if (type.includes("임산부정기검진")) return { category: "임산부정기검진", subType: typeRaw };
  if (type.includes("조퇴")) return { category: "조퇴", subType: typeRaw };
  if (type.includes("대체휴무") || type.includes("휴가") || type.includes("휴무")) {
    const subType = durationHours > 0 ? `휴무(${formatHoursLabel(durationHours)})` : "휴무";
    return { category: "휴무", subType, isDayOff: true };
  }
  return { category: "기타", subType: typeRaw };
}

function classifyTripType(rawType) {
  const type = rawType.replace(/\s+/g, "");
  if (type.includes("관내출장")) return { category: "출장", subType: "관내출장" };
  if (type.includes("관외출장")) return { category: "출장", subType: "관외출장" };
  if (type.includes("국외출장")) return { category: "출장", subType: "국외출장" };
  if (type.includes("국내출장")) return { category: "출장", subType: "국내출장" };
  if (type.includes("출장")) return { category: "출장", subType: rawType };
  return { category: "기타", subType: rawType };
}

function parseLeaveEntries(cellText) {
  const text = String(cellText ?? "").trim();
  if (!text) return [];

  const lines = text.split(/\r?\n/).map((v) => v.trim()).filter(Boolean);
  const entries = [];
  let current = null;

  lines.forEach((line) => {
    const typeMatch = line.match(/^종별\s*[:：]\s*(.+)$/);
    if (typeMatch) {
      if (current) entries.push(current);
      current = { typeRaw: typeMatch[1].trim(), durationHours: 0 };
      return;
    }

    const durationMatch = line.match(/^일수\/시간\s*[:：]\s*(.+)$/);
    if (durationMatch && current) current.durationHours = parseDurationFromText(durationMatch[1]);
  });

  if (current) entries.push(current);
  return entries;
}

function parseTripEntries(cellText) {
  const text = String(cellText ?? "").trim();
  if (!text) return [];

  const lines = text.split(/\r?\n/).map((v) => v.trim()).filter(Boolean);
  const entries = [];
  let current = null;

  lines.forEach((line) => {
    const typeMatch = line.match(/^종별\s*[:：]\s*(.+)$/);
    if (typeMatch) {
      if (current) entries.push(current);
      current = { typeRaw: typeMatch[1].trim(), fromLine: "", toLine: "" };
      return;
    }

    if (!current) return;
    if (/^부터\s*[:：]\s*/.test(line)) current.fromLine = line;
    if (/^까지\s*[:：]\s*/.test(line)) current.toLine = line;
  });

  if (current) entries.push(current);
  return entries;
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
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "시간외", subType: "시간외근무", overtimeHours, durationHours: 0, isDayOff: false });
  }

  const earlyLeaveHours = parseHourValue(row["조기퇴근"] ?? row["조퇴"]);
  if (earlyLeaveHours > 0) {
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: "조퇴", subType: `조기퇴근(${formatHoursLabel(earlyLeaveHours)})`, overtimeHours: 0, durationHours: earlyLeaveHours, isDayOff: false });
  }

  const leaveEntries = parseLeaveEntries(row["휴가관리"] ?? row["근태유형"] ?? row["유형"]);
  leaveEntries.forEach((entry) => {
    const c = classifyLeaveType(entry.typeRaw, entry.durationHours);
    results.push({
      month: sheetName,
      name,
      employeeId,
      date: dateIso,
      category: c.category,
      subType: c.subType,
      overtimeHours: 0,
      durationHours: entry.durationHours,
      isDayOff: Boolean(c.isDayOff),
    });
  });

  const tripEntries = parseTripEntries(row["출장관리"]);
  tripEntries.forEach((entry) => {
    const c = classifyTripType(entry.typeRaw);
    const tripHours = calcTripHours(entry.fromLine, entry.toLine);
    const subType = tripHours > 0 ? `${c.subType}(${formatHoursLabel(tripHours)})` : c.subType;
    results.push({ month: sheetName, name, employeeId, date: dateIso, category: c.category, subType, overtimeHours: 0, durationHours: tripHours, isDayOff: false });
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
    dayOffHours: 0,
    sickLeaveHours: 0,
    earlyLeaveHours: 0,
    domesticTrip: 0,
    internationalTrip: 0,
    localTrip: 0,
    outsideTrip: 0,
    maternityLeaveHours: 0,
    pregnancyCheckupHours: 0,
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
    if (r.category === "휴무") summary.dayOffHours += r.durationHours;
    if (r.category === "병가") summary.sickLeaveHours += r.durationHours;
    if (r.category === "조퇴") summary.earlyLeaveHours += r.durationHours;
    if (r.category === "산전후휴가") summary.maternityLeaveHours += r.durationHours;
    if (r.category === "임산부정기검진") summary.pregnancyCheckupHours += r.durationHours;

    if (r.subType.startsWith("국외출장")) summary.internationalTrip += 1;
    if (r.subType.startsWith("관내출장")) {
      summary.localTrip += 1;
      summary.domesticTrip += 1;
    }
    if (r.subType.startsWith("관외출장")) {
      summary.outsideTrip += 1;
      summary.domesticTrip += 1;
    }
    if (r.subType.startsWith("국내출장")) summary.domesticTrip += 1;
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
    .map((r) => `<tr><td>${r.date}</td><td>${r.category}</td><td>${r.subType}</td><td>${r.overtimeHours > 0 ? r.overtimeHours.toFixed(2) : "-"}</td></tr>`)
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
  dayOffEl.textContent = formatDays(summary.dayOffHours);
  sickLeaveEl.textContent = formatDays(summary.sickLeaveHours);
  earlyLeaveEl.textContent = formatDays(summary.earlyLeaveHours);
  domesticTripEl.textContent = String(summary.domesticTrip);
  internationalTripEl.textContent = String(summary.internationalTrip);
  localTripEl.textContent = String(summary.localTrip);
  outsideTripEl.textContent = String(summary.outsideTrip);
  maternityLeaveEl.textContent = formatDays(summary.maternityLeaveHours);
  pregnancyCheckupEl.textContent = formatDays(summary.pregnancyCheckupHours);

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

async function fetchWithTimeout(url, timeoutMs = 12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

function buildFetchTargets(downloadUrl) {
  const noProto = downloadUrl.replace(/^https?:\/\//, "");
  return [
    { label: "direct", url: downloadUrl },
    { label: "cors.isomorphic", url: `https://cors.isomorphic-git.org/${downloadUrl}` },
    { label: "corsproxy", url: `https://corsproxy.io/?${encodeURIComponent(downloadUrl)}` },
    { label: "allorigins", url: `https://api.allorigins.win/raw?url=${encodeURIComponent(downloadUrl)}` },
    { label: "jina", url: `https://r.jina.ai/http://${noProto}` },
  ];
}

async function fetchWorkbookBuffer(downloadUrl) {
  const targets = buildFetchTargets(downloadUrl);
  let lastError = null;

  for (const target of targets) {
    try {
      const response = await fetchWithTimeout(target.url);
      if (!response.ok) throw new Error(`${target.label}: HTTP ${response.status}`);
      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("text/html")) throw new Error(`${target.label}: HTML 응답(다운로드 링크/권한 확인 필요)`);
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
