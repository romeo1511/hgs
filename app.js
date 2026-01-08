
// DOM Elements
const dropArea = document.getElementById('drop-area');
const fileInput = document.createElement('input'); // Helper for click
fileInput.type = 'file';
fileInput.accept = '.xlsx, .xls';
fileInput.style.display = 'none';
document.body.appendChild(fileInput);

const fileInfo = document.getElementById('file-info');
const filenameDisplay = document.getElementById('filename');
const controlPanel = document.getElementById('control-panel');
const sheetSelect = document.getElementById('sheet-select');
const resultSection = document.getElementById('result-section');

// Mode Toggle Elements
const modeButtons = document.querySelectorAll('.mode-btn');
const sheetModeControls = document.getElementById('sheet-mode-controls');
const dayModeControls = document.getElementById('day-mode-controls');
const daySelect = document.getElementById('day-select');

// Search and Selection elements
const staffSearch = document.getElementById('staff-search');
const staffResults = document.getElementById('staff-results');
const staffSelect = document.getElementById('staff-select');

const dayStaffSearch = document.getElementById('day-staff-search');
const dayStaffResults = document.getElementById('day-staff-results');
const dayStaffSelect = document.getElementById('day-staff-select');

const availabilityModeControls = document.getElementById('availability-mode-controls');
const availabilityView = document.getElementById('availability-view');
const availDaySelect = document.getElementById('avail-day-select');
const availTimeStart = document.getElementById('avail-time-start');
const availTimeEnd = document.getElementById('avail-time-end');
const searchAvailBtn = document.getElementById('search-avail-btn');
const availList = document.getElementById('avail-list');
const availResultsTitle = document.getElementById('avail-results-title');

// Views
const attendanceView = document.getElementById('attendance-view');
const flightView = document.getElementById('flight-view');
const dayView = document.getElementById('day-view');

// Attendance Elements
const totalHoursEl = document.getElementById('total-hours');
const totalShiftsEl = document.getElementById('total-shifts');
const detailsTableBody = document.querySelector('#details-table tbody');

// Flight Elements
const flightsTableBody = document.querySelector('#flights-table tbody');

// Day View Elements
const shiftBadge = document.getElementById('shift-badge');
const shiftTime = document.getElementById('shift-time');
const shiftDuration = document.getElementById('shift-duration');
const flightTimeline = document.getElementById('flight-timeline');

const debugOutput = document.getElementById('debug-output');

// Global Data State
let workbook = null;
let currentSheetData = null; // Array of Arrays
let headerRowIndex = -1;
let staffMap = new Map(); // Name -> Row Data (For Attendance) OR Set of Names (For Flights)
let currentSheetType = 'UNKNOWN'; // 'QT' or 'FLIGHT'
let currentMode = 'sheet'; // 'sheet' or 'day'

// Cached QT Sheet Data (for day-based lookup)
let qtSheetData = null;
let qtStaffMap = new Map(); // Name -> Row Data from QT sheet
let cachedSheets = {}; // sheetName -> daySheetData (Array of Arrays)

// Shift Definitions
const BASE_SHIFTS = {
    'HC': { start: 7 * 60 + 30, end: 15 * 60 + 30, duration: 8 },
    'S': { start: 6 * 60, end: 14 * 60, duration: 8 },
    'X': { start: 14 * 60, end: 22 * 60, duration: 8 },
    'A': { start: 22 * 60, end: 6 * 60, duration: 8, nextDay: true },

    // Non-working codes
    'OFF': { start: 0, end: 0, duration: 0 },
    'LS': { start: 0, end: 0, duration: 0 },
    'LC': { start: 0, end: 0, duration: 0 },
    'H3': { start: 0, end: 0, duration: 0 },
    'NB': { start: 0, end: 0, duration: 0 },
    'CO': { start: 0, end: 0, duration: 0 }
};

// --- Event Listeners ---

/**
 * Normalizes Vietnamese strings by removing diacritics (accents)
 * @param {string} str - The string to normalize
 * @returns {string} Normalized string
 */
function removeVietnameseTones(str) {
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
    str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
    str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
    str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
    str = str.replace(/ỳ|ý|ỵ|ỷ|ĩ/g, "y");
    str = str.replace(/đ/g, "d");
    str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
    str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
    str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
    str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
    str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
    str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
    str = str.replace(/Đ/g, "D");
    // Some system encode vietnamese combining accent as individual utf-8 characters
    str = str.replace(/\u0300|\u0301|\u0303|\u0309|\u0323/g, ""); // Huyền sắc hỏi ngã nặng 
    str = str.replace(/\u02C6|\u0306|\u031B/g, ""); // Â, Ă, Ơ, Ư
    return str;
}

/**
 * Initialize a searchable selection component
 * @param {HTMLInputElement} inputEl - The search text input
 * @param {HTMLDivElement} resultsEl - The results list container
 * @param {HTMLSelectElement} selectEl - The hidden actual select element
 * @param {Function} onSelect - Callback when an item is selected
 */
function initSearchBox(inputEl, resultsEl, selectEl, onSelect) {
    let currentStaffNames = [];

    // Helper to refresh names from the underlying select
    const refreshNames = () => {
        currentStaffNames = Array.from(selectEl.options)
            .map(opt => opt.value)
            .filter(v => v !== "");
    };

    // Filter and show results
    const filterResults = (query) => {
        refreshNames();
        const normalizedQuery = removeVietnameseTones(query.toLowerCase());

        const matches = currentStaffNames.filter(name => {
            const normalizedName = removeVietnameseTones(name.toLowerCase());
            return normalizedName.includes(normalizedQuery);
        });

        if (matches.length > 0) {
            resultsEl.innerHTML = matches.map(name =>
                `<div class="search-result-item" data-value="${name}">${name}</div>`
            ).join('');
            resultsEl.classList.remove('hidden');
        } else if (query.trim() !== '') {
            resultsEl.innerHTML = '<div class="search-result-item no-results">Không tìm thấy kết quả</div>';
            resultsEl.classList.remove('hidden');
        } else {
            resultsEl.classList.add('hidden');
        }
    };

    inputEl.addEventListener('input', (e) => filterResults(e.target.value));

    inputEl.addEventListener('focus', (e) => {
        if (e.target.value.trim() !== '') filterResults(e.target.value);
    });

    // Close on blur (delayed to allow clicks)
    inputEl.addEventListener('blur', () => {
        setTimeout(() => resultsEl.classList.add('hidden'), 200);
    });

    resultsEl.addEventListener('click', (e) => {
        const item = e.target.closest('.search-result-item');
        if (item && !item.classList.contains('no-results')) {
            const val = item.dataset.value;
            inputEl.value = val;
            selectEl.value = val;
            resultsEl.classList.add('hidden');
            onSelect(val);
        }
    });
}

// Initialize search for both modes
initSearchBox(staffSearch, staffResults, staffSelect, (val) => {
    if (currentSheetType === 'QT') {
        calculateAttendanceStats(val);
    } else if (currentSheetType === 'FLIGHT') {
        findFlights(val);
    }
});

initSearchBox(dayStaffSearch, dayStaffResults, dayStaffSelect, (val) => {
    if (daySelect.value) {
        displayDayView(val, daySelect.value);
    }
});

dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.style.borderColor = '#6366f1';
    dropArea.style.background = 'rgba(99, 102, 241, 0.1)';
});

dropArea.addEventListener('dragleave', () => {
    dropArea.style.borderColor = 'rgba(255, 255, 255, 0.1)';
    dropArea.style.background = 'rgba(255, 255, 255, 0.05)';
});

dropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dropArea.style.borderColor = 'rgba(255, 255, 255, 0.1)';
    dropArea.style.background = 'rgba(255, 255, 255, 0.05)';
    const files = e.dataTransfer.files;
    if (files.length) handleFile(files[0]);
});

const visibleInput = document.getElementById('file-input');
visibleInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

sheetSelect.addEventListener('change', (e) => {
    if (e.target.value) {
        processSheet(e.target.value);
        staffSearch.value = ''; // Reset search on sheet change
    } else {
        staffSearch.disabled = true;
        staffSelect.innerHTML = '<option value="">-- Trước tiên chọn Sheet --</option>';
        resultSection.classList.add('hidden');
    }
});

// Mode Toggle Event Listeners
modeButtons.forEach(btn => {
    btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        currentMode = mode;

        // Update button states
        modeButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        // Hide everything first
        sheetModeControls.classList.add('hidden');
        dayModeControls.classList.add('hidden');
        availabilityModeControls.classList.add('hidden');
        attendanceView.classList.add('hidden');
        flightView.classList.add('hidden');
        dayView.classList.add('hidden');
        availabilityView.classList.add('hidden');
        resultSection.classList.add('hidden'); // Hide result section by default

        if (mode === 'sheet') {
            sheetModeControls.classList.remove('hidden');
            if (staffSelect.value) {
                resultSection.classList.remove('hidden');
                if (currentSheetType === 'QT') {
                    attendanceView.classList.remove('hidden');
                } else if (currentSheetType === 'FLIGHT') {
                    flightView.classList.remove('hidden');
                }
            }
        } else if (mode === 'day') {
            dayModeControls.classList.remove('hidden');
            if (daySelect.value && dayStaffSelect.value) {
                resultSection.classList.remove('hidden');
                dayView.classList.remove('hidden');
            }
        } else if (mode === 'availability') {
            availabilityModeControls.classList.remove('hidden');
            // Availability view will be shown after search
        }
    });
});

// Day Select Event Listener
daySelect.addEventListener('change', (e) => {
    if (e.target.value) {
        populateDayStaffList();
        dayStaffSearch.disabled = false;
        dayStaffSearch.value = '';
    } else {
        dayStaffSearch.disabled = true;
        dayStaffSearch.value = '';
        dayStaffSelect.innerHTML = '<option value="">-- Trước tiên chọn Ngày --</option>';
        resultSection.classList.add('hidden');
    }
});


// --- Core Functions ---

function handleFile(file) {
    filenameDisplay.textContent = file.name;
    fileInfo.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        cachedSheets = {}; // Clear old cache
        workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            cachedSheets[name] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        controlPanel.classList.remove('hidden');

        // Cache QT Sheet Data for day-based lookups
        const qtSheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('CHẤM CÔNG QT'));
        if (qtSheetName) {
            const qtSheet = workbook.Sheets[qtSheetName];
            qtSheetData = XLSX.utils.sheet_to_json(qtSheet, { header: 1, defval: '' });
            cacheQTStaffData();

            // Sync the main attendee search box
            populateStaffListFromQT();
            if (typeof staffSearch !== 'undefined' && typeof staffSearch.refreshNames === 'function') {
                staffSearch.refreshNames();
            }

            logDebug(`Cached QT sheet: ${qtSheetName}`);
        }

        // Auto-select priority: "BẢNG CHẤM CÔNG QT"
        const likelySheet = workbook.SheetNames.find(n => n.toUpperCase().includes('CHẤM CÔNG QT'));

        if (likelySheet) {
            sheetSelect.value = likelySheet;
            processSheet(likelySheet);
        } else if (workbook.SheetNames.length > 0) {
            const firstSheet = workbook.SheetNames[0];
            sheetSelect.value = firstSheet;
            processSheet(firstSheet);
        }

        logDebug(`File loaded. Sheets: ${workbook.SheetNames.join(', ')}`);
    };
    reader.readAsArrayBuffer(file);
}

function processSheet(sheetName) {
    currentSheetData = cachedSheets[sheetName];

    // Determine Sheet Type
    // QT rule: Contains "CHẤM CÔNG QT" in name OR specific structure
    if (sheetName.toUpperCase().includes('QT') || sheetName.toUpperCase().includes('CHẤM CÔNG')) {
        currentSheetType = 'QT';
        if (currentMode === 'sheet') {
            attendanceView.classList.remove('hidden');
            flightView.classList.add('hidden');
            dayView.classList.add('hidden');
        }
        processQTSheet(sheetName);
    } else {
        // Assume Flight Schedule (Daily)
        currentSheetType = 'FLIGHT';
        if (currentMode === 'sheet') {
            attendanceView.classList.add('hidden');
            flightView.classList.remove('hidden');
            dayView.classList.add('hidden');
        }
        processFlightSheet(sheetName);
    }
}

// --- QT Sheet Logic ---
function processQTSheet(sheetName) {
    headerRowIndex = 2; // Fixed Row 3 for QT
    const headerRow = currentSheetData[headerRowIndex] || [];

    let nameColIndex = 1; // Col B

    staffMap.clear();
    staffSelect.innerHTML = '<option value="">-- Chọn Nhân viên --</option>';

    for (let i = headerRowIndex + 1; i < currentSheetData.length; i++) {
        const row = currentSheetData[i];
        const rawName = row[nameColIndex];

        if (rawName && typeof rawName === 'string' && rawName.trim().length > 0) {
            const cleanName = rawName.trim();
            if (cleanName.toLowerCase().includes('ngày')) continue;
            staffMap.set(cleanName, row);
            const option = document.createElement('option');
            option.value = cleanName;
            option.textContent = cleanName;
            staffSelect.appendChild(option);
        }
    }

    staffSelect.disabled = false;
    resultSection.classList.add('hidden'); // waiting for selection
}

// --- Flight Schedule Logic ---
function processFlightSheet(sheetName) {
    // Strategy: Look for header row with "Flight", "SH", "Số hiệu"
    headerRowIndex = -1;
    for (let i = 0; i < Math.min(20, currentSheetData.length); i++) {
        const rowStr = JSON.stringify(currentSheetData[i]).toLowerCase();
        if (rowStr.includes('flight') || rowStr.includes('số hiệu') || rowStr.includes('gate') || rowStr.includes('etd')) {
            headerRowIndex = i;
            break;
        }
    }

    if (headerRowIndex === -1) {
        logDebug("Could not find Flight Header. Assuming Row 1 (Index 0).");
        headerRowIndex = 0;
    }

    // Collect ALL unique staff names from the grid
    const staffSet = new Set();
    staffSelect.innerHTML = '<option value="">-- Chọn Nhân viên --</option>';

    // Loop through all data rows
    for (let i = headerRowIndex + 1; i < currentSheetData.length; i++) {
        const row = currentSheetData[i];
        row.forEach(cell => {
            if (cell && typeof cell === 'string') {
                // Heuristic: Names usually > 3 chars, no numbers (mostly)
                // Filter out common keywords
                if (cell.length > 2 && !['HC', 'OFF', 'GATE'].includes(cell.toUpperCase())) {
                    staffSet.add(cell.trim());
                }
            }
        });
    }

    // Sort and populate
    const sortedStaff = Array.from(staffSet).sort();
    sortedStaff.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        staffSelect.appendChild(option);
    });

    staffSelect.disabled = false;
    resultSection.classList.add('hidden');
}

// --- Attendance Calculation (QT) ---
function parseShift(code) {
    if (!code || typeof code !== 'string') return null;
    code = code.toUpperCase().trim();

    if (code === 'H') return { start: 0, end: 0, duration: 0, isDirect: true, label: 'Học' };
    if (code === 'M') return { start: 0, end: 0, duration: 0, isDirect: true, label: 'Họp' };

    const directMatch = code.match(/^([HM])(\d+(\.\d+)?)([A-Z]*)?$/);
    if (directMatch && !code.startsWith('HC') && !code.startsWith('MC')) {
        const type = directMatch[1];
        const duration = parseFloat(directMatch[2]);
        const label = type === 'H' ? 'Học' : 'Họp';
        return { start: 0, end: duration * 60, duration: duration, isDirect: true, label: `${label} ${duration}h` };
    }

    const regex = /^([\d.]+-?)?([A-Z]+)(-?[\d.]+)?$/;
    const match = code.match(regex);
    if (!match) return null;

    // ... (Same logic as before) ...
    const prefix = match[1];
    const baseCode = match[2];
    const suffix = match[3];

    const baseShift = BASE_SHIFTS[baseCode];
    if (!baseShift) return null; // Only strict base codes

    let startMin = baseShift.start;
    let endMin = baseShift.end;

    if (prefix) {
        if (prefix.includes('-')) startMin += parseFloat(prefix.replace('-', '')) * 60;
        else startMin -= parseFloat(prefix) * 60;
    }
    if (suffix) {
        if (suffix.includes('-')) endMin -= parseFloat(suffix.replace('-', '')) * 60;
        else endMin += parseFloat(suffix) * 60;
    }

    let duration = (endMin - startMin) / 60;
    if (duration < 0) duration += 24;

    return { start: startMin, end: endMin, duration: duration };
}

function timeStr(minutes) {
    let m = minutes;
    while (m < 0) m += 1440;
    while (m >= 1440) m -= 1440;
    const h = Math.floor(m / 60);
    const min = Math.floor(m % 60);
    return `${h.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}`;
}

/**
 * Format Excel numeric time values to HH:mm string
 * Handles numeric values (including next day > 1), Date objects, and strings
 * @param {any} val - Value from Excel cell
 * @returns {string} Formatted time string
 */
function formatExcelTime(val) {
    if (val === null || val === undefined || val === '') return '-';

    // If it's already a string and looks like time HH:mm, HH:mm:ss, return as is
    if (typeof val === 'string') {
        const timeMatch = val.match(/(\d{1,2}:\d{2})(:\d{2})?/); // Look for HH:mm anywhere
        if (timeMatch) return timeMatch[1];
        return val;
    }

    // Handle Date object
    if (val instanceof Date || (typeof val === 'object' && typeof val.getHours === 'function')) {
        const h = val.getHours();
        const m = val.getMinutes();
        return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
    }

    // Handle Excel numeric time
    if (typeof val === 'number') {
        // Excel stores time as a fraction of a day (1 = 24h)
        // Values > 1 mean the next day(s)
        const totalMinutes = Math.round(val * 1440);
        if (isNaN(totalMinutes)) return '-';

        const h = Math.floor(totalMinutes / 60) % 24;
        const m = totalMinutes % 60;

        // If it's strictly > 1 day, user might want (+1) or similar, but for now HH:mm
        // The % 24 handles "25:00" as "01:00"
        return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
    }

    return val.toString();
}

function calculateAttendanceStats(staffName) {
    const row = staffMap.get(staffName);
    if (!row) return;

    let totalHours = 0;
    let shiftCount = 0;
    let detailsHtml = '';

    const days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'CN'];
    const startCol = 2; // Col C

    days.forEach((dayName, idx) => {
        const baseColIdx = startCol + (idx * 3);
        let combinedTokens = [];
        let displayValues = [];

        for (let offset = 0; offset < 3; offset++) {
            const val = row[baseColIdx + offset];
            if (val) {
                const sStr = val.toString();
                displayValues.push(sStr);
                const cellTokens = sStr.split(/[\s,+\n]+/);
                combinedTokens.push(...cellTokens);
            }
        }

        let dailyHours = 0;
        let validShifts = [];
        let unknownTokens = [];
        let isOff = false;

        combinedTokens.forEach(token => {
            if (!token.trim()) return;
            const res = parseShift(token);
            if (res && (res.duration > 0 || res.isDirect)) {
                dailyHours += res.duration;
                validShifts.push({ code: token, ...res });
            } else {
                const upper = token.toUpperCase();
                if (['OFF', 'LS', 'LC', 'NB', 'CO'].includes(upper)) {
                    isOff = true;
                } else {
                    unknownTokens.push(token);
                }
            }
        });

        if (validShifts.length > 0) {
            totalHours += dailyHours;
            shiftCount += validShifts.filter(s => s.duration > 0).length;
            const timeDetails = validShifts.map(s => {
                if (s.isDirect) return `[${s.code}] ${s.label}`;
                return `[${s.code}] ${timeStr(s.start)}-${timeStr(s.end)}`;
            }).join('<br>');
            detailsHtml += `<tr><td>${dayName}</td><td><span class="badge">${displayValues.join(' | ') || ''}</span></td><td colspan="2">${timeDetails}</td><td>${dailyHours}h</td></tr>`;
        } else if (isOff) {
            detailsHtml += `<tr><td>${dayName}</td><td><span class="badge" style="background:rgba(255,255,255,0.1)">${displayValues.join(' | ')}</span></td><td colspan="3" style="color:var(--text-muted)">Nghỉ/Không tính giờ</td></tr>`;
        } else if (unknownTokens.length > 0) {
            detailsHtml += `<tr><td>${dayName}</td><td colspan="4" style="color:red">Mã lạ: ${displayValues.join(' | ')}</td></tr>`;
        }
    });

    totalHoursEl.textContent = totalHours;
    totalShiftsEl.textContent = shiftCount;
    detailsTableBody.innerHTML = detailsHtml;
    resultSection.classList.remove('hidden');
}


// --- Flight Finding Logic ---
function findFlights(staffName) {
    if (!currentSheetData || !staffName) return;

    const headerRow = currentSheetData[headerRowIndex];

    // Identify Columns
    let flightColIdx = headerRow.findIndex(c => c && c.toString().match(/flight|số hiệu|sh/i));
    if (flightColIdx === -1) flightColIdx = 0; // Default A

    let etdColIdx = headerRow.findIndex(c => c && c.toString().match(/etd|giờ|time/i));
    if (etdColIdx === -1) etdColIdx = 1; // Default B

    let gateColIdx = headerRow.findIndex(c => c && c.toString().match(/gate|cửa/i));

    let html = '';

    for (let i = headerRowIndex + 1; i < currentSheetData.length; i++) {
        const row = currentSheetData[i];

        // Scan row for staffName
        let foundColIdx = row.indexOf(staffName);
        // Also try case-insensitive or trimmed match if exact fails
        if (foundColIdx === -1) {
            foundColIdx = row.findIndex(c => c && c.toString().trim() === staffName.trim());
        }

        if (foundColIdx !== -1) {
            // Found !
            const flight = row[flightColIdx] || '-';
            const etd = formatExcelTime(row[etdColIdx]);
            const position = headerRow[foundColIdx] || 'Unknown';
            const gate = (gateColIdx !== -1) ? row[gateColIdx] : '';

            html += `
                <tr>
                    <td><strong>${flight}</strong></td>
                    <td>${etd}</td>
                    <td><span class="badge">${position}</span></td>
                    <td>${gate}</td>
                </tr>
            `;
        }
    }

    if (html === '') {
        html = '<tr><td colspan="4" style="text-align:center">Không tìm thấy chuyến bay nào cho nhân viên này.</td></tr>';
    }

    flightsTableBody.innerHTML = html;
    resultSection.classList.remove('hidden');
}


// --- Day-Based Lookup Functions ---

/**
 * Cache QT Staff Data into qtStaffMap
 */
function cacheQTStaffData() {
    if (!qtSheetData) return;

    qtStaffMap.clear();
    const headerRowIndex = 2; // Row 3 (index 2)
    const nameColIndex = 1; // Col B

    for (let i = headerRowIndex + 1; i < qtSheetData.length; i++) {
        const row = qtSheetData[i];
        const rawName = row[nameColIndex];

        if (rawName && typeof rawName === 'string' && rawName.trim().length > 0) {
            const cleanName = rawName.trim();
            if (cleanName.toLowerCase().includes('ngày')) continue;
            qtStaffMap.set(cleanName, row);
        }
    }

    logDebug(`Cached ${qtStaffMap.size} staff from QT sheet`);
}

/**
 * Populate day staff list from QT sheet
 */
function populateDayStaffList() {
    dayStaffSelect.innerHTML = '<option value="">-- Chọn Nhân viên --</option>';

    if (qtStaffMap.size === 0) {
        dayStaffSelect.innerHTML = '<option value="">Không tìm thấy dữ liệu QT</option>';
        return;
    }

    const sortedStaff = Array.from(qtStaffMap.keys()).sort();
    sortedStaff.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        dayStaffSelect.appendChild(option);
    });
}

/**
 * Populate middle staff list (Attendance mode) from QT sheet
 */
function populateStaffListFromQT() {
    staffSelect.innerHTML = '<option value="">-- Chọn Nhân viên --</option>';
    if (qtStaffMap.size === 0) return;

    const sortedStaff = Array.from(qtStaffMap.keys()).sort();
    sortedStaff.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        staffSelect.appendChild(option);
    });

    // Refresh the searchable component
    if (staffSearch.refreshNames) staffSearch.refreshNames();
}

/**
 * Get shift information for a staff member on a specific day
 * @param {string} staffName - Name of staff member
 * @param {string} day - Day code (MON, TUE, etc.)
 * @returns {object} Shift info with code, start, end, duration
 */
function getShiftInfo(staffName, day) {
    const row = qtStaffMap.get(staffName);
    if (!row) return null;

    // Map day to column index
    // Days: MON=Thứ 2, TUE=Thứ 3, WED=Thứ 4, THU=Thứ 5, FRI=Thứ 6, SAT=Thứ 7, SUN=CN
    const dayMap = {
        'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4, 'SAT': 5, 'SUN': 6
    };

    const dayIndex = dayMap[day];
    if (dayIndex === undefined) return null;

    const startCol = 2; // Col C
    const baseColIdx = startCol + (dayIndex * 3);

    // Collect shift codes from 3 columns for this day
    let combinedTokens = [];
    for (let offset = 0; offset < 3; offset++) {
        const val = row[baseColIdx + offset];
        if (val) {
            const sStr = val.toString();
            const cellTokens = sStr.split(/[\s,+\n]+/);
            combinedTokens.push(...cellTokens);
        }
    }

    // Parse shift codes
    let totalDuration = 0;
    let validShifts = [];
    let isOff = false;

    combinedTokens.forEach(token => {
        if (!token.trim()) return;
        const res = parseShift(token);
        if (res && (res.duration > 0 || res.isDirect)) {
            totalDuration += res.duration;
            validShifts.push({ code: token, ...res });
        } else {
            const upper = token.toUpperCase();
            if (['OFF', 'LS', 'LC', 'NB', 'CO'].includes(upper)) {
                isOff = true;
            }
        }
    });

    return {
        shifts: validShifts,
        totalDuration,
        isOff,
        rawCodes: combinedTokens.filter(t => t.trim()).join(', ')
    };
}

/**
 * Get flight details for a staff member from a daily sheet
 * @param {string} staffName - Name of staff member
 * @param {Array} daySheetData - Sheet data as array of arrays
 * @returns {Array} Array of flight assignments
 */
function getFlightDetails(staffName, daySheetData) {
    if (!daySheetData || daySheetData.length === 0) return [];

    let flights = [];
    let currentSection = "PHỤC VỤ CHUYẾN BAY";
    let sectionHeaderRow = null;
    let flightColIdx = 12; // Default M
    let etdColIdx = 13;   // Default N
    let gateColIdx = -1;

    // Persist values for merged cells within a section
    let lastFlight = "";
    let lastEtd = "";

    const targetNameNormalized = removeVietnameseTones(staffName.toLowerCase().trim());

    for (let i = 0; i < daySheetData.length; i++) {
        const row = daySheetData[i];
        if (!row || row.length === 0) continue;

        const rowStr = row.join(" ").toUpperCase();

        // 1. Detect Section Anchors
        if (rowStr.includes("BẢNG PHÂN CÔNG CA TRỰC CARGO") || rowStr.includes("CARE CARGO") || rowStr.includes("CHARTER")) {
            currentSection = "CARE CARGO/CHARTER";
            sectionHeaderRow = null;
            lastFlight = "";
            lastEtd = "";
            continue;
        } else if (rowStr.includes("BẢNG PHÂN CÔNG EDIT CHUYẾN BAY")) {
            currentSection = "EDIT CHUYẾN BAY";
            sectionHeaderRow = null;
            lastFlight = "";
            lastEtd = "";
            continue;
        } else if (rowStr.includes("BẢNG PHÂN CÔNG PHỤC VỤ CHUYẾN BAY")) {
            currentSection = "PHỤC VỤ CHUYẾN BAY";
            sectionHeaderRow = null;
            lastFlight = "";
            lastEtd = "";
            continue;
        }

        // 2. Identify Section Header Row (FLT, ETD, etc.)
        if (rowStr.includes("FLT") || rowStr.includes("ETD") || rowStr.includes("SỐ HIỆU")) {
            sectionHeaderRow = row;
            flightColIdx = row.findIndex(c => c && c.toString().match(/flight|số hiệu|sh|flt/i));
            if (flightColIdx === -1) flightColIdx = 12;

            etdColIdx = row.findIndex((c, idx) =>
                idx !== flightColIdx && c && c.toString().match(/etd|giờ|time|bay/i)
            );
            if (etdColIdx === -1) etdColIdx = 13;

            gateColIdx = row.findIndex(c => c && c.toString().match(/gate|cửa/i));
            continue;
        }

        // 3. Scan for staffName if we have a header context
        if (sectionHeaderRow) {
            // Update last seen flight/etd for handle merged cells
            const currentFlightVal = row[flightColIdx];
            if (currentFlightVal && currentFlightVal.toString().trim() !== '' && currentFlightVal.toString().trim() !== '-') {
                lastFlight = currentFlightVal.toString();
            }

            let currentEtdVal = row[etdColIdx];
            // Fallback for ETD if column is empty but there's a time in the row
            if (!currentEtdVal || currentEtdVal.toString().trim() === '' || currentEtdVal.toString().trim() === '-') {
                for (let k = 0; k < Math.min(row.length, 15); k++) {
                    if (k === flightColIdx) continue;
                    const val = row[k];
                    if (val && (typeof val === 'number' || (typeof val === 'string' && val.match(/\d{1,2}:\d{2}/)))) {
                        currentEtdVal = val;
                        break;
                    }
                }
            }
            if (currentEtdVal && currentEtdVal.toString().trim() !== '' && currentEtdVal.toString().trim() !== '-') {
                lastEtd = currentEtdVal;
            }

            // Search name in columns (limited search to avoid matching unrelated data)
            // Usually staff assignments start after the first few columns
            for (let j = 2; j < row.length; j++) {
                const cell = row[j];
                if (!cell) continue;

                const cellStr = cell.toString().trim();
                const cellNormalized = removeVietnameseTones(cellStr.toLowerCase());

                // Check if name matches (allowing for roles like "Phuong 8 / SUP")
                // We use word-boundary-like check to avoid partial name matches (e.g. "ANH" vs "ANH 10")
                const nameRegex = new RegExp(`(^|\\s|[^a-zA-Z0-9])${targetNameNormalized}($|\\s|[^a-zA-Z0-9])`, 'i');

                if (cellNormalized === targetNameNormalized || nameRegex.test(cellNormalized)) {
                    const position = sectionHeaderRow[j] || "Assigned";
                    const gate = (gateColIdx !== -1) ? (row[gateColIdx] || '') : '';

                    flights.push({
                        flight: lastFlight !== "" ? lastFlight : (currentSection !== "PHỤC VỤ CHUYẾN BAY" ? currentSection : "-"),
                        etd: lastEtd !== "" ? lastEtd : "-",
                        position,
                        gate: gate.toString(),
                        section: currentSection
                    });
                    // Don't break! One row might have multiple roles for the same person
                }
            }
        }
    }

    const seenFlights = new Set();
    const uniqueFlights = [];

    for (const f of flights) {
        // Unique key per section, flight number and time
        const key = `${f.section}|${f.flight}|${f.etd}`;
        if (!seenFlights.has(key)) {
            seenFlights.add(key);
            uniqueFlights.push(f);
        }
    }

    // Return in Excel row order (no manual sorting)
    return uniqueFlights;
}

/**
 * Display combined day view (shift + flights)
 * @param {string} staffName - Name of staff member
 * @param {string} day - Day code (MON, TUE, etc.)
 */
function displayDayView(staffName, day) {
    // Get shift info
    const shiftInfo = getShiftInfo(staffName, day);

    if (!shiftInfo) {
        shiftBadge.textContent = 'Không tìm thấy dữ liệu';
        shiftTime.textContent = '';
        shiftDuration.textContent = '';
    } else if (shiftInfo.isOff) {
        shiftBadge.textContent = shiftInfo.rawCodes || 'OFF';
        shiftTime.textContent = 'Nghỉ';
        shiftDuration.textContent = '';
    } else if (shiftInfo.shifts.length > 0) {
        shiftBadge.textContent = shiftInfo.rawCodes;

        const timeDetails = shiftInfo.shifts.map(s => {
            if (s.isDirect) return s.label;
            return `${timeStr(s.start)} - ${timeStr(s.end)}`;
        }).join(' | ');

        shiftTime.textContent = timeDetails;
        shiftDuration.textContent = `Tổng: ${shiftInfo.totalDuration}h`;
    } else {
        shiftBadge.textContent = '--';
        shiftTime.textContent = 'Không có ca';
        shiftDuration.textContent = '';
    }

    // Get flight details
    const daySheetName = workbook.SheetNames.find(n => n.toUpperCase().includes(day.toUpperCase()));
    const daySheetData = daySheetName ? cachedSheets[daySheetName] : null;

    if (daySheetData) {
        const flights = getFlightDetails(staffName, daySheetData);

        // Render flight timeline
        if (flights.length === 0) {
            flightTimeline.innerHTML = '<div class="timeline-empty">Không có chuyến bay nào được phân công</div>';
        } else {
            // Group by section
            const grouped = {};
            flights.forEach(f => {
                if (!grouped[f.section]) grouped[f.section] = [];
                grouped[f.section].push(f);
            });

            let html = '';
            for (const section in grouped) {
                html += `<div class="timeline-section-header">${section}</div>`;

                grouped[section].forEach(f => {
                    const displayEtd = formatExcelTime(f.etd);
                    html += `
                        <div class="timeline-item">
                            <div class="timeline-time">${displayEtd}</div>
                            <div class="timeline-content">
                                <div class="timeline-flight">${f.flight}</div>
                            </div>
                        </div>
                    `;
                });
            }
            flightTimeline.innerHTML = html;
        }
    } else {
        flightTimeline.innerHTML = `<div class="timeline-empty">Không tìm thấy sheet cho ngày ${day}</div>`;
    }

    // Show day view
    attendanceView.classList.add('hidden');
    flightView.classList.add('hidden');
    dayView.classList.remove('hidden');
    resultSection.classList.remove('hidden');
}


function logDebug(msg) {
    const d = new Date();
    debugOutput.innerText += `[${d.toLocaleTimeString()}] ${msg}\n`;
}

/**
 * Helper to check if two time windows overlap, supporting midnight crossing.
 * Handles negative values and values > 1440 by normalizing to the daily cycle.
 */
function isOverlap(sStart, sEnd, qStart, qEnd) {
    function normalize(a, b) {
        // Shift range to be positive relative to midnight
        while (a < 0) { a += 1440; b += 1440; }
        while (a >= 1440) { a -= 1440; b -= 1440; }

        if (a > b) return [[a, 1440], [0, b]];
        if (b > 1440) return [[a, 1440], [0, b - 1440]];
        return [[a, b]];
    }

    const sRanges = normalize(sStart, sEnd);
    const qRanges = normalize(qStart, qEnd);

    for (const [s1, s2] of sRanges) {
        for (const [q1, q2] of qRanges) {
            // Overlap if max of starts < min of ends
            if (Math.max(s1, q1) < Math.min(s2, q2)) return true;
        }
    }
    return false;
}

/**
 * Find available staff for a specific day and time range
 * @param {string} day - MON, TUE, etc.
 * @param {string} startStr - HH:mm
 * @param {string} endStr - HH:mm
 */
function findAvailableStaff(day, startStr, endStr) {
    if (!qtStaffMap || qtStaffMap.size === 0) {
        alert("Vui lòng tải file Excel có BẢNG CHẤM CÔNG QT trước.");
        return;
    }

    const startMinutes = timeToMinutes(startStr);
    const endMinutes = timeToMinutes(endStr);

    // Get the daily sheet data for this day
    const daySheetName = workbook.SheetNames.find(n => n.toUpperCase().includes(day.toUpperCase()));
    const daySheetData = daySheetName ? cachedSheets[daySheetName] : null;

    if (!daySheetData) {
        alert("Không tìm thấy dữ liệu chuyến bay cho ngày " + day);
        return;
    }

    availList.innerHTML = '<div class="loading">Đang tìm kiếm...</div>';

    // Use setTimeout to allow UI to show loading state
    setTimeout(() => {
        const available = [];
        const busy = [];

        for (const [name, row] of qtStaffMap.entries()) {
            const shiftInfo = getShiftInfo(name, day);
            if (!shiftInfo || shiftInfo.isOff || shiftInfo.shifts.length === 0) continue;

            let isOnShift = false;
            let shiftLabels = [];
            for (const s of shiftInfo.shifts) {
                shiftLabels.push(`${s.code} (${timeStr(s.start)}-${timeStr(s.end)})`);
                if (isOverlap(s.start, s.end, startMinutes, endMinutes)) {
                    isOnShift = true;
                }
            }

            if (!isOnShift) continue;

            const flights = getFlightDetails(name, daySheetData);
            let kẹtLịch = [];

            for (const f of flights) {
                if (!f.etd || f.etd === '-') continue;

                const displayEtd = formatExcelTime(f.etd);
                const etdMinutes = timeToMinutes(displayEtd);
                const busyStart = etdMinutes - 240; // ETD - 4h
                const busyEnd = etdMinutes;         // ETD + 0h (Changed from + 60)

                if (isOverlap(busyStart, busyEnd, startMinutes, endMinutes)) {
                    kẹtLịch.push(f);
                }
            }

            const staffObj = {
                name,
                shifts: shiftLabels,
                busyFlights: kẹtLịch
            };

            if (kẹtLịch.length === 0) {
                available.push(staffObj);
            } else {
                busy.push(staffObj);
            }
        }

        renderAvailability(available, busy);
    }, 100);
}

function timeToMinutes(timeStr) {
    if (!timeStr || !timeStr.includes(':')) return 0;
    const [h, m] = timeStr.split(':').map(Number);
    return h * 60 + m;
}

function renderAvailability(available, busy) {
    resultSection.classList.remove('hidden');
    availabilityView.classList.remove('hidden');
    availResultsTitle.textContent = `Kết quả tìm kiếm (${available.length} người rảnh)`;

    if (available.length === 0) {
        availList.innerHTML = '<div class="timeline-empty">Không tìm thấy nhân viên nào rảnh trong khung giờ này.</div>';
    } else {
        availList.innerHTML = available.map(staff => `
            <div class="staff-card">
                <h4>${staff.name}</h4>
                <div>${staff.shifts.map(s => `<span class="shift-tag">${s}</span>`).join(' ')}</div>
                <div class="status-indicator ready">● Sẵn sàng làm việc</div>
            </div>
        `).join('');
    }

    if (busy.length > 0) {
        const busyHtml = `
            <div style="grid-column: 1 / -1; margin-top: 2rem;">
                <h3 style="margin-bottom: 1rem; color: #f87171;">Nhân viên đang kẹt lịch bay (${busy.length})</h3>
                <div class="avail-grid">
                    ${busy.map(staff => `
                        <div class="staff-card" style="opacity: 0.7;">
                            <h4>${staff.name}</h4>
                            <div class="assignment-list">
                                ${staff.busyFlights.map(f => `
                                    <div class="assignment-item">
                                        <strong>${f.flight}</strong> (${formatExcelTime(f.etd)}) - ${f.section}
                                    </div>
                                `).join('')}
                            </div>
                            <div class="status-indicator busy">● Đang kẹt lịch bay</div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;
        availList.insertAdjacentHTML('beforeend', busyHtml);
    }
}

searchAvailBtn.addEventListener('click', () => {
    const day = availDaySelect.value;
    const start = availTimeStart.value;
    const end = availTimeEnd.value;

    if (!day) {
        alert("Vui lòng chọn ngày.");
        return;
    }
    if (!start || !end) {
        alert("Vui lòng chọn khung giờ.");
        return;
    }

    findAvailableStaff(day, start, end);
});

