
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
const staffSelect = document.getElementById('staff-select');
const resultSection = document.getElementById('result-section');

// Views
const attendanceView = document.getElementById('attendance-view');
const flightView = document.getElementById('flight-view');

// Attendance Elements
const totalHoursEl = document.getElementById('total-hours');
const totalShiftsEl = document.getElementById('total-shifts');
const detailsTableBody = document.querySelector('#details-table tbody');

// Flight Elements
const flightsTableBody = document.querySelector('#flights-table tbody');

const debugOutput = document.getElementById('debug-output');

// Global Data State
let workbook = null;
let currentSheetData = null; // Array of Arrays
let headerRowIndex = -1;
let staffMap = new Map(); // Name -> Row Data (For Attendance) OR Set of Names (For Flights)
let currentSheetType = 'UNKNOWN'; // 'QT' or 'FLIGHT'

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
    } else {
        staffSelect.disabled = true;
        staffSelect.innerHTML = '<option value="">-- Trước tiên chọn Sheet --</option>';
        resultSection.classList.add('hidden');
    }
});

staffSelect.addEventListener('change', (e) => {
    if (e.target.value) {
        if (currentSheetType === 'QT') {
            calculateAttendanceStats(e.target.value);
        } else if (currentSheetType === 'FLIGHT') {
            findFlights(e.target.value);
        }
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

        sheetSelect.innerHTML = '<option value="">-- Chọn Sheet --</option>';
        workbook.SheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        controlPanel.classList.remove('hidden');

        // Auto-select priority: "BẢNG CHẤM CÔNG QT"
        const qtSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('CHẤM CÔNG QT'));
        // Or "Lịch" (Flight Schedule?)
        const likelySheet = qtSheet || workbook.SheetNames.find(n => n.toLowerCase().includes('lịch') || n.toLowerCase().includes('schedule'));

        if (likelySheet) {
            sheetSelect.value = likelySheet;
            processSheet(likelySheet);
        }

        logDebug(`File loaded. Sheets: ${workbook.SheetNames.join(', ')}`);
    };
    reader.readAsArrayBuffer(file);
}

function processSheet(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    currentSheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Determine Sheet Type
    // QT rule: Contains "CHẤM CÔNG QT" in name OR specific structure
    if (sheetName.toUpperCase().includes('QT') || sheetName.toUpperCase().includes('CHẤM CÔNG')) {
        currentSheetType = 'QT';
        attendanceView.classList.remove('hidden');
        flightView.classList.add('hidden');
        processQTSheet(sheetName);
    } else {
        // Assume Flight Schedule (Daily)
        currentSheetType = 'FLIGHT';
        attendanceView.classList.add('hidden');
        flightView.classList.remove('hidden');
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
            const etd = row[etdColIdx] || '-';
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


function logDebug(msg) {
    const d = new Date();
    debugOutput.innerText += `[${d.toLocaleTimeString()}] ${msg}\n`;
}
