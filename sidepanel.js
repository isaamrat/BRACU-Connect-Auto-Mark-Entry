const fillBtn = document.getElementById('fillBtn');
const btnProgress = document.getElementById('btnProgress');
const btnText = document.getElementById('btnText');
const fileInput = document.getElementById('fileInput');

function hideFillButton() { fillBtn.style.display = 'none'; }
function showFillButton() {
  fillBtn.style.display = 'flex';
  btnProgress.style.width = '0%';
  btnText.textContent = 'Fill Marks';
}


const fileSuccess = document.getElementById("fileSuccess");
const pageVerification = document.getElementById("pageVerification");
const fileUploadText = document.getElementById("fileUploadText");
const rawMarkChip = document.getElementById('rawMarkChip');
const roundedMarkChip = document.getElementById('roundedMarkChip');
const markTypeContainer = document.getElementById('markTypeContainer');
const configControls = document.getElementById('configControls');
const sheetSelect = document.getElementById('sheetSelect');

let studentData = [];
let currentWorkbook = null;
let isProcessing = false;
let selectedMarkIndex = 0;

rawMarkChip.addEventListener('click', () => {
  selectedMarkIndex = 0;
  rawMarkChip.classList.add('active');
  roundedMarkChip.classList.remove('active');
  showFillButton();
  clearUnmatchedStudents();
});

roundedMarkChip.addEventListener('click', () => {
  selectedMarkIndex = 1;
  roundedMarkChip.classList.add('active');
  rawMarkChip.classList.remove('active');
  showFillButton();
  clearUnmatchedStudents();
});

document.addEventListener('DOMContentLoaded', async () => {
  await verifyCurrentPage();
});

async function verifyCurrentPage() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const results = await chrome.scripting.executeScript({
      target: { tabId: tabs[0].id },
      function: checkForMarksInputFields
    });

    const hasMarksFields = results[0].result;

    if (hasMarksFields) {
      showUploadSection();
    } else {
      showPageVerificationError();
    }
  } catch (error) {
    console.error("Error verifying page:", error);
    showPageVerificationError();
  }
}


function checkForMarksInputFields() {
  const marksInputs = document.querySelectorAll('app-masked-input input[placeholder="Marks"], input[placeholder="Marks"]');
  const studentInputs = document.querySelectorAll('input[placeholder="Student"]');
  return marksInputs.length > 0 && studentInputs.length > 0;
}


function showUploadSection() {
  pageVerification.classList.remove('show');
  uploadSection.classList.add('show');
}

function showPageVerificationError() {
  uploadSection.classList.remove('show');
  pageVerification.classList.add('show');
}

function showMarkTypeContainer() {
  if (configControls) configControls.style.display = 'flex';
}

function hideMarkTypeContainer() {
  if (configControls) configControls.style.display = 'none';
}

// Listen for progress updates
chrome.runtime.onMessage.addListener((message) => {
  if (message.type === 'PROGRESS_UPDATE') {
    const progress = message.progress;
    btnProgress.style.width = `${progress}%`;
    btnText.textContent = `Filling... ${progress}%`;
  }
});

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) {
    resetFileUploadDisplay();
    hideFileSuccess();
    hideMarkTypeContainer();
    return;
  }

  showFillButton();


  updateFileUploadDisplay(file.name);

  const reader = new FileReader();

  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      currentWorkbook = XLSX.read(data, { type: "array" });

      // Populate sheet dropdown
      sheetSelect.innerHTML = '';
      currentWorkbook.SheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      });

      // Default to "Final GradeSheet" if it exists, otherwise first sheet
      const defaultSheet = currentWorkbook.SheetNames.includes("Final GradeSheet")
        ? "Final GradeSheet"
        : currentWorkbook.SheetNames[0];

      sheetSelect.value = defaultSheet;
      updateStudentDataFromSelectedSheet();

      showMessage(`Excel file loaded successfully`, "success");
      showFileSuccess();
      showMarkTypeContainer();
      clearUnmatchedStudents();
      showFillButton();
    } catch (error) {
      showMessage("Error reading the Excel file. Please check the file format and try again.", "error");
      resetFileUploadDisplay();
      hideFileSuccess();
      hideMarkTypeContainer();
    }
  };

  reader.onerror = () => {
    showMessage("Failed to read the file. Please try again with a different file.", "error");
    resetFileUploadDisplay();
    hideFileSuccess();
    hideMarkTypeContainer();
  };

  reader.readAsArrayBuffer(file);
});

function updateFileUploadDisplay(fileName) {
  fileUploadDisplay.classList.add('has-file');
  fileUploadText.innerHTML = `<span class="file-upload-icon">✅</span><span>${fileName}</span>`;
}

function resetFileUploadDisplay() {
  fileUploadDisplay.classList.remove('has-file');
  fileUploadText.innerHTML = `<span class="file-upload-icon">📁</span><span>Choose Grade Sheet Excel file</span>`;
  if (configControls) configControls.style.display = 'none';
  currentWorkbook = null;
  studentData = [];
}

function updateStudentDataFromSelectedSheet() {
  if (!currentWorkbook) return;
  const sheetName = sheetSelect.value;
  const worksheet = currentWorkbook.Sheets[sheetName];
  if (worksheet) {
    studentData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const validation = validateSheetRequirements(studentData);
    if (!validation.isValid) {
      disableFillButton(true, `Missing required column: <strong>${validation.missingField}</strong>`);
    } else {
      disableFillButton(false);
      showFillButton();
    }
  }
}

function validateSheetRequirements(data) {
  if (!data || data.length === 0) return { isValid: false, missingField: "Empty sheet" };

  let headerRowIndex = -1;
  let hasId = false;

  // Search for ID # column
  for (let r = 0; r < Math.min(data.length, 20); r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      if (row[c] && String(row[c]).trim().toLowerCase().includes('id #')) {
        headerRowIndex = r;
        hasId = true;
        break;
      }
    }
    if (hasId) break;
  }

  if (!hasId) return { isValid: false, missingField: "ID #" };

  // Check for Total column in same row
  const headerRow = data[headerRowIndex];
  let hasTotal = false;
  for (let i = 0; i < headerRow.length; i++) {
    const h = headerRow[i];
    if (h && String(h).trim().toLowerCase() === 'total') {
      hasTotal = true;
      break;
    }
  }

  if (!hasTotal) return { isValid: false, missingField: "Total" };

  return { isValid: true };
}

function disableFillButton(disabled, message = "") {
  fillBtn.disabled = disabled;
  fillBtn.style.opacity = disabled ? "0.5" : "1";
  fillBtn.style.cursor = disabled ? "not-allowed" : "pointer";

  const alert = document.getElementById('validationAlert');
  if (alert) {
    if (disabled && message) {
      alert.querySelector('.alert-message').innerHTML = message;
      alert.style.display = 'flex';
    } else {
      alert.style.display = 'none';
    }
  }
}

sheetSelect.addEventListener("change", () => {
  updateStudentDataFromSelectedSheet();
  showFillButton();
  clearUnmatchedStudents();
  showMessage(`Sheet switched to "${sheetSelect.value}"`, "info");
});


fillBtn.addEventListener("click", async () => {
  if (studentData.length === 0) {
    showMessage("Please select an Excel file first.", "error");
    return;
  }

  if (isProcessing) return;

  try {
    setButtonLoading(true);

    const results = await chrome.scripting.executeScript({
      target: { tabId: (await chrome.tabs.query({ active: true, currentWindow: true }))[0].id },
      function: fillMarksOnPage,
      args: [studentData, selectedMarkIndex]
    });

    const { unmatchedStudents, pageOnlyStudents, totalStudents, marksEnteredCount, absentCount } = results[0].result;
    const totalMatched = marksEnteredCount + absentCount;

    if (totalMatched > 0) {
      hideFillButton();
    }


    // If nothing matched at all, nudge the user
    if (unmatchedStudents.length === totalStudents && totalStudents > 0) {
      showMessage("No students were matched! Please verify that your Excel file course/section matches the page.", "error");
      displayUnmatchedStudents(unmatchedStudents, pageOnlyStudents);
      return;
    }

    // Summary message
    const parts = [];
    if (marksEnteredCount > 0) parts.push(`${marksEnteredCount} Marks`);
    if (absentCount > 0) parts.push(`${absentCount} Absent`);
    if (unmatchedStudents.length) parts.push(`${unmatchedStudents.length} Missing`);
    if (pageOnlyStudents.length) parts.push(`${pageOnlyStudents.length} Extras`);

    const msg = parts.join(" · ");

    if (unmatchedStudents.length === 0 && pageOnlyStudents.length === 0) {
      const successParts = [];
      if (marksEnteredCount > 0) successParts.push(`${marksEnteredCount} marks`);
      if (absentCount > 0) successParts.push(`${absentCount} absences`);
      const successDetail = successParts.join(" and ");
      showMessage(`All students matched! (${successDetail}) 🎉`, "success");
    } else {
      showMessage(msg, "warning");
    }

    displayUnmatchedStudents(unmatchedStudents, pageOnlyStudents);
  } catch (error) {
    showMessage("Error accessing the current tab. Please ensure you're on the correct page and try again.", "error");
  } finally {
    setButtonLoading(false);
  }
});


function setButtonLoading(loading) {
  isProcessing = loading;
  if (loading) {
    fillBtn.classList.add('loading');
    fillBtn.disabled = true;
    btnText.textContent = 'Processing...';
    btnProgress.style.width = '0%';
  } else {
    fillBtn.classList.remove('loading');
    fillBtn.disabled = false;
    btnText.textContent = 'Fill Marks';
    btnProgress.style.width = '0%';
  }
}

function showFileSuccess() {
  fileSuccess.classList.add('show');
}

function hideFileSuccess() {
  fileSuccess.classList.remove('show');
}

function showMessage(message, type = "info") {
  // Remove any existing message
  const existingMessage = document.querySelector('.message');
  if (existingMessage) {
    existingMessage.remove();
  }

  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${type}`;
  messageDiv.innerHTML = `
    <div class="message-content">
      <div class="message-icon"></div>
      <span class="message-text">${message}</span>
    </div>
  `;

  const popupContainer = document.querySelector('.popup-container');
  popupContainer.appendChild(messageDiv);

  setTimeout(() => messageDiv.classList.add('show'), 10);

  if (type === "success") {
    setTimeout(() => {
      if (messageDiv.parentNode) {
        messageDiv.classList.remove('show');
        setTimeout(() => messageDiv.remove(), 300);
      }
    }, 10000);
  }
}


function displayUnmatchedStudents(unmatchedStudents, pageOnlyStudents) {
  const container = document.getElementById("unmatchedContainer") || createUnmatchedContainer();

  const excelMisses = unmatchedStudents || [];
  const pageExtras = pageOnlyStudents || [];

  if (excelMisses.length === 0 && pageExtras.length === 0) {
    container.classList.add('success');
    container.style.display = 'block';
    container.innerHTML = `
      <div class="result-box success">
        <div class="success-icon">✓</div>
        <div class="success-text">All students matched successfully!</div>
      </div>
    `;
    return;
  }

  container.classList.remove('success');

  function table(rows, cols) {
    let thead = cols.map(c => `<th>${c}</th>`).join('');
    let tbody = rows.map(r => `
      <tr>
        <td><span class="student-id">${r.id ?? ''}</span></td>
        <td><span class="student-name">${r.name ?? ''}</span></td>
        ${'finalMark' in r ? `<td><span class="student-mark">${r.finalMark ?? ''}</span></td>` : ''}
      </tr>
    `).join('');
    return `
      <div class="table-container">
        <table class="unmatched-table">
          <thead><tr>${thead}</tr></thead>
          <tbody>${tbody}</tbody>
        </table>
      </div>
    `;
  }

  let html = `
    <div class="unmatched-header">
    <div class="flex">
      <div class="warning-icon">⚠</div>
      <h3>Review Needed</h3>
    </div>
      <p class="unmatched-subtitle">Some records didn’t line up between Excel and the page.</p>
    </div>
  `;

  if (excelMisses.length > 0) {
    html += `
      <h4 style="margin:8px 0 6px;color:#991b1b;font-size:14px;">In Excel, not on page (${excelMisses.length})</h4>
      ${table(excelMisses, ['Student ID', 'Name', 'Total Mark'])}
    `;
  }

  if (pageExtras.length > 0) {
    html += `
      <h4 style="margin:12px 0 6px;color:#991b1b;font-size:14px;">On page, not in Excel (${pageExtras.length})</h4>
      ${table(pageExtras.map(x => ({ id: x.id, name: x.name })), ['Student ID', 'Name'])}
    `;
  }

  container.innerHTML = html;
  container.style.display = 'block';
}


function createUnmatchedContainer() {
  const container = document.createElement('div');
  container.id = 'unmatchedContainer';
  container.className = 'unmatched-container';

  const style = document.createElement('style');
  style.textContent = `
    .message {
      padding: 16px;
      border-radius: 12px;
      margin: 8px 0;
      font-size: 14px;
      text-align: left;
      opacity: 0;
      transform: translateY(-10px);
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      border: 1px solid;
    }
    
    .message.show {
      opacity: 1;
      transform: translateY(0);
    }
    
    .message-content {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    
    .message-icon {
      width: 20px;
      height: 20px;
      border-radius: 50%;
      flex-shrink: 0;
      position: relative;
    }
    
    .message-icon::after {
      content: '';
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      font-weight: bold;
      font-size: 12px;
    }
    
    .message.success {
      background: linear-gradient(135deg, #ecfdf5, #d1fae5);
      color: #065f46;
      border-color: #a7f3d0;
    }
    
    .message.success .message-icon {
      background: #059669;
    }
    
    .message.success .message-icon::after {
      content: '✓';
      color: white;
    }
    
    .message.error {
      background: linear-gradient(135deg, #fef2f2, #fecaca);
      color: #991b1b;
      border-color: #fca5a5;
    }
    
    .message.error .message-icon {
      background: #dc2626;
    }
    
    .message.error .message-icon::after {
      content: '✕';
      color: white;
    }
    
    .message.warning {
      background: linear-gradient(135deg, #fffbeb, #fed7aa);
      color: #92400e;
      border-color: #fdba74;
    }
    
    .message.warning .message-icon {
      background: #d97706;
    }
    
    .message.warning .message-icon::after {
      content: '!';
      color: white;
    }
    
    .message.info {
      background: linear-gradient(135deg, #eff6ff, #dbeafe);
      color: #1e40af;
      border-color: #93c5fd;
    }
    
    .message.info .message-icon {
      background: #3b82f6;
    }
    
    .message.info .message-icon::after {
      content: 'i';
      color: white;
    }
    
    .unmatched-container {
      margin-top: 8px;
      padding: 16px;
      border: 1px solid #fca5a5;
      border-radius: 12px;
      background: linear-gradient(135deg, #fef2f2, #ffffff);
      display: none;
      max-height: 400px;
      overflow: hidden;
      width: 100%;
      box-sizing: border-box;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
      transition: all 0.3s ease;
    }

    .unmatched-container.success {
      border-color: #a7f3d0;
      background: linear-gradient(135deg, #ecfdf5, #ffffff);
      box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.1);
    }
    
    .unmatched-header {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin-bottom: 8px;
      padding-bottom: 6px;
    }
    
    .unmatched-header > div:first-child {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    
    .warning-icon {
      width: 24px;
      height: 24px;
      background: #dc2626;
      color: white;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      font-weight: bold;
      flex-shrink: 0;
    }
    
    .unmatched-header h3 {
      margin: 0;
      color: #991b1b;
      font-size: 16px;
      font-weight: 600;
    }
    
    .unmatched-subtitle {
      color: #7f1d1d;
      font-size: 12px;
      margin: 0;
      font-style: italic;
    }
    
    .result-box {
      display: flex;
      align-items: center;
      gap: 12px;
      color: #065f46;
    }
    
    .success-icon {
      width: 24px;
      height: 24px;
      background: #059669;
      color: white;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      font-weight: bold;
      flex-shrink: 0;
    }
    
    .success-text {
      font-weight: 600;
      font-size: 14px;
    }
    
    .table-container {
      max-height: 280px;
      overflow-y: auto;
      overflow-x: auto;
      border-radius: 8px;
      border: 1px solid #fecaca;
    }
    
    .unmatched-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 11px;
      background-color: white;
    }
    
    .unmatched-table th {
      background: linear-gradient(135deg, #dc2626, #b91c1c);
      color: white;
      padding: 12px 8px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
      z-index: 1;
      font-size: 12px;
      letter-spacing: 0.025em;
    }
    
    .unmatched-table td {
      padding: 10px 8px;
      border-bottom: 1px solid #fecaca;
    }
    
    .unmatched-table tr:nth-child(even) {
      background: #fef7f7;
    }
    
    .unmatched-table tr:hover {
      background: #fef2f2;
    }
    
    .student-id {
      font-weight: 600;
      color: #991b1b;
      background: #fecaca;
      padding: 4px 8px;
      border-radius: 6px;
      font-size: 12px;
    }
    
    .student-name {
      color: #991b1b;
      font-weight: 500;
    }
    
    .student-mark {
      color: #dc2626;
      font-weight: 600;
      background: #fee2e2;
      padding: 4px 8px;
      border-radius: 6px;
      font-size: 12px;
    }
    
    /* Custom scrollbar */
    .table-container::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    
    .table-container::-webkit-scrollbar-track {
      background: #f1f5f9;
      border-radius: 3px;
    }
    
    .table-container::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 3px;
    }
    
    .table-container::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }
  `;

  if (!document.head.querySelector('style[data-popup-styles]')) {
    style.setAttribute('data-popup-styles', 'true');
    document.head.appendChild(style);
  }

  const popupContainer = document.querySelector('.popup-container');
  popupContainer.appendChild(container);
  return container;
}

function clearUnmatchedStudents() {
  const container = document.getElementById("unmatchedContainer");
  if (container) {
    container.style.display = 'none';
    container.innerHTML = '';
  }
}



async function fillMarksOnPage(data, selectedMarkIndex = 0) {
  try {
    if (!data || data.length < 2) throw new Error("Excel data must have at least 2 rows");

    // ---------- Header detection ----------
    let headerRowIndex = -1;
    let idColumnIndex = -1;
    let totalColumnIndex = -1;
    let finalColumnIndex = -1;
    let nameColumnIndex = -1;

    for (let r = 0; r < data.length - 1; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (cell && String(cell).trim().toLowerCase().includes('id #')) {
          headerRowIndex = r;
          idColumnIndex = c;
          break;
        }
      }
      if (headerRowIndex !== -1) break;
    }
    if (headerRowIndex === -1) throw new Error("Could not find header row with 'ID #' column in the Excel file");

    const headerRow = data[headerRowIndex];

    // First 'Total' column (100.00)
    const totalColumns = [];
    for (let i = 0; i < headerRow.length; i++) {
      const h = headerRow[i];
      if (h && String(h).trim().toLowerCase() === 'total') totalColumns.push(i);
      if (h && String(h).trim().toLowerCase() === 'final') finalColumnIndex = i;
    }
    if (totalColumns.length === 0) throw new Error("Could not find any 'Total' column in the Excel file");

    // Pick the column based on user selection (Raw = index 0, Rounded = index 1)
    // If only one Total column exists, it will naturally pick the first one regardless.
    totalColumnIndex = totalColumns[Math.min(selectedMarkIndex, totalColumns.length - 1)];

    if (finalColumnIndex === -1) {
      for (let i = 0; i < headerRow.length; i++) {
        const txt = String(headerRow[i] ?? '').toLowerCase().replace(/\s+/g, ' ').trim();
        if (txt.startsWith('final')) { finalColumnIndex = i; break; }
      }
    }

    // Optional: detect a name column near ID
    for (let i = idColumnIndex + 1; i < Math.min(idColumnIndex + 3, headerRow.length); i++) {
      const h = headerRow[i];
      if (h && !/^\d+(\.\d+)?$/.test(String(h).trim())) { nameColumnIndex = i; break; }
    }

    // ---------- Extract "Full Mark" from meta row ----------
    let fullMark = 100;
    const metaRow = data[headerRowIndex + 1] || [];
    const metaCell = metaRow[totalColumnIndex];
    // if (metaCell != null) {
    //   const num = parseFloat(String(metaCell).replace(/[^\d.]/g, ''));
    //   if (!isNaN(num) && num > 0) fullMark = num;
    // }

    // ---------- Student rows ----------
    const studentDataStartIndex = headerRowIndex + 2;
    const studentRows = data.slice(studentDataStartIndex);

    // ---------- Helpers ----------
    const normId = (x) => String(x ?? '').replace(/\D/g, '').trim();
    const normName = (x) => String(x ?? '').toLowerCase().replace(/\s+/g, ' ').trim();
    function extractIdName(raw) {
      const val = String(raw ?? '').replace(/\s+/g, ' ').trim();
      const m = val.match(/^\s*([0-9]{4,})\s*[-–—]\s*(.+)$/);
      if (m) return { id: normId(m[1]), name: m[2].trim() };
      const parts = val.split(/\s*[-–—]\s*/);
      if (parts.length >= 2) return { id: normId(parts[0]), name: parts.slice(1).join(' - ').trim() };
      return { id: normId(val), name: val };
    }

    function isEffectivelyDisabled(inputEl) {
      if (!inputEl) return true;
      if (inputEl.disabled || inputEl.readOnly) return true;
      if (inputEl.hasAttribute('disabled')) return true;
      if (inputEl.getAttribute('aria-disabled') === 'true') return true;

      const wrapper = inputEl.closest('.mat-mdc-form-field, .mat-mdc-text-field-wrapper, .mdc-text-field, app-masked-input');
      if (wrapper) {
        if (
          wrapper.classList.contains('mat-mdc-form-field-disabled') ||
          wrapper.classList.contains('mdc-text-field--disabled') ||
          wrapper.getAttribute('aria-disabled') === 'true'
        ) return true;
      }
      return false;
    }

    function ensureKksRoundedBgStyles() {
      if (document.getElementById('kks-rounded-bg')) return;
      const s = document.createElement('style');
      s.id = 'kks-rounded-bg';
      s.textContent = `
    .kks-colored { position: relative; }
    .kks-colored::before{
      content:'';
      position:absolute;
      left:8px; right:8px;   /* inset horizontally */
      top:0px; bottom:26px;   /* trims bottom background */
      border-radius:8px;     /* rounded corners */
      background: var(--kks-bg, #dcfce7);
      pointer-events:none;   /* never intercept clicks */
    }
  `;
      document.head.appendChild(s);
    }

    function colorRow(rowEl, color) {
      if (!rowEl) return;
      rowEl.classList.add('kks-colored');
      rowEl.style.setProperty('--kks-bg', color);
    }



    async function setStatusAbsent(rowRoot) {
      let sel = rowRoot.querySelector('select[formcontrolname="status"], select[name*="status" i], select[placeholder="Status"], select');
      if (sel) {
        const opts = Array.from(sel.options || []);
        const opt = opts.find(o => String(o.textContent || '').trim().toLowerCase() === 'absent');
        if (opt) {
          sel.value = opt.value;
          sel.dispatchEvent(new Event('change', { bubbles: true }));
          sel.dispatchEvent(new Event('input', { bubbles: true }));
          return true;
        }
      }
      const matSelect = rowRoot.querySelector('mat-select, .mat-mdc-select') || rowRoot.querySelector('[role="combobox"]');
      if (matSelect) {
        const trigger = matSelect.querySelector('.mat-mdc-select-trigger') || matSelect;
        try {
          trigger.click();
          await new Promise(r => setTimeout(r, 120));
          const panelOptions = Array.from(document.querySelectorAll('.mat-mdc-option, mat-option'));
          const target = panelOptions.find(el => String(el.textContent || '').trim().toLowerCase() === 'absent');
          if (target) {
            target.click();
            await new Promise(r => setTimeout(r, 60));
            document.body.click();
            return true;
          }
          document.activeElement && document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        } catch (e) { }
      }
      return false;
    }

    // ---------- Build Excel lookup ----------
    const excelById = new Map();   // id -> { id, name, total, isAbsent }
    const excelByName = new Map(); // nameNorm -> { id, name, total, isAbsent }
    let totalStudentsProcessed = 0;

    for (const row of studentRows) {
      if (!row || row.length === 0) continue;

      const id = normId(row[idColumnIndex]);
      const totalMarks = row[totalColumnIndex];

      let isAbsent = false;
      if (finalColumnIndex !== -1) {
        const finalCell = row[finalColumnIndex];
        if (finalCell != null && typeof finalCell !== 'number') {
          const txt = String(finalCell).toLowerCase();
          if (txt.includes('absent') || txt === 'null' || txt === 'abs' || txt === 'a' || txt === 'na' || txt === 'nan') isAbsent = true;
        }
      }

      // name for fallback
      let name = (nameColumnIndex !== -1 ? row[nameColumnIndex] : undefined);
      if (!name) {
        for (let i = idColumnIndex + 1; i < Math.min(idColumnIndex + 4, row.length); i++) {
          if (row[i] && isNaN(row[i])) { name = row[i]; break; }
        }
      }

      if (!id || (totalMarks === undefined || totalMarks === null || totalMarks === '')) continue;

      totalStudentsProcessed++;
      const rec = { id, name: String(name ?? ''), total: totalMarks, isAbsent };
      excelById.set(id, rec);
      if (name) excelByName.set(normName(name), rec);
    }

    // ---------- Discover page rows ----------
    const pageById = new Map();
    const pageByName = new Map();
    const rows = [];

    const studentInputs = Array.from(document.querySelectorAll('input[placeholder="Student"]'));
    studentInputs.forEach(studentInput => {
      const rowRoot = studentInput.closest('.row') || studentInput.closest('formly-group') || studentInput.closest('.border-bottom') || studentInput.closest('[formly-field]') || studentInput.closest('form');
      if (!rowRoot) return;

      let marksInput = rowRoot.querySelector('app-masked-input input[placeholder="Marks"], input[placeholder="Marks"]');
      if (!marksInput) {
        const group = studentInput.closest('formly-group, .row, .border-bottom') || document;
        marksInput = group.querySelector('app-masked-input input[placeholder="Marks"], input[placeholder="Marks"]');
      }
      if (!marksInput) return;

      const marksContainer = marksInput.closest('.mdc-text-field, .mat-mdc-text-field-wrapper, .mat-mdc-form-field');
      const { id: pid, name: pname } = extractIdName(studentInput.value);

      const rowObj = { studentInput, marksInput, marksContainer, id: pid, name: pname, root: rowRoot };
      rows.push(rowObj);
      if (pid) pageById.set(pid, rowObj);
      if (pname) pageByName.set(normName(pname), rowObj);
    });

    // ---------- Fill top-level "Total marks" field ----------
    (function setTotalMarksOnPage(total) {
      let totalInput =
        document.querySelector('input[placeholder="Total marks"], input[placeholder="Total Marks"], input[name="totalMarks"], input[formcontrolname="totalMarks"]');
      if (!totalInput) {
        const candidates = Array.from(document.querySelectorAll('formly-field, formly-group, .row, .col, .container, .mat-mdc-form-field, .mat-mdc-text-field-wrapper, .border-bottom'));
        for (const node of candidates) {
          const text = (node.innerText || '').toLowerCase();
          if (text.includes('total marks')) {
            totalInput = node.querySelector('input');
            if (totalInput) break;
          }
        }
      }
      if (totalInput) {
        const safe = Number.isFinite(total) ? total : 100;
        totalInput.value = safe;
        totalInput.dispatchEvent(new Event('input', { bubbles: true }));
        totalInput.dispatchEvent(new Event('change', { bubbles: true }));

        // make the Total marks field green like per-student marks
        const totalContainer =
          totalInput.closest('.mdc-text-field, .mat-mdc-text-field-wrapper, .mat-mdc-form-field') || totalInput.parentElement;

        if (totalContainer) {
          totalContainer.style.backgroundColor = '#dcfce7';
          totalContainer.style.borderColor = '#22c55e';
          totalContainer.style.borderRadius = '8px';
        } else {
          // fallback for plain inputs
          totalInput.style.backgroundColor = '#dcfce7';
          totalInput.style.border = '1px solid #22c55e';
          totalInput.style.borderRadius = '8px';
        }
      }

    })(fullMark);

    // ---------- Match & fill per-student ----------
    const unmatchedStudents = [];   // in Excel, not on page
    let marksEnteredCount = 0;
    let absentCount = 0;
    const totalToProcess = excelById.size;
    let currentProcessed = 0;


    for (const rec of excelById.values()) {
      let pageRow = pageById.get(rec.id);
      if (!pageRow && rec.name) {
        pageRow = pageByName.get(normName(rec.name));
      }

      if (pageRow) {
        if (rec.isAbsent) {
          // Set status to Absent and DO NOT touch marks
          await setStatusAbsent(pageRow.root);
          // no matchedCount++ because we didn't fill marks
          colorRow(pageRow.root, '#fff7ed');   // orangy yellow
          absentCount++;

        } else if (!isEffectivelyDisabled(pageRow.marksInput)) {
          // Only fill if the marks input is not disabled/readonly
          pageRow.marksInput.value = rec.total;
          pageRow.marksInput.dispatchEvent(new Event('input', { bubbles: true }));

          setTimeout(() => {
            if (pageRow.marksContainer) {
              pageRow.marksContainer.style.backgroundColor = '#dcfce7';
              pageRow.marksContainer.style.borderColor = '#22c55e';
              pageRow.marksContainer.style.borderRadius = '8px';
              pageRow.marksContainer.style.transition = 'all 0.3s ease';
            }
          }, 50);

          colorRow(pageRow.root, '#dcfce7');   // same green as marks field

          marksEnteredCount++;
        }
        // else: input disabled → skip
      } else {
        unmatchedStudents.push({ id: rec.id, name: rec.name || 'Unknown', finalMark: rec.total });
      }

      currentProcessed++;
      const progress = Math.round((currentProcessed / totalToProcess) * 100);
      try {
        chrome.runtime.sendMessage({ type: 'PROGRESS_UPDATE', progress });
      } catch (e) {
        // extension context might be gone, but we continue
      }
    }
    ensureKksRoundedBgStyles();




    // ---------- NEW: students on page but not in Excel ----------
    const excelIds = new Set(excelById.keys());
    const excelNames = new Set(Array.from(excelByName.keys()));
    const pageOnlyStudents = [];
    for (const r of rows) {
      const inById = r.id && excelIds.has(r.id);
      const inByName = r.name && excelNames.has(normName(r.name));
      if (!inById && !inByName) {
        pageOnlyStudents.push({ id: r.id || '—', name: r.name || 'Unknown' });
      }
    }

    return {
      unmatchedStudents,           // Excel -> not on page
      pageOnlyStudents,            // Page -> not in Excel
      totalStudents: studentRows.filter(x => x && x.length).length,
      marksEnteredCount,
      absentCount
    };

  } catch (error) {
    console.error("Error processing Excel data:", error);
    return { unmatchedStudents: [], pageOnlyStudents: [], totalStudents: 0, marksEnteredCount: 0, absentCount: 0 };
  }
}




