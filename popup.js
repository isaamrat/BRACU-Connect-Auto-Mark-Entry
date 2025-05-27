const fileInput = document.getElementById("fileInput");
const fillBtn = document.getElementById("fillBtn");
const fileSuccess = document.getElementById("fileSuccess");
const btnText = fillBtn.querySelector('.btn-text');
const pageVerification = document.getElementById("pageVerification");
const uploadSection = document.getElementById("uploadSection");
const fileUploadDisplay = document.getElementById("fileUploadDisplay");
const fileUploadText = document.getElementById("fileUploadText");

let studentData = [];
let isProcessing = false;

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
  const marksInputs = document.querySelectorAll('input[placeholder="Marks"]');
  const studentIdInputs = document.querySelectorAll('input[placeholder="StudentId"]');
  
  return marksInputs.length > 0 && studentIdInputs.length > 0;
}

function showUploadSection() {
  pageVerification.classList.remove('show');
  uploadSection.classList.add('show');
}

function showPageVerificationError() {
  uploadSection.classList.remove('show');
  pageVerification.classList.add('show');
}

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) {
    resetFileUploadDisplay();
    hideFileSuccess();
    return;
  }
  
  updateFileUploadDisplay(file.name);
  
  const reader = new FileReader();

  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = "Final GradeSheet";
      const worksheet = workbook.Sheets[sheetName];

      if (!worksheet) {
        showMessage("Sheet 'Final GradeSheet' not found in the Excel file. Please ensure your Excel file contains the correct sheet name.", "error");
        resetFileUploadDisplay();
        hideFileSuccess();
        return;
      }

      studentData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      studentData.shift(); 

      showMessage(`Excel file loaded successfully`, "success");
      showFileSuccess();
      clearUnmatchedStudents();
    } catch (error) {
      showMessage("Error reading the Excel file. Please check the file format and try again.", "error");
      resetFileUploadDisplay();
      hideFileSuccess();
    }
  };

  reader.onerror = () => {
    showMessage("Failed to read the file. Please try again with a different file.", "error");
    resetFileUploadDisplay();
    hideFileSuccess();
  };

  reader.readAsArrayBuffer(file);
});

function updateFileUploadDisplay(fileName) {
  fileUploadDisplay.classList.add('has-file');
  fileUploadText.innerHTML = `<span class="file-upload-icon">✅</span><span>${fileName}</span>`;
}

function resetFileUploadDisplay() {
  fileUploadDisplay.classList.remove('has-file');
  fileUploadText.innerHTML = `<span class="file-upload-icon">📁</span><span>Choose Excel file or drag and drop</span>`;
}

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
      args: [studentData]
    });
    
    const { unmatchedStudents, totalStudents, matchedCount } = results[0].result;
    
    if (unmatchedStudents.length === totalStudents && totalStudents > 0) {
      showMessage("No students were matched! Please verify that your Excel file course and section match the BRACU Connect Mark Entry page course and section.", "error");
      displayUnmatchedStudents(unmatchedStudents);
      return;
    }
    
    displayUnmatchedStudents(unmatchedStudents);
    
    if (unmatchedStudents.length === 0) {
      showMessage(`All ${matchedCount} student marks filled successfully! 🎉`, "success");
    } else {
      showMessage(`${matchedCount} student marks filled successfully. ${unmatchedStudents.length} students could not be matched.`, "warning");
    }
  } catch (error) {
    showMessage("Error accessing the current tab. Please ensure you're on the correct BRACU Connect page and try again.", "error");
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
  } else {
    fillBtn.classList.remove('loading');
    fillBtn.disabled = false;
    btnText.textContent = 'Fill Marks';
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

function displayUnmatchedStudents(unmatchedStudents) {
  const container = document.getElementById("unmatchedContainer") || createUnmatchedContainer();
  
  if (unmatchedStudents.length === 0) {
    container.innerHTML = `
      <div class="success-message">
        <div class="success-icon">✓</div>
        <div class="success-text">All students matched successfully!</div>
      </div>
    `;
    container.style.display = 'block';
    return;
  }
  
  let html = `
    <div class="unmatched-header">
      <div class="warning-icon">⚠</div>
      <h3>Unmatched Students (${unmatchedStudents.length})</h3>
      <p class="unmatched-subtitle">These students from your Excel file could not be matched with the form:</p>
    </div>
    <div class="table-container">
      <table class="unmatched-table">
        <thead>
          <tr>
            <th>Student ID</th>
            <th>Name</th>
            <th>Total Mark</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  unmatchedStudents.forEach(student => {
    html += `
      <tr>
        <td><span class="student-id">${student.id}</span></td>
        <td><span class="student-name">${student.name}</span></td>
        <td><span class="student-mark">${student.finalMark}</span></td>
      </tr>
    `;
  });
  
  html += `
        </tbody>
      </table>
    </div>
  `;
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
      margin: 16px 0;
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
      margin-top: 24px;
      padding: 20px;
      border: 1px solid #fca5a5;
      border-radius: 16px;
      background: linear-gradient(135deg, #fef2f2, #ffffff);
      display: none;
      max-height: 400px;
      overflow: hidden;
      width: 100%;
      box-sizing: border-box;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .unmatched-header {
      display: flex;
      flex-direction: column;
      gap: 8px;
      margin-bottom: 16px;
      padding-bottom: 12px;
      border-bottom: 1px solid #fecaca;
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
    
    .success-message {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 16px;
      background: linear-gradient(135deg, #ecfdf5, #d1fae5);
      border-radius: 12px;
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
      color: #374151;
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

function fillMarksOnPage(data) {
  try {
    // Dynamic header detection logic
    if (!data || data.length < 2) {
      throw new Error("Excel data must have at least 2 rows");
    }
    
    let headerRowIndex = -1;
    let idColumnIndex = -1;
    let totalColumnIndex = -1;
    
    for (let rowIndex = 0; rowIndex < data.length - 1; rowIndex++) {
      const currentRow = data[rowIndex];
      if (!currentRow) continue;
      
      for (let colIndex = 0; colIndex < currentRow.length; colIndex++) {
        if (currentRow[colIndex] && String(currentRow[colIndex]).trim().toLowerCase().includes('id #')) {
          headerRowIndex = rowIndex;
          idColumnIndex = colIndex;
          break;
        }
      }
      
      if (headerRowIndex !== -1) break;
    }
    
    if (headerRowIndex === -1) {
      throw new Error("Could not find header row with 'ID #' column in the Excel file");
    }
    
    const headerRow = data[headerRowIndex];
    
    const totalColumns = [];
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] && String(headerRow[i]).trim().toLowerCase() === 'total') {
        totalColumns.push(i);
      }
    }
    
    if (totalColumns.length === 0) {
      throw new Error("Could not find any 'Total' column in the Excel file");
    } else if (totalColumns.length === 1) {
      totalColumnIndex = totalColumns[0];
    } else {
      totalColumnIndex = totalColumns[1]; 
    }
    
    const studentDataStartIndex = headerRowIndex + 2;
    const studentData = data.slice(studentDataStartIndex);
    
    const rows = Array.from(document.querySelectorAll('input[placeholder="StudentId"]'))
      .map(input => {
        const container = input.closest('.row');
        if (!container) return null;
        const marksInput = container.querySelector('input[placeholder="Marks"]');
        const marksContainer = marksInput ? marksInput.closest('.mat-mdc-text-field-wrapper') : null;
        return { studentIdInput: input, marksInput, marksContainer, container };
      }).filter(r => r && r.marksInput && r.marksContainer);
    
    const unmatchedStudents = [];
    let matchedCount = 0;
    let totalStudentsProcessed = 0;
    
    for (const row of studentData) {
      if (!row || row.length === 0) continue;
      
      const id = row[idColumnIndex];
      const totalMarks = row[totalColumnIndex];
      
      let name = 'Unknown';
      for (let i = idColumnIndex + 1; i < Math.min(idColumnIndex + 3, row.length); i++) {
        if (row[i] && String(row[i]).trim() && isNaN(row[i])) {
          name = row[i];
          break;
        }
      }
      
      if (!id || totalMarks === undefined || totalMarks === null || totalMarks === '') {
        continue;
      }
      
      totalStudentsProcessed++;
      
      const formRow = rows.find(r => r.studentIdInput.value.trim() === String(id).trim());
      
      if (formRow) {
        formRow.marksInput.value = totalMarks;
        
        formRow.marksInput.dispatchEvent(new Event('input', { bubbles: true }));

        setTimeout(() => {
          formRow.marksContainer.style.backgroundColor = '#dcfce7';
          formRow.marksContainer.style.borderColor = '#22c55e';
          formRow.marksContainer.style.borderRadius = '8px';
          formRow.marksContainer.style.transition = 'all 0.3s ease';
        }, 50);
        
        matchedCount++;
      } else {
        unmatchedStudents.push({
          id: id,
          name: name,
          finalMark: totalMarks
        });
      }
    }
    
    return {
      unmatchedStudents: unmatchedStudents,
      totalStudents: totalStudentsProcessed,
      matchedCount: matchedCount
    };
    
  } catch (error) {
    console.error("Error processing Excel data:", error);
    return {
      unmatchedStudents: [],
      totalStudents: 0,
      matchedCount: 0
    };
  }
}