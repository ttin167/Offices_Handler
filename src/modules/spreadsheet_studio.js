/**
 * src/modules/spreadsheet_studio.js ? Spreadsheet Studio Controller (12 Gates Validator, Mutator, Inspector)
 */

export function initSpreadsheetStudio({ showToast, refreshLucideIcons, uploadFileToServer, formatBytes }) {
// =========================================================================
// SPREADSHEET STUDIO CONTROLLER (12 Gates Validator, Mutator, Inspector)
// =========================================================================

// 1. Sub-tab switcher
const subTabBtns = document.querySelectorAll('.sub-tab-btn');
const subPanels = {
  validator: document.getElementById('subpanel-validator'),
  mutator: document.getElementById('subpanel-mutator'),
  inspector: document.getElementById('subpanel-inspector'),
  formulas: document.getElementById('subpanel-formulas'),
};

subTabBtns.forEach((btn) => {
  btn.addEventListener('click', () => {
    const targetSub = btn.getAttribute('data-subtab');
    if (!subPanels[targetSub]) return;

    subTabBtns.forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');

    Object.values(subPanels).forEach((p) => {
      if (p) p.style.display = 'none';
    });
    subPanels[targetSub].style.display = 'block';
  });
});

// 2. Sub-panel 1: 12 Quality Gates Validator
const xlsxValDropZone = document.getElementById('xlsx-val-drop-zone');
const xlsxValFileInput = document.getElementById('xlsx-val-file-input');
const btnBrowseXlsxVal = document.getElementById('btn-browse-xlsx-val');
const xlsxValFileInfo = document.getElementById('xlsx-val-file-info');
const xlsxValFileName = document.getElementById('xlsx-val-file-name');
const xlsxValFileSize = document.getElementById('xlsx-val-file-size');
const btnRemoveXlsxVal = document.getElementById('btn-remove-xlsx-val');
const btnRunXlsxValidate = document.getElementById('btn-run-xlsx-validate');
const xlsxValProgress = document.getElementById('xlsx-val-progress');
const xlsxValProgressFill = document.getElementById('xlsx-val-progress-fill');
const xlsxValProgressStatus = document.getElementById('xlsx-val-progress-status');
const xlsxValStatusBadge = document.getElementById('xlsx-val-status-badge');
const xlsxValEmptyState = document.getElementById('xlsx-val-empty-state');
const xlsxValResults = document.getElementById('xlsx-val-results');
const gatesStatusGrid = document.getElementById('gates-status-grid');
const discrepanciesBox = document.getElementById('discrepancies-box');
const discrepanciesList = document.getElementById('discrepancies-list');

let xlsxValFile = null;
let xlsxValServerPath = null;

function handleXlsxValFile(file) {
  if (!file) return;
  xlsxValFile = file;
  xlsxValServerPath = null;
  if (xlsxValFileName) xlsxValFileName.textContent = file.name;
  if (xlsxValFileSize) xlsxValFileSize.textContent = formatBytes(file.size);
  if (xlsxValDropZone) {
    const dropContent = xlsxValDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'none';
  }
  if (xlsxValFileInfo) xlsxValFileInfo.style.display = 'flex';
  if (btnRunXlsxValidate) btnRunXlsxValidate.disabled = false;
  showToast(`Đã chọn file Excel kiểm tra: ${file.name}`, 'info');
}

function clearXlsxValFile() {
  xlsxValFile = null;
  xlsxValServerPath = null;
  if (xlsxValFileInput) xlsxValFileInput.value = '';
  if (xlsxValFileInfo) xlsxValFileInfo.style.display = 'none';
  if (xlsxValDropZone) {
    const dropContent = xlsxValDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'block';
  }
  if (btnRunXlsxValidate) btnRunXlsxValidate.disabled = true;
}

if (btnBrowseXlsxVal && xlsxValFileInput) {
  btnBrowseXlsxVal.addEventListener('click', (e) => {
    e.stopPropagation();
    xlsxValFileInput.click();
  });
}

if (xlsxValDropZone && xlsxValFileInput) {
  xlsxValDropZone.addEventListener('click', () => xlsxValFileInput.click());
  xlsxValDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    xlsxValDropZone.classList.add('dragover');
  });
  xlsxValDropZone.addEventListener('dragleave', () => xlsxValDropZone.classList.remove('dragover'));
  xlsxValDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    xlsxValDropZone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleXlsxValFile(e.dataTransfer.files[0]);
    }
  });
  xlsxValFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleXlsxValFile(e.target.files[0]);
    }
  });
}

if (btnRemoveXlsxVal) {
  btnRemoveXlsxVal.addEventListener('click', (e) => {
    e.stopPropagation();
    clearXlsxValFile();
  });
}

const QUALITY_GATES = [
  { id: 1, name: 'Header Block', desc: 'Font, fill, border & alignment (R2-8)' },
  { id: 2, name: 'Row 7 KPI Formulas', desc: 'Dynamic COUNTIF/SUM formulas (E3)' },
  { id: 3, name: 'Row 9 TC Headers', desc: 'Navy fill & double border & alignment' },
  { id: 4, name: 'Col A Navy Fill', desc: 'Unbroken FF000080 continuous fill' },
  { id: 5, name: 'Col B-C-D Box', desc: 'Unified 3-column box & thin borders (E7)' },
  { id: 6, name: 'Duplicate Group Headers', desc: 'Prevents repeated group names (E8)' },
  { id: 7, name: 'Matrix Grid & O-Marks', desc: 'Grid border consistency & 8pt bold marks' },
  { id: 8, name: 'Result Footer Block', desc: 'Courier non-bold 8pt & mm/dd numfmt' },
  { id: 9, name: 'Merged Cells Integrity', desc: 'Header merges & footer B:D (ERR_XLSX_004)' },
  { id: 10, name: 'Outer Closing Borders', desc: 'Double bottom closing on sections' },
  { id: 11, name: 'Data Validations', desc: 'Dropdowns (O matrix, N/A/B, P/F)' },
  { id: 12, name: 'Geometry & Col Widths', desc: 'Column widths & row heights (E5)' },
];

if (btnRunXlsxValidate) {
  btnRunXlsxValidate.addEventListener('click', async () => {
    if (!xlsxValFile) return;

    btnRunXlsxValidate.disabled = true;
    if (xlsxValProgress) xlsxValProgress.style.display = 'block';
    if (xlsxValProgressFill) xlsxValProgressFill.style.width = '30%';
    if (xlsxValProgressStatus) xlsxValProgressStatus.textContent = 'Đang tải tệp Excel lên máy chủ...';

    try {
      const uploadRes = await uploadFileToServer(xlsxValFile);
      xlsxValServerPath = uploadRes.filepath || uploadRes.file_path || (uploadRes.first_file && uploadRes.first_file.file_path);

      if (xlsxValProgressFill) xlsxValProgressFill.style.width = '70%';
      if (xlsxValProgressStatus) xlsxValProgressStatus.textContent = 'Đang đối chiếu 12 Quality Gates...';

      const targetSheetInput = document.getElementById('xlsx-val-target-sheet');
      const refSheetInput = document.getElementById('xlsx-val-ref-sheet');

      const res = await fetch('/api/xlsx/validate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_path: xlsxValServerPath,
          target_sheet: targetSheetInput && targetSheetInput.value.trim() ? targetSheetInput.value.trim().replace(/^['"]|['"]$/g, '') : null,
          ref_sheet: refSheetInput && refSheetInput.value.trim() ? refSheetInput.value.trim().replace(/^['"]|['"]$/g, '') : 'Example',
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Kiểm tra thất bại (${res.status})`);
      }

      const data = await res.json();
      if (xlsxValProgressFill) xlsxValProgressFill.style.width = '100%';

      renderQualityGatesReport(data);
      showToast(data.passed ? '100% CLEAN! Đạt chuẩn 12 Quality Gates' : `Phát hiện ${data.diff_count} điểm lệch format!`, data.passed ? 'success' : 'error');

      setTimeout(() => {
        if (xlsxValProgress) xlsxValProgress.style.display = 'none';
      }, 1200);
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error', 6000);
      if (xlsxValProgressStatus) xlsxValProgressStatus.textContent = `Thất bại: ${err.message}`;
    } finally {
      btnRunXlsxValidate.disabled = false;
    }
  });
}

function renderQualityGatesReport(data) {
  if (xlsxValEmptyState) xlsxValEmptyState.style.display = 'none';
  if (xlsxValResults) xlsxValResults.style.display = 'block';

  if (xlsxValStatusBadge) {
    if (data.passed) {
      xlsxValStatusBadge.className = 'deck-count-badge badge-passed';
      xlsxValStatusBadge.textContent = '100% CLEAN (PASS)';
    } else {
      xlsxValStatusBadge.className = 'deck-count-badge badge-failed';
      xlsxValStatusBadge.textContent = `${data.diff_count} ĐIỂM LỆCH (${data.gates_failed} CỔNG FAIL)`;
    }
  }

  if (gatesStatusGrid) {
    gatesStatusGrid.innerHTML = '';
    QUALITY_GATES.forEach((gate) => {
      const gateDiffs = (data.diffs || []).filter((d) => d.includes(`[Gate ${gate.id} `) || d.includes(`[Gate ${gate.id}-`));
      const gatePassed = gateDiffs.length === 0;

      const gateCard = document.createElement('div');
      gateCard.className = `gate-card ${gatePassed ? 'gate-passed' : 'gate-failed'}`;
      gateCard.innerHTML = `
        <div class="gate-card-header">
          <span class="gate-id-badge">CỔNG ${gate.id}</span>
          <span class="gate-status-pill ${gatePassed ? 'pill-passed' : 'pill-failed'}">${gatePassed ? '✓ ĐẠT' : '✕ LỆCH (' + gateDiffs.length + ')'}</span>
        </div>
        <div class="gate-card-name">${gate.name}</div>
        <div class="gate-card-desc">${gate.desc}</div>
      `;
      gatesStatusGrid.appendChild(gateCard);
    });
  }

  if (discrepanciesList && discrepanciesBox) {
    discrepanciesList.innerHTML = '';
    if (!data.diffs || data.diffs.length === 0) {
      discrepanciesBox.style.display = 'none';
    } else {
      discrepanciesBox.style.display = 'block';
      data.diffs.forEach((diff) => {
        const item = document.createElement('div');
        item.className = 'discrepancy-item';
        item.textContent = diff.trim();
        discrepanciesList.appendChild(item);
      });
    }
  }
}

// 3. Sub-panel 2: Template Mutator
const xlsxMutDropZone = document.getElementById('xlsx-mut-drop-zone');
const xlsxMutFileInput = document.getElementById('xlsx-mut-file-input');
const btnBrowseXlsxMut = document.getElementById('btn-browse-xlsx-mut');
const xlsxMutFileInfo = document.getElementById('xlsx-mut-file-info');
const xlsxMutFileName = document.getElementById('xlsx-mut-file-name');
const xlsxMutFileSize = document.getElementById('xlsx-mut-file-size');
const btnRemoveXlsxMut = document.getElementById('btn-remove-xlsx-mut');
const xlsxMutSpecInput = document.getElementById('xlsx-mut-spec-input');
const btnPresetExpandRows = document.getElementById('btn-preset-expand-rows');
const btnPresetExpandCols = document.getElementById('btn-preset-expand-cols');
const btnRunXlsxMutate = document.getElementById('btn-run-xlsx-mutate');
const xlsxMutProgress = document.getElementById('xlsx-mut-progress');
const xlsxMutProgressFill = document.getElementById('xlsx-mut-progress-fill');
const xlsxMutProgressStatus = document.getElementById('xlsx-mut-progress-status');
const xlsxMutResultBox = document.getElementById('xlsx-mut-result-box');
const xlsxMutEmptyState = document.getElementById('xlsx-mut-empty-state');
const xlsxMutResultFilename = document.getElementById('xlsx-mut-result-filename');
const xlsxMutResultTime = document.getElementById('xlsx-mut-result-time');
const btnDownloadMutated = document.getElementById('btn-download-mutated');

let xlsxMutFile = null;
let xlsxMutServerPath = null;

const PRESET_MUT_EXPAND_ROWS = {
  target_sheet: "F1",
  ref_sheet: "Example",
  expand_rows: [
    { insert_after_row: 15, count: 10, clone_prototype_row: 14 }
  ],
  relink_charts: true
};

const PRESET_MUT_EXPAND_COLS = {
  target_sheet: "F1",
  ref_sheet: "Example",
  column_expansions: [
    { insert_after_col: "H", count: 5, clone_prototype_col: "H" }
  ],
  relink_charts: true
};

if (xlsxMutSpecInput) {
  xlsxMutSpecInput.value = JSON.stringify(PRESET_MUT_EXPAND_ROWS, null, 2);
}

if (btnPresetExpandRows && xlsxMutSpecInput) {
  btnPresetExpandRows.addEventListener('click', () => {
    xlsxMutSpecInput.value = JSON.stringify(PRESET_MUT_EXPAND_ROWS, null, 2);
  });
}

if (btnPresetExpandCols && xlsxMutSpecInput) {
  btnPresetExpandCols.addEventListener('click', () => {
    xlsxMutSpecInput.value = JSON.stringify(PRESET_MUT_EXPAND_COLS, null, 2);
  });
}

const btnPresetCellPatch = document.getElementById('btn-preset-cell-patch');
const PRESET_MUT_CELL_PATCH = {
  target_sheet: "F1",
  cell_patches: [
    { cell: "C7", value: 100 },
    { cell: "D7", formula: "=SUM(D2:D6)" },
    { cell: "E7", value: "Hoàn tất kiểm định" }
  ]
};

if (btnPresetCellPatch && xlsxMutSpecInput) {
  btnPresetCellPatch.addEventListener('click', () => {
    xlsxMutSpecInput.value = JSON.stringify(PRESET_MUT_CELL_PATCH, null, 2);
  });
}

async function handleXlsxMutFile(file) {
  if (!file) return;
  xlsxMutFile = file;
  xlsxMutServerPath = null;
  if (xlsxMutFileName) xlsxMutFileName.textContent = file.name;
  if (xlsxMutFileSize) xlsxMutFileSize.textContent = formatBytes(file.size);
  if (xlsxMutDropZone) {
    const dropContent = xlsxMutDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'none';
  }
  if (xlsxMutFileInfo) xlsxMutFileInfo.style.display = 'flex';
  if (btnRunXlsxMutate) btnRunXlsxMutate.disabled = false;
  showToast(`Đã chọn template Excel: ${file.name}`, 'info');

  // Auto-discover sheets to assist user
  try {
    const up = await uploadFileToServer(file);
    xlsxMutServerPath = up.filepath || up.file_path || (up.first_file && up.first_file.file_path);
    const readRes = await fetch('/api/xlsx/read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: xlsxMutServerPath, engine: 'openpyxl' })
    });
    if (readRes.ok) {
      const readData = await readRes.json();
      const sheetNames = readData.sheet_names || [];
      if (sheetNames.length > 0) {
        showToast(`Tìm thấy ${sheetNames.length} trang tính: ${sheetNames.slice(0, 3).join(', ')}${sheetNames.length > 3 ? '...' : ''}`, 'info', 3000);
      }
    }
  } catch (e) {
    // Non-blocking discovery
  }
}

function clearXlsxMutFile() {
  xlsxMutFile = null;
  xlsxMutServerPath = null;
  if (xlsxMutFileInput) xlsxMutFileInput.value = '';
  if (xlsxMutFileInfo) xlsxMutFileInfo.style.display = 'none';
  if (xlsxMutDropZone) {
    const dropContent = xlsxMutDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'block';
  }
  if (btnRunXlsxMutate) btnRunXlsxMutate.disabled = true;
}

if (btnBrowseXlsxMut && xlsxMutFileInput) {
  btnBrowseXlsxMut.addEventListener('click', (e) => {
    e.stopPropagation();
    xlsxMutFileInput.click();
  });
}

if (xlsxMutDropZone && xlsxMutFileInput) {
  xlsxMutDropZone.addEventListener('click', () => xlsxMutFileInput.click());
  xlsxMutDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    xlsxMutDropZone.classList.add('dragover');
  });
  xlsxMutDropZone.addEventListener('dragleave', () => xlsxMutDropZone.classList.remove('dragover'));
  xlsxMutDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    xlsxMutDropZone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleXlsxMutFile(e.dataTransfer.files[0]);
    }
  });
  xlsxMutFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleXlsxMutFile(e.target.files[0]);
    }
  });
}

if (btnRemoveXlsxMut) {
  btnRemoveXlsxMut.addEventListener('click', (e) => {
    e.stopPropagation();
    clearXlsxMutFile();
  });
}

if (btnRunXlsxMutate) {
  btnRunXlsxMutate.addEventListener('click', async () => {
    if (!xlsxMutFile) return;

    const specRaw = xlsxMutSpecInput ? xlsxMutSpecInput.value.trim() : '';
    let specObj = {};
    try {
      specObj = specRaw ? JSON.parse(specRaw) : {};
    } catch (e) {
      showToast('Cú pháp JSON spec biến đổi không hợp lệ: ' + e.message, 'error');
      return;
    }

    btnRunXlsxMutate.disabled = true;
    if (xlsxMutProgress) xlsxMutProgress.style.display = 'block';
    if (xlsxMutProgressFill) xlsxMutProgressFill.style.width = '30%';
    if (xlsxMutProgressStatus) xlsxMutProgressStatus.textContent = 'Đang tải template Excel lên máy chủ...';

    try {
      const uploadRes = await uploadFileToServer(xlsxMutFile);
      xlsxMutServerPath = uploadRes.filepath || uploadRes.file_path || (uploadRes.first_file && uploadRes.first_file.file_path);

      if (xlsxMutProgressFill) xlsxMutProgressFill.style.width = '65%';
      if (xlsxMutProgressStatus) xlsxMutProgressStatus.textContent = 'Đang thực hiện mở rộng dòng/cột & tái liên kết biểu đồ...';

      const res = await fetch('/api/xlsx/mutate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          template_path: xlsxMutServerPath,
          spec: specObj,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Biến đổi template thất bại (${res.status})`);
      }

      const data = await res.json();
      if (xlsxMutProgressFill) xlsxMutProgressFill.style.width = '100%';

      if (xlsxMutResultFilename) xlsxMutResultFilename.textContent = data.filename || 'mutated.xlsx';
      if (xlsxMutResultTime) xlsxMutResultTime.textContent = new Date().toLocaleTimeString();
      if (btnDownloadMutated) btnDownloadMutated.setAttribute('href', data.download_url || `/api/download/${data.filename}`);

      if (xlsxMutEmptyState) xlsxMutEmptyState.style.display = 'none';
      if (xlsxMutResultBox) xlsxMutResultBox.style.display = 'block';

      showToast('Biến đổi template Excel thành công (Invariant E12)!', 'success');

      setTimeout(() => {
        if (xlsxMutProgress) xlsxMutProgress.style.display = 'none';
      }, 1200);
    } catch (err) {
      showToast(`Lỗi biến đổi: ${err.message}`, 'error', 6000);
      if (xlsxMutProgressStatus) xlsxMutProgressStatus.textContent = `Thất bại: ${err.message}`;
    } finally {
      btnRunXlsxMutate.disabled = false;
    }
  });
}

// 4. Sub-panel 3: JSON Snapshot Inspector
const xlsxInsDropZone = document.getElementById('xlsx-ins-drop-zone');
const xlsxInsFileInput = document.getElementById('xlsx-ins-file-input');
const btnBrowseXlsxIns = document.getElementById('btn-browse-xlsx-ins');
const xlsxInsFileInfo = document.getElementById('xlsx-ins-file-info');
const xlsxInsFileName = document.getElementById('xlsx-ins-file-name');
const xlsxInsFileSize = document.getElementById('xlsx-ins-file-size');
const btnRemoveXlsxIns = document.getElementById('btn-remove-xlsx-ins');
const xlsxInsEngineSelect = document.getElementById('xlsx-ins-engine-select');
const xlsxInsSheetName = document.getElementById('xlsx-ins-sheet-name');
const btnRunXlsxInspect = document.getElementById('btn-run-xlsx-inspect');
const xlsxInsSummaryBadge = document.getElementById('xlsx-ins-summary-badge');
const xlsxInsEmptyState = document.getElementById('xlsx-ins-empty-state');
const xlsxInsViewContainer = document.getElementById('xlsx-ins-view-container');
const xlsxInsStatsBar = document.getElementById('xlsx-ins-stats-bar');
const xlsxInsJsonOutput = document.getElementById('xlsx-ins-json-output');

let xlsxInsFile = null;
let xlsxInsServerPath = null;

function handleXlsxInsFile(file) {
  if (!file) return;
  xlsxInsFile = file;
  xlsxInsServerPath = null;
  if (xlsxInsFileName) xlsxInsFileName.textContent = file.name;
  if (xlsxInsFileSize) xlsxInsFileSize.textContent = formatBytes(file.size);
  if (xlsxInsDropZone) {
    const dropContent = xlsxInsDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'none';
  }
  if (xlsxInsFileInfo) xlsxInsFileInfo.style.display = 'flex';
  if (btnRunXlsxInspect) btnRunXlsxInspect.disabled = false;
  showToast(`Đã chọn file soi snapshot: ${file.name}`, 'info');
}

function clearXlsxInsFile() {
  xlsxInsFile = null;
  xlsxInsServerPath = null;
  if (xlsxInsFileInput) xlsxInsFileInput.value = '';
  if (xlsxInsFileInfo) xlsxInsFileInfo.style.display = 'none';
  if (xlsxInsDropZone) {
    const dropContent = xlsxInsDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'block';
  }
  if (btnRunXlsxInspect) btnRunXlsxInspect.disabled = true;
}

if (btnBrowseXlsxIns && xlsxInsFileInput) {
  btnBrowseXlsxIns.addEventListener('click', (e) => {
    e.stopPropagation();
    xlsxInsFileInput.click();
  });
}

if (xlsxInsDropZone && xlsxInsFileInput) {
  xlsxInsDropZone.addEventListener('click', () => xlsxInsFileInput.click());
  xlsxInsDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    xlsxInsDropZone.classList.add('dragover');
  });
  xlsxInsDropZone.addEventListener('dragleave', () => xlsxInsDropZone.classList.remove('dragover'));
  xlsxInsDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    xlsxInsDropZone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleXlsxInsFile(e.dataTransfer.files[0]);
    }
  });
  xlsxInsFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleXlsxInsFile(e.target.files[0]);
    }
  });
}

if (btnRemoveXlsxIns) {
  btnRemoveXlsxIns.addEventListener('click', (e) => {
    e.stopPropagation();
    clearXlsxInsFile();
  });
}

if (btnRunXlsxInspect) {
  btnRunXlsxInspect.addEventListener('click', async () => {
    if (!xlsxInsFile) return;

    btnRunXlsxInspect.disabled = true;
    btnRunXlsxInspect.innerHTML = '<span>⏳ ĐANG TRÍCH XUẤT SNAPSHOT...</span>';

    try {
      const uploadRes = await uploadFileToServer(xlsxInsFile);
      xlsxInsServerPath = uploadRes.filepath || uploadRes.file_path || (uploadRes.first_file && uploadRes.first_file.file_path);

      const engine = xlsxInsEngineSelect ? xlsxInsEngineSelect.value : 'auto';
      const sheet = xlsxInsSheetName && xlsxInsSheetName.value.trim() ? xlsxInsSheetName.value.trim() : null;

      const res = await fetch('/api/xlsx/read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_path: xlsxInsServerPath,
          engine: engine,
          sheet_name: sheet,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Trích xuất snapshot thất bại (${res.status})`);
      }

      const data = await res.json();
      const snapshot = data.snapshot || {};

      if (xlsxInsEmptyState) xlsxInsEmptyState.style.display = 'none';
      if (xlsxInsViewContainer) xlsxInsViewContainer.style.display = 'block';

      if (xlsxInsJsonOutput) {
        xlsxInsJsonOutput.value = JSON.stringify(snapshot, null, 2);
      }

      const sheets = snapshot.sheets || (snapshot.sheet_name ? [snapshot] : []);
      let totalCells = 0;
      let totalMerges = 0;
      sheets.forEach((s) => {
        totalCells += (s.cells || []).length;
        totalMerges += (s.merged_cells || []).length;
      });

      if (xlsxInsSummaryBadge) {
        xlsxInsSummaryBadge.textContent = `${sheets.length} Sheets | ${totalCells} Cells | Engine: ${data.engine}`;
      }

      if (xlsxInsStatsBar) {
        xlsxInsStatsBar.innerHTML = `
          <div class="stat-pill"><span class="stat-lbl">Số trang tính:</span> <strong>${sheets.length}</strong></div>
          <div class="stat-pill"><span class="stat-lbl">Tổng ô có dữ liệu:</span> <strong>${totalCells}</strong></div>
          <div class="stat-pill"><span class="stat-lbl">Dải ô hợp nhất:</span> <strong>${totalMerges}</strong></div>
          <div class="stat-pill"><span class="stat-lbl">Động cơ:</span> <strong>${data.engine}</strong></div>
        `;
      }

      showToast('Đã trích xuất JSON snapshot thành công!', 'success');
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error', 6000);
    } finally {
      btnRunXlsxInspect.disabled = false;
      btnRunXlsxInspect.innerHTML = '<span class="btn-icon"><i data-lucide="scan-search"></i></span><span>TRÍCH XUẤT JSON SNAPSHOT</span>';
      refreshLucideIcons();
    }
  });
}

}
