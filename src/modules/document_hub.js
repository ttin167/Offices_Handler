/**
 * src/modules/document_hub.js ? Universal Document Hub, PDF Pro Suite, Security, OCR & Quick Patch Controller
 */

export function initDocumentHub({ state, showToast, refreshLucideIcons, uploadFileToServer, formatBytes, loadHistory }) {
// =========================================================================
// 7. TAB 2: UNIVERSAL DOCUMENT HUB CONTROLLER
// =========================================================================
const docDropZone = document.getElementById('doc-drop-zone');
const docFileInput = document.getElementById('doc-file-input');
const btnBrowseDocFile = document.getElementById('btn-browse-doc-file');
const docFileInfo = document.getElementById('doc-file-info');
const docFileName = document.getElementById('doc-file-name');
const docFileSize = document.getElementById('doc-file-size');
const docFileExt = document.getElementById('doc-file-ext');
const btnRemoveDocFile = document.getElementById('btn-remove-doc-file');
const btnConvertDoc = document.getElementById('btn-convert-doc');
const docProgress = document.getElementById('doc-progress-container');
const docProgressFill = document.getElementById('doc-progress-fill');
const docProgressStatus = document.getElementById('doc-progress-status');
const docResultBox = document.getElementById('doc-result-box');
const docEmptyState = document.getElementById('doc-empty-state');
const docResultFilename = document.getElementById('doc-result-filename');
const docResultTime = document.getElementById('doc-result-time');
const btnDownloadConverted = document.getElementById('btn-download-converted');

function handleDocFileSelected(file) {
  if (!file) return;
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!['.pdf', '.docx', '.md', '.xlsx'].includes(ext)) {
    showToast('Định dạng tệp không được hỗ trợ trong Document Hub', 'error');
    return;
  }

  state.docFile = file;
  state.docServerPath = null;

  if (docFileName) docFileName.textContent = file.name;
  if (docFileSize) docFileSize.textContent = formatBytes(file.size);
  if (docFileExt) docFileExt.textContent = ext.replace('.', '').toUpperCase();

  if (docDropZone) {
    const dropContent = docDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'none';
  }
  if (docFileInfo) docFileInfo.style.display = 'flex';
  if (btnConvertDoc) btnConvertDoc.disabled = false;
  showToast(`Đã chọn tài liệu: ${file.name}`, 'info');
}

function clearDocFile() {
  state.docFile = null;
  state.docServerPath = null;
  if (docFileInput) docFileInput.value = '';
  if (docFileInfo) docFileInfo.style.display = 'none';
  if (docDropZone) {
    const dropContent = docDropZone.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'block';
  }
  if (btnConvertDoc) btnConvertDoc.disabled = true;
}

if (btnBrowseDocFile && docFileInput) {
  btnBrowseDocFile.addEventListener('click', (e) => {
    e.stopPropagation();
    docFileInput.click();
  });
}

if (docDropZone && docFileInput) {
  docDropZone.addEventListener('click', () => docFileInput.click());
  docDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    docDropZone.classList.add('dragover');
  });
  docDropZone.addEventListener('dragleave', () => {
    docDropZone.classList.remove('dragover');
  });
  docDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    docDropZone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleDocFileSelected(e.dataTransfer.files[0]);
    }
  });
  docFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleDocFileSelected(e.target.files[0]);
    }
  });
}

if (btnRemoveDocFile) {
  btnRemoveDocFile.addEventListener('click', (e) => {
    e.stopPropagation();
    clearDocFile();
  });
}

// Target Format Matrix Selection
const matrixCards = document.querySelectorAll('.matrix-card');
const xlsxExtractionHint = document.getElementById('xlsx-extraction-hint');
matrixCards.forEach((card) => {
  card.addEventListener('click', () => {
    matrixCards.forEach((c) => c.classList.remove('active'));
    card.classList.add('active');
    const radio = card.querySelector('input[type="radio"]');
    if (radio) radio.checked = true;
    state.targetFormat = card.getAttribute('data-target') || '.docx';
    if (xlsxExtractionHint) {
      xlsxExtractionHint.style.display = state.targetFormat === '.xlsx' ? 'flex' : 'none';
    }
  });
});

// Convert Document Action
if (btnConvertDoc) {
  btnConvertDoc.addEventListener('click', async () => {
    if (!state.docFile) return;

    btnConvertDoc.disabled = true;
    if (docProgress) docProgress.style.display = 'block';
    if (docProgressFill) docProgressFill.style.width = '30%';
    if (docProgressStatus) docProgressStatus.textContent = 'Đang tải lên tài liệu...';

    try {
      const uploadRes = await uploadFileToServer(state.docFile);
      state.docServerPath = uploadRes.filepath || uploadRes.file_path || (uploadRes.first_file && uploadRes.first_file.file_path);

      if (docProgressFill) docProgressFill.style.width = '65%';
      if (docProgressStatus) docProgressStatus.textContent = `Đang chuyển đổi sang ${state.targetFormat.toUpperCase()}...`;

      const res = await fetch('/api/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          source_path: state.docServerPath,
          input_path: state.docServerPath,
          filepath: state.docServerPath,
          target_ext: state.targetFormat,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Chuyển đổi thất bại: ${res.statusText}`);
      }

      const data = await res.json();
      if (!data.success) {
        throw new Error(data.error || data.result_path || 'Chuyển đổi tài liệu thất bại');
      }
      if (docProgressFill) docProgressFill.style.width = '100%';
      if (docProgressStatus) docProgressStatus.textContent = 'Hoàn tất!';

      const outFileName = data.output_file || data.filename || (data.result_path ? data.result_path.split(/[\\/]/).pop() : '');
      if (docResultFilename) docResultFilename.textContent = outFileName;
      if (docResultTime) docResultTime.textContent = new Date().toLocaleTimeString();
      if (btnDownloadConverted) {
        btnDownloadConverted.setAttribute('href', data.download_url || `/api/download/${outFileName}`);
      }

      if (docEmptyState) docEmptyState.style.display = 'none';
      if (docResultBox) docResultBox.style.display = 'block';

      showToast('Chuyển đổi tài liệu thành công!', 'success');

      setTimeout(() => {
        if (docProgress) docProgress.style.display = 'none';
      }, 1200);
    } catch (err) {
      showToast(`Lỗi chuyển đổi: ${err.message}`, 'error', 6000);
      if (docProgressStatus) docProgressStatus.textContent = `Lỗi: ${err.message}`;
    } finally {
      btnConvertDoc.disabled = false;
    }
  });
}

// =========================================================================
// 7.1 DOCUMENT HUB SUB-TABS & EXTENDED CONTROLLERS (v4.0)
// =========================================================================
const docSubTabBtns = document.querySelectorAll('#doc-subtabs-nav .sub-tab-btn');
const docSubpanels = document.querySelectorAll('.doc-subpanel');
docSubTabBtns.forEach((btn) => {
  btn.addEventListener('click', () => {
    const target = btn.getAttribute('data-doc-subtab');
    docSubTabBtns.forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    docSubpanels.forEach((p) => {
      if (p.id === `doc-subpanel-${target}`) {
        p.style.display = 'block';
        p.classList.add('active');
      } else {
        p.style.display = 'none';
        p.classList.remove('active');
      }
    });
    refreshLucideIcons();
  });
});

// -------------------------------------------------------------------------
// SUB-TAB 2: BATCH MAIL MERGE CONTROLLER (Edge Case 3: Dual Dropzones)
// -------------------------------------------------------------------------
let mmTplFile = null;
let mmDataFile = null;

const mmTplDropZone = document.getElementById('mm-tpl-drop-zone');
const mmTplFileInput = document.getElementById('mm-tpl-file-input');
const btnBrowseMmTpl = document.getElementById('btn-browse-mm-tpl');
const mmTplFileInfo = document.getElementById('mm-tpl-file-info');
const mmTplFileName = document.getElementById('mm-tpl-file-name');
const mmTplFileSize = document.getElementById('mm-tpl-file-size');
const btnRemoveMmTpl = document.getElementById('btn-remove-mm-tpl');

const mmDataDropZone = document.getElementById('mm-data-drop-zone');
const mmDataFileInput = document.getElementById('mm-data-file-input');
const btnBrowseMmData = document.getElementById('btn-browse-mm-data');
const mmDataFileInfo = document.getElementById('mm-data-file-info');
const mmDataFileName = document.getElementById('mm-data-file-name');
const mmDataFileSize = document.getElementById('mm-data-file-size');
const mmDataExt = document.getElementById('mm-data-ext');
const btnRemoveMmData = document.getElementById('btn-remove-mm-data');

const mmTokensCard = document.getElementById('mm-tokens-card');
const mmTokensList = document.getElementById('mm-tokens-list');
const btnRunMailMerge = document.getElementById('btn-run-mail-merge');
const mmProgress = document.getElementById('mm-progress-container');
const mmProgressFill = document.getElementById('mm-progress-fill');
const mmProgressStatus = document.getElementById('mm-progress-status');
const mmResultBox = document.getElementById('mm-result-box');
const mmEmptyState = document.getElementById('mm-empty-state');
const mmResultCount = document.getElementById('mm-result-count');
const mmResultFilename = document.getElementById('mm-result-filename');
const btnDownloadMmZip = document.getElementById('btn-download-mm-zip');

function checkMmReady() {
  if (btnRunMailMerge) {
    btnRunMailMerge.disabled = !(mmTplFile && mmDataFile);
  }
}

async function inspectMmTemplate(file) {
  try {
    const uploadRes = await uploadFileToServer(file);
    const serverPath = uploadRes.filepath || uploadRes.file_path;
    if (!serverPath) return;

    const res = await fetch('/api/mail-merge/inspect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template: serverPath }),
    });
    if (res.ok) {
      const data = await res.json();
      if (data.tokens && data.tokens.length > 0 && mmTokensCard && mmTokensList) {
        mmTokensList.innerHTML = data.tokens
          .map((t) => `<span class="token-pill"><i data-lucide="tag"></i> {{${t}}}</span>`)
          .join('');
        mmTokensCard.style.display = 'block';
        refreshLucideIcons();
      }
    }
  } catch (_) {}
}

if (mmTplDropZone && mmTplFileInput) {
  mmTplDropZone.addEventListener('click', () => mmTplFileInput.click());
  if (btnBrowseMmTpl) btnBrowseMmTpl.addEventListener('click', (e) => { e.stopPropagation(); mmTplFileInput.click(); });
  mmTplFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      const f = e.target.files[0];
      mmTplFile = f;
      if (mmTplFileName) mmTplFileName.textContent = f.name;
      if (mmTplFileSize) mmTplFileSize.textContent = formatBytes(f.size);
      const content = mmTplDropZone.querySelector('.drop-zone-content');
      if (content) content.style.display = 'none';
      if (mmTplFileInfo) mmTplFileInfo.style.display = 'flex';
      checkMmReady();
      inspectMmTemplate(f);
    }
  });
  if (btnRemoveMmTpl) {
    btnRemoveMmTpl.addEventListener('click', (e) => {
      e.stopPropagation();
      mmTplFile = null;
      mmTplFileInput.value = '';
      if (mmTplFileInfo) mmTplFileInfo.style.display = 'none';
      const content = mmTplDropZone.querySelector('.drop-zone-content');
      if (content) content.style.display = 'block';
      if (mmTokensCard) mmTokensCard.style.display = 'none';
      checkMmReady();
    });
  }
}

if (mmDataDropZone && mmDataFileInput) {
  mmDataDropZone.addEventListener('click', () => mmDataFileInput.click());
  if (btnBrowseMmData) btnBrowseMmData.addEventListener('click', (e) => { e.stopPropagation(); mmDataFileInput.click(); });
  mmDataFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      const f = e.target.files[0];
      mmDataFile = f;
      const ext = f.name.substring(f.name.lastIndexOf('.')).replace('.', '').toUpperCase();
      if (mmDataFileName) mmDataFileName.textContent = f.name;
      if (mmDataFileSize) mmDataFileSize.textContent = formatBytes(f.size);
      if (mmDataExt) mmDataExt.textContent = ext;
      const content = mmDataDropZone.querySelector('.drop-zone-content');
      if (content) content.style.display = 'none';
      if (mmDataFileInfo) mmDataFileInfo.style.display = 'flex';
      checkMmReady();
    }
  });
  if (btnRemoveMmData) {
    btnRemoveMmData.addEventListener('click', (e) => {
      e.stopPropagation();
      mmDataFile = null;
      mmDataFileInput.value = '';
      if (mmDataFileInfo) mmDataFileInfo.style.display = 'none';
      const content = mmDataDropZone.querySelector('.drop-zone-content');
      if (content) content.style.display = 'block';
      checkMmReady();
    });
  }
}

if (btnRunMailMerge) {
  btnRunMailMerge.addEventListener('click', async () => {
    if (!mmTplFile || !mmDataFile) return;
    btnRunMailMerge.disabled = true;
    if (mmProgress) mmProgress.style.display = 'block';
    if (mmProgressFill) mmProgressFill.style.width = '30%';
    if (mmProgressStatus) mmProgressStatus.textContent = 'Đang tải lên template và dataset...';

    try {
      const [upTpl, upData] = await Promise.all([
        uploadFileToServer(mmTplFile),
        uploadFileToServer(mmDataFile),
      ]);
      const tplPath = upTpl.filepath || upTpl.file_path;
      const dataPath = upData.filepath || upData.file_path;

      if (mmProgressFill) mmProgressFill.style.width = '65%';
      if (mmProgressStatus) mmProgressStatus.textContent = 'Đang kết hợp dữ liệu và sinh tài liệu hàng loạt...';

      const patInput = document.getElementById('mm-pattern-input');
      const zipToggle = document.getElementById('mm-zip-toggle');
      const pattern = patInput && patInput.value.trim() ? patInput.value.trim() : null;
      const wantZip = zipToggle ? zipToggle.checked : true;

      const res = await fetch('/api/mail-merge', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          template: tplPath,
          data: dataPath,
          pattern: pattern,
          zip: wantZip ? `mail_merge_${Date.now()}.zip` : null,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'Trộn thư thất bại');
      }

      const data = await res.json();
      if (mmProgressFill) mmProgressFill.style.width = '100%';
      if (mmProgressStatus) mmProgressStatus.textContent = 'Hoàn tất!';

      if (mmResultCount) mmResultCount.textContent = data.count || data.total_generated || (data.files ? data.files.length : 'N/A');
      const zName = data.zip_path ? data.zip_path.split(/[\\/]/).pop() : (data.download_url ? data.download_url.split('/').pop() : 'mail_merge.zip');
      if (mmResultFilename) mmResultFilename.textContent = zName;
      if (btnDownloadMmZip) btnDownloadMmZip.setAttribute('href', data.download_url || `/api/download/${zName}`);

      if (mmEmptyState) mmEmptyState.style.display = 'none';
      if (mmResultBox) mmResultBox.style.display = 'block';
      showToast('Trộn thư hàng loạt thành công!', 'success');
    } catch (err) {
      showToast(`Lỗi trộn thư: ${err.message}`, 'error', 6000);
    } finally {
      btnRunMailMerge.disabled = false;
      setTimeout(() => { if (mmProgress) mmProgress.style.display = 'none'; }, 1500);
    }
  });
}

// -------------------------------------------------------------------------
// SUB-TAB 3: LEGAL REDLINE DIFF CONTROLLER (Edge Case 3: Dual Dropzones)
// -------------------------------------------------------------------------
let diffV1File = null;
let diffV2File = null;

const diffV1DropZone = document.getElementById('diff-v1-drop-zone');
const diffV1FileInput = document.getElementById('diff-v1-file-input');
const btnBrowseDiffV1 = document.getElementById('btn-browse-diff-v1');
const diffV1FileInfo = document.getElementById('diff-v1-file-info');
const diffV1FileName = document.getElementById('diff-v1-file-name');
const diffV1FileSize = document.getElementById('diff-v1-file-size');
const btnRemoveDiffV1 = document.getElementById('btn-remove-diff-v1');

const diffV2DropZone = document.getElementById('diff-v2-drop-zone');
const diffV2FileInput = document.getElementById('diff-v2-file-input');
const btnBrowseDiffV2 = document.getElementById('btn-browse-diff-v2');
const diffV2FileInfo = document.getElementById('diff-v2-file-info');
const diffV2FileName = document.getElementById('diff-v2-file-name');
const diffV2FileSize = document.getElementById('diff-v2-file-size');
const btnRemoveDiffV2 = document.getElementById('btn-remove-diff-v2');

const btnRunDocDiff = document.getElementById('btn-run-doc-diff');
const diffProgress = document.getElementById('diff-progress-container');
const diffProgressFill = document.getElementById('diff-progress-fill');
const diffProgressStatus = document.getElementById('diff-progress-status');
const diffResultBox = document.getElementById('diff-result-box');
const diffEmptyState = document.getElementById('diff-empty-state');
const diffResInsertions = document.getElementById('diff-res-insertions');
const diffResDeletions = document.getElementById('diff-res-deletions');
const diffResModParas = document.getElementById('diff-res-mod-paras');
const diffResultFilename = document.getElementById('diff-result-filename');
const btnDownloadDiff = document.getElementById('btn-download-diff');

function checkDiffReady() {
  if (btnRunDocDiff) btnRunDocDiff.disabled = !(diffV1File && diffV2File);
}

if (diffV1DropZone && diffV1FileInput) {
  diffV1DropZone.addEventListener('click', () => diffV1FileInput.click());
  if (btnBrowseDiffV1) btnBrowseDiffV1.addEventListener('click', (e) => { e.stopPropagation(); diffV1FileInput.click(); });
  diffV1FileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      diffV1File = e.target.files[0];
      if (diffV1FileName) diffV1FileName.textContent = diffV1File.name;
      if (diffV1FileSize) diffV1FileSize.textContent = formatBytes(diffV1File.size);
      const c = diffV1DropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'none';
      if (diffV1FileInfo) diffV1FileInfo.style.display = 'flex';
      checkDiffReady();
    }
  });
  if (btnRemoveDiffV1) {
    btnRemoveDiffV1.addEventListener('click', (e) => {
      e.stopPropagation();
      diffV1File = null;
      diffV1FileInput.value = '';
      if (diffV1FileInfo) diffV1FileInfo.style.display = 'none';
      const c = diffV1DropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'block';
      checkDiffReady();
    });
  }
}

if (diffV2DropZone && diffV2FileInput) {
  diffV2DropZone.addEventListener('click', () => diffV2FileInput.click());
  if (btnBrowseDiffV2) btnBrowseDiffV2.addEventListener('click', (e) => { e.stopPropagation(); diffV2FileInput.click(); });
  diffV2FileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      diffV2File = e.target.files[0];
      if (diffV2FileName) diffV2FileName.textContent = diffV2File.name;
      if (diffV2FileSize) diffV2FileSize.textContent = formatBytes(diffV2File.size);
      const c = diffV2DropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'none';
      if (diffV2FileInfo) diffV2FileInfo.style.display = 'flex';
      checkDiffReady();
    }
  });
  if (btnRemoveDiffV2) {
    btnRemoveDiffV2.addEventListener('click', (e) => {
      e.stopPropagation();
      diffV2File = null;
      diffV2FileInput.value = '';
      if (diffV2FileInfo) diffV2FileInfo.style.display = 'none';
      const c = diffV2DropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'block';
      checkDiffReady();
    });
  }
}

if (btnRunDocDiff) {
  btnRunDocDiff.addEventListener('click', async () => {
    if (!diffV1File || !diffV2File) return;
    btnRunDocDiff.disabled = true;
    if (diffProgress) diffProgress.style.display = 'block';
    if (diffProgressFill) diffProgressFill.style.width = '30%';
    if (diffProgressStatus) diffProgressStatus.textContent = 'Đang tải lên 2 văn bản v1 và v2...';

    try {
      const [upV1, upV2] = await Promise.all([
        uploadFileToServer(diffV1File),
        uploadFileToServer(diffV2File),
      ]);
      const v1Path = upV1.filepath || upV1.file_path;
      const v2Path = upV2.filepath || upV2.file_path;

      if (diffProgressFill) diffProgressFill.style.width = '70%';
      if (diffProgressStatus) diffProgressStatus.textContent = 'Đang thực hiện giải thuật so sánh pháp lý...';

      const authorInput = document.getElementById('diff-author-input');
      const author = authorInput && authorInput.value.trim() ? authorInput.value.trim() : 'Antigravity Legal AI';

      const res = await fetch('/api/docx/diff', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ doc_a: v1Path, doc_b: v2Path, author: author }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'So sánh văn bản thất bại');
      }

      const data = await res.json();
      if (diffProgressFill) diffProgressFill.style.width = '100%';
      if (diffProgressStatus) diffProgressStatus.textContent = 'Hoàn tất!';

      if (diffResInsertions) diffResInsertions.textContent = data.insertions || data.total_insertions || 0;
      if (diffResDeletions) diffResDeletions.textContent = data.deletions || data.total_deletions || 0;
      if (diffResModParas) diffResModParas.textContent = data.modified_paragraphs || 0;

      const outName = data.output_file || data.filename || 'legal_redline_diff.docx';
      if (diffResultFilename) diffResultFilename.textContent = outName;
      if (btnDownloadDiff) btnDownloadDiff.setAttribute('href', data.download_url || `/api/download/${outName}`);

      if (diffEmptyState) diffEmptyState.style.display = 'none';
      if (diffResultBox) diffResultBox.style.display = 'block';
      showToast('So sánh pháp lý Track Changes thành công!', 'success');
    } catch (err) {
      showToast(`Lỗi so sánh: ${err.message}`, 'error', 6000);
    } finally {
      btnRunDocDiff.disabled = false;
      setTimeout(() => { if (diffProgress) diffProgress.style.display = 'none'; }, 1500);
    }
  });
}

// -------------------------------------------------------------------------
// SUB-TAB 4: VIETNAMESE OCR CONTROLLER
// -------------------------------------------------------------------------
let ocrFile = null;
const ocrDropZone = document.getElementById('ocr-drop-zone');
const ocrFileInput = document.getElementById('ocr-file-input');
const btnBrowseOcr = document.getElementById('btn-browse-ocr');
const ocrFileInfo = document.getElementById('ocr-file-info');
const ocrFileName = document.getElementById('ocr-file-name');
const ocrFileSize = document.getElementById('ocr-file-size');
const ocrFileExt = document.getElementById('ocr-file-ext');
const btnRemoveOcr = document.getElementById('btn-remove-ocr');
const btnRunOcr = document.getElementById('btn-run-ocr');
const ocrProgress = document.getElementById('ocr-progress-container');
const ocrProgressFill = document.getElementById('ocr-progress-fill');
const ocrProgressStatus = document.getElementById('ocr-progress-status');
const ocrResultBox = document.getElementById('ocr-result-box');
const ocrEmptyState = document.getElementById('ocr-empty-state');
const ocrResPages = document.getElementById('ocr-res-pages');
const ocrResultFilename = document.getElementById('ocr-result-filename');
const btnDownloadOcr = document.getElementById('btn-download-ocr');

if (ocrDropZone && ocrFileInput) {
  ocrDropZone.addEventListener('click', () => ocrFileInput.click());
  if (btnBrowseOcr) btnBrowseOcr.addEventListener('click', (e) => { e.stopPropagation(); ocrFileInput.click(); });
  ocrFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      ocrFile = e.target.files[0];
      const ext = ocrFile.name.substring(ocrFile.name.lastIndexOf('.')).replace('.', '').toUpperCase();
      if (ocrFileName) ocrFileName.textContent = ocrFile.name;
      if (ocrFileSize) ocrFileSize.textContent = formatBytes(ocrFile.size);
      if (ocrFileExt) ocrFileExt.textContent = ext;
      const c = ocrDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'none';
      if (ocrFileInfo) ocrFileInfo.style.display = 'flex';
      if (btnRunOcr) btnRunOcr.disabled = false;
    }
  });
  if (btnRemoveOcr) {
    btnRemoveOcr.addEventListener('click', (e) => {
      e.stopPropagation();
      ocrFile = null;
      ocrFileInput.value = '';
      if (ocrFileInfo) ocrFileInfo.style.display = 'none';
      const c = ocrDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'block';
      if (btnRunOcr) btnRunOcr.disabled = true;
    });
  }
}

if (btnRunOcr) {
  btnRunOcr.addEventListener('click', async () => {
    if (!ocrFile) return;
    btnRunOcr.disabled = true;
    if (ocrProgress) ocrProgress.style.display = 'block';
    if (ocrProgressFill) ocrProgressFill.style.width = '30%';
    if (ocrProgressStatus) ocrProgressStatus.textContent = 'Đang tải lên tệp scan...';

    try {
      const upRes = await uploadFileToServer(ocrFile);
      const serverPath = upRes.filepath || upRes.file_path;

      if (ocrProgressFill) ocrProgressFill.style.width = '60%';
      if (ocrProgressStatus) ocrProgressStatus.textContent = 'Đang khử mộc đỏ và nhận diện tiếng Việt...';

      const filterStamp = document.getElementById('ocr-filter-stamp')?.checked ?? true;
      const deskew = document.getElementById('ocr-deskew')?.checked ?? true;
      const fmtSelect = document.getElementById('ocr-format-select');
      const fmt = fmtSelect ? fmtSelect.value : 'md';

      const res = await fetch('/api/ocr/scan', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_path: serverPath,
          filter_red_stamp: filterStamp,
          deskew: deskew,
          format: fmt,
          lang: 'vie+eng',
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'OCR thất bại');
      }

      const data = await res.json();
      if (ocrProgressFill) ocrProgressFill.style.width = '100%';
      if (ocrProgressStatus) ocrProgressStatus.textContent = 'Hoàn tất!';

      if (ocrResPages) ocrResPages.textContent = data.total_pages || 1;
      const outName = data.filename || data.output_file || 'scanned_ocr.md';
      if (ocrResultFilename) ocrResultFilename.textContent = outName;
      if (btnDownloadOcr) btnDownloadOcr.setAttribute('href', data.download_url || `/api/download/${outName}`);

      if (ocrEmptyState) ocrEmptyState.style.display = 'none';
      if (ocrResultBox) ocrResultBox.style.display = 'block';
      showToast('Bóc tách OCR tiếng Việt thành công!', 'success');
    } catch (err) {
      showToast(`Lỗi OCR: ${err.message}`, 'error', 6000);
    } finally {
      btnRunOcr.disabled = false;
      setTimeout(() => { if (ocrProgress) ocrProgress.style.display = 'none'; }, 1500);
    }
  });
}

// -------------------------------------------------------------------------
// SUB-TAB 5: UNIFIED QA AUDIT CONTROLLER
// -------------------------------------------------------------------------
let qaFile = null;
const qaDropZone = document.getElementById('qa-drop-zone');
const qaFileInput = document.getElementById('qa-file-input');
const btnBrowseQa = document.getElementById('btn-browse-qa');
const qaFileInfo = document.getElementById('qa-file-info');
const qaFileName = document.getElementById('qa-file-name');
const qaFileSize = document.getElementById('qa-file-size');
const qaFileExt = document.getElementById('qa-file-ext');
const btnRemoveQa = document.getElementById('btn-remove-qa');
const btnRunQaAudit = document.getElementById('btn-run-qa-audit');
const qaProgress = document.getElementById('qa-progress-container');
const qaProgressFill = document.getElementById('qa-progress-fill');
const qaProgressStatus = document.getElementById('qa-progress-status');
const qaResultBox = document.getElementById('qa-result-box');
const qaEmptyState = document.getElementById('qa-empty-state');
const qaBadgeStatus = document.getElementById('qa-badge-status');
const qaResCritical = document.getElementById('qa-res-critical');
const qaResWarning = document.getElementById('qa-res-warning');
const qaResRepaired = document.getElementById('qa-res-repaired');
const qaDownloadContainer = document.getElementById('qa-download-container');
const btnDownloadQaRepaired = document.getElementById('btn-download-qa-repaired');

if (qaDropZone && qaFileInput) {
  qaDropZone.addEventListener('click', () => qaFileInput.click());
  if (btnBrowseQa) btnBrowseQa.addEventListener('click', (e) => { e.stopPropagation(); qaFileInput.click(); });
  qaFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      qaFile = e.target.files[0];
      const ext = qaFile.name.substring(qaFile.name.lastIndexOf('.')).replace('.', '').toUpperCase();
      if (qaFileName) qaFileName.textContent = qaFile.name;
      if (qaFileSize) qaFileSize.textContent = formatBytes(qaFile.size);
      if (qaFileExt) qaFileExt.textContent = ext;
      const c = qaDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'none';
      if (qaFileInfo) qaFileInfo.style.display = 'flex';
      if (btnRunQaAudit) btnRunQaAudit.disabled = false;
    }
  });
  if (btnRemoveQa) {
    btnRemoveQa.addEventListener('click', (e) => {
      e.stopPropagation();
      qaFile = null;
      qaFileInput.value = '';
      if (qaFileInfo) qaFileInfo.style.display = 'none';
      const c = qaDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'block';
      if (btnRunQaAudit) btnRunQaAudit.disabled = true;
    });
  }
}

if (btnRunQaAudit) {
  btnRunQaAudit.addEventListener('click', async () => {
    if (!qaFile) return;
    btnRunQaAudit.disabled = true;
    if (qaProgress) qaProgress.style.display = 'block';
    if (qaProgressFill) qaProgressFill.style.width = '30%';
    if (qaProgressStatus) qaProgressStatus.textContent = 'Đang tải lên tệp kiểm định...';

    try {
      const upRes = await uploadFileToServer(qaFile);
      const serverPath = upRes.filepath || upRes.file_path;

      if (qaProgressFill) qaProgressFill.style.width = '70%';
      if (qaProgressStatus) qaProgressStatus.textContent = 'Đang quét 35+ quy chuẩn OpenXML & Geometry...';

      const repairToggle = document.getElementById('qa-repair-toggle');
      const doRepair = repairToggle ? repairToggle.checked : true;
      const securityToggle = document.getElementById('qa-security-toggle');
      const doSecurity = securityToggle ? securityToggle.checked : true;
      const deepCdrToggle = document.getElementById('qa-deep-cdr-toggle');
      const doDeepCdr = deepCdrToggle ? deepCdrToggle.checked : true;
      const redactPiiToggle = document.getElementById('qa-redact-pii-toggle');
      const doRedactPii = redactPiiToggle ? redactPiiToggle.checked : false;

      const res = await fetch('/api/qa/audit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_path: serverPath,
          repair: doRepair,
          security: doSecurity,
          deep_cdr: doDeepCdr,
          redact_pii: doRedactPii,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'Kiểm định QA thất bại');
      }

      const data = await res.json();
      if (qaProgressFill) qaProgressFill.style.width = '100%';
      if (qaProgressStatus) qaProgressStatus.textContent = 'Hoàn tất!';

      const crit = data.critical_count || 0;
      const warn = data.warning_count || 0;
      if (qaResCritical) qaResCritical.textContent = crit;
      if (qaResWarning) qaResWarning.textContent = warn;

      const qaResSecurity = document.getElementById('qa-res-security');
      const qaSecFindings = document.getElementById('qa-security-findings');
      const secIssues = (data.issues || []).filter(i => (i.domain === 'SECURITY' || (i.code && i.code.startsWith('ERR_SEC'))));
      if (qaResSecurity) {
        if (!doSecurity) {
          qaResSecurity.textContent = 'Bỏ qua (Tắt quét)';
          qaResSecurity.style.color = '#94a3b8';
        } else if (secIssues.length === 0) {
          qaResSecurity.textContent = '✓ 100% Sạch — An toàn';
          qaResSecurity.style.color = '#10b981';
        } else {
          qaResSecurity.textContent = `⚠ Phát hiện ${secIssues.length} nguy cơ (${secIssues.filter(i => (i.severity || '').toUpperCase() === 'CRITICAL').length} Critical)`;
          qaResSecurity.style.color = '#ef4444';
        }
      }

      if (qaSecFindings) {
        if (secIssues.length > 0) {
          qaSecFindings.style.display = 'block';
          qaSecFindings.innerHTML = `<strong>Chi tiết nguy cơ an ninh (Gate 13):</strong><ul style="margin: 6px 0 0 16px; padding: 0;">` +
            secIssues.map(i => `<li>[${escapeHtml(i.code || 'ERR_SEC')}] ${escapeHtml(i.message || '')}</li>`).join('') +
            `</ul>`;
        } else {
          qaSecFindings.style.display = 'none';
        }
      }

      if (qaBadgeStatus) {
        if (crit === 0) {
          qaBadgeStatus.className = 'result-badge-success';
          qaBadgeStatus.innerHTML = '<i data-lucide="check-circle-2"></i> Đạt chuẩn chất lượng 100%!';
        } else {
          qaBadgeStatus.className = 'result-badge-danger';
          qaBadgeStatus.innerHTML = `<i data-lucide="alert-triangle"></i> Phát hiện ${crit} lỗi nghiêm trọng!`;
        }
      }

      if (qaResRepaired) {
        if (data.repaired && data.repair_actions && data.repair_actions.length > 0) {
          qaResRepaired.textContent = `${data.repair_actions.length} hành động sửa an toàn`;
        } else {
          qaResRepaired.textContent = 'Không cần sửa hoặc chế độ chỉ kiểm định';
        }
      }

      if (qaDownloadContainer) {
        if (data.download_url) {
          qaDownloadContainer.style.display = 'block';
          if (btnDownloadQaRepaired) btnDownloadQaRepaired.setAttribute('href', data.download_url);
        } else {
          qaDownloadContainer.style.display = 'none';
        }
      }

      if (qaEmptyState) qaEmptyState.style.display = 'none';
      if (qaResultBox) qaResultBox.style.display = 'block';
      refreshLucideIcons();
      showToast('Kiểm định QA hoàn tất!', 'success');
    } catch (err) {
      showToast(`Lỗi kiểm định: ${err.message}`, 'error', 6000);
    } finally {
      btnRunQaAudit.disabled = false;
      setTimeout(() => { if (qaProgress) qaProgress.style.display = 'none'; }, 1500);
    }
  });
}

// -------------------------------------------------------------------------
// SUB-TAB 6: PDF PRO STUDIO & VISUAL PAGE ORGANIZER CONTROLLER
// -------------------------------------------------------------------------
let pdfProFile = null;
let pdfProServerPath = null;
let pdfPages = []; // [{ pageNumber: 1, rotation: 0, deleted: false, thumbnailUrl: '' }]
let giapLaiStampFile = null;
let giapLaiStampServerPath = null;

const pdfProDropZone = document.getElementById('pdf-pro-drop-zone');
const pdfProFileInput = document.getElementById('pdf-pro-file-input');
const btnBrowsePdfPro = document.getElementById('btn-browse-pdf-pro');
const pdfProFileInfo = document.getElementById('pdf-pro-file-info');
const pdfProFileName = document.getElementById('pdf-pro-file-name');
const pdfProFileSize = document.getElementById('pdf-pro-file-size');
const btnRemovePdfPro = document.getElementById('btn-remove-pdf-pro');

const pdfActionPills = document.querySelectorAll('.pdf-action-pill');
const pdfActionPanes = document.querySelectorAll('.pdf-action-card');

const pdfThumbnailGrid = document.getElementById('pdf-thumbnail-grid');
const pdfProEmptyState = document.getElementById('pdf-pro-empty-state');
const pdfOrganizerToolbar = document.getElementById('pdf-organizer-toolbar');
const btnRotateAllPages = document.getElementById('btn-rotate-all-pages');
const btnResetAllPages = document.getElementById('btn-reset-all-pages');
const pdfPageCountBadge = document.getElementById('pdf-page-count-badge');

const pdfStatKept = document.getElementById('pdf-stat-kept');
const pdfStatRotated = document.getElementById('pdf-stat-rotated');
const pdfStatDeleted = document.getElementById('pdf-stat-deleted');

const btnApplyReorganize = document.getElementById('btn-apply-reorganize');
const btnApplyRedact = document.getElementById('btn-apply-redact');
const btnApplyProtect = document.getElementById('btn-apply-protect');
const btnApplySign = document.getElementById('btn-apply-sign');
const btnApplyGiaplai = document.getElementById('btn-apply-giaplai');
const btnApplyNormalize = document.getElementById('btn-apply-normalize');
const btnApplyExtractTables = document.getElementById('btn-apply-extract-tables');

const pdfProProgress = document.getElementById('pdf-pro-progress-container');
const pdfProProgressFill = document.getElementById('pdf-pro-progress-fill');
const pdfProProgressStatus = document.getElementById('pdf-pro-progress-status');
const pdfProResultBox = document.getElementById('pdf-pro-result-box');
const pdfProBadgeStatus = document.getElementById('pdf-pro-badge-status');
const pdfProResAction = document.getElementById('pdf-pro-res-action');
const pdfProResultFilename = document.getElementById('pdf-pro-result-filename');
const pdfProExtraDetailRow = document.getElementById('pdf-pro-extra-detail-row');
const pdfProExtraDetail = document.getElementById('pdf-pro-extra-detail');
const btnDownloadPdfPro = document.getElementById('btn-download-pdf-pro');

// Electronic Signature & Stamp elements
let signMode = 'draw';
let signInkColor = '#0033aa';
let hasDrawnSignature = false;
let signStampServerPath = null;
let signSelectedPos = 'bottom-right';

const signModeTabs = document.querySelectorAll('.sign-mode-tab');
const signModePanelDraw = document.getElementById('sign-mode-panel-draw');
const signModePanelUpload = document.getElementById('sign-mode-panel-upload');
const signCanvas = document.getElementById('signature-pad-canvas');
const signCanvasHint = document.getElementById('signature-canvas-hint');
const btnClearSignature = document.getElementById('btn-clear-signature');
const signColorDots = document.querySelectorAll('.sign-color-dot');
const signStampDropZone = document.getElementById('sign-stamp-drop-zone');
const signStampFileInput = document.getElementById('sign-stamp-file-input');
const signStampFileInfo = document.getElementById('sign-stamp-file-info');
const signStampFileName = document.getElementById('sign-stamp-file-name');
const btnRemoveSignStamp = document.getElementById('btn-remove-sign-stamp');
const signTargetPageSelect = document.getElementById('sign-target-page-select');
const signCustomPageGroup = document.getElementById('sign-custom-page-group');
const signCustomPageInput = document.getElementById('sign-custom-page-input');
const signPosBtns = document.querySelectorAll('.sign-pos-btn');
const signScaleSlider = document.getElementById('sign-scale-slider');
const signScaleDisplay = document.getElementById('sign-scale-display');
const signOptMetadata = document.getElementById('sign-opt-metadata');
const signMetadataFields = document.getElementById('sign-metadata-fields');
const signMetaTitle = document.getElementById('sign-meta-title');
const signMetaName = document.getElementById('sign-meta-name');
const signMetaDate = document.getElementById('sign-meta-date');

// Giap lai stamp upload elements
const giaplaiStampDropZone = document.getElementById('giaplai-stamp-drop-zone');
const giaplaiStampFileInput = document.getElementById('giaplai-stamp-file-input');
const giaplaiStampFileInfo = document.getElementById('giaplai-stamp-file-info');
const giaplaiStampFileName = document.getElementById('giaplai-stamp-file-name');
const btnRemoveGiaplaiStamp = document.getElementById('btn-remove-giaplai-stamp');

// Action Pills Navigation
pdfActionPills.forEach((pill) => {
  pill.addEventListener('click', () => {
    const action = pill.getAttribute('data-action');
    pdfActionPills.forEach((p) => p.classList.remove('active'));
    pill.classList.add('active');

    pdfActionPanes.forEach((pane) => {
      if (pane.id === `pdf-action-pane-${action}`) {
        pane.style.display = 'block';
        pane.classList.add('active');
      } else {
        pane.style.display = 'none';
        pane.classList.remove('active');
      }
    });
    refreshLucideIcons();
  });
});

function updatePdfOrganizeStats() {
  const kept = pdfPages.filter((p) => !p.deleted).length;
  const rotated = pdfPages.filter((p) => !p.deleted && p.rotation !== 0).length;
  const deleted = pdfPages.filter((p) => p.deleted).length;

  if (pdfStatKept) pdfStatKept.textContent = kept;
  if (pdfStatRotated) pdfStatRotated.textContent = rotated;
  if (pdfStatDeleted) pdfStatDeleted.textContent = deleted;
  if (pdfPageCountBadge) pdfPageCountBadge.textContent = `${pdfPages.length} trang`;

  if (btnApplyReorganize) btnApplyReorganize.disabled = kept === 0;
}

function updateSignApplyButtonState() {
  if (!btnApplySign) return;
  const hasSourcePdf = Boolean(pdfProServerPath);
  const hasSignatureContent = signMode === 'draw' ? hasDrawnSignature : Boolean(signStampServerPath);
  btnApplySign.disabled = !(hasSourcePdf && hasSignatureContent);
}

function setPdfToolButtonsEnabled(enabled) {
  if (btnApplyReorganize) btnApplyReorganize.disabled = !enabled;
  if (btnApplyRedact) btnApplyRedact.disabled = !enabled;
  if (btnApplyProtect) btnApplyProtect.disabled = !enabled;
  if (btnApplyNormalize) btnApplyNormalize.disabled = !enabled;
  if (btnApplyExtractTables) btnApplyExtractTables.disabled = !enabled;
  if (btnApplyGiaplai) btnApplyGiaplai.disabled = !(enabled && giapLaiStampServerPath);
  updateSignApplyButtonState();
}

function renderPdfThumbnails() {
  if (!pdfThumbnailGrid) return;

  if (pdfPages.length === 0) {
    if (pdfProEmptyState) pdfProEmptyState.style.display = 'flex';
    if (pdfOrganizerToolbar) pdfOrganizerToolbar.style.display = 'none';
    return;
  }

  if (pdfProEmptyState) pdfProEmptyState.style.display = 'none';
  if (pdfOrganizerToolbar) pdfOrganizerToolbar.style.display = 'flex';

  // Remove existing cards
  const existingCards = pdfThumbnailGrid.querySelectorAll('.pdf-page-card');
  existingCards.forEach((c) => c.remove());

  pdfPages.forEach((pageObj) => {
    const card = document.createElement('div');
    card.className = `pdf-page-card ${pageObj.deleted ? 'is-deleted' : ''}`;
    card.setAttribute('data-page', pageObj.pageNumber);

    card.innerHTML = `
      <div class="page-card-header">
        <span class="page-idx-badge">Trang ${pageObj.pageNumber}</span>
        ${pageObj.rotation !== 0 ? `<span class="page-rot-badge">+${pageObj.rotation}°</span>` : ''}
      </div>
      <div class="page-thumbnail-wrap">
        <img src="${pageObj.thumbnailUrl}" class="page-thumb-img" style="transform: rotate(${pageObj.rotation}deg);" alt="Trang ${pageObj.pageNumber}" />
      </div>
      <div class="page-card-actions">
        <button type="button" class="page-action-btn rotate-btn" title="Xoay 90°"><i data-lucide="rotate-cw"></i></button>
        <button type="button" class="page-action-btn delete-btn" title="${pageObj.deleted ? 'Khôi phục trang' : 'Xóa trang'}"><i data-lucide="${pageObj.deleted ? 'undo-2' : 'trash-2'}"></i></button>
      </div>
    `;

    // Event listeners for individual card
    const rotateBtn = card.querySelector('.rotate-btn');
    const deleteBtn = card.querySelector('.delete-btn');
    const thumbImg = card.querySelector('.page-thumb-img');
    const cardHeader = card.querySelector('.page-card-header');

    rotateBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      pageObj.rotation = (pageObj.rotation + 90) % 360;
      thumbImg.style.transform = `rotate(${pageObj.rotation}deg)`;

      let rotBadge = cardHeader.querySelector('.page-rot-badge');
      if (pageObj.rotation !== 0) {
        if (!rotBadge) {
          rotBadge = document.createElement('span');
          rotBadge.className = 'page-rot-badge';
          cardHeader.appendChild(rotBadge);
        }
        rotBadge.textContent = `+${pageObj.rotation}°`;
      } else if (rotBadge) {
        rotBadge.remove();
      }
      updatePdfOrganizeStats();
    });

    deleteBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      pageObj.deleted = !pageObj.deleted;
      card.classList.toggle('is-deleted', pageObj.deleted);
      deleteBtn.title = pageObj.deleted ? 'Khôi phục trang' : 'Xóa trang';
      deleteBtn.innerHTML = `<i data-lucide="${pageObj.deleted ? 'undo-2' : 'trash-2'}"></i>`;
      refreshLucideIcons();
      updatePdfOrganizeStats();
    });

    pdfThumbnailGrid.appendChild(card);
  });

  updatePdfOrganizeStats();
  refreshLucideIcons();
}

// Batch actions
if (btnRotateAllPages) {
  btnRotateAllPages.addEventListener('click', () => {
    pdfPages.forEach((p) => {
      if (!p.deleted) p.rotation = (p.rotation + 90) % 360;
    });
    renderPdfThumbnails();
  });
}

if (btnResetAllPages) {
  btnResetAllPages.addEventListener('click', () => {
    pdfPages.forEach((p) => {
      p.rotation = 0;
      p.deleted = false;
    });
    renderPdfThumbnails();
  });
}

// Load and render PDF file
async function loadPdfProFile(file) {
  pdfProFile = file;
  if (pdfProFileName) pdfProFileName.textContent = file.name;
  if (pdfProFileSize) pdfProFileSize.textContent = formatBytes(file.size);
  const content = pdfProDropZone.querySelector('.drop-zone-content');
  if (content) content.style.display = 'none';
  if (pdfProFileInfo) pdfProFileInfo.style.display = 'flex';

  if (pdfProProgress) pdfProProgress.style.display = 'block';
  if (pdfProProgressFill) pdfProProgressFill.style.width = '35%';
  if (pdfProProgressStatus) pdfProProgressStatus.textContent = 'Đang tải lên và dựng ảnh xem trước từng trang...';

  try {
    const upRes = await uploadFileToServer(file);
    pdfProServerPath = upRes.filepath || upRes.file_path;

    if (pdfProProgressFill) pdfProProgressFill.style.width = '65%';
    if (pdfProProgressStatus) pdfProProgressStatus.textContent = 'Đang kết xuất ảnh nét 120 DPI (PyMuPDF)...';

    const prevRes = await fetch('/api/pdf/pages/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: pdfProServerPath }),
    });

    if (!prevRes.ok) {
      const err = await prevRes.json().catch(() => ({}));
      throw new Error(err.error || 'Không thể tạo ảnh xem trước');
    }

    const prevData = await prevRes.json();
    pdfPages = (prevData.pages || []).map((p) => ({
      pageNumber: p.page || p.page_num,
      rotation: 0,
      deleted: false,
      thumbnailUrl: p.thumbnail_url || p.preview_url,
    }));

    if (pdfProProgressFill) pdfProProgressFill.style.width = '100%';
    if (pdfProProgressStatus) pdfProProgressStatus.textContent = 'Sẵn sàng!';

    renderPdfThumbnails();
    setPdfToolButtonsEnabled(true);
    showToast(`Đã nạp ${pdfPages.length} trang PDF thành công!`, 'success');
  } catch (err) {
    showToast(`Lỗi xử lý PDF: ${err.message}`, 'error', 6000);
    setPdfToolButtonsEnabled(false);
  } finally {
    setTimeout(() => {
      if (pdfProProgress) pdfProProgress.style.display = 'none';
    }, 1200);
  }
}

if (pdfProDropZone && pdfProFileInput) {
  pdfProDropZone.addEventListener('click', () => pdfProFileInput.click());
  if (btnBrowsePdfPro) {
    btnBrowsePdfPro.addEventListener('click', (e) => {
      e.stopPropagation();
      pdfProFileInput.click();
    });
  }

  pdfProFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      loadPdfProFile(e.target.files[0]);
    }
  });

  // Drag and drop support
  pdfProDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    pdfProDropZone.classList.add('drag-over');
  });
  pdfProDropZone.addEventListener('dragleave', () => {
    pdfProDropZone.classList.remove('drag-over');
  });
  pdfProDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    pdfProDropZone.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.toLowerCase().endsWith('.pdf')) {
        loadPdfProFile(file);
      } else {
        showToast('Vui lòng chọn tệp có định dạng .pdf', 'warning');
      }
    }
  });

  if (btnRemovePdfPro) {
    btnRemovePdfPro.addEventListener('click', (e) => {
      e.stopPropagation();
      pdfProFile = null;
      pdfProServerPath = null;
      pdfPages = [];
      pdfProFileInput.value = '';
      if (pdfProFileInfo) pdfProFileInfo.style.display = 'none';
      const content = pdfProDropZone.querySelector('.drop-zone-content');
      if (content) content.style.display = 'block';
      renderPdfThumbnails();
      setPdfToolButtonsEnabled(false);
      if (pdfProResultBox) pdfProResultBox.style.display = 'none';
    });
  }
}

// Giáp Lai Stamp Dropzone Handling
if (giaplaiStampDropZone && giaplaiStampFileInput) {
  giaplaiStampDropZone.addEventListener('click', () => giaplaiStampFileInput.click());
  giaplaiStampFileInput.addEventListener('change', async (e) => {
    if (e.target.files && e.target.files.length > 0) {
      giapLaiStampFile = e.target.files[0];
      if (giaplaiStampFileName) giaplaiStampFileName.textContent = giapLaiStampFile.name;
      const c = giaplaiStampDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'none';
      if (giaplaiStampFileInfo) giaplaiStampFileInfo.style.display = 'flex';

      try {
        const up = await uploadFileToServer(giapLaiStampFile);
        giapLaiStampServerPath = up.filepath || up.file_path;
        if (btnApplyGiaplai) btnApplyGiaplai.disabled = !pdfProServerPath;
        showToast('Tải lên ảnh con dấu thành công!', 'success');
      } catch (err) {
        showToast(`Lỗi tải con dấu: ${err.message}`, 'error');
      }
    }
  });

  if (btnRemoveGiaplaiStamp) {
    btnRemoveGiaplaiStamp.addEventListener('click', (e) => {
      e.stopPropagation();
      giapLaiStampFile = null;
      giapLaiStampServerPath = null;
      giaplaiStampFileInput.value = '';
      if (giaplaiStampFileInfo) giaplaiStampFileInfo.style.display = 'none';
      const c = giaplaiStampDropZone.querySelector('.drop-zone-content');
      if (c) c.style.display = 'block';
      if (btnApplyGiaplai) btnApplyGiaplai.disabled = true;
    });
  }
}

// Unified PDF Pro Action Dispatcher
async function executePdfProAction(action, payload, actionTitle, triggerBtn) {
  if (!pdfProServerPath) {
    showToast('Vui lòng tải lên tệp PDF trước!', 'warning');
    return;
  }

  if (triggerBtn) triggerBtn.disabled = true;
  if (pdfProProgress) pdfProProgress.style.display = 'block';
  if (pdfProProgressFill) pdfProProgressFill.style.width = '30%';
  if (pdfProProgressStatus) pdfProProgressStatus.textContent = `Đang thực hiện ${actionTitle}...`;

  try {
    const bodyData = {
      file_path: pdfProServerPath,
      action: action,
      ...payload,
    };

    if (pdfProProgressFill) pdfProProgressFill.style.width = '65%';

    const res = await fetch('/api/pdf/tools', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyData),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `${actionTitle} thất bại`);
    }

    const data = await res.json();
    if (pdfProProgressFill) pdfProProgressFill.style.width = '100%';
    if (pdfProProgressStatus) pdfProProgressStatus.textContent = 'Hoàn tất!';

    const outName = data.output_file || data.filename || 'result.pdf';
    const dlUrl = data.download_url || `/api/download/${outName}`;

    if (pdfProResAction) pdfProResAction.textContent = actionTitle;
    if (pdfProResultFilename) pdfProResultFilename.textContent = outName;
    if (btnDownloadPdfPro) btnDownloadPdfPro.setAttribute('href', dlUrl);

    if (data.detail && pdfProExtraDetailRow && pdfProExtraDetail) {
      pdfProExtraDetailRow.style.display = 'flex';
      pdfProExtraDetail.textContent = data.detail;
    } else if (pdfProExtraDetailRow) {
      pdfProExtraDetailRow.style.display = 'none';
    }

    if (pdfProResultBox) pdfProResultBox.style.display = 'block';
    refreshLucideIcons();
    showToast(`${actionTitle} hoàn tất!`, 'success');
  } catch (err) {
    showToast(`Lỗi: ${err.message}`, 'error', 6000);
  } finally {
    if (triggerBtn) triggerBtn.disabled = false;
    setTimeout(() => {
      if (pdfProProgress) pdfProProgress.style.display = 'none';
    }, 1500);
  }
}

// 1. Action Reorganize
if (btnApplyReorganize) {
  btnApplyReorganize.addEventListener('click', () => {
    const pagesPlan = pdfPages
      .filter((p) => !p.deleted)
      .map((p) => ({ page: p.pageNumber, rotate: p.rotation }));

    if (pagesPlan.length === 0) {
      showToast('Vui lòng giữ lại ít nhất 1 trang!', 'warning');
      return;
    }

    executePdfProAction('reorganize', { pages_plan: pagesPlan }, 'Tổ chức & sắp xếp trang', btnApplyReorganize);
  });
}

// 2. Action True Redaction
if (btnApplyRedact) {
  btnApplyRedact.addEventListener('click', () => {
    const patterns = [];
    if (document.getElementById('redact-opt-cccd')?.checked) patterns.push('cccd');
    if (document.getElementById('redact-opt-phone')?.checked) patterns.push('phone');
    if (document.getElementById('redact-opt-email')?.checked) patterns.push('email');
    if (document.getElementById('redact-opt-bank')?.checked) patterns.push('bank_card');
    if (document.getElementById('redact-opt-salary')?.checked) patterns.push('salary');

    const termsInput = document.getElementById('redact-custom-terms');
    const customTerms = termsInput && termsInput.value.trim()
      ? termsInput.value.split(',').map((s) => s.trim()).filter(Boolean)
      : [];

    const fillColor = document.getElementById('redact-fill-color')?.value || 'black';

    if (patterns.length === 0 && customTerms.length === 0) {
      showToast('Vui lòng chọn ít nhất 1 danh mục mẫu hoặc nhập từ khóa cần bôi đen!', 'warning');
      return;
    }

    executePdfProAction(
      'redact',
      {
        patterns: patterns,
        custom_terms: customTerms,
        fill_color: fillColor,
      },
      'Bôi đen bảo mật True Redaction',
      btnApplyRedact
    );
  });
}

// 3. Action Protect AES-256
if (btnApplyProtect) {
  btnApplyProtect.addEventListener('click', () => {
    const userPwd = document.getElementById('protect-user-pwd')?.value || '';
    const ownerPwd = document.getElementById('protect-owner-pwd')?.value || '';
    const allowPrint = document.getElementById('protect-allow-print')?.checked ?? true;
    const allowCopy = document.getElementById('protect-allow-copy')?.checked ?? false;
    const allowModify = document.getElementById('protect-allow-modify')?.checked ?? false;

    if (!userPwd && !ownerPwd) {
      showToast('Vui lòng nhập ít nhất Mật khẩu mở file hoặc Mật khẩu chủ quyền!', 'warning');
      return;
    }

    executePdfProAction(
      'protect',
      {
        user_pwd: userPwd,
        owner_pwd: ownerPwd,
        allow_print: allowPrint,
        allow_copy: allowCopy,
        allow_modify: allowModify,
      },
      'Mã hóa AES-256 PDF',
      btnApplyProtect
    );
  });
}

// 4. Action Giáp Lai Stamp
if (btnApplyGiaplai) {
  btnApplyGiaplai.addEventListener('click', () => {
    if (!giapLaiStampServerPath) {
      showToast('Vui lòng tải lên ảnh con dấu mộc đỏ!', 'warning');
      return;
    }

    const side = document.getElementById('giaplai-side')?.value || 'right';

    executePdfProAction(
      'giap_lai',
      {
        stamp_file_path: giapLaiStampServerPath,
        side: side,
      },
      'Đóng dấu giáp lai đa trang',
      btnApplyGiaplai
    );
  });
}

// 5. Action Normalize A4 & Bates
if (btnApplyNormalize) {
  btnApplyNormalize.addEventListener('click', () => {
    const orientation = document.getElementById('norm-orientation')?.value || 'auto';
    const batesPrefix = document.getElementById('norm-bates-prefix')?.value || '';

    executePdfProAction(
      'normalize_a4',
      {
        orientation: orientation,
        bates_prefix: batesPrefix,
      },
      'Chuẩn hóa A4 & Đánh số Bates',
      btnApplyNormalize
    );
  });
}

// 6. Action Extract Tables to Excel
if (btnApplyExtractTables) {
  btnApplyExtractTables.addEventListener('click', () => {
    executePdfProAction(
      'extract_tables',
      {},
      'Bóc tách bảng biểu sang Excel',
      btnApplyExtractTables
    );
  });
}

// 7. Electronic Signature & Stamp Logic
let isDrawingSign = false;
let lastSignX = 0;
let lastSignY = 0;

function initSignatureCanvas() {
  if (!signCanvas) return;
  const ctx = signCanvas.getContext('2d');
  if (!ctx) return;

  function getCanvasCoordinates(e) {
    const rect = signCanvas.getBoundingClientRect();
    const scaleX = signCanvas.width / rect.width;
    const scaleY = signCanvas.height / rect.height;

    if (e.touches && e.touches.length > 0) {
      return {
        x: (e.touches[0].clientX - rect.left) * scaleX,
        y: (e.touches[0].clientY - rect.top) * scaleY,
      };
    }
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY,
    };
  }

  function startDrawing(e) {
    isDrawingSign = true;
    const pos = getCanvasCoordinates(e);
    lastSignX = pos.x;
    lastSignY = pos.y;
    if (signCanvasHint) signCanvasHint.style.opacity = '0';
  }

  function draw(e) {
    if (!isDrawingSign) return;
    const pos = getCanvasCoordinates(e);

    ctx.beginPath();
    ctx.moveTo(lastSignX, lastSignY);
    ctx.lineTo(pos.x, pos.y);
    ctx.strokeStyle = signInkColor;
    ctx.lineWidth = 2.8;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.stroke();

    lastSignX = pos.x;
    lastSignY = pos.y;
    hasDrawnSignature = true;
    updateSignApplyButtonState();
  }

  function stopDrawing() {
    isDrawingSign = false;
  }

  signCanvas.addEventListener('mousedown', startDrawing);
  signCanvas.addEventListener('mousemove', draw);
  signCanvas.addEventListener('mouseup', stopDrawing);
  signCanvas.addEventListener('mouseleave', stopDrawing);

  signCanvas.addEventListener('touchstart', (e) => {
    e.preventDefault();
    startDrawing(e);
  }, { passive: false });

  signCanvas.addEventListener('touchmove', (e) => {
    e.preventDefault();
    draw(e);
  }, { passive: false });

  signCanvas.addEventListener('touchend', stopDrawing);
  signCanvas.addEventListener('touchcancel', stopDrawing);
}

function clearSignatureCanvas() {
  if (!signCanvas) return;
  const ctx = signCanvas.getContext('2d');
  if (ctx) ctx.clearRect(0, 0, signCanvas.width, signCanvas.height);
  hasDrawnSignature = false;
  if (signCanvasHint) signCanvasHint.style.opacity = '1';
  updateSignApplyButtonState();
}

if (btnClearSignature) {
  btnClearSignature.addEventListener('click', clearSignatureCanvas);
}

// Color picker
signColorDots.forEach((dot) => {
  dot.addEventListener('click', () => {
    signColorDots.forEach((d) => d.classList.remove('active'));
    dot.classList.add('active');
    signInkColor = dot.getAttribute('data-color') || '#0033aa';
  });
});

// Mode Switcher (Draw vs Upload)
signModeTabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    signModeTabs.forEach((t) => t.classList.remove('active'));
    tab.classList.add('active');
    signMode = tab.getAttribute('data-sign-mode') || 'draw';

    if (signMode === 'draw') {
      if (signModePanelDraw) signModePanelDraw.style.display = 'block';
      if (signModePanelUpload) signModePanelUpload.style.display = 'none';
    } else {
      if (signModePanelDraw) signModePanelDraw.style.display = 'none';
      if (signModePanelUpload) signModePanelUpload.style.display = 'block';
    }
    updateSignApplyButtonState();
    refreshLucideIcons();
  });
});

// Preset Position buttons
signPosBtns.forEach((btn) => {
  btn.addEventListener('click', () => {
    signPosBtns.forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    signSelectedPos = btn.getAttribute('data-pos') || 'bottom-right';
  });
});

// Scale slider
if (signScaleSlider && signScaleDisplay) {
  signScaleSlider.addEventListener('input', () => {
    signScaleDisplay.textContent = `${signScaleSlider.value}%`;
  });
}

// Target page select
if (signTargetPageSelect) {
  signTargetPageSelect.addEventListener('change', () => {
    if (signCustomPageGroup) {
      signCustomPageGroup.style.display = signTargetPageSelect.value === 'custom' ? 'block' : 'none';
    }
  });
}

// Metadata Toggle
if (signOptMetadata && signMetadataFields) {
  signOptMetadata.addEventListener('change', () => {
    signMetadataFields.style.display = signOptMetadata.checked ? 'flex' : 'none';
    if (signOptMetadata.checked && signMetaDate && !signMetaDate.value) {
      const today = new Date();
      const d = String(today.getDate()).padStart(2, '0');
      const m = String(today.getMonth() + 1).padStart(2, '0');
      const y = today.getFullYear();
      signMetaDate.value = `Ngày ${d} tháng ${m} năm ${y}`;
    }
  });
}

// Stamp / Signature Image Upload
if (signStampDropZone && signStampFileInput) {
  signStampDropZone.addEventListener('click', () => signStampFileInput.click());

  signStampFileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const upRes = await uploadFileToServer(file);
      signStampServerPath = upRes.filepath || upRes.file_path;

      if (signStampFileName) signStampFileName.textContent = file.name;
      const dropContent = signStampDropZone.querySelector('.drop-zone-content');
      if (dropContent) dropContent.style.display = 'none';
      if (signStampFileInfo) signStampFileInfo.style.display = 'flex';

      updateSignApplyButtonState();
      showToast('Đã tải ảnh con dấu/chữ ký thành công!', 'success');
    } catch (err) {
      showToast(`Tải ảnh thất bại: ${err.message}`, 'error');
    }
  });
}

if (btnRemoveSignStamp) {
  btnRemoveSignStamp.addEventListener('click', (e) => {
    e.stopPropagation();
    signStampServerPath = null;
    if (signStampFileInput) signStampFileInput.value = '';
    if (signStampFileInfo) signStampFileInfo.style.display = 'none';
    const dropContent = signStampDropZone?.querySelector('.drop-zone-content');
    if (dropContent) dropContent.style.display = 'flex';
    updateSignApplyButtonState();
  });
}

// Action Button Apply Sign
if (btnApplySign) {
  btnApplySign.addEventListener('click', () => {
    const pageSelection = signTargetPageSelect?.value || 'last';
    let targetPage = pageSelection;
    if (pageSelection === 'custom') {
      const customVal = signCustomPageInput?.value?.trim() || '1';
      targetPage = customVal;
    }

    const scalePct = parseInt(signScaleSlider?.value || '25', 10) / 100.0;
    const wPct = scalePct;
    const hPct = scalePct * 0.6;

    let xPct = 0.68;
    let yPct = 0.76;

    const hasMeta = signOptMetadata?.checked;
    const metaOffset = hasMeta ? 0.08 : 0.0;

    if (signSelectedPos === 'bottom-right') {
      xPct = Math.max(0.05, 0.95 - wPct);
      yPct = Math.max(0.05, 0.95 - hPct - metaOffset);
    } else if (signSelectedPos === 'bottom-left') {
      xPct = 0.08;
      yPct = Math.max(0.05, 0.95 - hPct - metaOffset);
    } else if (signSelectedPos === 'center') {
      xPct = Math.max(0.05, (1.0 - wPct) / 2.0);
      yPct = 0.45;
    } else if (signSelectedPos === 'top-right') {
      xPct = Math.max(0.05, 0.95 - wPct);
      yPct = 0.08;
    }

    const sigObj = {
      page: targetPage,
      x_pct: xPct,
      y_pct: yPct,
      w_pct: wPct,
      h_pct: hPct,
    };

    if (signMode === 'draw') {
      if (!signCanvas || !hasDrawnSignature) {
        showToast('Vui lòng ký tên vào khung vẽ trước!', 'warning');
        return;
      }
      sigObj.image_base64 = signCanvas.toDataURL('image/png');
    } else {
      if (!signStampServerPath) {
        showToast('Vui lòng tải lên ảnh chữ ký / con dấu!', 'warning');
        return;
      }
      sigObj.image_path = signStampServerPath;
    }

    if (hasMeta) {
      sigObj.signer_title = signMetaTitle?.value?.trim() || '';
      sigObj.signer_name = signMetaName?.value?.trim() || '';
      sigObj.signer_date = signMetaDate?.value?.trim() || '';
    }

    executePdfProAction(
      'sign',
      {
        signatures: [sigObj],
      },
      'Ký tên & Đóng dấu điện tử',
      btnApplySign
    );
  });
}

initSignatureCanvas();


// -------------------------------------------------------------------------
// SUB-TAB 7: DOCX QUICK PATCH & INSPECTOR (Two-Way Sync Controller)
// -------------------------------------------------------------------------
function initDocxQuickPatch() {
  const patchDropZone = document.getElementById('doc-patch-drop-zone');
  const patchFileInput = document.getElementById('doc-patch-file-input');
  const btnBrowsePatch = document.getElementById('btn-browse-doc-patch');
  const patchFileInfo = document.getElementById('doc-patch-file-info');
  const patchFileName = document.getElementById('doc-patch-file-name');
  const patchFileSize = document.getElementById('doc-patch-file-size');
  const btnRemovePatch = document.getElementById('btn-remove-doc-patch');

  const btnToggleMode = document.getElementById('btn-toggle-patch-mode');
  const toggleLabel = document.getElementById('toggle-patch-mode-label');
  const formView = document.getElementById('doc-patch-form-view');
  const formContainer = document.getElementById('doc-patch-form-container');
  const jsonView = document.getElementById('doc-patch-json-view');
  const jsonInput = document.getElementById('doc-patch-json-input');

  const btnRunPatch = document.getElementById('btn-run-docx-patch');
  const progressContainer = document.getElementById('doc-patch-progress-container');
  const progressFill = document.getElementById('doc-patch-progress-fill');
  const progressStatus = document.getElementById('doc-patch-progress-status');

  const statsBox = document.getElementById('doc-patch-stats-box');
  const tokenCount = document.getElementById('doc-patch-token-count');
  const tokensPillList = document.getElementById('doc-patch-tokens-pill-list');

  const emptyState = document.getElementById('doc-patch-empty-state');
  const resultBox = document.getElementById('doc-patch-result-box');
  const resultFilename = document.getElementById('doc-patch-result-filename');
  const resultTime = document.getElementById('doc-patch-result-time');
  const btnDownload = document.getElementById('btn-download-docx-patched');

  let currentFile = null;
  let serverPath = null;
  let patchData = {};
  let isJsonMode = false;

  function renderFormFields() {
    if (!formContainer) return;
    const keys = Object.keys(patchData);
    if (keys.length === 0) {
      formContainer.innerHTML = `
        <div style="padding: 14px; text-align: center; background: rgba(255,255,255,0.02); border-radius: 8px; border: 1px dashed var(--border-color, rgba(255,255,255,0.1));">
          <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 10px;">
            Chưa phát hiện biến tự động. Bạn có thể thêm trường dữ liệu tùy chỉnh hoặc nhập JSON.
          </p>
          <button type="button" class="btn-sm-ghost" id="btn-add-custom-patch-key">
            <i data-lucide="plus"></i> <span>+ Thêm Biến Mới</span>
          </button>
        </div>
      `;
      const btnAdd = formContainer.querySelector('#btn-add-custom-patch-key');
      if (btnAdd) {
        btnAdd.addEventListener('click', () => {
          const newKey = prompt('Nhập tên biến (VD: HO_TEN hoặc {{HO_TEN}}):');
          if (newKey) {
            const cleanKey = newKey.trim();
            patchData[cleanKey] = '';
            syncToUi();
          }
        });
      }
      refreshLucideIcons();
      return;
    }

    formContainer.innerHTML = keys.map((key) => {
      const displayLabel = key.startsWith('{{') ? key : `{{${key}}}`;
      const val = patchData[key] || '';
      return `
        <div class="form-group" style="margin-bottom: 8px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
            <label class="form-label" style="margin: 0; font-family: var(--font-mono, monospace); font-weight: 600; color: var(--accent-color, #38bdf8); font-size: 0.85rem;">
              ${displayLabel}
            </label>
            <button type="button" class="btn-remove-patch-row" data-key="${key}" title="Xóa biến này" style="background: none; border: none; color: #ef4444; cursor: pointer; padding: 2px;">
              <i data-lucide="x" style="width: 14px; height: 14px;"></i>
            </button>
          </div>
          <input type="text" class="form-input patch-field" data-key="${key}" value="${String(val).replace(/"/g, '&quot;')}" placeholder="Nhập giá trị thay thế cho ${displayLabel}..." />
        </div>
      `;
    }).join('');

    // Wire change events
    formContainer.querySelectorAll('.patch-field').forEach((input) => {
      input.addEventListener('input', (e) => {
        const k = e.target.getAttribute('data-key');
        patchData[k] = e.target.value;
        if (jsonInput) {
          jsonInput.value = JSON.stringify(patchData, null, 2);
        }
      });
    });

    // Wire row deletion
    formContainer.querySelectorAll('.btn-remove-patch-row').forEach((btn) => {
      btn.addEventListener('click', () => {
        const k = btn.getAttribute('data-key');
        delete patchData[k];
        syncToUi();
      });
    });

    refreshLucideIcons();
  }

  function syncToUi() {
    renderFormFields();
    if (jsonInput) {
      jsonInput.value = JSON.stringify(patchData, null, 2);
    }
  }

  if (jsonInput) {
    jsonInput.addEventListener('input', () => {
      try {
        const parsed = JSON.parse(jsonInput.value);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          patchData = parsed;
          renderFormFields();
        }
      } catch (e) {
        // User is typing invalid JSON, do not disrupt
      }
    });
  }

  if (btnToggleMode) {
    btnToggleMode.addEventListener('click', () => {
      isJsonMode = !isJsonMode;
      if (isJsonMode) {
        if (formView) formView.style.display = 'none';
        if (jsonView) jsonView.style.display = 'block';
        if (toggleLabel) toggleLabel.textContent = 'Chuyển sang Chế Độ Form';
        btnToggleMode.innerHTML = '<i data-lucide="layout-grid"></i> <span id="toggle-patch-mode-label">Chuyển sang Chế Độ Form</span>';
      } else {
        if (formView) formView.style.display = 'block';
        if (jsonView) jsonView.style.display = 'none';
        if (toggleLabel) toggleLabel.textContent = 'Chuyển sang JSON Editor';
        btnToggleMode.innerHTML = '<i data-lucide="code"></i> <span id="toggle-patch-mode-label">Chuyển sang JSON Editor</span>';
      }
      refreshLucideIcons();
    });
  }

  async function handleFile(file) {
    if (!file) return;
    currentFile = file;
    serverPath = null;
    if (patchFileName) patchFileName.textContent = file.name;
    if (patchFileSize) patchFileSize.textContent = formatBytes(file.size);
    if (patchDropZone) {
      const dropContent = patchDropZone.querySelector('.drop-zone-content');
      if (dropContent) dropContent.style.display = 'none';
    }
    if (patchFileInfo) patchFileInfo.style.display = 'flex';
    if (btnRunPatch) btnRunPatch.disabled = true;

    // Upload and inspect tokens automatically
    if (progressContainer) progressContainer.style.display = 'block';
    if (progressFill) progressFill.style.width = '30%';
    if (progressStatus) progressStatus.textContent = 'Đang tải file Word và quét token {{placeholders}}...';

    try {
      const up = await uploadFileToServer(file);
      serverPath = up.filepath || up.file_path || (up.first_file && up.first_file.file_path);

      if (progressFill) progressFill.style.width = '70%';
      if (progressStatus) progressStatus.textContent = 'Đang bóc tách cây AST biến...';

      const res = await fetch('/api/docx/inspect-tokens', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ template_path: serverPath }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'Không thể bóc tách biến từ tệp Word');
      }

      const data = await res.json();
      const vars = data.variables || [];
      patchData = {};
      vars.forEach((v) => {
        patchData[v] = '';
      });

      syncToUi();

      // Render detected tokens badge & pills
      if (statsBox) statsBox.style.display = 'block';
      if (tokenCount) tokenCount.textContent = `${vars.length} biến`;
      if (tokensPillList) {
        if (vars.length > 0) {
          tokensPillList.innerHTML = vars.map((v) => `<span class="tag-badge">{{${v}}}</span>`).join('');
        } else {
          tokensPillList.innerHTML = '<span style="font-size: 0.8rem; color: var(--text-muted); font-style: italic;">Không tìm thấy placeholder dạng {{TOKEN}} trong file</span>';
        }
      }

      if (btnRunPatch) btnRunPatch.disabled = false;
      showToast(`Đã nhận diện ${vars.length} placeholder từ ${file.name}`, 'info');
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error', 6000);
    } finally {
      if (progressContainer) progressContainer.style.display = 'none';
    }
  }

  function clearFile() {
    currentFile = null;
    serverPath = null;
    patchData = {};
    syncToUi();
    if (patchFileInput) patchFileInput.value = '';
    if (patchFileInfo) patchFileInfo.style.display = 'none';
    if (patchDropZone) {
      const dropContent = patchDropZone.querySelector('.drop-zone-content');
      if (dropContent) dropContent.style.display = 'block';
    }
    if (statsBox) statsBox.style.display = 'none';
    if (btnRunPatch) btnRunPatch.disabled = true;
  }

  if (btnBrowsePatch && patchFileInput) {
    btnBrowsePatch.addEventListener('click', (e) => {
      e.stopPropagation();
      patchFileInput.click();
    });
  }

  if (patchDropZone && patchFileInput) {
    patchDropZone.addEventListener('click', () => patchFileInput.click());
    patchDropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      patchDropZone.classList.add('dragover');
    });
    patchDropZone.addEventListener('dragleave', () => patchDropZone.classList.remove('dragover'));
    patchDropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      patchDropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
      }
    });
    patchFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFile(e.target.files[0]);
      }
    });
  }

  if (btnRemovePatch) {
    btnRemovePatch.addEventListener('click', (e) => {
      e.stopPropagation();
      clearFile();
    });
  }

  if (btnRunPatch) {
    btnRunPatch.addEventListener('click', async () => {
      if (!currentFile || !serverPath) return;

      // If in JSON mode, ensure latest json input is parsed
      if (isJsonMode && jsonInput) {
        try {
          patchData = JSON.parse(jsonInput.value);
        } catch (e) {
          showToast('JSON không hợp lệ: ' + e.message, 'error');
          return;
        }
      }

      btnRunPatch.disabled = true;
      if (progressContainer) progressContainer.style.display = 'block';
      if (progressFill) progressFill.style.width = '40%';
      if (progressStatus) progressStatus.textContent = 'Đang áp dụng thay thế văn bản OpenXML bảo toàn style...';

      try {
        const res = await fetch('/api/docx/patch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            template_path: serverPath,
            text_replacements: patchData,
          }),
        });

        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error || `Quick Patch thất bại (${res.status})`);
        }

        const data = await res.json();
        if (progressFill) progressFill.style.width = '100%';

        if (resultFilename) resultFilename.textContent = data.filename || 'patched.docx';
        if (resultTime) resultTime.textContent = new Date().toLocaleTimeString();
        if (btnDownload) btnDownload.setAttribute('href', data.download_url || `/api/download/${data.filename}`);

        if (emptyState) emptyState.style.display = 'none';
        if (resultBox) resultBox.style.display = 'block';

        showToast('Vá văn bản Word thành công bảo toàn 100% định dạng (ERR_DOCX_002)!', 'success');

        setTimeout(() => {
          if (progressContainer) progressContainer.style.display = 'none';
        }, 1200);
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error', 6000);
        if (progressStatus) progressStatus.textContent = `Thất bại: ${err.message}`;
      } finally {
        btnRunPatch.disabled = false;
      }
    });
  }
}


  initDocxQuickPatch();
}
