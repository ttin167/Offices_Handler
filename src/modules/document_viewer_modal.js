/**
 * src/modules/document_viewer_modal.js — In-Browser Universal Document Viewer Modal
 * Supports: PDF, DOCX (AST-driven rich render), and XLSX (interactive multi-sheet grid).
 * Enforces: Chốt chặn G3 (Shimmer Skeleton UX during dynamic lazy-load and conversion).
 */

export function initDocumentViewerModal({ state, showToast, refreshLucideIcons }) {
  const modal = document.getElementById('modal-document-viewer');
  if (!modal) return {};

  const viewerTitle = document.getElementById('doc-viewer-filename');
  const viewerBadge = document.getElementById('doc-viewer-badge');
  const viewerContainer = document.getElementById('doc-viewer-content');
  const btnClose = document.getElementById('btn-close-doc-viewer');
  const btnZoomIn = document.getElementById('btn-viewer-zoom-in');
  const btnZoomOut = document.getElementById('btn-viewer-zoom-out');
  const btnZoomReset = document.getElementById('btn-viewer-zoom-reset');
  const btnFullscreen = document.getElementById('btn-viewer-fullscreen');
  const btnDownload = document.getElementById('btn-viewer-download');
  const btnPrint = document.getElementById('btn-viewer-print');
  const zoomDisplay = document.getElementById('viewer-zoom-display');

  let currentZoom = 1.0;
  let currentFileUrl = null;
  let currentFilePath = null;
  let currentFileType = null;
  let currentFilename = '';

  function setZoom(newZoom) {
    currentZoom = Math.max(0.5, Math.min(2.5, newZoom));
    if (zoomDisplay) zoomDisplay.textContent = `${Math.round(currentZoom * 100)}%`;
    const renderTarget = viewerContainer.querySelector('.doc-viewer-scalable');
    if (renderTarget) {
      renderTarget.style.transform = `scale(${currentZoom})`;
      renderTarget.style.transformOrigin = 'top center';
    }
  }

  if (btnZoomIn) btnZoomIn.addEventListener('click', () => setZoom(currentZoom + 0.15));
  if (btnZoomOut) btnZoomOut.addEventListener('click', () => setZoom(currentZoom - 0.15));
  if (btnZoomReset) btnZoomReset.addEventListener('click', () => setZoom(1.0));

  if (btnFullscreen) {
    btnFullscreen.addEventListener('click', () => {
      if (!document.fullscreenElement) {
        modal.requestFullscreen().catch(() => {});
      } else {
        document.exitFullscreen().catch(() => {});
      }
    });
  }

  function closeModal() {
    modal.style.display = 'none';
    if (viewerContainer) viewerContainer.innerHTML = '';
    currentZoom = 1.0;
    if (zoomDisplay) zoomDisplay.textContent = '100%';
    if (document.fullscreenElement) {
      document.exitFullscreen().catch(() => {});
    }
  }

  if (btnClose) btnClose.addEventListener('click', closeModal);

  window.addEventListener('keydown', (e) => {
    if (modal.style.display === 'flex' && e.key === 'Escape') {
      closeModal();
    }
  });

  if (btnDownload) {
    btnDownload.addEventListener('click', () => {
      if (currentFileUrl) {
        const a = document.createElement('a');
        a.href = currentFileUrl;
        a.download = currentFilename || 'downloaded_document';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
      }
    });
  }

  if (btnPrint) {
    btnPrint.addEventListener('click', () => {
      // If enterprise print modal exists, open it with this file
      const printInput = document.getElementById('print-file-path');
      const modalPrint = document.getElementById('modal-print');
      if (printInput && modalPrint && currentFilePath) {
        printInput.value = currentFilePath;
        modalPrint.style.display = 'flex';
        closeModal();
      } else if (currentFileUrl) {
        window.open(currentFileUrl, '_blank')?.print();
      }
    });
  }

  /**
   * Chốt chặn G3: Render Skeleton Shimmer Loading
   */
  function renderShimmerSkeleton(fileType) {
    if (!viewerContainer) return;
    viewerContainer.innerHTML = `
      <div class="viewer-shimmer-wrapper">
        <div class="shimmer-header">
          <div class="shimmer-box shimmer-title"></div>
          <div class="shimmer-box shimmer-tag"></div>
        </div>
        <div class="shimmer-toolbar">
          <div class="shimmer-box shimmer-btn"></div>
          <div class="shimmer-box shimmer-btn"></div>
          <div class="shimmer-box shimmer-btn"></div>
          <div class="shimmer-box shimmer-btn-wide"></div>
        </div>
        <div class="shimmer-body ${fileType === 'xlsx' ? 'shimmer-grid' : 'shimmer-page'}">
          ${fileType === 'xlsx' ? `
            <div class="shimmer-row shimmer-grid-header"></div>
            <div class="shimmer-row"></div>
            <div class="shimmer-row"></div>
            <div class="shimmer-row"></div>
            <div class="shimmer-row"></div>
            <div class="shimmer-row"></div>
            <div class="shimmer-row"></div>
          ` : `
            <div class="shimmer-line line-90"></div>
            <div class="shimmer-line line-75"></div>
            <div class="shimmer-line line-100"></div>
            <div class="shimmer-line line-60"></div>
            <div class="shimmer-card"></div>
            <div class="shimmer-line line-85"></div>
            <div class="shimmer-line line-95"></div>
          `}
        </div>
        <div class="shimmer-status">
          <div class="shimmer-spinner"></div>
          <span>Đang nạp và render cấu trúc ${fileType.toUpperCase()} (Chốt chặn G3 Shimmer UX)...</span>
        </div>
      </div>
    `;
  }

  /**
   * PDF Viewer: Embedded Object / Iframe with Controls
   */
  function renderPdfViewer(url) {
    viewerContainer.innerHTML = `
      <div class="doc-viewer-scalable pdf-frame-container" style="width: 100%; height: 100%;">
        <iframe
          src="${url}#toolbar=1&navpanes=1&scrollbar=1"
          class="pdf-viewer-iframe"
          title="PDF Document Preview"
          frameborder="0"
        ></iframe>
      </div>
    `;
  }

  /**
   * XLSX Viewer: Multi-Sheet Interactive Grid
   */
  function renderXlsxViewer(snapshot) {
    const sheets = snapshot.sheets || [];
    if (sheets.length === 0) {
      viewerContainer.innerHTML = '<div class="viewer-empty-msg">Tệp bảng tính không có dữ liệu để hiển thị.</div>';
      return;
    }

    let activeSheetIndex = 0;

    function buildSheetHtml(sheet) {
      const rows = sheet.rows || [];
      const stats = sheet.stats || {};
      const maxCol = stats.max_column || 10;

      // Generate column letters (A, B, C...)
      const colHeaders = [];
      for (let c = 1; c <= maxCol; c++) {
        let colLetter = '';
        let temp = c;
        while (temp > 0) {
          const mod = (temp - 1) % 26;
          colLetter = String.fromCharCode(65 + mod) + colLetter;
          temp = Math.floor((temp - mod) / 26);
        }
        colHeaders.push(colLetter);
      }

      let html = `
        <div class="xlsx-formula-bar">
          <div class="formula-cell-address" id="xlsx-active-cell">A1</div>
          <div class="formula-fx"><i data-lucide="function-square"></i> fx</div>
          <input type="text" class="formula-cell-value" id="xlsx-active-value" readonly value="" placeholder="Công thức / Giá trị ô tính" />
        </div>
        <div class="doc-viewer-scalable xlsx-grid-scroll">
          <table class="xlsx-render-table">
            <thead>
              <tr>
                <th class="xlsx-row-num-col"></th>
                ${colHeaders.map(ch => `<th class="xlsx-col-header">${ch}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
      `;

      rows.forEach((rowObj, rIdx) => {
        const rowNum = rowObj.row_number || (rIdx + 1);
        const cells = rowObj.cells || [];
        const cellsByCol = {};
        cells.forEach(c => { cellsByCol[c.column_number] = c; });

        html += `<tr><td class="xlsx-row-num">${rowNum}</td>`;
        for (let colIdx = 1; colIdx <= maxCol; colIdx++) {
          const cData = cellsByCol[colIdx];
          const val = cData ? (cData.formatted_value ?? cData.value ?? '') : '';
          const formula = cData ? (cData.formula || '') : '';
          const style = [];
          if (cData && cData.bold) style.push('font-weight: bold;');
          if (cData && cData.italic) style.push('font-style: italic;');
          if (cData && cData.font_color) style.push(`color: ${cData.font_color};`);
          if (cData && cData.fill_color) style.push(`background-color: ${cData.fill_color};`);
          if (cData && cData.alignment) style.push(`text-align: ${cData.alignment.toLowerCase()};`);

          html += `
            <td
              class="xlsx-cell ${formula ? 'has-formula' : ''}"
              data-addr="${colHeaders[colIdx - 1]}${rowNum}"
              data-val="${escapeHtml(String(val))}"
              data-formula="${escapeHtml(formula)}"
              style="${style.join(' ')}"
            >
              ${escapeHtml(String(val))}
            </td>
          `;
        }
        html += `</tr>`;
      });

      html += `
            </tbody>
          </table>
        </div>
      `;
      return html;
    }

    function renderWrapper() {
      const activeSheet = sheets[activeSheetIndex] || sheets[0];
      viewerContainer.innerHTML = `
        <div class="xlsx-viewer-container">
          <div class="xlsx-grid-wrapper">
            ${buildSheetHtml(activeSheet)}
          </div>
          <div class="xlsx-tabs-bar">
            ${sheets.map((s, idx) => `
              <button type="button" class="xlsx-tab-btn ${idx === activeSheetIndex ? 'active' : ''}" data-sheet-idx="${idx}">
                <i data-lucide="table"></i> ${escapeHtml(s.sheet_name || s.name || `Sheet${idx + 1}`)}
              </button>
            `).join('')}
          </div>
        </div>
      `;

      refreshLucideIcons();

      // Bind cell click to formula bar
      const cells = viewerContainer.querySelectorAll('.xlsx-cell');
      const addrEl = viewerContainer.querySelector('#xlsx-active-cell');
      const valEl = viewerContainer.querySelector('#xlsx-active-value');
      cells.forEach(c => {
        c.addEventListener('click', () => {
          cells.forEach(other => other.classList.remove('selected'));
          c.classList.add('selected');
          const addr = c.getAttribute('data-addr');
          const val = c.getAttribute('data-val');
          const formula = c.getAttribute('data-formula');
          if (addrEl) addrEl.textContent = addr;
          if (valEl) valEl.value = formula ? `=${formula}  [Hiển thị: ${val}]` : val;
        });
      });

      // Bind sheet switch tabs
      const tabBtns = viewerContainer.querySelectorAll('.xlsx-tab-btn');
      tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
          activeSheetIndex = parseInt(btn.getAttribute('data-sheet-idx'), 10);
          renderWrapper();
        });
      });
    }

    renderWrapper();
  }

  /**
   * DOCX Viewer: High-Fidelity AST Document Renderer
   */
  function renderDocxViewer(snapshot) {
    const body = snapshot.body || [];
    if (body.length === 0) {
      viewerContainer.innerHTML = '<div class="viewer-empty-msg">Tài liệu DOCX không có nội dung phần thân.</div>';
      return;
    }

    let html = `
      <div class="doc-viewer-scalable docx-page-scroll">
        <div class="docx-paper-page">
    `;

    body.forEach(elem => {
      const elemType = elem.type || 'paragraph';

      if (elemType === 'paragraph') {
        const styleName = (elem.style_name || '').toLowerCase();
        const align = (elem.alignment || 'LEFT').toLowerCase();
        const runs = elem.runs || [];

        let tag = 'p';
        let cls = 'docx-p';
        if (styleName.includes('heading 1') || styleName === 'title') {
          tag = 'h1'; cls = 'docx-h1';
        } else if (styleName.includes('heading 2')) {
          tag = 'h2'; cls = 'docx-h2';
        } else if (styleName.includes('heading 3')) {
          tag = 'h3'; cls = 'docx-h3';
        } else if (styleName.includes('list')) {
          cls = 'docx-list-item';
        }

        html += `<${tag} class="${cls}" style="text-align: ${align};">`;
        runs.forEach(r => {
          const rStyles = [];
          if (r.bold) rStyles.push('font-weight: 700;');
          if (r.italic) rStyles.push('font-style: italic;');
          if (r.underline) rStyles.push('text-decoration: underline;');
          if (r.color) rStyles.push(`color: #${r.color};`);
          if (r.font_size_pt) rStyles.push(`font-size: ${r.font_size_pt}pt;`);

          html += `<span style="${rStyles.join(' ')}">${escapeHtml(r.text || '')}</span>`;
        });
        html += `</${tag}>`;
      } else if (elemType === 'table') {
        const rows = elem.rows || [];
        html += `<table class="docx-rendered-table"><tbody>`;
        rows.forEach((r, rIdx) => {
          html += `<tr>`;
          (r.cells || []).forEach(c => {
            const isHeader = rIdx === 0 && elem.has_header;
            const cTag = isHeader ? 'th' : 'td';
            const cellStyle = [];
            if (c.shading) cellStyle.push(`background-color: #${c.shading};`);
            if (c.vertical_alignment) cellStyle.push(`vertical-align: ${c.vertical_alignment.toLowerCase()};`);

            html += `<${cTag} class="docx-td" style="${cellStyle.join(' ')}">`;
            (c.paragraphs || []).forEach(p => {
              html += `<div class="docx-cell-p">${escapeHtml(p.text || '')}</div>`;
            });
            html += `</${cTag}>`;
          });
          html += `</tr>`;
        });
        html += `</tbody></table>`;
      }
    });

    html += `
        </div>
      </div>
    `;

    viewerContainer.innerHTML = html;
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /**
   * Main Public API: Open Document Viewer
   */
  async function openDocumentViewer({ file, filepath, filename, fileType }) {
    modal.style.display = 'flex';
    currentZoom = 1.0;
    if (zoomDisplay) zoomDisplay.textContent = '100%';

    const name = filename || (file ? file.name : '') || (filepath ? filepath.split(/[\\/]/).pop() : 'Tài Liệu');
    const ext = (fileType || name.split('.').pop() || '').toLowerCase().replace('.', '');

    currentFilename = name;
    currentFilePath = filepath;
    currentFileType = ext;

    if (viewerTitle) viewerTitle.textContent = name;
    if (viewerBadge) {
      viewerBadge.textContent = ext.toUpperCase();
      viewerBadge.className = `viewer-type-badge badge-${ext}`;
    }

    // Step 1: Render Shimmer Skeleton immediately (Chốt chặn G3)
    renderShimmerSkeleton(ext);

    // Step 2: Lazy load based on format
    try {
      if (ext === 'pdf') {
        let pdfUrl = '';
        if (filepath) {
          const quoted = encodeURIComponent(name);
          pdfUrl = `/api/preview/${quoted}`;
        } else if (file) {
          pdfUrl = URL.createObjectURL(file);
        }
        currentFileUrl = pdfUrl;
        // Simulate smooth transition
        setTimeout(() => {
          renderPdfViewer(pdfUrl);
        }, 300);
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
        const payload = filepath ? { filepath: filepath } : { filename: name };
        const res = await fetch('/api/xlsx/read', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        if (!res.ok) {
          throw new Error(`Không thể đọc tệp bảng tính: ${res.statusText}`);
        }
        const data = await res.json();
        const snapshot = data.snapshot || data;
        setTimeout(() => {
          renderXlsxViewer(snapshot);
        }, 350);
      } else if (ext === 'docx') {
        const payload = filepath ? { filepath: filepath } : { filename: name };
        const res = await fetch('/api/docx/read', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        if (!res.ok) {
          throw new Error(`Không thể đọc tài liệu DOCX: ${res.statusText}`);
        }
        const data = await res.json();
        const snapshot = data.snapshot || data;
        setTimeout(() => {
          renderDocxViewer(snapshot);
        }, 350);
      } else {
        // Fallback generic preview
        viewerContainer.innerHTML = `
          <div class="viewer-fallback-box">
            <i data-lucide="file-text" style="width: 48px; height: 48px; color: var(--color-primary, #38bdf8);"></i>
            <h3>Xem trước định dạng .${ext.toUpperCase()}</h3>
            <p>Tệp "${name}" đã được tải lên thành công. Bạn có thể tải về hoặc gửi lệnh in trực tiếp.</p>
          </div>
        `;
        refreshLucideIcons();
      }
    } catch (err) {
      viewerContainer.innerHTML = `
        <div class="viewer-error-box">
          <i data-lucide="alert-circle" style="width: 48px; height: 48px; color: var(--rose-danger, #f43f5e);"></i>
          <h3>Không Thể Xem Trước Tài Liệu</h3>
          <p>${escapeHtml(err.message)}</p>
          <button type="button" class="btn btn-secondary" id="btn-viewer-retry">Thử lại</button>
        </div>
      `;
      refreshLucideIcons();
      const retryBtn = viewerContainer.querySelector('#btn-viewer-retry');
      if (retryBtn) {
        retryBtn.addEventListener('click', () => {
          openDocumentViewer({ file, filepath, filename, fileType });
        });
      }
    }
  }

  // Attach global hook
  window.openDocumentViewer = openDocumentViewer;

  return {
    openDocumentViewer,
    closeModal,
  };
}

