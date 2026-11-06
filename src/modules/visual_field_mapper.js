/**
 * src/modules/visual_field_mapper.js — Interactive Two-Column Visual Field Mapper for Batch Mail Merge
 * Matches Word template tokens ({{token}}) to dataset columns (CSV/XLSX/JSON) with Auto-Match and Live Preview.
 */

export function initVisualFieldMapper({ state, showToast, refreshLucideIcons, uploadFileToServer }) {
  const container = document.getElementById('visual-field-mapper-container');
  if (!container) return {};

  let currentTokens = [];
  let currentColumns = [];
  let currentSampleRecords = [];
  let fieldMappings = {}; // { token: columnName }

  function normalizeStr(str) {
    if (!str) return '';
    return str
      .toLowerCase()
      .replace(/[{}\s_\-–—]/g, '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  /**
   * Smart Auto-Match Algorithm
   */
  function autoMatchFields() {
    const newMappings = {};
    const usedCols = new Set();

    currentTokens.forEach(token => {
      const normToken = normalizeStr(token);

      // Pass 1: Exact normalized match
      let matchedCol = currentColumns.find(c => normalizeStr(c) === normToken);

      // Pass 2: Substring or semantic synonyms match
      if (!matchedCol) {
        matchedCol = currentColumns.find(c => {
          const normCol = normalizeStr(c);
          return normCol.includes(normToken) || normToken.includes(normCol);
        });
      }

      // Pass 3: Common Vietnamese/English BA field synonyms
      if (!matchedCol) {
        const synonyms = {
          name: ['ten', 'hoten', 'hovaten', 'fullname', 'nguoinhan'],
          fullname: ['ten', 'hoten', 'hovaten', 'name'],
          email: ['mail', 'thu', 'hopthu'],
          phone: ['sdt', 'so', 'dienthoai', 'mobile', 'tel'],
          address: ['diachi', 'noio', 'location'],
          id: ['ma', 'manv', 'code', 'stt', 'soid'],
          salary: ['luong', 'thunhap', 'gross', 'net', 'tienluong'],
          date: ['ngay', 'ngaythang', 'dob', 'ngaysinh'],
          department: ['phongban', 'phong', 'bophan', 'dept'],
          title: ['chucvu', 'vitri', 'position', 'role'],
        };

        const targetSyns = synonyms[normToken] || [];
        matchedCol = currentColumns.find(c => {
          const normCol = normalizeStr(c);
          return targetSyns.some(syn => normCol.includes(syn));
        });
      }

      if (matchedCol) {
        newMappings[token] = matchedCol;
        usedCols.add(matchedCol);
      } else {
        newMappings[token] = '';
      }
    });

    fieldMappings = newMappings;
    renderMappingUI();
    showToast(`Đã tự động khớp ${Object.values(newMappings).filter(Boolean).length}/${currentTokens.length} trường!`, 'info');
  }

  function renderMappingUI() {
    if (!currentTokens || currentTokens.length === 0) {
      container.innerHTML = `
        <div class="vfm-empty-state">
          <i data-lucide="layers" style="width: 36px; height: 36px; color: var(--text-muted, #64748b);"></i>
          <p>Tải lên tệp mẫu Word (.docx) và tệp dữ liệu nguồn (.xlsx/.csv) để kích hoạt ánh xạ trường trực quan.</p>
        </div>
      `;
      refreshLucideIcons();
      return;
    }

    const mappedCount = Object.values(fieldMappings).filter(Boolean).length;
    const totalCount = currentTokens.length;
    const pct = totalCount > 0 ? Math.round((mappedCount / totalCount) * 100) : 0;

    let html = `
      <div class="vfm-wrapper">
        <div class="vfm-header-bar">
          <div class="vfm-stats">
            <span class="vfm-badge ${mappedCount === totalCount ? 'badge-success' : 'badge-warning'}">
              <i data-lucide="${mappedCount === totalCount ? 'check-circle-2' : 'alert-circle'}"></i>
              ${mappedCount}/${totalCount} trường đã nối (${pct}%)
            </span>
          </div>
          <div class="vfm-actions">
            <button type="button" class="btn btn-sm btn-secondary" id="btn-vfm-auto-match">
              <i data-lucide="sparkles"></i> Tự Động Khớp (Auto-Match)
            </button>
            <button type="button" class="btn btn-sm btn-outline" id="btn-vfm-clear">
              <i data-lucide="trash-2"></i> Xóa Ánh Xạ
            </button>
          </div>
        </div>

        <div class="vfm-progress-track">
          <div class="vfm-progress-bar" style="width: ${pct}%;"></div>
        </div>

        <div class="vfm-grid">
          <div class="vfm-column vfm-left-col">
            <div class="vfm-col-title">
              <i data-lucide="tag"></i> Trường Biến Word (Template Tokens)
            </div>
          </div>
          <div class="vfm-column vfm-middle-col">
            <div class="vfm-col-title">
              <i data-lucide="link-2"></i> Trạng Thái Ánh Xạ
            </div>
          </div>
          <div class="vfm-column vfm-right-col">
            <div class="vfm-col-title">
              <i data-lucide="table"></i> Cột Dữ Liệu Nguồn (Dataset Columns)
            </div>
          </div>
        </div>

        <div class="vfm-rows-container">
    `;

    currentTokens.forEach(token => {
      const selectedCol = fieldMappings[token] || '';
      const isMapped = Boolean(selectedCol);

      // Find sample value if column selected
      let sampleVal = '';
      if (selectedCol && currentSampleRecords.length > 0) {
        sampleVal = currentSampleRecords[0][selectedCol] ?? '';
      }

      html += `
        <div class="vfm-row ${isMapped ? 'is-mapped' : 'is-unmapped'}" data-token="${escapeHtml(token)}">
          <div class="vfm-token-cell">
            <span class="vfm-token-pill">
              <i data-lucide="tag"></i> {{${escapeHtml(token)}}}
            </span>
          </div>

          <div class="vfm-connector-cell">
            <div class="vfm-link-arrow ${isMapped ? 'active' : ''}">
              <i data-lucide="${isMapped ? 'arrow-right' : 'minus'}"></i>
            </div>
          </div>

          <div class="vfm-select-cell">
            <select class="vfm-select" data-token="${escapeHtml(token)}">
              <option value="">-- Chọn cột nguồn --</option>
              ${currentColumns.map(col => `
                <option value="${escapeHtml(col)}" ${col === selectedCol ? 'selected' : ''}>
                  ${escapeHtml(col)}
                </option>
              `).join('')}
            </select>
            ${sampleVal ? `<span class="vfm-sample-hint" title="Giá trị dòng 1: ${escapeHtml(String(sampleVal))}">Mẫu: ${escapeHtml(String(sampleVal))}</span>` : ''}
          </div>
        </div>
      `;
    });

    html += `
        </div>

        <!-- Live Record Preview Card -->
        <div class="vfm-preview-card">
          <div class="vfm-preview-title">
            <i data-lucide="eye"></i> Xem Trước Dữ Liệu Sau Khi Ghép (Dòng 1):
          </div>
          <div class="vfm-preview-chips">
            ${currentTokens.map(token => {
              const col = fieldMappings[token];
              const val = col && currentSampleRecords.length > 0 ? (currentSampleRecords[0][col] ?? '') : '(chưa ánh xạ)';
              return `
                <div class="vfm-chip ${col ? 'chip-filled' : 'chip-empty'}">
                  <strong>{{${escapeHtml(token)}}}:</strong> <span>${escapeHtml(String(val))}</span>
                </div>
              `;
            }).join('')}
          </div>
        </div>
      </div>
    `;

    container.innerHTML = html;
    refreshLucideIcons();

    // Bind Auto-match button
    const btnAuto = container.querySelector('#btn-vfm-auto-match');
    if (btnAuto) btnAuto.addEventListener('click', autoMatchFields);

    // Bind Clear button
    const btnClear = container.querySelector('#btn-vfm-clear');
    if (btnClear) {
      btnClear.addEventListener('click', () => {
        currentTokens.forEach(t => { fieldMappings[t] = ''; });
        renderMappingUI();
      });
    }

    // Bind select changes
    const selects = container.querySelectorAll('.vfm-select');
    selects.forEach(sel => {
      sel.addEventListener('change', (e) => {
        const token = sel.getAttribute('data-token');
        fieldMappings[token] = e.target.value;
        renderMappingUI();
      });
    });
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
   * Set Template Tokens (called when DOCX template uploaded)
   */
  function setTokens(tokens) {
    currentTokens = tokens || [];
    currentTokens.forEach(t => {
      if (!(t in fieldMappings)) fieldMappings[t] = '';
    });
    if (currentColumns.length > 0) {
      autoMatchFields();
    } else {
      renderMappingUI();
    }
  }

  /**
   * Set Dataset Columns (called when CSV/XLSX uploaded)
   */
  function setDataset({ columns, sample }) {
    currentColumns = columns || [];
    currentSampleRecords = sample || [];
    if (currentTokens.length > 0) {
      autoMatchFields();
    } else {
      renderMappingUI();
    }
  }

  /**
   * Get Active Mappings for Mail Merge execution
   */
  function getMappings() {
    const active = {};
    for (const [token, col] of Object.entries(fieldMappings)) {
      if (col) active[token] = col;
    }
    return active;
  }

  return {
    setTokens,
    setDataset,
    getMappings,
    autoMatchFields,
    renderMappingUI,
  };
}

