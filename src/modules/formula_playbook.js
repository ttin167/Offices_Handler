/**
 * src/modules/formula_playbook.js — Enterprise Excel Formula Playbook Controller.
 * Manages formula catalog fetching, live multi-field filtering, 1-click clipboard copying,
 * and direct dual-sheet Excel (.xlsx) export.
 */

export function initFormulaPlaybook({ showToast, refreshLucideIcons }) {
  const searchInput = document.getElementById('formula-search-input');
  const btnClearSearch = document.getElementById('btn-clear-formula-search');
  const categoryPills = document.querySelectorAll('.category-pill');
  const cardsGrid = document.getElementById('formula-cards-grid');
  const emptyState = document.getElementById('formula-empty-state');
  const filteredCountBadge = document.getElementById('formula-filtered-count');
  const btnExport = document.getElementById('btn-export-formula-playbook');

  // Internal state
  let allFormulas = [];
  let currentCategory = 'all';
  let currentSearch = '';
  let searchDebounceTimer = null;

  // Escape HTML to prevent injection
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Category Color Map for Visual Distinction
  const CATEGORY_COLORS = {
    lookup: { bg: 'rgba(56, 189, 248, 0.12)', border: 'rgba(56, 189, 248, 0.35)', text: '#38bdf8' },
    aggregation: { bg: 'rgba(16, 185, 129, 0.12)', border: 'rgba(16, 185, 129, 0.35)', text: '#10b981' },
    text: { bg: 'rgba(245, 158, 11, 0.12)', border: 'rgba(245, 158, 11, 0.35)', text: '#f59e0b' },
    logic: { bg: 'rgba(168, 85, 247, 0.12)', border: 'rgba(168, 85, 247, 0.35)', text: '#a855f7' },
    dynamic: { bg: 'rgba(236, 72, 153, 0.12)', border: 'rgba(236, 72, 153, 0.35)', text: '#ec4899' },
    financial: { bg: 'rgba(99, 102, 241, 0.12)', border: 'rgba(99, 102, 241, 0.35)', text: '#818cf8' },
  };

  // Fetch formulas from API
  async function loadFormulaCatalog() {
    try {
      const res = await fetch('/api/xlsx/formulas');
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}`);
      }
      const data = await res.json();
      if (data && data.formulas) {
        allFormulas = data.formulas;
        applyFilterAndRender();
      }
    } catch (err) {
      console.error('Failed to load formula catalog:', err);
      if (cardsGrid) {
        cardsGrid.innerHTML = `
          <div class="formula-error-box">
            <p>Không thể kết nối danh mục công thức: ${escapeHtml(err.message)}</p>
            <button type="button" class="btn-sm-ghost" id="btn-retry-load-formulas">Thử lại</button>
          </div>
        `;
        const retryBtn = document.getElementById('btn-retry-load-formulas');
        if (retryBtn) {
          retryBtn.addEventListener('click', loadFormulaCatalog);
        }
      }
    }
  }

  // Copy text to clipboard with fallback
  async function copyToClipboard(text, btnElement) {
    let success = false;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        success = true;
      } else {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.style.position = 'fixed';
        textarea.style.left = '-9999px';
        document.body.appendChild(textarea);
        textarea.select();
        success = document.execCommand('copy');
        document.body.removeChild(textarea);
      }
    } catch (err) {
      console.warn('Copy error:', err);
      success = false;
    }

    if (success) {
      if (btnElement) {
        const originalHtml = btnElement.innerHTML;
        btnElement.innerHTML = '<i data-lucide="check"></i> <span>Đã chép!</span>';
        btnElement.classList.add('copied');
        refreshLucideIcons();
        setTimeout(() => {
          btnElement.innerHTML = originalHtml;
          btnElement.classList.remove('copied');
          refreshLucideIcons();
        }, 1800);
      }
      showToast('Đã sao chép công thức vào bộ nhớ tạm!', 'success', 2500);
    } else {
      showToast('Không thể sao chép tự động. Hãy bôi đen và bấm Ctrl+C.', 'error');
    }
  }

  // Filter & Render Cards
  function applyFilterAndRender() {
    if (!cardsGrid) return;

    const query = currentSearch.trim().toLowerCase();
    const filtered = allFormulas.filter((f) => {
      // 1. Category check
      if (currentCategory !== 'all' && f.category !== currentCategory) {
        return false;
      }
      // 2. Search query check
      if (!query) return true;

      const matchName = (f.name || '').toLowerCase().includes(query);
      const matchCategoryVi = (f.category_vi || '').toLowerCase().includes(query);
      const matchTemplate = (f.template || '').toLowerCase().includes(query);
      const matchExample = (f.example || '').toLowerCase().includes(query);
      const matchWhen = (f.when_to_use || '').toLowerCase().includes(query);
      const matchWhy = (f.why_choose || '').toLowerCase().includes(query);

      return matchName || matchCategoryVi || matchTemplate || matchExample || matchWhen || matchWhy;
    });

    if (filteredCountBadge) {
      filteredCountBadge.textContent = filtered.length;
    }

    if (filtered.length === 0) {
      cardsGrid.innerHTML = '';
      if (emptyState) emptyState.style.display = 'block';
      return;
    }

    if (emptyState) emptyState.style.display = 'none';

    cardsGrid.innerHTML = filtered.map((item, index) => {
      const colors = CATEGORY_COLORS[item.category] || {
        bg: 'rgba(255, 255, 255, 0.08)',
        border: 'rgba(255, 255, 255, 0.2)',
        text: 'var(--text-primary)',
      };

      return `
        <article class="formula-card glass-panel" data-formula-index="${index}">
          <!-- Card Header -->
          <div class="formula-card-header">
            <div class="formula-card-title-group">
              <span class="formula-num-badge">#${index + 1}</span>
              <h3 class="formula-func-name">${escapeHtml(item.name)}</h3>
              <span class="formula-category-pill" style="background:${colors.bg}; border-color:${colors.border}; color:${colors.text};">
                ${escapeHtml(item.category_vi)}
              </span>
            </div>
            <button type="button" class="btn-copy-formula" data-copy-target="${escapeHtml(item.example || item.template)}" title="Sao chép công thức thực tế">
              <i data-lucide="copy"></i>
              <span>Copy</span>
            </button>
          </div>

          <!-- 1. Cú pháp & Công thức mẫu -->
          <div class="formula-code-section">
            <div class="formula-code-block">
              <span class="code-block-label">Cú pháp chuẩn:</span>
              <code class="formula-code-syntax">${escapeHtml(item.template)}</code>
            </div>
            <div class="formula-code-block formula-example-block">
              <div class="code-block-header">
                <span class="code-block-label">Công thức mẫu thực chiến:</span>
                <button type="button" class="btn-copy-inline" data-copy-text="${escapeHtml(item.example)}" title="Chép công thức này">
                  <i data-lucide="copy"></i>
                </button>
              </div>
              <code class="formula-code-example">${escapeHtml(item.example)}</code>
            </div>
          </div>

          <!-- 2. Dùng trong trường hợp nào? -->
          <div class="formula-info-section when-to-use-section">
            <div class="info-section-header">
              <span class="info-icon when-icon"><i data-lucide="briefcase"></i></span>
              <h4 class="info-title">Dùng trong trường hợp nào?</h4>
            </div>
            <p class="info-text">${escapeHtml(item.when_to_use)}</p>
          </div>

          <!-- 3. Tại sao lại chọn hàm đó của Excel? -->
          <div class="formula-info-section why-choose-section">
            <div class="info-section-header">
              <span class="info-icon why-icon"><i data-lucide="lightbulb"></i></span>
              <h4 class="info-title">Tại sao lại chọn hàm này?</h4>
            </div>
            <p class="info-text">${escapeHtml(item.why_choose)}</p>
          </div>
        </article>
      `;
    }).join('');

    // Attach click handlers to copy buttons
    const copyBtns = cardsGrid.querySelectorAll('.btn-copy-formula, .btn-copy-inline');
    copyBtns.forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const textToCopy = btn.getAttribute('data-copy-target') || btn.getAttribute('data-copy-text');
        if (textToCopy) {
          copyToClipboard(textToCopy, btn);
        }
      });
    });

    refreshLucideIcons();
  }

  // Bind category pill clicks
  categoryPills.forEach((pill) => {
    pill.addEventListener('click', () => {
      categoryPills.forEach((p) => p.classList.remove('active'));
      pill.classList.add('active');
      currentCategory = pill.getAttribute('data-category') || 'all';
      applyFilterAndRender();
    });
  });

  // Bind live search input
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const val = e.target.value;
      currentSearch = val;
      if (btnClearSearch) {
        btnClearSearch.style.display = val.length > 0 ? 'flex' : 'none';
      }
      clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(() => {
        applyFilterAndRender();
      }, 150);
    });
  }

  if (btnClearSearch && searchInput) {
    btnClearSearch.addEventListener('click', () => {
      searchInput.value = '';
      currentSearch = '';
      btnClearSearch.style.display = 'none';
      applyFilterAndRender();
      searchInput.focus();
    });
  }

  // Bind Excel Export Button
  if (btnExport) {
    btnExport.addEventListener('click', () => {
      btnExport.disabled = true;
      const originalText = btnExport.innerHTML;
      btnExport.innerHTML = '<i data-lucide="refresh-cw"></i> <span>ĐANG KHỞI TẠO TỆP EXCEL...</span>';
      refreshLucideIcons();

      // Download via hidden iframe or anchor to avoid page reload
      const downloadLink = document.createElement('a');
      downloadLink.href = '/api/xlsx/formulas/export';
      downloadLink.download = 'Excel_Formula_Playbook.xlsx';
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);

      showToast('Đang tạo và tải về tệp Excel 2 Sheets (Playbook & Live Demo)...', 'success', 3500);

      setTimeout(() => {
        btnExport.disabled = false;
        btnExport.innerHTML = originalText;
        refreshLucideIcons();
      }, 2000);
    });
  }

  // Initial load
  loadFormulaCatalog();
}
