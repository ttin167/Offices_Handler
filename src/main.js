/**
 * ANTIGRAVITY OFFICE STUDIO — CLIENT LOGIC CONTROLLER
 * Full-Cycle PDF Hub, Document Matrix, Presentation Studio & Diagram Engine
 */

document.addEventListener('DOMContentLoaded', () => {
  // Global Application State
  const state = {
    theme: localStorage.getItem('antigravity_theme') || 'dark',
    activeTab: 'presentation',
    apiConnected: false,

    // Presentation Studio State
    pptxFile: null,
    pptxServerPath: null,
    pptxTheme: 'thesis_blue',
    templateServerPath: null,

    // Document Hub State
    docFile: null,
    docServerPath: null,
    targetFormat: '.docx',

    // Diagram Studio State
    diagramEngine: 'mermaid',
    diagramCode: '',
  };

  // Preset Diagram Templates
  const DIAGRAM_PRESETS = {
    mermaid_sequence: `sequenceDiagram
  autonumber
  actor User as Khách Hàng
  participant Web as Web Studio (Vite)
  participant API as Python API Server
  participant Engine as Office Engine (OpenXML)

  User->>Web: Tải lên tài liệu PDF / DOCX
  Web->>API: POST /api/upload (Multipart)
  API-->>Web: Trả về file_path & metadata
  Web->>API: POST /api/pdf-to-pptx (Theme: Thesis Blue)
  API->>Engine: Bóc tách AST & Render 16:9 Widescreen
  Engine-->>API: Trả về presentation.pptx
  API-->>Web: 200 OK + Spec JSON & Tải Tệp
  Web-->>User: Hiển thị bộ thẻ Slide & Nút Tải Về`,

    plantuml_c4: `@startuml
!include <C4/C4_Context>
title Sơ Đồ Kiến Trúc Hệ Thống (C4 Context)

Person(user, "Chuyên Viên / Giảng Viên", "Người soạn thảo tài liệu kỹ thuật & slide thuyết trình")
System(studio, "Antigravity Office Studio", "Nền tảng chuyển đổi tài liệu & biên tập slide đa năng")
System_Ext(ms_office, "Microsoft Office / 365", "Đọc và trình chiếu tệp PPTX, DOCX chuẩn OpenXML")
System_Ext(pdf_reader, "Trình Đọc PDF", "Hiển thị tài liệu phân giải cao 300+ DPI")

Rel(user, studio, "Tải lên DOCX/PDF & Tùy chọn Theme")
Rel(studio, ms_office, "Xuất bản tệp 16:9 Zero-Overflow")
Rel(studio, pdf_reader, "Xuất bản PDF chất lượng cao")
@enduml`,

    mermaid_erd: `erDiagram
  DOCUMENT ||--o{ SLIDE : contains
  SLIDE ||--o{ BLOCK : composed_of
  BLOCK ||--o| TABLE_DATA : renders
  DOCUMENT {
    string id PK
    string filename
    string original_format
    int page_count
  }
  SLIDE {
    int slide_index PK
    string title
    string layout_type
    string background_color
  }
  BLOCK {
    string block_id PK
    string block_type
    string content
  }
  TABLE_DATA {
    string table_id PK
    int row_count
    int col_count
  }`,
  };

  // =========================================================================
  // 1. THEME CONTROLLER
  // =========================================================================
  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('antigravity_theme', theme);
    const themeBtn = document.getElementById('btn-toggle-theme');
    if (themeBtn) {
      themeBtn.innerHTML = theme === 'dark' ? '<span>🌙</span>' : '<span>☀️</span>';
    }
  }

  const btnToggleTheme = document.getElementById('btn-toggle-theme');
  if (btnToggleTheme) {
    btnToggleTheme.addEventListener('click', () => {
      applyTheme(state.theme === 'dark' ? 'light' : 'dark');
    });
  }
  applyTheme(state.theme);

  // =========================================================================
  // 2. TOAST NOTIFICATIONS
  // =========================================================================
  function showToast(message, type = 'info', duration = 4000) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';

    toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = 'toastSlideOut 0.3s cubic-bezier(0.4, 0, 0.2, 1) forwards';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, duration);
  }

  // =========================================================================
  // 3. SERVER HEALTH POLLING
  // =========================================================================
  async function checkServerHealth() {
    const badge = document.getElementById('server-status-badge');
    const label = badge ? badge.querySelector('.status-label') : null;

    try {
      const res = await fetch('/api/health');
      if (res.ok) {
        state.apiConnected = true;
        if (badge) {
          badge.className = 'badge-status';
          if (label) label.textContent = 'API Bridge Sẵn Sàng (:8000)';
        }
      } else {
        throw new Error(`HTTP ${res.status}`);
      }
    } catch {
      state.apiConnected = false;
      if (badge) {
        badge.className = 'badge-status disconnected';
        if (label) label.textContent = 'Mất kết nối API Server';
      }
    }
  }

  checkServerHealth();
  setInterval(checkServerHealth, 10000);

  // =========================================================================
  // 4. NAVIGATION TABS
  // =========================================================================
  const navTabs = document.querySelectorAll('.nav-tab');
  const panels = {
    presentation: document.getElementById('panel-presentation'),
    spreadsheet: document.getElementById('panel-spreadsheet'),
    documents: document.getElementById('panel-documents'),
    diagrams: document.getElementById('panel-diagrams'),
  };

  navTabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      const targetTab = tab.getAttribute('data-tab');
      if (!panels[targetTab]) return;

      navTabs.forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');

      Object.values(panels).forEach((p) => p && p.classList.remove('active'));
      panels[targetTab].classList.add('active');
      state.activeTab = targetTab;
    });
  });

  // =========================================================================
  // 5. HELPER: UPLOAD FILE
  // =========================================================================
  async function uploadFileToServer(file) {
    const formData = new FormData();
    formData.append('file', file);

    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData,
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Upload thất bại: ${res.statusText}`);
    }

    return await res.json();
  }

  function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }

  // =========================================================================
  // 6. TAB 1: PRESENTATION STUDIO CONTROLLER
  // =========================================================================
  const pptxDropZone = document.getElementById('pptx-drop-zone');
  const pptxFileInput = document.getElementById('pptx-file-input');
  const btnBrowsePptxFile = document.getElementById('btn-browse-pptx-file');
  const pptxFileInfo = document.getElementById('pptx-file-info');
  const pptxFileName = document.getElementById('pptx-file-name');
  const pptxFileSize = document.getElementById('pptx-file-size');
  const pptxFileExt = document.getElementById('pptx-file-ext');
  const btnRemovePptxFile = document.getElementById('btn-remove-pptx-file');
  const pptxPdfPagesGroup = document.getElementById('pptx-pdf-pages-group');
  const btnGeneratePptx = document.getElementById('btn-generate-pptx');
  const pptxProgress = document.getElementById('pptx-progress-container');
  const pptxProgressFill = document.getElementById('pptx-progress-fill');
  const pptxProgressStatus = document.getElementById('pptx-progress-status');

  function handlePresentationFileSelected(file) {
    if (!file) return;
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!['.docx', '.pdf', '.md'].includes(ext)) {
      showToast('Định dạng tệp không được hỗ trợ. Hãy chọn .docx, .pdf hoặc .md', 'error');
      return;
    }

    state.pptxFile = file;
    state.pptxServerPath = null;

    if (pptxFileName) pptxFileName.textContent = file.name;
    if (pptxFileSize) pptxFileSize.textContent = formatBytes(file.size);
    if (pptxFileExt) pptxFileExt.textContent = ext.replace('.', '').toUpperCase();

    if (pptxDropZone) {
      const dropContent = pptxDropZone.querySelector('.drop-zone-content');
      if (dropContent) dropContent.style.display = 'none';
    }
    if (pptxFileInfo) pptxFileInfo.style.display = 'flex';

    // Show PDF page range option if PDF
    if (pptxPdfPagesGroup) {
      pptxPdfPagesGroup.style.display = ext === '.pdf' ? 'block' : 'none';
    }

    if (btnGeneratePptx) btnGeneratePptx.disabled = false;
    showToast(`Đã chọn tệp nguồn: ${file.name}`, 'info');
  }

  function clearPresentationFile() {
    state.pptxFile = null;
    state.pptxServerPath = null;
    if (pptxFileInput) pptxFileInput.value = '';
    if (pptxFileInfo) pptxFileInfo.style.display = 'none';
    if (pptxDropZone) {
      const dropContent = pptxDropZone.querySelector('.drop-zone-content');
      if (dropContent) dropContent.style.display = 'block';
    }
    if (pptxPdfPagesGroup) pptxPdfPagesGroup.style.display = 'none';
    if (btnGeneratePptx) btnGeneratePptx.disabled = true;
  }

  if (btnBrowsePptxFile && pptxFileInput) {
    btnBrowsePptxFile.addEventListener('click', (e) => {
      e.stopPropagation();
      pptxFileInput.click();
    });
  }

  if (pptxDropZone && pptxFileInput) {
    pptxDropZone.addEventListener('click', () => pptxFileInput.click());

    pptxDropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      pptxDropZone.classList.add('dragover');
    });

    pptxDropZone.addEventListener('dragleave', () => {
      pptxDropZone.classList.remove('dragover');
    });

    pptxDropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      pptxDropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handlePresentationFileSelected(e.dataTransfer.files[0]);
      }
    });

    pptxFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handlePresentationFileSelected(e.target.files[0]);
      }
    });
  }

  if (btnRemovePptxFile) {
    btnRemovePptxFile.addEventListener('click', (e) => {
      e.stopPropagation();
      clearPresentationFile();
    });
  }

  // Theme selector radio buttons
  const themeCards = document.querySelectorAll('.theme-card');
  themeCards.forEach((card) => {
    card.addEventListener('click', () => {
      themeCards.forEach((c) => c.classList.remove('active'));
      card.classList.add('active');
      const radio = card.querySelector('input[type="radio"]');
      if (radio) radio.checked = true;
      state.pptxTheme = card.getAttribute('data-theme-val') || 'thesis_blue';
    });
  });

  // Master Slide Template Selector
  const btnBrowseTemplate = document.getElementById('btn-browse-template');
  const pptxTemplateFileInput = document.getElementById('pptx-template-file-input');
  const pptxTemplateInput = document.getElementById('pptx-template-input');

  if (btnBrowseTemplate && pptxTemplateFileInput) {
    btnBrowseTemplate.addEventListener('click', () => pptxTemplateFileInput.click());

    pptxTemplateFileInput.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files.length > 0) {
        const file = e.target.files[0];
        try {
          showToast(`Đang tải lên Master Slide: ${file.name}...`, 'info');
          const res = await uploadFileToServer(file);
          state.templateServerPath = res.filepath;
          if (pptxTemplateInput) pptxTemplateInput.value = file.name;
          showToast('Master Slide đã được nạp thành công!', 'success');
        } catch (err) {
          showToast(`Lỗi upload template: ${err.message}`, 'error');
        }
      }
    });
  }

  // Render Slide Cards into Inspector Deck
  function renderSlideDeckInspector(spec, pptxDownloadUrl, specDownloadUrl) {
    const deckGrid = document.getElementById('slide-cards-grid');
    const emptyState = document.getElementById('deck-empty-state');
    const countBadge = document.getElementById('deck-count-badge');
    const actionsGroup = document.getElementById('deck-download-actions');
    const btnDownloadPptx = document.getElementById('btn-download-pptx');
    const btnDownloadSpec = document.getElementById('btn-download-spec');

    if (!deckGrid || !spec || !spec.slides) return;

    deckGrid.innerHTML = '';
    const slides = spec.slides;

    if (countBadge) countBadge.textContent = `${slides.length} SLIDES (16:9)`;
    if (actionsGroup) actionsGroup.style.display = 'flex';
    if (btnDownloadPptx) btnDownloadPptx.setAttribute('href', pptxDownloadUrl);
    if (btnDownloadSpec) btnDownloadSpec.setAttribute('href', specDownloadUrl);

    slides.forEach((slide, idx) => {
      const card = document.createElement('div');
      card.className = 'slide-card-item';

      const padIndex = String(idx + 1).padStart(2, '0');
      const layoutName = slide.layout || 'content';

      let contentHtml = '';
      if (slide.subtitle) {
        contentHtml += `<p class="slide-card-subtitle" style="font-size: 13px; color: var(--text-secondary); margin-bottom: 8px;">${slide.subtitle}</p>`;
      }

      if (slide.bullets && Array.isArray(slide.bullets) && slide.bullets.length > 0) {
        contentHtml += `<ul class="slide-card-bullets">`;
        slide.bullets.forEach((b) => {
          contentHtml += `<li class="slide-card-bullet-item"><span class="bullet-bullet-icon">▸</span><span>${b}</span></li>`;
        });
        contentHtml += `</ul>`;
      }

      if (slide.table && slide.table.rows && slide.table.rows.length > 0) {
        contentHtml += `<table class="slide-card-table-preview">`;
        slide.table.rows.slice(0, 4).forEach((row, rIdx) => {
          contentHtml += `<tr>`;
          row.forEach((cell) => {
            if (rIdx === 0) {
              contentHtml += `<th>${cell}</th>`;
            } else {
              contentHtml += `<td>${cell}</td>`;
            }
          });
          contentHtml += `</tr>`;
        });
        if (slide.table.rows.length > 4) {
          contentHtml += `<tr><td colspan="${slide.table.rows[0].length}" style="text-align: center; color: var(--text-muted); font-size: 11px;">... và ${slide.table.rows.length - 4} dòng khác</td></tr>`;
        }
        contentHtml += `</table>`;
      }

      let notesHtml = '';
      if (slide.notes) {
        notesHtml = `<div class="slide-card-notes">📝 Ghi chú: ${slide.notes}</div>`;
      }

      card.innerHTML = `
        <div class="slide-card-header">
          <div class="slide-card-badges">
            <span class="slide-num-badge">SLIDE ${padIndex}</span>
            <span class="slide-layout-badge">${layoutName}</span>
          </div>
        </div>
        <div class="slide-card-title">${slide.title || 'Untitled Slide'}</div>
        ${contentHtml}
        ${notesHtml}
      `;
      deckGrid.appendChild(card);
    });

    if (emptyState) emptyState.style.display = 'none';
    deckGrid.style.display = 'flex';
  }

  // Slide Previews & Core v2 State
  state.currentPptxPath = null;
  state.currentPptxFile = null;
  state.pptxSlidePreviews = [];
  state.currentPreviewIndex = 0;
  state.activeDeckView = 'cards';

  // Toggle View: Cards vs Gallery
  const deckViewSwitcher = document.getElementById('deck-view-switcher');
  const viewBtns = deckViewSwitcher ? deckViewSwitcher.querySelectorAll('.view-btn') : [];
  const slideCardsGrid = document.getElementById('slide-cards-grid');
  const slideGalleryGrid = document.getElementById('slide-gallery-grid');

  function setDeckView(view) {
    state.activeDeckView = view;
    viewBtns.forEach((btn) => {
      btn.classList.toggle('active', btn.getAttribute('data-view') === view);
    });
    if (slideCardsGrid) slideCardsGrid.style.display = view === 'cards' ? 'flex' : 'none';
    if (slideGalleryGrid) slideGalleryGrid.style.display = view === 'gallery' ? 'grid' : 'none';
  }

  viewBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      setDeckView(btn.getAttribute('data-view'));
    });
  });

  // Export Slide Previews Action (1080p)
  const btnExportPreviews = document.getElementById('btn-export-previews');
  const galleryCountBadge = document.getElementById('gallery-count-badge');

  async function exportSlidePreviews() {
    if (!state.currentPptxPath && !state.currentPptxFile) {
      showToast('Chưa có tệp PowerPoint để xuất ảnh xem trước!', 'error');
      return;
    }

    if (btnExportPreviews) {
      btnExportPreviews.disabled = true;
      btnExportPreviews.innerHTML = '<span>⏳ Đang render 1080p...</span>';
    }

    try {
      showToast('Đang xuất ảnh xem trước chuẩn 1080p bằng Win32 COM / headless...', 'info');
      const res = await fetch('/api/pptx/preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pptx_path: state.currentPptxPath || state.currentPptxFile,
          slides: 'all',
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Xuất ảnh thất bại (${res.status})`);
      }

      const data = await res.json();
      state.pptxSlidePreviews = data.previews || [];
      if (galleryCountBadge) galleryCountBadge.textContent = state.pptxSlidePreviews.length;
      renderSlideGallery(state.pptxSlidePreviews);
      setDeckView('gallery');
      showToast(`Đã xuất thành công ${state.pptxSlidePreviews.length} ảnh slide xem trước 1080p!`, 'success');
    } catch (err) {
      showToast(`Lỗi xuất ảnh xem trước: ${err.message}`, 'error');
    } finally {
      if (btnExportPreviews) {
        btnExportPreviews.disabled = false;
        btnExportPreviews.innerHTML = '<span>📸 Xuất Thư Viện Ảnh (1080p)</span>';
      }
    }
  }

  if (btnExportPreviews) {
    btnExportPreviews.addEventListener('click', exportSlidePreviews);
  }

  // Render Slide Gallery Thumbnails
  function renderSlideGallery(previews) {
    if (!slideGalleryGrid) return;
    slideGalleryGrid.innerHTML = '';

    if (!previews || previews.length === 0) {
      slideGalleryGrid.innerHTML = '<div class="empty-desc" style="grid-column: 1/-1; text-align: center; padding: 20px;">Chưa có ảnh slide nào được xuất. Bấm "Xuất Thư Viện Ảnh (1080p)" để tạo.</div>';
      return;
    }

    previews.forEach((p, idx) => {
      const card = document.createElement('div');
      card.className = 'slide-gallery-card';
      card.innerHTML = `
        <div class="gallery-card-badge">SLIDE ${String(p.slide || idx + 1).padStart(2, '0')}</div>
        <div class="gallery-card-thumb">
          <img src="${p.preview_url}?t=${Date.now()}" alt="Slide ${p.slide}" loading="lazy" />
          <div class="gallery-hover-overlay">
            <span class="overlay-icon">🔍 Phóng to</span>
          </div>
        </div>
      `;
      card.addEventListener('click', () => openSlideModal(idx));
      slideGalleryGrid.appendChild(card);
    });
  }

  // Slide Modal Zoom Controller
  const slideModal = document.getElementById('slide-preview-modal');
  const modalBackdrop = document.getElementById('modal-backdrop');
  const modalCloseBtn = document.getElementById('btn-close-modal');
  const modalSlideImg = document.getElementById('modal-slide-img');
  const modalSlideCaption = document.getElementById('modal-slide-caption');
  const btnModalPrev = document.getElementById('btn-modal-prev');
  const btnModalNext = document.getElementById('btn-modal-next');

  function openSlideModal(index) {
    if (!state.pptxSlidePreviews || state.pptxSlidePreviews.length === 0) return;
    state.currentPreviewIndex = index;
    updateModalContent();
    if (slideModal) slideModal.style.display = 'flex';
  }

  function closeSlideModal() {
    if (slideModal) slideModal.style.display = 'none';
  }

  function updateModalContent() {
    const item = state.pptxSlidePreviews[state.currentPreviewIndex];
    if (!item) return;
    if (modalSlideImg) modalSlideImg.src = `${item.preview_url}?t=${Date.now()}`;
    if (modalSlideCaption) {
      modalSlideCaption.textContent = `Slide ${state.currentPreviewIndex + 1} / ${state.pptxSlidePreviews.length}`;
    }
  }

  if (modalCloseBtn) modalCloseBtn.addEventListener('click', closeSlideModal);
  if (modalBackdrop) modalBackdrop.addEventListener('click', closeSlideModal);

  if (btnModalPrev) {
    btnModalPrev.addEventListener('click', () => {
      if (state.currentPreviewIndex > 0) {
        state.currentPreviewIndex--;
        updateModalContent();
      }
    });
  }

  if (btnModalNext) {
    btnModalNext.addEventListener('click', () => {
      if (state.currentPreviewIndex < state.pptxSlidePreviews.length - 1) {
        state.currentPreviewIndex++;
        updateModalContent();
      }
    });
  }

  document.addEventListener('keydown', (e) => {
    if (!slideModal || slideModal.style.display === 'none') return;
    if (e.key === 'Escape') closeSlideModal();
    if (e.key === 'ArrowLeft' && state.currentPreviewIndex > 0) {
      state.currentPreviewIndex--;
      updateModalContent();
    }
    if (e.key === 'ArrowRight' && state.currentPreviewIndex < state.pptxSlidePreviews.length - 1) {
      state.currentPreviewIndex++;
      updateModalContent();
    }
  });

  // Spec Builder Accordion & Action
  const btnToggleSpecBuilder = document.getElementById('btn-toggle-spec-builder');
  const specBuilderBody = document.getElementById('spec-builder-body');
  const specAccordionArrow = document.getElementById('spec-accordion-arrow');
  const pptxSpecJsonInput = document.getElementById('pptx-spec-json-input');
  const btnBuildFromSpec = document.getElementById('btn-build-from-spec');

  if (btnToggleSpecBuilder && specBuilderBody) {
    btnToggleSpecBuilder.addEventListener('click', () => {
      const isVisible = specBuilderBody.style.display !== 'none';
      specBuilderBody.style.display = isVisible ? 'none' : 'block';
      if (specAccordionArrow) specAccordionArrow.textContent = isVisible ? '▼' : '▲';
    });
  }

  if (btnBuildFromSpec && pptxSpecJsonInput) {
    btnBuildFromSpec.addEventListener('click', async () => {
      const specText = pptxSpecJsonInput.value.trim();
      if (!specText) {
        showToast('Vui lòng nhập nội dung JSON Spec!', 'error');
        return;
      }

      let parsedSpec;
      try {
        parsedSpec = JSON.parse(specText);
      } catch (e) {
        showToast('Cú pháp JSON không hợp lệ: ' + e.message, 'error');
        return;
      }

      btnBuildFromSpec.disabled = true;
      btnBuildFromSpec.innerHTML = '<span>⏳ Đang render PPTX...</span>';

      try {
        const res = await fetch('/api/pptx/build', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            spec: parsedSpec,
            theme: state.pptxTheme,
            template_path: state.templateServerPath,
          }),
        });

        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error || `Build PPTX thất bại (${res.status})`);
        }

        const data = await res.json();
        state.currentPptxFile = data.filename;
        state.currentPptxPath = data.pptx_path;
        state.pptxSlidePreviews = [];
        if (galleryCountBadge) galleryCountBadge.textContent = '0';
        if (deckViewSwitcher) deckViewSwitcher.style.display = 'flex';

        renderSlideDeckInspector(parsedSpec, data.download_url, '#');
        setDeckView('cards');
        showToast(`Tạo thành công file ${data.filename}!`, 'success');
      } catch (err) {
        showToast(`Lỗi build spec: ${err.message}`, 'error');
      } finally {
        btnBuildFromSpec.disabled = false;
        btnBuildFromSpec.innerHTML = '<span>🛠️ Render PPTX Từ Spec JSON</span>';
      }
    });
  }

  // Generate PPTX Handler (Word / PDF)
  if (btnGeneratePptx) {
    btnGeneratePptx.addEventListener('click', async () => {
      if (!state.pptxFile) return;

      btnGeneratePptx.disabled = true;
      if (pptxProgress) pptxProgress.style.display = 'block';
      if (pptxProgressFill) pptxProgressFill.style.width = '20%';
      if (pptxProgressStatus) pptxProgressStatus.textContent = '1/3 Đang tải tệp nguồn lên máy chủ...';

      try {
        // Step 1: Upload source file
        const uploadRes = await uploadFileToServer(state.pptxFile);
        state.pptxServerPath = uploadRes.filepath || uploadRes.file_path || (uploadRes.first_file && uploadRes.first_file.file_path);

        if (pptxProgressFill) pptxProgressFill.style.width = '55%';
        if (pptxProgressStatus) pptxProgressStatus.textContent = '2/3 Đang bóc tách AST & phân mảnh nội dung...';

        const ext = state.pptxFile.name.substring(state.pptxFile.name.lastIndexOf('.')).toLowerCase();
        let apiEndpoint = '/api/docx-to-pptx';

        const preserveBranding = document.getElementById('pptx-preserve-branding')?.checked ?? true;
        const title28pt = document.getElementById('pptx-title-28pt')?.checked ?? true;

        let payload = {
          theme: state.pptxTheme,
          template_pptx: state.templateServerPath || null,
          template_path: state.templateServerPath || null,
          source_path: state.pptxServerPath,
          input_path: state.pptxServerPath,
          preserve_branding: preserveBranding,
          title_font_size: title28pt ? 28.0 : 24.0,
          title_align: title28pt ? 'left' : 'center',
        };

        if (ext === '.pdf') {
          apiEndpoint = '/api/pdf-to-pptx';
          const pagesInput = document.getElementById('pptx-pdf-pages');
          payload.pdf_path = state.pptxServerPath;
          payload.pages = pagesInput && pagesInput.value.trim() ? pagesInput.value.trim() : null;
        } else {
          payload.docx_path = state.pptxServerPath;
        }

        if (pptxProgressFill) pptxProgressFill.style.width = '80%';
        if (pptxProgressStatus) pptxProgressStatus.textContent = '3/3 Đang định dạng 16:9 và render OpenXML PPTX...';

        const genRes = await fetch(apiEndpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        if (!genRes.ok) {
          const err = await genRes.json().catch(() => ({}));
          throw new Error(err.error || `Sinh slide thất bại (${genRes.status})`);
        }

        const data = await genRes.json();
        if (pptxProgressFill) pptxProgressFill.style.width = '100%';
        if (pptxProgressStatus) pptxProgressStatus.textContent = 'Hoàn tất!';

        const pptxFile = data.pptx_file || data.filename || data.output_file || (data.pptx_path ? data.pptx_path.split(/[\\/]/).pop() : '');
        const specFile = data.spec_file || (data.spec_path ? data.spec_path.split(/[\\/]/).pop() : '');
        const pptxDownloadUrl = data.download_url || `/api/download/${pptxFile}`;
        const specDownloadUrl = `/api/download/${specFile}`;

        state.currentPptxFile = pptxFile;
        state.currentPptxPath = data.pptx_path;
        state.pptxSlidePreviews = [];
        if (galleryCountBadge) galleryCountBadge.textContent = '0';
        if (deckViewSwitcher) deckViewSwitcher.style.display = 'flex';

        renderSlideDeckInspector(data.spec, pptxDownloadUrl, specDownloadUrl);
        setDeckView('cards');
        showToast(`Tổng hợp thành công ${data.slides_count} slides chất lượng cao!`, 'success');

        setTimeout(() => {
          if (pptxProgress) pptxProgress.style.display = 'none';
        }, 1200);
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error', 6000);
        if (pptxProgressStatus) pptxProgressStatus.textContent = `Thất bại: ${err.message}`;
      } finally {
        btnGeneratePptx.disabled = false;
      }
    });
  }

  // =========================================================================
  // SPREADSHEET STUDIO CONTROLLER (12 Gates Validator, Mutator, Inspector)
  // =========================================================================

  // 1. Sub-tab switcher
  const subTabBtns = document.querySelectorAll('.sub-tab-btn');
  const subPanels = {
    validator: document.getElementById('subpanel-validator'),
    mutator: document.getElementById('subpanel-mutator'),
    inspector: document.getElementById('subpanel-inspector'),
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
            target_sheet: targetSheetInput && targetSheetInput.value.trim() ? targetSheetInput.value.trim() : null,
            ref_sheet: refSheetInput && refSheetInput.value.trim() ? refSheetInput.value.trim() : 'Example',
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

  function handleXlsxMutFile(file) {
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
        btnRunXlsxInspect.innerHTML = '<span>🔍 TRÍCH XUẤT JSON SNAPSHOT</span>';
      }
    });
  }

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
  matrixCards.forEach((card) => {
    card.addEventListener('click', () => {
      matrixCards.forEach((c) => c.classList.remove('active'));
      card.classList.add('active');
      const radio = card.querySelector('input[type="radio"]');
      if (radio) radio.checked = true;
      state.targetFormat = card.getAttribute('data-target') || '.docx';
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
  // 8. TAB 3: DIAGRAM STUDIO CONTROLLER
  // =========================================================================
  const engineBtns = document.querySelectorAll('.engine-btn');
  const diagramCodeInput = document.getElementById('diagram-code-input');
  const btnSampleSequence = document.getElementById('btn-sample-sequence');
  const btnSampleC4 = document.getElementById('btn-sample-c4');
  const btnSampleErd = document.getElementById('btn-sample-erd');
  const btnRenderDiagram = document.getElementById('btn-render-diagram');
  const diagramCanvasContainer = document.getElementById('diagram-canvas-container');
  const diagramViewport = document.getElementById('diagram-viewport');
  const diagramEmptyState = document.getElementById('diagram-empty-state');
  const diagramPreviewImg = document.getElementById('diagram-preview-img');
  const diagramDownloadActions = document.getElementById('diagram-download-actions');
  const btnDownloadDiagram = document.getElementById('btn-download-diagram');
  const diagramZoomToolbar = document.getElementById('diagram-zoom-toolbar');
  const diagramHintBar = document.getElementById('diagram-hint-bar');
  const btnZoomOut = document.getElementById('btn-zoom-out');
  const btnZoomIn = document.getElementById('btn-zoom-in');
  const btnZoomFit = document.getElementById('btn-zoom-fit');
  const btnZoom100 = document.getElementById('btn-zoom-100');
  const btnZoomOpen = document.getElementById('btn-zoom-open');
  const zoomLevelBadge = document.getElementById('zoom-level-badge');

  let currentZoom = 1.0;
  let viewMode = '100'; // '100' or 'fit'

  function setViewMode(mode) {
    viewMode = mode;
    if (!diagramViewport || !diagramPreviewImg) return;

    if (mode === 'fit') {
      diagramViewport.classList.remove('mode-100');
      diagramViewport.classList.add('mode-fit');
      diagramPreviewImg.style.transform = 'none';
      if (btnZoomFit) btnZoomFit.classList.add('active');
      if (btnZoom100) btnZoom100.classList.remove('active');
      if (zoomLevelBadge) zoomLevelBadge.textContent = 'Fit';
      if (diagramCanvasContainer) diagramCanvasContainer.classList.remove('panning');
    } else {
      diagramViewport.classList.remove('mode-fit');
      diagramViewport.classList.add('mode-100');
      currentZoom = 1.0;
      applyZoomTransform();
      if (btnZoomFit) btnZoomFit.classList.remove('active');
      if (btnZoom100) btnZoom100.classList.add('active');
      if (zoomLevelBadge) zoomLevelBadge.textContent = '100%';
      if (diagramCanvasContainer) diagramCanvasContainer.classList.add('panning');
    }
  }

  function applyZoomTransform() {
    if (!diagramPreviewImg) return;
    if (diagramViewport) {
      diagramViewport.classList.remove('mode-fit');
      diagramViewport.classList.add('mode-100');
    }
    if (btnZoomFit) btnZoomFit.classList.remove('active');
    if (btnZoom100) btnZoom100.classList.toggle('active', Math.abs(currentZoom - 1.0) < 0.01);
    diagramPreviewImg.style.transform = `scale(${currentZoom})`;
    if (zoomLevelBadge) zoomLevelBadge.textContent = `${Math.round(currentZoom * 100)}%`;
    if (diagramCanvasContainer) diagramCanvasContainer.classList.add('panning');
  }

  if (btnZoomIn) {
    btnZoomIn.addEventListener('click', () => {
      currentZoom = Math.min(3.0, currentZoom + 0.25);
      applyZoomTransform();
    });
  }

  if (btnZoomOut) {
    btnZoomOut.addEventListener('click', () => {
      currentZoom = Math.max(0.25, currentZoom - 0.25);
      applyZoomTransform();
    });
  }

  if (btnZoomFit) {
    btnZoomFit.addEventListener('click', () => setViewMode('fit'));
  }

  if (btnZoom100) {
    btnZoom100.addEventListener('click', () => setViewMode('100'));
  }

  if (btnZoomOpen) {
    btnZoomOpen.addEventListener('click', () => {
      if (diagramPreviewImg && diagramPreviewImg.src) {
        window.open(diagramPreviewImg.src, '_blank');
      }
    });
  }

  // Toggle on double-click image
  if (diagramPreviewImg) {
    diagramPreviewImg.addEventListener('dblclick', () => {
      setViewMode(viewMode === 'fit' ? '100' : 'fit');
    });
  }

  // Mouse wheel zoom when hovering canvas with Ctrl key
  if (diagramCanvasContainer) {
    diagramCanvasContainer.addEventListener('wheel', (e) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const delta = e.deltaY < 0 ? 0.15 : -0.15;
        currentZoom = Math.min(3.0, Math.max(0.25, currentZoom + delta));
        applyZoomTransform();
      }
    }, { passive: false });

    // Drag to Pan functionality
    let isDown = false;
    let startX = 0;
    let startY = 0;
    let scrollLeft = 0;
    let scrollTop = 0;

    diagramCanvasContainer.addEventListener('mousedown', (e) => {
      if (viewMode === 'fit' && currentZoom <= 1.0) return;
      isDown = true;
      diagramCanvasContainer.classList.add('is-dragging');
      startX = e.pageX - diagramCanvasContainer.offsetLeft;
      startY = e.pageY - diagramCanvasContainer.offsetTop;
      scrollLeft = diagramCanvasContainer.scrollLeft;
      scrollTop = diagramCanvasContainer.scrollTop;
    });

    diagramCanvasContainer.addEventListener('mouseleave', () => {
      isDown = false;
      diagramCanvasContainer.classList.remove('is-dragging');
    });

    diagramCanvasContainer.addEventListener('mouseup', () => {
      isDown = false;
      diagramCanvasContainer.classList.remove('is-dragging');
    });

    diagramCanvasContainer.addEventListener('mousemove', (e) => {
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - diagramCanvasContainer.offsetLeft;
      const y = e.pageY - diagramCanvasContainer.offsetTop;
      const walkX = (x - startX) * 1.5;
      const walkY = (y - startY) * 1.5;
      diagramCanvasContainer.scrollLeft = scrollLeft - walkX;
      diagramCanvasContainer.scrollTop = scrollTop - walkY;
    });
  }

  // Initial code template
  if (diagramCodeInput) {
    diagramCodeInput.value = DIAGRAM_PRESETS.mermaid_sequence;
  }

  engineBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      engineBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      const engine = btn.getAttribute('data-engine') || 'mermaid';
      state.diagramEngine = engine;

      if (engine === 'mermaid' && diagramCodeInput) {
        diagramCodeInput.value = DIAGRAM_PRESETS.mermaid_sequence;
      } else if (engine === 'plantuml' && diagramCodeInput) {
        diagramCodeInput.value = DIAGRAM_PRESETS.plantuml_c4;
      } else if (engine === 'drawio' && diagramCodeInput) {
        diagramCodeInput.value = `<mxfile host="Antigravity Studio" modified="2026-10-02" agent="WebStudio" version="21.0.0">
  <diagram id="sample" name="System Overview">
    <mxGraphModel dx="1200" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1200" pageHeight="800" background="#FFFFFF">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="hub" value="Antigravity Hub" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#0051E2;strokeColor=#003DB3;fontColor=#FFFFFF;fontStyle=1;fontSize=14;" vertex="1" parent="1">
          <mxGeometry x="450" y="250" width="180" height="70" as="geometry"/>
        </mxCell>
        <mxCell id="spoke1" value="PDF Engine" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#F8F9FA;strokeColor=#1E293B;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="180" y="255" width="150" height="60" as="geometry"/>
        </mxCell>
        <mxCell id="spoke2" value="PPTX 16:9 Generator" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#F8F9FA;strokeColor=#1E293B;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="750" y="255" width="170" height="60" as="geometry"/>
        </mxCell>
        <mxCell id="e1" edge="1" parent="1" source="spoke1" target="hub" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#0051E2;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="e2" edge="1" parent="1" source="hub" target="spoke2" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=2;strokeColor=#0051E2;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>`;
      }
    });
  });

  if (btnSampleSequence && diagramCodeInput) {
    btnSampleSequence.addEventListener('click', () => {
      engineBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-engine') === 'mermaid'));
      state.diagramEngine = 'mermaid';
      diagramCodeInput.value = DIAGRAM_PRESETS.mermaid_sequence;
    });
  }

  if (btnSampleC4 && diagramCodeInput) {
    btnSampleC4.addEventListener('click', () => {
      engineBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-engine') === 'plantuml'));
      state.diagramEngine = 'plantuml';
      diagramCodeInput.value = DIAGRAM_PRESETS.plantuml_c4;
    });
  }

  if (btnSampleErd && diagramCodeInput) {
    btnSampleErd.addEventListener('click', () => {
      engineBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-engine') === 'mermaid'));
      state.diagramEngine = 'mermaid';
      diagramCodeInput.value = DIAGRAM_PRESETS.mermaid_erd;
    });
  }

  if (btnRenderDiagram) {
    btnRenderDiagram.addEventListener('click', async () => {
      const code = diagramCodeInput ? diagramCodeInput.value.trim() : '';
      if (!code) {
        showToast('Vui lòng nhập mã nguồn sơ đồ!', 'error');
        return;
      }

      btnRenderDiagram.disabled = true;
      btnRenderDiagram.innerHTML = '<span>⏳</span><span>ĐANG RENDER SƠ ĐỒ...</span>';

      try {
        const res = await fetch('/api/render-diagram', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            engine: state.diagramEngine,
            code: code,
            format: 'png',
          }),
        });

        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error || `Render sơ đồ thất bại (${res.status})`);
        }

        const data = await res.json();
        const imgFile = data.image_file || data.filename || data.output_file || (data.image_path ? data.image_path.split(/[\\/]/).pop() : '');
        const imageUrl = `/api/preview/${imgFile}?t=${Date.now()}`;
        const downloadUrl = data.download_url || `/api/download/${imgFile}`;

        if (diagramPreviewImg) {
          diagramPreviewImg.src = imageUrl;
        }
        if (diagramViewport) diagramViewport.style.display = 'flex';
        if (diagramEmptyState) diagramEmptyState.style.display = 'none';
        if (diagramDownloadActions) diagramDownloadActions.style.display = 'flex';
        if (diagramZoomToolbar) diagramZoomToolbar.style.display = 'flex';
        if (diagramHintBar) diagramHintBar.style.display = 'flex';
        if (btnDownloadDiagram) btnDownloadDiagram.setAttribute('href', downloadUrl);

        // Auto default to 100% Gốc for high-fidelity uncompressed typography
        setViewMode('100');

        showToast('Render sơ đồ kỹ thuật thành công!', 'success');
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error', 6000);
      } finally {
        btnRenderDiagram.disabled = false;
        btnRenderDiagram.innerHTML = '<span>🎨</span><span>RENDER SƠ ĐỒ ĐỘ PHÂN GIẢI CAO</span>';
      }
    });
  }
});
