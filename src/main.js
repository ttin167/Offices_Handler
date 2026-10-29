import {
  createIcons,
  Sparkles,
  Presentation,
  Table,
  RefreshCw,
  Network,
  Sigma,
  History,
  Moon,
  Sun,
  Trash2,
  X,
  FolderOpen,
  ShieldCheck,
  Maximize2,
  UploadCloud,
  Shield,
  Type,
  FileCode,
  ChevronDown,
  Play,
  Images,
  Download,
  Code,
  LayoutGrid,
  Image,
  Layers,
  ChevronLeft,
  ChevronRight,
  Zap,
  ScanSearch,
  CheckCircle2,
  AlertCircle,
  Info,
  FileText,
  FileCheck,
  ZoomOut,
  ZoomIn,
  Maximize,
  ExternalLink,
  User,
  Database,
  HelpCircle,
  Cloud,
  Box,
  Calculator,
  Search,
  Copy,
  Check,
  BookOpen,
  Briefcase,
  Lightbulb,
  Printer,
  Package,
  Eye,
  Sliders,
  SlidersHorizontal,
  GraduationCap,
  Award,
  Users,
  Building,
  UserCheck,
  Pin,
  Plus,
  Book,
} from 'lucide';

import { DIAGRAM_PRESETS } from './modules/constants.js';
import { initHistoryDrawer } from './modules/history_drawer.js';
import { initMathStudio } from './modules/math_studio.js';
import { initFormulaPlaybook } from './modules/formula_playbook.js';

// Optional local-only module (excluded in GitHub/Docker builds via .gitignore)
const _optionalYearbookModules = import.meta.glob('./modules/yearbook_studio.js', { eager: true });
const initYearbookStudio = _optionalYearbookModules['./modules/yearbook_studio.js']?.initYearbookStudio || null;

const studioIcons = {
  Sparkles,
  Presentation,
  Table,
  RefreshCw,
  Network,
  Sigma,
  History,
  Moon,
  Sun,
  Trash2,
  X,
  FolderOpen,
  ShieldCheck,
  Maximize2,
  UploadCloud,
  Shield,
  Type,
  FileCode,
  ChevronDown,
  Play,
  Images,
  Download,
  Code,
  LayoutGrid,
  Image,
  Layers,
  ChevronLeft,
  ChevronRight,
  Zap,
  ScanSearch,
  CheckCircle2,
  AlertCircle,
  Info,
  FileText,
  FileCheck,
  ZoomOut,
  ZoomIn,
  Maximize,
  ExternalLink,
  User,
  Database,
  HelpCircle,
  Cloud,
  Box,
  Calculator,
  Search,
  Copy,
  Check,
  BookOpen,
  Briefcase,
  Lightbulb,
  Printer,
  Package,
  Eye,
  Sliders,
  SlidersHorizontal,
  GraduationCap,
  Award,
  Users,
  Building,
  UserCheck,
  Pin,
  Plus,
  Book,
};

/**
 * ANTIGRAVITY OFFICE STUDIO — CLIENT LOGIC CONTROLLER
 * Full-Cycle PDF Hub, Document Matrix, Presentation Studio & Diagram Engine
 */

function refreshLucideIcons() {
  try {
    createIcons({ icons: studioIcons });
  } catch (err) {
    console.warn('Lucide icon initialization warning:', err);
  }
}

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

  // DIAGRAM_PRESETS imported from ./modules/constants.js

  // =========================================================================
  // 1. THEME CONTROLLER
  // =========================================================================
  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('antigravity_theme', theme);
    const themeBtn = document.getElementById('btn-toggle-theme');
    if (themeBtn) {
      themeBtn.innerHTML = theme === 'dark' ? '<i data-lucide="sun"></i>' : '<i data-lucide="moon"></i>';
      refreshLucideIcons();
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

    let iconName = 'info';
    if (type === 'success') iconName = 'check-circle-2';
    if (type === 'error') iconName = 'alert-circle';

    toast.innerHTML = `<i data-lucide="${iconName}"></i><span>${message}</span>`;
    container.appendChild(toast);
    refreshLucideIcons();

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
        const data = await res.json().catch(() => ({}));
        if (data.version) {
          const brandSubtitle = document.querySelector('.brand-subtitle');
          if (brandSubtitle) {
            brandSubtitle.textContent = `Office Studio v${data.version}`;
          }
        }
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
    math: document.getElementById('panel-math'),
    yearbook: document.getElementById('panel-yearbook'),
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

      if (targetTab === 'yearbook') {
        window.dispatchEvent(new CustomEvent('yearbook:activated'));
      }
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

  // Deck Scope selector radio buttons (Executive Deck Sizing)
  const scopeCards = document.querySelectorAll('.scope-card');
  scopeCards.forEach((card) => {
    card.addEventListener('click', () => {
      scopeCards.forEach((c) => c.classList.remove('active'));
      card.classList.add('active');
      const radio = card.querySelector('input[type="radio"]');
      if (radio) radio.checked = true;
      state.pptxDeckScope = card.getAttribute('data-scope-val') || 'auto';
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
        notesHtml = `<div class="slide-card-notes"><i data-lucide="file-text"></i> Ghi chú: ${slide.notes}</div>`;
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
    refreshLucideIcons();
  }

  // Slide Previews & Core v2 State
  state.currentPptxPath = null;
  state.currentPptxFile = null;
  state.pptxSlidePreviews = [];
  state.currentPreviewIndex = 0;
  state.activeDeckView = 'cards';
  state.pptxDeckScope = 'auto';
  state.totalSlideCount = 0;
  state.previewBatchLoaded = 0;

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

  // Export Slide Previews Action (1080p) with Batch Progressive Loading
  const btnExportPreviews = document.getElementById('btn-export-previews');
  const galleryCountBadge = document.getElementById('gallery-count-badge');
  const loadMoreContainer = document.getElementById('gallery-load-more-container');
  const btnLoadMorePreviews = document.getElementById('btn-load-more-previews');
  const loadMoreText = document.getElementById('load-more-text');

  async function exportSlidePreviews(batchRange = null) {
    if (!state.currentPptxPath && !state.currentPptxFile) {
      showToast('Chưa có tệp PowerPoint để xuất ảnh xem trước!', 'error');
      return;
    }

    const total = state.totalSlideCount || 0;
    const isBatchMode = total > 16;
    let slidesArg = 'all';

    if (batchRange) {
      slidesArg = batchRange;
    } else if (isBatchMode) {
      slidesArg = '1-16';
    }

    if (btnExportPreviews && !batchRange) {
      btnExportPreviews.disabled = true;
      btnExportPreviews.innerHTML = '<span>⏳ Đang render 1080p...</span>';
    }

    try {
      showToast(`Đang xuất ảnh xem trước (phạm vi: ${slidesArg}) chuẩn 1080p...`, 'info');
      const res = await fetch('/api/pptx/preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pptx_path: state.currentPptxPath || state.currentPptxFile,
          slides: slidesArg,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Xuất ảnh thất bại (${res.status})`);
      }

      const data = await res.json();
      const newPreviews = data.previews || [];

      if (batchRange) {
        state.pptxSlidePreviews = state.pptxSlidePreviews.concat(newPreviews);
      } else {
        state.pptxSlidePreviews = newPreviews;
      }

      state.previewBatchLoaded = state.pptxSlidePreviews.length;
      if (galleryCountBadge) galleryCountBadge.textContent = state.pptxSlidePreviews.length;
      renderSlideGallery(state.pptxSlidePreviews);
      setDeckView('gallery');

      if (loadMoreContainer) {
        if (total > state.previewBatchLoaded) {
          loadMoreContainer.style.display = 'block';
          if (loadMoreText) {
            const nextBatchEnd = Math.min(state.previewBatchLoaded + 16, total);
            loadMoreText.textContent = `Tải tiếp slides ${state.previewBatchLoaded + 1}–${nextBatchEnd} (còn ${total - state.previewBatchLoaded} slides)...`;
          }
        } else {
          loadMoreContainer.style.display = 'none';
        }
      }

      showToast(`Đã xuất thành công ${state.pptxSlidePreviews.length} ảnh slide xem trước 1080p!`, 'success');
    } catch (err) {
      showToast(`Lỗi xuất ảnh xem trước: ${err.message}`, 'error');
    } finally {
      if (btnExportPreviews && !batchRange) {
        btnExportPreviews.disabled = false;
        btnExportPreviews.innerHTML = '<span><i data-lucide="images"></i> Xuất Thư Viện Ảnh (1080p)</span>';
        refreshLucideIcons();
      }
    }
  }

  if (btnExportPreviews) {
    btnExportPreviews.addEventListener('click', () => exportSlidePreviews());
  }

  if (btnLoadMorePreviews) {
    btnLoadMorePreviews.addEventListener('click', async () => {
      const start = state.previewBatchLoaded + 1;
      const end = Math.min(state.previewBatchLoaded + 16, state.totalSlideCount || (start + 15));
      btnLoadMorePreviews.disabled = true;
      if (loadMoreText) loadMoreText.textContent = '⏳ Đang tải thêm slides...';
      await exportSlidePreviews(`${start}-${end}`);
      btnLoadMorePreviews.disabled = false;
    });
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
            <span class="overlay-icon"><i data-lucide="zoom-in"></i> Phóng to</span>
          </div>
        </div>
      `;
      card.addEventListener('click', () => openSlideModal(idx));
      slideGalleryGrid.appendChild(card);
    });
    refreshLucideIcons();
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
        btnBuildFromSpec.innerHTML = '<span><i data-lucide="play"></i> Render PPTX Từ Spec JSON</span>';
        refreshLucideIcons();
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
          deck_scope: state.pptxDeckScope || 'auto',
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
        state.totalSlideCount = (data.spec && data.spec.slides) ? data.spec.slides.length : (data.slides_count || 0);
        state.previewBatchLoaded = 0;
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
    } catch (_) { }
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
      if (zoomLevelBadge) zoomLevelBadge.textContent = 'Khớp';
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
      } else if (engine === 'canvas' && diagramCodeInput) {
        diagramCodeInput.value = DIAGRAM_PRESETS.canvas_flow;
      } else if (engine === 'drawio' && diagramCodeInput) {
        diagramCodeInput.value = DIAGRAM_PRESETS.drawio_overview;
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

  const btnSampleCanvas = document.getElementById('btn-sample-canvas');
  if (btnSampleCanvas && diagramCodeInput) {
    btnSampleCanvas.addEventListener('click', () => {
      engineBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-engine') === 'canvas'));
      state.diagramEngine = 'canvas';
      diagramCodeInput.value = DIAGRAM_PRESETS.canvas_flow;
    });
  }

  const diagramPresetSelect = document.getElementById('diagram-preset-select');
  if (diagramPresetSelect && diagramCodeInput) {
    diagramPresetSelect.addEventListener('change', (e) => {
      const presetKey = e.target.value;
      if (!presetKey || !DIAGRAM_PRESETS[presetKey]) return;

      let targetEngine = 'mermaid';
      if (presetKey.startsWith('plantuml_')) {
        targetEngine = 'plantuml';
      } else if (presetKey.startsWith('drawio_')) {
        targetEngine = 'drawio';
      } else if (presetKey.startsWith('canvas_')) {
        targetEngine = 'canvas';
      }

      state.diagramEngine = targetEngine;
      engineBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-engine') === targetEngine));
      diagramCodeInput.value = DIAGRAM_PRESETS[presetKey];
      showToast(`Đã tải mẫu: ${presetKey.replace('_', ' ').toUpperCase()}`, 'info');
    });
  }

  function insertSnippetIntoEditor(snippet) {
    if (!diagramCodeInput) return;
    const start = diagramCodeInput.selectionStart;
    const end = diagramCodeInput.selectionEnd;
    const text = diagramCodeInput.value;
    if (typeof start === 'number' && typeof end === 'number') {
      diagramCodeInput.value = text.substring(0, start) + snippet + text.substring(end);
      diagramCodeInput.selectionStart = diagramCodeInput.selectionEnd = start + snippet.length;
    } else {
      diagramCodeInput.value += snippet;
    }
    diagramCodeInput.focus();
  }

  const toolstripBtns = document.querySelectorAll('.toolstrip-btn');
  toolstripBtns.forEach((tBtn) => {
    tBtn.addEventListener('click', () => {
      const sym = tBtn.getAttribute('data-symbol');
      const randId = Math.floor(1000 + Math.random() * 9000);
      let snippet = '';

      if (state.diagramEngine === 'mermaid') {
        const map = {
          actor: `\n  actor User_${randId} as Người Dùng\n`,
          database: `\n  DB_${randId}[(Cơ Sở Dữ Liệu)]\n`,
          decision: `\n  Cond_${randId}{"Điều Kiện Kiểm Tra?"}\n`,
          cloud: `\n  subgraph Cloud_${randId} [Hệ Thống Đám Mây]\n    Svc_${randId}[Dịch Vụ Ngoài]\n  end\n`,
          terminator: `\n  Term_${randId}([Bắt Đầu / Kết Thúc])\n`,
          queue: `\n  Queue_${randId}[("(Hàng Đợi Queue)")]\n`,
          component: `\n  Comp_${randId}[Mô-đun Hệ Thống]\n`,
          document: `\n  Doc_${randId}[/"Tài Liệu Báo Cáo"/]\n`,
        };
        snippet = map[sym] || `\n  Node_${randId}[Khối Mới]\n`;
      } else if (state.diagramEngine === 'plantuml') {
        const map = {
          actor: `\nactor "Người Dùng" as User_${randId}\n`,
          database: `\ndatabase "Cơ Sở Dữ Liệu" as DB_${randId}\n`,
          decision: `\nif (Điều kiện kiểm tra?) then (Đúng)\n  :Xử lý tiếp;\nelse (Sai)\n  :Báo lỗi;\nendif\n`,
          cloud: `\ncloud "Dịch Vụ Cloud" as Cloud_${randId}\n`,
          terminator: `\nstart\n:Bắt đầu tác vụ;\nstop\n`,
          queue: `\nqueue "Hàng Đợi Tin Nhắn" as Queue_${randId}\n`,
          component: `\n[Mô-đun Thành Phần] as Comp_${randId}\n`,
          document: `\nfile "Tệp Tài Liệu" as Doc_${randId}\n`,
        };
        snippet = map[sym] || `\n:Tác Vụ Mới;\n`;
      } else if (state.diagramEngine === 'canvas') {
        const map = {
          actor: `,\n    {\n      "id": "actor_${randId}",\n      "title": "Actor User",\n      "type": "actor",\n      "x": 100,\n      "y": 200,\n      "width": 70,\n      "height": 100,\n      "icon": "user"\n    }`,
          database: `,\n    {\n      "id": "db_${randId}",\n      "title": "Database",\n      "type": "database",\n      "x": 300,\n      "y": 200,\n      "width": 140,\n      "height": 90,\n      "icon": "database"\n    }`,
          decision: `,\n    {\n      "id": "decision_${randId}",\n      "title": "Check Valid?",\n      "type": "decision",\n      "x": 300,\n      "y": 200,\n      "width": 140,\n      "height": 90\n    }`,
          cloud: `,\n    {\n      "id": "cloud_${randId}",\n      "title": "External Cloud",\n      "type": "cloud",\n      "x": 300,\n      "y": 200,\n      "width": 180,\n      "height": 90,\n      "icon": "cloud"\n    }`,
          terminator: `,\n    {\n      "id": "term_${randId}",\n      "title": "Start/End",\n      "type": "terminator",\n      "x": 300,\n      "y": 200,\n      "width": 140,\n      "height": 60\n    }`,
          queue: `,\n    {\n      "id": "queue_${randId}",\n      "title": "Message Queue",\n      "type": "queue",\n      "x": 300,\n      "y": 200,\n      "width": 160,\n      "height": 70\n    }`,
          component: `,\n    {\n      "id": "comp_${randId}",\n      "title": "Component",\n      "type": "component",\n      "x": 300,\n      "y": 200,\n      "width": 180,\n      "height": 80,\n      "icon": "server"\n    }`,
          document: `,\n    {\n      "id": "doc_${randId}",\n      "title": "Document",\n      "type": "document",\n      "x": 300,\n      "y": 200,\n      "width": 150,\n      "height": 80\n    }`,
        };
        snippet = map[sym] || `,\n    {\n      "id": "node_${randId}",\n      "title": "New Node",\n      "type": "standard",\n      "x": 300,\n      "y": 200,\n      "width": 180,\n      "height": 80\n    }`;
      } else if (state.diagramEngine === 'drawio') {
        const map = {
          actor: `\n        <mxCell id="actor_${randId}" value="User Actor" style="shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;outlineConnect=0;fillColor=#FFFFFF;strokeColor=#1E293B;strokeWidth=1.8;" vertex="1" parent="1">\n          <mxGeometry x="100" y="200" width="60" height="90" as="geometry"/>\n        </mxCell>`,
          database: `\n        <mxCell id="db_${randId}" value="Database" style="shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;backgroundOutline=1;size=15;fillColor=#F8FAFC;strokeColor=#2563EB;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="120" height="90" as="geometry"/>\n        </mxCell>`,
          decision: `\n        <mxCell id="decision_${randId}" value="Condition?" style="rhombus;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#D97706;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="130" height="85" as="geometry"/>\n        </mxCell>`,
          cloud: `\n        <mxCell id="cloud_${randId}" value="Cloud API" style="ellipse;shape=cloud;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#0284C7;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="170" height="90" as="geometry"/>\n        </mxCell>`,
          terminator: `\n        <mxCell id="term_${randId}" value="Start/Stop" style="rounded=1;arcSize=50;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#10B981;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="140" height="60" as="geometry"/>\n        </mxCell>`,
          queue: `\n        <mxCell id="queue_${randId}" value="Task Queue" style="shape=tape;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#7C3AED;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="150" height="70" as="geometry"/>\n        </mxCell>`,
          component: `\n        <mxCell id="comp_${randId}" value="Component" style="shape=module;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#2563EB;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="160" height="80" as="geometry"/>\n        </mxCell>`,
          document: `\n        <mxCell id="doc_${randId}" value="Report Doc" style="shape=document;whiteSpace=wrap;html=1;boundedLbl=1;fillColor=#F8FAFC;strokeColor=#475569;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="140" height="80" as="geometry"/>\n        </mxCell>`,
        };
        snippet = map[sym] || `\n        <mxCell id="node_${randId}" value="New Node" style="rounded=1;whiteSpace=wrap;html=1;" vertex="1" parent="1">\n          <mxGeometry x="300" y="200" width="180" height="70" as="geometry"/>\n        </mxCell>`;
      }

      insertSnippetIntoEditor(snippet);
      showToast(`Đã chèn ký hiệu ${sym.toUpperCase()}!`, 'info');
    });
  });

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
        refreshLucideIcons();

        // Auto default to fit mode for instant full overview
        setViewMode('fit');

        showToast('Render sơ đồ kỹ thuật thành công!', 'success');
        if (typeof loadHistory === 'function') loadHistory();
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error', 6000);
      } finally {
        btnRenderDiagram.disabled = false;
        btnRenderDiagram.innerHTML = '<span class="btn-icon"><i data-lucide="network"></i></span><span>RENDER SƠ ĐỒ ĐỘ PHÂN GIẢI CAO</span>';
        refreshLucideIcons();
      }
    });
  }

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

  // =========================================================================
  // 10. HISTORY DRAWER & RECENT FILES CONTROLLER (Modularized)
  // =========================================================================
  const { loadHistory, openHistoryDrawer, closeHistoryDrawer } = initHistoryDrawer({
    formatBytes,
    refreshLucideIcons,
    showToast,
  });

  // =========================================================================
  // 11. TAB 5: MATH & EQUATION STUDIO CONTROLLER (Modularized)
  // =========================================================================
  initMathStudio({ showToast, refreshLucideIcons, loadHistory });

  // =========================================================================
  // 12. TAB 2 & 3: DOCX QUICK PATCH & ATOMIC CONTROLLERS
  // =========================================================================
  initDocxQuickPatch();

  // =========================================================================
  // 13. SPREADSHEET STUDIO: EXCEL FORMULA PLAYBOOK CONTROLLER
  // =========================================================================
  initFormulaPlaybook({ showToast, refreshLucideIcons });

  // =========================================================================
  // 14. TAB 6: DIGITAL YEARBOOK & PRE-PRESS PHOTOBOOK STUDIO CONTROLLER
  // =========================================================================
  if (typeof initYearbookStudio === 'function') {
    initYearbookStudio({ showToast, refreshLucideIcons, loadHistory });
  } else {
    const ybTabBtn = document.querySelector('[data-tab="yearbook"]');
    const ybPanel = document.getElementById('panel-yearbook');
    if (ybTabBtn) ybTabBtn.style.display = 'none';
    if (ybPanel) ybPanel.style.display = 'none';
  }

  // Initial Lucide icon generation
  refreshLucideIcons();
});
