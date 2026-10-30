/**
 * src/modules/presentation_studio.js ? Presentation Studio Controller (16:9 PPTX Synthesis & 1080p Gallery)
 */

export function initPresentationStudio({ state, showToast, refreshLucideIcons, uploadFileToServer, formatBytes, loadHistory }) {
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

}
