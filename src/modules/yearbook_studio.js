/**
 * src/modules/yearbook_studio.js — Digital Yearbook & Pre-Press Photobook Studio Controller
 *
 * Implements:
 * - 3D Realistic Double-Page Spread View with Spine Gutter Shadow.
 * - Interactive HTML5 Drag & Drop (Tray-to-Slot & Slot-to-Slot Swapping).
 * - Pre-Press Guides Overlay (Bleed 3mm, Trim line, Safe zone).
 * - Multi-photo Upload & High-Density Auto-Layout (50–100+ photos).
 * - 9 Editorial Layout Templates (Cover, Editorial, Duo, Mag-4, Polaroid, Mosaic-6, Classroom-8, Megagrid-12, Panoramic).
 * - Designer Handoff 5-in-1 Bundle Export & PDF Generation.
 */

// Gradient SVG Placeholders for Mock Photos (High Aesthetic)
const MOCK_PLACEHOLDERS = {
  'mock:grad_cap': {
    title: 'Biện Trung Tín - Lễ Tốt Nghiệp',
    sub: 'Khoa Công Nghệ Thông Tin 2026',
    bg: 'linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #1E40AF 100%)',
    icon: 'graduation-cap',
    tag: 'Chân Dung Cử Nhân'
  },
  'mock:ceremony': {
    title: 'Khoảnh Khắc Nhận Bằng',
    sub: 'Hội Trường Trọng Thể',
    bg: 'linear-gradient(135deg, #1E1B4B 0%, #312E81 50%, #4338CA 100%)',
    icon: 'award',
    tag: 'Nghi Lễ Trao Bằng'
  },
  'mock:friends': {
    title: 'Tình Bạn Đại Học',
    sub: 'Nhóm Bạn Thân Cùng Lớp',
    bg: 'linear-gradient(135deg, #064E3B 0%, #065F46 50%, #059669 100%)',
    icon: 'users',
    tag: 'Bạn Thân Kỷ Niệm'
  },
  'mock:campus': {
    title: 'Sân Trường Thân Quen',
    sub: 'Giảng Đường & Hàng Cây',
    bg: 'linear-gradient(135deg, #1C1917 0%, #292524 50%, #44403C 100%)',
    icon: 'building',
    tag: 'Toàn Cảnh Giảng Đường'
  },
  'mock:library': {
    title: 'Ký Ức Thư Viện',
    sub: 'Những Đêm Ôn Thi Đồ Án',
    bg: 'linear-gradient(135deg, #451A03 0%, #78350F 50%, #92400E 100%)',
    icon: 'book-open',
    tag: 'Góc Học Tập Đồ Án'
  },
  'mock:portrait': {
    title: 'Áo Cử Nhân Rạng Rỡ',
    sub: 'Nụ Cười Ngày Vinh Quy',
    bg: 'linear-gradient(135deg, #4C0519 0%, #881337 50%, #9F1239 100%)',
    icon: 'user-check',
    tag: 'Chân Dung Ngoại Cảnh'
  },
  'mock:diploma': {
    title: 'Tấm Bằng Đại Học Kỹ Sư',
    sub: 'Minh Chứng 4 Năm Nỗ Lực',
    bg: 'linear-gradient(135deg, #3B0764 0%, #581C87 50%, #7E22CE 100%)',
    icon: 'file-check',
    tag: 'Bằng Tốt Nghiệp'
  },
  'mock:sunset': {
    title: 'Hoàng Hôn Giảng Đường',
    sub: 'Khép Lại Thanh Xuân',
    bg: 'linear-gradient(135deg, #431407 0%, #7C2D12 50%, #C2410C 100%)',
    icon: 'sun',
    tag: 'Hoàng Hôn Tốt Nghiệp'
  }
};

export function initYearbookStudio({ showToast, refreshLucideIcons, loadHistory }) {
  // DOM Elements
  const panelYearbook = document.getElementById('panel-yearbook');
  if (!panelYearbook) return;

  const albumTitleInput = document.getElementById('yb-input-album-title') || document.getElementById('yb-input-title');
  const albumSubInput = document.getElementById('yb-input-album-sub') || document.getElementById('yb-input-author');
  const selectSize = document.getElementById('yb-select-size') || document.getElementById('yb-select-preset');
  const sliderBleed = document.getElementById('yb-slider-bleed');
  const bleedVal = document.getElementById('yb-bleed-val');

  const photoDropZone = document.getElementById('yearbook-drop-zone');
  const photoFileInput = document.getElementById('yearbook-file-input');
  const btnBrowsePhotos = document.getElementById('btn-browse-yearbook-photos');
  const trayCountBadge = document.getElementById('yb-tray-count') || document.getElementById('yearbook-tray-count');
  const trayPhotosContainer = document.getElementById('yb-tray-photos') || document.getElementById('yearbook-photo-tray');
  const btnClearTray = document.getElementById('btn-clear-tray');
  const btnAutoLayout = document.getElementById('btn-yearbook-autolayout');

  const spreadIndicator = document.getElementById('yb-spread-indicator');
  const btnPrevSpread = document.getElementById('btn-yb-prev-spread');
  const btnNextSpread = document.getElementById('btn-yb-next-spread');
  const spreadSlider = document.getElementById('yb-spread-slider');
  const togglePrepressGuides = document.getElementById('yb-toggle-prepress-guides');
  const bookContainer = document.getElementById('yearbook-book-container');

  const selectLayout = document.getElementById('yb-select-layout');
  const inputCaption = document.getElementById('yb-input-spread-caption');
  const inputNote = document.getElementById('yb-input-spread-note');
  const quoteGroup = document.getElementById('yb-quote-group');
  const inputQuote = document.getElementById('yb-input-spread-quote');

  const btnAddSpread = document.getElementById('btn-yb-add-spread');
  const btnDuplicateSpread = document.getElementById('btn-yb-duplicate-spread');
  const btnDeleteSpread = document.getElementById('btn-yb-delete-spread');

  const btnExportSpreadPdf = document.getElementById('btn-yb-export-spread-pdf');
  const btnExportPrintPdf = document.getElementById('btn-yb-export-print-pdf');
  const btnExportHandoffZip = document.getElementById('btn-yb-export-handoff-zip');
  const exportStatusBox = document.getElementById('yb-export-status-box');
  const statusSpinner = document.getElementById('yb-status-spinner');
  const statusText = document.getElementById('yb-status-text');
  const btnDownloadResult = document.getElementById('yb-btn-download-result');

  // Application State
  const state = {
    spec: null,
    currentSpreadIndex: 0,
    guidesVisible: false,
    photoTray: [],
    draggedItem: null, // { type: 'tray-photo' | 'slot-photo', data: any }
  };

  // =========================================================================
  // 1. HELPER FUNCTIONS
  // =========================================================================

  function normalizeSpec(spec) {
    if (!spec) return null;
    const rawSpreads = spec.spreads || spec.pages || [];
    const normalizedSpreads = rawSpreads.map((sp, idx) => {
      const layout = sp.layout || sp.type || 'magazine_grid_4';
      const rawSlots = sp.slots || [];
      const slots = rawSlots.map((sl, sIdx) => {
        let photoVal = sl.photo || sl.photo_path || null;
        if (photoVal && typeof photoVal === 'string' && photoVal.startsWith('mock://')) {
          const theme = (sl.mock_theme || photoVal.replace('mock://', '').split('.')[0]).toLowerCase();
          if (theme.includes('gown') || theme.includes('hero')) photoVal = 'mock:grad_cap';
          else if (theme.includes('ceremony') || theme.includes('diploma')) photoVal = 'mock:ceremony';
          else if (theme.includes('friend') || theme.includes('student')) photoVal = 'mock:friends';
          else if (theme.includes('campus') || theme.includes('gate')) photoVal = 'mock:campus';
          else if (theme.includes('library')) photoVal = 'mock:library';
          else if (theme.includes('portrait')) photoVal = 'mock:portrait';
          else if (theme.includes('backstage') || theme.includes('sunset')) photoVal = 'mock:sunset';
          else photoVal = 'mock:grad_cap';
        }
        return {
          id: sl.id || `slot_${sIdx + 1}`,
          photo: photoVal,
          photo_path: photoVal,
          caption: sl.caption || '',
          note: sl.note || sl.designer_note || '',
          ratio: sl.ratio || getSlotDefaultRatio(layout, sIdx)
        };
      });

      return {
        ...sp,
        spread_index: sp.spread_index || idx + 1,
        layout: layout,
        type: layout,
        caption: sp.caption || '',
        quote: sp.quote || '',
        designer_note: sp.designer_note || sp.note || '',
        slots: slots
      };
    });

    return {
      ...spec,
      title: spec.title || 'KỶ YẾU TỐT NGHIỆP 2026',
      subtitle: spec.subtitle || 'BIỆN TRUNG TÍN',
      page_size: spec.page_size || spec.preset || 'photobook_square_25x25',
      spreads: normalizedSpreads,
      pages: normalizedSpreads
    };
  }

  function getPhotoDisplayData(photoRef) {
    if (!photoRef) return null;
    let ref = photoRef;
    if (typeof photoRef === 'object') {
      ref = photoRef.photo || photoRef.photo_path || photoRef.url || photoRef.filepath;
    }
    if (!ref) return null;

    if (typeof ref === 'string' && (ref.startsWith('mock:') || ref.startsWith('mock://'))) {
      let mockKey = ref;
      if (ref.startsWith('mock://')) {
        const theme = ref.replace('mock://', '').split('.')[0].toLowerCase();
        if (theme.includes('gown') || theme.includes('hero')) mockKey = 'mock:grad_cap';
        else if (theme.includes('ceremony') || theme.includes('diploma')) mockKey = 'mock:ceremony';
        else if (theme.includes('friend') || theme.includes('student')) mockKey = 'mock:friends';
        else if (theme.includes('campus') || theme.includes('gate')) mockKey = 'mock:campus';
        else if (theme.includes('library')) mockKey = 'mock:library';
        else if (theme.includes('portrait')) mockKey = 'mock:portrait';
        else if (theme.includes('backstage') || theme.includes('sunset')) mockKey = 'mock:sunset';
        else mockKey = 'mock:grad_cap';
      }

      const mockMeta = MOCK_PLACEHOLDERS[mockKey] || {
        title: 'Kỷ Yếu Tốt Nghiệp',
        sub: '2026',
        bg: 'linear-gradient(135deg, #1E293B, #334155)',
        icon: 'image',
        tag: 'Mock Photo'
      };
      return {
        isMock: true,
        mockId: mockKey,
        title: mockMeta.title,
        sub: mockMeta.sub,
        bg: mockMeta.bg,
        icon: mockMeta.icon,
        tag: mockMeta.tag
      };
    }
    // Real uploaded photo
    const url = typeof ref === 'string' ? ref : (ref.url || ref.filepath);
    return {
      isMock: false,
      url: url,
      title: (typeof photoRef === 'object' && photoRef.filename) ? photoRef.filename : (typeof ref === 'string' ? ref.split('/').pop() : 'Ảnh Kỷ Yếu'),
      sub: (typeof photoRef === 'object' && photoRef.aspect) ? photoRef.aspect : '3:2'
    };
  }

  function getSlotCountForLayout(layout) {
    switch (layout) {
      case 'cover_hero':
      case 'editorial_quote':
      case 'panoramic_spread':
        return 1;
      case 'symmetric_duo':
        return 2;
      case 'polaroid_scrapbook':
        return 3;
      case 'magazine_grid_4':
        return 4;
      case 'mosaic_gallery_6':
        return 6;
      case 'classroom_grid_8':
        return 8;
      case 'megagrid_12':
        return 12;
      default:
        return 4;
    }
  }

  function createSlotsForLayout(layout, existingSlots = []) {
    const targetCount = getSlotCountForLayout(layout);
    const newSlots = [];
    for (let i = 0; i < targetCount; i++) {
      const existing = existingSlots[i] || {};
      newSlots.push({
        id: `slot_${i + 1}`,
        photo: existing.photo || null,
        caption: existing.caption || '',
        note: existing.note || '',
        ratio: getSlotDefaultRatio(layout, i)
      });
    }
    return newSlots;
  }

  function getSlotDefaultRatio(layout, slotIndex) {
    if (layout === 'cover_hero' || layout === 'editorial_quote') return '3:4';
    if (layout === 'symmetric_duo') return '2:3';
    if (layout === 'polaroid_scrapbook') return '1:1';
    if (layout === 'panoramic_spread') return '16:9';
    if (layout === 'mosaic_gallery_6') {
      return slotIndex === 0 || slotIndex === 3 ? '16:9' : '1:1';
    }
    if (layout === 'classroom_grid_8') return '1:1';
    if (layout === 'megagrid_12') return '1:1';
    return slotIndex === 0 ? '3:2' : '1:1';
  }

  // =========================================================================
  // 2. RENDERING: 3D FLIPBOOK DOUBLE-PAGE SPREAD
  // =========================================================================

  function renderCurrentSpread() {
    if (!state.spec || !state.spec.spreads || state.spec.spreads.length === 0) {
      bookContainer.innerHTML = `
        <div class="empty-state-deck">
          <div class="empty-illustration"><i data-lucide="book-open"></i></div>
          <div class="empty-title">Chưa Có Dữ Liệu Kỷ Yếu</div>
          <p class="empty-desc">Bấm "Dàn Trang Tự Động" hoặc tải ảnh lên để bắt đầu thiết kế photobook.</p>
        </div>
      `;
      refreshLucideIcons();
      return;
    }

    const spreads = state.spec.spreads;
    if (state.currentSpreadIndex >= spreads.length) {
      state.currentSpreadIndex = spreads.length - 1;
    }
    if (state.currentSpreadIndex < 0) {
      state.currentSpreadIndex = 0;
    }

    const currentSpread = spreads[state.currentSpreadIndex];
    const leftPageNum = (state.currentSpreadIndex * 2) + 1;
    const rightPageNum = (state.currentSpreadIndex * 2) + 2;

    // Update Toolbar Indicators
    spreadIndicator.textContent = `Spread ${state.currentSpreadIndex + 1} / ${spreads.length} (Trang ${leftPageNum}–${rightPageNum})`;
    spreadSlider.max = spreads.length;
    spreadSlider.value = state.currentSpreadIndex + 1;

    // Update Inspector Controls
    selectLayout.value = currentSpread.layout || 'magazine_grid_4';
    inputCaption.value = currentSpread.caption || '';
    inputNote.value = currentSpread.designer_note || '';

    if (currentSpread.layout === 'editorial_quote') {
      quoteGroup.style.display = 'block';
      inputQuote.value = currentSpread.quote || '';
    } else {
      quoteGroup.style.display = 'none';
      inputQuote.value = '';
    }

    const slots = currentSpread.slots || [];
    const layout = currentSpread.layout || 'magazine_grid_4';
    const guidesClass = state.guidesVisible ? 'visible' : '';

    // Render Slots according to template layout
    let leftContentHtml = '';
    let rightContentHtml = '';

    if (layout === 'cover_hero') {
      // Cover on Right page, Intro spine on Left
      leftContentHtml = `
        <div class="page-editorial-box">
          <span class="editorial-kicker">KỶ YẾU TỐT NGHIỆP 2026</span>
          <h2 class="editorial-title">${escapeHtml(state.spec.title || 'Biện Trung Tín')}</h2>
          <div class="editorial-divider"></div>
          <p class="editorial-body">${escapeHtml(state.spec.subtitle || 'Kỹ Sư Công Nghệ Thông Tin')}</p>
          <div class="editorial-meta-box">
            <span>Niên khóa: 2022 – 2026</span>
            <span>Khổ in: ${escapeHtml(state.spec.page_size || '200x200mm')}</span>
          </div>
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-cover-hero">
          ${renderSlotHtml(slots[0], 0, layout)}
        </div>
      `;
    } else if (layout === 'editorial_quote') {
      leftContentHtml = `
        <div class="page-editorial-box editorial-quote-box">
          <span class="quote-mark">“</span>
          <blockquote class="quote-text">${escapeHtml(currentSpread.quote || 'Cảm ơn Thầy Cô, Gia đình và Bạn bè đã luôn đồng hành trên suốt chặng đường 4 năm thanh xuân rực rỡ.')}</blockquote>
          <span class="quote-author">— Biện Trung Tín</span>
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-editorial-hero">
          ${renderSlotHtml(slots[0], 0, layout)}
        </div>
      `;
    } else if (layout === 'symmetric_duo') {
      leftContentHtml = `
        <div class="page-slots-grid layout-single-full">
          ${renderSlotHtml(slots[0], 0, layout)}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-single-full">
          ${renderSlotHtml(slots[1], 1, layout)}
        </div>
      `;
    } else if (layout === 'polaroid_scrapbook') {
      leftContentHtml = `
        <div class="page-slots-grid layout-polaroid-left">
          ${renderSlotHtml(slots[0], 0, layout, 'polaroid-tilt-left')}
          ${renderSlotHtml(slots[1], 1, layout, 'polaroid-tilt-right')}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-polaroid-right">
          ${renderSlotHtml(slots[2], 2, layout, 'polaroid-tilt-center')}
          <div class="scrapbook-note-card">
            <span class="note-pin"><i data-lucide="pin"></i></span>
            <p class="note-handwriting">${escapeHtml(currentSpread.caption || 'Những khoảnh khắc không thể nào quên')}</p>
          </div>
        </div>
      `;
    } else if (layout === 'mosaic_gallery_6') {
      // 3 on Left, 3 on Right
      leftContentHtml = `
        <div class="page-slots-grid layout-mosaic-side">
          <div class="mosaic-top">${renderSlotHtml(slots[0], 0, layout)}</div>
          <div class="mosaic-bottom-duo">
            ${renderSlotHtml(slots[1], 1, layout)}
            ${renderSlotHtml(slots[2], 2, layout)}
          </div>
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-mosaic-side">
          <div class="mosaic-bottom-duo">
            ${renderSlotHtml(slots[3], 3, layout)}
            ${renderSlotHtml(slots[4], 4, layout)}
          </div>
          <div class="mosaic-top">${renderSlotHtml(slots[5], 5, layout)}</div>
        </div>
      `;
    } else if (layout === 'classroom_grid_8') {
      // 4 on Left (2x2), 4 on Right (2x2)
      leftContentHtml = `
        <div class="page-slots-grid layout-grid-2x2">
          ${renderSlotHtml(slots[0], 0, layout)}
          ${renderSlotHtml(slots[1], 1, layout)}
          ${renderSlotHtml(slots[2], 2, layout)}
          ${renderSlotHtml(slots[3], 3, layout)}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-grid-2x2">
          ${renderSlotHtml(slots[4], 4, layout)}
          ${renderSlotHtml(slots[5], 5, layout)}
          ${renderSlotHtml(slots[6], 6, layout)}
          ${renderSlotHtml(slots[7], 7, layout)}
        </div>
      `;
    } else if (layout === 'megagrid_12') {
      // 6 on Left (3x2), 6 on Right (3x2)
      leftContentHtml = `
        <div class="page-slots-grid layout-megagrid-side">
          ${renderSlotHtml(slots[0], 0, layout)}
          ${renderSlotHtml(slots[1], 1, layout)}
          ${renderSlotHtml(slots[2], 2, layout)}
          ${renderSlotHtml(slots[3], 3, layout)}
          ${renderSlotHtml(slots[4], 4, layout)}
          ${renderSlotHtml(slots[5], 5, layout)}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-megagrid-side">
          ${renderSlotHtml(slots[6], 6, layout)}
          ${renderSlotHtml(slots[7], 7, layout)}
          ${renderSlotHtml(slots[8], 8, layout)}
          ${renderSlotHtml(slots[9], 9, layout)}
          ${renderSlotHtml(slots[10], 10, layout)}
          ${renderSlotHtml(slots[11], 11, layout)}
        </div>
      `;
    } else if (layout === 'panoramic_spread') {
      // Single continuous image spanning both pages
      leftContentHtml = `
        <div class="page-slots-grid layout-panoramic-half">
          ${renderSlotHtml(slots[0], 0, layout, 'panoramic-left')}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-panoramic-half">
          ${renderSlotHtml(slots[0], 0, layout, 'panoramic-right')}
        </div>
      `;
    } else {
      // Default: Magazine Grid 4 (1 big on Left, 3 on Right)
      leftContentHtml = `
        <div class="page-slots-grid layout-magazine-main">
          ${renderSlotHtml(slots[0], 0, layout)}
        </div>
      `;
      rightContentHtml = `
        <div class="page-slots-grid layout-magazine-sub">
          ${renderSlotHtml(slots[1], 1, layout)}
          ${renderSlotHtml(slots[2], 2, layout)}
          ${renderSlotHtml(slots[3], 3, layout)}
        </div>
      `;
    }

    const captionFooter = currentSpread.caption ? `
      <div class="page-caption-footer" contenteditable="true" title="Bấm vào để sửa lời tựa">${escapeHtml(currentSpread.caption)}</div>
    ` : '';

    bookContainer.innerHTML = `
      <div class="yearbook-double-spread" data-spread-idx="${state.currentSpreadIndex}">
        <!-- Left Page -->
        <div class="yearbook-page yearbook-page-left" data-page-side="left">
          <div class="page-corner-number left">${leftPageNum}</div>
          <div class="prepress-guides-layer ${guidesClass}">
            <div class="guide-bleed" title="Bleed 3mm (Vùng xén ngoài)"></div>
            <div class="guide-trim" title="Trim Line (Vạch cắt thành phẩm)"></div>
            <div class="guide-safe" title="Safe Zone 8mm (Vùng an toàn)"></div>
          </div>
          <div class="page-content-wrapper">
            ${leftContentHtml}
          </div>
          ${captionFooter}
        </div>

        <!-- Center Realistic Gutter Shadow Spine -->
        <div class="yearbook-book-spine" title="Rãnh gáy sách (Spine)"></div>

        <!-- Right Page -->
        <div class="yearbook-page yearbook-page-right" data-page-side="right">
          <div class="page-corner-number right">${rightPageNum}</div>
          <div class="prepress-guides-layer ${guidesClass}">
            <div class="guide-bleed" title="Bleed 3mm (Vùng xén ngoài)"></div>
            <div class="guide-trim" title="Trim Line (Vạch cắt thành phẩm)"></div>
            <div class="guide-safe" title="Safe Zone 8mm (Vùng an toàn)"></div>
          </div>
          <div class="page-content-wrapper">
            ${rightContentHtml}
          </div>
        </div>
      </div>
    `;

    bindSlotInteractions();
    refreshLucideIcons();
  }

  function renderSlotHtml(slot, index, layout, extraClass = '') {
    if (!slot) {
      slot = { id: `slot_${index + 1}`, photo: null, caption: '' };
    }

    const photoData = getPhotoDisplayData(slot.photo || slot.photo_path);
    const isPolaroid = layout === 'polaroid_scrapbook';

    let innerContent = '';
    if (photoData) {
      if (photoData.isMock) {
        innerContent = `
          <div class="mock-photo-card" style="background: ${photoData.bg};">
            <div class="mock-photo-content">
              <div class="mock-badge"><i data-lucide="${photoData.icon}"></i> ${escapeHtml(photoData.tag)}</div>
              <div class="mock-title">${escapeHtml(photoData.title)}</div>
              <div class="mock-sub">${escapeHtml(photoData.sub)}</div>
            </div>
            <div class="slot-actions-overlay">
              <span class="slot-idx-badge">#${index + 1}</span>
              <button type="button" class="btn-slot-remove" data-slot-idx="${index}" title="Gỡ ảnh khỏi khung"><i data-lucide="x"></i></button>
            </div>
          </div>
        `;
      } else {
        innerContent = `
          <div class="real-photo-card">
            <img src="${escapeHtml(photoData.url)}" alt="Yearbook Photo" class="slot-image" />
            <div class="slot-actions-overlay">
              <span class="slot-idx-badge">#${index + 1}</span>
              <button type="button" class="btn-slot-remove" data-slot-idx="${index}" title="Gỡ ảnh khỏi khung"><i data-lucide="x"></i></button>
            </div>
          </div>
        `;
      }
    } else {
      innerContent = `
        <div class="slot-empty-state">
          <i data-lucide="image"></i>
          <span>Kéo ảnh vào khung #${index + 1}</span>
        </div>
      `;
    }

    if (isPolaroid) {
      return `
        <div class="yearbook-slot slot-polaroid ${extraClass}" 
             data-slot-idx="${index}" 
             draggable="${slot.photo ? 'true' : 'false'}">
          <div class="polaroid-photo-frame">
            ${innerContent}
          </div>
          <div class="polaroid-caption" contenteditable="true" data-slot-idx="${index}">
            ${escapeHtml(slot.caption || `Khoảnh khắc #${index + 1}`)}
          </div>
        </div>
      `;
    }

    return `
      <div class="yearbook-slot ${extraClass}" 
           data-slot-idx="${index}" 
           draggable="${slot.photo ? 'true' : 'false'}">
        ${innerContent}
        ${slot.caption ? `<div class="slot-caption-tag" contenteditable="true" data-slot-idx="${index}">${escapeHtml(slot.caption)}</div>` : ''}
      </div>
    `;
  }

  // =========================================================================
  // 3. DRAG & DROP AND EVENT INTERACTIONS
  // =========================================================================

  function bindSlotInteractions() {
    const slots = bookContainer.querySelectorAll('.yearbook-slot');
    const removeButtons = bookContainer.querySelectorAll('.btn-slot-remove');
    const captionFooter = bookContainer.querySelector('.page-caption-footer');
    const polaroidCaptions = bookContainer.querySelectorAll('.polaroid-caption');

    // Remove Photo from Slot
    removeButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const slotIdx = parseInt(btn.getAttribute('data-slot-idx'), 10);
        const currentSpread = state.spec.spreads[state.currentSpreadIndex];
        if (currentSpread && currentSpread.slots && currentSpread.slots[slotIdx]) {
          currentSpread.slots[slotIdx].photo = null;
          renderCurrentSpread();
          showToast('Đã gỡ ảnh khỏi khung', 'info');
        }
      });
    });

    // Inline Caption Footer Live Edit
    if (captionFooter) {
      captionFooter.addEventListener('blur', () => {
        const text = captionFooter.innerText.trim();
        const currentSpread = state.spec.spreads[state.currentSpreadIndex];
        if (currentSpread) {
          currentSpread.caption = text;
          inputCaption.value = text;
        }
      });
    }

    // Polaroid Captions Live Edit
    polaroidCaptions.forEach(capEl => {
      capEl.addEventListener('blur', () => {
        const slotIdx = parseInt(capEl.getAttribute('data-slot-idx'), 10);
        const text = capEl.innerText.trim();
        const currentSpread = state.spec.spreads[state.currentSpreadIndex];
        if (currentSpread && currentSpread.slots && currentSpread.slots[slotIdx]) {
          currentSpread.slots[slotIdx].caption = text;
        }
      });
    });

    // Drag & Drop for Slots
    slots.forEach(slotEl => {
      // Slot as Drag Source (for Slot-to-Slot Swapping)
      slotEl.addEventListener('dragstart', (e) => {
        const slotIdx = parseInt(slotEl.getAttribute('data-slot-idx'), 10);
        const currentSpread = state.spec.spreads[state.currentSpreadIndex];
        if (currentSpread && currentSpread.slots && currentSpread.slots[slotIdx] && currentSpread.slots[slotIdx].photo) {
          state.draggedItem = {
            type: 'slot-photo',
            slotIdx: slotIdx,
            photo: currentSpread.slots[slotIdx].photo
          };
          e.dataTransfer.setData('text/plain', JSON.stringify({ type: 'slot-photo', slotIdx }));
          slotEl.classList.add('dragging');
        }
      });

      slotEl.addEventListener('dragend', () => {
        slotEl.classList.remove('dragging');
        slots.forEach(s => s.classList.remove('drag-over'));
      });

      // Slot as Drop Target
      slotEl.addEventListener('dragover', (e) => {
        e.preventDefault();
        slotEl.classList.add('drag-over');
      });

      slotEl.addEventListener('dragleave', () => {
        slotEl.classList.remove('drag-over');
      });

      slotEl.addEventListener('drop', (e) => {
        e.preventDefault();
        slotEl.classList.remove('drag-over');
        const targetSlotIdx = parseInt(slotEl.getAttribute('data-slot-idx'), 10);
        const currentSpread = state.spec.spreads[state.currentSpreadIndex];
        if (!currentSpread || !currentSpread.slots) return;

        // Ensure slot exists
        while (currentSpread.slots.length <= targetSlotIdx) {
          currentSpread.slots.push({ id: `slot_${currentSpread.slots.length + 1}`, photo: null });
        }

        if (state.draggedItem && state.draggedItem.type === 'tray-photo') {
          // Drop from Photo Tray into Slot
          currentSpread.slots[targetSlotIdx].photo = state.draggedItem.photo;
          state.draggedItem = null;
          renderCurrentSpread();
          updatePhotoTray();
          showToast(`Đã gán ảnh vào khung #${targetSlotIdx + 1}`, 'success');
        } else if (state.draggedItem && state.draggedItem.type === 'slot-photo') {
          // Slot-to-Slot Swapping
          const sourceIdx = state.draggedItem.slotIdx;
          if (sourceIdx !== targetSlotIdx) {
            const tempPhoto = currentSpread.slots[targetSlotIdx].photo;
            currentSpread.slots[targetSlotIdx].photo = currentSpread.slots[sourceIdx].photo;
            currentSpread.slots[sourceIdx].photo = tempPhoto;
            state.draggedItem = null;
            renderCurrentSpread();
            showToast(`Đã hoán đổi vị trí khung #${sourceIdx + 1} và #${targetSlotIdx + 1}`, 'success');
          }
        }
      });
    });
  }

  // =========================================================================
  // 4. PHOTO TRAY MANAGEMENT & UPLOAD
  // =========================================================================

  function updatePhotoTray() {
    if (trayCountBadge) trayCountBadge.textContent = `${state.photoTray.length} ảnh`;
    const subtabBadge = document.getElementById('yb-subtab-photo-badge');
    if (subtabBadge) subtabBadge.textContent = state.photoTray.length;

    if (state.photoTray.length === 0) {
      if (trayPhotosContainer) {
        trayPhotosContainer.innerHTML = `
          <div class="tray-empty-hint">
            <i data-lucide="upload-cloud"></i>
            <span>Khay ảnh trống. Hãy tải ảnh kỷ yếu vào đây!</span>
          </div>
        `;
      }
      refreshLucideIcons();
      return;
    }

    // Check which photos are already placed in album
    const usedPhotos = new Set();
    if (state.spec && state.spec.spreads) {
      state.spec.spreads.forEach(sp => {
        (sp.slots || []).forEach(sl => {
          if (sl.photo) {
            const pKey = typeof sl.photo === 'string' ? sl.photo : (sl.photo.filepath || sl.photo.url);
            usedPhotos.add(pKey);
          }
        });
      });
    }

    trayPhotosContainer.innerHTML = state.photoTray.map((item, idx) => {
      const pKey = typeof item === 'string' ? item : (item.filepath || item.url);
      const isUsed = usedPhotos.has(pKey);
      const photoData = getPhotoDisplayData(item);

      let thumbHtml = '';
      if (photoData.isMock) {
        thumbHtml = `
          <div class="tray-thumb-mock" style="background: ${photoData.bg}">
            <i data-lucide="${photoData.icon}"></i>
          </div>
        `;
      } else {
        thumbHtml = `<img src="${escapeHtml(photoData.url)}" alt="Thumb" class="tray-thumb-img" />`;
      }

      return `
        <div class="tray-photo-chip ${isUsed ? 'used' : ''}" 
             data-tray-idx="${idx}" 
             draggable="true" 
             title="${escapeHtml(photoData.title)} (${photoData.sub}) - Kéo vào khung để đặt ảnh">
          ${thumbHtml}
          <div class="chip-info">
            <span class="chip-name">${escapeHtml(photoData.title)}</span>
            <span class="chip-ratio-badge">${escapeHtml(photoData.sub)}</span>
          </div>
          ${isUsed ? '<span class="chip-used-badge"><i data-lucide="check"></i></span>' : ''}
        </div>
      `;
    }).join('');

    // Bind Tray Drag events
    const chips = trayPhotosContainer.querySelectorAll('.tray-photo-chip');
    chips.forEach(chip => {
      chip.addEventListener('dragstart', (e) => {
        const idx = parseInt(chip.getAttribute('data-tray-idx'), 10);
        const photo = state.photoTray[idx];
        state.draggedItem = {
          type: 'tray-photo',
          photo: photo
        };
        e.dataTransfer.setData('text/plain', JSON.stringify({ type: 'tray-photo', idx }));
        chip.classList.add('dragging');
      });

      chip.addEventListener('dragend', () => {
        chip.classList.remove('dragging');
      });
    });

    refreshLucideIcons();
  }

  async function handlePhotoUpload(files) {
    if (!files || files.length === 0) return;

    showToast(`Đang tải lên ${files.length} ảnh kỷ yếu...`, 'info');

    let uploadedCount = 0;
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (!file.type.startsWith('image/')) continue;

      const formData = new FormData();
      formData.append('file', file);

      try {
        const res = await fetch('/api/upload', {
          method: 'POST',
          body: formData
        });
        const data = await res.json();
        if (data && data.filepath) {
          // Detect image orientation
          const aspect = await detectImageAspect(file);
          state.photoTray.push({
            filepath: data.filepath,
            filename: file.name,
            url: `/api/files/download?path=${encodeURIComponent(data.filepath)}`,
            aspect: aspect,
            size: file.size
          });
          uploadedCount++;
        }
      } catch (err) {
        console.error('Upload error:', err);
      }
    }

    if (uploadedCount > 0) {
      updatePhotoTray();
      showToast(`Đã thêm ${uploadedCount} ảnh vào khay! Bấm "Dàn Trang Tự Động" để tạo album ngay.`, 'success');
    }
  }

  function detectImageAspect(file) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const ratio = img.width / img.height;
        if (ratio >= 1.2) resolve('landscape');
        else if (ratio <= 0.85) resolve('portrait');
        else resolve('square');
        URL.revokeObjectURL(img.src);
      };
      img.onerror = () => resolve('landscape');
      img.src = URL.createObjectURL(file);
    });
  }

  // =========================================================================
  // 5. SERVER API CALLS: INITIAL MOCK, AUTO-LAYOUT & EXPORTS
  // =========================================================================

  async function loadInitialAlbum() {
    try {
      const res = await fetch('/api/yearbook/mock-album');
      if (res.ok) {
        const data = await res.json();
        if (data && data.spec) {
          state.spec = normalizeSpec(data.spec);
          populateInitialFormFields();
          extractPhotosIntoTray();
          renderCurrentSpread();
          showToast('Đã tải cuốn Kỷ yếu mẫu 8 trang đôi cử nhân!', 'info');
          return;
        }
      }
    } catch (e) {
      console.warn('Could not fetch server mock album, using embedded fallback:', e);
    }
    // Fallback embedded mock spec
    state.spec = normalizeSpec(getEmbeddedFallbackSpec());
    populateInitialFormFields();
    extractPhotosIntoTray();
    renderCurrentSpread();
  }

  function populateInitialFormFields() {
    if (!state.spec) return;
    if (albumTitleInput) albumTitleInput.value = state.spec.title || 'KỶ YẾU TỐT NGHIỆP 2026';
    if (albumSubInput) albumSubInput.value = state.spec.subtitle || 'BIỆN TRUNG TÍN';
    if (state.spec.page_size && selectSize) selectSize.value = state.spec.page_size;
    if (state.spec.print_settings && state.spec.print_settings.bleed_mm && sliderBleed) {
      sliderBleed.value = state.spec.print_settings.bleed_mm;
      if (bleedVal) bleedVal.textContent = `${state.spec.print_settings.bleed_mm.toFixed(1)} mm`;
    }
  }

  function extractPhotosIntoTray() {
    state.photoTray = [];
    // Populate tray with available mock photos
    Object.keys(MOCK_PLACEHOLDERS).forEach(mockId => {
      state.photoTray.push(mockId);
    });
    updatePhotoTray();
  }

  async function runAutoLayout() {
    if (state.photoTray.length === 0) {
      showToast('Khay ảnh đang trống. Hãy tải ảnh lên trước khi dàn trang!', 'warning');
      return;
    }

    showToast('Đang chạy thuật toán tự động dàn trang 9 mẫu kỷ yếu...', 'info');

    try {
      const res = await fetch('/api/yearbook/auto-layout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          photos: state.photoTray,
          album_title: albumTitleInput.value.trim() || 'KỶ YẾU TỐT NGHIỆP 2026',
          album_subtitle: albumSubInput.value.trim() || 'BIỆN TRUNG TÍN',
          page_size: selectSize.value
        })
      });

      if (res.ok) {
        const data = await res.json();
        if (data && data.spec) {
          state.spec = normalizeSpec(data.spec);
          state.currentSpreadIndex = 0;
          renderCurrentSpread();
          updatePhotoTray();
          const spreadCount = state.spec?.spreads?.length || 0;
          showToast(`Dàn trang thành công! Đã tạo ${spreadCount} trang đôi chuẩn editorial.`, 'success');
          return;
        }
      }
      throw new Error('Server returned invalid spec');
    } catch (err) {
      console.error('Auto layout failed:', err);
      showToast(`Lỗi tự động dàn trang: ${err.message}`, 'error');
    }
  }

  async function triggerExport(exportType, successMsg) {
    if (!state.spec) return;

    exportStatusBox.style.display = 'block';
    statusSpinner.style.display = 'block';
    btnDownloadResult.style.display = 'none';
    statusText.textContent = 'Đang biên soạn và xuất bản tài liệu...';

    // Synchronize latest spec inputs
    if (albumTitleInput) state.spec.title = albumTitleInput.value.trim() || state.spec.title || 'KỶ YẾU TỐT NGHIỆP 2026';
    if (albumSubInput) state.spec.subtitle = albumSubInput.value.trim() || state.spec.subtitle || 'BIỆN TRUNG TÍN';
    if (selectSize) state.spec.page_size = selectSize.value || state.spec.page_size || 'photobook_square_25x25';
    state.spec.print_settings = state.spec.print_settings || {};
    if (sliderBleed) state.spec.print_settings.bleed_mm = parseFloat(sliderBleed.value) || 3.0;

    try {
      const res = await fetch('/api/yearbook/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spec: state.spec,
          export_type: exportType
        })
      });

      const data = await res.json();
      if (res.ok && data.success && data.download_url) {
        statusSpinner.style.display = 'none';
        statusText.textContent = 'Xuất bản thành công!';
        btnDownloadResult.style.display = 'inline-flex';
        btnDownloadResult.href = data.download_url;
        btnDownloadResult.download = data.filename || 'yearbook_export';
        showToast(successMsg, 'success');

        if (loadHistory) loadHistory();
      } else {
        throw new Error(data.error || 'Xuất bản thất bại');
      }
    } catch (err) {
      console.error('Export error:', err);
      statusSpinner.style.display = 'none';
      statusText.textContent = `Thất bại: ${err.message}`;
      showToast(`Lỗi xuất bản: ${err.message}`, 'error');
    }
  }

  // =========================================================================
  // 6. EVENT LISTENERS SETUP
  // =========================================================================

  // Bleed Slider change
  sliderBleed.addEventListener('input', () => {
    bleedVal.textContent = `${parseFloat(sliderBleed.value).toFixed(1)} mm`;
  });

  // Toggle Pre-Press Guides
  togglePrepressGuides.addEventListener('change', () => {
    state.guidesVisible = togglePrepressGuides.checked;
    const guidesLayers = bookContainer.querySelectorAll('.prepress-guides-layer');
    guidesLayers.forEach(gl => {
      if (state.guidesVisible) gl.classList.add('visible');
      else gl.classList.remove('visible');
    });
    showToast(state.guidesVisible ? 'Đã bật Vạch Xén In Ấn (Bleed, Trim, Safe Zone)' : 'Đã ẩn Vạch Xén In Ấn', 'info');
  });

  // Spread Navigation
  btnPrevSpread.addEventListener('click', () => {
    if (state.currentSpreadIndex > 0) {
      state.currentSpreadIndex--;
      renderCurrentSpread();
    }
  });

  btnNextSpread.addEventListener('click', () => {
    if (state.spec && state.spec.spreads && state.currentSpreadIndex < state.spec.spreads.length - 1) {
      state.currentSpreadIndex++;
      renderCurrentSpread();
    }
  });

  spreadSlider.addEventListener('input', () => {
    const targetIdx = parseInt(spreadSlider.value, 10) - 1;
    if (targetIdx >= 0 && targetIdx < (state.spec?.spreads?.length || 0)) {
      state.currentSpreadIndex = targetIdx;
      renderCurrentSpread();
    }
  });

  // Keyboard navigation when Yearbook tab is active
  window.addEventListener('keydown', (e) => {
    if (!panelYearbook.classList.contains('active')) return;
    if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA' || document.activeElement.isContentEditable)) {
      return;
    }
    if (e.key === 'ArrowLeft') {
      if (state.currentSpreadIndex > 0) {
        state.currentSpreadIndex--;
        renderCurrentSpread();
      }
    } else if (e.key === 'ArrowRight') {
      if (state.spec && state.spec.spreads && state.currentSpreadIndex < state.spec.spreads.length - 1) {
        state.currentSpreadIndex++;
        renderCurrentSpread();
      }
    }
  });

  // Layout Selector Change
  selectLayout.addEventListener('change', () => {
    if (!state.spec || !state.spec.spreads) return;
    const currentSpread = state.spec.spreads[state.currentSpreadIndex];
    if (!currentSpread) return;

    const newLayout = selectLayout.value;
    currentSpread.layout = newLayout;
    currentSpread.slots = createSlotsForLayout(newLayout, currentSpread.slots);
    renderCurrentSpread();
    showToast(`Đã chuyển sang mẫu bố cục: ${newLayout}`, 'info');
  });

  // Inspector Inputs Live Sync
  inputCaption.addEventListener('input', () => {
    if (state.spec && state.spec.spreads && state.spec.spreads[state.currentSpreadIndex]) {
      state.spec.spreads[state.currentSpreadIndex].caption = inputCaption.value;
      const footer = bookContainer.querySelector('.page-caption-footer');
      if (footer) footer.innerText = inputCaption.value;
    }
  });

  inputNote.addEventListener('input', () => {
    if (state.spec && state.spec.spreads && state.spec.spreads[state.currentSpreadIndex]) {
      state.spec.spreads[state.currentSpreadIndex].designer_note = inputNote.value;
    }
  });

  inputQuote.addEventListener('input', () => {
    if (state.spec && state.spec.spreads && state.spec.spreads[state.currentSpreadIndex]) {
      state.spec.spreads[state.currentSpreadIndex].quote = inputQuote.value;
      const quoteEl = bookContainer.querySelector('.quote-text');
      if (quoteEl) quoteEl.innerText = inputQuote.value;
    }
  });

  // Add Spread
  btnAddSpread.addEventListener('click', () => {
    if (!state.spec) return;
    state.spec.spreads = state.spec.spreads || [];
    const newSpread = {
      spread_index: state.spec.spreads.length + 1,
      layout: 'magazine_grid_4',
      caption: 'Trang kỷ yếu mới',
      designer_note: '',
      slots: createSlotsForLayout('magazine_grid_4')
    };
    state.spec.spreads.push(newSpread);
    state.currentSpreadIndex = state.spec.spreads.length - 1;
    renderCurrentSpread();
    showToast('Đã thêm 1 trang đôi mới vào album!', 'success');
  });

  // Duplicate Spread
  btnDuplicateSpread.addEventListener('click', () => {
    if (!state.spec || !state.spec.spreads || state.spec.spreads.length === 0) return;
    const current = state.spec.spreads[state.currentSpreadIndex];
    const cloned = JSON.parse(JSON.stringify(current));
    cloned.spread_index = state.spec.spreads.length + 1;
    state.spec.spreads.splice(state.currentSpreadIndex + 1, 0, cloned);
    state.currentSpreadIndex++;
    renderCurrentSpread();
    showToast('Đã nhân bản trang đôi hiện tại!', 'success');
  });

  // Delete Spread
  btnDeleteSpread.addEventListener('click', () => {
    if (!state.spec || !state.spec.spreads || state.spec.spreads.length <= 1) {
      showToast('Cuốn kỷ yếu phải giữ lại ít nhất 1 trang đôi!', 'warning');
      return;
    }
    state.spec.spreads.splice(state.currentSpreadIndex, 1);
    if (state.currentSpreadIndex >= state.spec.spreads.length) {
      state.currentSpreadIndex = state.spec.spreads.length - 1;
    }
    renderCurrentSpread();
    showToast('Đã xóa trang đôi!', 'info');
  });

  // Auto Layout Button
  btnAutoLayout.addEventListener('click', runAutoLayout);

  // Clear Tray Button
  btnClearTray.addEventListener('click', () => {
    state.photoTray = [];
    updatePhotoTray();
    showToast('Đã làm trống khay ảnh!', 'info');
  });

  // Photo Upload Browse & Drop
  btnBrowsePhotos.addEventListener('click', () => photoFileInput.click());
  photoFileInput.addEventListener('change', () => {
    if (photoFileInput.files && photoFileInput.files.length > 0) {
      handlePhotoUpload(Array.from(photoFileInput.files));
      photoFileInput.value = '';
    }
  });

  photoDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    photoDropZone.classList.add('drag-over');
  });
  photoDropZone.addEventListener('dragleave', () => {
    photoDropZone.classList.remove('drag-over');
  });
  photoDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    photoDropZone.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handlePhotoUpload(Array.from(e.dataTransfer.files));
    }
  });

  // Export Buttons
  btnExportSpreadPdf.addEventListener('click', () => {
    triggerExport('spread_pdf', 'Tải thành công PDF Mở Phẳng (Spread View)!');
  });

  btnExportPrintPdf.addEventListener('click', () => {
    triggerExport('print_pdf', 'Tải thành công PDF In Ấn 300 DPI kèm Bleed & Crop Marks!');
  });

  btnExportHandoffZip.addEventListener('click', () => {
    triggerExport('handoff_zip', 'Gói Designer Handoff 5-trong-1 đã sẵn sàng tải về!');
  });

  // Sub-tabs in Column 1 (Photos vs Print Settings)
  const subtabBtns = panelYearbook.querySelectorAll('.yb-subtab-btn');
  const subpanels = {
    photos: document.getElementById('yb-subpanel-photos'),
    settings: document.getElementById('yb-subpanel-settings'),
  };

  subtabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.getAttribute('data-ybtab');
      if (!subpanels[target]) return;
      subtabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      Object.values(subpanels).forEach(p => p && p.classList.remove('active'));
      subpanels[target].classList.add('active');
      refreshLucideIcons();
    });
  });

  // Re-render when Yearbook tab becomes active
  window.addEventListener('yearbook:activated', () => {
    if (state.spec && state.spec.spreads) {
      renderCurrentSpread();
      updatePhotoTray();
    }
  });

  // Initialize Initial Album
  loadInitialAlbum();
}

// Helper escape HTML
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Embedded Fallback Spec (Guarantees zero-config instant render)
function getEmbeddedFallbackSpec() {
  return {
    title: 'KỶ YẾU TỐT NGHIỆP 2026',
    subtitle: 'BIỆN TRUNG TÍN - KỸ SƯ CÔNG NGHỆ THÔNG TIN',
    page_size: '200x200mm',
    binding: 'layflat',
    print_settings: {
      bleed_mm: 3.0,
      safe_margin_mm: 8.0,
      target_dpi: 300,
      crop_marks: true
    },
    spreads: [
      {
        spread_index: 1,
        layout: 'cover_hero',
        caption: 'Bìa Kỷ Yếu Cử Nhân 2026',
        designer_note: 'Bìa cứng phủ màng mờ bồi carton 3 lớp',
        slots: [{ id: 'slot_1', photo: 'mock:grad_cap', caption: 'Chân dung cử nhân' }]
      },
      {
        spread_index: 2,
        layout: 'editorial_quote',
        caption: 'Lời Tri Ân Thầy Cô & Gia Đình',
        quote: 'Cảm ơn Cha Mẹ đã nuôi dưỡng, Thầy Cô đã dạy dỗ và Bạn bè đã đồng hành trong suốt 4 năm đại học.',
        designer_note: 'Font chữ trang trọng, lề thoáng',
        slots: [{ id: 'slot_1', photo: 'mock:portrait', caption: 'Áo cử nhân' }]
      },
      {
        spread_index: 3,
        layout: 'symmetric_duo',
        caption: 'Lễ Trao Bằng Trọng Thể & Thư Viện',
        designer_note: 'Canh chỉnh thẳng hàng ngang 2 ảnh',
        slots: [
          { id: 'slot_1', photo: 'mock:ceremony', caption: 'Giây phút nhận bằng' },
          { id: 'slot_2', photo: 'mock:library', caption: 'Góc học tập' }
        ]
      },
      {
        spread_index: 4,
        layout: 'magazine_grid_4',
        caption: 'Bạn Thân & Kỷ Niệm Giảng Đường',
        designer_note: 'Ảnh bên trái điểm nhấn lớn',
        slots: [
          { id: 'slot_1', photo: 'mock:friends', caption: 'Nhóm bạn thân' },
          { id: 'slot_2', photo: 'mock:campus', caption: 'Khuôn viên trường' },
          { id: 'slot_3', photo: 'mock:sunset', caption: 'Hoàng hôn' },
          { id: 'slot_4', photo: 'mock:diploma', caption: 'Tấm bằng cử nhân' }
        ]
      },
      {
        spread_index: 5,
        layout: 'polaroid_scrapbook',
        caption: 'Góc Lưu Bút - Những Khoảnh Khắc Nghịch Ngợm',
        designer_note: 'Hiệu ứng viền trắng polaroid nghiêng tự nhiên',
        slots: [
          { id: 'slot_1', photo: 'mock:friends', caption: 'Nụ cười thanh xuân' },
          { id: 'slot_2', photo: 'mock:campus', caption: 'Check-in giảng đường' },
          { id: 'slot_3', photo: 'mock:portrait', caption: 'Tung mũ cử nhân' }
        ]
      },
      {
        spread_index: 6,
        layout: 'mosaic_gallery_6',
        caption: 'Bộ Sưu Tập Mosaic - 6 Khung Hình Đa Sắc',
        designer_note: 'Bố cục cân đối giữa ảnh ngang 16:9 và ảnh vuông 1:1',
        slots: [
          { id: 'slot_1', photo: 'mock:campus', caption: 'Toàn cảnh' },
          { id: 'slot_2', photo: 'mock:friends', caption: 'Bạn thân' },
          { id: 'slot_3', photo: 'mock:grad_cap', caption: 'Mũ cử nhân' },
          { id: 'slot_4', photo: 'mock:library', caption: 'Thư viện' },
          { id: 'slot_5', photo: 'mock:ceremony', caption: 'Nhận bằng' },
          { id: 'slot_6', photo: 'mock:sunset', caption: 'Hoàng hôn' }
        ]
      },
      {
        spread_index: 7,
        layout: 'classroom_grid_8',
        caption: 'Kỷ Yếu Tập Thể Lớp - 8 Thành Viên Tiêu Biểu',
        designer_note: 'Đều đặn 2x2 mỗi trang, khung ảnh vuông 1:1',
        slots: [
          { id: 'slot_1', photo: 'mock:portrait', caption: 'Thành viên 1' },
          { id: 'slot_2', photo: 'mock:grad_cap', caption: 'Thành viên 2' },
          { id: 'slot_3', photo: 'mock:friends', caption: 'Thành viên 3' },
          { id: 'slot_4', photo: 'mock:library', caption: 'Thành viên 4' },
          { id: 'slot_5', photo: 'mock:ceremony', caption: 'Thành viên 5' },
          { id: 'slot_6', photo: 'mock:campus', caption: 'Thành viên 6' },
          { id: 'slot_7', photo: 'mock:diploma', caption: 'Thành viên 7' },
          { id: 'slot_8', photo: 'mock:sunset', caption: 'Thành viên 8' }
        ]
      },
      {
        spread_index: 8,
        layout: 'panoramic_spread',
        caption: 'Toàn Cảnh Giảng Đường - Tạm Biệt Giảng Đường Thanh Xuân',
        designer_note: 'Ảnh trải rộng 2 trang mở phẳng tuyệt đối (Layflat)',
        slots: [{ id: 'slot_1', photo: 'mock:campus', caption: 'Toàn cảnh giảng đường' }]
      }
    ]
  };
}
