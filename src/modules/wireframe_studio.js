/**
 * src/modules/wireframe_studio.js — Rapid Low-Fi UI Wireframe & Mockup Studio Controller
 *
 * Dual-Mode Authoring:
 *  - Mode 1 (JSON Spec): Declarative JSON specification with One-Way Generator to PlantUML Salt.
 *  - Mode 2 (Salt Code): Direct @startsalt ... @endsalt syntax editor with quick widget snippets.
 * Live Vector SVG Preview (-tsvg) + 300 DPI PNG Export + 1-Click Word (.docx) OpenXML Injection (ERR_DOCX_005).
 */

const FALLBACK_TEMPLATES = {
  login_auth: {
    id: 'login_auth',
    title: 'Màn Hình Đăng Nhập & Xác Thực 2FA',
    category: 'Authentication',
    description: 'Khung đăng nhập hệ thống doanh nghiệp kèm xác thực hai lớp OTP SMS và ghi nhớ thiết bị.',
    json_spec: {
      title: 'CỔNG ĐĂNG NHẬP HỆ THỐNG DOANH NGHIỆP (SSO & 2FA)',
      frame: '+',
      elements: [
        { type: 'separator' },
        { type: 'input', label: 'Tài khoản / Email:', value: 'ba.lead@enterprise.vn' },
        { type: 'password', label: 'Mật khẩu:', value: '************' },
        { type: 'droplist', label: 'Vai trò truy cập:', options: ['Business Analyst', 'Product Owner', 'System Admin'] },
        { type: 'input', label: 'Mã xác thực OTP (6 số):', value: '849201' },
        { type: 'checkbox', label: 'Ghi nhớ thiết bị tin cậy trong 30 ngày', checked: true },
        { type: 'separator' },
        { type: 'button_group', buttons: ['Đăng Nhập Hệ Thống', 'Gửi Lại OTP', 'Quên Mật Khẩu?'] }
      ]
    },
    salt_code: `@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>CỔNG ĐĂNG NHẬP HỆ THỐNG DOANH NGHIỆP (SSO & 2FA)
  --
  Tài khoản / Email:      | "ba.lead@enterprise.vn   "
  Mật khẩu:               | "************            "
  Vai trò truy cập:       | ^Business Analyst|Product Owner|System Admin^
  Mã xác thực OTP (6 số): | "849201                  "
  [X] Ghi nhớ thiết bị tin cậy trong 30 ngày
  --
  { [ Đăng Nhập Hệ Thống ] | [ Gửi Lại OTP ] | [ Quên Mật Khẩu? ] }
}
@endsalt`
  }
};

export function initWireframeStudio({ showToast, refreshLucideIcons, uploadFileToServer, loadHistory }) {
  const templateSelect = document.getElementById('wf-template-select');
  const btnLoadTemplate = document.getElementById('btn-wf-load-template');
  const presetBtns = document.querySelectorAll('.wf-preset-chip');
  const modeBtns = document.querySelectorAll('.wf-mode-btn');
  const editorTextarea = document.getElementById('wf-code-editor');
  const modeHintLabel = document.getElementById('wf-editor-mode-label');
  const snippetToolbar = document.getElementById('wf-snippet-toolbar');
  const snippetBtns = document.querySelectorAll('.wf-snippet-btn');

  const btnRender = document.getElementById('btn-wf-render');
  const btnCopyCode = document.getElementById('btn-wf-copy-code');
  const btnDownloadSvg = document.getElementById('btn-wf-download-svg');
  const btnDownloadPng = document.getElementById('btn-wf-download-png');

  const svgViewport = document.getElementById('wf-svg-viewport');
  const emptyState = document.getElementById('wf-empty-state');
  const dimBadge = document.getElementById('wf-dimensions-badge');
  const zoomBadge = document.getElementById('wf-zoom-badge');
  const btnZoomIn = document.getElementById('btn-wf-zoom-in');
  const btnZoomOut = document.getElementById('btn-wf-zoom-out');
  const btnZoomReset = document.getElementById('btn-wf-zoom-reset');

  const docxFileInput = document.getElementById('wf-docx-file');
  const docxHeadingInput = document.getElementById('wf-docx-heading');
  const docxCaptionInput = document.getElementById('wf-docx-caption');
  const btnInjectDocx = document.getElementById('btn-wf-inject-docx');

  if (!editorTextarea) return;

  let templatesCatalog = { ...FALLBACK_TEMPLATES };
  let activeMode = 'salt'; // 'salt' or 'json'
  let currentTemplateId = 'login_auth';
  let currentJsonStr = JSON.stringify(FALLBACK_TEMPLATES.login_auth.json_spec, null, 2);
  let currentSaltStr = FALLBACK_TEMPLATES.login_auth.salt_code;
  let lastSvgUrl = null;
  let lastPngUrl = null;
  let currentZoom = 1.0;

  // 1. Load templates catalog from backend
  async function fetchTemplates() {
    try {
      const res = await fetch('/api/diagram/wireframe/templates');
      if (res.ok) {
        const data = await res.json();
        if (data.success && data.templates) {
          templatesCatalog = data.templates;
          populateTemplateDropdown();
        }
      }
    } catch {
      // Fallback already populated
    }
  }

  function populateTemplateDropdown() {
    if (!templateSelect) return;
    templateSelect.innerHTML = '';
    Object.values(templatesCatalog).forEach((tpl) => {
      const opt = document.createElement('option');
      opt.value = tpl.id;
      opt.textContent = `[${tpl.category}] ${tpl.title}`;
      if (tpl.id === currentTemplateId) opt.selected = true;
      templateSelect.appendChild(opt);
    });
  }

  // 2. Apply selected template
  function applyTemplate(templateId, autoRender = true) {
    const tpl = templatesCatalog[templateId];
    if (!tpl) return;
    currentTemplateId = templateId;
    currentJsonStr = JSON.stringify(tpl.json_spec, null, 2);
    currentSaltStr = tpl.salt_code;

    if (templateSelect) templateSelect.value = templateId;
    presetBtns.forEach((btn) => {
      btn.classList.toggle('active', btn.getAttribute('data-tpl') === templateId);
    });

    if (activeMode === 'json') {
      editorTextarea.value = currentJsonStr;
    } else {
      editorTextarea.value = currentSaltStr;
    }

    if (docxCaptionInput && tpl.title) {
      docxCaptionInput.value = `Hình: Phác thảo giao diện — ${tpl.title}`;
    }

    if (autoRender) {
      renderWireframe(false);
    }
  }

  // 3. Switch between Salt Code Mode & JSON Spec Mode (One-Way JSON -> Salt support)
  function setAuthoringMode(mode) {
    if (mode === activeMode) return;

    // Save current buffer before switching
    if (activeMode === 'json') {
      currentJsonStr = editorTextarea.value;
    } else {
      currentSaltStr = editorTextarea.value;
    }

    activeMode = mode;
    modeBtns.forEach((b) => {
      b.classList.toggle('active', b.getAttribute('data-wf-mode') === mode);
    });

    if (mode === 'json') {
      editorTextarea.value = currentJsonStr;
      if (modeHintLabel) {
        modeHintLabel.textContent = 'Chế độ JSON Spec: Sửa cấu trúc JSON -> Hệ thống tự động sinh mã Salt khi kết xuất';
      }
      if (snippetToolbar) snippetToolbar.style.display = 'none';
    } else {
      editorTextarea.value = currentSaltStr;
      if (modeHintLabel) {
        modeHintLabel.textContent = 'Chế độ Salt Code: Chỉnh sửa trực tiếp cú pháp @startsalt ... @endsalt';
      }
      if (snippetToolbar) snippetToolbar.style.display = 'flex';
    }
  }

  modeBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetMode = btn.getAttribute('data-wf-mode');
      if (targetMode) setAuthoringMode(targetMode);
    });
  });

  if (btnLoadTemplate && templateSelect) {
    btnLoadTemplate.addEventListener('click', () => {
      applyTemplate(templateSelect.value, true);
      showToast('Đã nạp mẫu giao diện Wireframe chuẩn BA', 'info');
    });
    templateSelect.addEventListener('change', () => {
      applyTemplate(templateSelect.value, true);
    });
  }

  presetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const tid = btn.getAttribute('data-tpl');
      if (tid) applyTemplate(tid, true);
    });
  });

  // 4. Quick Widget Snippets in Salt Mode
  const SNIPPET_MAP = {
    button: '\n  [ Nút Hành Động ]',
    input: '\n  Nhãn trường: | "Giá trị mặc định     "',
    droplist: '\n  Danh mục: | ^Tùy chọn 1|Tùy chọn 2|Tùy chọn 3^',
    checkbox: '\n  [X] Tùy chọn đánh dấu kiểm (Checkbox)',
    radio: '\n  (X) Lựa chọn loại trừ (Radio Button)',
    table: '\n  {#\n    <b>Cột 1 | <b>Cột 2 | <b>Thao Tác\n    Dữ liệu A1 | Dữ liệu A2 | [Chi tiết]\n    Dữ liệu B1 | Dữ liệu B2 | [Chi tiết]\n  }',
    tabs: '\n  {/ <b>Tab Chính | Tab Phụ 2 | Tab Cấu Hình }'
  };

  snippetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      if (activeMode !== 'salt') {
        setAuthoringMode('salt');
      }
      const key = btn.getAttribute('data-snippet');
      const snippet = SNIPPET_MAP[key];
      if (!snippet) return;

      const val = editorTextarea.value;
      const closeIdx = val.lastIndexOf('}');
      if (closeIdx !== -1) {
        editorTextarea.value = val.slice(0, closeIdx) + snippet + '\n' + val.slice(closeIdx);
      } else {
        editorTextarea.value += snippet;
      }
      currentSaltStr = editorTextarea.value;
      editorTextarea.focus();
    });
  });

  // 5. Zoom Controls for SVG Viewport
  function applyZoom() {
    if (!svgViewport) return;
    const svgEl = svgViewport.querySelector('svg');
    if (svgEl) {
      svgEl.style.transform = `scale(${currentZoom})`;
      svgEl.style.transformOrigin = 'top center';
      svgEl.style.transition = 'transform 0.15s ease';
    }
    if (zoomBadge) {
      zoomBadge.textContent = `${Math.round(currentZoom * 100)}%`;
    }
  }

  if (btnZoomIn) {
    btnZoomIn.addEventListener('click', () => {
      currentZoom = Math.min(2.5, currentZoom + 0.2);
      applyZoom();
    });
  }
  if (btnZoomOut) {
    btnZoomOut.addEventListener('click', () => {
      currentZoom = Math.max(0.4, currentZoom - 0.2);
      applyZoom();
    });
  }
  if (btnZoomReset) {
    btnZoomReset.addEventListener('click', () => {
      currentZoom = 1.0;
      applyZoom();
    });
  }

  // 6. Render Wireframe (Live Vector SVG + 300 DPI PNG)
  async function renderWireframe(notify = true) {
    const rawText = editorTextarea.value.trim();
    if (!rawText) {
      showToast('Vui lòng nhập đặc tả JSON hoặc mã Salt trước khi kết xuất', 'error');
      return;
    }

    let payloadSpec = rawText;
    if (activeMode === 'json') {
      currentJsonStr = rawText;
      try {
        payloadSpec = JSON.parse(rawText);
      } catch (err) {
        showToast(`Lỗi cú pháp JSON Spec: ${err.message}`, 'error');
        return;
      }
    } else {
      currentSaltStr = rawText;
    }

    if (btnRender) {
      btnRender.disabled = true;
      btnRender.innerHTML = '<i data-lucide="refresh-cw" class="spin"></i> Đang kết xuất...';
      refreshLucideIcons();
    }

    try {
      const res = await fetch('/api/diagram/wireframe/render', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spec: payloadSpec,
          format: 'both',
          scale: 2.0
        })
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Không thể kết xuất Wireframe');
      }

      // One-Way Sync: if user was in JSON mode, update Salt buffer with generated Salt code
      if (data.salt_code) {
        currentSaltStr = data.salt_code;
      }

      lastSvgUrl = data.svg_url || null;
      lastPngUrl = data.png_url || data.download_url || null;

      if (emptyState) emptyState.style.display = 'none';
      if (svgViewport && data.svg) {
        svgViewport.style.display = 'flex';
        svgViewport.innerHTML = data.svg;
        currentZoom = 1.0;
        applyZoom();
      }

      if (dimBadge && Array.isArray(data.dimensions)) {
        dimBadge.textContent = `${data.dimensions[0]} × ${data.dimensions[1]} px (300 DPI + Vector SVG)`;
      }

      if (btnDownloadSvg) btnDownloadSvg.disabled = !lastSvgUrl;
      if (btnDownloadPng) btnDownloadPng.disabled = !lastPngUrl;

      if (loadHistory) loadHistory();
      if (notify) {
        showToast('Đã kết xuất Wireframe Vector SVG & PNG 300 DPI!', 'success');
      }
    } catch (err) {
      if (notify) {
        showToast(`Lỗi kết xuất Wireframe: ${err.message}`, 'error');
      }
    } finally {
      if (btnRender) {
        btnRender.disabled = false;
        btnRender.innerHTML = '<i data-lucide="play"></i> Kết Xuất Wireframe (Live SVG)';
        refreshLucideIcons();
      }
    }
  }

  if (btnRender) {
    btnRender.addEventListener('click', () => renderWireframe(true));
  }

  if (btnCopyCode) {
    btnCopyCode.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(editorTextarea.value);
        showToast('Đã sao chép mã Wireframe vào bộ nhớ tạm!', 'success');
      } catch {
        showToast('Không thể sao chép tự động', 'error');
      }
    });
  }

  if (btnDownloadSvg) {
    btnDownloadSvg.addEventListener('click', () => {
      if (lastSvgUrl) window.open(lastSvgUrl, '_blank');
    });
  }

  if (btnDownloadPng) {
    btnDownloadPng.addEventListener('click', () => {
      if (lastPngUrl) window.open(lastPngUrl, '_blank');
    });
  }

  // 7. 1-Click Word (.docx) OpenXML Injection (ERR_DOCX_005)
  if (btnInjectDocx) {
    btnInjectDocx.addEventListener('click', async () => {
      const rawText = editorTextarea.value.trim();
      if (!rawText) {
        showToast('Vui lòng nhập mã Wireframe trước khi chèn vào Word', 'error');
        return;
      }

      let payloadSpec = rawText;
      if (activeMode === 'json') {
        try {
          payloadSpec = JSON.parse(rawText);
        } catch (err) {
          showToast(`Lỗi cú pháp JSON: ${err.message}`, 'error');
          return;
        }
      }

      btnInjectDocx.disabled = true;
      btnInjectDocx.innerHTML = '<i data-lucide="refresh-cw" class="spin"></i> Đang nhúng vào Word...';
      refreshLucideIcons();

      try {
        let uploadedDocxPath = null;
        if (docxFileInput && docxFileInput.files && docxFileInput.files[0]) {
          const upRes = await uploadFileToServer(docxFileInput.files[0]);
          uploadedDocxPath = upRes.filepath || upRes.filename;
        }

        const heading = docxHeadingInput ? docxHeadingInput.value.trim() : '';
        const caption = docxCaptionInput ? docxCaptionInput.value.trim() : '';

        const res = await fetch('/api/diagram/wireframe/inject-doc', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            spec: payloadSpec,
            docx_path: uploadedDocxPath,
            heading: heading || undefined,
            caption: caption || undefined,
            max_width_cm: 14.0
          })
        });

        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Lỗi khi chèn Wireframe vào tài liệu Word');
        }

        showToast(
          `Đã nhúng Wireframe vào Word (${data.width_cm}×${data.height_cm} cm, chuẩn lề in ERR_DOCX_005)!`,
          'success'
        );
        if (data.download_url) {
          window.open(data.download_url, '_blank');
        }
        if (loadHistory) loadHistory();
      } catch (err) {
        showToast(`Lỗi nhúng Word: ${err.message}`, 'error');
      } finally {
        btnInjectDocx.disabled = false;
        btnInjectDocx.innerHTML = '<i data-lucide="file-text"></i> Chèn 1-Click Vào Word BRD (.docx)';
        refreshLucideIcons();
      }
    });
  }

  // Initialize templates and default preview
  populateTemplateDropdown();
  editorTextarea.value = currentSaltStr;
  fetchTemplates().then(() => {
    applyTemplate('login_auth', false);
  });
}

