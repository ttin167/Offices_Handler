/**
 * src/modules/diagram_studio.js ? Technical Diagram Studio Controller (Mermaid, PlantUML, Canvas, Draw.io)
 */

import { DIAGRAM_PRESETS } from './constants.js';

export function initDiagramStudio({ state, showToast, refreshLucideIcons, loadHistory }) {
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

}
