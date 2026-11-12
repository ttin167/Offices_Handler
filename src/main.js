import {

  createIcons,

  Sparkles,

  Presentation,

  Building2,

  Video,

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
  Minus,
  Book,
  Terminal,
  Archive,
  Scale,
  SearchX,
} from 'lucide';






import { initHistoryDrawer } from './modules/history_drawer.js';
import { initWireframeStudio } from './modules/wireframe_studio.js';
import { initPresentationStudio } from './modules/presentation_studio.js';
import { initSpreadsheetStudio } from './modules/spreadsheet_studio.js';
import { initDocumentHub } from './modules/document_hub.js';
import { initDiagramStudio } from './modules/diagram_studio.js';
import { initFormulaPlaybook } from './modules/formula_playbook.js';
import { initPrintStudio } from './modules/print_studio.js';
import { initGuideModal } from './modules/guide_modal.js';
import { initDocumentViewerModal } from './modules/document_viewer_modal.js';
import { initVisualFieldMapper } from './modules/visual_field_mapper.js';
import { initAiCopilotSidebar } from './modules/ai_copilot_sidebar.js';

const optionalBaModules = import.meta.glob('./modules/ba_bootstrap.js', { eager: true });



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
  Minus,
  Book,
  Terminal,
  Archive,
  Scale,
  SearchX,
};



/**

 * ANTIGRAVITY OFFICE STUDIO — CLIENT LOGIC ORCHESTRATOR

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

    activeTab: 'documents',

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

  // 4. NAVIGATION TABS (Single Source of Truth Router)

  // =========================================================================

  window.__ag_main_router_active = true;

  const mainNav = document.getElementById('main-nav');

  const panels = {
    wireframe: document.getElementById('panel-wireframe'),
    documents: document.getElementById('panel-documents'),
    spreadsheet: document.getElementById('panel-spreadsheet'),
    diagrams: document.getElementById('panel-diagrams'),
    presentation: document.getElementById('panel-presentation'),
  };

  if (mainNav) {
    mainNav.addEventListener('click', (e) => {
      const tab = e.target.closest('.nav-tab');
      if (!tab) return;
      const targetTab = tab.getAttribute('data-tab');
      const targetPanel = document.getElementById(`panel-${targetTab}`) || panels[targetTab];
      if (!targetPanel || tab.classList.contains('active')) return;

      document.querySelectorAll('.nav-tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');

      document.querySelectorAll('.tab-panel').forEach((p) => {
        if (p !== targetPanel && p.classList.contains('active')) {
          p.classList.remove('active');
        }
      });

      requestAnimationFrame(() => {
        targetPanel.classList.add('active');
        state.activeTab = targetTab;
        if (window.appStore && typeof window.appStore.set === 'function') {
          window.appStore.set('activeTab', targetTab);
        }
        if (typeof refreshLucideIcons === 'function') {
          refreshLucideIcons();
        }
      });
    });
  }



  // =========================================================================

  // 5. SHARED HELPERS: UPLOAD FILE & FORMAT BYTES

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

  // 6. INITIALIZE MODULAR STUDIO CONTROLLERS

  // =========================================================================

  const { loadHistory } = initHistoryDrawer({

    formatBytes,

    refreshLucideIcons,

    showToast,

  });



  const docViewer = initDocumentViewerModal({ state, showToast, refreshLucideIcons });
  const visualFieldMapper = initVisualFieldMapper({ state, showToast, refreshLucideIcons, uploadFileToServer });
  window.docViewer = docViewer;
  window.visualFieldMapper = visualFieldMapper;

  const sharedCtx = {

    state,

    showToast,

    refreshLucideIcons,

    uploadFileToServer,

    formatBytes,

    loadHistory,

    docViewer,

    visualFieldMapper,

  };



  const baBootstrapMod = Object.values(optionalBaModules)[0];
  if (baBootstrapMod && typeof baBootstrapMod.initBABootstrap === 'function') {
    baBootstrapMod.initBABootstrap({ state, sharedCtx });
  }
  initWireframeStudio(sharedCtx);
  initPresentationStudio(sharedCtx);
  initSpreadsheetStudio(sharedCtx);
  initFormulaPlaybook({ showToast, refreshLucideIcons });
  initDocumentHub(sharedCtx);
  initDiagramStudio(sharedCtx);
  initPrintStudio(sharedCtx);
  initGuideModal(sharedCtx);
  initAiCopilotSidebar(sharedCtx);



  // Global Quick Copy Handler for Code Previews
  document.addEventListener('click', (e) => {
    const copyBtn = e.target.closest('[data-copy-target]');
    if (!copyBtn) return;
    const targetId = copyBtn.getAttribute('data-copy-target');
    const targetEl = targetId ? document.getElementById(targetId) : null;
    if (!targetEl) return;
    const textToCopy = (targetEl.value !== undefined ? targetEl.value : targetEl.textContent) || '';
    if (!textToCopy.trim()) {
      showToast('Không có nội dung để sao chép.', 'warning');
      return;
    }
    navigator.clipboard.writeText(textToCopy).then(() => {
      showToast('Đã sao chép vào bộ nhớ tạm!', 'success');
    }).catch(() => {
      showToast('Không thể sao chép nội dung.', 'error');
    });
  });

  // Initial Lucide icon generation

  refreshLucideIcons();

});

