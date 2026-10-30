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

import { initHistoryDrawer } from './modules/history_drawer.js';
import { initPresentationStudio } from './modules/presentation_studio.js';
import { initSpreadsheetStudio } from './modules/spreadsheet_studio.js';
import { initDocumentHub } from './modules/document_hub.js';
import { initDiagramStudio } from './modules/diagram_studio.js';
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

  const sharedCtx = {
    state,
    showToast,
    refreshLucideIcons,
    uploadFileToServer,
    formatBytes,
    loadHistory,
  };

  initPresentationStudio(sharedCtx);
  initSpreadsheetStudio(sharedCtx);
  initFormulaPlaybook({ showToast, refreshLucideIcons });
  initDocumentHub(sharedCtx);
  initDiagramStudio(sharedCtx);
  initMathStudio({ showToast, refreshLucideIcons, loadHistory });

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
