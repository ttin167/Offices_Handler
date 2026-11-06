/**
 * src/modules/ai_copilot_sidebar.js — AI Copilot Workspace & Universal FTS5 RAG Sidebar Controller
 * Supports Hybrid LLM Connectors (Local Ollama + Cloud Gemini) and interactive Action Cards.
 */

export function initAiCopilotSidebar({ state, showToast, refreshLucideIcons, uploadFileToServer }) {
  const drawer = document.getElementById('copilot-sidebar-drawer');
  const backdrop = document.getElementById('copilot-drawer-backdrop');
  const btnToggle = document.getElementById('btn-toggle-copilot');
  const btnClose = document.getElementById('btn-close-copilot');
  const providerSelect = document.getElementById('copilot-provider-select');
  const ragStatsBadge = document.getElementById('copilot-rag-stats');
  const btnIndexFile = document.getElementById('btn-copilot-index-file');
  const fileInput = document.getElementById('copilot-rag-file-input');
  const messagesBox = document.getElementById('copilot-messages-list');
  const chatInput = document.getElementById('copilot-chat-input');
  const btnSend = document.getElementById('btn-copilot-send');
  const quickChips = document.querySelectorAll('.copilot-quick-chip');

  if (!drawer) return {};

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function formatAnswerHtml(text) {
    const escaped = escapeHtml(text || '');
    return escaped
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\n/g, '<br/>');
  }

  async function refreshStatus() {
    try {
      const res = await fetch('/api/copilot/status');
      if (!res.ok) return;
      const data = await res.json();
      if (data.rag && ragStatsBadge) {
        const docs = data.rag.total_documents || 0;
        const chunks = data.rag.total_chunks || 0;
        ragStatsBadge.textContent = `FTS5 RAG: ${docs} tài liệu (${chunks} phân đoạn)`;
      }
    } catch (_) {}
  }

  function openCopilot() {
    drawer.classList.add('open');
    if (backdrop) backdrop.classList.add('open');
    refreshStatus();
    if (chatInput) setTimeout(() => chatInput.focus(), 120);
  }

  function closeCopilot() {
    drawer.classList.remove('open');
    if (backdrop) backdrop.classList.remove('open');
  }

  if (btnToggle) {
    btnToggle.addEventListener('click', () => {
      if (drawer.classList.contains('open')) {
        closeCopilot();
      } else {
        openCopilot();
      }
    });
  }

  if (btnClose) btnClose.addEventListener('click', closeCopilot);
  if (backdrop) backdrop.addEventListener('click', closeCopilot);

  function executeActionCard(card) {
    const cmd = card.command_type;
    const payload = card.payload || {};

    if (cmd === 'open_modal' && payload.target) {
      const modal = document.getElementById(payload.target);
      if (modal) {
        modal.style.display = 'flex';
        closeCopilot();
        showToast(`Đã mở: ${card.title}`, 'info');
      }
      return;
    }

    if (cmd === 'switch_tab' && payload.tab) {
      const navBtn = document.querySelector(`.nav-btn[data-tab="${payload.tab}"]`);
      if (navBtn) navBtn.click();
      if (payload.subtab) {
        setTimeout(() => {
          const subBtn = document.querySelector(`[data-doc-subtab="${payload.subtab}"], [data-subtab="${payload.subtab}"]`);
          if (subBtn) subBtn.click();
        }, 80);
      }
      closeCopilot();
      showToast(`Đã chuyển tới: ${card.title}`, 'info');
      return;
    }

    if (cmd === 'open_viewer') {
      const fname = payload.filename || (state && state.docFile ? state.docFile.name : '');
      if (window.openDocumentViewer) {
        window.openDocumentViewer({
          file: state ? state.docFile : null,
          filename: fname || 'Preview_Document.docx',
        });
        closeCopilot();
      }
    }
  }

  function appendMessage({ role, text, provider, model, actionCards = [], citations = [], latencyMs = 0 }) {
    if (!messagesBox) return;

    const msgEl = document.createElement('div');
    msgEl.className = `copilot-msg copilot-msg-${role}`;

    let cardsHtml = '';
    if (actionCards && actionCards.length > 0) {
      cardsHtml = `
        <div class="copilot-cards-list">
          ${actionCards.map((c, idx) => `
            <div class="copilot-action-card" data-card-idx="${idx}">
              <div class="copilot-card-top">
                <span class="copilot-card-badge">${escapeHtml(c.badge || 'Action')}</span>
                <i data-lucide="arrow-up-right" style="width: 14px; height: 14px;"></i>
              </div>
              <div class="copilot-card-title">${escapeHtml(c.title)}</div>
              <div class="copilot-card-desc">${escapeHtml(c.description)}</div>
            </div>
          `).join('')}
        </div>
      `;
    }

    let citationsHtml = '';
    if (citations && citations.length > 0) {
      citationsHtml = `
        <div class="copilot-citations">
          ${citations.map(cit => `
            <span class="copilot-citation-pill">
              <i data-lucide="file-text" style="width: 11px; height: 11px;"></i>
              ${escapeHtml(cit.filename || cit.doc_id)} (${escapeHtml(cit.section_title || 'Section')})
            </span>
          `).join('')}
        </div>
      `;
    }

    const metaHtml = role === 'assistant' && model
      ? `<div class="copilot-msg-meta">${escapeHtml(provider || 'hybrid')} • ${escapeHtml(model)} • ${latencyMs}ms</div>`
      : '';

    msgEl.innerHTML = `
      <div class="copilot-msg-bubble">
        <div class="copilot-msg-text">${formatAnswerHtml(text)}</div>
        ${citationsHtml}
        ${cardsHtml}
        ${metaHtml}
      </div>
    `;

    messagesBox.appendChild(msgEl);
    messagesBox.scrollTop = messagesBox.scrollHeight;
    refreshLucideIcons();

    if (actionCards && actionCards.length > 0) {
      const cardEls = msgEl.querySelectorAll('.copilot-action-card');
      cardEls.forEach(el => {
        el.addEventListener('click', () => {
          const idx = parseInt(el.getAttribute('data-card-idx'), 10);
          if (actionCards[idx]) executeActionCard(actionCards[idx]);
        });
      });
    }
  }

  async function sendPrompt(rawPrompt) {
    const prompt = (rawPrompt || (chatInput ? chatInput.value : '')).trim();
    if (!prompt) return;

    if (chatInput) chatInput.value = '';
    appendMessage({ role: 'user', text: prompt });

    const mode = providerSelect ? providerSelect.value : 'local';
    if (btnSend) btnSend.disabled = true;

    try {
      const res = await fetch('/api/copilot/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt,
          provider: mode,
          top_k: 5,
        }),
      });
      const data = await res.json();
      if (!res.ok || data.error && !data.answer) {
        throw new Error(data.error || 'Không thể kết nối AI Copilot');
      }

      appendMessage({
        role: 'assistant',
        text: data.answer,
        provider: data.provider,
        model: data.model,
        actionCards: data.action_cards || [],
        citations: data.citations || [],
        latencyMs: data.latency_ms || 0,
      });
      refreshStatus();
    } catch (err) {
      appendMessage({
        role: 'assistant',
        text: `Cảnh báo kết nối: ${err.message}`,
      });
    } finally {
      if (btnSend) btnSend.disabled = false;
    }
  }

  if (btnSend) {
    btnSend.addEventListener('click', () => sendPrompt());
  }

  if (chatInput) {
    chatInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendPrompt();
      }
    });
  }

  quickChips.forEach(chip => {
    chip.addEventListener('click', () => {
      const q = chip.getAttribute('data-prompt');
      if (q) sendPrompt(q);
    });
  });

  if (btnIndexFile && fileInput) {
    btnIndexFile.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async (e) => {
      if (!e.target.files || e.target.files.length === 0) return;
      const file = e.target.files[0];
      try {
        showToast(`Đang nạp "${file.name}" vào Universal FTS5 RAG...`, 'info');
        const up = await uploadFileToServer(file);
        const serverPath = up.filepath || up.file_path;
        const res = await fetch('/api/copilot/index', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ file_path: serverPath }),
        });
        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Lỗi lập chỉ mục');
        }
        showToast(`Đã lập chỉ mục ${data.chunks_indexed} phân đoạn từ ${data.filename}!`, 'success');
        refreshStatus();
      } catch (err) {
        showToast(`Lỗi nạp RAG: ${err.message}`, 'error');
      } finally {
        fileInput.value = '';
      }
    });
  }

  return {
    openCopilot,
    closeCopilot,
    sendPrompt,
    refreshStatus,
  };
}

