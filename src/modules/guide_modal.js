/**
 * src/modules/guide_modal.js — Interactive User Guide & Documentation Controller.
 *
 * Provides a searchable, categorized, in-app documentation center
 * covering all modules: Document Hub, Spreadsheet Studio, Diagram Studio,
 * Presentation Studio, Enterprise Printing, BA Enterprise Hub, Labor Law & Payroll Vault,
 * MCP & CLI Tooling, and System Performance Tips.
 */

import guideModalHtml from '../templates/modals/user_guide_modal.html?raw';

export function initGuideModal(ctx) {
  const { showToast, refreshLucideIcons } = ctx;

  let modal = document.getElementById('modal-user-guide');
  if (!modal) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = guideModalHtml.trim();
    modal = wrapper.querySelector('#modal-user-guide') || wrapper.firstElementChild;
    const host = document.getElementById('app') || document.body;
    if (modal && host) host.appendChild(modal);
  }

  const btnOpen = document.getElementById('btn-open-guide-modal');
  const btnClose = document.getElementById('btn-close-guide-modal');
  const searchInput = document.getElementById('guide-search-input');
  const categoryBtns = document.querySelectorAll('.guide-nav-item');
  const guidePanels = document.querySelectorAll('.guide-content-panel');

  if (!btnOpen || !modal) return;

  function openModal() {
    modal.style.display = 'flex';
    refreshLucideIcons();
  }

  function closeModal() {
    modal.style.display = 'none';
  }

  btnOpen.addEventListener('click', openModal);
  if (btnClose) btnClose.addEventListener('click', closeModal);

  // Close when clicking modal backdrop
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  // Category switching
  categoryBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetCat = btn.getAttribute('data-guide-cat');
      categoryBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');

      guidePanels.forEach((panel) => {
        if (panel.getAttribute('data-guide-cat') === targetCat) {
          panel.classList.add('active');
        } else {
          panel.classList.remove('active');
        }
      });
      refreshLucideIcons();
    });
  });

  // Real-time search filter
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const q = (e.target.value || '').trim().toLowerCase();
      if (!q) {
        // Restore active category panel
        const activeCatBtn = document.querySelector('.guide-nav-item.active');
        const activeCat = activeCatBtn ? activeCatBtn.getAttribute('data-guide-cat') : 'quickstart';
        guidePanels.forEach((p) => {
          p.classList.toggle('active', p.getAttribute('data-guide-cat') === activeCat);
          // Show all items inside panels
          p.querySelectorAll('.guide-card, .guide-section').forEach((card) => {
            card.style.display = '';
          });
        });
        return;
      }

      // In search mode, search across all cards in all panels
      guidePanels.forEach((panel) => {
        let panelHasMatch = false;
        panel.querySelectorAll('.guide-card, .guide-section').forEach((card) => {
          const text = card.textContent.toLowerCase();
          const matches = text.includes(q);
          card.style.display = matches ? '' : 'none';
          if (matches) panelHasMatch = true;
        });
        panel.classList.toggle('active', panelHasMatch);
      });
      refreshLucideIcons();
    });
  }

  // Quick Copy to Clipboard for code blocks & commands
  modal.querySelectorAll('.guide-copy-btn').forEach((btn) => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const codeSnippet = btn.getAttribute('data-copy');
      if (!codeSnippet) return;
      try {
        await navigator.clipboard.writeText(codeSnippet);
        const originalHtml = btn.innerHTML;
        btn.innerHTML = '<i data-lucide="check" style="color:#22c55e;"></i> Đã Copy';
        refreshLucideIcons();
        if (showToast) showToast('Đã sao chép lệnh vào clipboard!', 'success', 2000);
        setTimeout(() => {
          btn.innerHTML = originalHtml;
          refreshLucideIcons();
        }, 2000);
      } catch {
        if (showToast) showToast('Không thể sao chép tự động', 'error');
      }
    });
  });
}

