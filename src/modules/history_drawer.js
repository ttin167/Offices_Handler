/**
 * src/modules/history_drawer.js — History Drawer & Recent Files Controller
 */

export function initHistoryDrawer({ formatBytes, refreshLucideIcons, showToast }) {
  const btnToggleHistory = document.getElementById('btn-toggle-history');
  const btnCloseHistory = document.getElementById('btn-close-history');
  const drawerBackdrop = document.getElementById('drawer-backdrop');
  const historyDrawer = document.getElementById('history-drawer');
  const btnClearHistory = document.getElementById('btn-clear-history');
  const historyListContainer = document.getElementById('history-list-container');
  const historyBadgeCount = document.getElementById('history-badge-count');

  function openHistoryDrawer() {
    if (historyDrawer) historyDrawer.classList.add('open');
    if (drawerBackdrop) drawerBackdrop.classList.add('active');
    loadHistory();
  }

  function closeHistoryDrawer() {
    if (historyDrawer) historyDrawer.classList.remove('open');
    if (drawerBackdrop) drawerBackdrop.classList.remove('active');
  }

  if (btnToggleHistory) btnToggleHistory.addEventListener('click', openHistoryDrawer);
  if (btnCloseHistory) btnCloseHistory.addEventListener('click', closeHistoryDrawer);
  if (drawerBackdrop) drawerBackdrop.addEventListener('click', closeHistoryDrawer);

  function getTaskTypeIcon(taskType) {
    switch (taskType) {
      case 'convert': return '<i data-lucide="refresh-cw"></i>';
      case 'docx_to_pptx': return '<i data-lucide="presentation"></i>';
      case 'pdf_to_pptx': return '<i data-lucide="file-check"></i>';
      case 'diagram_render': return '<i data-lucide="network"></i>';
      case 'math_export': return '<i data-lucide="sigma"></i>';
      case 'mutate_excel': return '<i data-lucide="table"></i>';
      default: return '<i data-lucide="file-text"></i>';
    }
  }

  async function loadHistory() {
    try {
      const res = await fetch('/api/history');
      if (!res.ok) return;
      const data = await res.json();
      const items = data.items || [];

      if (historyBadgeCount) {
        historyBadgeCount.textContent = items.length;
      }

      if (!historyListContainer) return;

      if (items.length === 0) {
        historyListContainer.innerHTML = `
          <div class="history-empty" id="history-empty-state">
            <span class="empty-icon"><i data-lucide="folder-open"></i></span>
            <p>Chưa có tệp nào được xử lý gần đây.</p>
          </div>
        `;
        refreshLucideIcons();
        return;
      }

      historyListContainer.innerHTML = items.map((item) => {
        const icon = getTaskTypeIcon(item.task_type);
        const sizeStr = item.file_size > 0 ? formatBytes(item.file_size) : '';
        const dateStr = item.created_at || '';
        const downloadUrl = item.download_url || `/api/download/${item.output_filename}`;

        return `
          <div class="history-item-card" data-id="${item.id}">
            <div class="history-item-info">
              <div class="history-item-title" title="${item.output_filename}">
                ${icon} ${item.output_filename}
              </div>
              <div class="history-item-meta">
                <span class="history-item-type-badge">${item.task_type}</span>
                ${sizeStr ? `<span>${sizeStr}</span> • ` : ''}
                <span>${dateStr}</span>
              </div>
            </div>
            <div class="history-item-ops">
              <a href="${downloadUrl}" class="btn-hist-action" title="Tải xuống tệp" download="${item.output_filename}"><i data-lucide="download"></i></a>
              <button class="btn-hist-action danger btn-delete-single-history" data-id="${item.id}" title="Xóa lịch sử"><i data-lucide="trash-2"></i></button>
            </div>
          </div>
        `;
      }).join('');
      refreshLucideIcons();

      // Attach delete handlers
      const deleteButtons = historyListContainer.querySelectorAll('.btn-delete-single-history');
      deleteButtons.forEach((btn) => {
        btn.addEventListener('click', async (e) => {
          e.stopPropagation();
          const id = btn.getAttribute('data-id');
          if (!id) return;
          try {
            const delRes = await fetch(`/api/history/${id}`, { method: 'DELETE' });
            if (delRes.ok) {
              loadHistory();
            }
          } catch (err) {
            console.error('Lỗi khi xóa mục lịch sử:', err);
          }
        });
      });

    } catch (err) {
      console.error('Lỗi khi tải lịch sử:', err);
    }
  }

  if (btnClearHistory) {
    btnClearHistory.addEventListener('click', async () => {
      if (!confirm('Bạn có chắc chắn muốn xóa toàn bộ lịch sử không?')) return;
      try {
        const res = await fetch('/api/history', { method: 'DELETE' });
        if (res.ok) {
          loadHistory();
          showToast('Đã xóa sạch lịch sử chuyển đổi!', 'info');
        }
      } catch (err) {
        showToast('Không thể xóa lịch sử', 'error');
      }
    });
  }

  // Load history count on startup
  loadHistory();

  return {
    openHistoryDrawer,
    closeHistoryDrawer,
    loadHistory,
  };
}
