/**
 * src/modules/print_studio.js — Enterprise Printing Studio Controller.
 *
 * Interacts with:
 * - GET  /api/print/printers (hardware/virtual printer enumeration)
 * - POST /api/print/job (silent printing & pre-print optimization)
 */

export function initPrintStudio(ctx) {
  const { showToast, refreshLucideIcons, uploadFileToServer } = ctx;

  const btnOpen = document.getElementById('btn-open-print-modal');
  const modal = document.getElementById('modal-print');
  const btnClose = document.getElementById('btn-close-print-modal');
  const btnCancel = document.getElementById('btn-cancel-print');
  const btnExecute = document.getElementById('btn-execute-print');

  const fileInput = document.getElementById('print-input-file');
  const printerSelect = document.getElementById('print-select-printer');
  const copiesInput = document.getElementById('print-copies');
  const nupSelect = document.getElementById('print-nup');
  const checkGrayscale = document.getElementById('print-opt-grayscale');
  const checkBooklet = document.getElementById('print-opt-booklet');
  const watermarkInput = document.getElementById('print-opt-watermark');
  const statusArea = document.getElementById('print-job-status-area');

  if (!btnOpen || !modal) return;

  function openModal() {
    modal.style.display = 'flex';
    refreshLucideIcons();
    loadPrinters();
  }

  function closeModal() {
    modal.style.display = 'none';
    if (statusArea) statusArea.innerHTML = '';
  }

  btnOpen.addEventListener('click', openModal);
  if (btnClose) btnClose.addEventListener('click', closeModal);
  if (btnCancel) btnCancel.addEventListener('click', closeModal);

  // Close modal when clicking on backdrop
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  async function loadPrinters() {
    if (!printerSelect) return;
    try {
      const res = await fetch('/api/print/printers').then((r) => r.json());
      if (res.success && res.printers) {
        printerSelect.innerHTML = '';
        res.printers.forEach((p) => {
          const opt = document.createElement('option');
          opt.value = p.name;
          const defMark = p.is_default ? ' ★ (Mặc định)' : '';
          const typeMark = p.is_virtual ? ' [Virtual]' : ' [Hardware]';
          opt.textContent = `${p.name}${defMark}${typeMark}`;
          if (p.is_default) opt.selected = true;
          printerSelect.appendChild(opt);
        });
      } else {
        printerSelect.innerHTML = '<option value="">Không tìm thấy máy in hệ thống</option>';
      }
    } catch (err) {
      printerSelect.innerHTML = `<option value="">Lỗi tải danh sách máy in: ${err.message}</option>`;
    }
  }

  if (btnExecute) {
    btnExecute.addEventListener('click', async () => {
      const file = fileInput?.files[0];
      if (!file) {
        showToast('Vui lòng chọn tệp tin cần in (.pdf, .docx, .xlsx)', 'warning');
        return;
      }

      const selectedPrinter = printerSelect?.value;
      const copies = parseInt(copiesInput?.value || '1', 10);
      const nUp = parseInt(nupSelect?.value || '1', 10);
      const grayscale = checkGrayscale?.checked || false;
      const booklet = checkBooklet?.checked || false;
      const watermark = watermarkInput?.value || '';

      btnExecute.disabled = true;
      btnExecute.innerHTML = '<i data-lucide="loader-2" class="spin"></i> Đang In...';
      refreshLucideIcons();

      try {
        showToast('Đang tải tệp tin lên máy chủ...', 'info');
        const uploadRes = await uploadFileToServer(file);
        if (!uploadRes.success) throw new Error(uploadRes.error || 'Upload tệp in thất bại');

        showToast('Đang xử lý tối ưu hóa trang in & gửi đến Spooler...', 'info');
        const printRes = await fetch('/api/print/job', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file: uploadRes.filepath || uploadRes.filename,
            printer: selectedPrinter,
            copies,
            n_up: nUp,
            grayscale,
            booklet,
            watermark,
          }),
        }).then((r) => r.json());

        if (printRes.success) {
          showToast(`Lệnh in đã được gửi thành công: ${printRes.job_id}`, 'success');
          if (statusArea) {
            statusArea.innerHTML = `
              <div style="background: rgba(16, 185, 129, 0.15); border: 1px solid #10b981; border-radius: 8px; padding: 10px; color: #10b981;">
                <strong>Thành Công:</strong> Lệnh in <code>${printRes.job_id}</code> đã gửi tới <em>${printRes.printer}</em> (${printRes.status}).
              </div>
            `;
          }
        } else {
          showToast(`Lỗi in ấn: ${printRes.error}`, 'error');
          if (statusArea) {
            statusArea.innerHTML = `
              <div style="background: rgba(239, 68, 68, 0.15); border: 1px solid #ef4444; border-radius: 8px; padding: 10px; color: #ef4444;">
                <strong>Lỗi In:</strong> ${printRes.error}
              </div>
            `;
          }
        }
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error');
      } finally {
        btnExecute.disabled = false;
        btnExecute.innerHTML = '<i data-lucide="printer"></i> Bắt Đầu In Ngầm';
        refreshLucideIcons();
      }
    });
  }
}

