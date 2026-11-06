/**
 * src/modules/print_studio.js — Enterprise Printing Studio Controller.
 *
 * Interacts with:
 * - GET  /api/print/printers (hardware/virtual printer enumeration)
 * - GET  /api/print/status (hardware queue, paper-out, and jam detection)
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
  const filePathInput = document.getElementById('print-file-path');
  const printerSelect = document.getElementById('print-select-printer');
  const statusPill = document.getElementById('printer-status-pill');

  const rangeAllRadio = document.getElementById('print-range-all');
  const rangeCustomRadio = document.getElementById('print-range-custom');
  const rangeInput = document.getElementById('print-range-input');

  const copiesInput = document.getElementById('print-copies');
  const pageSizeSelect = document.getElementById('print-select-page-size');
  const duplexSelect = document.getElementById('print-select-duplex');
  const scalingSelect = document.getElementById('print-select-scaling');

  const nupSelect = document.getElementById('print-nup');
  const checkGrayscale = document.getElementById('print-opt-grayscale');
  const checkBooklet = document.getElementById('print-opt-booklet');
  const watermarkInput = document.getElementById('print-opt-watermark');
  const statusArea = document.getElementById('print-job-status-area');

  if (!modal) return;

  // Toggle Page Range Input
  if (rangeAllRadio && rangeCustomRadio && rangeInput) {
    rangeAllRadio.addEventListener('change', () => {
      if (rangeAllRadio.checked) {
        rangeInput.disabled = true;
        rangeInput.style.opacity = '0.6';
      }
    });
    rangeCustomRadio.addEventListener('change', () => {
      if (rangeCustomRadio.checked) {
        rangeInput.disabled = false;
        rangeInput.style.opacity = '1.0';
        rangeInput.focus();
      }
    });
  }

  // Update Printer Hardware Status Pill
  async function updatePrinterStatus(printerName) {
    if (!statusPill) return;
    statusPill.className = 'printer-status-pill ready';
    statusPill.textContent = '● Đang kiểm tra...';
    try {
      const url = printerName
        ? `/api/print/status?printer=${encodeURIComponent(printerName)}`
        : '/api/print/status';
      const res = await fetch(url).then((r) => r.json());
      if (res.success) {
        if (res.status === 'PAPER_OUT') {
          statusPill.className = 'printer-status-pill paperout';
          statusPill.textContent = '● Hết Giấy (Paper Out)';
        } else if (res.status === 'PAPER_JAM') {
          statusPill.className = 'printer-status-pill paperout';
          statusPill.textContent = '● Kẹt Giấy (Jam)';
        } else if (res.status === 'OFFLINE') {
          statusPill.className = 'printer-status-pill offline';
          statusPill.textContent = '● Ngoại Tuyến (Offline)';
        } else if (res.status === 'DOOR_OPEN') {
          statusPill.className = 'printer-status-pill warning';
          statusPill.textContent = '● Nắp Máy Mở';
        } else {
          statusPill.className = 'printer-status-pill ready';
          const queueText = res.c_jobs > 0 ? ` (${res.c_jobs} job)` : '';
          statusPill.textContent = `● Sẵn Sàng${queueText}`;
        }
      }
    } catch {
      statusPill.className = 'printer-status-pill warning';
      statusPill.textContent = '● Sẵn Sàng (Virtual)';
    }
  }

  function openModal() {
    modal.style.display = 'flex';
    refreshLucideIcons();
    loadPrinters();
  }

  function closeModal() {
    modal.style.display = 'none';
    if (statusArea) statusArea.innerHTML = '';
  }

  if (btnOpen) btnOpen.addEventListener('click', openModal);
  if (btnClose) btnClose.addEventListener('click', closeModal);
  if (btnCancel) btnCancel.addEventListener('click', closeModal);

  // Close modal when clicking on backdrop
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  if (printerSelect) {
    printerSelect.addEventListener('change', () => {
      updatePrinterStatus(printerSelect.value);
    });
  }

  async function loadPrinters() {
    if (!printerSelect) return;
    try {
      const res = await fetch('/api/print/printers').then((r) => r.json());
      if (res.success && res.printers) {
        printerSelect.innerHTML = '';
        let selectedName = '';
        res.printers.forEach((p) => {
          const opt = document.createElement('option');
          opt.value = p.name;
          const defMark = p.is_default ? ' ★ (Mặc định)' : '';
          const typeMark = p.is_virtual ? ' [Virtual]' : ' [Hardware]';
          opt.textContent = `${p.name}${defMark}${typeMark}`;
          if (p.is_default) {
            opt.selected = true;
            selectedName = p.name;
          }
          printerSelect.appendChild(opt);
        });
        updatePrinterStatus(selectedName || printerSelect.value);
      } else {
        printerSelect.innerHTML = '<option value="">Không tìm thấy máy in hệ thống</option>';
        if (statusPill) {
          statusPill.className = 'printer-status-pill offline';
          statusPill.textContent = '● Không tìm thấy';
        }
      }
    } catch (err) {
      printerSelect.innerHTML = `<option value="">Lỗi tải danh sách máy in: ${err.message}</option>`;
    }
  }

  if (btnExecute) {
    btnExecute.addEventListener('click', async () => {
      const file = fileInput?.files?.[0];
      const directFilePath = filePathInput?.value?.trim();

      if (!file && !directFilePath) {
        showToast('Vui lòng chọn tệp tin cần in (.pdf, .docx, .xlsx)', 'warning');
        return;
      }

      // Validate Page Range input
      let pageRange = 'all';
      if (rangeCustomRadio?.checked) {
        const customVal = rangeInput?.value?.trim();
        if (!customVal) {
          showToast('Vui lòng nhập dải trang in (ví dụ: 1-5, 8, 11-20)', 'warning');
          rangeInput?.focus();
          return;
        }
        pageRange = customVal;
      }

      const selectedPrinter = printerSelect?.value;
      const copies = parseInt(copiesInput?.value || '1', 10);
      const pageSize = pageSizeSelect?.value || 'a4';
      const duplex = duplexSelect?.value || 'vertical';
      const scaling = scalingSelect?.value || 'fit';

      const nUp = parseInt(nupSelect?.value || '1', 10);
      const grayscale = checkGrayscale?.checked || false;
      const booklet = checkBooklet?.checked || false;
      const watermark = watermarkInput?.value || '';

      btnExecute.disabled = true;
      btnExecute.innerHTML = '<i data-lucide="loader-2" class="spin"></i> Đang Tạo Luồng In...';
      refreshLucideIcons();

      try {
        let printTargetFile = directFilePath;

        // Step 1: Upload file if chosen via file input
        if (file) {
          if (statusArea) {
            statusArea.innerHTML = `
              <div style="background: rgba(56, 189, 248, 0.1); border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 8px; padding: 10px; color: #38bdf8; display: flex; align-items: center; gap: 8px;">
                <i data-lucide="loader-2" class="spin" style="width: 14px; height: 14px;"></i>
                <span>Bước 1/3: Đang tải tệp tin lên máy chủ...</span>
              </div>
            `;
            refreshLucideIcons();
          }
          showToast('Đang tải tệp tin lên máy chủ...', 'info');
          const uploadRes = await uploadFileToServer(file);
          if (!uploadRes.success) throw new Error(uploadRes.error || 'Upload tệp in thất bại');
          printTargetFile = uploadRes.filepath || uploadRes.filename;
        }

        // Step 2: Processing & Spooling
        if (statusArea) {
          statusArea.innerHTML = `
            <div style="background: rgba(56, 189, 248, 0.1); border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 8px; padding: 10px; color: #38bdf8; display: flex; align-items: center; gap: 8px;">
              <i data-lucide="loader-2" class="spin" style="width: 14px; height: 14px;"></i>
              <span>Bước 2/3: Đang tối ưu hóa trang in (Range, Duplex, Khổ giấy) & gửi Spooler...</span>
            </div>
          `;
          refreshLucideIcons();
        }
        showToast('Đang tối ưu hóa trang in & gửi đến Spooler...', 'info');

        const printRes = await fetch('/api/print/job', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file: printTargetFile,
            printer: selectedPrinter,
            copies,
            page_range: pageRange,
            page_size: pageSize,
            duplex,
            scaling,
            n_up: nUp,
            grayscale,
            booklet,
            watermark,
          }),
        }).then((r) => r.json());

        // Step 3: Result & Spooler Job Tracker
        if (printRes.success) {
          showToast(`Đã gửi lệnh in (Job ID: #${printRes.job_id}) tới máy in thành công`, 'success');

          let warnHtml = '';
          if (printRes.hardware_warning) {
            warnHtml = `
              <div style="margin-top: 8px; background: rgba(239, 68, 68, 0.15); border: 1px solid #ef4444; border-radius: 6px; padding: 8px; color: #f87171;">
                <strong>⚠️ Cảnh Báo Máy In:</strong> ${printRes.hardware_warning}
              </div>
            `;
          }

          const optInfo = printRes.optimization || {};
          const pagesInfo = optInfo.optimized_pages
            ? ` (${optInfo.optimized_pages} trang xuất xưởng)`
            : '';

          if (statusArea) {
            statusArea.innerHTML = `
              <div style="background: rgba(16, 185, 129, 0.12); border: 1px solid #10b981; border-radius: 8px; padding: 12px; color: #10b981;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                  <strong>✓ Lệnh In Đã Khởi Tạo:</strong>
                  <span style="font-family: monospace; background: rgba(16, 185, 129, 0.2); padding: 2px 6px; border-radius: 4px;">#${printRes.job_id}</span>
                </div>
                <div style="font-size: 11.5px; color: #cbd5e1; line-height: 1.5;">
                  Máy in: <strong>${printRes.printer}</strong> (${printRes.status}) • Khổ: <strong>${(printRes.page_size || 'A4').toUpperCase()}</strong> • Trang: <strong>${printRes.page_range}</strong>${pagesInfo} • Duplex: <strong>${printRes.duplex}</strong>
                </div>
                ${warnHtml}
              </div>
            `;
          }
          // Refresh printer status pill after print dispatch
          updatePrinterStatus(selectedPrinter);
        } else {
          showToast(`Lỗi in ấn: ${printRes.error}`, 'error');
          if (statusArea) {
            statusArea.innerHTML = `
              <div style="background: rgba(239, 68, 68, 0.15); border: 1px solid #ef4444; border-radius: 8px; padding: 12px; color: #ef4444;">
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

