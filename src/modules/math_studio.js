/**
 * src/modules/math_studio.js — Math & Equation Studio Controller
 */

export function initMathStudio({ showToast, refreshLucideIcons, loadHistory }) {
  const mathDocTitle = document.getElementById('math-doc-title');
  const mathPresetSelect = document.getElementById('math-preset-select');
  const mathLatexInput = document.getElementById('math-latex-input');
  const btnExportMathDocx = document.getElementById('btn-export-math-docx');
  const previewPaperTitle = document.getElementById('preview-paper-title');
  const previewPaperContent = document.getElementById('preview-paper-content');
  const mathDownloadActions = document.getElementById('math-download-actions');
  const btnDownloadMathDocx = document.getElementById('btn-download-math-docx');
  const symbolButtons = document.querySelectorAll('.btn-symbol');

  const MATH_PRESETS = {
    kinematics: `Vận tốc dài và gia tốc tiếp tuyến của cơ cấu truyền động được mô tả bởi:
$$v(t) = v_{ref} + a \\cdot t$$
Trong đó vận tốc góc thỏa mãn quan hệ:
$$\\omega = 2\\pi f = \\frac{v}{r}$$
Với góc lệch cực đại $\\Delta \\theta \\le 5^{\\degree}$ và gia tốc trọng trường $g \\approx 9.81\\text{ m/s}^{2}$.`,

    physics: `Công suất tiêu thụ danh định của động cơ điện:
$$P = U \\cdot I \\cdot \\cos(\\phi)$$
Năng lượng tích lũy trong cuộn cảm và tụ điện:
$$W_L = \\frac{1}{2} L \\cdot I^2$$
$$W_C = \\frac{1}{2} C \\cdot U^2$$
Với dung sai điện áp cho phép $\\pm 5\\%$.`,

    calculus: `Đạo hàm bậc nhất theo định nghĩa giới hạn vi phân:
$$f'(x) = \\lim_{\\Delta x \\to 0} \\frac{f(x + \\Delta x) - f(x)}{\\Delta x}$$
Tích phân diện tích dưới đường cong chuẩn hóa:
$$S = \\int_{a}^{b} \\sqrt{1 + [f'(x)]^2} \\, dx$$`,

    circle: `Diện tích và chu vi hình tròn bán kính $R$:
$$S = \\pi R^{2}$$
$$C = 2\\pi R$$
Bán kính hiệu dụng đối với tiết diện vành khuyên:
$$R_{eff} = \\sqrt{R_1^2 - R_2^2}$$`,
  };

  // Helper to replace LaTeX tokens with HTML elements
  function formatLatexToHtml(latex) {
    if (!latex) return '';

    const replacements = [
      [/\\alpha/g, 'α'],
      [/\\beta/g, 'β'],
      [/\\gamma/g, 'γ'],
      [/\\delta/g, 'δ'],
      [/\\Delta/g, 'Δ'],
      [/\\omega/g, 'ω'],
      [/\\Omega/g, 'Ω'],
      [/\\pi/g, 'π'],
      [/\\sigma/g, 'σ'],
      [/\\theta/g, 'θ'],
      [/\\mu/g, 'μ'],
      [/\\lambda/g, 'λ'],
      [/\\phi/g, 'ϕ'],
      [/\\Phi/g, 'Φ'],
      [/\\approx/g, '≈'],
      [/\\le/g, '≤'],
      [/\\ge/g, '≥'],
      [/\\pm/g, '±'],
      [/\\cdot/g, '·'],
      [/\\degree/g, '°'],
      [/\\lim/g, 'lim'],
      [/\\cos/g, 'cos'],
      [/\\sin/g, 'sin'],
      [/\\int/g, '∫'],
      [/\\to/g, '→'],
    ];

    let s = latex;

    // Apply token replacements first so symbols inside fractions or sub/sup are converted
    for (const [pattern, repl] of replacements) {
      s = s.replace(pattern, repl);
    }

    // Text: \text{...}
    s = s.replace(/\\text\{([^{}]+)\}/g, '<span style="font-family: var(--font-sans); font-style: normal;">$1</span>');

    // Fractions: \frac{num}{den}
    s = s.replace(/\\frac\{([^{}]+)\}\{([^{}]+)\}/g, (match, num, den) => {
      return `<span class="math-fraction-view"><span class="math-frac-num">${formatLatexToHtml(num)}</span><span class="math-frac-den">${formatLatexToHtml(den)}</span></span>`;
    });

    // Square root: \sqrt{rad}
    s = s.replace(/\\sqrt\{([^{}]+)\}/g, (match, rad) => {
      return `<span class="math-sqrt-view"><span class="math-sqrt-sym">√</span><span class="math-sqrt-rad">${formatLatexToHtml(rad)}</span></span>`;
    });

    // Superscripts: ^{...} or ^x
    s = s.replace(/\^\{([^{}]+)\}/g, '<sup>$1</sup>');
    s = s.replace(/\^([0-9a-zA-Z]+)/g, '<sup>$1</sup>');

    // Subscripts: _{...} or _x
    s = s.replace(/_\{([^{}]+)\}/g, '<sub>$1</sub>');
    s = s.replace(/_([0-9a-zA-Z]+)/g, '<sub>$1</sub>');

    return s;
  }

  function updateMathPreview() {
    if (previewPaperTitle && mathDocTitle) {
      previewPaperTitle.textContent = mathDocTitle.value.trim() || 'Báo Cáo Kỹ Thuật & Công Thức Toán';
    }

    if (!previewPaperContent || !mathLatexInput) return;

    const raw = mathLatexInput.value.trim();
    if (!raw) {
      previewPaperContent.innerHTML = '<p class="math-paragraph" style="color: #94a3b8; font-style: italic;">Nhập văn bản hoặc công thức toán vào ô bên trái để xem trước tại đây...</p>';
      return;
    }

    const lines = raw.split(/\r?\n/);
    let html = '';

    lines.forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed) {
        html += '<p style="margin-bottom: 8px;"></p>';
        return;
      }

      // Check for display math: $$...$$
      if (trimmed.startsWith('$$') && trimmed.endsWith('$$') && trimmed.length >= 4) {
        const formula = trimmed.slice(2, -2).trim();
        html += `<div class="math-display-block">${formatLatexToHtml(formula)}</div>`;
      } else {
        // Inline math parsing $...$
        let renderedLine = trimmed.replace(/\$([^$]+)\$/g, (match, formula) => {
          return `<span class="math-inline-elem">${formatLatexToHtml(formula)}</span>`;
        });
        html += `<p class="math-paragraph">${renderedLine}</p>`;
      }
    });

    previewPaperContent.innerHTML = html;
  }

  // Symbol palette click -> insert at textarea cursor
  symbolButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const insertText = btn.getAttribute('data-insert');
      if (!insertText || !mathLatexInput) return;

      const start = mathLatexInput.selectionStart;
      const end = mathLatexInput.selectionEnd;
      const cur = mathLatexInput.value;

      mathLatexInput.value = cur.substring(0, start) + insertText + cur.substring(end);
      mathLatexInput.focus();
      mathLatexInput.selectionStart = mathLatexInput.selectionEnd = start + insertText.length;

      updateMathPreview();
    });
  });

  if (mathLatexInput) {
    mathLatexInput.addEventListener('input', updateMathPreview);
  }

  if (mathDocTitle) {
    mathDocTitle.addEventListener('input', updateMathPreview);
  }

  if (mathPresetSelect && mathLatexInput) {
    mathPresetSelect.addEventListener('change', () => {
      const preset = MATH_PRESETS[mathPresetSelect.value];
      if (preset) {
        mathLatexInput.value = preset;
        updateMathPreview();
      }
    });

    // Initialize with kinematics preset if empty
    if (!mathLatexInput.value) {
      mathLatexInput.value = MATH_PRESETS.kinematics;
    }
  }

  // Initial render of preview
  updateMathPreview();

  // 1-Click Export to DOCX
  if (btnExportMathDocx) {
    btnExportMathDocx.addEventListener('click', async () => {
      const latex = mathLatexInput ? mathLatexInput.value.trim() : '';
      const title = mathDocTitle ? mathDocTitle.value.trim() : 'Công Thức Toán Học';

      if (!latex) {
        showToast('Vui lòng nhập công thức toán trước khi xuất!', 'error');
        return;
      }

      btnExportMathDocx.disabled = true;
      btnExportMathDocx.innerHTML = '<span>⏳</span><span>ĐANG BIÊN DỊCH WORD OMML...</span>';

      try {
        const res = await fetch('/api/math/export', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            title: title,
            latex: latex,
          }),
        });

        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error || `Xuất Word thất bại (${res.status})`);
        }

        const data = await res.json();
        const downloadUrl = data.download_url || `/api/download/${data.filename}`;

        if (btnDownloadMathDocx) {
          btnDownloadMathDocx.href = downloadUrl;
          btnDownloadMathDocx.download = data.filename;
        }

        if (mathDownloadActions) {
          mathDownloadActions.style.display = 'flex';
        }

        showToast(`Xuất tệp Word thành công: ${data.filename}`, 'success');

        // Automatically trigger file download
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = data.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);

        // Update history drawer
        if (typeof loadHistory === 'function') {
          loadHistory();
        }

      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error', 6000);
      } finally {
        btnExportMathDocx.disabled = false;
        btnExportMathDocx.innerHTML = '<span class="btn-icon"><i data-lucide="download"></i></span><span class="btn-text">Xuất Tệp Word (.docx) Chuẩn OMML</span>';
        refreshLucideIcons();
      }
    });
  }
}
