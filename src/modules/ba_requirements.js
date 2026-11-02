/**
 * src/modules/ba_requirements.js — Requirements Hub for BA Enterprise Workstation.
 *
 * Implements:
 * 1. Interactive Guided Authoring for User Stories & Gherkin Acceptance Criteria.
 * 2. Industry-standard Templates (Banking, E-commerce, Logistics).
 * 3. Real-Time BABOK Quality Linter & Scorecard.
 * 4. Use Case Specification Form & OpenXML Word Exporter.
 * 5. 1-Click Complete Enterprise BRD / SRS Document Synthesizer (.docx).
 */

export function initRequirementsHub(ctx) {
  const { showToast, refreshLucideIcons, loadHistory } = ctx;

  const panel = document.getElementById('panel-requirements');
  if (!panel) return;

  // State
  let industryTemplates = null;
  let currentStories = [];
  let currentUsecase = null;

  // DOM Elements
  const tplSelect = document.getElementById('ba-template-select');
  const btnLoadTpl = document.getElementById('btn-load-ba-template');
  const btnLintStory = document.getElementById('btn-lint-story');
  const btnExportDocx = document.getElementById('btn-export-story-docx');
  const btnExportMd = document.getElementById('btn-export-story-md');
  const btnGenerateBRD = document.getElementById('btn-generate-brd');

  // Input Fields - Story
  const fldStoryId = document.getElementById('story-id');
  const fldStoryTitle = document.getElementById('story-title');
  const fldStoryDomain = document.getElementById('story-domain');
  const fldStoryRole = document.getElementById('story-role');
  const fldStoryGoal = document.getElementById('story-goal');
  const fldStoryBenefit = document.getElementById('story-benefit');
  const fldStoryRules = document.getElementById('story-rules');
  const scenariosContainer = document.getElementById('scenarios-list');
  const btnAddScenario = document.getElementById('btn-add-scenario');
  const linterResultBox = document.getElementById('story-linter-result');

  // Fetch Templates on Init
  fetch('/api/ba/templates')
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        industryTemplates = data;
        populateTemplateSelect(data.stories);
        // Load default banking template
        if (data.stories && data.stories.banking_fund_transfer) {
          loadStoryIntoForm('US-PAY-001', data.stories.banking_fund_transfer);
        }
      }
    })
    .catch((err) => console.warn('Cannot fetch BA templates:', err));

  function populateTemplateSelect(storiesMap) {
    if (!tplSelect || !storiesMap) return;
    tplSelect.innerHTML = '<option value="">-- Chọn Khung Mẫu Nghiệp Vụ Chuẩn --</option>';
    for (const [key, tpl] of Object.entries(storiesMap)) {
      const opt = document.createElement('option');
      opt.value = key;
      opt.textContent = `[${tpl.domain}] ${tpl.title}`;
      tplSelect.appendChild(opt);
    }
  }

  if (btnLoadTpl) {
    btnLoadTpl.addEventListener('click', () => {
      const key = tplSelect ? tplSelect.value : '';
      if (!key || !industryTemplates || !industryTemplates.stories[key]) {
        showToast('Vui lòng chọn một khung mẫu từ danh sách', 'warning');
        return;
      }
      loadStoryIntoForm('US-001', industryTemplates.stories[key]);
      showToast(`Đã nạp khung mẫu: ${industryTemplates.stories[key].title}`, 'success');
      runStoryLinting();
    });
  }

  function loadStoryIntoForm(id, data) {
    if (fldStoryId) fldStoryId.value = id || 'US-001';
    if (fldStoryTitle) fldStoryTitle.value = data.title || '';
    if (fldStoryDomain) fldStoryDomain.value = data.domain || 'General Business';
    if (fldStoryRole) fldStoryRole.value = data.role || '';
    if (fldStoryGoal) fldStoryGoal.value = data.goal || '';
    if (fldStoryBenefit) fldStoryBenefit.value = data.benefit || '';
    if (fldStoryRules) {
      fldStoryRules.value = (data.business_rules || []).join('\n');
    }

    renderScenarios(data.scenarios || []);
    refreshLucideIcons();
  }

  function renderScenarios(scenarios) {
    if (!scenariosContainer) return;
    scenariosContainer.innerHTML = '';

    scenarios.forEach((sc, idx) => {
      addScenarioCard(sc, idx + 1);
    });
  }

  function addScenarioCard(sc = {}, idx = 1) {
    if (!scenariosContainer) return;
    const card = document.createElement('div');
    card.className = 'scenario-card';
    card.dataset.index = idx;

    card.innerHTML = `
      <div class="scenario-card-header">
        <div class="scenario-badge">Kịch Bản #${idx}</div>
        <select class="scenario-type-select">
          <option value="Happy Path" ${sc.type === 'Happy Path' ? 'selected' : ''}>Happy Path (Luồng Chuẩn)</option>
          <option value="Negative Path" ${sc.type === 'Negative Path' ? 'selected' : ''}>Negative Path (Xử Lý Lỗi)</option>
          <option value="Boundary / Edge Case" ${sc.type && sc.type.includes('Boundary') ? 'selected' : ''}>Boundary / Edge Case (Giới Hạn)</option>
        </select>
        <button class="btn-remove-scenario" title="Xóa kịch bản này"><i data-lucide="trash-2"></i></button>
      </div>
      <div class="scenario-card-body">
        <div class="form-group">
          <label>Tên kịch bản:</label>
          <input type="text" class="sc-title" value="${sc.title || ''}" placeholder="Ví dụ: Chuyển tiền thành công trong hạn mức" />
        </div>
        <div class="form-row-3">
          <div class="form-group">
            <label class="kw-label kw-given">Given (Tiền điều kiện):</label>
            <textarea class="sc-given" rows="2" placeholder="Khách hàng đã đăng nhập và có số dư...">${sc.given || ''}</textarea>
          </div>
          <div class="form-group">
            <label class="kw-label kw-when">When (Hành động):</label>
            <textarea class="sc-when" rows="2" placeholder="Khách hàng nhập số tài khoản và xác thực OTP...">${sc.when || ''}</textarea>
          </div>
          <div class="form-group">
            <label class="kw-label kw-then">Then (Kết quả mong đợi):</label>
            <textarea class="sc-then" rows="2" placeholder="Hệ thống trừ tiền và gửi biên lai...">${sc.then || ''}</textarea>
          </div>
        </div>
      </div>
    `;

    card.querySelector('.btn-remove-scenario').addEventListener('click', () => {
      card.remove();
      refreshLucideIcons();
    });

    scenariosContainer.appendChild(card);
    refreshLucideIcons();
  }

  if (btnAddScenario) {
    btnAddScenario.addEventListener('click', () => {
      const currentCount = scenariosContainer ? scenariosContainer.children.length : 0;
      addScenarioCard({ title: '', type: 'Happy Path', given: '', when: '', then: '' }, currentCount + 1);
    });
  }

  function collectStoryFromForm() {
    const scenarios = [];
    if (scenariosContainer) {
      scenariosContainer.querySelectorAll('.scenario-card').forEach((card) => {
        scenarios.push({
          title: card.querySelector('.sc-title')?.value || 'Kịch bản kiểm thử',
          type: card.querySelector('.scenario-type-select')?.value || 'Happy Path',
          given: card.querySelector('.sc-given')?.value || '',
          when: card.querySelector('.sc-when')?.value || '',
          then: card.querySelector('.sc-then')?.value || '',
        });
      });
    }

    const rules = (fldStoryRules?.value || '')
      .split('\n')
      .map((l) => l.trim())
      .filter((l) => l.length > 0);

    return {
      id: fldStoryId?.value || 'US-001',
      title: fldStoryTitle?.value || 'Chức năng mới',
      domain: fldStoryDomain?.value || 'General Business',
      role: fldStoryRole?.value || '',
      goal: fldStoryGoal?.value || '',
      benefit: fldStoryBenefit?.value || '',
      business_rules: rules,
      scenarios,
    };
  }

  function runStoryLinting() {
    const story = collectStoryFromForm();
    if (!linterResultBox) return;

    linterResultBox.innerHTML = '<span class="loading-spinner"></span> Đang thẩm định chất lượng theo chuẩn BABOK...';

    fetch('/api/ba/story/lint', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ story }),
    })
      .then((r) => r.json())
      .then((data) => {
        if (!data.success) return;
        const audit = data.audit;
        let scoreBadgeClass = audit.score >= 80 ? 'badge-success' : audit.score >= 50 ? 'badge-warning' : 'badge-danger';

        let html = `
          <div class="linter-header">
            <div class="linter-score-pill ${scoreBadgeClass}">
              Điểm Chất Lượng: <strong>${audit.score}/100</strong> (${audit.grade})
            </div>
            <div class="linter-summary">Đã phân tích ${audit.total_scenarios} kịch bản nghiệm thu</div>
          </div>
        `;

        if (audit.findings && audit.findings.length > 0) {
          html += '<ul class="linter-findings">';
          audit.findings.forEach((f) => {
            const icon = f.severity === 'CRITICAL' ? 'alert-triangle' : f.severity === 'WARNING' ? 'alert-circle' : 'info';
            const color = f.severity === 'CRITICAL' ? 'text-danger' : f.severity === 'WARNING' ? 'text-warning' : 'text-info';
            html += `<li class="${color}"><i data-lucide="${icon}"></i> <span>[${f.severity}] ${f.message}</span></li>`;
          });
          html += '</ul>';
        } else {
          html += '<div class="linter-clean"><i data-lucide="check-circle-2"></i> Yêu cầu đã đáp ứng hoàn hảo các chuẩn mực BABOK & Gherkin!</div>';
        }

        linterResultBox.innerHTML = html;
        refreshLucideIcons();
      })
      .catch((err) => {
        linterResultBox.innerHTML = `<span class="text-danger">Lỗi kiểm tra: ${err.message}</span>`;
      });
  }

  if (btnLintStory) {
    btnLintStory.addEventListener('click', runStoryLinting);
  }

  // Export User Stories (DOCX / Markdown)
  if (btnExportDocx) {
    btnExportDocx.addEventListener('click', () => {
      const story = collectStoryFromForm();
      showToast('Đang đóng gói tài liệu Word OpenXML...', 'info');

      fetch('/api/ba/story/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ stories: [story], format: 'docx' }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast('Xuất tệp Word thành công!', 'success');
            if (res.download_url) {
              window.open(res.download_url, '_blank');
            }
            if (loadHistory) loadHistory();
          } else {
            showToast(`Lỗi: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi kết nối: ${err.message}`, 'error'));
    });
  }

  if (btnExportMd) {
    btnExportMd.addEventListener('click', () => {
      const story = collectStoryFromForm();
      showToast('Đang xuất Markdown Gherkin...', 'info');

      fetch('/api/ba/story/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ stories: [story], format: 'markdown' }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast('Xuất Markdown thành công!', 'success');
            if (res.download_url) {
              window.open(res.download_url, '_blank');
            }
          } else {
            showToast(`Lỗi: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi: ${err.message}`, 'error'));
    });
  }

  // 1-Click Generate Full BRD Document
  if (btnGenerateBRD) {
    btnGenerateBRD.addEventListener('click', () => {
      const story = collectStoryFromForm();
      const projName = document.getElementById('brd-proj-name')?.value || 'Hệ Thống Số Hóa Nghiệp Vụ Doanh Nghiệp';
      const author = document.getElementById('brd-author')?.value || 'Lead Business Analyst';

      showToast('Đang tổng hợp và lắp ráp tài liệu BRD/SRS hoàn chỉnh...', 'info');

      fetch('/api/ba/brd/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          metadata: {
            title: 'TÀI LIỆU YÊU CẦU NGHIỆP VỤ (BRD / SRS SPECIFICATION)',
            project_name: projName,
            version: '1.0',
            author,
            organization: 'Khối Công Nghệ & Vận Hành Doanh Nghiệp',
          },
          stories: [story],
        }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast('Đã sinh tài liệu BRD hoàn chỉnh (.docx)!', 'success');
            if (res.download_url) {
              window.open(res.download_url, '_blank');
            }
            if (loadHistory) loadHistory();
          } else {
            showToast(`Lỗi sinh BRD: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi mạng: ${err.message}`, 'error'));
    });
  }

  // Generate Requirements Traceability Matrix (RTM)
  const btnGenerateRTM = document.getElementById('btn-generate-rtm');
  if (btnGenerateRTM) {
    btnGenerateRTM.addEventListener('click', () => {
      const story = collectStoryFromForm();
      const projName = document.getElementById('brd-proj-name')?.value || 'Hệ Thống Doanh Nghiệp';

      showToast('Đang lắp ráp ma trận truy xuất nguồn gốc RTM...', 'info');

      fetch('/api/ba/rtm/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          project_name: projName,
          stories: [story],
        }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast(`Ma trận RTM hoàn tất! Độ bao phủ: ${res.coverage?.coverage_pct}%`, 'success');
            if (res.download_url) {
              window.open(res.download_url, '_blank');
            }
            if (loadHistory) loadHistory();
          } else {
            showToast(`Lỗi: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi mạng: ${err.message}`, 'error'));
    });
  }
}

