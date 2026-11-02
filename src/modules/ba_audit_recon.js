/**
 * src/modules/ba_audit_recon.js — Audit, Reconciliation & Change Control for BA Enterprise Workstation.
 *
 * Implements:
 * 1. Two-Way Automated Data Reconciliation Engine (A vs B, Composite Keys, Tolerance, 4-Sheet Excel).
 * 2. UAT Acceptance Checklist Matrix Generator (with OpenXML In-cell Dropdown validations).
 * 3. SRS / BRD Legal Redline Diff (Track Changes Word audit).
 */

export function initAuditReconStudio(ctx) {
  const { showToast, refreshLucideIcons, loadHistory } = ctx;

  const panel = document.getElementById('panel-audit-recon');
  if (!panel) return;

  // DOM Elements - Reconciliation
  const fileAInput = document.getElementById('recon-file-a');
  const fileBInput = document.getElementById('recon-file-b');
  const nameAInput = document.getElementById('recon-name-a');
  const nameBInput = document.getElementById('recon-name-b');
  const keyColsInput = document.getElementById('recon-key-cols');
  const compColsInput = document.getElementById('recon-comp-cols');
  const toleranceInput = document.getElementById('recon-tolerance');
  const btnRunRecon = document.getElementById('btn-run-recon');
  const reconResultsArea = document.getElementById('recon-results-area');

  // DOM Elements - UAT Checklist
  const uatProjName = document.getElementById('uat-proj-name');
  const uatModuleName = document.getElementById('uat-module-name');
  const btnGenerateUAT = document.getElementById('btn-generate-uat');

  // DOM Elements - Redline Diff
  const diffOldFileInput = document.getElementById('ba-diff-old-doc');
  const diffNewFileInput = document.getElementById('ba-diff-new-doc');
  const btnRunDocDiff = document.getElementById('btn-run-ba-diff');

  // ---------------------------------------------------------------------------
  // 1. RECONCILIATION ENGINE
  // ---------------------------------------------------------------------------
  if (btnRunRecon) {
    btnRunRecon.addEventListener('click', async () => {
      const fileA = fileAInput?.files[0];
      const fileB = fileBInput?.files[0];

      if (!fileA || !fileB) {
        showToast('Vui lòng chọn đủ 2 tệp dữ liệu (Nguồn A và Nguồn B)', 'warning');
        return;
      }

      showToast('Đang tải lên và phân tích dữ liệu 2 nguồn...', 'info');

      try {
        // Upload File A
        const formDataA = new FormData();
        formDataA.append('file', fileA);
        const resA = await fetch('/api/upload', { method: 'POST', body: formDataA }).then((r) => r.json());
        if (!resA.success) throw new Error(resA.error || 'Upload File A thất bại');

        // Upload File B
        const formDataB = new FormData();
        formDataB.append('file', fileB);
        const resB = await fetch('/api/upload', { method: 'POST', body: formDataB }).then((r) => r.json());
        if (!resB.success) throw new Error(resB.error || 'Upload File B thất bại');

        const keyCols = (keyColsInput?.value || 'id')
          .split(',')
          .map((c) => c.trim())
          .filter((c) => c.length > 0);

        const compColsRaw = compColsInput?.value || '';
        const compCols = compColsRaw
          ? compColsRaw
              .split(',')
              .map((c) => c.trim())
              .filter((c) => c.length > 0)
          : null;

        const tolerance = parseFloat(toleranceInput?.value || '0.0') || 0.0;
        const nameA = nameAInput?.value || 'Hệ Thống Nội Bộ';
        const nameB = nameBInput?.value || 'Cổng Thanh Toán Đối Tác';

        showToast('Đang khớp bản ghi theo khóa liên kết & tính toán chênh lệch...', 'info');

        const reconPayload = {
          file_a: resA.filepath || resA.filename,
          file_b: resB.filepath || resB.filename,
          key_columns: keyCols,
          compare_columns: compCols,
          tolerance: tolerance,
          name_a: nameA,
          name_b: nameB,
        };

        const reconRes = await fetch('/api/ba/reconciliation/run', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(reconPayload),
        }).then((r) => r.json());

        if (!reconRes.success) {
          throw new Error(reconRes.error || 'Đối soát thất bại');
        }

        renderReconSummary(reconRes);
        showToast('Hoàn tất đối soát 2 nguồn dữ liệu!', 'success');
        if (loadHistory) loadHistory();
      } catch (err) {
        showToast(`Lỗi đối soát: ${err.message}`, 'error');
      }
    });
  }

  function renderReconSummary(data) {
    if (!reconResultsArea) return;
    const matchRate = (data.match_rate_pct || 0).toFixed(1);
    const rateColor = matchRate >= 95 ? 'text-success' : matchRate >= 80 ? 'text-warning' : 'text-danger';

    reconResultsArea.innerHTML = `
      <div class="recon-dashboard-card">
        <div class="recon-header-stats">
          <div class="stat-box primary">
            <div class="stat-label">Tỷ Lệ Khớp</div>
            <div class="stat-value ${rateColor}">${matchRate}%</div>
          </div>
          <div class="stat-box success">
            <div class="stat-label">Khớp Tuyệt Đối</div>
            <div class="stat-value">${(data.matched_exact || 0).toLocaleString()}</div>
          </div>
          <div class="stat-box warning">
            <div class="stat-label">Lệch Giá Trị</div>
            <div class="stat-value">${(data.matched_discrepant || 0).toLocaleString()}</div>
          </div>
          <div class="stat-box danger">
            <div class="stat-label">Thiếu ở Nguồn B</div>
            <div class="stat-value">${(data.missing_in_b || 0).toLocaleString()}</div>
          </div>
          <div class="stat-box info">
            <div class="stat-label">Thiếu ở Nguồn A</div>
            <div class="stat-value">${(data.missing_in_a || 0).toLocaleString()}</div>
          </div>
        </div>

        <div class="recon-summary-details">
          <p><strong>Tổng bản ghi Nguồn A:</strong> ${(data.total_a || 0).toLocaleString()} | <strong>Nguồn B:</strong> ${(data.total_b || 0).toLocaleString()}</p>
          <p><strong>Tổng chênh lệch giá trị (Variance Sum):</strong> ${Math.abs(data.total_variance_sum || 0).toLocaleString()} VNĐ</p>
        </div>

        <div class="recon-actions">
          <a href="${data.download_url}" target="_blank" class="btn btn-primary btn-download-recon">
            <i data-lucide="download"></i> Tải Báo Cáo Đối Soát Đầy Đủ 4 Sheet (.xlsx)
          </a>
        </div>
      </div>
    `;
    refreshLucideIcons();
  }

  // ---------------------------------------------------------------------------
  // 2. UAT CHECKLIST MATRIX GENERATOR
  // ---------------------------------------------------------------------------
  if (btnGenerateUAT) {
    btnGenerateUAT.addEventListener('click', async () => {
      const projName = uatProjName?.value || 'Hệ Thống Nghiệp Vụ Doanh Nghiệp';
      const modName = uatModuleName?.value || 'Phân Hệ Thanh Toán & Giao Dịch';

      showToast('Đang thu thập kịch bản và sinh bảng kiểm thử nghiệm thu UAT...', 'info');

      try {
        // Fetch current templates to construct standard UAT test cases
        const tplRes = await fetch('/api/ba/templates').then((r) => r.json());
        let stories = [];
        if (tplRes.success && tplRes.stories) {
          stories = Object.values(tplRes.stories);
        }

        const uatRes = await fetch('/api/ba/uat/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            project_name: projName,
            module_name: modName,
            stories: stories,
          }),
        }).then((r) => r.json());

        if (uatRes.success) {
          showToast(`Đã sinh ${uatRes.total_testcases} ca kiểm thử UAT chuẩn OpenXML!`, 'success');
          if (uatRes.download_url) {
            window.open(uatRes.download_url, '_blank');
          }
          if (loadHistory) loadHistory();
        } else {
          showToast(`Lỗi: ${uatRes.error}`, 'error');
        }
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error');
      }
    });
  }

  // ---------------------------------------------------------------------------
  // 3. SRS / BRD REDLINE DIFF (LEGAL TRACK CHANGES)
  // ---------------------------------------------------------------------------
  if (btnRunDocDiff) {
    btnRunDocDiff.addEventListener('click', async () => {
      const oldDoc = diffOldFileInput?.files[0];
      const newDoc = diffNewFileInput?.files[0];

      if (!oldDoc || !newDoc) {
        showToast('Vui lòng chọn 2 phiên bản tài liệu Word (.docx) để so sánh', 'warning');
        return;
      }

      showToast('Đang tải tài liệu và thực hiện thẩm định Track Changes...', 'info');

      try {
        const formDataA = new FormData();
        formDataA.append('file', oldDoc);
        const resA = await fetch('/api/upload', { method: 'POST', body: formDataA }).then((r) => r.json());

        const formDataB = new FormData();
        formDataB.append('file', newDoc);
        const resB = await fetch('/api/upload', { method: 'POST', body: formDataB }).then((r) => r.json());

        const diffRes = await fetch('/api/docx-diff', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            base_file: resA.filepath || resA.filename,
            revised_file: resB.filepath || resB.filename,
          }),
        }).then((r) => r.json());

        if (diffRes.success) {
          showToast('Đã tạo tài liệu Redline Diff so sánh thay đổi thành công!', 'success');
          if (diffRes.download_url) {
            window.open(diffRes.download_url, '_blank');
          }
          if (loadHistory) loadHistory();
        } else {
          showToast(`Lỗi so sánh: ${diffRes.error}`, 'error');
        }
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error');
      }
    });
  }

  // ---------------------------------------------------------------------------
  // 4. 7 CORE SQL INTEGRITY TESTS GENERATOR (prachitparikh/ba-toolkit)
  // ---------------------------------------------------------------------------
  const btnGenerateSqlTests = document.getElementById('btn-generate-sql-tests');
  const sqlTestsResult = document.getElementById('sql-tests-result');

  if (btnGenerateSqlTests) {
    btnGenerateSqlTests.addEventListener('click', async () => {
      const targetTable = document.getElementById('sql-test-target')?.value || 'orders';
      const primaryKey = document.getElementById('sql-test-pk')?.value || 'order_id';
      const parentTable = document.getElementById('sql-test-parent')?.value || 'customers';
      const foreignKey = document.getElementById('sql-test-fk')?.value || 'customer_id';
      const childTable = document.getElementById('sql-test-child')?.value || 'order_items';
      const amountColumn = document.getElementById('sql-test-amount')?.value || 'total_amount';
      const dateColumn = document.getElementById('sql-test-date')?.value || 'order_date';
      const statusColumn = document.getElementById('sql-test-status')?.value || 'status';
      const dialect = document.getElementById('sql-test-dialect')?.value || 'postgres';

      showToast('Đang tạo bộ 7 câu truy vấn SQL kiểm toán dữ liệu...', 'info');

      try {
        const res = await fetch('/api/ba/sql-tests/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            target_table: targetTable,
            primary_key: primaryKey,
            parent_table: parentTable,
            foreign_key: foreignKey,
            child_table: childTable,
            amount_column: amountColumn,
            date_column: dateColumn,
            status_column: statusColumn,
            dialect: dialect,
          }),
        }).then((r) => r.json());

        if (res.success && res.tests) {
          renderSqlTests(res.tests);
          showToast(`Đã sinh thành công 7 bài kiểm tra SQL chuẩn ${dialect.toUpperCase()}!`, 'success');
        } else {
          showToast(`Lỗi: ${res.error || 'Không thể sinh SQL tests'}`, 'error');
        }
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error');
      }
    });
  }

  function renderSqlTests(tests) {
    if (!sqlTestsResult) return;
    sqlTestsResult.innerHTML = '';

    const listContainer = document.createElement('div');
    listContainer.className = 'sql-tests-list';

    tests.forEach((test, idx) => {
      const card = document.createElement('div');
      card.className = `sql-test-card severity-${(test.severity || 'major').toLowerCase()}`;

      const severityBadgeColor =
        test.severity === 'CRITICAL' ? 'badge-danger' : test.severity === 'MAJOR' ? 'badge-warning' : 'badge-info';

      card.innerHTML = `
        <div class="sql-test-card-header">
          <div class="sql-test-title-group">
            <span class="sql-test-num">#${idx + 1}</span>
            <span class="sql-test-name">${test.name}</span>
            <span class="ba-badge ${severityBadgeColor}">${test.severity}</span>
          </div>
          <button class="btn btn-sm btn-secondary btn-copy-sql">
            <i data-lucide="copy"></i> Sao Chép SQL
          </button>
        </div>
        <p class="sql-test-desc">${test.description}</p>
        <div class="sql-code-container">
          <pre class="sql-code-block"><code>${escapeHtml(test.sql)}</code></pre>
        </div>
      `;

      const copyBtn = card.querySelector('.btn-copy-sql');
      copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(test.sql).then(() => {
          copyBtn.innerHTML = '<i data-lucide="check"></i> Đã Sao Chép!';
          refreshLucideIcons();
          setTimeout(() => {
            copyBtn.innerHTML = '<i data-lucide="copy"></i> Sao Chép SQL';
            refreshLucideIcons();
          }, 2000);
        });
      });

      listContainer.appendChild(card);
    });

    sqlTestsResult.appendChild(listContainer);
    refreshLucideIcons();
  }

  // ---------------------------------------------------------------------------
  // 5. SCOPE CREEP & CHANGE IMPACT ANALYSIS (ScopeCreepAuditor)
  // ---------------------------------------------------------------------------
  const btnAnalyzeScopeCreep = document.getElementById('btn-analyze-scope-creep');
  const scopeOldFileInput = document.getElementById('ba-scope-old-doc');
  const scopeNewFileInput = document.getElementById('ba-scope-new-doc');
  const scopeProjNameInput = document.getElementById('ba-scope-proj-name');
  const scopeResultsArea = document.getElementById('scope-creep-results-area');

  if (btnAnalyzeScopeCreep) {
    btnAnalyzeScopeCreep.addEventListener('click', async () => {
      const oldDoc = scopeOldFileInput?.files[0];
      const newDoc = scopeNewFileInput?.files[0];
      const projName = scopeProjNameInput?.value || 'Dự Án Nghiệp Vụ';

      if (!oldDoc || !newDoc) {
        showToast('Vui lòng chọn đủ 2 phiên bản tài liệu (Baseline v1 và Revised v2)', 'warning');
        return;
      }

      showToast('Đang tải lên 2 tài liệu và bóc tách Semantic Delta...', 'info');

      try {
        const formDataA = new FormData();
        formDataA.append('file', oldDoc);
        const resA = await fetch('/api/upload', { method: 'POST', body: formDataA }).then((r) => r.json());
        if (!resA.success) throw new Error(resA.error || 'Upload tài liệu v1 thất bại');

        const formDataB = new FormData();
        formDataB.append('file', newDoc);
        const resB = await fetch('/api/upload', { method: 'POST', body: formDataB }).then((r) => r.json());
        if (!resB.success) throw new Error(resB.error || 'Upload tài liệu v2 thất bại');

        showToast('Đang phân tích mức độ phình phạm vi & ước tính Man-days...', 'info');

        const scopeRes = await fetch('/api/ba/scope-diff/analyze', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            base_file: resA.filepath || resA.filename,
            revised_file: resB.filepath || resB.filename,
            project_name: projName,
          }),
        }).then((r) => r.json());

        if (scopeRes.success && scopeRes.summary) {
          renderScopeCreepSummary(scopeRes.summary, scopeRes.download_url);
          showToast('Phân tích Scope Creep & Impact hoàn tất!', 'success');
          if (loadHistory) loadHistory();
        } else {
          showToast(`Lỗi phân tích: ${scopeRes.error}`, 'error');
        }
      } catch (err) {
        showToast(`Lỗi: ${err.message}`, 'error');
      }
    });
  }

  function renderScopeCreepSummary(summary, downloadUrl) {
    if (!scopeResultsArea) return;

    const impactBadge =
      summary.impact_level === 'HIGH'
        ? '<span class="ba-badge badge-danger">MỨC ĐỘ TÁC ĐỘNG CAO (HIGH IMPACT)</span>'
        : summary.impact_level === 'MEDIUM'
        ? '<span class="ba-badge badge-warning">MỨC ĐỘ TÁC ĐỘNG TRUNG BÌNH (MEDIUM)</span>'
        : '<span class="ba-badge badge-success">MỨC ĐỘ TÁC ĐỘNG THẤP (LOW)</span>';

    let html = `
      <div class="scope-dashboard-card">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
          <h4 style="font-size: 15px; font-weight: 700; color: #f8fafc; margin: 0;">Tổng Quan Đánh Giá Phạm Vi (Scope Creep Executive KPI)</h4>
          <div>${impactBadge}</div>
        </div>

        <div class="scope-kpi-grid">
          <div class="stat-box primary">
            <div class="stat-label">Tổng Số Thay Đổi</div>
            <div class="stat-value">${summary.total_changes || 0}</div>
          </div>
          <div class="stat-box success">
            <div class="stat-label">Yêu Cầu Mới</div>
            <div class="stat-value text-success">${summary.new_requirements_count || 0}</div>
          </div>
          <div class="stat-box warning">
            <div class="stat-label">Thay Đổi Logic</div>
            <div class="stat-value text-warning">${summary.modified_logic_count || 0}</div>
          </div>
          <div class="stat-box danger">
            <div class="stat-label">Cắt Giảm Phạm Vi</div>
            <div class="stat-value text-danger">${summary.removed_scope_count || 0}</div>
          </div>
          <div class="stat-box info">
            <div class="stat-label">Ước Tính Effort</div>
            <div class="stat-value text-info">${summary.total_estimated_mandays || 0} <span style="font-size: 14px; font-weight: 500;">Man-days</span></div>
          </div>
        </div>

        <div style="margin-top: 16px;">
          <h5 style="font-size: 13px; font-weight: 700; color: #38bdf8; margin: 0 0 10px;">Chi Tiết Các Thay Đổi Ngữ Nghĩa (Semantic Deltas):</h5>
          <div class="ba-table-container">
            <table class="ba-editable-table">
              <thead>
                <tr>
                  <th style="width: 130px;">Phân Loại</th>
                  <th style="width: 90px; text-align: center;">Mức Độ</th>
                  <th style="width: 140px;">Mục / Điều Khoản</th>
                  <th>Tác Động Nghiệp Vụ</th>
                  <th style="width: 100px; text-align: right;">Man-days</th>
                </tr>
              </thead>
              <tbody>
    `;

    const deltas = summary.deltas || [];
    if (deltas.length === 0) {
      html += `<tr><td colspan="5" style="text-align: center; color: #94a3b8; padding: 16px;">Không phát hiện thay đổi đáng kể giữa 2 tài liệu.</td></tr>`;
    } else {
      deltas.forEach((d) => {
        const sevClass = d.severity === 'CRITICAL' ? 'badge-danger' : d.severity === 'MAJOR' ? 'badge-warning' : 'badge-info';
        html += `
          <tr>
            <td><strong style="color: #cbd5e1;">${d.category}</strong></td>
            <td style="text-align: center;"><span class="ba-badge ${sevClass}">${d.severity}</span></td>
            <td><code>${escapeHtml(d.section || '-')}</code></td>
            <td style="font-size: 12px; color: #cbd5e1;">${escapeHtml(d.impact || '-')}</td>
            <td style="text-align: right; font-weight: 700; color: #38bdf8;">+${d.estimated_mandays || 0} d</td>
          </tr>
        `;
      });
    }

    html += `
              </tbody>
            </table>
          </div>
        </div>

        <div style="display: flex; justify-content: flex-end; margin-top: 16px;">
          <a href="${downloadUrl}" target="_blank" class="btn btn-primary">
            <i data-lucide="download"></i> Tải Báo Cáo Đánh Giá Scope Creep (.docx)
          </a>
        </div>
      </div>
    `;

    scopeResultsArea.innerHTML = html;
    refreshLucideIcons();
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
}

