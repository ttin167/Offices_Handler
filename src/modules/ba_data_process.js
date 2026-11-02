/**
 * src/modules/ba_data_process.js — Data & Process Studio for BA Enterprise Workstation.
 *
 * Implements:
 * 1. Data Dictionary Scanner (from Excel workbook or SQL DDL statements).
 * 2. Industry Standard Blueprints (Banking Customer Account, Core Ledger).
 * 3. Interactive Schema Editor (Add/Edit/Remove fields, PK/FK/Type/Rules).
 * 4. Enterprise Styled Data Dictionary Excel Exporter (.xlsx).
 * 5. Source-to-Target Data Mapping Matrix Builder.
 */

export function initDataProcessStudio(ctx) {
  const { showToast, refreshLucideIcons, loadHistory } = ctx;

  const panel = document.getElementById('panel-data-process');
  if (!panel) return;

  // State
  let currentFields = [];
  let currentMappings = [];
  let defaultTemplates = null;

  // DOM Elements - Data Dictionary
  const fldTableName = document.getElementById('dict-table-name');
  const btnLoadSampleDict = document.getElementById('btn-load-sample-dict');
  const ddlInput = document.getElementById('ddl-input-text');
  const btnParseDDL = document.getElementById('btn-parse-ddl');
  const excelFileInput = document.getElementById('dict-excel-file');
  const btnScanExcel = document.getElementById('btn-scan-excel');
  const fieldsTbody = document.getElementById('dict-fields-tbody');
  const btnAddField = document.getElementById('btn-add-dict-field');
  const btnExportDictExcel = document.getElementById('btn-export-dict-excel');

  // DOM Elements - Data Mapping
  const mappingTbody = document.getElementById('mapping-tbody');
  const btnAddMapping = document.getElementById('btn-add-mapping');
  const btnExportMappingExcel = document.getElementById('btn-export-mapping-excel');

  // Fetch Templates on Init
  fetch('/api/ba/templates')
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        defaultTemplates = data;
        if (data.data_dictionaries && data.data_dictionaries.banking_customer_account) {
          renderDataDictionaryTable('CUSTOMER_ACCOUNTS', data.data_dictionaries.banking_customer_account);
        }
      }
    })
    .catch((err) => console.warn('Cannot fetch BA Data Dict templates:', err));

  // Load Sample Template Button
  if (btnLoadSampleDict) {
    btnLoadSampleDict.addEventListener('click', () => {
      if (defaultTemplates && defaultTemplates.data_dictionaries?.banking_customer_account) {
        renderDataDictionaryTable('CUSTOMER_ACCOUNTS', defaultTemplates.data_dictionaries.banking_customer_account);
        showToast('Đã nạp Từ điển dữ liệu chuẩn: Tài khoản Ngân hàng (CIF)', 'success');
      }
    });
  }

  // Parse SQL DDL
  if (btnParseDDL) {
    btnParseDDL.addEventListener('click', () => {
      const sql = ddlInput?.value || '';
      if (!sql.trim()) {
        showToast('Vui lòng dán câu lệnh SQL CREATE TABLE để quét cấu trúc', 'warning');
        return;
      }
      showToast('Đang phân tích cú pháp SQL DDL...', 'info');

      fetch('/api/ba/data-dict/scan-ddl', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sql }),
      })
        .then((r) => r.json())
        .then((data) => {
          if (data.success && data.tables) {
            const tableNames = Object.keys(data.tables);
            if (tableNames.length === 0) {
              showToast('Không tìm thấy bảng hợp lệ trong câu lệnh SQL', 'warning');
              return;
            }
            const firstTable = tableNames[0];
            const rawFields = data.tables[firstTable];
            const mapped = rawFields.map((f) => ({
              field: f.field_name,
              name: f.display_name || f.field_name,
              type: f.data_type,
              nullable: f.nullable,
              pk: f.is_primary_key ? 'YES' : 'NO',
              fk: f.is_foreign_key ? f.references : '-',
              description: f.description || '',
              rules: f.business_rules || '',
            }));
            renderDataDictionaryTable(firstTable, mapped);
            showToast(`Đã trích xuất thành công ${mapped.length} trường từ bảng ${firstTable}!`, 'success');
          } else {
            showToast(`Lỗi phân tích: ${data.error || 'Cú pháp SQL chưa chuẩn'}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi kết nối: ${err.message}`, 'error'));
    });
  }

  // Scan Excel File
  if (btnScanExcel && excelFileInput) {
    btnScanExcel.addEventListener('click', () => {
      const file = excelFileInput.files[0];
      if (!file) {
        showToast('Vui lòng chọn một tệp Excel (.xlsx) để quét', 'warning');
        return;
      }

      showToast('Đang tải tệp và suy luận kiểu dữ liệu...', 'info');
      const formData = new FormData();
      formData.append('file', file);

      fetch('/api/upload', {
        method: 'POST',
        body: formData,
      })
        .then((r) => r.json())
        .then((uploadRes) => {
          if (!uploadRes.success || !uploadRes.filename) {
            throw new Error(uploadRes.error || 'Upload thất bại');
          }
          return fetch('/api/ba/data-dict/scan-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filepath: uploadRes.filepath || uploadRes.filename }),
          });
        })
        .then((r) => r.json())
        .then((scanRes) => {
          if (scanRes.success && scanRes.fields) {
            const mapped = scanRes.fields.map((f) => ({
              field: f.field_name,
              name: f.display_name || f.field_name,
              type: f.data_type,
              nullable: f.nullable,
              pk: f.is_primary_key ? 'YES' : 'NO',
              fk: f.is_foreign_key ? f.references : '-',
              description: f.description || '',
              rules: f.business_rules || '',
            }));
            renderDataDictionaryTable(scanRes.table_name || 'SCANNED_TABLE', mapped);
            showToast(`Đã quét thành công ${mapped.length} cột dữ liệu từ Excel!`, 'success');
          } else {
            showToast(`Lỗi quét: ${scanRes.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi: ${err.message}`, 'error'));
    });
  }

  function renderDataDictionaryTable(tableName, fields) {
    if (fldTableName) fldTableName.value = tableName;
    currentFields = fields || [];
    if (!fieldsTbody) return;
    fieldsTbody.innerHTML = '';

    currentFields.forEach((f, idx) => {
      addFieldRow(f, idx);
    });
    refreshLucideIcons();
  }

  function addFieldRow(f = {}, idx = 0) {
    if (!fieldsTbody) return;
    const tr = document.createElement('tr');
    tr.className = 'dict-field-row';
    tr.dataset.index = idx;

    tr.innerHTML = `
      <td><input type="text" class="field-code" value="${f.field || ''}" placeholder="column_name" /></td>
      <td><input type="text" class="field-name" value="${f.name || ''}" placeholder="Tên hiển thị" /></td>
      <td>
        <select class="field-type">
          <option value="VARCHAR(50)" ${f.type && f.type.startsWith('VARCHAR') ? 'selected' : ''}>VARCHAR</option>
          <option value="INT" ${f.type === 'INT' || f.type === 'INTEGER' ? 'selected' : ''}>INT</option>
          <option value="BIGINT" ${f.type === 'BIGINT' ? 'selected' : ''}>BIGINT</option>
          <option value="DECIMAL(18,2)" ${f.type && f.type.startsWith('DECIMAL') ? 'selected' : ''}>DECIMAL</option>
          <option value="DATE" ${f.type === 'DATE' ? 'selected' : ''}>DATE</option>
          <option value="TIMESTAMP" ${f.type === 'TIMESTAMP' ? 'selected' : ''}>TIMESTAMP</option>
          <option value="BOOLEAN" ${f.type === 'BOOLEAN' ? 'selected' : ''}>BOOLEAN</option>
          <option value="TEXT" ${f.type === 'TEXT' ? 'selected' : ''}>TEXT</option>
          <option value="JSON" ${f.type === 'JSON' ? 'selected' : ''}>JSON</option>
        </select>
      </td>
      <td>
        <select class="field-nullable">
          <option value="NO" ${f.nullable === 'NO' ? 'selected' : ''}>NO (Bắt buộc)</option>
          <option value="YES" ${f.nullable === 'YES' ? 'selected' : ''}>YES (Tùy chọn)</option>
        </select>
      </td>
      <td class="text-center">
        <input type="checkbox" class="field-pk" ${f.pk === 'YES' ? 'checked' : ''} />
      </td>
      <td><input type="text" class="field-fk" value="${f.fk && f.fk !== '-' ? f.fk : ''}" placeholder="Bảng(khoá)" /></td>
      <td><input type="text" class="field-desc" value="${f.description || ''}" placeholder="Mô tả ý nghĩa nghiệp vụ" /></td>
      <td><input type="text" class="field-rules" value="${f.rules || ''}" placeholder="Quy tắc ràng buộc validation" /></td>
      <td class="text-center">
        <button class="btn-icon-danger btn-del-field" title="Xóa trường"><i data-lucide="trash-2"></i></button>
      </td>
    `;

    tr.querySelector('.btn-del-field').addEventListener('click', () => {
      tr.remove();
      refreshLucideIcons();
    });

    fieldsTbody.appendChild(tr);
    refreshLucideIcons();
  }

  if (btnAddField) {
    btnAddField.addEventListener('click', () => {
      addFieldRow({ field: '', name: '', type: 'VARCHAR(100)', nullable: 'YES', pk: 'NO', fk: '-', description: '', rules: '' });
    });
  }

  // Export Data Dictionary to Excel
  if (btnExportDictExcel) {
    btnExportDictExcel.addEventListener('click', () => {
      const rows = fieldsTbody ? fieldsTbody.querySelectorAll('.dict-field-row') : [];
      if (rows.length === 0) {
        showToast('Danh sách trường dữ liệu đang trống', 'warning');
        return;
      }

      const fields = [];
      rows.forEach((r) => {
        fields.push({
          field: r.querySelector('.field-code')?.value || '',
          name: r.querySelector('.field-name')?.value || '',
          type: r.querySelector('.field-type')?.value || 'VARCHAR',
          nullable: r.querySelector('.field-nullable')?.value || 'YES',
          pk: r.querySelector('.field-pk')?.checked ? 'YES' : 'NO',
          fk: r.querySelector('.field-fk')?.value || '-',
          description: r.querySelector('.field-desc')?.value || '',
          rules: r.querySelector('.field-rules')?.value || '',
        });
      });

      const tblName = fldTableName?.value || 'DATA_DICTIONARY';
      showToast('Đang đóng gói Từ điển dữ liệu sang Excel OpenXML...', 'info');

      fetch('/api/ba/data-dict/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ table_name: tblName, fields }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast('Xuất Excel Từ Điển Dữ Liệu thành công!', 'success');
            if (res.download_url) {
              window.open(res.download_url, '_blank');
            }
            if (loadHistory) loadHistory();
          } else {
            showToast(`Lỗi: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi: ${err.message}`, 'error'));
    });
  }

  // ---------------------------------------------------------------------------
  // DATA MAPPING SUB-MODULE
  // ---------------------------------------------------------------------------
  function addMappingRow(m = {}) {
    if (!mappingTbody) return;
    const tr = document.createElement('tr');
    tr.className = 'mapping-row';

    tr.innerHTML = `
      <td><input type="text" class="map-src-field" value="${m.src || ''}" placeholder="source_col" /></td>
      <td><input type="text" class="map-src-type" value="${m.src_type || 'VARCHAR'}" placeholder="VARCHAR(50)" /></td>
      <td>
        <select class="map-rule-type">
          <option value="Direct Copy" ${m.rule === 'Direct Copy' ? 'selected' : ''}>Direct 1:1 Copy</option>
          <option value="Lookup / Mapping" ${m.rule === 'Lookup / Mapping' ? 'selected' : ''}>Bảng Tra Cứu (Lookup)</option>
          <option value="Concatenate" ${m.rule === 'Concatenate' ? 'selected' : ''}>Ghép Chuỗi (CONCAT)</option>
          <option value="Format / Cast" ${m.rule === 'Format / Cast' ? 'selected' : ''}>Chuyển Kiểu (CAST/TO_CHAR)</option>
          <option value="Mask PII" ${m.rule === 'Mask PII' ? 'selected' : ''}>Che Giấu Dữ Liệu (PII Masking)</option>
          <option value="Business Logic" ${m.rule === 'Business Logic' ? 'selected' : ''}>Công Thức Nghiệp Vụ Riêng</option>
        </select>
      </td>
      <td><input type="text" class="map-target-field" value="${m.target || ''}" placeholder="target_col" /></td>
      <td><input type="text" class="map-target-type" value="${m.target_type || 'VARCHAR'}" placeholder="VARCHAR(50)" /></td>
      <td><input type="text" class="map-notes" value="${m.notes || ''}" placeholder="Giải thích logic chuyển đổi" /></td>
      <td class="text-center">
        <button class="btn-icon-danger btn-del-mapping" title="Xóa dòng ánh xạ"><i data-lucide="trash-2"></i></button>
      </td>
    `;

    tr.querySelector('.btn-del-mapping').addEventListener('click', () => {
      tr.remove();
      refreshLucideIcons();
    });

    mappingTbody.appendChild(tr);
    refreshLucideIcons();
  }

  if (btnAddMapping) {
    btnAddMapping.addEventListener('click', () => {
      addMappingRow({ src: '', src_type: 'VARCHAR(50)', rule: 'Direct Copy', target: '', target_type: 'VARCHAR(50)', notes: '' });
    });
  }

  // Default sample mapping rows
  if (mappingTbody && mappingTbody.children.length === 0) {
    addMappingRow({ src: 'cust_id', src_type: 'VARCHAR(20)', rule: 'Direct Copy', target: 'cif_number', target_type: 'VARCHAR(20)', notes: 'Khóa chính khách hàng' });
    addMappingRow({ src: 'id_number', src_type: 'VARCHAR(12)', rule: 'Mask PII', target: 'national_id_masked', target_type: 'VARCHAR(12)', notes: 'Che 6 số giữa: 0010******89' });
    addMappingRow({ src: 'first_name, last_name', src_type: 'VARCHAR(50)', rule: 'Concatenate', target: 'full_name', target_type: 'VARCHAR(100)', notes: 'Ghép Họ + Tên in hoa: UPPER(CONCAT(last_name, " ", first_name))' });
  }

  if (btnExportMappingExcel) {
    btnExportMappingExcel.addEventListener('click', () => {
      showToast('Đang phát hành ma trận Source-to-Target Data Mapping...', 'info');
      // Export logic using data-dict endpoint with custom formatting
      const rows = mappingTbody ? mappingTbody.querySelectorAll('.mapping-row') : [];
      const fields = [];
      rows.forEach((r) => {
        const src = r.querySelector('.map-src-field')?.value || '';
        const target = r.querySelector('.map-target-field')?.value || '';
        const rule = r.querySelector('.map-rule-type')?.value || '';
        const notes = r.querySelector('.map-notes')?.value || '';
        fields.push({
          field: `${src} -> ${target}`,
          name: target,
          type: r.querySelector('.map-target-type')?.value || 'VARCHAR',
          nullable: 'YES',
          pk: 'NO',
          fk: '-',
          description: `Chuyển đổi từ [${src}] sang [${target}] theo quy tắc [${rule}]`,
          rules: notes,
        });
      });

      fetch('/api/ba/data-dict/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ table_name: 'SOURCE_TO_TARGET_MAPPING', fields }),
      })
        .then((r) => r.json())
        .then((res) => {
          if (res.success) {
            showToast('Xuất Ma Trận Ánh Xạ Dữ Liệu thành công!', 'success');
            if (res.download_url) window.open(res.download_url, '_blank');
          } else {
            showToast(`Lỗi: ${res.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi: ${err.message}`, 'error'));
    });
  }

  // ---------------------------------------------------------------------------
  // DBML BIDIRECTIONAL STUDIO (From @software-mansion/dbml)
  // ---------------------------------------------------------------------------
  const dbmlInput = document.getElementById('dbml-input-code');
  const dbmlOutput = document.getElementById('dbml-output-code');
  const btnDbmlToSql = document.getElementById('btn-dbml-to-sql');
  const btnSqlToDbml = document.getElementById('btn-sql-to-dbml');
  const btnDbmlToErd = document.getElementById('btn-dbml-to-erd');

  if (btnDbmlToSql) {
    btnDbmlToSql.addEventListener('click', () => {
      const code = dbmlInput?.value || '';
      if (!code.trim()) {
        showToast('Vui lòng nhập mã DBML', 'warning');
        return;
      }
      showToast('Đang biên dịch DBML sang SQL DDL...', 'info');
      fetch('/api/ba/dbml/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode: 'dbml_to_sql', code }),
      })
        .then((r) => r.json())
        .then((data) => {
          if (data.success && dbmlOutput) {
            dbmlOutput.value = data.sql || '';
            showToast('Biên dịch DBML sang SQL DDL thành công!', 'success');
          } else {
            showToast(`Lỗi: ${data.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi kết nối: ${err.message}`, 'error'));
    });
  }

  if (btnSqlToDbml) {
    btnSqlToDbml.addEventListener('click', () => {
      const code = dbmlInput?.value || '';
      if (!code.trim()) {
        showToast('Vui lòng nhập câu lệnh SQL DDL', 'warning');
        return;
      }
      showToast('Đang dịch ngược SQL sang DBML...', 'info');
      fetch('/api/ba/dbml/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode: 'sql_to_dbml', code }),
      })
        .then((r) => r.json())
        .then((data) => {
          if (data.success && dbmlOutput) {
            dbmlOutput.value = data.dbml || '';
            showToast('Chuyển đổi SQL sang DBML thành công!', 'success');
          } else {
            showToast(`Lỗi: ${data.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi kết nối: ${err.message}`, 'error'));
    });
  }

  if (btnDbmlToErd) {
    btnDbmlToErd.addEventListener('click', () => {
      const code = dbmlInput?.value || '';
      if (!code.trim()) {
        showToast('Vui lòng nhập mã DBML', 'warning');
        return;
      }
      showToast('Đang sinh sơ đồ ERD...', 'info');
      fetch('/api/ba/dbml/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode: 'dbml_to_mermaid', code }),
      })
        .then((r) => r.json())
        .then((data) => {
          if (data.success && dbmlOutput) {
            dbmlOutput.value = data.mermaid_erd || '';
            showToast('Đã sinh mã Mermaid ERD từ DBML!', 'success');
          } else {
            showToast(`Lỗi: ${data.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi kết nối: ${err.message}`, 'error'));
    });
  }

  // ---------------------------------------------------------------------------
  // SATSUMA DATA MAPPING DSL (From EqualExperts/satsuma-lang)
  // ---------------------------------------------------------------------------
  const satsumaInput = document.getElementById('satsuma-input-dsl');
  const btnSatsumaVerify = document.getElementById('btn-satsuma-verify');
  const satsumaResultBox = document.getElementById('satsuma-result-box');

  if (btnSatsumaVerify) {
    btnSatsumaVerify.addEventListener('click', () => {
      const dsl = satsumaInput?.value || '';
      if (!dsl.trim()) {
        showToast('Vui lòng nhập cú pháp Satsuma DSL (Source.Col -> Target.Col [Rule])', 'warning');
        return;
      }
      showToast('Đang phân tích cú pháp Satsuma & kiểm tra lỗi kiểu dữ liệu...', 'info');

      fetch('/api/ba/mapping/satsuma-verify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ dsl_text: dsl }),
      })
        .then((r) => r.json())
        .then((data) => {
          if (data.success && satsumaResultBox) {
            const rules = data.rules || [];
            let html = `<div style="font-weight: 700; color: #38bdf8; margin-bottom: 8px;">Đã phân tích ${rules.length} quy tắc ánh xạ:</div>`;
            html += '<ul style="list-style: none; padding: 0; margin: 0; display: flex; flex-direction: column; gap: 6px; font-size: 12px;">';
            rules.forEach((r) => {
              const statusColor = r.is_valid ? '#4ade80' : '#f87171';
              const statusIcon = r.is_valid ? 'check' : 'alert-triangle';
              html += `<li style="padding: 6px 10px; background: #090d16; border-radius: 6px; border-left: 3px solid ${statusColor};">
                <strong>${r.source}</strong> &rarr; <strong>${r.target}</strong> <span style="color: #fbbf24;">[${r.rule}]</span>
                ${r.warnings && r.warnings.length ? `<div style="color: #f87171; margin-top: 4px;">⚠️ ${r.warnings.join('<br/>')}</div>` : ''}
              </li>`;
            });
            html += '</ul>';

            if (data.lineage_mermaid) {
              html += `<div style="margin-top: 12px; font-weight: 700; color: #4ade80;">Sơ đồ Luồng Dữ Liệu (Data Lineage):</div>
              <pre style="background: #090d16; padding: 10px; border-radius: 6px; font-family: monospace; font-size: 11px; color: #cbd5e1; overflow-x: auto; margin-top: 6px;">${data.lineage_mermaid}</pre>`;
            }

            satsumaResultBox.innerHTML = html;
            showToast('Kiểm tra cú pháp Satsuma hoàn tất!', 'success');
          } else {
            showToast(`Lỗi: ${data.error}`, 'error');
          }
        })
        .catch((err) => showToast(`Lỗi: ${err.message}`, 'error'));
    });
  }
}

