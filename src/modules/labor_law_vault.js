/**
 * src/modules/labor_law_vault.js — Vietnam Labor Law & HR Knowledge Vault Controller
 *
 * Provides:
 *  - 3-Column Knowledge Workstation (Obsidian / Notion style)
 *  - Category groups navigation (5 statutory groups + All)
 *  - Sub-10ms unaccented SQLite FTS5 search
 *  - Markdown reader view with legal grounds banner
 *  - Quick Parameter Inspector (copy YAML rates / ceilings / brackets)
 *  - 1-Click OpenXML Word (.docx) export & Obsidian Vault (.zip) export
 */

function sanitizeLatexMath(text) {
  if (!text) return '';
  let s = text;
  // Replace \text{...}, \mathbf{...}, \mathit{...}
  s = s.replace(/\\(?:text|mathbf|mathit)\{([^}]+)\}/g, '$1');
  // Replace math operators
  s = s.replace(/\\le\b/g, '≤');
  s = s.replace(/\\ge\b/g, '≥');
  s = s.replace(/\\times\b/g, '×');
  s = s.replace(/\\rightarrow\b/g, '→');
  s = s.replace(/\\leftarrow\b/g, '←');
  s = s.replace(/\\sum\b/g, 'Tổng');
  s = s.replace(/\\%/g, '%');
  // Remove enclosing double and single dollar signs
  s = s.replace(/\$\$([^\$]+)\$\$/g, '$1');
  s = s.replace(/\$([^\$\n]+)\$/g, '$1');
  return s;
}

/**
 * Lightweight, safe Zero-Dependency Markdown to HTML Converter
 * Handles headers, blockquotes, code blocks, tables, lists, bold/italic, and inline code.
 */
function renderMarkdownToHtml(md) {
  if (!md) return '';

  const cleanMd = sanitizeLatexMath(md);
  const lines = cleanMd.split('\n');
  const output = [];
  let inCodeBlock = false;
  let codeLang = '';
  let codeContent = [];
  let inTable = false;
  let tableHeaderDone = false;
  let inList = false;
  let listType = 'ul';

  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function inlineFormat(text) {
    let s = escapeHtml(text);
    // Bold
    s = s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Italic
    s = s.replace(/\*([^*]+?)\*/g, '<em>$1</em>');
    // Inline code
    s = s.replace(/`([^`]+?)`/g, '<code class="law-inline-code">$1</code>');
    // Links
    s = s.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>');
    return s;
  }

  function closeList() {
    if (inList) {
      output.push(`</${listType}>`);
      inList = false;
    }
  }

  function closeTable() {
    if (inTable) {
      output.push('</tbody></table></div>');
      inTable = false;
      tableHeaderDone = false;
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const trimmed = rawLine.trim();

    // Code blocks
    if (trimmed.startsWith('```')) {
      if (inCodeBlock) {
        output.push(`<pre class="law-code-block"><code class="language-${codeLang}">${escapeHtml(codeContent.join('\n'))}</code></pre>`);
        inCodeBlock = false;
        codeContent = [];
        codeLang = '';
      } else {
        closeList();
        closeTable();
        inCodeBlock = true;
        codeLang = trimmed.substring(3).trim();
      }
      continue;
    }

    if (inCodeBlock) {
      codeContent.push(rawLine);
      continue;
    }

    // Horizontal Rule
    if (/^---$|^___$|^\*\*\*$/.test(trimmed)) {
      closeList();
      closeTable();
      output.push('<hr class="law-divider" />');
      continue;
    }

    // Table rows
    if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
      closeList();
      // Check if it's separator row |---|---|
      if (/^\|[\s\-:|]+\|$/.test(trimmed)) {
        tableHeaderDone = true;
        continue;
      }

      const cells = trimmed
        .slice(1, -1)
        .split('|')
        .map((c) => c.trim());

      if (!inTable) {
        inTable = true;
        tableHeaderDone = false;
        output.push('<div class="law-table-wrapper"><table class="law-table"><thead><tr>');
        cells.forEach((c) => output.push(`<th>${inlineFormat(c)}</th>`));
        output.push('</tr></thead><tbody>');
      } else {
        output.push('<tr>');
        cells.forEach((c) => output.push(`<td>${inlineFormat(c)}</td>`));
        output.push('</tr>');
      }
      continue;
    } else {
      closeTable();
    }

    // Headers
    if (trimmed.startsWith('#')) {
      closeList();
      if (trimmed.startsWith('#### ')) {
        output.push(`<h4 class="law-h4">${inlineFormat(trimmed.substring(5))}</h4>`);
      } else if (trimmed.startsWith('### ')) {
        output.push(`<h3 class="law-h3">${inlineFormat(trimmed.substring(4))}</h3>`);
      } else if (trimmed.startsWith('## ')) {
        output.push(`<h2 class="law-h2">${inlineFormat(trimmed.substring(3))}</h2>`);
      } else if (trimmed.startsWith('# ')) {
        output.push(`<h1 class="law-h1">${inlineFormat(trimmed.substring(2))}</h1>`);
      }
      continue;
    }

    // Blockquotes
    if (trimmed.startsWith('>')) {
      closeList();
      const quoteText = trimmed.replace(/^>\s*/, '');
      output.push(`<blockquote class="law-blockquote">${inlineFormat(quoteText)}</blockquote>`);
      continue;
    }

    // Bullet Lists
    const bulletMatch = trimmed.match(/^[-*]\s+(.+)$/);
    if (bulletMatch) {
      if (!inList || listType !== 'ul') {
        closeList();
        inList = true;
        listType = 'ul';
        output.push('<ul class="law-list">');
      }
      output.push(`<li>${inlineFormat(bulletMatch[1])}</li>`);
      continue;
    }

    // Numbered Lists
    const numMatch = trimmed.match(/^(\d+)\.\s+(.+)$/);
    if (numMatch) {
      if (!inList || listType !== 'ol') {
        closeList();
        inList = true;
        listType = 'ol';
        output.push('<ol class="law-list-num">');
      }
      output.push(`<li>${inlineFormat(numMatch[2])}</li>`);
      continue;
    }

    // Empty line
    if (!trimmed) {
      closeList();
      continue;
    }

    // Standard Paragraph
    closeList();
    output.push(`<p class="law-p">${inlineFormat(trimmed)}</p>`);
  }

  closeList();
  closeTable();
  if (inCodeBlock) {
    output.push(`<pre class="law-code-block"><code>${escapeHtml(codeContent.join('\n'))}</code></pre>`);
  }

  return output.join('\n');
}

export function initLaborLawVault({ showToast, refreshLucideIcons, loadHistory }) {
  const groupsContainer = document.getElementById('law-groups-container');
  const searchInput = document.getElementById('law-search-input');
  const btnClearSearch = document.getElementById('btn-law-clear-search');
  const articlesCountLabel = document.getElementById('law-articles-count-label');
  const searchPerfBadge = document.getElementById('law-search-perf-badge');
  const articlesContainer = document.getElementById('law-articles-container');

  const readerEmpty = document.getElementById('law-reader-empty');
  const readerContent = document.getElementById('law-reader-content');
  const readerTags = document.getElementById('law-reader-tags');
  const readerTitle = document.getElementById('law-reader-title');
  const readerSource = document.getElementById('law-reader-source');
  const readerMarkdown = document.getElementById('law-reader-markdown');

  const paramsContainer = document.getElementById('law-params-container');
  const paramsCode = document.getElementById('law-params-code');
  const btnCopyParams = document.getElementById('btn-law-copy-params');

  const btnExportDocx = document.getElementById('btn-law-export-docx');
  const btnExportVault = document.getElementById('btn-law-export-vault');

  if (!groupsContainer || !articlesContainer) return;

  let activeCategory = 'all';
  let activeArticleId = null;
  let currentArticleDetail = null;
  let searchTimer = null;
  let cachedCategories = [];
  let cachedArticles = [];

  // Group Lucide icons mapping
  const GROUP_ICONS = {
    all: 'book-open',
    '01_Luat_Lao_Dong_2019': 'file-text',
    '02_Bao_Hiem_BHXH_BHYT_BHTN': 'shield-check',
    '03_Thue_TNCN_Va_Tien_Luong_2026': 'calculator',
    '04_Quy_Chuan_Nhan_Su_Doanh_Nghiep': 'briefcase',
    '05_Bieu_Mau_Doanh_Nghiep': 'table',
  };

  /**
   * 1. Load Categories
   */
  async function loadCategories() {
    try {
      const res = await fetch('/api/law/categories');
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Lỗi tải danh mục');

      cachedCategories = data.categories || [];
      renderCategories();
    } catch (err) {
      console.error('[LaborLawVault] Failed to load categories:', err);
    }
  }

  function renderCategories() {
    groupsContainer.innerHTML = '';
    cachedCategories.forEach((group) => {
      const item = document.createElement('div');
      item.className = `law-group-item ${group.id === activeCategory ? 'active' : ''}`;
      item.setAttribute('data-group-id', group.id);

      const iconName = GROUP_ICONS[group.id] || group.icon || 'folder';
      item.innerHTML = `
        <div class="law-group-icon"><i data-lucide="${iconName}"></i></div>
        <div class="law-group-info">
          <div class="law-group-title">${group.title}</div>
          <div class="law-group-desc">${group.description || ''}</div>
        </div>
        <div class="law-group-count">${group.count ?? 0}</div>
      `;

      item.addEventListener('click', () => {
        if (activeCategory === group.id) return;
        activeCategory = group.id;
        renderCategories();
        loadArticles();
      });

      groupsContainer.appendChild(item);
    });

    refreshLucideIcons();
  }

  /**
   * 2. Load Articles
   */
  async function loadArticles() {
    const query = searchInput ? searchInput.value.trim() : '';

    if (query) {
      btnClearSearch && (btnClearSearch.style.display = 'flex');
    } else {
      btnClearSearch && (btnClearSearch.style.display = 'none');
    }

    try {
      articlesContainer.innerHTML = '<div class="law-loading"><div class="spinner"></div> Đang tra cứu kho tri thức...</div>';

      let url = '/api/law/articles?';
      const params = new URLSearchParams();
      if (activeCategory && activeCategory !== 'all') {
        params.set('category', activeCategory);
      }
      if (query) {
        params.set('query', query);
      }
      params.set('limit', '50');
      url += params.toString();

      const startTime = performance.now();
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      const clientLatency = (performance.now() - startTime).toFixed(1);

      if (!data.success) throw new Error(data.error || 'Lỗi tra cứu bài viết');

      let docs = [];
      if (query && data.results) {
        // Search results
        docs = data.results.map((r) => ({
          doc_id: r.doc_id,
          title: r.title,
          category: r.category,
          source: r.source,
          tags: r.tags,
          summary: sanitizeLatexMath(r.snippet ? r.snippet.replace(/<\/?mark>/g, '') : ''),
          raw_snippet: sanitizeLatexMath(r.snippet || ''),
        }));

        const perfMs = data.elapsed_ms !== undefined ? data.elapsed_ms.toFixed(1) : clientLatency;
        if (searchPerfBadge) {
          searchPerfBadge.style.display = 'inline-flex';
          searchPerfBadge.textContent = `⚡ ${perfMs} ms (FTS5)`;
        }
        if (articlesCountLabel) {
          articlesCountLabel.textContent = `Tìm thấy ${docs.length} kết quả`;
        }
      } else {
        // Direct category listing
        docs = (data.documents || []).map((d) => ({
          ...d,
          summary: sanitizeLatexMath(d.summary || ''),
        }));
        if (searchPerfBadge) searchPerfBadge.style.display = 'none';
        if (articlesCountLabel) {
          articlesCountLabel.textContent = `${docs.length} tài liệu quy chuẩn`;
        }
      }

      cachedArticles = docs;
      renderArticles(docs);

      // Auto-select first article if none or current not in list
      if (docs.length > 0) {
        const hasCurrent = docs.some((d) => d.doc_id === activeArticleId);
        if (!activeArticleId || !hasCurrent) {
          loadArticleDetail(docs[0].doc_id);
        }
      } else {
        showEmptyReader('Không tìm thấy tài liệu nào phù hợp với bộ lọc.');
      }
    } catch (err) {
      console.error('[LaborLawVault] Failed to load articles:', err);
      articlesContainer.innerHTML = `<div class="law-error"><i data-lucide="alert-circle"></i> ${err.message}</div>`;
      refreshLucideIcons();
    }
  }

  function renderArticles(docs) {
    articlesContainer.innerHTML = '';
    if (!docs || docs.length === 0) {
      articlesContainer.innerHTML = `
        <div class="law-articles-empty">
          <i data-lucide="search-x"></i>
          <div>Không tìm thấy văn bản nào</div>
          <small>Thử đổi từ khóa hoặc chọn "Tất Cả Văn Bản"</small>
        </div>
      `;
      refreshLucideIcons();
      return;
    }

    docs.forEach((doc) => {
      const card = document.createElement('div');
      card.className = `law-article-card ${doc.doc_id === activeArticleId ? 'active' : ''}`;
      card.setAttribute('data-doc-id', doc.doc_id);

      const tagList = (doc.tags || '')
        .split(' ')
        .filter((t) => t.trim().length > 0)
        .slice(0, 3);
      const tagsHtml = tagList.map((t) => `<span class="law-tag-pill">${t}</span>`).join('');

      card.innerHTML = `
        <div class="law-card-header">
          <span class="law-card-badge">${doc.category || 'Luật'}</span>
          <span class="law-card-id">${doc.doc_id}</span>
        </div>
        <div class="law-card-title">${doc.title}</div>
        <div class="law-card-summary">${doc.raw_snippet || doc.summary || ''}</div>
        ${tagsHtml ? `<div class="law-card-tags">${tagsHtml}</div>` : ''}
      `;

      card.addEventListener('click', () => {
        if (activeArticleId === doc.doc_id) return;
        loadArticleDetail(doc.doc_id);
      });

      articlesContainer.appendChild(card);
    });

    refreshLucideIcons();
  }

  /**
   * 3. Load Article Detail
   */
  async function loadArticleDetail(docId) {
    activeArticleId = docId;

    // Highlight card in col 2
    const allCards = articlesContainer.querySelectorAll('.law-article-card');
    allCards.forEach((c) => {
      if (c.getAttribute('data-doc-id') === docId) {
        c.classList.add('active');
        c.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      } else {
        c.classList.remove('active');
      }
    });

    try {
      const res = await fetch(`/api/law/article?id=${encodeURIComponent(docId)}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const detail = await res.json();
      if (!detail.success) throw new Error(detail.error || 'Lỗi tải nội dung bài viết');

      currentArticleDetail = detail;

      // Update Reader UI
      if (readerEmpty) readerEmpty.style.display = 'none';
      if (readerContent) readerContent.style.display = 'block';

      // Title & Tags
      if (readerTitle) readerTitle.textContent = detail.title;
      if (readerTags) {
        const tagList = (detail.tags || '').split(' ').filter((t) => t.trim());
        readerTags.innerHTML = tagList.map((t) => `<span class="law-reader-tag">${t}</span>`).join('');
      }

      // Legal Source Banner
      if (readerSource) {
        if (detail.source) {
          readerSource.style.display = 'flex';
          readerSource.innerHTML = `<i data-lucide="scale"></i> <strong>Căn cứ pháp lý:</strong> <span>${detail.source}</span>`;
        } else {
          readerSource.style.display = 'none';
        }
      }

      // Parameters Inspector
      if (detail.law_params && detail.law_params.trim()) {
        if (paramsContainer) paramsContainer.style.display = 'block';
        if (paramsCode) paramsCode.textContent = detail.law_params.trim();
      } else {
        if (paramsContainer) paramsContainer.style.display = 'none';
      }

      // Main Markdown Content
      if (readerMarkdown) {
        readerMarkdown.innerHTML = renderMarkdownToHtml(detail.content);
      }

      refreshLucideIcons();
    } catch (err) {
      console.error('[LaborLawVault] Failed to load article detail:', err);
      showEmptyReader(`Lỗi tải bài viết: ${err.message}`);
    }
  }

  function showEmptyReader(msg) {
    if (readerContent) readerContent.style.display = 'none';
    if (readerEmpty) {
      readerEmpty.style.display = 'flex';
      const p = readerEmpty.querySelector('p');
      if (p && msg) p.textContent = msg;
    }
  }

  /**
   * 4. Actions: Copy YAML, Export Word & Export Vault
   */
  if (btnCopyParams) {
    btnCopyParams.addEventListener('click', async () => {
      if (!currentArticleDetail || !currentArticleDetail.law_params) return;
      try {
        await navigator.clipboard.writeText(currentArticleDetail.law_params.trim());
        showToast('Đã sao chép tham số cấu hình YAML vào clipboard!', 'success');
      } catch {
        // Fallback
        const textarea = document.createElement('textarea');
        textarea.value = currentArticleDetail.law_params.trim();
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        showToast('Đã sao chép tham số cấu hình YAML vào clipboard!', 'success');
      }
    });
  }

  if (btnExportDocx) {
    btnExportDocx.addEventListener('click', async () => {
      if (!activeArticleId) {
        showToast('Vui lòng chọn bài viết cần xuất Word!', 'warning');
        return;
      }

      try {
        btnExportDocx.disabled = true;
        btnExportDocx.innerHTML = '<div class="spinner"></div> Đang xuất Word...';

        const res = await fetch('/api/law/export-docx', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ id: activeArticleId }),
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data.success) throw new Error(data.error || 'Lỗi xuất Word');

        showToast(`Đã xuất Word thành công: ${data.filename}`, 'success');
        if (loadHistory) loadHistory();

        // Download file
        if (data.download_url) {
          const a = document.createElement('a');
          a.href = data.download_url;
          a.download = data.filename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }
      } catch (err) {
        console.error('[LaborLawVault] Export Word error:', err);
        showToast(`Lỗi xuất Word: ${err.message}`, 'error');
      } finally {
        btnExportDocx.disabled = false;
        btnExportDocx.innerHTML = '<i data-lucide="download"></i> Xuất Word (.docx)';
        refreshLucideIcons();
      }
    });
  }

  if (btnExportVault) {
    btnExportVault.addEventListener('click', async () => {
      try {
        btnExportVault.disabled = true;
        btnExportVault.innerHTML = '<div class="spinner"></div> Đang nén...';

        const res = await fetch('/api/law/export-vault', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({}),
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data.success) throw new Error(data.error || 'Lỗi đóng gói Vault');

        showToast(`Đã tải xuống Obsidian Vault: ${data.filename}`, 'success');
        if (loadHistory) loadHistory();

        if (data.download_url) {
          const a = document.createElement('a');
          a.href = data.download_url;
          a.download = data.filename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }
      } catch (err) {
        console.error('[LaborLawVault] Export Vault error:', err);
        showToast(`Lỗi đóng gói Vault: ${err.message}`, 'error');
      } finally {
        btnExportVault.disabled = false;
        btnExportVault.innerHTML = '<i data-lucide="archive"></i> Vault .zip';
        refreshLucideIcons();
      }
    });
  }

  const btnExportFullBook = document.getElementById('btn-law-export-full-book');
  if (btnExportFullBook) {
    btnExportFullBook.addEventListener('click', async () => {
      try {
        btnExportFullBook.disabled = true;
        btnExportFullBook.innerHTML = '<div class="spinner"></div> Đang biên dịch...';

        const res = await fetch('/api/law/export-full-book', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({}),
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data.success) throw new Error(data.error || 'Lỗi xuất cẩm nang Word');

        showToast(`Đã xuất trọn bộ Cẩm Nang (${data.total_articles} tài liệu): ${data.filename}`, 'success');
        if (loadHistory) loadHistory();

        if (data.download_url) {
          const a = document.createElement('a');
          a.href = data.download_url;
          a.download = data.filename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }
      } catch (err) {
        console.error('[LaborLawVault] Export Full Book error:', err);
        showToast(`Lỗi xuất Cẩm Nang Word: ${err.message}`, 'error');
      } finally {
        btnExportFullBook.disabled = false;
        btnExportFullBook.innerHTML = '<i data-lucide="book"></i> Toàn Bộ (.docx)';
        refreshLucideIcons();
      }
    });
  }

  /**
   * 5. Search Input Handling (250ms debounce)
   */
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      if (searchTimer) clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        loadArticles();
      }, 250);
    });

    searchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        searchInput.value = '';
        loadArticles();
      }
    });
  }

  if (btnClearSearch) {
    btnClearSearch.addEventListener('click', () => {
      if (searchInput) {
        searchInput.value = '';
        searchInput.focus();
        loadArticles();
      }
    });
  }

  // Initial load
  loadCategories();
  loadArticles();
}

