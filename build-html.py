#!/usr/bin/env python3
"""Bundle xlsx.mini.min.js into a single excel-compare.html."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "xlsx.mini.min.js"
OUT = ROOT / "excel-compare.html"

HTML_HEAD = r"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>مقایسه دو فایل اکسل</title>
  <style>
    :root {
      --bg: #0f1419;
      --surface: #1a2332;
      --border: #2d3a4d;
      --text: #e7ecf3;
      --muted: #8b9cb3;
      --accent: #3b82f6;
      --accent-hover: #2563eb;
      --danger: #ef4444;
      --warn: #f59e0b;
      --ok: #22c55e;
      --only-a: #7c3aed;
      --only-b: #0891b2;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Vazirmatn", "Segoe UI", Tahoma, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
      min-height: 100vh;
    }
    .wrap { max-width: 1100px; margin: 0 auto; padding: 1.5rem 1.25rem 3rem; }
    h1 { font-size: 1.5rem; font-weight: 700; margin: 0 0 0.35rem; }
    .subtitle { color: var(--muted); font-size: 0.9rem; margin-bottom: 1.5rem; }
    .card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1.25rem;
      margin-bottom: 1rem;
    }
    .grid { display: grid; gap: 1rem; }
    @media (min-width: 640px) { .grid-2 { grid-template-columns: 1fr 1fr; } }
    label { display: block; font-size: 0.85rem; color: var(--muted); margin-bottom: 0.35rem; }
    input[type="file"] {
      width: 100%;
      padding: 0.65rem;
      border: 1px dashed var(--border);
      border-radius: 8px;
      background: var(--bg);
      color: var(--text);
      cursor: pointer;
    }
    input[type="file"]::file-selector-button {
      margin-left: 0.75rem;
      padding: 0.4rem 0.75rem;
      border: none;
      border-radius: 6px;
      background: var(--accent);
      color: #fff;
      cursor: pointer;
      font-family: inherit;
    }
    select, input[type="number"], input[type="text"] {
      width: 100%;
      padding: 0.55rem 0.65rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: var(--bg);
      color: var(--text);
      font-family: inherit;
    }
    .file-meta { font-size: 0.8rem; color: var(--muted); margin-top: 0.35rem; min-height: 1.2rem; }
    .opts { display: grid; gap: 0.75rem; }
    @media (min-width: 640px) {
      .opts { grid-template-columns: 1fr 1fr 1fr; align-items: end; }
    }
  .check-row { display: flex; align-items: center; gap: 0.5rem; padding: 0.35rem 0; }
  .check-row input { width: auto; accent-color: var(--accent); }
    .check-row span { font-size: 0.9rem; color: var(--text); }
    .actions { display: flex; flex-wrap: wrap; gap: 0.75rem; margin-top: 1rem; align-items: center; }
    button {
      padding: 0.65rem 1.25rem;
      border: none;
      border-radius: 8px;
      font-family: inherit;
      font-size: 0.95rem;
      cursor: pointer;
      transition: background 0.15s;
    }
    button.primary { background: var(--accent); color: #fff; }
    button.primary:hover:not(:disabled) { background: var(--accent-hover); }
    button.secondary { background: transparent; color: var(--muted); border: 1px solid var(--border); }
    button:disabled { opacity: 0.45; cursor: not-allowed; }
    .stats {
      display: flex; flex-wrap: wrap; gap: 0.75rem; margin-bottom: 1rem;
    }
    .stat {
      flex: 1; min-width: 140px;
      padding: 0.75rem 1rem;
      border-radius: 8px;
      background: var(--bg);
      border: 1px solid var(--border);
    }
    .stat strong { display: block; font-size: 1.35rem; }
    .stat span { font-size: 0.8rem; color: var(--muted); }
    .stat.diff strong { color: var(--warn); }
    .stat.only-a strong { color: var(--only-a); }
    .stat.only-b strong { color: var(--only-b); }
    .table-wrap { overflow: auto; max-height: 55vh; border: 1px solid var(--border); border-radius: 8px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
    th, td { padding: 0.5rem 0.65rem; text-align: right; border-bottom: 1px solid var(--border); white-space: nowrap; }
    th { position: sticky; top: 0; background: #243044; z-index: 1; }
    tr:hover td { background: rgba(59, 130, 246, 0.06); }
    .tag {
      display: inline-block; padding: 0.15rem 0.45rem; border-radius: 4px; font-size: 0.75rem;
    }
    .tag.changed { background: rgba(245, 158, 11, 0.2); color: var(--warn); }
    .tag.only-a { background: rgba(124, 58, 237, 0.2); color: #c4b5fd; }
    .tag.only-b { background: rgba(8, 145, 178, 0.2); color: #67e8f9; }
    .val-a { color: #fca5a5; max-width: 200px; overflow: hidden; text-overflow: ellipsis; }
    .val-b { color: #86efac; max-width: 200px; overflow: hidden; text-overflow: ellipsis; }
    .empty { text-align: center; color: var(--muted); padding: 2rem; }
    .error { color: var(--danger); font-size: 0.9rem; margin-top: 0.5rem; }
    .hidden { display: none !important; }
    #filter { max-width: 280px; }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>مقایسه دو فایل اکسل</h1>
    <p class="subtitle">کاملاً در مرورگر اجرا می‌شود؛ فایل‌ها از دستگاه شما خارج نمی‌شوند.</p>

    <div class="card">
      <div class="grid grid-2">
        <div>
          <label for="fileA">فایل اول (.xlsx / .xls)</label>
          <input type="file" id="fileA" accept=".xlsx,.xls,.xlsm,.xlsb,.ods,.csv" />
          <div class="file-meta" id="metaA"></div>
        </div>
        <div>
          <label for="fileB">فایل دوم (.xlsx / .xls)</label>
          <input type="file" id="fileB" accept=".xlsx,.xls,.xlsm,.xlsb,.ods,.csv" />
          <div class="file-meta" id="metaB"></div>
        </div>
      </div>

      <div class="opts" style="margin-top: 1rem;">
        <div>
          <label for="sheetSelect">برگه (شیت)</label>
          <select id="sheetSelect" disabled>
            <option value="">— ابتدا هر دو فایل را انتخاب کنید —</option>
          </select>
        </div>
        <div>
          <label for="keyCol">ستون کلید (اختیاری، از ۱)</label>
          <input type="number" id="keyCol" min="1" placeholder="خالی = مقایسه سطر به سطر" />
        </div>
        <div>
          <label for="headerRow">ردیف عنوان (از ۱، ۰ = بدون عنوان)</label>
          <input type="number" id="headerRow" value="1" min="0" />
        </div>
      </div>
      <div class="check-row" style="margin-top: 0.75rem;">
        <input type="checkbox" id="ignoreCase" />
        <span>نادیده گرفتن بزرگی و کوچکی حروف</span>
      </div>
      <div class="check-row">
        <input type="checkbox" id="trimCells" checked />
        <span>حذف فاصله اضافه از ابتدا و انتهای متن</span>
      </div>

      <div class="actions">
        <button type="button" class="primary" id="btnCompare" disabled>مقایسه</button>
        <button type="button" class="secondary hidden" id="btnExport">دانلود گزارش CSV</button>
      </div>
      <div class="error" id="errBox" role="alert"></div>
    </div>

    <div class="card hidden" id="resultCard">
      <div class="stats" id="stats"></div>
      <div style="margin-bottom: 0.75rem;">
        <label for="filter">فیلتر نتایج</label>
        <input type="text" id="filter" placeholder="جستجو در شیت، سطر، ستون، مقدار..." />
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>نوع</th>
              <th>شیت</th>
              <th>سطر</th>
              <th>ستون</th>
              <th>فایل اول</th>
              <th>فایل دوم</th>
            </tr>
          </thead>
          <tbody id="diffBody"></tbody>
        </table>
      </div>
      <p class="empty hidden" id="noDiff">تفاوتی پیدا نشد؛ دو فایل در این شیت یکسان‌اند.</p>
    </div>
  </div>

  <script>
"""

HTML_FOOT = r"""
  </script>
  <script>
(function () {
  const fileA = document.getElementById('fileA');
  const fileB = document.getElementById('fileB');
  const metaA = document.getElementById('metaA');
  const metaB = document.getElementById('metaB');
  const sheetSelect = document.getElementById('sheetSelect');
  const keyCol = document.getElementById('keyCol');
  const headerRow = document.getElementById('headerRow');
  const ignoreCase = document.getElementById('ignoreCase');
  const trimCells = document.getElementById('trimCells');
  const btnCompare = document.getElementById('btnCompare');
  const btnExport = document.getElementById('btnExport');
  const errBox = document.getElementById('errBox');
  const resultCard = document.getElementById('resultCard');
  const stats = document.getElementById('stats');
  const diffBody = document.getElementById('diffBody');
  const noDiff = document.getElementById('noDiff');
  const filter = document.getElementById('filter');

  let wbA = null;
  let wbB = null;
  let lastRows = [];

  function clearError() {
    errBox.textContent = '';
  }

  function showError(msg) {
    errBox.textContent = msg;
  }

  function colLetter(n) {
    let s = '';
    n++;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function normalize(val) {
    if (val === null || val === undefined) return '';
    if (val instanceof Date) {
      return val.toISOString().slice(0, 19);
    }
    let s = String(val);
    if (trimCells.checked) s = s.trim();
    if (ignoreCase.checked) s = s.toLowerCase();
    return s;
  }

  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error('خواندن فایل ممکن نشد'));
      reader.readAsArrayBuffer(file);
    });
  }

  function parseBuffer(buf) {
    return XLSX.read(buf, { type: 'array', cellDates: true });
  }

  function matrixFromSheet(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  }

  function updateSheetOptions() {
    sheetSelect.innerHTML = '';
    if (!wbA || !wbB) {
      sheetSelect.disabled = true;
      sheetSelect.innerHTML = '<option value="">— ابتدا هر دو فایل را انتخاب کنید —</option>';
      btnCompare.disabled = true;
      return;
    }
    const namesA = new Set(wbA.SheetNames);
    const common = wbB.SheetNames.filter((n) => namesA.has(n));
    const options = common.length ? common : wbA.SheetNames;
    if (!common.length && wbA.SheetNames.length && wbB.SheetNames.length) {
      showError('هیچ شیت مشترکی نیست؛ شیت‌های فایل اول برای مقایسه نمایش داده می‌شوند.');
    } else {
      clearError();
    }
    options.forEach((name) => {
      const o = document.createElement('option');
      o.value = name;
      o.textContent = name;
      sheetSelect.appendChild(o);
    });
    sheetSelect.disabled = options.length === 0;
    btnCompare.disabled = options.length === 0;
  }

  async function loadSide(input, metaEl, which) {
    const file = input.files && input.files[0];
    if (!file) {
      metaEl.textContent = '';
      if (which === 'A') wbA = null;
      else wbB = null;
      updateSheetOptions();
      return;
    }
    metaEl.textContent = file.name + ' (' + formatSize(file.size) + ')';
    try {
      const buf = await readFile(file);
      const wb = parseBuffer(buf);
      if (which === 'A') wbA = wb;
      else wbB = wb;
      clearError();
    } catch (e) {
      showError(e.message || String(e));
      if (which === 'A') wbA = null;
      else wbB = null;
    }
    updateSheetOptions();
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function getHeaders(matrix, hRow) {
    if (!hRow || hRow < 1) return null;
    const idx = hRow - 1;
    if (idx >= matrix.length) return null;
    return matrix[idx].map((c, i) => (c !== '' && c != null ? String(c) : colLetter(i)));
  }

  function compareSheets(sheetName) {
    const matA = matrixFromSheet(wbA, sheetName);
    const matB = matrixFromSheet(wbB, sheetName);
    const hRow = parseInt(headerRow.value, 10) || 0;
    const headers = getHeaders(matA, hRow) || getHeaders(matB, hRow);
    const dataStart = hRow > 0 ? hRow : 0;
    const rowsA = matA.slice(dataStart);
    const rowsB = matB.slice(dataStart);

    const keyIdx = keyCol.value ? parseInt(keyCol.value, 10) - 1 : -1;
    const diffs = [];

    function pushDiff(type, rowHuman, colIdx, rawA, rawB) {
      const colLabel = headers && colIdx >= 0 && colIdx < headers.length
        ? headers[colIdx]
        : colLetter(colIdx);
      diffs.push({
        type,
        sheet: sheetName,
        row: rowHuman,
        col: colLabel,
        rawA: rawA === undefined || rawA === null ? '' : String(rawA),
        rawB: rawB === undefined || rawB === null ? '' : String(rawB),
      });
    }

    if (keyIdx >= 0) {
      const mapB = new Map();
      rowsB.forEach((row, i) => {
        const k = normalize(row[keyIdx]);
        if (k === '' && row.every((c) => normalize(c) === '')) return;
        if (!mapB.has(k)) mapB.set(k, { row, index: i });
      });
      const seenB = new Set();

      rowsA.forEach((rowA, i) => {
        const k = normalize(rowA[keyIdx]);
        if (k === '' && rowA.every((c) => normalize(c) === '')) return;
        const rowHuman = dataStart + i + 1;
        const hit = mapB.get(k);
        if (!hit) {
          pushDiff('only-a', rowHuman, keyIdx, rowA[keyIdx], '');
          rowA.forEach((cell, c) => {
            if (c !== '' && c != null) pushDiff('only-a', rowHuman, c, cell, '');
          });
          return;
        }
        seenB.add(k);
        const rowB = hit.row;
        const maxC = Math.max(rowA.length, rowB.length);
        for (let c = 0; c < maxC; c++) {
          const a = rowA[c];
          const b = rowB[c];
          if (normalize(a) !== normalize(b)) {
            pushDiff('changed', rowHuman, c, a, b);
          }
        }
      });

      rowsB.forEach((rowB, i) => {
        const k = normalize(rowB[keyIdx]);
        if (k === '' && rowB.every((c) => normalize(c) === '')) return;
        if (seenB.has(k)) return;
        const rowHuman = dataStart + i + 1;
        pushDiff('only-b', rowHuman, keyIdx, '', rowB[keyIdx]);
        rowB.forEach((cell, c) => {
          if (cell !== '' && cell != null) pushDiff('only-b', rowHuman, c, '', cell);
        });
      });
    } else {
      const maxRows = Math.max(rowsA.length, rowsB.length);
      let maxCols = 0;
      rowsA.forEach((r) => { maxCols = Math.max(maxCols, r.length); });
      rowsB.forEach((r) => { maxCols = Math.max(maxCols, r.length); });

      for (let r = 0; r < maxRows; r++) {
        const rowA = rowsA[r] || [];
        const rowB = rowsB[r] || [];
        const rowHuman = dataStart + r + 1;
        const onlyA = r >= rowsB.length;
        const onlyB = r >= rowsA.length;

        if (onlyA) {
          rowA.forEach((cell, c) => {
            if (normalize(cell) !== '') pushDiff('only-a', rowHuman, c, cell, '');
          });
          continue;
        }
        if (onlyB) {
          rowB.forEach((cell, c) => {
            if (normalize(cell) !== '') pushDiff('only-b', rowHuman, c, '', cell);
          });
          continue;
        }

        for (let c = 0; c < maxCols; c++) {
          const a = rowA[c];
          const b = rowB[c];
          if (normalize(a) !== normalize(b)) {
            pushDiff('changed', rowHuman, c, a, b);
          }
        }
      }
    }

    return diffs;
  }

  function typeLabel(t) {
    if (t === 'changed') return 'تغییر';
    if (t === 'only-a') return 'فقط فایل اول';
    if (t === 'only-b') return 'فقط فایل دوم';
    return t;
  }

  function typeClass(t) {
    if (t === 'changed') return 'changed';
    if (t === 'only-a') return 'only-a';
    if (t === 'only-b') return 'only-b';
    return '';
  }

  function renderStats(rows) {
    const changed = rows.filter((r) => r.type === 'changed').length;
    const oa = rows.filter((r) => r.type === 'only-a').length;
    const ob = rows.filter((r) => r.type === 'only-b').length;
    stats.innerHTML =
      '<div class="stat diff"><strong>' + changed + '</strong><span>سلول تغییر کرده</span></div>' +
      '<div class="stat only-a"><strong>' + oa + '</strong><span>فقط در فایل اول</span></div>' +
      '<div class="stat only-b"><strong>' + ob + '</strong><span>فقط در فایل دوم</span></div>' +
      '<div class="stat"><strong>' + rows.length + '</strong><span>جمع موارد</span></div>';
  }

  function renderTable(rows) {
    const q = (filter.value || '').trim().toLowerCase();
    const filtered = q
      ? rows.filter((r) =>
          [r.type, r.sheet, String(r.row), r.col, r.rawA, r.rawB].join(' ').toLowerCase().includes(q)
        )
      : rows;

    diffBody.innerHTML = '';
    filtered.forEach((r) => {
      const tr = document.createElement('tr');
      tr.innerHTML =
        '<td><span class="tag ' + typeClass(r.type) + '">' + typeLabel(r.type) + '</span></td>' +
        '<td>' + escapeHtml(r.sheet) + '</td>' +
        '<td>' + r.row + '</td>' +
        '<td>' + escapeHtml(r.col) + '</td>' +
        '<td class="val-a" title="' + escapeAttr(r.rawA) + '">' + escapeHtml(r.rawA) + '</td>' +
        '<td class="val-b" title="' + escapeAttr(r.rawB) + '">' + escapeHtml(r.rawB) + '</td>';
      diffBody.appendChild(tr);
    });

    noDiff.classList.toggle('hidden', rows.length !== 0);
    diffBody.parentElement.parentElement.classList.toggle('hidden', rows.length === 0);
  }

  function escapeHtml(s) {
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, '&#39;');
  }

  function runCompare() {
    clearError();
    if (!wbA || !wbB) {
      showError('هر دو فایل را انتخاب کنید.');
      return;
    }
    const sheet = sheetSelect.value;
    if (!sheet) {
      showError('شیت را انتخاب کنید.');
      return;
    }
    if (!wbB.Sheets[sheet]) {
      showError('شیت «' + sheet + '» در فایل دوم وجود ندارد.');
      return;
    }
    try {
      lastRows = compareSheets(sheet);
      resultCard.classList.remove('hidden');
      renderStats(lastRows);
      renderTable(lastRows);
      btnExport.classList.toggle('hidden', lastRows.length === 0);
    } catch (e) {
      showError(e.message || String(e));
    }
  }

  function exportCsv() {
    if (!lastRows.length) return;
    const header = ['نوع', 'شیت', 'سطر', 'ستون', 'فایل اول', 'فایل دوم'];
    const lines = [header.join(',')];
    lastRows.forEach((r) => {
      const line = [typeLabel(r.type), r.sheet, r.row, r.col, r.rawA, r.rawB]
        .map((cell) => '"' + String(cell).replace(/"/g, '""') + '"')
        .join(',');
      lines.push(line);
    });
    const blob = new Blob(['\ufeff' + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'excel-diff-report.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  fileA.addEventListener('change', () => loadSide(fileA, metaA, 'A'));
  fileB.addEventListener('change', () => loadSide(fileB, metaB, 'B'));
  btnCompare.addEventListener('click', runCompare);
  btnExport.addEventListener('click', exportCsv);
  filter.addEventListener('input', () => renderTable(lastRows));
})();
  </script>
</body>
</html>
"""

def main() -> None:
    xlsx_src = XLSX.read_text(encoding="utf-8", errors="surrogateescape")
    content = HTML_HEAD + xlsx_src + "\n" + HTML_FOOT
    OUT.write_text(content, encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()