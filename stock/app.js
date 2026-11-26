(function () {
  const STORAGE_KEY = 'stock-report-settings-v1';

  const fileProduct = document.getElementById('fileProduct');
  const fileInventory = document.getElementById('fileInventory');
  const metaProduct = document.getElementById('metaProduct');
  const metaInventory = document.getElementById('metaInventory');
  const sheetProduct = document.getElementById('sheetProduct');
  const sheetInventory = document.getElementById('sheetInventory');
  const headerRowProduct = document.getElementById('headerRowProduct');
  const headerRowInventory = document.getElementById('headerRowInventory');
  const colProductCode = document.getElementById('colProductCode');
  const colInventoryCode = document.getElementById('colInventoryCode');
  const colInventoryQty = document.getElementById('colInventoryQty');
  const btnRun = document.getElementById('btnRun');
  const btnPdf = document.getElementById('btnPdf');
  const btnXlsx = document.getElementById('btnXlsx');
  const errBox = document.getElementById('errBox');
  const resultCard = document.getElementById('resultCard');
  const previewBody = document.getElementById('previewBody');
  const statsEl = document.getElementById('stats');
  const pdfRoot = document.getElementById('pdfRoot');
  const pdfLogo = document.getElementById('pdfLogo');
  const pdfTitle = document.getElementById('pdfTitle');
  const pdfDate = document.getElementById('pdfDate');
  const pdfTableBody = document.getElementById('pdfTableBody');
  const pdfFooter = document.getElementById('pdfFooter');

  const companyName = document.getElementById('companyName');
  const companyPhone = document.getElementById('companyPhone');
  const companyAddress = document.getElementById('companyAddress');
  const reportTitle = document.getElementById('reportTitle');
  const logoFile = document.getElementById('logoFile');
  const logoPreview = document.getElementById('logoPreview');

  let wbProduct = null;
  let wbInventory = null;
  let lastResult = null;

  const QTY_NAMES = ['مانده', 'موجودی', 'تعداد', 'مقدار', 'qty', 'quantity', 'stock'];

  function clearError() {
    errBox.textContent = '';
  }

  function showError(msg) {
    errBox.textContent = msg;
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

  function matrixFrom(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  }

  function normalizeCode(val) {
    if (val === null || val === undefined) return '';
    let s = String(val).trim();
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    return s;
  }

  function parseQty(val) {
    if (val === null || val === undefined || val === '') return 0;
    const s = String(val).replace(/,/g, '').trim();
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  }

  function formatQty(n) {
    if (!Number.isFinite(n)) return '—';
    if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
    return String(n);
  }

  function formatMoney(val) {
    if (val === '' || val === null || val === undefined) return '';
    const n = Number(String(val).replace(/,/g, ''));
    if (!Number.isFinite(n)) return String(val);
    return n.toLocaleString('fa-IR');
  }

  function findHeaderRow(matrix, maxScan) {
    const limit = Math.min(matrix.length, maxScan || 15);
    for (let r = 0; r < limit; r++) {
      const row = matrix[r];
      if (!row) continue;
      const joined = row.map((c) => String(c)).join(' ');
      if (joined.includes('کد کالا') || joined.includes('کدکالا')) return r;
    }
    return 0;
  }

  function headerCells(matrix, headerIdx) {
    const row = matrix[headerIdx] || [];
    return row.map((c, i) => {
      const t = String(c || '').trim();
      return t || 'ستون ' + (i + 1);
    });
  }

  function indexByHeader(headers, names, fallbackIdx) {
    const lower = headers.map((h) => String(h).trim().toLowerCase());
    for (const name of names) {
      const idx = lower.findIndex((h) => h === name.toLowerCase() || h.includes(name));
      if (idx >= 0) return idx;
    }
    return fallbackIdx;
  }

  function fillSelect(select, headers, selectedIdx) {
    select.innerHTML = '';
    headers.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = h + ' (ستون ' + (i + 1) + ')';
      select.appendChild(opt);
    });
    if (selectedIdx >= 0 && selectedIdx < headers.length) {
      select.value = String(selectedIdx);
    }
  }

  function loadSettings() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const s = JSON.parse(raw);
      if (s.companyName) companyName.value = s.companyName;
      if (s.companyPhone) companyPhone.value = s.companyPhone;
      if (s.companyAddress) companyAddress.value = s.companyAddress;
      if (s.reportTitle) reportTitle.value = s.reportTitle;
      if (s.logoDataUrl) {
        logoPreview.src = s.logoDataUrl;
        logoPreview.classList.remove('hidden');
        pdfLogo.src = s.logoDataUrl;
        pdfLogo.classList.remove('hidden');
      }
    } catch (_) {
      /* ignore */
    }
  }

  function saveSettings() {
    const logoSrc = logoPreview.src && !logoPreview.classList.contains('hidden') ? logoPreview.src : '';
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        companyName: companyName.value,
        companyPhone: companyPhone.value,
        companyAddress: companyAddress.value,
        reportTitle: reportTitle.value,
        logoDataUrl: logoSrc.startsWith('data:') ? logoSrc : '',
      })
    );
  }

  function updatePdfMeta() {
    pdfTitle.textContent = reportTitle.value.trim() || 'گزارش موجودی لیست محصول';
    const parts = [];
    if (companyName.value.trim()) parts.push(companyName.value.trim());
    if (companyPhone.value.trim()) parts.push('تلفن: ' + companyPhone.value.trim());
    if (companyAddress.value.trim()) parts.push(companyAddress.value.trim());
    pdfFooter.textContent = parts.join(' · ');
    pdfDate.textContent = 'تاریخ گزارش: ' + new Intl.DateTimeFormat('fa-IR', {
      dateStyle: 'long',
      timeStyle: 'short',
    }).format(new Date());
  }

  async function loadWorkbook(input, meta, setter) {
    const file = input.files && input.files[0];
    if (!file) {
      meta.textContent = '';
      setter(null);
      syncRunButton();
      return;
    }
    meta.textContent = file.name + ' — ' + formatSize(file.size);
    try {
      const buf = await readFile(file);
      setter(parseBuffer(buf));
      clearError();
    } catch (e) {
      showError(e.message || String(e));
      setter(null);
    }
    syncRunButton();
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function refreshSheetSelect(wb, select) {
    select.innerHTML = '';
    if (!wb) {
      select.disabled = true;
      const o = document.createElement('option');
      o.textContent = '— فایل را انتخاب کنید —';
      select.appendChild(o);
      return;
    }
    wb.SheetNames.forEach((name) => {
      const o = document.createElement('option');
      o.value = name;
      o.textContent = name;
      select.appendChild(o);
    });
    select.disabled = false;
  }

  function configureColumns(which) {
    const wb = which === 'product' ? wbProduct : wbInventory;
  const sheetSel = which === 'product' ? sheetProduct : sheetInventory;
  const hRowInput = which === 'product' ? headerRowProduct : headerRowInventory;
  const codeSel = which === 'product' ? colProductCode : colInventoryCode;
  const qtySel = colInventoryQty;

    if (!wb) return;
    const matrix = matrixFrom(wb, sheetSel.value);
    const autoHeader = findHeaderRow(matrix);
    hRowInput.value = String(autoHeader + 1);
    const headers = headerCells(matrix, autoHeader);
    if (which === 'product') {
      const codeIdx = indexByHeader(headers, ['کد کالا', 'کدکالا', 'کد'], 1);
      fillSelect(codeSel, headers, codeIdx);
    } else {
      const codeIdx = indexByHeader(headers, ['کد کالا', 'کدکالا', 'کد'], 0);
      const qtyIdx = indexByHeader(headers, QTY_NAMES, 2);
      fillSelect(codeSel, headers, codeIdx);
      fillSelect(qtySel, headers, qtyIdx);
    }
  }

  function onSheetsReady() {
    refreshSheetSelect(wbProduct, sheetProduct);
    refreshSheetSelect(wbInventory, sheetInventory);
    if (wbProduct) configureColumns('product');
    if (wbInventory) configureColumns('inventory');
    syncRunButton();
  }

  function syncRunButton() {
    btnRun.disabled = !(wbProduct && wbInventory);
  }

  function buildInventoryMap() {
    const matrix = matrixFrom(wbInventory, sheetInventory.value);
    const headerIdx = Math.max(0, parseInt(headerRowInventory.value, 10) - 1);
    const headers = headerCells(matrix, headerIdx);
    const codeCol = parseInt(colInventoryCode.value, 10);
    const qtyCol = parseInt(colInventoryQty.value, 10);
    const map = new Map();

    for (let r = headerIdx + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row) continue;
      const code = normalizeCode(row[codeCol]);
      if (!code) continue;
      const qty = parseQty(row[qtyCol]);
      const prev = map.get(code) || 0;
      map.set(code, prev + qty);
    }

    return { map, headers, headerIdx, codeCol, qtyCol };
  }

  function mergeProductWithInventory() {
    const matrix = matrixFrom(wbProduct, sheetProduct.value);
    const headerIdx = Math.max(0, parseInt(headerRowProduct.value, 10) - 1);
    const headers = headerCells(matrix, headerIdx);
    const codeCol = parseInt(colProductCode.value, 10);
    const inv = buildInventoryMap();
    const outHeaders = headers.slice();
    if (!outHeaders.some((h) => String(h).includes('موجودی'))) {
      outHeaders.push('موجودی');
    }
    const qtyColOut = outHeaders.length - 1;

    const rows = [];
    const missing = [];

    for (let r = headerIdx + 1; r < matrix.length; r++) {
      const row = matrix[r].slice();
      const code = normalizeCode(row[codeCol]);
      const allEmpty = row.every((c) => normalizeCode(c) === '');
      if (!code && allEmpty) continue;

      let qty = 0;
      let found = false;
      if (code && inv.map.has(code)) {
        qty = inv.map.get(code);
        found = true;
      }

      while (row.length < outHeaders.length - 1) row.push('');
      row[qtyColOut] = found ? qty : '';
      rows.push({
        cells: row,
        code,
        qty,
        found,
        sourceRow: r + 1,
      });
      if (code && !found) missing.push(code);
    }

    return {
      headers: outHeaders,
      rows,
      missing,
      productSheet: sheetProduct.value,
      inventorySheet: sheetInventory.value,
    };
  }

  function renderPreview(result) {
    const headRow = document.querySelector('#previewHead tr');
    headRow.innerHTML = '';
    result.headers.forEach((h) => {
      const th = document.createElement('th');
      th.textContent = h;
      headRow.appendChild(th);
    });
    previewBody.innerHTML = '';
    const frag = document.createDocumentFragment();
    result.rows.forEach((row) => {
      const tr = document.createElement('tr');
      if (!row.found && row.code) tr.className = 'row-missing';
      result.headers.forEach((_, ci) => {
        const td = document.createElement('td');
        let val = row.cells[ci];
        if (result.headers[ci] === 'مبلغ' || (ci === 4 && String(result.headers[ci]).includes('مبلغ'))) {
          val = formatMoney(val);
        } else if (result.headers[ci] === 'موجودی') {
          val = row.found ? formatQty(row.qty) : '—';
        }
        td.textContent = val === undefined || val === null ? '' : String(val);
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });
    previewBody.appendChild(frag);

    const total = result.rows.length;
    const withStock = result.rows.filter((r) => r.found).length;
    const withoutStock = total - withStock;
    const sumQty = result.rows.reduce((a, r) => a + (r.found ? r.qty : 0), 0);
    statsEl.innerHTML =
      '<div class="stat"><strong>' + total + '</strong><span>ردیف لیست محصول</span></div>' +
      '<div class="stat ok"><strong>' + withStock + '</strong><span>با موجودی در انبار</span></div>' +
      '<div class="stat warn"><strong>' + withoutStock + '</strong><span>بدون ردیف موجودی</span></div>' +
      '<div class="stat"><strong>' + formatQty(sumQty) + '</strong><span>جمع موجودی (مطابق کد)</span></div>';
  }

  function fillPdfTable(result) {
    pdfTableBody.innerHTML = '';
    const headRow = document.getElementById('pdfTableHead');
    headRow.innerHTML = '';
    result.headers.forEach((h) => {
      const th = document.createElement('th');
      th.textContent = h;
      headRow.appendChild(th);
    });
    result.rows.forEach((row) => {
      const tr = document.createElement('tr');
      result.headers.forEach((h, ci) => {
        const td = document.createElement('td');
        let val = row.cells[ci];
        if (h === 'مبلغ' || String(h).includes('مبلغ')) val = formatMoney(val);
        else if (h === 'موجودی') val = row.found ? formatQty(row.qty) : '—';
        td.textContent = val === undefined || val === null ? '' : String(val);
        tr.appendChild(td);
      });
      pdfTableBody.appendChild(tr);
    });
  }

  function runMerge() {
    clearError();
    if (!wbProduct || !wbInventory) {
      showError('هر دو فایل (لیست محصول و موجودی) را انتخاب کنید.');
      return;
    }
    try {
      lastResult = mergeProductWithInventory();
      saveSettings();
      updatePdfMeta();
      renderPreview(lastResult);
      fillPdfTable(lastResult);
      resultCard.classList.remove('hidden');
      btnPdf.disabled = false;
      btnXlsx.disabled = false;
      if (lastResult.missing.length) {
        showError(
          lastResult.missing.length +
            ' کد کالا در فایل موجودی پیدا نشد (در جدول با خط نمایش داده می‌شود).'
        );
      } else {
        clearError();
      }
    } catch (e) {
      showError(e.message || String(e));
    }
  }

  function waitFrames(n) {
    return new Promise((resolve) => {
      let i = 0;
      function tick() {
        i += 1;
        if (i >= n) resolve();
        else requestAnimationFrame(tick);
      }
      requestAnimationFrame(tick);
    });
  }

  async function exportPdf() {
    if (!lastResult) return;
    updatePdfMeta();
    saveSettings();
    fillPdfTable(lastResult);

    const mask = document.createElement('div');
    mask.className = 'pdf-export-mask';
    mask.textContent = 'در حال ساخت PDF…';
    document.body.appendChild(mask);

    pdfRoot.removeAttribute('aria-hidden');
    pdfRoot.classList.add('pdf-exporting');

    const captureW = Math.max(pdfRoot.scrollWidth, pdfRoot.offsetWidth, 1);
    const captureH = Math.max(pdfRoot.scrollHeight, pdfRoot.offsetHeight, 1);

    const opt = {
      margin: [8, 8, 12, 8],
      filename: 'stock-report.pdf',
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        width: captureW,
        height: captureH,
        scrollX: 0,
        scrollY: 0,
        onclone: (clonedDoc) => {
          const node = clonedDoc.getElementById('pdfRoot');
          if (!node) return;
          node.style.opacity = '1';
          node.style.background = '#ffffff';
          node.style.position = 'relative';
          node.style.left = '0';
          node.style.top = '0';
          node.style.width = captureW + 'px';
          node.style.minHeight = captureH + 'px';
          node.style.height = 'auto';
          node.style.display = 'block';
          node.style.overflow = 'visible';
        },
      },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
      pagebreak: { mode: ['css', 'legacy'] },
    };

    btnPdf.disabled = true;
    try {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      await waitFrames(3);
      await html2pdf().set(opt).from(pdfRoot).save();
    } catch (e) {
      showError('ساخت PDF ناموفق بود: ' + (e.message || String(e)));
    } finally {
      pdfRoot.classList.remove('pdf-exporting');
      pdfRoot.setAttribute('aria-hidden', 'true');
      mask.remove();
      btnPdf.disabled = false;
    }
  }

  function exportXlsx() {
    if (!lastResult) return;
    const aoa = [lastResult.headers];
    lastResult.rows.forEach((r) => {
      const line = r.cells.slice();
      const qIdx = lastResult.headers.indexOf('موجودی');
      if (qIdx >= 0) line[qIdx] = r.found ? r.qty : '';
      aoa.push(line);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'گزارش');
    XLSX.writeFile(wb, 'stock-merged.xlsx');
  }

  logoFile.addEventListener('change', () => {
    const file = logoFile.files && logoFile.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      logoPreview.src = reader.result;
      logoPreview.classList.remove('hidden');
      pdfLogo.src = reader.result;
      pdfLogo.classList.remove('hidden');
      saveSettings();
    };
    reader.readAsDataURL(file);
  });

  [companyName, companyPhone, companyAddress, reportTitle].forEach((el) => {
    el.addEventListener('change', () => {
      updatePdfMeta();
      saveSettings();
    });
  });

  fileProduct.addEventListener('change', async () => {
    await loadWorkbook(fileProduct, metaProduct, (wb) => {
      wbProduct = wb;
      onSheetsReady();
    });
  });

  fileInventory.addEventListener('change', async () => {
    await loadWorkbook(fileInventory, metaInventory, (wb) => {
      wbInventory = wb;
      onSheetsReady();
    });
  });

  sheetProduct.addEventListener('change', () => configureColumns('product'));
  sheetInventory.addEventListener('change', () => configureColumns('inventory'));

  btnRun.addEventListener('click', runMerge);
  btnPdf.addEventListener('click', exportPdf);
  btnXlsx.addEventListener('click', exportXlsx);

  loadSettings();
  updatePdfMeta();
})();