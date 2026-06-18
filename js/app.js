/* BOQ Extractor – Application Controller */
(function () {
  'use strict';

  // ── State ──────────────────────────────────────────────────────────────────
  var state = {
    file: null,
    workbook: null,
    extractedRows: [],
    filteredRows: [],
    sortCol: null,
    sortDir: 'asc',
    page: 1,
    pageSize: 25,
    darkMode: false
  };

  var COLS = window.BOQExtractor.OUTPUT_COLUMNS;

  // ── DOM refs ───────────────────────────────────────────────────────────────
  var $ = function (id) { return document.getElementById(id); };

  var els = {
    // Upload
    dropzone:        $('dropzone'),
    fileInput:       $('fileInput'),
    fileBtn:         $('fileBtn'),
    fileMeta:        $('fileMeta'),
    fileName:        $('fileName'),
    fileSize:        $('fileSize'),
    removeFile:      $('removeFile'),
    // Options
    modeBasic:       $('modeBasic'),
    modeSteel:       $('modeSteel'),
    itemNumberInput: $('itemNumber'),
    itemNameWrap:    $('itemNameWrap'),
    itemNameInput:   $('itemName'),
    extractBtn:      $('extractBtn'),
    // Progress
    progressSection: $('progressSection'),
    progressBar:     $('progressBar'),
    progressPct:     $('progressPct'),
    progressMsg:     $('progressMsg'),
    spinner:         $('spinner'),
    // Results
    resultsSection:  $('resultsSection'),
    totalRowsBadge:  $('totalRowsBadge'),
    searchInput:     $('searchInput'),
    pageSizeSelect:  $('pageSizeSelect'),
    tableHead:       $('tableHead'),
    tableBody:       $('tableBody'),
    pagination:      $('pagination'),
    // Download
    downloadSection: $('downloadSection'),
    downloadXlsx:    $('downloadXlsx'),
    downloadCsv:     $('downloadCsv'),
    // Dark mode
    darkToggle:      $('darkToggle'),
    darkIcon:        $('darkIcon'),
    // Toast
    toastContainer:  $('toastContainer'),
    // Status
    statusArea:      $('statusArea')
  };

  // ── Dark Mode ──────────────────────────────────────────────────────────────
  function setDark(on) {
    state.darkMode = on;
    if (on) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('boq_dark', '1');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('boq_dark', '0');
    }
    els.darkIcon.innerHTML = on
      ? '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364-6.364l-.707.707M6.343 17.657l-.707.707M17.657 6.343l-.707-.707M6.343 6.343l-.707-.707M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>'
      : '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"/>';
  }

  els.darkToggle.addEventListener('click', function () { setDark(!state.darkMode); });

  // Restore preference
  setDark(localStorage.getItem('boq_dark') === '1' || window.matchMedia('(prefers-color-scheme: dark)').matches);

  // ── Toast ──────────────────────────────────────────────────────────────────
  function toast(msg, type) {
    type = type || 'info';
    var colours = {
      info:    'bg-indigo-600 text-white',
      success: 'bg-green-600 text-white',
      error:   'bg-red-600 text-white',
      warning: 'bg-yellow-500 text-gray-900'
    };
    var div = document.createElement('div');
    div.className = 'toast-in pointer-events-auto flex items-center gap-3 px-4 py-3 rounded-xl shadow-lg text-sm font-medium ' + (colours[type] || colours.info);
    div.textContent = msg;
    els.toastContainer.appendChild(div);
    setTimeout(function () {
      div.classList.remove('toast-in');
      div.classList.add('toast-out');
      setTimeout(function () { if (div.parentNode) div.parentNode.removeChild(div); }, 350);
    }, 3500);
  }

  // ── Status area ────────────────────────────────────────────────────────────
  function setStatus(msg, type) {
    var colours = { info: 'text-indigo-600 dark:text-indigo-400', error: 'text-red-600 dark:text-red-400', success: 'text-green-600 dark:text-green-400' };
    els.statusArea.className = 'mt-2 text-sm font-medium min-h-[1.25rem] ' + (colours[type] || colours.info);
    els.statusArea.textContent = msg;
  }

  // ── File upload ────────────────────────────────────────────────────────────
  function fmtBytes(b) {
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
    return (b / 1048576).toFixed(1) + ' MB';
  }

  function applyFile(file) {
    if (!file) return;
    var ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xlsx' && ext !== 'xls') {
      toast('Please upload an Excel file (.xlsx or .xls)', 'error');
      return;
    }
    state.file = file;
    state.workbook = null;
    state.extractedRows = [];
    state.filteredRows = [];
    clearResults();

    els.fileName.textContent = file.name;
    els.fileSize.textContent = fmtBytes(file.size);
    els.fileMeta.classList.remove('hidden');
    els.dropzone.classList.add('border-indigo-500', 'bg-indigo-50', 'dark:bg-indigo-950/30');
    els.dropzone.classList.remove('dropzone-pulse');
    setStatus('File ready — choose extraction options and click Extract.', 'info');
    toast('File loaded: ' + file.name, 'success');
  }

  els.fileBtn.addEventListener('click', function () { els.fileInput.click(); });
  els.fileInput.addEventListener('change', function () {
    if (els.fileInput.files[0]) applyFile(els.fileInput.files[0]);
  });

  els.removeFile.addEventListener('click', function (e) {
    e.stopPropagation();
    state.file = null;
    state.workbook = null;
    els.fileInput.value = '';
    els.fileMeta.classList.add('hidden');
    els.dropzone.classList.remove('border-indigo-500', 'bg-indigo-50', 'dark:bg-indigo-950/30');
    els.dropzone.classList.add('dropzone-pulse');
    clearResults();
    setStatus('', '');
  });

  // Drag & drop
  var dragCounter = 0;
  els.dropzone.addEventListener('click', function () { els.fileInput.click(); });
  els.dropzone.addEventListener('dragenter', function (e) { e.preventDefault(); dragCounter++; els.dropzone.classList.add('drag-over'); });
  els.dropzone.addEventListener('dragleave', function () { if (--dragCounter <= 0) { dragCounter = 0; els.dropzone.classList.remove('drag-over'); } });
  els.dropzone.addEventListener('dragover', function (e) { e.preventDefault(); });
  els.dropzone.addEventListener('drop', function (e) {
    e.preventDefault();
    els.dropzone.classList.remove('drag-over');
    dragCounter = 0;
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) applyFile(f);
  });

  // ── Mode toggle ────────────────────────────────────────────────────────────
  function updateModeUI() {
    var isSteel = els.modeSteel.checked;
    if (isSteel) {
      els.itemNameWrap.classList.remove('hidden');
    } else {
      els.itemNameWrap.classList.add('hidden');
    }
  }
  els.modeBasic.addEventListener('change', updateModeUI);
  els.modeSteel.addEventListener('change', updateModeUI);
  updateModeUI();

  // ── Progress helpers ───────────────────────────────────────────────────────
  function showProgress(show) {
    els.progressSection.classList.toggle('hidden', !show);
    els.spinner.classList.toggle('hidden', !show);
  }

  function setProgress(pct, msg) {
    pct = Math.max(0, Math.min(100, Math.round(pct)));
    els.progressBar.style.width = pct + '%';
    els.progressPct.textContent = pct + '%';
    if (msg) els.progressMsg.textContent = msg;
  }

  // ── Parse workbook (async / chunked via setTimeout to keep UI alive) ───────
  function parseWorkbook(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var data = new Uint8Array(e.target.result);
          // cellDates:true → date-formatted cells become JS Date objects (the
          // engine formats them to DD-MM-YYYY); numbers stay numeric for L/B/D.
          var wb = XLSX.read(data, { type: 'array', cellDates: true });
          resolve(wb);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = function () { reject(new Error('Failed to read file')); };
      reader.readAsArrayBuffer(file);
    });
  }

  function wbToSheets(wb) {
    return wb.SheetNames.map(function (name) {
      var ws = wb.Sheets[name];
      var arr = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      return { name: name, rows: arr };
    });
  }

  // ── Extract ────────────────────────────────────────────────────────────────
  els.extractBtn.addEventListener('click', function () { doExtract(); });

  async function doExtract() {
    if (!state.file) { toast('Please upload an Excel file first.', 'error'); return; }
    var itemNum = els.itemNumberInput.value.trim();
    if (!itemNum) { toast('Please enter an Item Number.', 'error'); els.itemNumberInput.focus(); return; }
    var isSteel = els.modeSteel.checked;
    var itemName = els.itemNameInput.value.trim(); // optional for Steel

    clearResults();
    showProgress(true);
    setProgress(0, 'Reading file…');
    setStatus('Extracting data…', 'info');
    els.extractBtn.disabled = true;

    try {
      // Parse the workbook
      setProgress(5, 'Parsing workbook…');
      var wb;
      if (state.workbook && state.workbook._file === state.file) {
        wb = state.workbook;
      } else {
        wb = await parseWorkbook(state.file);
        wb._file = state.file;
        state.workbook = wb;
      }

      setProgress(15, 'Converting sheets…');
      var sheets = wbToSheets(wb);
      var total = sheets.length;

      setProgress(20, 'Starting extraction across ' + total + ' sheet(s)…');

      // Run extraction with progress callbacks
      var result = await new Promise(function (resolve) {
        setTimeout(function () {
          var r = window.BOQExtractor.extractFromSheets(sheets, {
            mode: isSteel ? 'steel' : 'basic',
            itemNumber: itemNum,
            itemName: itemName,
            onProgress: function (pct, msg) {
              // Remap engine 0-100 to our 20-90 band
              var overall = 20 + Math.round(pct * 0.7);
              setProgress(overall, msg);
            }
          });
          resolve(r);
        }, 0);
      });

      setProgress(95, 'Rendering results…');

      state.extractedRows = result.rows;
      state.filteredRows = result.rows.slice();
      state.sortCol = null;
      state.sortDir = 'asc';
      state.page = 1;

      setTimeout(function () {
        renderTable();
        renderDownloadSection(isSteel, itemNum, itemName);
        setProgress(100, 'Done');

        if (result.rows.length === 0) {
          setStatus('No matching rows found. Check Item Number' + (isSteel ? ' and Item Name' : '') + '.', 'warning');
          toast('Extraction complete – no rows matched.', 'warning');
        } else {
          setStatus('Extraction complete — ' + result.rows.length + ' row(s) extracted from ' + result.stats.sheetsProcessed + ' sheet(s).', 'success');
          toast('Extracted ' + result.rows.length + ' row(s) successfully.', 'success');
        }

        setTimeout(function () { showProgress(false); }, 800);
        els.extractBtn.disabled = false;
      }, 50);

    } catch (err) {
      setProgress(0, '');
      showProgress(false);
      setStatus('Error: ' + err.message, 'error');
      toast('Extraction failed: ' + err.message, 'error');
      console.error(err);
      els.extractBtn.disabled = false;
    }
  }

  // ── Table rendering ────────────────────────────────────────────────────────
  function clearResults() {
    els.resultsSection.classList.add('hidden');
    els.downloadSection.classList.add('hidden');
    els.tableHead.innerHTML = '';
    els.tableBody.innerHTML = '';
    els.pagination.innerHTML = '';
  }

  function renderTable() {
    els.resultsSection.classList.remove('hidden');
    renderHeader();
    renderBody();
    renderPagination();
    els.totalRowsBadge.textContent = state.filteredRows.length + ' row(s)';
  }

  function renderHeader() {
    var tr = document.createElement('tr');
    COLS.forEach(function (col) {
      var th = document.createElement('th');
      th.setAttribute('data-col', col);
      th.className = 'px-3 py-2.5 text-left text-xs font-semibold uppercase tracking-wide text-gray-500 dark:text-gray-400 whitespace-nowrap select-none';
      if (state.sortCol === col) th.classList.add(state.sortDir);
      th.innerHTML = col + '<span class="sort-icon"></span>';
      th.addEventListener('click', function () { sortBy(col); });
      tr.appendChild(th);
    });
    els.tableHead.innerHTML = '';
    els.tableHead.appendChild(tr);
  }

  function renderBody() {
    var start = (state.page - 1) * state.pageSize;
    var end = Math.min(start + state.pageSize, state.filteredRows.length);
    var frag = document.createDocumentFragment();

    for (var i = start; i < end; i++) {
      var row = state.filteredRows[i];
      var tr = document.createElement('tr');
      tr.className = (i % 2 === 0)
        ? 'bg-white dark:bg-gray-900'
        : 'bg-gray-50 dark:bg-gray-800/60';

      COLS.forEach(function (col) {
        var td = document.createElement('td');
        td.className = 'px-3 py-2 text-xs text-gray-700 dark:text-gray-300 whitespace-nowrap max-w-xs truncate';
        var v = row[col];
        td.textContent = (v === null || v === undefined) ? '' : String(v);
        td.title = td.textContent;
        tr.appendChild(td);
      });

      frag.appendChild(tr);
    }

    els.tableBody.innerHTML = '';
    if (state.filteredRows.length === 0) {
      var empty = document.createElement('tr');
      var td = document.createElement('td');
      td.colSpan = COLS.length;
      td.className = 'px-4 py-10 text-center text-sm text-gray-400 dark:text-gray-500';
      td.textContent = 'No rows to display.';
      empty.appendChild(td);
      els.tableBody.appendChild(empty);
    } else {
      els.tableBody.appendChild(frag);
    }
  }

  function renderPagination() {
    var total = state.filteredRows.length;
    var totalPages = Math.max(1, Math.ceil(total / state.pageSize));
    els.pagination.innerHTML = '';

    if (totalPages <= 1) return;

    var wrap = document.createElement('div');
    wrap.className = 'flex items-center gap-1';

    function btn(label, pg, disabled, active) {
      var b = document.createElement('button');
      b.textContent = label;
      b.disabled = disabled;
      b.className = [
        'px-2.5 py-1 text-xs rounded-lg font-medium border transition',
        active
          ? 'bg-indigo-600 border-indigo-600 text-white'
          : 'border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700',
        disabled ? 'opacity-40 cursor-not-allowed' : ''
      ].join(' ');
      if (!disabled) b.addEventListener('click', function () { state.page = pg; renderBody(); renderPagination(); });
      return b;
    }

    wrap.appendChild(btn('«', 1, state.page === 1, false));
    wrap.appendChild(btn('‹', state.page - 1, state.page === 1, false));

    // Page window
    var lo = Math.max(1, state.page - 2);
    var hi = Math.min(totalPages, state.page + 2);
    if (lo > 1) { wrap.appendChild(btn('1', 1, false, false)); if (lo > 2) { var sp = document.createElement('span'); sp.className = 'px-1 text-xs text-gray-400'; sp.textContent = '…'; wrap.appendChild(sp); } }
    for (var p = lo; p <= hi; p++) wrap.appendChild(btn(String(p), p, false, p === state.page));
    if (hi < totalPages) { if (hi < totalPages - 1) { var sp2 = document.createElement('span'); sp2.className = 'px-1 text-xs text-gray-400'; sp2.textContent = '…'; wrap.appendChild(sp2); } wrap.appendChild(btn(String(totalPages), totalPages, false, false)); }

    wrap.appendChild(btn('›', state.page + 1, state.page === totalPages, false));
    wrap.appendChild(btn('»', totalPages, state.page === totalPages, false));

    // Info label
    var start = (state.page - 1) * state.pageSize + 1;
    var end = Math.min(state.page * state.pageSize, total);
    var info = document.createElement('span');
    info.className = 'ml-3 text-xs text-gray-400 dark:text-gray-500';
    info.textContent = start + '–' + end + ' of ' + total;

    els.pagination.appendChild(wrap);
    els.pagination.appendChild(info);
  }

  // ── Search ─────────────────────────────────────────────────────────────────
  els.searchInput.addEventListener('input', function () {
    var q = els.searchInput.value.toLowerCase();
    if (!q) {
      state.filteredRows = state.extractedRows.slice();
    } else {
      state.filteredRows = state.extractedRows.filter(function (row) {
        return COLS.some(function (col) {
          var v = row[col];
          return v !== null && v !== undefined && String(v).toLowerCase().indexOf(q) !== -1;
        });
      });
    }
    state.page = 1;
    renderHeader();
    renderBody();
    renderPagination();
    els.totalRowsBadge.textContent = state.filteredRows.length + ' row(s)';
  });

  // ── Page size ──────────────────────────────────────────────────────────────
  els.pageSizeSelect.addEventListener('change', function () {
    state.pageSize = parseInt(els.pageSizeSelect.value, 10);
    state.page = 1;
    renderBody();
    renderPagination();
  });

  // ── Sort ───────────────────────────────────────────────────────────────────
  function sortBy(col) {
    if (state.sortCol === col) {
      state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      state.sortCol = col;
      state.sortDir = 'asc';
    }
    state.filteredRows.sort(function (a, b) {
      var av = a[col], bv = b[col];
      if (av === null || av === undefined) av = '';
      if (bv === null || bv === undefined) bv = '';
      var n = typeof av === 'number' && typeof bv === 'number';
      var result = n ? (av - bv) : String(av).localeCompare(String(bv));
      return state.sortDir === 'asc' ? result : -result;
    });
    state.page = 1;
    renderHeader();
    renderBody();
    renderPagination();
  }

  // ── Download section ───────────────────────────────────────────────────────
  function buildFilename(isSteel, itemNum, itemName, ext) {
    var base = isSteel
      ? 'Steel_Extract_Item_' + itemNum + '_' + (itemName || '').replace(/[^a-zA-Z0-9]/g, '_')
      : 'Basic_Extract_Item_' + itemNum;
    return base + '.' + ext;
  }

  function renderDownloadSection(isSteel, itemNum, itemName) {
    if (state.extractedRows.length === 0) return;
    els.downloadSection.classList.remove('hidden');
    els.downloadXlsx.onclick = function () { exportXlsx(isSteel, itemNum, itemName); };
    els.downloadCsv.onclick  = function () { exportCsv(isSteel, itemNum, itemName); };
  }

  function exportXlsx(isSteel, itemNum, itemName) {
    var ws = XLSX.utils.json_to_sheet(state.extractedRows, { header: COLS });

    // Column widths
    ws['!cols'] = COLS.map(function (c) { return { wch: c === 'Description' ? 40 : 14 }; });

    // Style header row bold (basic styling via cell metadata)
    COLS.forEach(function (col, i) {
      var addr = XLSX.utils.encode_cell({ r: 0, c: i });
      if (ws[addr]) {
        ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: 'E0E7FF' } } };
      }
    });

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Extract');
    XLSX.writeFile(wb, buildFilename(isSteel, itemNum, itemName, 'xlsx'));
    toast('Excel file downloaded.', 'success');
  }

  function exportCsv(isSteel, itemNum, itemName) {
    var lines = [COLS.join(',')];
    state.extractedRows.forEach(function (row) {
      var line = COLS.map(function (c) {
        var v = row[c] === null || row[c] === undefined ? '' : String(row[c]);
        // Wrap in quotes if contains comma, quote, or newline
        if (/[",\n\r]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
        return v;
      });
      lines.push(line.join(','));
    });
    var blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = buildFilename(isSteel, itemNum, itemName, 'csv');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast('CSV file downloaded.', 'success');
  }

  // ── Initial state ──────────────────────────────────────────────────────────
  setStatus('Upload an Excel file to get started.', 'info');
  els.dropzone.classList.add('dropzone-pulse');

})();
