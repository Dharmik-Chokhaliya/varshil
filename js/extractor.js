/*
 * BOQ Extractor - Extraction Engine
 * ---------------------------------
 * Pure, dependency-free extraction logic for RA Bill / BOQ measurement sheets.
 *
 * The engine operates on plain data ("sheets") so it can run identically in the
 * browser (fed by SheetJS) and in Node (fed by synthetic fixtures for tests).
 *
 *   sheets : Array<{ name: string, rows: any[][] }>
 *
 * Public API:
 *   BOQExtractor.extractFromSheets(sheets, options) -> { rows, stats }
 *
 * options:
 *   mode        : 'basic' | 'steel'
 *   itemNumber  : string | number   (e.g. "28")
 *   itemName    : string            (Steel only, e.g. "Fe500")
 *   onProgress  : (percent, message, sheetName) => void   (optional)
 */
(function (global) {
  'use strict';

  // -------------------------------------------------------------------------
  // Output column order (kept in one place so export + preview stay in sync)
  // -------------------------------------------------------------------------
  var OUTPUT_COLUMNS = [
    'RecordDate',
    'Description',
    'No',
    'Length',
    'Breadth',
    'Depth',
    'Remarks',
    'ItemType'
  ];

  // -------------------------------------------------------------------------
  // Small value helpers
  // -------------------------------------------------------------------------

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  function cleanText(v) {
    if (isBlank(v)) return '';
    return String(v).replace(/\s+/g, ' ').trim();
  }

  // True when a value represents a real number (incl. numeric strings like "10.50").
  function isNumericVal(v) {
    if (v === null || v === undefined) return false;
    if (typeof v === 'number') return !isNaN(v);
    if (typeof v === 'string') {
      var s = v.trim();
      if (s === '') return false;
      // Allow leading/trailing spaces and thousands separators e.g. "1,250.00"
      var cleaned = s.replace(/,/g, '');
      return /^[-+]?\d*\.?\d+$/.test(cleaned);
    }
    return false;
  }

  // Convert a measurement cell to a number when possible, otherwise keep the
  // trimmed string, otherwise return '' for blanks.
  function numOrBlank(v) {
    if (isBlank(v)) return '';
    if (typeof v === 'number') return v;
    var s = String(v).trim();
    var cleaned = s.replace(/,/g, '');
    if (/^[-+]?\d*\.?\d+$/.test(cleaned)) return parseFloat(cleaned);
    return s;
  }

  // Preserve a "No." value exactly. Numeric strings become numbers (so Excel
  // keeps them numeric) but we never renumber.
  function preserveNo(v) {
    if (isBlank(v)) return '';
    if (typeof v === 'number') return v;
    var s = String(v).trim();
    var cleaned = s.replace(/,/g, '');
    if (/^[-+]?\d*\.?\d+$/.test(cleaned)) return parseFloat(cleaned);
    return s;
  }

  function normalizeIntStr(v) {
    if (isBlank(v)) return null;
    var s = String(v).trim();
    var m = s.match(/^0*(\d+)$/);
    if (m) return String(parseInt(m[1], 10));
    return null;
  }

  // Aggressively normalize text for fuzzy matching (item names, grades, etc.)
  function normalize(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  // -------------------------------------------------------------------------
  // Date handling -> "DD-MM-YYYY" (time stripped)
  // -------------------------------------------------------------------------

  function pad2(n) { return n < 10 ? '0' + n : String(n); }

  function fmtDate(d) {
    return pad2(d.getDate()) + '-' + pad2(d.getMonth() + 1) + '-' + d.getFullYear();
  }

  function excelSerialToDate(serial) {
    // Excel day 0 is 1899-12-30. Build the date in LOCAL time (start at the
    // local epoch and add whole days) so that getDate()/getMonth()/getFullYear()
    // always return the correct calendar date regardless of timezone.
    var whole = Math.floor(serial);
    var d = new Date(1899, 11, 30);
    d.setDate(d.getDate() + whole);
    return isNaN(d.getTime()) ? null : d;
  }

  var MONTHS = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12
  };

  function monthNum(name) {
    var k = String(name).toLowerCase().slice(0, 4);
    if (MONTHS[k] != null) return MONTHS[k];
    k = k.slice(0, 3);
    return MONTHS[k] != null ? MONTHS[k] : null;
  }

  // Normalise a (time-stripped) date string to DD-MM-YYYY when recognisable.
  function reformatDateString(s) {
    var m;
    // ISO: YYYY-MM-DD
    m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (m) return pad2(+m[3]) + '-' + pad2(+m[2]) + '-' + m[1];
    // DD-MM-YYYY (day first – the RA Bill standard)
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (m) return pad2(+m[1]) + '-' + pad2(+m[2]) + '-' + m[3];
    // DD-MM-YY
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
    if (m) { var yy = +m[3]; return pad2(+m[1]) + '-' + pad2(+m[2]) + '-' + (yy < 70 ? 2000 + yy : 1900 + yy); }
    // DD-MMM-YYYY (e.g. 15-May-2026 / 15 May 2026)
    m = s.match(/^(\d{1,2})[-/ ]([A-Za-z]{3,9})[-/ ,]+(\d{4})$/);
    if (m) { var mo = monthNum(m[2]); if (mo) return pad2(+m[1]) + '-' + pad2(mo) + '-' + m[3]; }
    // MMM DD, YYYY (e.g. May 15, 2026)
    m = s.match(/^([A-Za-z]{3,9})[-/ ]+(\d{1,2})[-/ ,]+(\d{4})$/);
    if (m) { var mo2 = monthNum(m[1]); if (mo2) return pad2(+m[2]) + '-' + pad2(mo2) + '-' + m[3]; }
    return s; // unrecognised – return unchanged
  }

  // Scan ANY row (including pre-colMap rows) for a recognizable date value.
  // Only accepts Date objects, date-like strings, and Excel serials in the
  // year-2000+ range — avoids false positives from measurement numbers.
  function extractAnyDate(row) {
    for (var i = 0; i < row.length; i++) {
      var v = row[i];
      if (isBlank(v)) continue;

      if (v instanceof Date) {
        return isNaN(v.getTime()) ? '' : fmtDate(v);
      }

      if (typeof v === 'number') {
        // Only accept Excel serials for 2000-01-01 (36526) through 2099-12-31 (73050)
        if (v > 36526 && v < 73051) {
          var dn = excelSerialToDate(v);
          if (dn) return fmtDate(dn);
        }
        continue; // skip ordinary numbers (quantities, item nos, measurements)
      }

      if (typeof v === 'string') {
        var str = v.trim();
        if (!str) continue;
        // Strip a leading label like "Date:", "Record Date:", etc.
        str = str.replace(/^(record\s*date|measurement\s*date|meas\s*date|date|dt|dated)\s*[:：]\s*/i, '');
        // Strip trailing time
        str = str.replace(/[ T]+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*[Zz]?\s*([AaPp]\.?\s*[Mm]\.?)?\s*$/, '').trim();
        if (!str) continue;
        var result = reformatDateString(str);
        if (/^\d{2}-\d{2}-\d{4}$/.test(result)) return result;
      }
    }
    return '';
  }

  function formatDate(v) {
    if (isBlank(v)) return '';

    if (v instanceof Date) {
      return isNaN(v.getTime()) ? '' : fmtDate(v);
    }

    if (typeof v === 'number') {
      // Plausible Excel date serials only; otherwise leave the number as text.
      if (v > 59 && v < 600000) {
        var dn = excelSerialToDate(v);
        if (dn) return fmtDate(dn);
      }
      return String(v);
    }

    var s = String(v).trim();
    if (!s) return '';

    // Pure-number text that is an Excel serial date (e.g. "45801")
    if (/^\d+(\.\d+)?$/.test(s)) {
      var n = parseFloat(s);
      if (n > 59 && n < 600000) {
        var ds = excelSerialToDate(n);
        if (ds) return fmtDate(ds);
      }
      return s;
    }

    // Strip a trailing time portion in any common form:
    //   "15-05-2026 12:00 PM", "... 00:00:00", "2026-05-15T00:00:00(.000)(Z)"
    s = s.replace(/[ T]+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*[Zz]?\s*([AaPp]\.?\s*[Mm]\.?)?\s*$/, '').trim();

    // Normalise to DD-MM-YYYY.
    return reformatDateString(s);
  }

  // -------------------------------------------------------------------------
  // Total / summary detection
  // -------------------------------------------------------------------------

  var TOTAL_RE = /\b(grand\s*total|sub[\s-]*total|total|totel|summary|carried\s*(over|forward)|brought\s*forward|b\/?f|c\/?o|c\/?f)\b/i;

  function isTotalRow(text) {
    if (!text) return false;
    return TOTAL_RE.test(text);
  }

  // "Say 208.00" style rounding rows.
  function isSayRow(text) {
    if (!text) return false;
    return /^\s*say\b/i.test(text);
  }

  // -------------------------------------------------------------------------
  // Column header classification
  // -------------------------------------------------------------------------

  function classifyHeaderCell(text) {
    if (isBlank(text)) return null;
    var t = String(text).toLowerCase().replace(/[.:#()]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!t) return null;

    // Description-like (check first so "item of work" isn't read as the item-no column)
    if (/(descrip|particular|item of work|nature of work|name of work|details of)/.test(t)) return 'description';

    // Serial number column (must NOT be treated as the measurement "No.")
    if (/^(sr no|sr|s no|sl no|sl|sno|srno|serial no|serial|r no)$/.test(t)) return 'serial';

    // Item number column
    if (/^(item|item no|item nos|item number|item code|boq no|boq item)$/.test(t)) return 'itemno';

    if (/^(date|dt|dated|record date|recorddate|rec date|rec dt|date of record|measurement date|meas date)$/.test(t)) return 'date';

    // Measurement "No." / "Nos" (count)
    if (/^(no|nos|no of|nos of|number|qnty no|qty no)$/.test(t)) return 'no';

    if (/^(l|length|len)$/.test(t)) return 'length';
    if (/^(b|breadth|width|w|brth)$/.test(t)) return 'breadth';
    if (/^(d|depth|height|ht|h|dia depth)$/.test(t)) return 'depth';

    if (/^(qty|quantity|quantities|content|contents|cnt|total qty|quantity in)$/.test(t)) return 'qty';

    return null;
  }

  function detectColumns(row) {
    var map = {};
    for (var i = 0; i < row.length; i++) {
      var type = classifyHeaderCell(row[i]);
      if (type && !(type in map)) map[type] = i;
    }
    return map;
  }

  var MEASURE_KEYS = ['length', 'breadth', 'depth', 'no', 'qty', 'date'];

  function isHeaderRow(row) {
    var map = detectColumns(row);
    var keys = Object.keys(map);
    if (keys.length === 0) return false;

    var hasMeasure = MEASURE_KEYS.some(function (k) { return k in map; });
    var hasDesc = 'description' in map;

    if (hasDesc && hasMeasure) return true;
    // L/B/D trio is a very strong signal even without an explicit description label
    if ('length' in map && 'breadth' in map && 'depth' in map) return true;
    if (keys.length >= 3 && hasMeasure) return true;
    return false;
  }

  // When the header row had no recognizable "description" column, guess it from
  // the data below: the column (not already mapped to a known field) that most
  // often carries alphabetic text.
  function guessDescCol(rows, startIdx, colMap) {
    var used = {};
    Object.keys(colMap).forEach(function (k) { used[colMap[k]] = true; });

    var scores = {};
    var limit = Math.min(rows.length, startIdx + 40);
    for (var r = startIdx; r < limit; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (used[c]) continue;
        var v = row[c];
        if (!isBlank(v) && /[a-zA-Z]{3,}/.test(String(v))) {
          scores[c] = (scores[c] || 0) + String(v).length;
        }
      }
    }

    var best = null, bestScore = -1;
    Object.keys(scores).forEach(function (c) {
      if (scores[c] > bestScore) { bestScore = scores[c]; best = parseInt(c, 10); }
    });
    return best;
  }

  // -------------------------------------------------------------------------
  // Item-section markers
  // -------------------------------------------------------------------------

  function rowText(row) {
    var parts = [];
    for (var i = 0; i < row.length; i++) {
      if (!isBlank(row[i])) parts.push(String(row[i]));
    }
    return parts.join(' ');
  }

  function isEmptyRow(row) {
    if (!row || row.length === 0) return true;
    for (var i = 0; i < row.length; i++) {
      if (!isBlank(row[i])) return false;
    }
    return true;
  }

  // Explicit "Item No. 28" style marker (strong, unambiguous).
  function detectStrongItemMarker(row) {
    for (var i = 0; i < row.length; i++) {
      if (isBlank(row[i])) continue;
      var s = String(row[i]).trim();
      var m = s.match(/\bitem\s*(?:no\.?|nos\.?|number|code|#)?\s*[:\-.]?\s*0*(\d+)\b/i);
      if (m) return { number: String(parseInt(m[1], 10)), text: rowText(row) };
    }
    return null;
  }

  // Weak marker: a row whose first non-empty cell is a bare integer followed by
  // descriptive text and NO measurement values. Used when item labels are not
  // prefixed with the word "Item".
  function detectWeakItemMarker(row, colMap) {
    var firstIdx = -1;
    for (var i = 0; i < row.length; i++) {
      if (!isBlank(row[i])) { firstIdx = i; break; }
    }
    if (firstIdx < 0) return null;

    var first = String(row[firstIdx]).trim();
    // Accept "28", "28.", "28)" as item labels.
    var m = first.match(/^0*(\d+)[).]?$/);
    if (!m) return null;

    // Must contain descriptive text somewhere after the number.
    var hasText = false;
    for (var j = firstIdx + 1; j < row.length; j++) {
      if (!isBlank(row[j]) && /[a-zA-Z]{3,}/.test(String(row[j]))) { hasText = true; break; }
    }
    if (!hasText) return null;

    // Reject if any known measurement column holds a value (i.e. it's a data row).
    if (colMap) {
      var measureCols = ['no', 'length', 'breadth', 'depth', 'qty'];
      for (var k = 0; k < measureCols.length; k++) {
        var ci = colMap[measureCols[k]];
        if (ci != null && isNumericVal(row[ci])) return null;
      }
    }

    return { number: String(parseInt(m[1], 10)), text: rowText(row) };
  }

  // -------------------------------------------------------------------------
  // Diameter detection (Steel mode)
  // -------------------------------------------------------------------------

  function formatDia(n) {
    var num = parseFloat(n);
    if (isNaN(num)) return '';
    return num + ' mm';
  }

  function detectDiameter(text) {
    if (isBlank(text)) return '';
    var s = String(text);

    // "dia 8", "Ø8", "diameter: 12 mm"
    var m = s.match(/(?:dia\.?|diameter|ø|φ|⌀)\s*[:\-]?\s*0*(\d+(?:\.\d+)?)\s*(?:mm)?/i);
    if (m) return formatDia(m[1]);

    // "8 mm", "8mm", "10 m m"
    m = s.match(/\b0*(\d+(?:\.\d+)?)\s*m\s*m\b/i);
    if (m) return formatDia(m[1]);

    return '';
  }

  // True when a header cell is essentially just a diameter (e.g. "8 mm",
  // "Dia 10 mm") rather than a descriptive section header.
  function isMostlyDiameter(text) {
    if (isBlank(text)) return false;
    var stripped = String(text)
      .replace(/(?:dia\.?|diameter|ø|φ|⌀)/gi, ' ')
      .replace(/\d+(?:\.\d+)?/g, ' ')
      .replace(/m\s*m/gi, ' ')
      .replace(/[^a-z0-9]/gi, '');
    return stripped.length <= 2;
  }

  // -------------------------------------------------------------------------
  // Core extraction
  // -------------------------------------------------------------------------

  function extractFromSheets(sheets, options) {
    options = options || {};
    var mode = options.mode === 'steel' ? 'steel' : 'basic';
    var targetItem = normalizeIntStr(options.itemNumber);
    if (targetItem === null && !isBlank(options.itemNumber)) {
      targetItem = String(options.itemNumber).trim();
    }
    // Item Name is OPTIONAL for Steel: when omitted, match by item number alone.
    var itemNameNorm = (mode === 'steel' && !isBlank(options.itemName)) ? normalize(options.itemName) : '';
    var onProgress = typeof options.onProgress === 'function' ? options.onProgress : null;

    var out = [];
    var stats = {
      totalRows: 0,
      sheetsProcessed: 0,
      sheetsTotal: sheets.length,
      sectionsFound: 0
    };

    for (var si = 0; si < sheets.length; si++) {
      var sheet = sheets[si] || {};
      var rows = sheet.rows || [];

      processSheet(rows, {
        mode: mode,
        targetItem: targetItem,
        itemNameNorm: itemNameNorm
      }, out, stats);

      stats.sheetsProcessed++;
      if (onProgress) {
        var pct = Math.round((stats.sheetsProcessed / Math.max(1, stats.sheetsTotal)) * 100);
        onProgress(pct, 'Processed sheet "' + (sheet.name || ('#' + (si + 1))) + '"', sheet.name);
      }
    }

    stats.totalRows = out.length;
    return { rows: out, stats: stats };
  }

  function processSheet(rows, cfg, out, stats) {
    var colMap = null;
    var descCol = null;
    var currentItem = null;
    var sectionText = '';
    var pendingHeader = '';
    var currentDiameter = '';
    var lastDate = '';
    var sawTargetSection = false;

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];

      if (isEmptyRow(row)) {
        // Blank rows do NOT clear a pending header — a header must still reach
        // the first subsequent data row even across blank separators.
        continue;
      }

      // 1. Strong item marker ("Item No. 28")
      var strong = detectStrongItemMarker(row);
      if (strong) {
        currentItem = strong.number;
        sectionText = strong.text;
        pendingHeader = '';
        currentDiameter = '';
        lastDate = '';
        if (currentItem === String(cfg.targetItem)) {
          sawTargetSection = true;
          stats.sectionsFound++;
        }
        continue;
      }

      // 2. Table header row
      if (isHeaderRow(row)) {
        colMap = detectColumns(row);
        descCol = (colMap.description != null) ? colMap.description : guessDescCol(rows, r + 1, colMap);
        pendingHeader = '';
        continue;
      }

      // 3. Weak item marker ("28  Providing PCC ...")
      var weak = detectWeakItemMarker(row, colMap);
      if (weak) {
        currentItem = weak.number;
        sectionText = weak.text;
        pendingHeader = '';
        currentDiameter = '';
        lastDate = '';
        if (currentItem === String(cfg.targetItem)) {
          sawTargetSection = true;
          stats.sectionsFound++;
        }
        continue;
      }

      // Need a table context to interpret the row.
      // Before the table header is found, still scan for standalone date rows
      // (e.g. "Record Date: 15/05/2026" written above the measurement table).
      if (!colMap) {
        var earlyDate = extractAnyDate(row);
        if (earlyDate) lastDate = earlyDate;
        continue;
      }

      // Resolve the item this row belongs to (column value wins when present).
      var rowItem = currentItem;
      if (colMap.itemno != null) {
        var iv = normalizeIntStr(row[colMap.itemno]);
        if (iv != null) { rowItem = iv; currentItem = iv; }
      }

      var desc = (descCol != null) ? cleanText(row[descCol]) : '';
      var fullText = rowText(row);

      // Carry the most recent date forward within the section.
      if (colMap.date != null) {
        var d = formatDate(row[colMap.date]);
        if (d) lastDate = d;
      }

      // Measurement presence
      var noVal = colMap.no != null ? row[colMap.no] : null;
      var lVal = colMap.length != null ? row[colMap.length] : null;
      var bVal = colMap.breadth != null ? row[colMap.breadth] : null;
      var dVal = colMap.depth != null ? row[colMap.depth] : null;
      var qVal = colMap.qty != null ? row[colMap.qty] : null;

      var hasDim = isNumericVal(noVal) || isNumericVal(lVal) || isNumericVal(bVal) || isNumericVal(dVal);
      var hasQty = isNumericVal(qVal);

      // 4. TOTAL / Grand Total / Say rows  (skipped, but do not consume a header)
      if (isTotalRow(fullText) || isTotalRow(desc) || isSayRow(desc)) {
        continue;
      }

      // 5. Quantity-summary row (only a quantity, no dimensions / no count)
      if (!hasDim && hasQty) {
        continue;
      }

      // 6. Section header (descriptive text, no measurements)
      if (!hasDim) {
        // In real RA Bills, section headers often use merged cells that SheetJS
        // places in column 0 — but descCol may point to a different column.
        // Fall back to the first alphabetic text in any non-measurement column.
        if (!desc) {
          var skipMeas = {};
          ['no', 'length', 'breadth', 'depth', 'qty'].forEach(function (k) {
            if (colMap[k] != null) skipMeas[colMap[k]] = true;
          });
          for (var hc = 0; hc < row.length; hc++) {
            if (skipMeas[hc]) continue;
            var hv = cleanText(row[hc]);
            if (hv.length > 2 && /[a-zA-Z]{2,}/.test(hv)) { desc = hv; break; }
          }
        }
        if (desc) {
          sectionText += ' ' + desc;
          if (cfg.mode === 'steel') {
            var dia = detectDiameter(desc);
            if (dia && isMostlyDiameter(desc)) {
              // A pure diameter header feeds Remarks, NOT the description merge,
              // and must leave any pending section header untouched.
              currentDiameter = dia;
            } else {
              // Latest header REPLACES any previous pending header.
              pendingHeader = desc;
              if (dia) currentDiameter = dia;
            }
          } else {
            // Latest header REPLACES any previous pending header.
            pendingHeader = desc;
          }
        }
        continue;
      }

      // 7. Data row -------------------------------------------------------
      // Header Merge Rule: the FIRST data row after a header consumes it.
      // Merge once, then clear immediately so it is never reused.
      var finalDesc = desc;
      if (pendingHeader) {
        finalDesc = desc ? (pendingHeader + ' - ' + desc) : pendingHeader;
        pendingHeader = '';
      }

      // Item-number filter
      if (String(rowItem) !== String(cfg.targetItem)) {
        continue;
      }

      // Steel: optional item-name (grade) filter (skipped when no name is given)
      if (cfg.mode === 'steel' && cfg.itemNameNorm) {
        var hay = normalize(sectionText + ' ' + fullText);
        if (hay.indexOf(cfg.itemNameNorm) === -1) {
          continue;
        }
      }

      var remarks = '';
      if (cfg.mode === 'steel') {
        var rowDia = detectDiameter(fullText);
        if (rowDia) currentDiameter = rowDia;
        remarks = currentDiameter || '';
      }

      out.push({
        RecordDate: lastDate || '',
        Description: finalDesc,
        No: cfg.mode === 'steel' ? 1 : preserveNo(noVal),
        Length: numOrBlank(lVal),
        Breadth: numOrBlank(bVal),
        Depth: numOrBlank(dVal),
        Remarks: remarks,
        ItemType: cfg.mode === 'steel' ? 'Diameter' : 'Basic'
      });
    }

    return sawTargetSection;
  }

  // -------------------------------------------------------------------------
  // Exports
  // -------------------------------------------------------------------------

  var api = {
    OUTPUT_COLUMNS: OUTPUT_COLUMNS,
    extractFromSheets: extractFromSheets,
    // exposed for testing / reuse
    _internals: {
      formatDate: formatDate,
      detectColumns: detectColumns,
      isHeaderRow: isHeaderRow,
      detectStrongItemMarker: detectStrongItemMarker,
      detectWeakItemMarker: detectWeakItemMarker,
      detectDiameter: detectDiameter,
      isMostlyDiameter: isMostlyDiameter,
      isTotalRow: isTotalRow,
      normalize: normalize
    }
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.BOQExtractor = api;
  }
})(typeof self !== 'undefined' ? self : (typeof window !== 'undefined' ? window : this));
