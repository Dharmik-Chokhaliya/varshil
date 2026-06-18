/*
 * Node test harness for the BOQ extraction engine.
 * Run with:  node test/extractor.test.js
 *
 * Uses synthetic "sheets" (array-of-arrays) that mimic real RA Bill layouts so
 * the core logic can be validated without a browser or a real .xlsx file.
 */
var assert = require('assert');
var BOQ = require('../js/extractor.js');

var passed = 0;
var failed = 0;

function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('  ✓ ' + name);
  } catch (e) {
    failed++;
    console.log('  ✗ ' + name);
    console.log('      ' + e.message);
  }
}

function eq(a, b, msg) { assert.deepStrictEqual(a, b, msg); }

// ---------------------------------------------------------------------------
console.log('\nBasic Extract');
// ---------------------------------------------------------------------------

test('extracts a single item table, header merge applies once', function () {
  var sheet = {
    name: 'Sheet1',
    rows: [
      ['RA Bill - Abstract of Measurements'],
      [],
      ['Item No. 28 : Providing and laying PCC 1:2:4'],
      [],
      ['Sr No', 'Date', 'Description', 'No.', 'L', 'B', 'D', 'Qty'],
      [null, '15-05-2026 12:00 PM', 'Excavation', null, null, null, null, null],
      [1, null, 'Earth Work', 2, 10, 5, 2, 200],
      [2, null, 'PCC', 1, 10, 5, 0.15, 7.5],
      [null, null, 'TOTAL', null, null, null, null, 207.5]
    ]
  };

  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 2, 'should extract 2 data rows');

  eq(res.rows[0].Description, 'Excavation - Earth Work', 'header merges into first row');
  eq(res.rows[1].Description, 'PCC', 'header cleared after first row');

  eq(res.rows[0].RecordDate, '15-05-2026', 'time stripped from date');
  eq(res.rows[1].RecordDate, '15-05-2026', 'date carried forward');

  eq(res.rows[0].No, 2, 'No preserved exactly');
  eq(res.rows[1].No, 1, 'No preserved exactly');

  eq(res.rows[0].Length, 10);
  eq(res.rows[0].Breadth, 5);
  eq(res.rows[0].Depth, 2);

  eq(res.rows[0].Remarks, '', 'remarks blank for basic');
  eq(res.rows[0].ItemType, 'Basic', 'item type Basic');
});

test('ignores TOTAL / Grand Total / quantity-summary / empty rows', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D', 'Qty'],
      ['01-01-2026', 'Footing', 4, 2, 2, 0.5, 8],
      [],
      [null, null, null, null, null, null, 8],          // qty-only summary
      [null, 'Sub Total', null, null, null, null, 8],     // sub total
      [null, 'Grand Total', null, null, null, null, 8]    // grand total
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 1, 'only the real data row survives');
  eq(res.rows[0].Description, 'Footing');
});

test('preserves original No sequence (no renumbering)', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 5'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['a', 2, 1, 1, 1],
      ['b', 1, 1, 1, 1],
      ['c', 1, 1, 1, 1],
      ['d', 1, 1, 1, 1],
      ['e', 2, 1, 1, 1],
      ['f', 2, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '5' });
  eq(res.rows.map(function (r) { return r.No; }), [2, 1, 1, 1, 2, 2]);
});

test('finds the same item across multiple sheets and multiple tables', function () {
  var s1 = {
    name: 'A',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['x', 1, 1, 1, 1],
      ['Item No. 30'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['ignore me', 1, 1, 1, 1]
    ]
  };
  var s2 = {
    name: 'B',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['y', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([s1, s2], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.Description; }), ['x', 'y']);
});

test('supports an item-number column layout', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item', 'Description', 'No.', 'L', 'B', 'D'],
      [28, 'keep1', 1, 1, 1, 1],
      [31, 'drop', 1, 1, 1, 1],
      [28, 'keep2', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.Description; }), ['keep1', 'keep2']);
});

test('weak item marker ("28  Providing ...") without the word Item', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['28', 'Providing and laying PCC'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D'],
      ['10-10-2026', 'Bed concrete', 3, 4, 4, 0.1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 1);
  eq(res.rows[0].Description, 'Bed concrete');
  eq(res.rows[0].RecordDate, '10-10-2026');
});

// ---------------------------------------------------------------------------
console.log('\nHeader Merge Rule');
// ---------------------------------------------------------------------------

test('header applies to FIRST data row only (Example 1)', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['Excavation'],            // header
      ['Earth Work', 1, 1, 1, 1],
      ['PCC Work', 1, 1, 1, 1],
      ['Brick Work', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.Description; }),
     ['Excavation - Earth Work', 'PCC Work', 'Brick Work']);
});

test('header survives blank rows between header and first data row', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['Excavation'],            // header
      [],                        // blank row in between
      [],
      ['Earth Work', 1, 1, 1, 1],
      ['PCC Work', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.Description; }),
     ['Excavation - Earth Work', 'PCC Work']);
});

test('latest header REPLACES previous pending header (Example 3)', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['Foundation'],            // header 1
      ['Excavation'],            // header 2 (replaces "Foundation")
      ['Earth Work', 1, 1, 1, 1],
      ['PCC', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.Description; }),
     ['Excavation - Earth Work', 'PCC']);
});

test('total / summary rows do not consume the pending header', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Description', 'No.', 'L', 'B', 'D', 'Qty'],
      ['Excavation', null, null, null, null, null], // header
      [null, null, null, null, null, 99],            // stray qty-summary
      ['TOTAL', null, null, null, null, 99],         // stray total
      ['Earth Work', 1, 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 1);
  eq(res.rows[0].Description, 'Excavation - Earth Work');
});

// ---------------------------------------------------------------------------
console.log('\nSteel Extract');
// ---------------------------------------------------------------------------

test('extracts steel rows with diameter -> Remarks and No = 1', function () {
  var sheet = {
    name: 'Steel',
    rows: [
      ['Item No. 31 : Steel reinforcement Fe500'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D', 'Qty'],
      [null, '8 mm', null, null, null, null, null],            // diameter header
      ['15-05-2026 12:00 PM', 'Column Reinforcement', null, null, null, null, null], // section header
      [null, 'Main Bar', 4, 3.5, null, null, 14],
      [null, 'Ring Bar', 20, 1.2, null, null, 24],
      [null, '12 mm', null, null, null, null, null],           // new diameter
      [null, 'Footing Bar', 10, 2.0, null, null, 20],
      [null, 'TOTAL', null, null, null, null, 58]
    ]
  };

  var res = BOQ.extractFromSheets([sheet], { mode: 'steel', itemNumber: '31', itemName: 'Fe500' });
  eq(res.rows.length, 3, 'three steel data rows');

  eq(res.rows[0].Description, 'Column Reinforcement - Main Bar', 'header merge once');
  eq(res.rows[1].Description, 'Ring Bar', 'header cleared');
  eq(res.rows[2].Description, 'Footing Bar');

  eq(res.rows[0].Remarks, '8 mm', 'diameter stored in remarks');
  eq(res.rows[1].Remarks, '8 mm', 'diameter persists across rows');
  eq(res.rows[2].Remarks, '12 mm', 'diameter updates on new header');

  res.rows.forEach(function (r) {
    eq(r.No, 1, 'No is always 1 for steel');
    eq(r.ItemType, 'Diameter', 'item type Diameter');
  });

  eq(res.rows[0].RecordDate, '15-05-2026', 'date carried + time stripped');
});

test('steel respects item-name (grade) filter', function () {
  var sheet = {
    name: 'Steel',
    rows: [
      ['Item No. 31 : Reinforcement Fe415'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['8 mm'],
      ['Bar A', 1, 1, 1, 1],
      ['Item No. 31 : Reinforcement Fe500'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['10 mm'],
      ['Bar B', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'steel', itemNumber: '31', itemName: 'Fe500' });
  eq(res.rows.length, 1, 'only the Fe500 section matches');
  eq(res.rows[0].Description, 'Bar B');
  eq(res.rows[0].Remarks, '10 mm');
});

test('steel extracts by item number alone when no item name is given', function () {
  var sheet = {
    name: 'Steel',
    rows: [
      ['Item No. 31 : Reinforcement Fe415'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['8 mm'],
      ['Bar A', 1, 1, 1, 1],
      ['Item No. 31 : Reinforcement Fe500'],
      ['Description', 'No.', 'L', 'B', 'D'],
      ['10 mm'],
      ['Bar B', 1, 1, 1, 1]
    ]
  };
  // No itemName supplied -> both Fe415 and Fe500 sections of item 31 are returned.
  var res = BOQ.extractFromSheets([sheet], { mode: 'steel', itemNumber: '31' });
  eq(res.rows.map(function (r) { return r.Description; }), ['Bar A', 'Bar B']);
  eq(res.rows.map(function (r) { return r.Remarks; }), ['8 mm', '10 mm']);
  res.rows.forEach(function (r) { eq(r.No, 1); eq(r.ItemType, 'Diameter'); });
});

// ---------------------------------------------------------------------------
console.log('\nRecordDate handling');
// ---------------------------------------------------------------------------

test('detects Record Date / Dt / Dated column headers', function () {
  var variants = ['Date', 'DATE', 'Record Date', 'RecordDate', 'Dt', 'Dated'];
  variants.forEach(function (label) {
    var sheet = {
      name: 'S',
      rows: [
        ['Item No. 28'],
        [label, 'Description', 'No.', 'L', 'B', 'D'],
        ['15-05-2026', 'Earth Work', 1, 1, 1, 1]
      ]
    };
    var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
    eq(res.rows.length, 1, 'row extracted for header "' + label + '"');
    eq(res.rows[0].RecordDate, '15-05-2026', 'date read from column "' + label + '"');
  });
});

test('carries a single heading date across all subsequent rows', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D'],
      ['15-05-2026', 'Earth Work', 1, 1, 1, 1],
      [null, 'PCC Work', 1, 1, 1, 1],
      [null, 'Brick Work', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.RecordDate; }),
     ['15-05-2026', '15-05-2026', '15-05-2026']);
});

test('a new date replaces the carried-forward date', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D'],
      ['15-05-2026', 'A', 1, 1, 1, 1],
      [null, 'B', 1, 1, 1, 1],
      ['20-05-2026', 'C', 1, 1, 1, 1],
      [null, 'D', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.map(function (r) { return r.RecordDate; }),
     ['15-05-2026', '15-05-2026', '20-05-2026', '20-05-2026']);
});

test('converts Excel serial dates (number and numeric string)', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Date', 'Description', 'No.', 'L', 'B', 'D'],
      [45797, 'A', 1, 1, 1, 1],     // serial number -> 20-05-2025
      ['45797', 'B', 1, 1, 1, 1]    // serial as text
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows[0].RecordDate, '20-05-2025', 'serial number converted');
  eq(res.rows[1].RecordDate, '20-05-2025', 'serial string converted');
});

test('RecordDate from standalone row before table header', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      ['Record Date:', '15-05-2026'],   // pre-header label+date row
      ['Description', 'No.', 'L', 'B', 'D'],
      ['Earth Work', 1, 1, 1, 1],
      ['PCC Work',   1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 2);
  eq(res.rows[0].RecordDate, '15-05-2026', 'date from pre-header row');
  eq(res.rows[1].RecordDate, '15-05-2026', 'date carried forward');
});

test('RecordDate from Date object (SheetJS cellDates:true) before table header', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      [new Date(2026, 4, 15), 'Dated'],   // Date object in first cell
      ['Description', 'No.', 'L', 'B', 'D'],
      ['Earth Work', 1, 1, 1, 1]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 1);
  eq(res.rows[0].RecordDate, '15-05-2026', 'Date object in pre-header row');
});

test('section header in merged cell (col 0) sets pendingHeader for Basic Extract', function () {
  var sheet = {
    name: 'S',
    rows: [
      ['Item No. 28'],
      // column layout: Sr(0), Date(1), Description(2), No(3), L(4), B(5), D(6)
      ['Sr', 'Date', 'Description', 'No.', 'L', 'B', 'D'],
      // Section header appears as merged cell -> text only in col 0; descCol (2) is null
      ['RCC WORK', null, null, null, null, null, null],
      [1, '15-05-2026', 'Column Footing', 4, 0.9, 0.9, 0.5],
      [2, null,         'Pedestal',        4, 0.45, 0.45, 0.6]
    ]
  };
  var res = BOQ.extractFromSheets([sheet], { mode: 'basic', itemNumber: '28' });
  eq(res.rows.length, 2);
  eq(res.rows[0].Description, 'RCC WORK - Column Footing', 'merged-cell header applied once');
  eq(res.rows[1].Description, 'Pedestal',                  'header cleared after first row');
  eq(res.rows[0].RecordDate, '15-05-2026');
  eq(res.rows[1].RecordDate, '15-05-2026', 'date carried forward');
});

// ---------------------------------------------------------------------------
console.log('\nUnit helpers');
// ---------------------------------------------------------------------------

test('formatDate strips time and normalises to DD-MM-YYYY', function () {
  var f = BOQ._internals.formatDate;
  eq(f('15-05-2026 12:00 PM'), '15-05-2026');
  eq(f('15-05-2026 00:00:00'), '15-05-2026');
  eq(f('2026-05-15T00:00:00'), '15-05-2026');       // ISO -> DD-MM-YYYY
  eq(f('2026-05-15T09:30'), '15-05-2026');
  eq(f('15/05/2026'), '15-05-2026');                 // slashes -> dashes
  eq(f('15-May-2026'), '15-05-2026');                // month name
  eq(f(new Date(2026, 4, 15)), '15-05-2026');
  eq(f(''), '');
});

test('detectDiameter parses common formats', function () {
  var d = BOQ._internals.detectDiameter;
  eq(d('8 mm'), '8 mm');
  eq(d('10mm'), '10 mm');
  eq(d('Dia 12'), '12 mm');
  eq(d('Ø16'), '16 mm');
  eq(d('no diameter here'), '');
});

test('isHeaderRow recognizes BOQ headers and rejects data', function () {
  var h = BOQ._internals.isHeaderRow;
  eq(h(['Date', 'Description', 'No.', 'L', 'B', 'D', 'Qty']), true);
  eq(h(['Sr', 'Particulars', 'Nos', 'Length', 'Breadth', 'Depth']), true);
  eq(h(['10-10-2026', 'Earth work', 2, 10, 5, 2]), false);
});

// ---------------------------------------------------------------------------
console.log('\n' + (failed === 0 ? '✅ ' : '❌ ') + passed + ' passed, ' + failed + ' failed\n');
process.exit(failed === 0 ? 0 : 1);
