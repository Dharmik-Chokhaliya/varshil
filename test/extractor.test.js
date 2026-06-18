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

// ---------------------------------------------------------------------------
console.log('\nUnit helpers');
// ---------------------------------------------------------------------------

test('formatDate strips time in various forms', function () {
  var f = BOQ._internals.formatDate;
  eq(f('15-05-2026 12:00 PM'), '15-05-2026');
  eq(f('15-05-2026 12:00:30'), '15-05-2026');
  eq(f('2026-05-15T09:30'), '2026-05-15');
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
