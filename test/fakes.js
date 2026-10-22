// Minimal fakes of the Sheets API pieces recipe_timer.js uses.
// Like the real API, rows are 1-based and a header-only sheet has lastRow 1.

function fakeSheet(name, rows = []) {
  const sheet = {
    rows: rows.map(r => r.slice()),
    writes: [],
    getName: () => name,
    getLastRow: () => sheet.rows.length,
    getRange(...args) {
      let first, col, count, width;
      if (typeof args[0] === 'string') {
        const m = /^([A-D])(\d+):([A-D])(\d+)$/.exec(args[0]);
        if (!m || Number(m[4]) < Number(m[2])) {
          throw new Error(`Invalid range: ${args[0]}`);
        }
        first = Number(m[2]);
        col = m[1].charCodeAt(0) - 64;
        count = Number(m[4]) - first + 1;
        width = m[3].charCodeAt(0) - m[1].charCodeAt(0) + 1;
      } else {
        [first, col, count, width] = args;
      }
      return {
        getValues: () => sheet.rows.slice(first - 1, first - 1 + count).map(r => r.slice(col - 1, col - 1 + width)),
        setValues(values) {
          sheet.writes.push({ row: first, col, rows: values.length, cols: width });
          values.forEach((v, i) => {
            const row = sheet.rows[first - 1 + i] || [];
            v.forEach((cell, j) => { row[col - 1 + j] = cell; });
            sheet.rows[first - 1 + i] = row;
          });
        },
      };
    },
  };
  return sheet;
}

function fakeSummary() {
  const calls = { setBackground: 0, setBackgrounds: 0 };
  const cell = {
    setValues() {}, setFontWeight() {}, setBorder() {}, setHorizontalAlignment() {},
    setBackground() { calls.setBackground++; },
    setBackgrounds() { calls.setBackgrounds++; },
  };
  return { calls, getName: () => 'Summary', clear() {}, getRange: () => cell };
}

function fakeSpreadsheetApp(sheets) {
  const spreadsheet = {
    getSheets: () => sheets,
    getSheetByName: name => sheets.find(s => s.getName() === name) || null,
  };
  return { getActiveSpreadsheet: () => spreadsheet };
}

// A minimal onEdit event for an edit in the given sheet and column.
function fakeEdit(sheet, column) {
  return { range: { getSheet: () => sheet, getColumn: () => column, getA1Notation: () => 'X' } };
}

module.exports = { fakeSheet, fakeSummary, fakeSpreadsheetApp, fakeEdit };
