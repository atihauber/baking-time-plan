// Minimal fakes of the Sheets API pieces recipe_timer.js uses.
// Like the real API, rows are 1-based and a header-only sheet has lastRow 1.

function fakeSheet(name, rows = []) {
  const sheet = {
    rows: rows.map(r => r.slice()),
    writes: [],
    reads: 0,
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
        getValues() {
          sheet.reads++;
          return sheet.rows.slice(first - 1, first - 1 + count).map(r => r.slice(col - 1, col - 1 + width));
        },
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

// A Summary sheet that keeps its data rows and their background colors.
function fakeSummary(rows = [], colors = []) {
  const calls = { setBackground: 0, setBackgrounds: 0 };
  const summary = {
    calls,
    rows: rows.map(r => r.slice()),
    colors: colors.slice(),
    getName: () => 'Summary',
    getLastRow: () => (summary.rows.length ? summary.rows.length + 1 : 0),
    clear() { summary.rows = []; summary.colors = []; },
    getRange(row, col, count) {
      const data = typeof row === 'number' && row >= 2;
      return {
        getValues: () => summary.rows.slice(row - 2, row - 2 + count).map(r => r.slice()),
        getBackgrounds: () => summary.colors.slice(row - 2, row - 2 + count).map(c => Array(4).fill(c)),
        setValues(values) {
          if (data && col === 1) summary.rows.splice(row - 2, values.length, ...values.map(v => v.slice()));
        },
        setBackgrounds(values) {
          calls.setBackgrounds++;
          summary.colors.splice(row - 2, values.length, ...values.map(v => v[0]));
        },
        setBackground() { calls.setBackground++; },
        setFontWeight() {}, setBorder() {}, setHorizontalAlignment() {},
      };
    },
  };
  return summary;
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

// A minimal onChange event; the active sheet is where the user made the change.
function fakeChange(changeType, activeSheet) {
  return { changeType, source: { getActiveSheet: () => activeSheet } };
}

// Records the triggers setUpTrigger creates.
function fakeScriptApp() {
  const created = [];
  const builder = name => ({
    forSpreadsheet() {
      return {
        onEdit: () => ({ create: () => created.push([name, 'edit']) }),
        onChange: () => ({ create: () => created.push([name, 'change']) }),
      };
    },
  });
  return { created, getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: builder };
}

module.exports = { fakeSheet, fakeSummary, fakeSpreadsheetApp, fakeEdit, fakeChange, fakeScriptApp };
