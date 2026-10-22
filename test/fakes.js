// Minimal fakes of the Sheets API pieces recipe_timer.js uses.
// Like the real API, rows are 1-based and a header-only sheet has lastRow 1.

function fakeSheet(name, rows = []) {
  const sheet = {
    rows: rows.map(r => r.slice()),
    getName: () => name,
    getLastRow: () => sheet.rows.length,
    getRange(a1) {
      const m = /^([A-D])(\d+):([A-D])(\d+)$/.exec(a1);
      if (!m || Number(m[4]) < Number(m[2])) {
        throw new Error(`Invalid range: ${a1}`);
      }
      const first = Number(m[2]);
      const last = Number(m[4]);
      return {
        getValues: () => sheet.rows.slice(first - 1, last).map(r => r.slice()),
        setValues: values => values.forEach((v, i) => { sheet.rows[first - 1 + i] = v.slice(); }),
      };
    },
  };
  return sheet;
}

function fakeSummary() {
  const cell = {
    setValues() {}, setFontWeight() {}, setBorder() {}, setBackground() {}, setHorizontalAlignment() {},
  };
  return { getName: () => 'Summary', clear() {}, getRange: () => cell };
}

function fakeSpreadsheetApp(sheets) {
  const spreadsheet = {
    getSheets: () => sheets,
    getSheetByName: name => sheets.find(s => s.getName() === name) || null,
  };
  return { getActiveSpreadsheet: () => spreadsheet };
}

module.exports = { fakeSheet, fakeSummary, fakeSpreadsheetApp };
