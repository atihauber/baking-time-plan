const { calculateTimes, parseDelayMinutes, findTaskIndex, insertDataToSummary } = require('./load');

const t = (h, m) => new Date(1899, 11, 30, h, m);
const row = (phase, delay, override = '') => [phase, delay, override, ''];
const times = rows => calculateTimes(rows).map(r => r[3]);

describe('calculateTimes', () => {
  test('matches the README example (15:15, 15:45, 16:30, 17:15)', () => {
    expect(times([
      row('Autolízis', 0),
      row('Dagasztás', 60, t(15, 15)),
      row('Dagasztás sóval', 30),
      row('Laminálás', 45),
      row('Hajtogatás 1.', 45),
    ])).toEqual(['', t(15, 15), t(15, 45), t(16, 30), t(17, 15)]);
  });

  test('a later override restarts the chain', () => {
    expect(times([row('a', 0, t(10, 0)), row('b', 30), row('c', 30, t(12, 0)), row('d', 15)]))
      .toEqual([t(10, 0), t(10, 30), t(12, 0), t(12, 15)]);
  });

  test('stale calculated times are cleared when there is no override', () => {
    expect(times([['a', 0, '', t(8, 0)], ['b', 5, '', t(8, 5)]])).toEqual(['', '']);
  });

  test('rows without a phase get no time and the chain continues past them', () => {
    expect(times([row('Mix', 0, t(8, 0)), row('', ''), row('Bake', 30), ['', '', '', ''], row('', '')]))
      .toEqual([t(8, 0), '', t(8, 30), '', '']);
  });

  test('an invalid override is marked and skipped', () => {
    expect(times([row('Mix', 0, t(8, 0)), row('Fold', 30, '9.30'), row('Bake', 30)]))
      .toEqual([t(8, 0), '⚠ invalid time', t(8, 30)]);
  });

  test('an invalid override before any start is marked', () => {
    expect(times([row('Mix', 0, 930), row('Bake', 30)])).toEqual(['⚠ invalid time', '']);
  });

  test('an invalid delay is marked and skipped', () => {
    expect(times([row('Mix', 0, t(8, 0)), row('Fold', '1 pack'), row('Bake', 30)]))
      .toEqual([t(8, 0), '⚠ invalid delay', t(8, 30)]);
  });

  test('a delay stored as text is added, not concatenated', () => {
    expect(times([row('Mix', 0, t(8, 15)), row('Fold', 30), row('Shape', '45')]))
      .toEqual([t(8, 15), t(8, 45), t(9, 30)]);
  });
});

describe('parseDelayMinutes', () => {
  test.each([
    [45, 45],
    ['45', 45],
    [' 45 ', 45],
    ['', 0],
    [1.5, 1.5],
    [t(0, 45), 45],
    [new Date(1899, 11, 31, 1, 0), 25 * 60],
    ['1 pack', null],
    [new Date('x'), null],
  ])('%p -> %p', (value, expected) => {
    expect(parseDelayMinutes(value)).toBe(expected);
  });
});

describe('summary de-duplication', () => {
  test('identical rows are inserted once', () => {
    const tasks = [];
    insertDataToSummary(tasks, 'Kenyér', ['Dagasztás', 60, '', t(15, 15)]);
    insertDataToSummary(tasks, 'Kenyér', ['Dagasztás', 60, '', new Date(t(15, 15))]);
    expect(tasks).toHaveLength(1);
  });

  test('same phase in another recipe is kept', () => {
    const tasks = [];
    insertDataToSummary(tasks, 'Kenyér', ['Dagasztás', 60, '', t(15, 15)]);
    insertDataToSummary(tasks, 'Pizza', ['Dagasztás', 60, '', t(15, 15)]);
    expect(tasks).toHaveLength(2);
  });

  test('findTaskIndex returns -1 when absent', () => {
    expect(findTaskIndex([], ['x', 'y', t(1, 0), 'x'])).toBe(-1);
  });
});

describe('sheets without data rows', () => {
  const { load } = require('./load');
  const { fakeSheet, fakeSummary, fakeSpreadsheetApp } = require('./fakes');
  const header = ['Phase', 'Delay (min)', 'Time override', 'Calculated time'];

  const setup = (...sheets) => load({ SpreadsheetApp: fakeSpreadsheetApp([...sheets, fakeSummary()]) });

  test.each([
    ['an empty', []],
    ['a header-only', [header]],
  ])('updateCalculatedTimes skips %s sheet without touching it', (_, rows) => {
    const blank = fakeSheet('Blank', rows);
    setup(blank).updateCalculatedTimes();
    expect(blank.rows).toEqual(rows);
  });

  test('a blank sheet does not stop other recipes from recalculating', () => {
    const blank = fakeSheet('Blank', []);
    const recipe = fakeSheet('Kenyér', [header, ['Mix', 0, t(8, 0), ''], ['Bake', 30, '', '']]);
    setup(blank, recipe).updateCalculatedTimes();
    expect(recipe.rows[2][3]).toEqual(t(8, 30));
  });

  test('updateSummary tolerates empty and header-only sheets', () => {
    const sheets = [fakeSheet('Empty', []), fakeSheet('Headers', [header])];
    expect(() => setup(...sheets).updateSummary()).not.toThrow();
  });

  test('updateSummary leaves out rows marked as invalid', () => {
    const summary = fakeSummary();
    const recipe = fakeSheet('Kenyér', [header, ['Mix', 0, t(8, 0), '⚠ invalid time'], ['Bake', 30, '', '']]);
    load({ SpreadsheetApp: fakeSpreadsheetApp([recipe, summary]) }).updateSummary();
    expect(summary.calls.setBackgrounds).toBe(0);
  });

  test('updateSummary colors the rows with a single batched call', () => {
    const summary = fakeSummary();
    const recipe = fakeSheet('Kenyér', [header, ['Mix', 0, t(8, 0), t(8, 0)], ['Bake', 30, '', t(8, 30)]]);
    load({ SpreadsheetApp: fakeSpreadsheetApp([recipe, summary]) }).updateSummary();
    expect(summary.calls).toEqual({ setBackground: 0, setBackgrounds: 1 });
  });
});

describe('edit handling', () => {
  const { load } = require('./load');
  const { fakeSheet, fakeSummary, fakeSpreadsheetApp, fakeEdit } = require('./fakes');
  const header = ['Phase', 'Delay (min)', 'Time override', 'Calculated time'];
  const recipeRows = () => [header, ['Mix', 0, t(8, 0), ''], ['Bake', 30, '', '']];
  const setup = (...sheets) => load({ SpreadsheetApp: fakeSpreadsheetApp([...sheets, fakeSummary()]) });

  test('updateCalculatedTimes writes only the calculated column', () => {
    const recipe = fakeSheet('Kenyér', recipeRows());
    setup(recipe).updateCalculatedTimes();
    expect(recipe.writes).toEqual([{ row: 2, col: 4, rows: 2, cols: 1 }]);
    expect(recipe.rows[2]).toEqual(['Bake', 30, '', t(8, 30)]);
  });

  test('an edit recalculates only the edited sheet', () => {
    const edited = fakeSheet('Kenyér', recipeRows());
    const other = fakeSheet('Pizza', recipeRows());
    setup(edited, other).handleEdit(fakeEdit(edited, 3));
    expect(edited.writes).toHaveLength(1);
    expect(other.writes).toHaveLength(0);
  });

  test('edits right of the calculated column are ignored', () => {
    const recipe = fakeSheet('Kenyér', recipeRows());
    setup(recipe).handleEdit(fakeEdit(recipe, 6));
    expect(recipe.writes).toHaveLength(0);
  });

  test('edits on the Summary sheet are ignored', () => {
    const recipe = fakeSheet('Kenyér', recipeRows());
    const summary = fakeSummary();
    setup(recipe).handleEdit(fakeEdit(summary, 1));
    expect(recipe.writes).toHaveLength(0);
  });

  test('without an event every recipe is recalculated', () => {
    const a = fakeSheet('Kenyér', recipeRows());
    const b = fakeSheet('Pizza', recipeRows());
    setup(a, b).handleEdit();
    expect([a.writes.length, b.writes.length]).toEqual([1, 1]);
  });
});
