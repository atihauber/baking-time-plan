const {
  findFirstOverrideTime,
  recalculateTimes,
  findTaskIndex,
  insertDataToSummary,
} = require('./load');

const t = (h, m) => new Date(1899, 11, 30, h, m);
const row = (phase, delay, override = '') => [phase, delay, override, ''];

describe('findFirstOverrideTime', () => {
  test('returns the first non-empty override', () => {
    const values = [row('a', 0), row('b', 60, t(15, 15)), row('c', 30, t(18, 0))];
    expect(findFirstOverrideTime(values)).toEqual({ initialTime: t(15, 15), index: 1 });
  });

  test('returns no start when there is no override', () => {
    expect(findFirstOverrideTime([row('a', 0), row('b', 5)]).index).toBe(-1);
  });
});

describe('recalculateTimes', () => {
  test('matches the README example (15:15, 15:45, 16:30, 17:15)', () => {
    const values = [
      row('Autolízis', 0),
      row('Dagasztás', 60, t(15, 15)),
      row('Dagasztás sóval', 30),
      row('Laminálás', 45),
      row('Hajtogatás 1.', 45),
    ];
    recalculateTimes(values, 1, t(15, 15));
    expect(values.map(r => r[3])).toEqual(['', t(15, 15), t(15, 45), t(16, 30), t(17, 15)]);
  });

  test('a later override restarts the chain', () => {
    const values = [row('a', 0, t(10, 0)), row('b', 30), row('c', 30, t(12, 0)), row('d', 15)];
    recalculateTimes(values, 0, t(10, 0));
    expect(values.map(r => r[3])).toEqual([t(10, 0), t(10, 30), t(12, 0), t(12, 15)]);
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
    ['empty', []],
    ['header-only', [header]],
  ])('updateCalculatedTimes skips a %s sheet without touching it', (_, rows) => {
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

  test.each([
    ['empty', []],
    ['header-only', [header]],
  ])('hasCalculatedTimes is false for a %s sheet', (_, rows) => {
    expect(setup().hasCalculatedTimes(fakeSheet('Blank', rows))).toBe(false);
  });

  test('updateSummary tolerates empty and header-only sheets', () => {
    const sheets = [fakeSheet('Empty', []), fakeSheet('Headers', [header])];
    expect(() => setup(...sheets).updateSummary()).not.toThrow();
  });
});
