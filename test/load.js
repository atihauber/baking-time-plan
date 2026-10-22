const fs = require('fs');
const path = require('path');
const vm = require('vm');

// Apps Script has no module system, so evaluate the .gs source in a sandbox
// and hand back the functions for testing. Pass a fake SpreadsheetApp to
// exercise the code that touches sheets.
const source = fs.readFileSync(path.join(__dirname, '..', 'src', 'recipe_timer.js'), 'utf8');
const exported = [
  'findFirstOverrideTime',
  'recalculateTimes',
  'findTaskIndex',
  'insertDataToSummary',
  'handleEdit',
  'updateCalculatedTimes',
  'updateSummary',
];

function load(globals = {}) {
  const context = vm.createContext({ Logger: { log() {} }, Date, ...globals });
  return vm.runInContext(`${source}\n;({ ${exported.join(', ')} })`, context);
}

module.exports = { ...load(), load };
