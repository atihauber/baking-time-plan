const fs = require('fs');
const path = require('path');
const vm = require('vm');

// Apps Script has no module system, so evaluate the .gs source in a sandbox
// and hand back the pure-logic functions for testing.
const source = fs.readFileSync(path.join(__dirname, '..', 'src', 'recipe_timer.js'), 'utf8');
const context = vm.createContext({ Logger: { log() {} }, Date });

module.exports = vm.runInContext(
  source + '\n;({ findFirstOverrideTime, recalculateTimes, findTaskIndex, insertDataToSummary })',
  context
);
