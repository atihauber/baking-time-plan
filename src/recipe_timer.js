const DEBUG = 0;

const TASK_COLUMN = 0;
const TIME_DELTA_COLUMN = 1;
const OVERRIDE_TIME_COLUMN = 2;
const CALCULATED_TIME_COLUMN = 3;
const FIRST_DATA_ROW = 2;
const SUMMARY_HEADER = [["Recept", "Fázis", "Számolt Időpont", "Recept"]];
// Written to the calculated column instead of a time when a row cannot be used
const INVALID_TIME_MARK = "⚠ invalid time";
const INVALID_DELAY_MARK = "⚠ invalid delay";
const RECIPE_COLORS = {
  "Recipe1": "#FFC8C8", // Light Red
  "Recipe2": "#C8FFC8", // Light Green
  "Recipe3": "#C8C8FF", // Light Blue
  "Recipe4": "#FFFFC8", // Light Yellow
  "Recipe5": "#FFC8FF"  // Light Pink
};

function setUpTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

function clearSummarySheet(summarySheet) {
  summarySheet.clear();
  const headerRange = summarySheet.getRange("A1:D1");
  headerRange.setValues(SUMMARY_HEADER);
  headerRange.setFontWeight('bold');
  headerRange.setBorder(true, true, true, true, true, true);
}

function getRecipeColors() {
  return Object.values(RECIPE_COLORS);
}

function debugLog(message) {
  if (DEBUG) {
    Logger.log(message);
  }
}

function findTaskIndex(data, task) {
  for (let i = 0; i < data.length; i++) {
    if (
      data[i][0] === task[0] &&
      data[i][1] === task[1] &&
      new Date(data[i][2]).getTime() === new Date(task[2]).getTime() &&
      data[i][3] === task[3]
    ) {
      return i;
    }
  }
  return -1;
}

function insertDataToSummary(summaryTasks, recipeName, task) {
  const row = [recipeName, task[TASK_COLUMN], task[CALCULATED_TIME_COLUMN], recipeName];
  const taskIndex = findTaskIndex(summaryTasks, row);
  if (taskIndex === -1) {
    summaryTasks.push(row);
  }
}

// Sheets hands durations such as 0:45 to scripts as dates on 1899-12-30
const SHEETS_EPOCH = new Date(1899, 11, 30);

// Returns the delay in minutes, or null if the cell is not a usable delay
function parseDelayMinutes(value) {
  if (value === "") {
    return 0;
  }
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : (value.getTime() - SHEETS_EPOCH.getTime()) / 60000;
  }
  const minutes = typeof value === "number" ? value : Number(String(value).trim());
  return Number.isFinite(minutes) ? minutes : null;
}

function isValidTime(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

// Fills the calculated column. Rows without a phase are skipped, an override
// restarts the chain, and every other row adds its delay to the last good time.
// Rows that cannot be used are marked and the chain continues past them.
function calculateTimes(values) {
  let previousTime = null;

  values.forEach(row => {
    row[CALCULATED_TIME_COLUMN] = "";
    if (row[TASK_COLUMN] === "") {
      return;
    }

    const override = row[OVERRIDE_TIME_COLUMN];
    if (override !== "") {
      if (!isValidTime(override)) {
        row[CALCULATED_TIME_COLUMN] = INVALID_TIME_MARK;
        return;
      }
      previousTime = new Date(override);
    } else if (previousTime) {
      const delay = parseDelayMinutes(row[TIME_DELTA_COLUMN]);
      if (delay === null) {
        row[CALCULATED_TIME_COLUMN] = INVALID_DELAY_MARK;
        return;
      }
      previousTime = new Date(previousTime.getTime() + delay * 60000);
    } else {
      return;
    }

    row[CALCULATED_TIME_COLUMN] = previousTime;
  });

  return values;
}

// Named handleEdit, not onEdit: Sheets runs a function called onEdit as a simple
// trigger on its own, which would fire alongside the installable trigger.
// An edit touches only the edited sheet and the Summary; without an event
// (run from the editor) every recipe and the whole Summary are rebuilt.
function handleEdit(e) {
  try {
    const editedSheet = e && e.range ? e.range.getSheet() : null;

    if (!editedSheet) {
      rebuildAll();
      return;
    }

    // The Summary and anything right of the calculated column cannot change a schedule.
    if (editedSheet.getName() === "Summary" || e.range.getColumn() > CALCULATED_TIME_COLUMN + 1) {
      return;
    }
    debugLog(`Edit detected in sheet: ${editedSheet.getName()}, range: ${e.range.getA1Notation()}`);

    refreshRecipe(editedSheet);
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}

function rebuildAll() {
  updateCalculatedTimes();
  updateSummary();
}

function refreshRecipe(sheet) {
  const values = recalculateSheet(sheet);
  updateSummaryForRecipe(sheet.getName(), values);
}

function updateSummary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadsheet.getSheetByName("Summary");

  const sheets = spreadsheet.getSheets();
  const colorList = getRecipeColors();
  let colorIndex = 0;
  const summaryTasks = [];
  const recipeColors = {};

  sheets.forEach(sheet => {
    if (sheet.getName() !== "Summary") {
      const recipeName = sheet.getName();
      const lastRow = sheet.getLastRow();
      if (lastRow < FIRST_DATA_ROW) {
        return;
      }
      const range = sheet.getRange("A2:D" + lastRow);
      const tasks = range.getValues();
      const color = colorList[colorIndex % colorList.length];
      recipeColors[recipeName] = color;
      let hasCalculatedTimes = false;

      tasks.forEach(task => {
        if (isValidTime(task[CALCULATED_TIME_COLUMN])) {
          debugLog(`Adding task to summary: ${task}`);
          insertDataToSummary(summaryTasks, recipeName, task);
          hasCalculatedTimes = true;
        }
      });

      if (hasCalculatedTimes) {
        colorIndex += 1;
      }
    }
  });

  writeSummary(summarySheet, summaryTasks, recipeColors);
}

// Replaces one recipe's rows in the Summary without reading the other recipe
// sheets. Other recipes keep their rows and colors as they are in the Summary.
function updateSummaryForRecipe(recipeName, values) {
  const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  const lastRow = summarySheet.getLastRow();
  let rows = [];
  let backgrounds = [];

  if (lastRow >= FIRST_DATA_ROW) {
    const range = summarySheet.getRange(FIRST_DATA_ROW, 1, lastRow - 1, 4);
    rows = range.getValues();
    backgrounds = range.getBackgrounds();
  }

  const recipeColors = {};
  rows.forEach((row, i) => {
    recipeColors[row[0]] = backgrounds[i][0];
  });
  if (!recipeColors[recipeName]) {
    recipeColors[recipeName] = pickUnusedColor(recipeColors);
  }

  const summaryTasks = rows.filter(row => row[0] !== recipeName);
  values.forEach(task => {
    if (isValidTime(task[CALCULATED_TIME_COLUMN])) {
      insertDataToSummary(summaryTasks, recipeName, task);
    }
  });

  writeSummary(summarySheet, summaryTasks, recipeColors);
}

function pickUnusedColor(recipeColors) {
  const colorList = getRecipeColors();
  const used = Object.values(recipeColors).map(color => String(color).toLowerCase());
  const unused = colorList.find(color => !used.includes(color.toLowerCase()));
  return unused || colorList[Object.keys(recipeColors).length % colorList.length];
}

function writeSummary(summarySheet, summaryTasks, recipeColors) {
  clearSummarySheet(summarySheet);
  summaryTasks.sort((a, b) => new Date(a[2]) - new Date(b[2]));

  // Insert the sorted tasks into the summary sheet and apply coloring
  if (summaryTasks.length > 0) {
    summarySheet.getRange(FIRST_DATA_ROW, 1, summaryTasks.length, 4).setValues(summaryTasks);
    const backgrounds = summaryTasks.map(task => Array(4).fill(recipeColors[task[0]]));
    summarySheet.getRange(FIRST_DATA_ROW, 1, summaryTasks.length, 4).setBackgrounds(backgrounds);

    // Center-align the calculated time column
    summarySheet.getRange(FIRST_DATA_ROW, 3, summaryTasks.length, 1).setHorizontalAlignment('center');
  }
}

function updateCalculatedTimes() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
    if (sheet.getName() !== "Summary") {
      recalculateSheet(sheet);
    }
  });
}

// Recalculates one recipe sheet and returns its data rows (empty if it has none)
function recalculateSheet(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < FIRST_DATA_ROW) {
    debugLog(`Skipping empty sheet: ${sheet.getName()}`);
    return [];
  }

  const range = sheet.getRange("A2:D" + lastRow);
  const values = range.getValues();

  debugLog(`Processing sheet: ${sheet.getName()}, number of rows: ${values.length}`);
  calculateTimes(values);

  // Write only the calculated column so formulas in A-C are left alone
  sheet.getRange(FIRST_DATA_ROW, CALCULATED_TIME_COLUMN + 1, values.length, 1)
    .setValues(values.map(row => [row[CALCULATED_TIME_COLUMN]]));
  return values;
}
