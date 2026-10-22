const DEBUG = 0;

const TASK_COLUMN = 0;
const TIME_DELTA_COLUMN = 1;
const OVERRIDE_TIME_COLUMN = 2;
const CALCULATED_TIME_COLUMN = 3;
const FIRST_DATA_ROW = 2;
const SUMMARY_HEADER = [["Recept", "Fázis", "Számolt Időpont", "Recept"]];
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

function findFirstOverrideTime(values) {
  for (let i = 0; i < values.length; i++) {
    if (values[i][OVERRIDE_TIME_COLUMN] !== "") {
      return { initialTime: new Date(values[i][OVERRIDE_TIME_COLUMN]), index: i };
    }
  }
  return { initialTime: null, index: -1 };
}

function recalculateTimes(values, initialIndex, initialTime) {
  let calculatedTime = initialTime;
  values[initialIndex][CALCULATED_TIME_COLUMN] = calculatedTime;

  for (let i = initialIndex + 1; i < values.length; i++) {
    if (values[i][OVERRIDE_TIME_COLUMN] !== "") {
      calculatedTime = new Date(values[i][OVERRIDE_TIME_COLUMN]);
    } else if (values[i - 1][CALCULATED_TIME_COLUMN] !== "") {
      calculatedTime = new Date(values[i - 1][CALCULATED_TIME_COLUMN]);
      calculatedTime.setMinutes(calculatedTime.getMinutes() + values[i][TIME_DELTA_COLUMN]);
    } else {
      calculatedTime = null;
    }
    if (calculatedTime) {
      values[i][CALCULATED_TIME_COLUMN] = calculatedTime;
    }
  }
  return values;
}

// Named handleEdit, not onEdit: Sheets runs a function called onEdit as a simple
// trigger on its own, which would fire alongside the installable trigger.
function handleEdit(e) {
  try {
    const editedSheet = e && e.range ? e.range.getSheet() : null;

    if (editedSheet) {
      // The Summary and anything right of the calculated column cannot change a schedule.
      if (editedSheet.getName() === "Summary" || e.range.getColumn() > CALCULATED_TIME_COLUMN + 1) {
        return;
      }
      debugLog(`Edit detected in sheet: ${editedSheet.getName()}, range: ${e.range.getA1Notation()}`);
    }

    updateCalculatedTimes(editedSheet ? [editedSheet] : undefined);
    updateSummary();
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}

function updateSummary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadsheet.getSheetByName("Summary");
  clearSummarySheet(summarySheet);

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
        if (task[CALCULATED_TIME_COLUMN] !== "") {
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

function updateCalculatedTimes(sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
  sheets.forEach(sheet => {
    if (sheet.getName() !== "Summary") {
      const lastRow = sheet.getLastRow();

      if (lastRow < FIRST_DATA_ROW) {
        debugLog(`Skipping empty sheet: ${sheet.getName()}`);
        return;
      }

      const range = sheet.getRange("A2:D" + lastRow);
      const values = range.getValues();

      debugLog(`Processing sheet: ${sheet.getName()}, number of rows: ${values.length}`);
      values.forEach(row => row[CALCULATED_TIME_COLUMN] = "");

      const { initialTime, index } = findFirstOverrideTime(values);

      if (initialTime !== null) {
        debugLog(`Recalculating times from index: ${index}, initial time: ${initialTime}`);
        recalculateTimes(values, index, initialTime);
      }

      // Write only the calculated column so formulas in A-C are left alone
      sheet.getRange(FIRST_DATA_ROW, CALCULATED_TIME_COLUMN + 1, values.length, 1)
        .setValues(values.map(row => [row[CALCULATED_TIME_COLUMN]]));
    }
  });
}
