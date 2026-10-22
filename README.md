# Baking Time Plan

A Google Sheets and Google Apps Script tool for scheduling recipe phases by time.

Enter a time override for a phase, and the script calculates subsequent phase times using each row's delay in minutes. A `Summary` sheet collects calculated phases from the recipe sheets, sorts them chronologically, and colors rows by recipe.

The project is designed for baking workflows, but the same approach can be adapted to other sequences of timed tasks.

## Features

- Define an ordered list of phases for each recipe.
- Enter a manual time override at any phase.
- Calculate following phase times from the previous calculated time and the configured delay.
- Recalculate recipe times when the spreadsheet is edited.
- Collect calculated phases into a chronological summary.
- Color-code recipes in the summary.
- Use Hungarian column labels.

## Example

A recipe sheet with an override time (`8:00`) and the times calculated from it:

![A recipe sheet with an override time and calculated times](docs/images/recipe-sheet.png)

Each recipe has its own sheet; here is a second one (`Pizza`) with a different start time:

![A second recipe sheet with its own override time](docs/images/recipe-sheet-2.png)

The `Summary` sheet merges all recipes in chronological order, with one color per recipe:

![The Summary sheet showing two recipes interleaved by time](docs/images/summary.png)

## Requirements

- A Google account with access to Google Sheets and Google Apps Script.
- A spreadsheet containing one sheet per recipe and a sheet named `Summary`.
- The Apps Script source from [`src/recipe_timer.js`](src/recipe_timer.js).

## Spreadsheet layout

Each recipe sheet uses these columns, starting in row 1:

| Column | Hungarian header | Purpose |
| --- | --- | --- |
| A | `Fázis` | Phase or task name |
| B | `Késleltetés (min)` | Minutes after the previous calculated phase |
| C | `Időpont megadása / felülírás` | Optional manually entered time |
| D | `Számolt időpont` | Calculated time written by the script |

The script expects recipe data to start in row 2. Keep the phases in the order in which they should be processed.

Create a sheet named exactly `Summary`. Its output columns are:

| Column | Header | Purpose |
| --- | --- | --- |
| A | `Recept` | Recipe sheet name |
| B | `Fázis` | Phase name |
| C | `Számolt Időpont` | Calculated time |
| D | `Recept` | Recipe name repeated for color identification |

The script writes the summary; avoid manually editing its generated rows.

See [`examples/spreadsheet-layout.md`](examples/spreadsheet-layout.md) for a compact setup example.

## Installation

1. Create a Google spreadsheet.
2. Create a `Summary` sheet.
3. Create a sheet for each recipe, using the column layout above.
4. Open **Extensions → Apps Script**.
5. Copy the contents of [`src/recipe_timer.js`](src/recipe_timer.js) into the editor, replacing the default `Code.gs` content.
6. Save the project.
7. Select `setUpTrigger` in the Apps Script editor and run it once.
8. Review and grant the requested permissions.
9. Enter a time in column C of a recipe sheet. The calculated times and summary should update automatically.

Run `setUpTrigger` only when you need to create or recreate the edit trigger. The function removes existing project triggers before creating the `onEdit` trigger.

## How calculations work

- The script finds the first non-empty override time in column C.
- That row becomes the starting point for calculation.
- Each following row uses its override time if present; otherwise, its calculated time is the previous calculated time plus the row's delay in minutes.
- Rows before the first override do not receive calculated times.
- The summary includes rows with calculated times and sorts them by time.

## Configuration and limitations

- The summary sheet name is `Summary`.
- Recipe data is expected in columns A–D, starting at row 2.
- The current design has a five-color palette (`RECIPE_COLORS` in the script). Colors are assigned by position, in sheet order, to recipes that have calculated times; the keys (`Recipe1`–`Recipe5`) are labels only and do not need to match sheet names. If more than five recipes have calculated times, colors repeat.
- Every edit, on any sheet, recalculates all recipes and rebuilds the `Summary`.
- `DEBUG` is `0` by default. Set it to `1` to write detailed progress to the Apps Script log.
- The whole `Summary` sheet is cleared and rebuilt on each run, so anything typed there is lost.
- Times are handled as time-of-day values; a schedule that runs past midnight sorts after earlier times in the summary.
- Recipe sheet names are used as recipe names in the summary.
- The script is a Google Apps Script project intended to be copied into a spreadsheet; this repository is not a standalone web application.

## Troubleshooting

- **Nothing updates:** check that the installable edit trigger exists in Apps Script under **Triggers**, and review **Executions** for errors.
- **Authorization is requested:** run `setUpTrigger` from the Apps Script editor and complete Google's authorization flow.
- **No calculated times appear:** ensure at least one row in column C contains a valid time override.
- **The summary cannot be found:** create a sheet named exactly `Summary`.

When reporting a bug, include the steps to reproduce it and the relevant error from Apps Script **Executions**. Do not share private spreadsheet data or spreadsheet IDs.

## Development

The script lives in `src/` and is synced with the spreadsheet's Apps Script project using [clasp](https://github.com/google/clasp). Install [Node.js](https://nodejs.org/) first.

```
npm install
npx clasp login
npx clasp clone <scriptId> --rootDir src   # once; creates the gitignored .clasp.json
npm run pull    # Apps Script -> src/
npm run push    # src/ -> Apps Script (overwrites the project's files)
npm test        # unit tests for the calculation and summary logic
```

Find the script ID under **Project Settings** in the Apps Script editor. Keep `.clasp.json` and `.clasprc.json` out of version control; both are in `.gitignore`.

The tests in `test/` load `src/recipe_timer.js` in a sandbox and cover the pure logic (`findFirstOverrideTime`, `recalculateTimes`, summary de-duplication). Code that calls `SpreadsheetApp` must still be tested in a copy of a spreadsheet.

## Contributing

Bug reports and focused improvements are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

Released under the MIT License. See [LICENSE](LICENSE).
