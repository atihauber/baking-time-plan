# Contributing

Thanks for considering a contribution to Baking Time Plan.

## Reporting a bug

Open a GitHub Issue and include:

- The steps needed to reproduce the problem.
- What you expected to happen.
- What actually happened.
- The relevant error from Apps Script **Executions**, if available.

Do not include private spreadsheet contents, personal information, or spreadsheet IDs.

## Development setup

The script in `src/` is synced with a spreadsheet's Apps Script project through [clasp](https://github.com/google/clasp). Install [Node.js](https://nodejs.org/), then:

```
npm install
npx clasp login
npx clasp clone <scriptId> --rootDir src   # once; creates .clasp.json
npm run pull    # Apps Script -> src/
npm run push    # src/ -> Apps Script (overwrites the project's files)
npm test        # run the tests
```

The script ID is under **Project Settings** in the Apps Script editor. `.clasp.json` and `.clasprc.json` are gitignored; keep them out of commits.

The tests in `test/` load `src/recipe_timer.js` in a sandbox. They cover the calculation logic directly and the sheet handling through small fakes of the Sheets API (`test/fakes.js`). Anything that depends on real Sheets behavior still needs a check in a copy of a spreadsheet.

Set `DEBUG = 1` at the top of the script to log each step to **Executions**.

## Proposing a change

1. Fork the repository.
2. Create a branch for your change.
3. Keep the change focused and avoid adding dependencies unless they are necessary.
4. Run `npm test`, add tests for the behavior you change, and try it in a copy of a Google spreadsheet.
5. Open a pull request describing the behavior changed and how you tested it.

## Compatibility expectations

- Preserve the `Summary` headers unless the change also updates the documentation.
- Preserve chronological summary sorting.
- Prevent duplicate summary entries.
- Keep calculated times blank before the first override.
- Check empty or header-only sheets when changing range handling.
- Verify that recipe row colors remain associated with the correct recipe after sorting.
