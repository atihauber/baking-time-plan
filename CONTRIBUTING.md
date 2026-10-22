# Contributing

Thanks for considering a contribution to Baking Time Plan.

## Reporting a bug

Open a GitHub Issue and include:

- The steps needed to reproduce the problem.
- What you expected to happen.
- What actually happened.
- The relevant error from Apps Script **Executions**, if available.

Do not include private spreadsheet contents, personal information, or spreadsheet IDs.

## Proposing a change

1. Fork the repository.
2. Create a branch for your change.
3. Keep the change focused and avoid adding dependencies unless they are necessary.
4. Test the change in a copy of a Google spreadsheet.
5. Open a pull request describing the behavior changed and how you tested it.

## Compatibility expectations

- Preserve the Hungarian spreadsheet headers unless the change explicitly updates the documentation too.
- Preserve chronological summary sorting.
- Prevent duplicate summary entries.
- Keep calculated times blank before the first override.
- Check empty or header-only sheets when changing range handling.
- Verify that recipe row colors remain associated with the correct recipe after sorting.
