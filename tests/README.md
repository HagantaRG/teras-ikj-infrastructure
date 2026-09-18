# Local tests

Run the local mock checks from the repository root:

```powershell
node tests/manifest-identity.cjs
```

These checks cover manifest validation, ID-based column and question matching,
renames, deactivation, dashboard IDs, trigger routing, response writes, and the
explicit reset. They load the deployed source directly and require Node.js.
They do not connect to Google or modify live data.

Tests that run locally will need mocks for Apps Script services such as
`SpreadsheetApp`, `FormApp`, and `Logger`. Test pure data transformations
independently of these services where possible.

This directory is outside clasp's `src` root and is not deployed to Apps Script.
