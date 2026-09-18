# teras-ikj-infrastructure
Repository for the JS scripts for automation of daily reporting for the TERAS IKJ cafe.

## Directory structure

```text
.
|-- .clasp.json              # Connected Apps Script project and source directory
|-- .gitignore
|-- AGENTS.md
|-- LICENSE
|-- README.md
|-- src/                    # Files deployed to Apps Script
|   |-- appsscript.json     # Runtime, timezone, and service configuration
|   `-- automations/        # Sheet automation workflows
|       `-- Code.js          # Existing createFormFromSheet entry point
|-- docs/                   # Sheet schema and operational documentation
|   `-- README.md
`-- tests/                  # Local tests and synthetic fixtures
    `-- README.md
```

Keep deployable JavaScript and any Apps Script HTML under `src/`. The clasp
configuration sets `rootDir` to `src` and enables subdirectory discovery, so
documentation and local tests stay outside the deployment directory.

Add each new workflow under `src/automations/` with a descriptive filename.
When needed, add `src/sheets/` for reusable sheet access helpers and
`src/triggers/` for menu, edit, and scheduled trigger entry points. Create these
directories when they have an implementation to hold.

Folders organize source files; they do not create JavaScript module scopes.
Keep Apps Script function names unique across the project.

## Local workflow

Run clasp commands from the repository root. With clasp installed and
authenticated, use `clasp status` to review the deployment file list before
running `clasp push` to update the connected Apps Script project.

The existing automation and manifest were moved without changing their contents.
No build step is required. Run `node tests/manifest-identity.cjs` for local mock
checks. See [inventory setup and operation](docs/README.md) for the manifest
schema, clean-start reset, and trigger setup.
