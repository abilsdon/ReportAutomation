# Report Automation for Microsoft Word

This repository contains a Microsoft Word task-pane add-in. Its first tool is **Gender Swap**, which previews and replaces gendered words in the entire open document or the current selection.

## Gender Swap rules

- `he` → `her`, `she` → `him`
- `him` → `her`, `her` → `him`
- `his` → `hers`, `hers` → `his`
- `himself` ↔ `herself`
- `man` ↔ `woman`, `men` ↔ `women`

Matching is case-insensitive and whole-word only. Common capitalisation and surrounding Word formatting are preserved. These direct rules follow the requested examples; users should review results because English pronouns have different grammatical roles.

## Install for local development

Prerequisites: Microsoft Word desktop, Node.js 18 or later, and npm.

1. Open PowerShell in this repository.
2. Run `npm install`.
3. Run `npm.cmd start` and keep that terminal open. The Office tooling installs a trusted localhost certificate, serves the add-in over HTTPS, sideloads `manifest.xml`, and opens Word.
4. In Word, open a document and select **Home → Report Automation → Open tools**.
5. Run `npm.cmd stop` when finished.

If Word says it cannot load the add-in, confirm the `npm.cmd start` terminal is still open and browse to `https://localhost:3000/taskpane.html`. The Gender Swap pane should appear there without a certificate warning.

## Test and validate

- `npm test` runs replacement-rule tests.
- `npm run validate` validates the Office manifest.

## Organisational deployment

Host the static files on HTTPS, replace all `https://localhost:3000` values in `manifest.xml` with that host, and upload the manifest through Microsoft 365 Admin Center under **Settings → Integrated apps → Upload custom apps**.

Localhost is development-only. Replace the placeholder support URL before production deployment.
