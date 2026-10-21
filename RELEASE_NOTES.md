# v1.0.0 — Office Authoring Suite

Editable Word, PowerPoint and Excel work plus preservation-aware revision, cross-format synchronization and final-file verification.

This is the group's first stable release. For editable Word, PowerPoint, and Excel work, plus complex document revisions that must preserve comments, tracked changes, package relationships, and cross-format consistency.

## Included in this release

- `word-authoring`
- `powerpoint-authoring`
- `excel-authoring`
- `research-document-delivery`

Install members by task; other groups are optional companions. File helpers require a compatible Python 3.10+ environment. Install pinned libraries, native applications, and rendering tools separately. PowerPoint uses Word's shared helpers, so install both. The two support kits are resources, not additional Skills.

## Installation

```sh
npx skills@1.7.4 add https://github.com/Ynzuses/office-authoring-suite/tree/v1.0.0 --skill "*" --agent codex
```

Project scope is the default. Add `--global` for user-wide use or select one member with `--skill <skill-name>`. Check Node/npm/Git and existing installations first. See the [installation guide](INSTALL.md) for complete setup, dependencies, and removal.

## Migration

Read [MIGRATION.md](MIGRATION.md) before upgrading. Existing tags and versions retain their identities. Upgrading does not automatically remove old installations, install companion groups, or migrate local settings. `1.0.0` marks the first stable contract; internal review rounds are not public version numbers.

## Validation and limits

Microsoft Office, a spreadsheet calculation engine, Poppler, and live connectors are not bundled. Report package checks, supplementary rendering, and native Office acceptance separately. See [VALIDATION.md](VALIDATION.md) for applicable structural, script, and entrypoint checks. Static checks do not guarantee future agent invocation, native application acceptance, or an external connection.

## Downloads and checksums

Custom assets: `office-authoring-suite-1.0.0.zip` and `office-authoring-suite-1.0.0.zip.sha256`. The extracted ZIP contains per-file `SHA256SUMS` and `RELEASE_MANIFEST.json`. GitHub's automatic Source code archives come from the tagged tree and have different packaging bytes and generated inventories.

Use `Get-FileHash <zip-path> -Algorithm SHA256` on Windows, `shasum -a 256 <zip-path>` on macOS, or `sha256sum <zip-path>` on Linux, then compare the companion checksum. Signatures, platform listings, and native application acceptance are reported only when actually completed.

Copyright 2026 Ynzuses. [Apache-2.0](LICENSE). See [NOTICE](NOTICE) for retained terms.

## Feedback and contributions

Use GitHub Issues for synthetic reproductions and installation feedback. See [CONTRIBUTING.md](CONTRIBUTING.md) for changes and validation expectations.
