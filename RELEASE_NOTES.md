# v1.0.1 — Office Authoring Suite

Editable Word, PowerPoint and Excel work plus preservation-aware revision, cross-format synchronization and final-file verification.

This patch extends the installation and invocation guidance from Codex to Claude Code and fixes one page-count check. Member instructions, templates and schemas are unchanged from v1.0.0 apart from the version metadata, and v1.0.0 remains available at its own tag.

## What changed

- The README and the [installation guide](INSTALL.md) give the Claude Code command (`--agent claude-code`) next to the Codex one and explain how each host invokes a Skill.
- The discovery check in the installation guide covers Claude Code.
- Fixed-tag links, the product tag and asset names now refer to v1.0.1.
- **Fix:** the PowerPoint export check reads the PDF page count from `pdfinfo` as bytes. On systems where `pdfinfo` prints metadata such as a time-zone name in a non-UTF-8 code page, the previous text decode failed and the check crashed instead of returning a result. A regression test covers this case.

## Included in this release

- `word-authoring`
- `powerpoint-authoring`
- `excel-authoring`
- `research-document-delivery`

Install members by task; other groups are optional companions. File helpers require a compatible Python 3.10+ environment. Install pinned libraries, native applications, and rendering tools separately. PowerPoint uses Word's shared helpers, so install both. The two support kits are resources, not additional Skills.

## Install or update

```sh
npx skills@1.7.4 add https://github.com/Ynzuses/office-authoring-suite/tree/v1.0.1 --skill "*" --agent codex
npx skills@1.7.4 add https://github.com/Ynzuses/office-authoring-suite/tree/v1.0.1 --skill "*" --agent claude-code
```

Project scope is the default. Add `--global` for user-wide use or select one member with `--skill <skill-name>`. Replace complete member directories; upgrading does not remove old copies, install companion groups or change local settings. See [MIGRATION.md](MIGRATION.md) and the [installation guide](INSTALL.md).

## Validation and limits

Microsoft Office, a spreadsheet calculation engine, Poppler, and live connectors are not bundled. Report package checks, supplementary rendering, and native Office acceptance separately. See [VALIDATION.md](VALIDATION.md) for applicable structural, script, and entrypoint checks. Static checks do not guarantee future agent invocation, native application acceptance, or an external connection.

For this patch the source and an independent candidate copy were re-run with the group's regressions and official Skill validation; see [VALIDATION.md](VALIDATION.md).

## Downloads and checksums

Custom assets: `office-authoring-suite-1.0.1.zip` and `office-authoring-suite-1.0.1.zip.sha256`. The extracted ZIP contains per-file `SHA256SUMS` and `RELEASE_MANIFEST.json`. GitHub's automatic Source code archives come from the tagged tree and have different packaging bytes and generated inventories.

Use `Get-FileHash <zip-path> -Algorithm SHA256` on Windows, `shasum -a 256 <zip-path>` on macOS, or `sha256sum <zip-path>` on Linux, then compare the companion checksum. Signatures, platform listings, and native application acceptance are reported only when actually completed.

Copyright 2026 Ynzuses. [Apache-2.0](LICENSE). See [NOTICE](NOTICE) for retained terms.

## Feedback and contributions

Use GitHub Issues for synthetic reproductions and installation feedback. See [CONTRIBUTING.md](CONTRIBUTING.md) for changes and validation expectations.
