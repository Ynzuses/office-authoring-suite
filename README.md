# Office Authoring Suite

Editable Word, PowerPoint and Excel work plus preservation-aware revision, cross-format synchronization and final-file verification.

For editable Word, PowerPoint, and Excel work, plus complex document revisions that must preserve comments, tracked changes, package relationships, and cross-format consistency.

## Included Skills

| Skill | Purpose |
| --- | --- |
| [word-authoring](word-authoring/SKILL.md) | Editable DOCX creation with semantic styles, bilingual typography, package checks and application acceptance. |
| [powerpoint-authoring](powerpoint-authoring/SKILL.md) | Editable presentations, relevant images, content-led layouts and native/visual verification. |
| [excel-authoring](excel-authoring/SKILL.md) | Preservation-aware workbook files or targeted edits through an actually advertised live Excel session. |
| [research-document-delivery](research-document-delivery/SKILL.md) | Rich-document revision, comments and tracked changes, OpenXML preservation, synchronization and final-file verification; historical ID retained. |

## Install with skills.sh

After the `v1.0.0` tag is published, run this from the project where you want to use the group:

```sh
npx skills@1.7.4 add https://github.com/Ynzuses/office-authoring-suite/tree/v1.0.0 --skill "*" --agent codex
```

This selects every member for Codex and defaults to the current project. Add `--list` to inspect available members or `--global` for user-wide installation. Use `--skill <skill-name>` to select one member, or choose another supported target such as `--agent claude-code`. Keep the first installation interactive so you can review its destination. Use `--copy` if symbolic links are unavailable.

See **[INSTALL.md](INSTALL.md)** for prerequisites, pinned versions, individual installation, discovery paths, dependencies, updates, and removal. Verify discovery and an actual invocation in the target agent after installing the files.

## Example requests

> Use $word-authoring to create an editable DOCX from the approved text, preserve semantic styles, and verify the final file.

> Use $excel-authoring to inspect this workbook and the available calculation engine, make the requested edits, and check formulas and preserved content.

These are requests to an agent. Codex supports explicit `$skill-name` invocation; other agents use their own invocation interface.

## Requirements and scope

File helpers require a compatible Python 3.10+ environment. Install pinned libraries, native applications, and rendering tools separately. PowerPoint uses Word's shared helpers, so install both. The two support kits are resources, not additional Skills.

Microsoft Office, a spreadsheet calculation engine, Poppler, and live connectors are not bundled. Report package checks, supplementary rendering, and native Office acceptance separately.

## Versions, support, and license

`1.0.0` is this group's first stable contract. Read [MIGRATION.md](MIGRATION.md) before replacing an older installation and [RELEASE_NOTES.md](RELEASE_NOTES.md) for this release. Compatibility and verification limits are documented in [STABILITY.md](STABILITY.md) and [VALIDATION.md](VALIDATION.md).

Optional companion groups: `research-skills-suite`, `workspace-skills-suite`, `voice-preserving-writing`. They are independent distributions and are not installed with this group. When reporting a problem, include agent/CLI versions, reproducible steps, errors, and a synthetic example. Remove credentials, personal data, and unpublished material.

Copyright 2026 Ynzuses. [Apache-2.0](LICENSE). See [NOTICE](NOTICE) and [PROVENANCE.md](PROVENANCE.md) for attribution and retained component terms.

For proposed changes and reproducible reports, see [CONTRIBUTING.md](CONTRIBUTING.md).
