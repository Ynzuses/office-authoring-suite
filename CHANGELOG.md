# Changelog

## 1.0.1 - 2026-11-11

Installation and invocation guidance now covers Claude Code (`--agent claude-code`, invocation by name) alongside Codex, and fixed-tag links, the installer pin and asset names refer to v1.0.1. Fixed the PowerPoint export check's PDF page count on systems whose `pdfinfo` prints metadata in a non-UTF-8 code page (for example a Chinese-locale Windows time-zone name): output is now read as bytes and decoded leniently, so the check no longer fails with a decoding error; a regression test covers it. No other instruction, script, schema or interface change.

## 1.0.0 - 2026-10-20

First stable Office Authoring Suite group; 4 declared members. See MIGRATION.md and RELEASE_NOTES.md.

## Pre-stable history

### Release documentation preparation (2026-10-20, before publication)

Added scope and task examples, pinned skills.sh setup and removal instructions, dependency guidance, and release notes. The public version remains 1.0.0; documentation preparation alone does not establish release publication or deployment.

### English public documentation (2026-10-21, before the stable release)

Standardized README, installation, release notes, and changelog prose in English. Existing runtime behavior, Skill IDs, and the first-stable 1.0.0 contract are unchanged.

### Community reporting before stable publication

Added a structured bug-report form and completed the contribution guidance for the stable repository.
