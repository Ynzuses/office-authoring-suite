# Excel authoring provenance

This package is an independently authored implementation. It is informed by public workflow patterns and by the installed OpenAI Spreadsheets runtime, but it does not bundle third-party skill files, downloaded templates, or private course material.

## Public references reviewed on 2026-09-18

The `skills.sh` API was queried for `excel` and `xlsx` results. The following public entries were read or compared:

| Source | skills.sh installs at review | Repository stars at review | License signal | Decision |
| --- | ---: | ---: | --- | --- |
| `claude-office-skills/skills@excel-automation` | 14,893 | 472 | MIT | Borrowed the distinction between batch and live Excel work, typed ranges, and explicit formatting; replaced its `xlwings`-first route with the local artifact runtime. |
| `claude-office-skills/skills@xlsx-manipulation` | 5,623 | 472 | MIT | Borrowed preservation-aware editing, formulas, validations, and tables; did not copy its openpyxl implementation. |
| `anthropics/skills@xlsx` | 169,743 | 176,852 | Proprietary | Read for formula recalculation, external-link, macro, merged-cell, and source-protection gotchas. No text or proprietary files are bundled. |
| `davila7/claude-code-templates@excel-analysis` | 2,486 | 30,771 | MIT | Borrowed the clear split between loading, analysis, summaries, charts, and validation; replaced pandas/openpyxl authoring with the local runtime contract. |
| `sbroenne/mcp-server-excel@excel-mcp` | 1,613 | 748 | MIT | Borrowed session identity, targeted updates, calculation-mode discipline, table-first structure, and suggested-next-action error recovery; rejected Windows-only assumptions. |
| `daymade/claude-code-skills@excel-automation` | 1,241 | 1,396 | MIT | Borrowed file-type detection, macro/complex-model caution, and safe AppleScript timeouts; kept native checks non-destructive and macOS-specific. |
| `zgldh/xlsx-populate-skill@processing-excel-files` | 379 | 8 | MIT | Borrowed the explicit preservation goal for existing formatting and merged cells; did not use its library because the installed runtime mandates artifact-tool for authoring. |

Install counts and star counts are discovery metadata, not evidence that a workflow is correct for this machine. Windows-only MCP, live-session, or proprietary implementation details are not copied into the skill. The authoritative local constraints remain the workspace `AGENTS.md`, the artifact runtime Spreadsheets skill, and the existing Office authoring kit.

## Local design decisions

- `excel-authoring` is the single Excel entry point. Its description and first routing section deliberately distinguish the standalone file route from the connected-session route so a detached file task is not sent to a live workbook and a live task is not silently rewritten as a detached file.
- `@oai/artifact-tool` is the authoring engine because it is the installed runtime contract for this machine. The checker uses only Python standard-library ZIP/XML APIs and never mutates a workbook.
- Existing workbooks are written to a new path by default. Macro-enabled or feature-rich workbooks require native Excel when the artifact runtime cannot preserve their package parts. Structural and visual evidence are reported separately from native Excel evidence.
- The shared Office deployment script now links three sibling sources and remains conflict-protective and idempotent. The former standalone `excel-live-control` source and its managed discovery links were removed as part of the 0.3.0 merge; the merged `excel-authoring` guidance was refined to 0.4.0 in this handoff. Unrelated sources, settings, and user-level skills are not removed or overwritten.

## 0.2.0 layout-readability update

The 0.2.0 guidance was added after reviewing a real multi-pass editable Gantt workbook. The reusable failure modes were undersized task labels, long detail text crowding the primary chart, wrapped rows clipped after nominal AutoFit, and a vertical freeze-pane divider crossing title/notes rows. The skills now separate primary/detail reading, normalize labels and ownership colours, check type at the intended insertion size, inspect wrapped rows, distinguish native pane dividers from borders, and require post-edit visual review. The original read-only checker now reports configurable small-font risks, shrink-to-fit, unresolved styles, frozen-merge crossings and print scaling; synthetic regression tests cover these invariants. These diagnostics do not claim human readability or native viewport acceptance. No course names, workbook contents or local acceptance receipts are included in the reusable fixtures.

## Receiver packaging clarification

The engine choices above describe historical development environments. The current excel-authoring 0.4.0 contract selects an available, permitted host engine by preservation and calculation requirements; artifact-tool is conditional on the host, not a universal dependency. No course documents, private runtime or source machine configuration is distributed.
