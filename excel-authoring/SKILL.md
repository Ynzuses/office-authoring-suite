---
name: excel-authoring
description: "Create, edit, analyze, and verify local Excel workbook files, or make targeted changes in an open Microsoft Excel workbook through an advertised connected session. Route cloud spreadsheets to their connector."
license: Apache-2.0
metadata:
  version: "1.0.0"
---

# Excel authoring

Preserve user work and deliver an editable workbook whose content, calculations and appearance have been checked separately. Follow the current host's tool requirements; this portable skill does not supply a calculation engine or connected Excel service.

## Route by the requested target

- **Open workbook:** use the advertised connected Excel commands and read that runtime's live-control instructions. Read [connected-session.md](references/connected-session.md) before writes. A session can expose editing without save/export; resolve that distinction before promising a saved artifact.
- **Local file or new workbook:** read [file-authoring.md](references/file-authoring.md). Choose an installed, permitted engine by preservation and calculation capabilities. Prefer the host's spreadsheet runtime when provided or required; its private helpers are conditional, not a dependency every recipient must have.
- **Cloud sheet:** use the target connector's own workflow. Do not redirect an open/cloud workbook into a detached file rewrite.
- **CSV/TSV:** preserve encoding, delimiter, types and literal text; these formats cannot carry sheets, native charts or workbook formatting. Do not relabel text as `.xlsx`. For `.xls`, encrypted files or unsupported native features, diagnose read-only and use an authorized supported conversion.

Resolve ambiguous file/workbook identity before mutation. A filename, active window, or familiar sheet name alone may not identify the target.

## Keep the workbook auditable

Put raw inputs and assumptions in labelled cells; use formulas for derived values and typed dates/numbers with explicit units. Preserve existing formulas outside the requested change. Quote cross-sheet names, distinguish blanks from zero, and guard invalid inputs deliberately. A source column or concise note should explain researched inputs without exposing private build paths.

Before expanding a model, calculate representative formulas and check them against independently expected values. Include a boundary case relevant to the task: blank or zero denominator, reversed dates, first/last period, or milestone. Use a disposable copy for destructive input probes. A cache that looks plausible is not evidence of recalculation, and an error-free formula can still reference the wrong rows.

Use tables for sortable data and keep complete records on each row. Put dense supporting prose on a linked detail sheet only when that fits the requested deliverable. Read [layout-readability.md](references/layout-readability.md) for layout changes: review at the actual Excel, print, projection or document-insertion scale; shrink-to-fit and high-resolution zoom do not establish readability.

## Verify the exact final state

1. Reconcile representative totals, ranges, formulas and displayed results to the source. Recalculate with a supported engine and record the engine and any unverified functions, external data or missing caches. If an error predates the edit, compare the original cell and report it rather than silently changing unrelated logic.
2. Scan formula errors and compare sheet order, names, tables, charts, validations, names, hidden/protected content, links and macros before/after. A macro part's presence alone does not prove complete preservation.
3. For OOXML workbooks, run `python3 "<skill-dir>/scripts/check_xlsx.py" "final.xlsx" --layout --min-font-pt 10`, choosing the threshold for the reader. Read [inspection.md](references/inspection.md). Structural `passed` and `layout_review` are separate; this checker neither calculates nor establishes visual acceptance.
4. Render all visible sheets for a new workbook, or affected sheets and dependent views for a targeted edit. Inspect titles, labels, units, wrapping, charts, `#####`, merges, conditional formatting and print boundaries at ordinary reading scale. After layout changes, re-render the affected views. Pane dividers may require a native viewport check.
5. When available and permitted, open the exact candidate normally in Microsoft Excel and inspect key calculated values. See [native-verification.md](references/native-verification.md). The bundled read-only helper does not persist recalculated caches or prove formula correctness; do not present it as a substitute for saved-file calculation verification.
6. Bind evidence to the final path/hash (or live session identity and re-read state). Any later save invalidates that identity. Report content/formula, package, native-open and visual results with precise gaps. Missing a renderer or calculation engine leaves that check unverified; a library reload is not an equivalent check.

## Boundaries

Use available dependencies; do not install software, alter settings, enable macros, bypass protection, update external links, close unrelated workbooks, upload or submit files as an implicit part of authoring. Preserve sources and user edits. Keep builders, previews and receipts in the task workspace, outside reusable skill directories. A user request to edit the open workbook authorizes the targeted in-memory edit; it does not prove that the host saved it.
