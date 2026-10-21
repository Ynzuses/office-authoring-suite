# Spreadsheet layout and readability QA

This reference turns recurring workbook-revision failures into reusable design and acceptance rules. It applies to standalone workbooks and, in the same visual spirit, to targeted edits in an already-open Excel workbook.

## Design before formatting

Decide what the reader must scan first and where the workbook will be read: in Excel, printed, projected, or inserted into a document. A primary sheet should answer that question quickly. When long descriptions crowd it, move them to a detail sheet or an agreed text appendix, retaining stable task IDs between the two. Comments suit optional context, not required text that must appear in a printout. Do not keep shrinking the primary view to fit text that has a different reading purpose.

For a Gantt chart, keep the task/person labels aligned with a continuous time axis. Show dates, durations, phases and milestones where they help the requested view, with editable schedule inputs available. A concise legend is useful only if it adds meaning: avoid repeating all names above a timeline when labels and matching colours already identify them. Normalize name order and capitalization consistently using the supplied identities, without inventing or altering names. Keep ownership colours stable across sheets; use neutral phase grouping if colour already means person, and label milestones distinctly.

Group small tasks only when dates, responsibility and meaning support the same summary; retain their action details under the summary's ID. Do not merge different owners or time spans merely to shorten the chart. Match timeline granularity to available evidence instead of inventing daily activity. Keep shared-workbook notes intelligible to recipients; private local paths and build/audit details belong in the local record, not a mandatory `Sources & Notes` tab. Preserve useful public citations and necessary assumptions in an appropriate form.

## Type and cell fit

Use role-based starting sizes and review the rendered result rather than trusting a nominal style:

| Role | Starting point for an ordinary editable workbook | Review rule |
| --- | ---: | --- |
| Sheet title | 14–16 pt bold | Must remain readable without zooming into a cropped image |
| Main task/person labels | 11–12 pt; start near 12 pt for a compact Gantt task column | Below 10 pt is a review trigger; reduce density before type |
| Column headers and phase labels | 10–11 pt bold | Keep the label and its unit visible together |
| Detail prose | 10–11 pt | Wrap intentionally and give the row enough height |
| Dense date headers or metadata | 9–10 pt only when the full sheet remains legible | Treat anything below 9 pt as a review exception, not a default |

These are adjustable starting points, not an Excel standard or a guarantee. Preserve mandatory templates and consider the intended reader. Turn off shrink-to-fit for key labels rather than letting the workbook silently reduce their type. If text does not fit, first shorten duplicated wording, widen the relevant column, increase row height, or move supporting detail to the detail sheet. AutoFit is a starting point; inspect wrapping and add modest row-height headroom where needed. Merged cells may not AutoFit reliably and need explicit heights.

For insertion into a report or slide, inspect a preview at the intended physical width as well as the full-resolution sheet. A 12 pt cell becomes roughly 6 pt when the whole chart is scaled to half its original width; increasing image resolution alone cannot restore readable type. Check print scaling, paper size, margins and print area when printing is intended. Do not force a wide chart onto one page or claim report-size readability from an enlarged image. If the destination width is unknown, record the preview width used and leave actual embedded-size acceptance unverified.

## Lines, panes, and layering

Freeze panes are navigation aids, not decorative rules. Put the freeze boundary below the title/notes block and column headers. In a compact chart whose title or notes span the sheet, prefer rows only (`xSplit = 0`); a vertical divider runs through the full viewport. Retain column freezing where it materially helps navigation and native inspection shows that it does not cut text or merged titles. Ordinary gridlines, cell borders, pane dividers, page-break guides and drawing objects have different causes: diagnose the line before changing formulas or disabling unrelated formatting.

Keep task-bar and milestone conditional formats inside their data ranges. Header borders and fills can be intentional, but do not apply blanket rules across title/notes or draw over text. Check actual layering; valid formulas and styles do not establish unobstructed reading. A sheet image may omit Excel's pane dividers and page-break guides: inspect saved pane settings and the native viewport when available, and disclose when native visual inspection is missing. A third-party image alone cannot prove that a native divider is harmless.

For presentation sheets, merge repeated phase labels only across their contiguous rows in the phase column; preserve task rows and the date grid. Keep sortable/filterable source or detail tables unmerged with one complete record per row. Recheck heights and editing after any merge; merged presentation ranges are not a substitute for a data table.

## Gantt-specific checks

- Keep the left task information and right date grid aligned on the same rows.
- Use a 0-day milestone with a distinct diamond or equivalent marker; do not render it as a misleading one-day bar.
- State whether duration means inclusive calendar days, working days or another unit; use that convention consistently and do not silently introduce a holiday calendar. Guard blank, text, or reversed dates with an explicit empty or `Check dates` result.
- Apply task bars only to data rows and milestone markers only to the date grid; do not let conditional formatting reach the title, notes, or header rows.
- Merge repeated phase labels only in the phase column and keep the task rows individually editable.
- Move a long `Details` field out of the primary chart when it prevents readable sizing; use a linked detail sheet or the user's agreed text appendix.
- Edit representative dates on a disposable exact-byte copy and verify duration, bar movement, blank/invalid-date behavior and milestone position. Preserve the delivered file. A re-import that returns only cached shared-formula values is not recalculation evidence; use native Excel or a supported engine for the behavior check.

## Static preflight

Run `python3 <skill-dir>/scripts/check_xlsx.py <file.xlsx> --layout --min-font-pt 10`, adapting the review threshold to the intended sheet. This read-only check inventories populated-cell font sizes, shrink-to-fit, saved panes, merges and print scaling. It flags freeze boundaries intersecting merged ranges and cells below the selected threshold, including supported row/column/named-style fallbacks and explicit rich-text runs. Review the affected roles; dense date labels can be a justified exception.

Its `layout_review.status` is `review_required`, `no_flagged_risks`, `incomplete` or `not_assessed`; `visual_acceptance` remains `not_assessed`. Structural `status: passed` and exit code 0 do not clear layout risks. The tool cannot measure glyph wrapping, nonmerged text spilling across dividers, conditional-format font overrides, shapes/charts, target-width readability or all application-specific style effects. Unknown style sizes remain review items. A clear preflight never replaces a rendered/native check.

## Required visual acceptance

Render every visible sheet after the last layout-sensitive edit and inspect it at ordinary reading scale, not only as a zoomed crop. For each sheet, explicitly check:

1. title and explanatory notes are unobscured;
2. headers, task/person labels, phase labels, dates, units, legend, and milestones are readable;
3. wrapped text has enough row height and no text is clipped, overlapped, or hidden behind a border, pane divider, fill, or merged cell;
4. no `#####`, accidental blank sheet, excessive whitespace, or content outside the intended working area is present;
5. frozen rows/columns help navigation and do not introduce a grey divider through the primary reading area;
6. colours remain distinguishable and are not the sole carrier of a necessary label;
7. the full width and height still communicate the intended hierarchy without requiring extreme zoom.

After any failure, change the layout or content hierarchy, then re-render every affected sheet. Do not treat a package check, formula scan, or successful Excel open as evidence of visual readability. When native Excel is available, open the exact final path, recalculate/read back key cells, and bind the final hash only after the last save.

## Acceptance record

Record the workbook path, sheet names, final hash, the type-size decisions, any intentional small-type exceptions, pane splits, merged ranges, and visual findings. Keep temporary renders and native receipts outside the reusable skill and outside the course deliverable unless the user requests them.
