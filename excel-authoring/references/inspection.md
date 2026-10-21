# Excel inspection boundaries

`check_xlsx.py` is a small, read-only OOXML package check. It validates the ZIP container, required workbook parts, workbook-to-sheet relationships, duplicate sheet names, worksheet dimensions, and formula-error tokens in stored values. It also reports tables, charts, drawings, external links, defined names, and `vbaProject.bin` so a reviewer can notice preservation changes.

The check does not calculate formulas, prove that cached values are current, validate every Excel feature, or establish that Microsoft Excel will render the workbook exactly as the preview does. Use the artifact runtime for value/formula inspection and rendering, then use Microsoft Excel's ordinary open/calculation/export path when native acceptance matters.

When a workbook is edited, compare the pre-edit and post-edit reports. Unexpected changes in sheet count, relationships, formulas, package parts, external links, charts, tables, or VBA parts are review triggers. A warning is evidence to inspect, not permission to delete the affected feature.

`--layout` adds a separate static layout inventory; `--min-font-pt 10` also enables it and selects a configurable threshold. Review `layout_review` even when structural `status` is `passed`: it reports small populated-cell fonts, shrink-to-fit, unresolved styles, merged ranges crossing saved freeze boundaries, other split panes and print scaling. It reports risks, not proven collisions. Row/column/named-style fallbacks and rich-text sizes are inspected where present; arbitrary glyph wrapping, conditional-format overrides, drawings, zoom and physical output size are not resolved. See [layout-readability.md](layout-readability.md). No automatic flag or absence of flags establishes visual acceptance.
