# Native Excel verification on macOS

Native acceptance is separate from package and render checks. Microsoft Excel must open the exact final file in ordinary mode without repair or recovery, and the exported proof must come from that same final file. Keep the user's other workbooks open and do not close or save them.

Use the bundled `scripts/verify_xlsx_macos.py` only after structural checks. It opens the target read-only, records the workbook identity and sheet count, and asks Excel to calculate. It does not save or export the workbook, enable macros, or update external links. If Apple Events are denied, Excel is unavailable, or the target cannot be opened normally, report `native_open_unverified` with the concrete reason.

Visual evidence comes from the artifact runtime's sheet renders; if a task specifically requires a native Excel PDF, perform that export as a separate, user-authorized native workflow and inspect the resulting pages. A successful AppleScript call without a matching workbook identity and final SHA-256 is not an acceptance result.

Sheet renders and PDFs may omit native pane dividers or page-break guides. After changing panes, inspect the saved split rows/columns and use an available advertised native screenshot/UI route for obstruction checks; report `native_visual_unverified` if unavailable. The open/calculation helper does not capture the viewport or test date edits. For interactive schedules, use a disposable exact-byte copy to test representative input changes in a supported calculation engine, particularly when an importer supplies shared-formula caches instead of executable formulas. Do not save those test changes to the deliverable.
