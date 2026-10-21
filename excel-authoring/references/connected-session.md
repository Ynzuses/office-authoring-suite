# Connected Excel session

Use this route only for an open Microsoft Excel workbook with advertised session commands. Read the selected runtime's instructions and schemas; do not invent commands or switch to Office.js, macros, shell automation or a detached rewrite to bypass the session.

1. Read the workbook identity and relevant sheets, ranges, formulas, tables, charts and formatting. Reconfirm the identity before writing. An instruction to edit the current workbook is enough authorization for its targeted edit; no extra output filename is required.
2. Determine whether the request requires a saved/exported file or only an in-app change. If saved delivery is essential but the session offers no save/export capability, explain that gap before making a write that cannot satisfy the request. If editing in memory is acceptable, proceed within the authorized scope and report persistence as unverified unless the host proves it.
3. Capture the affected range image before a substantial layout change. Make the smallest mutation, preserving validations, hidden sheets, formulas and native objects outside scope.
4. Re-read changed ranges and dependent formula results. Recalculate only with an advertised command; restore calculation mode if it was changed. For layout-sensitive changes, follow [layout-readability.md](layout-readability.md), recapture the affected view, and inspect native pane dividers when available.
5. Save/export only through an advertised command when requested, then verify the resulting identity and relevant values. A successful edit reply does not prove saving; lack of save capability does not undo an in-memory mutation. Report separately what changed, what was re-read, and what persistence/visual checks remain unknown.

Do not install add-ins, enable macros, refresh links, bypass protection, close unrelated workbooks or change host settings without corresponding authorization. If the target session changes mid-task, stop further writes and resolve identity before continuing.
