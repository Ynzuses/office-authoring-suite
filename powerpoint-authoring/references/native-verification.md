# Native PowerPoint verification on macOS

Use the companion Word skill's `scripts/verify_office_macos.py`; the PowerPoint backend lives beside it as the single shared implementation. It needs `osascript`, installed PowerPoint, normal macOS automation access and `pdfinfo` from Poppler. No global application/security settings are changed.

Native exports here are verification evidence, not an automatic request for a user-facing PDF. Keep them in an internal/temporary QA location when the deliverable is PPTX only. If the user prohibits conversion itself, inspect through permitted native views and explicitly record which checks could not run.

```bash
python3 "<skills-root>/word-authoring/scripts/verify_office_macos.py" "final.pptx" \
  --pdf "/tmp/slide-proof/final.pdf" --report "/tmp/slide-proof/native.json"
```

When the task authorizes local Office export, the caller may add `--grant-export-dir`. This handles a recognised English/Chinese Office file-access prompt **only for the unique temporary directory created by this invocation**. It is not permission to click an unknown dialog or grant access to a parent directory. Other localisations or unexpected UI require inspection; use existing task authorization rather than inventing a new approval ritual.

## Tool approval and retry protocol

The native command has two separate failure surfaces: the verifier can start and return a JSON receipt, or the surrounding execution/approval layer can reject or time out before the verifier starts. Keep these surfaces separate.

1. Finish structural and package checks before starting native verification. Invoke the native command in its own sequential tool call; do not run it through `Promise.all`, `Promise.allSettled`, a parallel shell batch, or alongside another escalated command. Approval prompts compete for the same review channel and can time out even when Office is healthy.
2. Use a fresh PDF and JSON receipt path for every attempt. Never reuse a partial output; the verifier deliberately refuses existing outputs. Record the exact source path and hash before each attempt.
3. If the tool call itself reports an approval timeout or a review-service disconnect and no verifier JSON receipt exists, wait 20–30 seconds and retry the identical command once. If it fails again for the same infrastructure reason, wait once more and issue one final identical attempt. These are at most three total attempts; do not loop indefinitely and do not change the command to bypass review.
4. If the verifier returns a receipt, inspect its fields before retrying. A receipt with `native_open: passed`, `native_export: passed`, matching source hash, complete PDF and expected page count is usable evidence. A helper error, incomplete PDF, source-hash mismatch, or target-left-open state requires diagnosis of that concrete condition; an approval retry cannot repair it.
5. If all bounded approval attempts fail before a receipt is produced, record `approval_layer_unresolved` separately from `office_native_failed`. Preserve the exact error text, keep the candidate distinguishable, and report native acceptance as unresolved. A prior candidate's receipt or a PDF from different bytes cannot validate the final file. Structural/package checks and a third-party render may support diagnosis, but they do not close the native acceptance gap.

Do not parallelize the recovery attempts. A normal, non-escalated call that fails because the sandbox cannot access macOS automation is not evidence that PowerPoint is unavailable; after recording that result, use the same command with the supported escalation mechanism. If the escalation reviewer rejects the action for a policy reason rather than timing out, stop and report that reason instead of trying indirect automation.

## Completion conditions

1. Refuse existing outputs and already-open exact target documents, preserving their edits. Launch the requested file normally through macOS Launch Services. Enumerate presentations by numeric index and compare normalised full paths; collection-object iteration and direct alias opening can fail in some Office builds.
2. Export **once** to a unique HFS-form temporary destination. A fresh filename prevents an overwrite prompt from masquerading as an automation hang.
3. Wait within a bounded deadline. Require stable PDF size, a complete trailer and `pdfinfo` agreement with the actual slide count. A successful AppleEvent with no file is a failed export; a reply timeout is not proof that the PPTX is damaged.
4. Verify the source hash, copy the completed PDF unchanged to the requested output, and close only the helper-opened exact presentation without saving. Refuse to close a target with unsaved changes. Never close other documents or terminate Office to make the check pass.
5. Record ordinary-open and export results separately, source/PDF hashes, app version and any cleanup gap. Then render and review the PDF; native acceptance is not visual acceptance.

## A failed or interrupted run

An installed application's Launch Services/AppleScript connection failure in a sandbox is not evidence that Office is missing or the deck is corrupt. Inspect the application path and helper receipt; when local native verification is authorized, retry the same bounded operation using the environment's supported permission mechanism before declaring the capability unavailable. Apply the sequential approval-timeout protocol above rather than a blind retry loop. A syntax failure calls for checking the helper/version, not another approval attempt. Preserve other documents and their lock files; an Office `~$` owner file may represent active editing and must not be deleted to obtain a clean Git status. Exclude generated lock files from version control using the workspace's ignore policy.

Read its receipt and inspect the exact app/target before retrying. The helper leaves an incomplete export's own presentation open and temporary evidence intact; it does not blindly repeat save, close midway through export, accept repair, or clear unrelated dialogs. The receipt identifies whether this helper initiated opening and whether its target remains open. A subsequent agent can finish or close that task-owned target after checking its actual state and preserving edits; do not assume every open presentation belongs to the agent.

If export completed but the helper-owned buffer now reports `saved=false`, retain the original receipt and diagnose cleanup separately. Do not mark it saved artificially, overwrite the canonical candidate or discard the buffer. If a task-authorized SaveAs is needed, preserve only that verified helper-owned buffer to a fresh file in its unique export directory. Check that the directory and exact target still exist: a successful helper run may already have removed its temporary directory, and a receipt's path is not proof of current state. Construct POSIX-to-HFS text **outside** the application `tell` block; neither coercion nor a successful save reply proves a destination file was created. After SaveAs, verify the actual full path, created file and `saved` state, then re-enumerate by the new full path before closing. Never reuse the previous ordinal reference. An unexpected path or identity is a failed recovery step: retain evidence and inspect it before any further close or retry. Preserve unrelated user windows and report a remaining cleanup gap honestly.

Use a new output/receipt for a changed candidate and keep the earlier evidence distinguishable. Temporary files can disappear between sessions; retain receipts and the evidence needed for the claimed verification in an internal QA location, without adding an unwanted preview deliverable. A previous successful opening of another path or earlier bytes cannot validate the current file. If evidence is relocated, record its old/new path and unchanged hash rather than editing the historical receipt to imply a new native check.

Without completed native and visual checks, report the exact gap and retain a candidate status. Package validity, slide count, library reload and nominal box bounds are not a submission-readiness certificate. User screenshots are useful defect evidence; they do not substitute for reviewing all final pages or prove that a repair worked.
