# Existing decks and templates

Read this for a local revision, a supplied template, or a deck containing user edits.

## Preserve before editing

Identify the authoritative deck and requested change. Keep a byte-identical source and write a new candidate unless the user explicitly authorizes replacement. Record slide order and IDs, layouts/masters, notes, media, chart workbooks and any animations or embedded objects relevant to the edit. Static PDF review cannot verify interactive features.

Distinguish the assigned section from a full reference deck. Follow its theme and section rhythm, but generate a combined deck only when integration is requested. Do not create an extra preview format by default. If the previous candidate is open, preserve the user's buffer; use an unambiguous new candidate path and update the current entrypoint rather than closing their presentation to satisfy a verifier.

Inspect the actual inherited artwork and visible regions, including backgrounds, masters and grouped shapes. Use the measured protected-region contract in [design-and-validation.md](design-and-validation.md); placeholder coordinates alone do not describe every visible object.

## Choose a bounded mutation

- Replace text within the relevant runs when formatting must survive. Assigning a complete `text_frame.text` value through `python-pptx` removes the previous paragraphs/runs and their local formatting. If replacement spans multiple runs, reconstruct the intended formatting explicitly and inspect the result.
- When shortening projected content is authorized, preserve existing notes and the displaced explanations, with current speaking guidance distinct from historical source wording. Verify notes parts and relationships as well as slide parts; an archive merge that restores old notes can silently drop the new guidance.
- Keep charts editable and preserve their embedded workbook. Changing a visible chart label does not update its source series. Check categories, values, units and notes together.
- For a template with a different number of people or process items, adapt the complete visual group, including images and captions, within the user's allowed template changes. Empty text can leave misleading orphan graphics.
- Do not duplicate slide shape XML alone. Media, layouts, chart parts, relationships and registration must remain coherent; shared chart parts may couple two slides unintentionally. Use a dependency-aware tool or native duplication and verify independent edits where required.
- Do not rebuild a rich deck with the starter solely to make editing easier. If the required features cannot be preserved with available tools, report the exact gap and retain the source.

## Verify the candidate

Compare source and candidate: requested changes are present; slide count/order, unaffected text and notes, image identities, source links and editable data remain as intended. Classify each unexpected difference before delivery. Then run the main skill's package, native and visual checks. Review all final slides for template/layout changes; reuse prior page-level visual evidence only for byte-identical rendered pages under the same settings.
