---
name: powerpoint-authoring
description: "Create or revise editable PowerPoint PPTX presentations with content-led layouts, relevant cited images and native acceptance. Use for academic/professional decks and image, alignment or crowded-slide corrections. Not for prose outlines without a PPTX request or complex template migration."
license: Apache-2.0
metadata:
  version: "1.0.0"
---

# PowerPoint authoring

Build slides around what the audience needs to understand. Use the existing brief, sources, language and time limit; proceed with sensible defaults when optional style choices are absent.

## Choose the change scope

Establish the output contract: requested formats, assigned slides, template authority and whether integration into another deck is requested. A group deck can be a style reference without becoming an authorized merged deliverable. Do not add PDF previews or combined decks merely because they are easy to generate. Verification intermediates belong in temporary/internal evidence locations; if conversion itself is prohibited, use permitted native views or record the precise visual-check gap.

For a new deck, keep a compact slide plan: principal claim or purpose, source, visual role, essential visible qualifications, and detail reserved for notes. A title, transition or reference slide need not assert a claim. Reuse the plan during final content and image checks.

For a targeted edit or a mandatory template, first read [existing-deck editing](references/existing-deck-editing.md). Preserve slide/layout/master relationships, theme, rich text and editable data; the bundled new-deck generator does not implement template-preserving editing. Apply the review to the changed content and its dependencies without turning a local correction into an unsolicited redesign.

## Plan and build

1. Map requirements to slide claims, evidence and visual roles. Establish author voice (individual/team/institution), viewing mode (projection/desktop/both), and the required template and page limit before layout. For individual authored work use I/my for the author's actions and judgements, including AI disclosure; preserve attributed claims and quotations. Use available feedback to improve analogous slides, not just the marked example. Separate projected text, notes and citations. For coursework and AI disclosures, read [authorship-and-disclosure.md](references/authorship-and-disclosure.md).
2. Select a canvas and a small set of fonts, colors and spacing choices. A plain or template-free deck can still contain relevant images and icons; do not silently replace a requested image-based deliverable with text boxes. When a supplied template is mandatory, preserve its required structure and branding, inspect slide/layout/master elements in a native preview, and reserve their visible areas before adding anything. A deck loaded from the template can still violate it through foreground overlays. Inspect actual font availability; a font name alone does not guarantee the recipient has it.
3. Match layouts to relationships: ordered stages for a process, aligned alternatives for comparison, appropriate charts and units for data. For new layout decisions or density/fragmentation/whitespace feedback, read [information and layout](references/information-and-layout.md). Decide what the audience needs on screen before adjusting boxes. See [design-and-validation.md](references/design-and-validation.md) for fit and acceptance.
4. Inspect candidate images before placing them. Judge recognisability and label size at the intended slide size, not by download success or pixel dimensions alone. For images, read [images-and-attribution.md](references/images-and-attribution.md). There is no required picture quota; use images where they explain or orient, and tables or native diagrams where those communicate better.
5. Create editable text, shapes, tables and native charts; embed selected images with provenance. `scripts/create_pptx.py` supports image/text layouts and images in comparison/process items as well as the basic layouts. Read [creation.md](references/creation.md) for its bounded contract. Adapt a task generator when needed; do not force an unsuitable brief into the starter.
6. For projection or mixed use, start with 22–26 pt substantive text, including table cells, chart labels and essential qualifications; treat smaller type as a review trigger, not a way to fit more text. Disable shrink-to-fit in task generators. When revision is authorized, remove repeated explanation and move supporting detail to notes before shrinking fonts; preserve existing notes, source wording and essential limitations. Keep required content visible and resolve page-limit conflicts rather than silently adding slides. Synchronize current wording, references and layout plans. Preserve submitted/approved artifacts: feedback authorizes learning, not silent replacement of their bytes.
7. Preview a representative image slide and the densest text/data slide early. Resolve information density and uneven visual weight before repeating the layout across the deck. Then build the requested presentation and inspect it completely. When integration is requested, compile modules with their dependencies; do not join decks by copying shape XML alone. Preserve previous user-editable versions before replacement.

## Verify the final file

Run the content/geometry inventory and package check first. Run the native PowerPoint check as a separate, sequential operation; do not place it in a parallel tool batch with other checks, especially when `--grant-export-dir` requires an escalated approval:

```bash
"<office-python>" "<skill-dir>/scripts/check_pptx.py" "final.pptx" --spec "presentation.json"
python3 "<skills-root>/word-authoring/scripts/check_office_package.py" "final.pptx"
python3 "<skills-root>/word-authoring/scripts/verify_office_macos.py" "final.pptx" --pdf "/tmp/ppt-proof/final.pdf" --report "/tmp/ppt-proof/native.json"
```

Use `--spec` for the starter's JSON; otherwise compare actual embedded images against the task's plan. The checker reports missing planned images, image identity, crop/stretch and text-box risks. Warnings require inspection, not blind fixes or automatic acceptance. It cannot establish exact glyph wrapping, visual quality or group geometry. Package validation also checks namespace-valued attributes.

For projected or mixed-use decks also run `--min-font-pt 22`. For supplied templates pass inspected `--protected-regions regions.json`; see [design-and-validation.md](references/design-and-validation.md) for the contract. These checks flag explicit small type, shrink risks and objects crossing reserved areas, including decorative bars. They do not infer whether a source line is an acceptable small-type exception or discover every inherited element.

For macOS native checks and file-access prompts, read [native-verification.md](references/native-verification.md). Follow its approval-timeout recovery protocol. Ordinary open, completed export and visual acceptance are distinct results; a successful AppleEvent or a missing temporary PDF proves neither layout quality nor document corruption.

Render the internal verification export and view **every slide**, including references, both fitted to an ordinary screen and at readable resolution for defects. Zoomed-in legibility is not evidence of audience readability; record whether projection/distance was actually tested or only simulated. Check the principal message, reading order and balance as well as image relevance, legibility, clipping, glyphs, chart values, source lines and intentional layering. Compare text and numbers against the plan; inspect notes separately. After any final-file change, rerun structural/native checks and render again. Previously reviewed pages can carry visual evidence only when the new rendered image bytes match exactly under the same render settings; inspect changed pages again. Deliver only the contracted artifacts, with a single current entrypoint; keep QA intermediates out of the user-facing deliverable list.

If PowerPoint is unavailable, use an available renderer and record the native acceptance gap. Diagnose an installed application's sandbox/automation failure using the native reference before declaring it unavailable. Deliver with an explicit acceptance gap when unresolved; never call the file submission-ready from package/geometry checks alone or silently transfer first visual QA to the user. The helper is for macOS; Windows can use native automation or manual ordinary-open/export checks. A filename, library reload, montage or another application's rendering does not establish native compatibility.

## Dependencies and boundaries

Use `python-pptx==1.0.2` in an isolated environment. Install both owned skills with the suite installer; Word supplies shared read-only checks with a single source of truth. If the companion is missing, restore it before claiming those checks ran. Keep content and receipts outside reusable skills. Creation authorization does not imply publishing or sending files.
