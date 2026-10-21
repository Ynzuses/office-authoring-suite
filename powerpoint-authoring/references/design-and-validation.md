# Design and acceptance

Choose layouts from the audience's question and the relationship between facts. Processes, comparisons, charts and explanations have different reading orders. Do not impose a fixed proportion of pictures or decorative cards. For information selection, fragmented text and uneven visual weight, use [information and layout](information-and-layout.md) before changing geometry.

Write in the audience's reading order: introduce an unfamiliar idea before using it to explain a consequence or comparison. Use concrete subjects and direct verbs; avoid fragment piles, repetitive lead-ins and forced contrasts. When local writing skills or user-approved revisions provide a voice model, apply that guidance across analogous slides while preserving the evidence and scope. Shortening should preserve the reasoning, not merely reduce the word count.

## Readability for the actual audience

For projected 16:9 slides, start around 30–36 pt titles and 22–26 pt substantive text. This includes table cells, explanatory captions, axis/data labels and essential qualifications, not just paragraph text. Text below 22 pt requires a role-specific decision; 18–20 pt can suit secondary labels after visual review, while 10–15 pt paragraphs are generally unsuitable for projection. A larger projection does not automatically solve small type: viewers are also farther away. These are working defaults, not universal physical guarantees or a teacher's prescribed limits.

Use compact author/year or numbered citations near claims. Short source/credit lines and page numbers can use smaller type when they are metadata, with full references readable in the allowed appendix or handout. Do not hide substantive reasoning, AI disclosure or a qualification essential to the claim in tiny footnotes. Avoid displaying long encoded URLs; use readable bibliographic labels with tested hyperlinks and retain the full address in the permitted reference record. Dense reading appendices have a different viewing purpose; label them accordingly rather than pretending all pages project equally well.

Preview at fit-to-slide on an ordinary screen without zoom, then inspect at higher resolution for clipping. Use distance/projection checks if available; otherwise record the desktop proxy and its limitation. For screen-only reports, adapt type to that declared use. When the brief includes both reading and presenting, protect the projected core and allocate allowed detail separately. Do not add a second deliverable or extra pages silently.

Use readable contrast and avoid colour-only distinctions. When a plain deck is requested, choose the starter's `style: plain` or an equivalent blank-canvas build; do not impose the starter's optional accent palette.

## A supplied template is a constraint

Inspect representative source-template slides in native PowerPoint, including cover and content layouts. Inventory actual visible branding, footer text, page fields and icons across slide, layout, master, background and grouped objects; they may not appear in `slide.shapes`. Reserve their measured areas plus padding. Preserve required headings, order, identity and template elements unless the brief permits changes. Loading the original file is not proof of template fidelity: a full-canvas background or thin top bar can cover inherited artwork.

Check all foreground objects against reserved areas, including decorative bars and custom footer/page numbers. Prefer omitting an unnecessary shape over adding an overlay to conceal the collision. Do not hide/delete inherited branding to make a custom design fit. Keep title text inside its own fill with adequate left/top and rounded-corner padding; a colour strip is not a substitute for matching the text and fill bounds. Allow the final heading to wrap and measure its actual height, rather than guessing width from character count. Use a simple rectangle when large corner radii consume the usable region.

## Fit text with the visual

Allocate independent regions for title, body/visual, captions and sources, with space between them. Use the final wording during layout. A text box's nominal height does not clip overflowing glyphs reliably: text may run into the next heading even while every shape is inside the canvas.

When space is tight, first shorten redundant text or move explanations to notes if rewriting is authorized. Retain required comparisons, key causal steps and qualifications that change the claim. Preserve existing notes and source wording rather than replacing them with only the abbreviated slide text. Then adjust the image/text proportions, spacing or an allowed slide split. Do not rescue a crowded layout by making all text smaller. Fixed-height reference blocks are especially fragile when a citation wraps; allow space for the rendered citation before its hyperlink.

Disable automatic shrink-to-fit in the generator. Audit any inherited autofit settings and effective exported size; a nominal 24 pt run can render smaller. If a fixed template/page limit prevents a readable fit, cut repetition or use permitted supporting material; do not silently override either the limit or readability.

For new plain shapes, suppress unwanted theme effects such as shadows. Set hyperlink theme colours deliberately when monochrome styling is requested; direct text colour alone may not override PowerPoint's hyperlink theme. Do not normalize an existing branded template indiscriminately. Check inherited logos and master elements in native rendering before adding overlapping content or deleting legitimate branding.

## Inspect the result, not just the specification

| Check | Evidence | Limit |
|---|---|---|
| Source comparison | Claims, text, data and notes match the current input | Does not prove placement |
| Image plan comparison | Intended images are embedded on the correct slides with provenance | Does not prove relevance, recognisability or permission |
| Package/namespace checks | Selected ZIP/XML/reference invariants | Not full schema validation |
| Shape and risk inventory | Bounds, image dimensions, crop/stretch, intersecting text boxes, explicit-line height risks | Wrapping and intentional overlaps still need native inspection; group transforms are not resolved |
| Type and template risk checks | Explicit font sizes, autofit and supplied protected-region intersections | Does not resolve every theme/inherited size, glyph, raster label or visible master element; review unresolved cases |
| Ordinary native open/export | Exact file opens and a complete PDF is produced | Does not prove visual quality |
| Individual rendered-slide review | Images, type, labels, references and margins are readable | Not evidence that factual claims are correct |

For side-by-side layouts, `pdftotext -layout` may interleave columns, falsely suggesting missing phrases. Compare text per object/region, or use `-raw` when it preserves the export's object order; investigate actual omissions. Neither extraction mode establishes visual acceptance. Inspect tables, image labels, notes and charts with their appropriate evidence.

Review every page including citation/credit appendices, not only a montage. Fix demonstrated defects and repeat checks on the final file. A later save invalidates earlier file-bound acceptance. Identical final render bytes under identical settings can reuse a prior individual page review; changed pages require fresh inspection. Record the actual paths/hashes and explicit gaps in the task's receipt. Keep a single current entrypoint; prior accepted versions and drafts remain distinguishable. Store internal render/export evidence separately from the requested deliverables. Moving a checked artifact requires an explicit path/hash mapping; a same-named file must not inherit another file's receipt.

For animations, links, media and interactions, test behavior in the target application separately. Static PDF export does not exercise them. One macOS acceptance does not promise universal cross-platform fidelity.

## Optional machine-readable review inputs

Run `check_pptx.py deck.pptx --min-font-pt 22` for projected or mixed-use decks. This inventories explicit run/paragraph sizes in shapes, tables and supported chart labels. All smaller text is flagged for role-based review, including references; the checker does not silently exempt footers based on position. Unresolved inherited sizes, chart coverage limits and shrink settings stay visible as review items. No warning-free result establishes readability.

For a template, add `--protected-regions regions.json`. This is a task-local array of regions measured from the source template (inches, top-left origin), not a bundled template:

```json
[
  {"slide": "all", "label": "header branding", "rect": [10.6, 0.0, 2.6, 0.7]},
  {"slide": 2, "label": "footer on this layout", "rect": [8.5, 6.9, 4.5, 0.5]}
]
```

`rect` is `[left, top, width, height]`; use a 1-based slide number or `all`. Adapt coordinates to the actual source; do not copy these synthetic measurements. The checker flags top-level objects that intersect a region, even empty decorative shapes. It cannot know which overlapping objects are legitimate template elements; adjudicate with source/final native previews. Group transforms remain unassessed. Retain the policy and its review alongside the task receipt.
