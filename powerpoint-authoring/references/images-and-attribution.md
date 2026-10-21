# Images and attribution

## Choose for a specific role

Distinguish evidence figures (data, anatomy, clinical appearance, mechanisms), explanatory illustrations and navigation/decorative icons. A cited scientific claim does not turn a decorative icon into scientific evidence. A native diagram can be the best representation of a process, but drawing boxes around prose does not fulfil a request to use relevant supplied pictures.

Before building an image-bearing deck, inspect the existing asset library and plan each selected asset's slide, purpose, placement and readable size. Record excluded candidates briefly when their rejection matters: unclear subject, unreadable labels, outdated content, unsuitable reuse terms or duplication. A download inventory is not an insertion plan, and not every downloaded image needs to appear.

- View the actual file, then the final slide. A technically adequate resolution can still show an unrecognisable crop or blurry subject. Replace a poor source when possible; enlargement does not recover information.
- A figure should show the claimed phenomenon and enough anatomical or experimental context to interpret it. A label cannot rescue an image whose subject the audience cannot recognise.
- Match the evidence level and population: normal anatomy locates an organ but does not demonstrate pathology; a clinical example illustrates appearance but does not diagnose the presented case. Keep an age group, model or experimental condition visible when it changes interpretation.
- Prefer already suitable assets and local layout work. Do not invoke paid image generation just to fill space or create basic arrows/icons. Do not invent clinical photographs or present a generated illustration as observed evidence.
- Plain backgrounds and black text do not imply removing colour from scientific images. Preserve meaningful source colours unless an authorized, clearly identified adaptation is warranted.

## Place without crowding

Reserve the image, caption and short credit area before allocating body text. Give an explanatory scientific figure enough area for its useful content and labels; a small icon beside a long paragraph is not a substitute. Let the caption identify what to look at; remove prose that only repeats the image. Keep essential interpretive limitations visible. Text-only slides remain appropriate for compact comparisons or explanations with no useful external image.

Use contain sizing with the original aspect ratio by default. Do not crop, recolour, overlay annotations or split a published multi-panel figure without checking the applicable terms and scientific meaning. For an unchanged reproduction, preserve source bytes and verify zero crop. Record legitimate adaptations instead of claiming that they are unchanged.

When rasterising an SVG, distinguish intrinsic display dimensions from its viewBox. A large PNG containing a tiny drawing in one corner is not a high-resolution icon. Set the intended display size, retain the viewBox/path geometry, and visually inspect the result. Keep the original SVG and licence; distinguish conversion from artwork revision.

## Retain provenance

For third-party images retain the source page, author/organisation, year when available, title/figure number, original file URL, licence or stated reuse terms, and any adaptation. Keep checksums and slide mapping in task-local records when reproducibility matters. A DOI alone is not a reuse licence; being online or hosted by a public institution does not establish public-domain status. Preserve separate credits for third-party components of a larger figure.

Use a readable short credit beside the figure or in the slide footer; put the full reference, source link and licence in notes and/or an image-credit appendix. Keep complete icon licence notices with the asset package. A hypothetical image approval workflow is unnecessary when the user has supplied the image or the intended reuse is already authorized; unresolved external reuse terms should guide asset choice, not trigger unrelated permissions.

Verify the selected assets actually appear on their intended slides. The generator's `assets`/`image` fields and the checker's `--spec` option provide this check for the bounded starter. A custom generator should compare its plan with PPTX picture objects and embedded bytes rather than infer insertion from captions or file existence.
