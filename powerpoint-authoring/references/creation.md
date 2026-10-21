# Structured slide creation

```bash
"<office-python>" "<skill-dir>/scripts/create_pptx.py" "presentation.json" "deck.pptx"
"<office-python>" "<skill-dir>/scripts/check_pptx.py" "deck.pptx" --spec "presentation.json"
```

Requires `python-pptx==1.0.2` and the suite's locked dependencies. Output creation refuses existing files/symlinks. Paths to local image assets resolve relative to the JSON file. See [basic example](../examples/presentation.json) and [illustrated example](../examples/illustrated.json); all example data/artwork are synthetic, not scientific evidence.

Top-level fields: `title`, `font`, `style` (`accent` default, or `plain`), optional six-digit RGB `accent`, optional `assets`, and nonempty `slides`. The plain style uses black text/shapes on white, suppresses default shape effects and normalises new-deck hyperlink colours. It preserves image colours. Neither style is a supplied external template.

Each slide has `layout`, `title`, optional `source` and `notes`, plus:

| layout | Fields |
|---|---|
| cover / content | `body`, with optional line breaks |
| comparison | exactly two `items`, each with `title`, `body`, optional `image` asset ID |
| process | two to four `items` with the same fields |
| chart | matching `categories` and finite numeric `values`; `series`, `unit` |
| image | `image` asset ID, `body`, optional `caption`, `side` (`left` default or `right`) |

The image layout reserves a large image pane, adjacent editable text, caption and credit area. Item images use compact contain-sized regions above aligned headings; they suit short icon/illustration labels, not complex figures. Mixed or denser layouts need an adapted task generator. There is no instruction to add images to every item or slide.

## Asset contract

`assets` maps IDs to objects:

```json
{
  "figure-a": {
    "path": "images/figure.png",
    "kind": "figure",
    "credit": "Author (year), Fig. 2",
    "license": "CC BY 4.0; unchanged",
    "source_url": "https://example.org/source-page",
    "alt": "A short description of the relevant visible content"
  }
}
```

`path`, `kind` (`figure`, `photo` or `icon`), nonempty `credit` and `license` are required; `source_url`, `sha256` and `alt` are optional. Supply the source page for third-party work. An original/user-provided asset may lack a public URL; state its real provenance and reuse status rather than inventing a licence. The code checks metadata presence, not legal permission or factual accuracy. It loads only referenced assets; candidates can remain unused.

Use inspected PNG/JPEG images. Convert SVG separately, retaining the original and its notices. The generator preserves aspect ratios, embeds the original raster bytes, attaches the source URL to the image, writes a visible short credit and adds provenance to notes. Use concise credits that fit; full bibliographies can have a separate appendix.

`--spec` checks planned image hashes and occurrence counts per slide, zero crop for this contain-only contract, and provenance in notes/links. Without it the checker inventories actual pictures but cannot know what was omitted. Diagnostics distinguish errors from `review_required` warnings; warnings do not cause a nonzero exit by themselves and must be adjudicated in native visual review.

The canvas is 13.333 × 7.5 inches. Substantive body text, explanatory captions and chart axes/data labels use 22–24 pt with shrink-to-fit disabled; short source lines, credits and page numbers use 12 pt as metadata. These defaults do not guarantee that long input fits. Run `--min-font-pt 22` and review smaller metadata in context; adapt the layout for projected references or a different audience. Text, shapes and column-chart workbook data remain editable. Unsupported fields fail. Adapt the task generator for tables, other charts or supplied templates; do not shrink content to force it into these bounded layouts. The library does not measure PowerPoint's exact glyph wrapping, and box bounds are not a text-fit proof.
