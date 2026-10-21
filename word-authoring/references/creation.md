# Structured Word creation

Run with a Python environment containing `python-docx==1.2.0`:

```bash
"<office-python>" "<skill-dir>/scripts/create_docx.py" "document.json" "report.docx"
```

Start from [the executable example](../examples/document.json). The output is created exclusively: an existing file or symbolic link is never overwritten. Inputs are UTF-8 JSON. Image paths resolve relative to the JSON file; no network resource is downloaded.

Top-level fields: `title`, `subtitle`, `language`, `font`, `east_asian_font`, `body_pt`, `page_size` (`A4` or `Letter`), `margin_cm`, and nonempty `blocks`. The macOS starter uses Arial with Arial Unicode MS for East Asian text; other systems default to Noto Sans CJK SC for East Asian text. These are availability-dependent defaults, not required fonts. Theme font attributes in the blank template are normalized and runs receive explicit mappings. Fonts are not embedded; verify availability and native rendering. Defaults are for a simple report, not a mandatory house style.

Supported blocks:

| type | Content fields |
|---|---|
| paragraph | `text` OR `runs` |
| heading | `level` 1–3; `text` OR `runs` |
| list | string `items`; optional Boolean `ordered` |
| table | `headers`, rectangular `rows` |
| image | local `path`, optional `width_cm`, `caption` |
| page_break | no other fields |

Runs contain `text` and optional Boolean `bold`, `italic`, `superscript`, `subscript`. Lists use Word styles. Tables have repeating headers. The footer contains a native PAGE field. Unknown keys and unsupported types stop generation rather than silently omitting content.

This starter does not implement automatic Markdown parsing, hyperlinks, nested list restarts, footnotes, TOC, equations, tracked changes, comments or arbitrary template preservation. Convert a source into the supported contract only if every required feature is represented; otherwise use a task-specific generator or a richer tool. Keep an authoritative source and compare all text, values and references after conversion. A successful generator exit only means a file was written; run the full skill acceptance workflow.
