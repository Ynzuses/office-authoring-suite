---
name: word-authoring
description: "Create editable Word DOCX documents from approved text or structured content, with semantic styles, bilingual typography, package checks, and native Word acceptance. Use for new reports, letters, handouts, or Markdown-to-Word delivery. Compose with research-document-delivery for complex revisions, comments, tracked changes or frozen submissions. Do not use for prose-only answers or courseware explanation."
license: Apache-2.0
metadata:
  version: "1.0.0"
---

# Word authoring

Create an editable document whose content, package integrity and native appearance have each been checked. Follow the user's scope, destination and deadline; use a modest layout when no template is supplied.

## Select the route

- For new documents, map the approved content to [the generator contract](references/creation.md). It accepts structured JSON, not arbitrary Markdown.
- For an existing file or mandatory template, inventory styles, sections, headers, numbering, fields, equations and review markup before selecting a tool. The starter cannot preserve these by rebuilding from extracted text. Use an available document-editing workflow, or a task-specific implementation that preserves the required features.
- For conversion, account for every source feature. Stop on missing images or unrepresentable content instead of silently omitting it. Keep relative asset paths anchored to the input file.

## Workflow

1. Identify authoritative text, audience, language, template and final filename. Reuse established choices and proceed with authorized local generation. Keep citations tied to supplied or verified sources; this skill does not invent evidence.
2. Inspect relevant templates and available fonts. Distinguish a blank document from a rich existing file with comments, equations, fields or tracked changes. For the latter, preserve a byte-identical backup, record the source hash, and use the document-delivery workflow before touching XML. Keep the original and candidate paths distinct.
3. Use semantic headings, paragraphs, lists and tables. Set page geometry, margins and Latin/East Asian fonts explicitly. Match the requested template; do not impose a cover, contents page or ornament on a short assignment.
4. The bundled `scripts/create_docx.py` is a small structured-content generator using `python-docx`, not a complete Markdown parser. See [creation.md](references/creation.md). For unsupported features, write a task-specific generator or use native Word; never silently flatten equations, links, numbering or citations.
5. Save a candidate to a new path. Compare text, numbers, tables, references and required document features with the approved source. Check that local images are present and correctly resolved. Missing images, unsupported input or silently dropped features are failures, not placeholders to deliver.
6. After the last save, run both read-only checks:

   ```bash
   python3 "<skill-dir>/scripts/check_docx_namespaces.py" "final.docx"
   python3 "<skill-dir>/scripts/check_office_package.py" "final.docx"
   ```

7. On macOS with Word, use the native helper. It refuses already-open targets and existing output paths, opens normally with repair disabled, exports a PDF and checks source SHA-256. It never accepts recovery dialogs or globally closes the user's documents.

   ```bash
   python3 "<skill-dir>/scripts/verify_office_macos.py" "final.docx" --pdf "/tmp/word-proof/final.pdf" --report "/tmp/word-proof/native.json"
   ```

8. Render that PDF with available PDF tools and inspect every page. Prefer the workspace's existing rendering entry point. Check clipping, glyph substitution, page breaks, table overflow and references. Compare content against the source. Fix actual defects and repeat affected checks; do not invent defects to satisfy a loop.
9. Deliver the exact checked file. Any save invalidates earlier evidence. Report content, structural, native-open and visual checks separately, including the final path and hash. If Word is unavailable, deliver the best verified artifact with native compatibility explicitly unverified; a third-party renderer does not establish Word acceptance. Keep unsupported features and intentional exceptions visible in the handoff.

## Prevent unreadable content

An XML parser resolves element and attribute names, but may not resolve prefixes stored **inside attribute values**. A document can parse while `mc:Ignorable="future"` refers to an undeclared prefix and Word refuses to open it.

Do not round-trip rich OOXML through a serializer that drops these declarations. Preserve namespace scopes and extension markup; do not delete compatibility attributes or invent namespace URIs to silence a checker. Read [compatibility.md](references/compatibility.md) for the regression and repair boundaries.

## Tools and maintenance

Generator dependency: `python-docx==1.2.0`; validators and native helper use standard-library Python. Use an isolated environment. Setup is a separate operation, not an implicit download on every document task.

The paired `powerpoint-authoring` skill uses the generic package and native Office helpers here. Edit owned source, test, then deploy through the suite installer. Keep private content and verification receipts in the task workspace. Use specialist skills for evidence research or complex revision; file generation does not replace them.
