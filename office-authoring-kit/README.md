# Office authoring support kit · stable group 1.0.0

This support kit accompanies Office Authoring Suite 1.0.0. The four sibling Skills cover Word, PowerPoint, Excel and preservation-aware document delivery, with separate content, structural, native-application and visual acceptance:

- [word-authoring](../word-authoring/SKILL.md): editable DOCX construction, explicit new-vs-existing document routing, namespace regression checks, generic package checks and native macOS Office verification.
- [powerpoint-authoring](../powerpoint-authoring/SKILL.md): editable slide construction, information selection and crowding repair, images with provenance, projected-type and template-region review, existing-deck/notes preservation and native verification/recovery. It depends on Word's shared helpers.
- [excel-authoring](../excel-authoring/SKILL.md): local workbook creation/revision or targeted edits in an advertised Excel session, with engine selection, formula/structure checks and readability review.
- [research-document-delivery](../research-document-delivery/SKILL.md): rich-document preservation, comments and tracked changes, cross-format synchronization and final-file validation. The historical ID remains for compatibility.

The packages contain original instructions and implementation, using public library APIs and independently written synthetic examples. The 0.6.0 iteration adds bounded existing-document/deck routes, host-runtime selection for Excel, and voice/profile evidence gates; it does not vendor third-party skill text or code. [PROVENANCE.md](PROVENANCE.md) records reviewed sources, licenses and selection decisions. Apache-2.0 applies to the authored files; earlier MIT notices are retained in the repository LICENSES directory. Dependency licenses remain separate. No proprietary skill source or private course material is bundled. Publication status is determined by a tagged release, not by a local installation.

The 0.6.1 documentation update accompanies PowerPoint 0.5.0: repair information density before resizing, preserve supporting explanations in notes, and keep internal QA exports distinct from requested deliverables. The generators, checkers and shared Office backend are unchanged.

## Setup and use

Use Python 3.10+ in an isolated environment. The locked dependency set was tested on Python 3.14; other platforms/interpreters require a compatible wheel set and their own smoke test.

```bash
python3 -m venv /path/to/office-env
/path/to/office-env/bin/python -m pip install -r requirements.txt
```

Generation and structural checks are portable. `verify_office_macos.py` requires installed Microsoft Word/PowerPoint and normal OS automation authorization. PowerPoint's backend also uses Poppler `pdfinfo` to validate completed export and slide count. It never disables security settings or accepts recovery dialogs. The optional `--grant-export-dir` handles only a recognised access prompt for its own temporary export directory. See [PowerPoint native verification](../powerpoint-authoring/references/native-verification.md). LibreOffice rendering is supplementary evidence, not Microsoft Office acceptance.

The installer links all four owned Skill sources into a chosen discovery directory. Excel host runtime availability is checked by the receiving environment; this package does not install a private runtime implicitly. It does not download third-party Skills or overwrite conflicting installations:

```bash
python3 deploy.py --skills-root /path/to/agent/skills
python3 deploy.py --skills-root /path/to/agent/skills --apply
```

Keep all source directories adjacent. Their deployment links resolve to those sources, so editing the owned source updates the installed skill without competing copies. If a source is moved, inspect and deliberately replace its old deployment link; do not let an update tool replace it with a marketplace package. Use source control to review and roll back changes.

## Iteration and tests

1. Change the relevant source, examples and instructions together.
2. Run `python -m unittest discover -s tests -v` with the isolated interpreter. This includes malformed compatibility attributes, broken relationships, content/data preservation, omitted/changed/cropped images, provenance, plain styling, text overlap/height risks, projected type and autofit, protected template regions, invalid/partial/wrong-page PDFs, false export success, output preservation, workbook package checks, and deployment conflicts.
   For Excel changes also run `python -m unittest discover -s ../excel-authoring-kit/tests -v`: read-only package checks, small/inherited/rich-text fonts, shrink-to-fit, frozen-merge crossings, ordinary split units, unresolved styles and print-scaling review.
3. Use the skill creator's `quick_validate.py` on all four Skills where available; PyYAML is included in the isolated dependency set. Run document-delivery regressions from the group root with `python -m unittest discover -s tests -v`.
4. When generation/rendering behavior changes, generate relevant examples to fresh paths and perform their structural/native/visual checks. For read-only diagnostic changes, use synthetic regression fixtures and unchanged real files without rewriting them or claiming fresh native acceptance. Review warnings explicitly; a risk is not a proven collision. Excel range images may omit pane dividers, and enlarged renders do not prove report-size readability. Keep receipts in the project audit directory, not in reusable skills. Any change to a final artifact invalidates its previous hash-bound acceptance.
5. Inspect discovery links and hashes. Update the version when behavior changes. Do not claim automatic activation in an already-open agent session solely because files exist; verify the discovery inventory or use the explicit skill path.

## Preparing a future public repository

Export the selected owned skill directories and required sibling helpers/test kits to a clean independent repository, preserving their relative links. Include LICENSE and PROVENANCE, run the tests, and scan both current files and any included history for private paths, content and credentials. The maintainer is Ynzuses and the distribution license is Apache-2.0; retain the component notices and dependency license boundaries. Do not export the parent coursework repository or local verification receipts. Local deployment does not authorize publication.
