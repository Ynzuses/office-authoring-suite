# Local workbook authoring

## Select an available engine

Inspect the current host's spreadsheet tools and instructions before choosing a library. Host requirements take precedence over these fallback choices.

| Available capability | Appropriate route | Verification consequence |
| --- | --- | --- |
| Host spreadsheet runtime, such as `@oai/artifact-tool` | Use its actual documented APIs for creation, import, calculation, rendering and export | Follow operation markers/output conventions only when advertised; verify supported functions and preserved features |
| Installed `openpyxl` | Bounded `.xlsx` edits when the inspected features survive its round trip | It writes formulas but does not calculate them; arrange supported recalculation and rendering separately |
| Installed `xlsxwriter` | New `.xlsx` files | It cannot edit an existing workbook or calculate formulas; caches require independent validation |
| Native Excel or a permitted connector | Features requiring native fidelity, or calculation/visual acceptance | Preserve workbook identity and other open documents; use only capabilities allowed by the environment |
| None capable of preserving required features | Read-only diagnosis and a concrete dependency gap | Do not silently drop features, change extension or claim a verified workbook |

Do not install dependencies as a side effect. If a host mandates its own runtime, a missing runtime is a blocker for authoring in that host. A standalone recipient with no such mandate may select a suitable already-installed engine explicitly; record its name/version and limits. Neither a skill name nor a package being installed establishes its capabilities.

## Protect the existing file

Record source path, extension, hash and size. Preserve a byte-identical backup and write a new candidate unless replacement of that exact source is already authorized. Inspect sheets/order, formulas and caches, tables, validation, merged cells, conditional formats, charts, names, external links, hidden/protected sheets, macros and embedded objects before choosing the route. Render relevant sheets first where a renderer is available.

Do not round-trip a macro-enabled or externally linked workbook through a tool that cannot preserve it. `keep_vba=True` alone is not proof of full fidelity, and merely retaining `vbaProject.bin` does not validate signatures, controls or connections. A workbook loaded with `data_only=True` is for cached-value inspection; saving it substitutes those values for formulas. Read formulas and values separately without using the cache-only view as the editing authority.

Compare source and candidate package/feature inventories after edits. Preserve existing visual conventions and keep the mutation within requested ranges. Do not delete a feature merely to make an exporter succeed.

## Calculation and rendering

Use documented functions supported by the selected calculation engine. Test representative formulas before populating a large grid; do not impose a universal ban on modern Excel functions based on one engine's limitations. Check spill ranges and shared/array formulas explicitly when used.

Setting “recalculate on open” does not populate saved caches. If recalculation requires a program that rewrites the workbook, use a disposable candidate and compare features afterward; conversions may alter charts, links or formulas. Do not automatically refresh external data or enable VBA. Report unavailable source workbooks and unresolved calculation dependencies.

Retain a reproducible task builder where substantial generation benefits from it. Export to the user-requested path or the host's required output location. Complete the entrypoint's acceptance checks against the final bytes.
