# Namespace compatibility regression

This well-formed XML has declared element and attribute names but an undeclared prefix inside a compatibility value:

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 mc:Ignorable="future"><w:body/></w:document>
```

A serializer may preserve the value while dropping an `xmlns:future` declaration used only there. A local Word failure exhibited this mismatch; our portable fixture uses synthetic names and no private document.

The checker tracks namespaces at each element, including inherited/local scopes and sibling boundaries. It covers MC `Ignorable`, `MustUnderstand`, `ProcessContent`, `PreserveAttributes`, `PreserveElements`, `Choice/@Requires` and `xsi:type` across all XML/rels parts. Aliases are resolved by namespace identity. Ordinary XML parsing is a separate necessary check.

For repairs, preserve the original bytes and hash. Inventory rich features and retain unrelated ZIP parts. Prefer a library that preserves the required features; when XML surgery is necessary, preserve namespace mappings and element order. `lxml` helps preserve declarations but still requires post-save checks. Resolve missing prefixes using a trusted original/template and relevant OOXML documentation; never guess URIs or remove compatibility attributes wholesale.

Compare content and formatting with the original. Recovery opening may support diagnosis but is not acceptance. Check the **final path and SHA-256** in normal Word; changing a candidate's name requires checking identical bytes and opening the actual final path. Re-saving invalidates previous acceptance.

These helpers check selected structural invariants, not the full ISO/ECMA schema, all extension rules, rendering or semantics. Keep content, structural, native and visual evidence separate.

## Native automation and font checks

PowerPoint can return a POSIX path as `full name` on macOS; coercing that string directly to an AppleScript alias fails even though the presentation opened. The helper handles POSIX and HFS forms separately, checks the exact target, and never treats automation failure as document corruption.

The shared verifier's PowerPoint backend is `scripts/powerpoint_macos.py`. It uses indexed document access, Launch Services ordinary opening, one unique HFS export destination and a parser-checked complete PDF before closing its own document. It preserves already-open targets and refuses unsaved changes; the Word branch remains separate. See the sibling PPT skill's [native verification guide](../../powerpoint-authoring/references/native-verification.md) for the bounded access-prompt option and failed-run handling.

Word's blank-template theme attributes can override explicit font names. New-document generation normalizes those theme font attributes and writes explicit run mappings. Do not apply that normalization indiscriminately to an existing formatted document. Review actual rendered fonts and glyphs; successful XML assignment does not prove the application used the named font.
