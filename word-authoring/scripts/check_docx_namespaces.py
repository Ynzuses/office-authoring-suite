#!/usr/bin/env python3
"""Read-only check of namespace-valued attributes in every DOCX XML part.

This catches undeclared prefixes in markup-compatibility attribute values,
which an ordinary XML parser does not resolve. It is not an OOXML schema
validator and does not replace opening the final document in Microsoft Word.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile


MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
XML = "http://www.w3.org/XML/1998/namespace"
PREFIX_LISTS = {f"{{{MC}}}{name}" for name in ("Ignorable", "MustUnderstand")}
QNAME_LISTS = {
    f"{{{MC}}}{name}"
    for name in ("ProcessContent", "PreserveAttributes", "PreserveElements")
} | {f"{{{XSI}}}type"}


def inspect_xml(payload: bytes) -> list[dict[str, str]]:
    """Resolve prefixes at the element carrying each attribute, including scope."""
    findings = []
    scopes: list[dict[str, str]] = []
    pending: dict[str, str] = {}
    events = ET.iterparse(io.BytesIO(payload), events=("start-ns", "start", "end"))
    for event, node in events:
        if event == "start-ns":
            prefix, uri = node
            pending[prefix] = uri
        elif event == "start":
            scope = dict(scopes[-1]) if scopes else {"xml": XML}
            scope.update(pending)
            pending.clear()
            scopes.append(scope)
            for attribute, value in node.attrib.items():
                prefix_list = attribute in PREFIX_LISTS or (
                    node.tag == f"{{{MC}}}Choice" and attribute == "Requires"
                )
                qname_list = attribute in QNAME_LISTS
                if not (prefix_list or qname_list):
                    continue
                for token in value.split():
                    # Unprefixed QNames do not reference an undeclared prefix.
                    if qname_list and ":" not in token:
                        continue
                    prefix = token.split(":", 1)[0] if qname_list else token
                    if not scope.get(prefix):
                        findings.append({
                            "element": node.tag,
                            "attribute": attribute,
                            "undefined_prefix": prefix,
                        })
        else:
            scopes.pop()
            node.clear()
    return findings


def check(path: Path, required_part: str = "word/document.xml") -> dict:
    result = {"path": str(path), "sha256": None, "xml_parts_checked": 0,
              "issues": [], "read_errors": []}
    try:
        # Hash and inspect exactly the same byte snapshot.
        payload = path.read_bytes()
        result["sha256"] = hashlib.sha256(payload).hexdigest()
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            names = archive.namelist()
            if required_part not in names:
                result["read_errors"].append(f"Missing {required_part}")
            if len(names) != len(set(names)):
                result["read_errors"].append("Duplicate ZIP member names")
            corrupt = archive.testzip()
            if corrupt:
                result["read_errors"].append(f"ZIP CRC failure: {corrupt}")
            for member in archive.infolist():
                if not member.filename.endswith((".xml", ".rels")):
                    continue
                try:
                    findings = inspect_xml(archive.read(member))
                    result["xml_parts_checked"] += 1
                    result["issues"].extend({"part": member.filename, **f} for f in findings)
                except ET.ParseError as exc:
                    result["read_errors"].append(f"{member.filename}: {exc}")
    except (OSError, zipfile.BadZipFile, RuntimeError, NotImplementedError) as exc:
        result["read_errors"].append(str(exc))
    result["status"] = "unreadable" if result["read_errors"] else (
        "failed" if result["issues"] else "passed"
    )
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("documents", nargs="+", type=Path)
    parser.add_argument("--json", action="store_true", help="Print structured results")
    args = parser.parse_args()
    results = [check(path) for path in args.documents]
    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for result in results:
            print(f"{result['status'].upper()}: {result['path']}")
            print(f"  SHA-256: {result['sha256']}; XML parts: {result['xml_parts_checked']}")
            for finding in result["issues"]:
                print(f"  {finding['part']}: {finding['attribute']} -> "
                      f"undefined prefix {finding['undefined_prefix']!r}")
            for error in result["read_errors"]:
                print(f"  READ ERROR: {error}")
        print("Scope: namespace references only; Microsoft Word normal-open verification is still required.")
    return max((2 if r["read_errors"] else 1 if r["issues"] else 0) for r in results)


if __name__ == "__main__":
    raise SystemExit(main())
