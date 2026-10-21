#!/usr/bin/env python3
"""Read-only selected DOCX/PPTX package checks, not full schema validation."""
import argparse
import io
import json
import posixpath
from pathlib import Path
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

from check_docx_namespaces import check as namespaces

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REF = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def check(path):
    main = {".docx": "word/document.xml", ".pptx": "ppt/presentation.xml"}.get(path.suffix.lower())
    if not main:
        raise ValueError("Only .docx and .pptx are supported")
    result = namespaces(path, required_part=main)
    result["package_errors"] = []
    result["external_relationships"] = 0
    if result["read_errors"]:
        return result
    errors = result["package_errors"]
    # Verify both passes used the same bytes, even if another process is writing.
    payload = path.read_bytes()
    import hashlib
    if hashlib.sha256(payload).hexdigest() != result["sha256"]:
        errors.append("Input changed during inspection")
        result["status"] = "failed"
        return result
    with zipfile.ZipFile(io.BytesIO(payload)) as z:
        names = set(z.namelist())
        for name in names:
            if name.startswith("/") or ".." in name.split("/") or "\\" in name:
                errors.append(f"Noncanonical package member: {name}")
        for required in ("[Content_Types].xml", "_rels/.rels"):
            if required not in names:
                errors.append(f"Missing {required}")
        if "[Content_Types].xml" in names:
            root = ET.fromstring(z.read("[Content_Types].xml"))
            defaults = {e.get("Extension") for e in root.findall(f"{{{CT}}}Default")}
            overrides = {unquote(e.get("PartName", "")).lstrip("/") for e in root.findall(f"{{{CT}}}Override")}
            for name in overrides - names:
                errors.append(f"Content type references missing part: {name}")
            for name in names:
                if name.endswith("/") or name == "[Content_Types].xml":
                    continue
                if name not in overrides and name.rsplit(".", 1)[-1] not in defaults:
                    errors.append(f"No content type: {name}")
        rel_ids = {}
        root_targets = []
        for name in sorted(n for n in names if n.endswith(".rels")):
            source = "" if name == "_rels/.rels" else posixpath.join(
                posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
            if source and source not in names:
                errors.append(f"Relationships source missing: {source}")
            ids = set()
            for rel in ET.fromstring(z.read(name)).findall(f"{{{REL}}}Relationship"):
                rid = rel.get("Id")
                if not rid or rid in ids:
                    errors.append(f"Missing or duplicate relationship ID: {name}: {rid}")
                ids.add(rid)
                if rel.get("TargetMode") == "External":
                    result["external_relationships"] += 1
                    continue  # Never follow links or download external resources.
                target = unquote(urlsplit(rel.get("Target", "")).path)
                dest = posixpath.normpath(target.lstrip("/") if target.startswith("/") else
                                         posixpath.join(posixpath.dirname(source), target))
                if not target or dest not in names:
                    errors.append(f"Missing relationship target: {name}: {dest}")
                if not source and rel.get("Type", "").endswith("/officeDocument"):
                    root_targets.append(dest)
            rel_ids[source] = ids
        if root_targets != [main]:
            errors.append(f"Root officeDocument target must be {main}: {root_targets}")
        for name in sorted(n for n in names if n.endswith(".xml")):
            for element in ET.fromstring(z.read(name)).iter():
                for key, value in element.attrib.items():
                    if key.startswith("{" + REF + "}") and value not in rel_ids.get(name, set()):
                        errors.append(f"Unresolved relationship reference: {name}: {value}")
    if errors:
        result["status"] = "failed"
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("file", type=Path)
    a = p.parse_args()
    try:
        result = check(a.file)
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as e:
        result = {"status": "unreadable", "error": str(e)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
