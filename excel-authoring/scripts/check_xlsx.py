#!/usr/bin/env python3
"""Read-only structural and preservation checks for OOXML Excel workbooks."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
ERROR_RE = re.compile(r"(?:#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A|#NUM!|#NULL!)")


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def parse_xml(zf: zipfile.ZipFile, name: str, issues: list[dict]):
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        issues.append({"kind": "missing_part", "part": name})
    except ET.ParseError as exc:
        issues.append({"kind": "invalid_xml", "part": name, "detail": str(exc)})
    return None


def relationship_targets(zf: zipfile.ZipFile, rels_name: str, issues: list[dict]):
    root = parse_xml(zf, rels_name, issues)
    result = {}
    if root is None:
        return result
    for rel in root.findall(q(PKG_REL, "Relationship")):
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            issues.append({"kind": "invalid_relationship", "part": rels_name})
            continue
        if rel.get("TargetMode") == "External":
            result[rid] = {"target": target, "external": True}
            continue
        base = posixpath.dirname(rels_name)
        if rels_name.endswith(".rels"):
            # A relationship part lives in a _rels directory beside its source.
            source_dir = base.removesuffix("/_rels")
        else:
            source_dir = base
        result[rid] = {
            # A leading slash is package-root-relative, not relative to xl/.
            "target": (posixpath.normpath(target).lstrip("/") if target.startswith("/")
                       else posixpath.normpath(posixpath.join(source_dir, target))),
            "external": False,
        }
    return result


def cell_error_tokens(root: ET.Element):
    found = []
    for cell in root.iter(q(MAIN, "c")):
        ref = cell.get("r")
        for node in cell.iter():
            if node.text:
                for token in ERROR_RE.findall(node.text):
                    found.append({"cell": ref, "token": token})
    return found


def positive_number(value: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise ValueError("font threshold must be a finite positive number")
    return number


def cell_position(ref: str):
    match = re.fullmatch(r"\$?([A-Z]+)\$?([1-9][0-9]*)", ref)
    if not match:
        return None
    col = 0
    for char in match[1]:
        col = col * 26 + ord(char) - ord("A") + 1
    return col, int(match[2])


def layout_inventory(root, styles, shared_strings, min_font_pt):
    """Selected static risks, never a rendering or font-metrics proof."""
    def children(parent, name):
        element = None if parent is None else parent.find(q(MAIN, name))
        return [] if element is None else list(element)

    def at(items, value):
        index = int(value)
        if not 0 <= index < len(items):
            raise IndexError(index)
        return items[index]

    xfs, bases, fonts = (children(styles, n) for n in ("cellXfs", "cellStyleXfs", "fonts"))
    risk_list = []
    panes = [dict(p.attrib) for p in root.iter(q(MAIN, "pane"))]
    merges = [m.get("ref", "") for m in root.iter(q(MAIN, "mergeCell"))]
    for pane in panes:
        if pane.get("state") not in ("frozen", "frozenSplit"):
            risk_list.append({"kind": "split_pane_requires_native_review", "pane": pane})
            continue  # Non-frozen split offsets are twips, not row/column counts.
        for axis, key, index in (("vertical", "xSplit", 0), ("horizontal", "ySplit", 1)):
            try:
                split = float(pane.get(key, "0"))
            except ValueError:
                risk_list.append({"kind": "unresolved_pane_split", "pane": pane})
                continue
            if split <= 0:
                continue
            for ref in merges:
                bounds = ref.split(":")
                a, b = cell_position(bounds[0]), cell_position(bounds[-1])
                if a and b and a[index] <= split < b[index]:
                    risk_list.append({"kind": "freeze_divider_crosses_merge", "axis": axis,
                                      "range": ref, "split_after": split})

    col_styles = children(root, "cols")
    small, shrink, unknown = [], [], []
    font_sizes = []
    populated = 0
    for row in root.iter(q(MAIN, "row")):
        for cell in row.findall(q(MAIN, "c")):
            value = cell.find(q(MAIN, "v"))
            if (cell.find(q(MAIN, "f")) is None and (value is None or not value.text)
                    and not any(t.text for t in cell.iter(q(MAIN, "t")))):
                continue
            populated += 1
            ref = cell.get("r", "?")
            try:
                style_id = cell.get("s")
                if style_id is None and row.get("customFormat") in ("1", "true"):
                    style_id = row.get("s")
                if style_id is None:
                    pos = cell_position(ref)
                    for col in col_styles:
                        if pos and int(col.get("min", "0")) <= pos[0] <= int(col.get("max", "0")):
                            style_id = col.get("style", style_id)
                xf = at(xfs, style_id or 0)
                base_id = int(xf.get("xfId", "0"))
                base = bases[base_id] if 0 <= base_id < len(bases) else None
                # Explicit child overrides plus the named-style fallback. Missing
                # or disabled properties fall back to the base where available.
                font_id = xf.get("fontId")
                if base is not None and (font_id is None or xf.get("applyFont") in ("0", "false")):
                    font_id = base.get("fontId")
                font = at(fonts, font_id or 0)
                size_node = font.find(q(MAIN, "sz"))
                size = None if size_node is None else positive_number(size_node.get("val", ""))
                alignment = xf.find(q(MAIN, "alignment"))
                if base is not None and (alignment is None or xf.get("applyAlignment") in ("0", "false")):
                    alignment = base.find(q(MAIN, "alignment"))
                if alignment is not None and alignment.get("shrinkToFit") in ("1", "true"):
                    shrink.append(ref)
                rich = cell.find(q(MAIN, "is"))
                if cell.get("t") == "s":
                    rich = at(shared_strings, cell.find(q(MAIN, "v")).text)
                runs = [] if rich is None else rich.findall(q(MAIN, "r"))
                sizes = []
                if runs:
                    for run in runs:
                        run_size = run.find(f"{q(MAIN, 'rPr')}/{q(MAIN, 'sz')}")
                        sizes.append(size if run_size is None else positive_number(run_size.get("val", "")))
                else:
                    sizes.append(size)
                if None in sizes:
                    unknown.append(ref)
                known = [s for s in sizes if s is not None]
                font_sizes.extend(known)
                if min_font_pt is not None and known and min(known) < min_font_pt:
                    small.append({"cell": ref, "min_font_pt": min(known)})
            except (ValueError, IndexError, TypeError, AttributeError):
                unknown.append(ref)

    for kind, items in (("small_cell_font", small), ("shrink_to_fit", shrink), ("unresolved_cell_style", unknown)):
        if items:
            risk_list.append({"kind": kind, "count": len(items), "samples": items[:100]})
    page = root.find(q(MAIN, "pageSetup"))
    page_setup = {} if page is None else dict(page.attrib)
    if any(key in page_setup for key in ("scale", "fitToWidth", "fitToHeight")):
        risk_list.append({"kind": "print_scaling_requires_target_size_review", "page_setup": page_setup})
    return {"populated_cells": populated, "font_size_range_pt":
            [min(font_sizes), max(font_sizes)] if font_sizes else None,
            "panes": panes, "merged_ranges": merges, "page_setup": page_setup,
            "risks": risk_list}


def check(path: Path, *, layout: bool = False, min_font_pt: float | None = None) -> dict:
    if min_font_pt is not None:
        min_font_pt = positive_number(str(min_font_pt))
        layout = True
    issues: list[dict] = []
    warnings: list[dict] = []
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    report = {
        "path": str(path),
        "sha256": digest,
        "extension": path.suffix.lower(),
        "status": "failed",
        "issues": issues,
        "warnings": warnings,
        "sheets": [],
        "formula_count": 0,
        "formula_error_tokens": [],
        "package_parts": {},
    }
    if layout:
        report["layout_review"] = {"status": "not_assessed", "min_font_pt": min_font_pt,
                                   "visual_acceptance": "not_assessed", "sheets": []}
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        issues.append({"kind": "not_ooxml_zip", "detail": str(exc)})
        return report

    with zf:
        names = zf.namelist()
        inspected_sheets = set()
        styles = parse_xml(zf, "xl/styles.xml", issues) if layout and "xl/styles.xml" in names else None
        shared_root = parse_xml(zf, "xl/sharedStrings.xml", issues) if layout and "xl/sharedStrings.xml" in names else None
        shared_strings = [] if shared_root is None else list(shared_root)
        report["package_parts"] = {
            "count": len(names),
            "tables": sum(n.startswith("xl/tables/") and n.endswith(".xml") for n in names),
            "charts": sum(n.startswith("xl/charts/") and n.endswith(".xml") for n in names),
            "drawings": sum(n.startswith("xl/drawings/") and n.endswith(".xml") for n in names),
            "external_links": sum(n.startswith("xl/externalLinks/") for n in names),
            "vba_project": "xl/vbaProject.bin" in names,
            "custom_xml": sum(n.startswith("customXml/") for n in names),
        }
        bad_names = [n for n in names if n.startswith("/") or ".." in Path(n).parts]
        for n in bad_names:
            issues.append({"kind": "unsafe_zip_path", "part": n})
        for required in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
            if required not in names:
                issues.append({"kind": "missing_part", "part": required})

        workbook = parse_xml(zf, "xl/workbook.xml", issues)
        rels = relationship_targets(zf, "xl/_rels/workbook.xml.rels", issues)
        if workbook is not None:
            defined = workbook.find(q(MAIN, "definedNames"))
            report["package_parts"]["defined_names"] = 0 if defined is None else len(list(defined))
            sheets = workbook.find(q(MAIN, "sheets"))
            seen_names = set()
            if sheets is None:
                issues.append({"kind": "missing_element", "part": "xl/workbook.xml", "element": "sheets"})
            else:
                for sheet in sheets.findall(q(MAIN, "sheet")):
                    name = sheet.get("name")
                    rid = sheet.get(q(DOC_REL, "id"))
                    entry = {"name": name, "relationship_id": rid, "state": sheet.get("state", "visible")}
                    report["sheets"].append(entry)
                    if not name:
                        issues.append({"kind": "sheet_without_name"})
                    elif name in seen_names:
                        issues.append({"kind": "duplicate_sheet_name", "name": name})
                    seen_names.add(name)
                    if not rid or rid not in rels:
                        issues.append({"kind": "unresolved_sheet_relationship", "name": name, "relationship_id": rid})
                        continue
                    rel = rels[rid]
                    if rel["external"]:
                        issues.append({"kind": "external_sheet_relationship", "name": name, "target": rel["target"]})
                        continue
                    target = rel["target"]
                    if target not in names:
                        issues.append({"kind": "missing_sheet_part", "name": name, "target": target})
                        continue
                    sheet_root = parse_xml(zf, target, issues)
                    if sheet_root is None:
                        continue
                    inspected_sheets.add(target)
                    dimension = sheet_root.find(q(MAIN, "dimension"))
                    entry["dimension"] = None if dimension is None else dimension.get("ref")
                    if layout:
                        inventory = layout_inventory(sheet_root, styles, shared_strings, min_font_pt)
                        report["layout_review"]["sheets"].append({"name": name, "state": entry["state"], **inventory})
                    formulas = list(sheet_root.iter(q(MAIN, "f")))
                    report["formula_count"] += len(formulas)
                    report["formula_error_tokens"].extend(
                        {"sheet": name, **item} for item in cell_error_tokens(sheet_root)
                    )

        # Parse all worksheet parts too, catching orphaned sheets and formula tokens.
        for ws_name in sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")):
            if ws_name in inspected_sheets:
                continue
            root = parse_xml(zf, ws_name, issues)
            if root is not None:
                report["formula_error_tokens"].extend({"sheet_part": ws_name, **item} for item in cell_error_tokens(root))

        if report["formula_error_tokens"]:
            warnings.append({"kind": "cached_formula_errors", "count": len(report["formula_error_tokens"]), "tokens": report["formula_error_tokens"][:100]})
        if report["package_parts"].get("external_links", 0):
            warnings.append({"kind": "external_links_present", "count": report["package_parts"]["external_links"]})
        if path.suffix.lower() == ".xlsm" and not report["package_parts"].get("vba_project"):
            warnings.append({"kind": "macro_extension_without_vba_project"})
        if path.suffix.lower() != ".xlsm" and report["package_parts"].get("vba_project"):
            warnings.append({"kind": "vba_project_in_non_xlsm_extension"})

    report["status"] = "passed" if not issues else "failed"
    if layout:
        entries = report["layout_review"]["sheets"]
        report["layout_review"]["status"] = ("incomplete" if issues or not entries else
            "review_required" if any(s["risks"] for s in entries) else "no_flagged_risks")
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--layout", action="store_true", help="Inventory static layout risks; does not prove visual readability")
    parser.add_argument("--min-font-pt", type=positive_number, help="Flag populated cells below this point size; implies --layout")
    args = parser.parse_args(argv)
    if not args.workbook.is_file():
        print(json.dumps({"status": "failed", "issues": [{"kind": "missing_file", "path": str(args.workbook)}]}, ensure_ascii=False, indent=2))
        return 2
    result = check(args.workbook, layout=args.layout, min_font_pt=args.min_font_pt)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    sys.exit(main())
