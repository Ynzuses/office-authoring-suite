import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("check_xlsx", ROOT / "excel-authoring/scripts/check_xlsx.py")
check_xlsx = importlib.util.module_from_spec(spec)
spec.loader.exec_module(check_xlsx)


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
ROOT_RELS = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
WORKBOOK_RELS = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'''
WORKBOOK = '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Total">Summary!$B$2</definedName></definedNames></workbook>'''
SHEET = '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c><c r="B1" t="inlineStr"><is><t>Total</t></is></c></row><row r="2"><c r="A2"><v>4</v></c><c r="B2"><f>SUM(A2:A2)</f><v>4</v></c></row></sheetData></worksheet>'''


def write_fixture(path: Path, bad=False, macro=False, sheet_target="worksheets/sheet1.xml"):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in {
            "[Content_Types].xml": CONTENT_TYPES,
            "_rels/.rels": ROOT_RELS,
            "xl/workbook.xml": WORKBOOK if not bad else WORKBOOK.replace('r:id="rId1"', 'r:id="missing"'),
            "xl/_rels/workbook.xml.rels": WORKBOOK_RELS.replace("worksheets/sheet1.xml", sheet_target),
            "xl/worksheets/sheet1.xml": SHEET if not bad else SHEET.replace("<v>4</v>", "<v>#REF!</v>"),
        }.items():
            z.writestr(name, data)
        if macro:
            z.writestr("xl/vbaProject.bin", b"synthetic macro part")


class ExcelCheckerTests(unittest.TestCase):
    def test_package_absolute_sheet_target_is_resolved_without_rewriting_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "absolute.xlsx"
            write_fixture(path, sheet_target="/xl/worksheets/sheet1.xml")
            before = path.read_bytes()
            result = check_xlsx.check(path)
            self.assertEqual(result["status"], "passed")
            self.assertEqual(result["formula_count"], 1)
            self.assertEqual(before, path.read_bytes())

    def test_missing_absolute_sheet_target_still_fails(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "missing.xlsx"
            write_fixture(path, sheet_target="/xl/worksheets/absent.xml")
            result = check_xlsx.check(path)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any(i.get("target") == "xl/worksheets/absent.xml" for i in result["issues"]))

    def test_valid_package_reports_structure_without_modifying_bytes(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "book.xlsx"
            write_fixture(path)
            before = path.read_bytes()
            result = check_xlsx.check(path)
            self.assertEqual(result["status"], "passed")
            self.assertEqual(result["sheets"][0]["name"], "Summary")
            self.assertEqual(result["formula_count"], 1)
            self.assertEqual(result["package_parts"]["defined_names"], 1)
            self.assertEqual(before, path.read_bytes())

    def test_broken_relationship_fails_and_cached_formula_error_warns(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "bad.xlsx"
            write_fixture(path, bad=True)
            result = check_xlsx.check(path)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any(i["kind"] == "unresolved_sheet_relationship" for i in result["issues"]))
            self.assertTrue(any(w["kind"] == "cached_formula_errors" for w in result["warnings"]))

    def test_xlsm_macro_signal_is_reported(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "macro.xlsm"
            write_fixture(path, macro=True)
            result = check_xlsx.check(path)
            self.assertTrue(result["package_parts"]["vba_project"])
            self.assertEqual(result["status"], "passed")

    def test_multiple_error_sheets_are_scanned_once_without_key_error(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "errors.xlsx"
            write_fixture(path)
            with zipfile.ZipFile(path, "a") as z:
                z.writestr("xl/worksheets/orphan2.xml", SHEET.replace("<v>4</v>", "<v>#REF!</v>"))
                z.writestr("xl/worksheets/orphan3.xml", SHEET.replace("<v>4</v>", "<v>#VALUE!</v>"))
            report = check_xlsx.check(path)
            self.assertEqual(len(report['formula_error_tokens']), 4)
            self.assertEqual({e['sheet_part'] for e in report['formula_error_tokens']},
                             {'xl/worksheets/orphan2.xml', 'xl/worksheets/orphan3.xml'})


if __name__ == "__main__":
    unittest.main()
