"""Regression tests for namespace references that ordinary XML parsing misses."""

import importlib.util
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile


TOOL = Path(__file__).resolve().parents[2] / "word-authoring/scripts/check_docx_namespaces.py"
spec = importlib.util.spec_from_file_location("docx_namespace_checker", TOOL)
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = checker.MC


class NamespaceChecks(unittest.TestCase):
    def test_parser_success_does_not_hide_missing_ignorable_prefix(self):
        xml = f'<w:document xmlns:w="{W}" xmlns:m="{MC}" m:Ignorable="ext"/>'.encode()
        ET.fromstring(xml)  # A well-formed XML document can still be incompatible.
        self.assertEqual(checker.inspect_xml(xml)[0]["undefined_prefix"], "ext")

    def test_local_declarations_apply_to_children_but_not_siblings(self):
        xml = f'''<w:document xmlns:w="{W}" xmlns:m="{MC}">
          <w:p xmlns:ext="urn:test" m:Ignorable="ext"><w:r m:Ignorable="ext"/></w:p>
          <w:p m:Ignorable="ext"/>
        </w:document>'''.encode()
        self.assertEqual(len(checker.inspect_xml(xml)), 1)

    def test_choice_requires_uses_namespace_identity_not_literal_mc_prefix(self):
        xml = f'''<w:document xmlns:w="{W}" xmlns:alias="{MC}">
          <alias:AlternateContent><alias:Choice Requires="future"/></alias:AlternateContent>
        </w:document>'''.encode()
        issue = checker.inspect_xml(xml)[0]
        self.assertEqual((issue["attribute"], issue["undefined_prefix"]), ("Requires", "future"))

    def test_qname_lists_allow_bound_wildcards_and_builtin_xml_prefix(self):
        good = f'''<w:document xmlns:w="{W}" xmlns:m="{MC}" xmlns:ext="urn:test"
          m:PreserveElements="ext:*" m:PreserveAttributes="xml:space"/>'''.encode()
        self.assertEqual(checker.inspect_xml(good), [])
        bad = good.replace(b'ext:*', b'missing:item')
        self.assertEqual(checker.inspect_xml(bad)[0]["undefined_prefix"], "missing")

    def test_style_parts_are_checked_and_input_bytes_are_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / "example.docx"
            with zipfile.ZipFile(p, "w") as z:
                z.writestr("word/document.xml", f'<w:document xmlns:w="{W}"/>')
                z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W}" xmlns:m="{MC}" m:Ignorable="lost"/>')
            before = p.read_bytes()
            result = checker.check(p)
            self.assertEqual(result["status"], "failed")
            self.assertEqual(result["issues"][0]["part"], "word/styles.xml")
            self.assertEqual(p.read_bytes(), before)

    def test_malformed_xml_and_non_zip_inputs_are_unreadable(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / "bad.docx"
            p.write_bytes(b"not a zip")
            self.assertEqual(checker.check(p)["status"], "unreadable")
            with zipfile.ZipFile(p, "w") as z:
                z.writestr("word/document.xml", "<broken>")
            self.assertEqual(checker.check(p)["status"], "unreadable")

    def test_zip_without_a_document_is_not_a_pass(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / "empty.docx"
            with zipfile.ZipFile(p, "w"):
                pass
            self.assertEqual(checker.check(p)["status"], "unreadable")


if __name__ == "__main__":
    unittest.main()
