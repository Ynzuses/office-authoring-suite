"""Synthetic OOXML diagnostics, not rendering or native Excel acceptance."""
import hashlib
import tempfile
import unittest
import zipfile
from pathlib import Path

from test_excel import check_xlsx, CONTENT_TYPES, ROOT_RELS, WORKBOOK, WORKBOOK_RELS

MAIN = check_xlsx.MAIN
STYLES = f'''<styleSheet xmlns="{MAIN}">
<fonts count="2"><font><sz val="12"/></font><font><sz val="8"/></font></fonts>
<cellStyleXfs count="1"><xf fontId="0"><alignment shrinkToFit="1"/></xf></cellStyleXfs>
<cellXfs count="4">
<xf fontId="0"><alignment wrapText="1"/></xf>
<xf fontId="1"><alignment wrapText="1"/></xf>
<xf fontId="0"><alignment shrinkToFit="1"/></xf>
<xf xfId="0" fontId="1" applyFont="0" applyAlignment="0"/>
</cellXfs></styleSheet>'''


def fixture(path, data, *, panes="", merges="", extras="", styles=STYLES, shared=None):
    sheet = f'<worksheet xmlns="{MAIN}">{panes}{extras}<sheetData>{data}</sheetData>{merges}</worksheet>'
    parts = {"[Content_Types].xml": CONTENT_TYPES, "_rels/.rels": ROOT_RELS,
             "xl/workbook.xml": WORKBOOK, "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
             "xl/worksheets/sheet1.xml": sheet}
    if styles is not None:
        parts["xl/styles.xml"] = styles
    if shared is not None:
        parts["xl/sharedStrings.xml"] = shared
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            z.writestr(name, text)


TEXT = '<row r="1"><c r="A1" t="inlineStr"><is><t>Plan title</t></is></c></row>'


class ExcelLayoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "layout.xlsx"

    def inspect(self, **kwargs):
        before = self.path.read_bytes()
        report = check_xlsx.check(self.path, layout=True, **kwargs)
        self.assertEqual(before, self.path.read_bytes())
        self.assertEqual(hashlib.sha256(before).hexdigest(), report['sha256'])
        self.assertEqual(report['layout_review']['visual_acceptance'], 'not_assessed')
        return report['layout_review']['sheets'][0]

    def test_freeze_crossing_merged_title_is_found_but_row_only_is_clear(self):
        merge = '<mergeCells><mergeCell ref="A1:F1"/></mergeCells>'
        for x, expected in [('2', True), ('0', False)]:
            pane = f'<sheetViews><sheetView workbookViewId="0"><pane state="frozenSplit" xSplit="{x}" ySplit="4"/></sheetView></sheetViews>'
            fixture(self.path, TEXT, panes=pane, merges=merge)
            risks = self.inspect()['risks']
            self.assertEqual(any(r['kind'] == 'freeze_divider_crosses_merge' for r in risks), expected)
            if expected:
                self.assertIn({'kind': 'freeze_divider_crosses_merge', 'axis': 'vertical',
                               'range': 'A1:F1', 'split_after': 2.0}, risks)

    def test_horizontal_crossing_and_boundary_touch_are_distinct(self):
        pane = '<sheetViews><sheetView><pane state="frozen" ySplit="2"/></sheetView></sheetViews>'
        merges = '<mergeCells><mergeCell ref="A1:A2"/><mergeCell ref="B2:B4"/></mergeCells>'
        fixture(self.path, TEXT, panes=pane, merges=merges)
        crossings = [r['range'] for r in self.inspect()['risks'] if r['kind'] == 'freeze_divider_crosses_merge']
        self.assertEqual(crossings, ['B2:B4'])

    def test_split_panes_use_twips_not_cell_counts(self):
        pane = '<sheetViews><sheetView><pane state="split" xSplit="1500"/></sheetView></sheetViews>'
        fixture(self.path, TEXT, panes=pane)
        self.assertEqual(self.inspect()['risks'][0]['kind'], 'split_pane_requires_native_review')

    def test_font_threshold_ignores_empty_style_only_cells_and_finds_shrink(self):
        fixture(self.path, '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"/><c r="C1" s="2"><v>2</v></c><c r="D1" s="1"><v/></c></row>')
        inventory = self.inspect(min_font_pt=10)
        small = next(r for r in inventory['risks'] if r['kind'] == 'small_cell_font')
        self.assertEqual(small['samples'], [{'cell': 'A1', 'min_font_pt': 8.0}])
        self.assertEqual(inventory['populated_cells'], 2)
        self.assertIn('shrink_to_fit', [r['kind'] for r in inventory['risks']])
        self.assertNotIn('small_cell_font', [r['kind'] for r in self.inspect(min_font_pt=8)['risks']])

    def test_row_column_and_named_style_inheritance(self):
        data = '<row r="1" s="1" customFormat="1"><c r="A1"><v>1</v></c><c r="B1" s="0"><v>2</v></c></row><row r="2"><c r="C2"><v>3</v></c><c r="D2" s="3"><v>4</v></c></row>'
        fixture(self.path, data, extras='<cols><col min="3" max="3" style="1"/></cols>')
        risks = self.inspect(min_font_pt=10)['risks']
        small = next(r for r in risks if r['kind'] == 'small_cell_font')
        self.assertEqual([r['cell'] for r in small['samples']], ['A1', 'C2'])
        self.assertEqual(next(r for r in risks if r['kind'] == 'shrink_to_fit')['samples'], ['D2'])

    def test_rich_inline_and_shared_string_fonts_are_inspected(self):
        rich = '<r><rPr><sz val="7"/></rPr><t>small label</t></r><r><t>normal label</t></r>'
        data = f'<row r="1"><c r="A1" t="inlineStr"><is>{rich}</is></c><c r="B1" t="s"><v>0</v></c></row>'
        fixture(self.path, data, shared=f'<sst xmlns="{MAIN}"><si>{rich}</si></sst>')
        small = next(r for r in self.inspect(min_font_pt=10)['risks'] if r['kind'] == 'small_cell_font')
        self.assertEqual(small['samples'], [{'cell': 'A1', 'min_font_pt': 7.0}, {'cell': 'B1', 'min_font_pt': 7.0}])

    def test_unknown_or_invalid_style_remains_unassessed(self):
        for styles, data in [(None, TEXT), (STYLES, TEXT.replace('r="A1"', 'r="A1" s="-1"'))]:
            fixture(self.path, data, styles=styles)
            inventory = self.inspect(min_font_pt=10)
            self.assertEqual(inventory['font_size_range_pt'], None)
            self.assertEqual(inventory['risks'][0]['kind'], 'unresolved_cell_style')

    def test_fit_to_page_is_reviewed_without_claiming_effective_font_size(self):
        fixture(self.path, TEXT, extras='<pageSetup fitToWidth="1" fitToHeight="1"/>')
        inventory = self.inspect()
        self.assertEqual(inventory['font_size_range_pt'], [12.0, 12.0])
        self.assertEqual(inventory['risks'][0]['kind'], 'print_scaling_requires_target_size_review')

    def test_layout_opt_in_and_invalid_thresholds(self):
        fixture(self.path, TEXT)
        self.assertNotIn('layout_review', check_xlsx.check(self.path))
        self.assertEqual(check_xlsx.check(self.path, min_font_pt=12)['layout_review']['status'], 'no_flagged_risks')
        for bad in [0, -1, float('nan'), float('inf')]:
            with self.assertRaises(ValueError):
                check_xlsx.check(self.path, min_font_pt=bad)


if __name__ == '__main__':
    unittest.main()
