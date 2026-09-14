import base64
import io
import unittest
import zipfile
from xml.dom import minidom
from app import mapping_rules, mask_xlsx, process

def fixture(sheet='Sheet1', extra=None):
    parts = {
        'xl/workbook.xml': f'<workbook><sheets><sheet name="{sheet}"/></sheets></workbook>',
        'xl/sharedStrings.xml': '<sst><si><r><t>玉</t></r><r><t>崎</t></r></si><si><t>玉崎太郎・玉崎・高橋</t></si></sst>',
        'xl/worksheets/sheet1.xml': '<worksheet><c><v>42</v></c><c><f>SUM(A1:A2)</f><v>42</v></c><c><f>IF(A1=1,"高橋","")</f></c></worksheet>',
        'xl/worksheets/sheet2.xml': '<worksheet><is><t>高橋</t></is></worksheet>',
        'xl/styles.xml': '<styleSheet><fonts count="1"/></styleSheet>',
    }
    parts.update(extra or {})
    b = io.BytesIO()
    with zipfile.ZipFile(b, 'w') as z:
        for n, s in parts.items():
            z.writestr(n, s)
    return b.getvalue()

class Tests(unittest.TestCase):
    def setUp(self):
        self.m = {'玉崎': 'A社のメンバー1', '玉崎太郎': 'A社のメンバー3', '高橋': 'B社のメンバー1'}
    def test_longest_rich_multisheet_formula_and_preservation(self):
        original = fixture()
        result, counts, warnings = mask_xlsx(original, self.m)
        self.assertEqual(counts, {'玉崎': 2, '玉崎太郎': 1, '高橋': 3})
        with zipfile.ZipFile(io.BytesIO(result)) as z, zipfile.ZipFile(io.BytesIO(original)) as src:
            self.assertEqual(z.read('xl/styles.xml'), src.read('xl/styles.xml'))
            self.assertIn(b'SUM(A1:A2)', z.read('xl/worksheets/sheet1.xml'))
            self.assertIn(b'<v>42</v>', z.read('xl/worksheets/sheet1.xml'))
            for n in z.namelist():
                self.assertNotIn('玉崎', z.read(n).decode())
                self.assertNotIn('高橋', z.read(n).decode())
    def test_sheet_name_blocks(self):
        with self.assertRaises(ValueError):
            mask_xlsx(fixture(sheet='玉崎'), self.m)
    def test_formula_reference_blocks(self):
        with self.assertRaises(ValueError):
            mask_xlsx(fixture(extra={'xl/worksheets/sheet3.xml': '<worksheet><f>玉崎!A1</f></worksheet>'}), self.m)
    def test_duplicate_and_alias_contamination(self):
        for rows in [[{'name':'a','alias':'b'},{'name':'a','alias':'c'}],[{'name':'a','alias':'aa'}]]:
            with self.assertRaises(ValueError):
                mapping_rules(rows)
    def test_batch_partial_failure_and_neutral_names(self):
        p = process({'rows':[{'name':n,'alias':a} for n,a in self.m.items()], 'files':[{'name':'玉崎.xlsx','data':base64.b64encode(fixture()).decode()},{'name':'bad.xlsx','data':'!!!'}]})
        self.assertEqual([r['status'] for r in p['reports']], ['ok','error'])
        with zipfile.ZipFile(io.BytesIO(base64.b64decode(p['zip']))) as z:
            self.assertEqual(z.namelist(), ['masked_001.xlsx'])
    def test_embedded_blocks_and_image_warns(self):
        with self.assertRaises(ValueError):
            mask_xlsx(fixture(extra={'xl/embeddings/a.bin': 'x'}), self.m)
        self.assertTrue(mask_xlsx(fixture(extra={'xl/media/a.png': 'x'}), self.m)[2])
    def test_formula_escaping(self):
        data = fixture(extra={'xl/worksheets/sheet3.xml':'<worksheet><f>"高橋"</f></worksheet>'})
        result, _, _ = mask_xlsx(data, {'高橋':'B"1'})
        with zipfile.ZipFile(io.BytesIO(result)) as z:
            doc = minidom.parseString(z.read('xl/worksheets/sheet3.xml'))
            self.assertEqual('"B""1"', doc.getElementsByTagName('f')[0].firstChild.data)

if __name__ == '__main__':
    unittest.main()
