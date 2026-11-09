from pathlib import Path
import contextlib
import io
import json
import tempfile
import unittest
import zipfile
import solver

class SolverTests(unittest.TestCase):
    def test_batch_and_xml(self):
        with tempfile.TemporaryDirectory(dir=solver.ROOT, prefix='test-') as folder:
            root = Path(folder).resolve()
            self.assertEqual(root.parent, solver.ROOT.resolve())
            incoming = root / 'input'
            incoming.mkdir()
            source = incoming / 'sample.xlsx'
            sheet = b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><sheetProtection password="ABCD" sheet="1"/></worksheet>'
            with zipfile.ZipFile(source, 'w') as z:
                z.writestr('xl/workbook.xml', b'<workbook><workbookProtection lockStructure="1" workbookPassword="ABCD"/></workbook>')
                z.writestr('xl/worksheets/sheet1.xml', sheet)
                z.writestr('xl/other.xml', b'unchanged')
            before = source.read_bytes()
            (incoming / 'broken.xlsm').write_bytes(b'broken')
            original_root = solver.ROOT
            try:
                solver.ROOT = root
                with contextlib.redirect_stdout(io.StringIO()):
                    self.assertEqual(solver.main(), 1)
                    self.assertEqual(solver.main(), 1)
            finally:
                solver.ROOT = original_root
            self.assertEqual(source.read_bytes(), before)
            self.assertTrue((root / 'output' / 'sample (2).xlsx').exists())
            self.assertFalse((root / 'output' / 'broken.xlsm').exists())
            self.assertFalse(list((root / 'output').glob('*.tmp')))
            with zipfile.ZipFile(root / 'output' / 'sample.xlsx') as z:
                self.assertNotIn(b'sheetProtection', z.read('xl/worksheets/sheet1.xml'))
                self.assertNotIn(b'workbookProtection', z.read('xl/workbook.xml'))
                self.assertEqual(z.read('xl/other.xml'), b'unchanged')
            report = json.loads(next((root / 'output').glob('*.json')).read_text(encoding='utf-8'))
            self.assertEqual({r['status'] for r in report}, {'success', 'error'})

    def test_real_xlsb_and_second_pass(self):
        samples = sorted((solver.ROOT.parent / 'bbsk rapor').glob('*.xlsb'))
        if not samples:
            self.skipTest('No local XLSB sample is available')
        source = samples[0]
        before = solver.digest(source)
        with tempfile.TemporaryDirectory(dir=solver.ROOT, prefix='test-') as folder:
            output = Path(folder).resolve()
            self.assertEqual(output.parent, solver.ROOT.resolve())
            result = solver.process(source, output)
            self.assertTrue(result['vba_processed'])
            self.assertGreater(result['sheet_protection_records'], 0)
            self.assertGreater(result['workbook_protection_records'], 0)
            unlocked = output / result['output']
            result2 = solver.process(unlocked, output)
            self.assertEqual(result2['sheet_protection_records'], 0)
            self.assertEqual(result2['workbook_protection_records'], 0)
            self.assertEqual(result2['modified_parts'], [])
        self.assertEqual(before, solver.digest(source))

if __name__ == '__main__':
    unittest.main()
