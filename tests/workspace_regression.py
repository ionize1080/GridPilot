"""Offline parser, grant boundary, pagination, analysis and embedded-file regressions."""
import base64
import hashlib
import io
import json
import os
from pathlib import Path
import sys
import unittest
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'bridge'))
import workspace_files as files
from workspace_analysis import analyze
from embedded_files import process, native_payload


def pdf(path, jpeg=None):
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>', b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>']
    objects.append(b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 400] /Resources << /Font << /F1 4 0 R >> /XObject << /Im 6 0 R >> >> /Contents 5 0 R >>')
    objects.append(b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    content = b'q 600 0 0 400 0 0 cm /Im Do Q' if jpeg else b'BT /F1 18 Tf 30 320 Td (Research revenue 120. Growth 20 percent.) Tj ET'
    objects.append(b'<< /Length ' + str(len(content)).encode() + b' >>\nstream\n' + content + b'\nendstream')
    if jpeg:
        objects.append(b'<< /Type /XObject /Subtype /Image /Width 600 /Height 400 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + str(len(jpeg)).encode() + b' >>\nstream\n' + jpeg + b'\nendstream')
    else:
        objects.append(b'null')
    output, offsets = bytearray(b'%PDF-1.4\n'), [0]
    for i, obj in enumerate(objects, 1):
        offsets.append(len(output)); output.extend(str(i).encode() + b' 0 obj\n' + obj + b'\nendobj\n')
    xref = len(output); output.extend(f'xref\n0 {len(offsets)}\n0000000000 65535 f \n'.encode())
    for value in offsets[1:]: output.extend(f'{value:010d} 00000 n \n'.encode())
    output.extend(f'trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF'.encode()); path.write_bytes(output)


def fixtures(folder):
    folder.mkdir(parents=True, exist_ok=True)
    (folder / '研究说明.md').write_text('# Research\nRevenue grew from 100 to 120.\nActual growth is 20%.\nEvidence marker: ORCHID-731.\n', encoding='utf-8')
    (folder / 'page.html').write_text('<h1>Market</h1><script>SHOULD_NOT_APPEAR</script><p>revenue &amp; profit 120</p>', encoding='utf-8')
    (folder / 'long.txt').write_text('first\n' + 'x' * 10000 + '\nlast', encoding='utf-8')
    (folder / '.env').write_text('hidden', encoding='utf-8')
    (folder / 'invalid.pdf').write_bytes(b'not a PDF')
    (folder / 'table.csv').write_text('Region,Year,Revenue,Cost\nEast,2025,100,60\nEast,2026,120,70\nWest,2025,80,50\nWest,2026,100,55\n', encoding='utf-8')
    from PIL import Image, ImageDraw, ImageFont
    im = Image.new('RGB', (600, 400), 'white'); d = ImageDraw.Draw(im)
    font = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 32)
    d.text((30, 30), 'ORCHID 731', font=font, fill='black'); d.rectangle((80, 140, 210, 330), fill='#D93838'); d.rectangle((320, 100, 450, 330), fill='#286CD5')
    d.text((80, 340), '40', font=font, fill='black'); d.text((320, 340), '60', font=font, fill='black')
    im.save(folder / 'chart.png'); out = io.BytesIO(); im.save(out, format='JPEG')
    pdf(folder / 'text.pdf'); pdf(folder / 'scan.pdf', out.getvalue())
    from docx import Document
    doc = Document(); doc.add_heading('Research heading', 1); doc.add_paragraph('Office evidence ORCHID-731, revenue 120.')
    table = doc.add_table(rows=2, cols=2); table.cell(0, 0).text = 'Revenue'; table.cell(0, 1).text = '120'; table.cell(1, 0).text = 'Profit'; table.cell(1, 1).text = '50'
    doc.add_picture(str(folder / 'chart.png')); doc.save(folder / 'report.docx')
    import xlsxwriter
    with xlsxwriter.Workbook(str(folder / 'revenue.xlsx')) as wb:
        ws = wb.add_worksheet('Revenue')
        for i, row in enumerate([['Region', 'Year', 'Revenue', 'Cost'], ['East', 2025, 100, 60], ['East', 2026, 120, 70], ['West', 2025, 80, 50], ['West', 2026, 100, 55]]): ws.write_row(i, 0, row)
        ws = wb.add_worksheet('Formulas'); ws.write_number('A1', 100); ws.write_number('A2', 120); ws.write_formula('A3', '=SUM(A1:A2)', None, 220)
    from pptx import Presentation
    from pptx.util import Inches
    pres = Presentation(); slide = pres.slides.add_slide(pres.slide_layouts[1]); slide.shapes.title.text = 'Research slide'; slide.placeholders[1].text = 'ORCHID-731 revenue 120'
    slide.notes_slide.notes_text_frame.text = 'Evidence note 20 percent'; slide.shapes.add_picture(str(folder / 'chart.png'), Inches(6), Inches(2), width=Inches(3)); pres.save(folder / 'slides.pptx')
    return folder


class WorkspaceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = ROOT / '.temp' / 'workspace-tests'; cls.base.mkdir(parents=True, exist_ok=True)
        cls.source = fixtures(cls.base / 'references')
        cls.store = cls.base / 'store'; cls.store.mkdir(exist_ok=True)
        os.environ['GRIDPILOT_WORKSPACES'] = str(cls.store)
        cls.wid, cls.sid = uuid.uuid4().hex, uuid.uuid4().hex
        cls.grant = cls.store / (cls.wid + '.json')
        cls.grant.write_text(json.dumps(dict(id=cls.wid, sources=[dict(id=cls.sid, kind='folder', path=str(cls.source))])), encoding='utf-8')
        (cls.base / 'session.json').write_text(json.dumps(dict(workspace_id=cls.wid, folder_id=cls.sid + ':', store=str(cls.store), references=str(cls.source))), encoding='utf-8')
    def fid(self, name):
        return self.sid + ':' + base64.urlsafe_b64encode(name.encode()).decode().rstrip('=')
    def read(self, name, **kw): return files.read_file(self.wid, self.fid(name), **kw)
    def test_01_listing_is_metadata_only(self):
        listed = files.list_files(self.wid, self.sid + ':', limit=3)
        self.assertEqual(len(listed['entries']), 3); self.assertIsNotNone(listed['next_cursor']); self.assertFalse(listed['contentIncluded'])
        self.assertNotIn('ORCHID', json.dumps(listed)); self.assertNotIn('.env', json.dumps(files.list_files(self.wid, self.sid + ':', limit=100)))
    def test_02_paths_and_grants(self):
        for path in ['../secret.txt', '/etc/passwd', 'C:/Windows/test.txt', '.env', '..\\secret']:
            with self.assertRaises(ValueError): files.read_file(self.wid, self.fid(path))
        with self.assertRaises(ValueError): files.read_file('0' * 32, self.fid('研究说明.md'))
        with self.assertRaises(ValueError): files.read_file(self.wid, '0' * 32 + ':')
    def test_03_text_html_pdf(self):
        self.assertIn('20%', str(self.read('研究说明.md', count=10)))
        self.assertNotIn('SHOULD_NOT_APPEAR', str(self.read('page.html')))
        self.assertIn('120', str(self.read('text.pdf')))
        self.assertIn('workspace_view', str(self.read('scan.pdf')))
        with self.assertRaises(Exception): self.read('invalid.pdf')
    def test_04_continuation_and_revisions(self):
        result = self.read('long.txt', start=2, max_chars=512)
        self.assertEqual(result['chars'], 512); self.assertEqual(result['next'], {'start': 2, 'offset': 512})
        second = self.read('long.txt', expected_revision=result['revision'], **result['next'], max_chars=512)
        self.assertEqual(second['units'][0]['offset'], 512)
        with self.assertRaises(ValueError): self.read('long.txt', expected_revision='stale')
    def test_05_search_range(self):
        first = files.search_file(self.wid, self.fid('研究说明.md'), 'ORCHID', scan_units=2)
        self.assertFalse(first['complete']); self.assertEqual(first['hits'], [])
        last = files.search_file(self.wid, self.fid('研究说明.md'), 'ORCHID', start=first['next_start'])
        self.assertTrue(last['complete']); self.assertEqual(last['hits'][0]['unit'], 4)
    def test_06_office_structure(self):
        result = self.read('report.docx', count=50)
        self.assertIn('style: Heading1', str(result)); self.assertIn('table 1 row 2', str(result)); self.assertIn('Profit', str(result))
        result = self.read('revenue.xlsx', count=50)
        self.assertIn('cached value: 220', str(result)); self.assertIn('C3=120', str(result))
        self.assertIn('Evidence note', str(self.read('slides.pptx')))
    def test_07_images_and_outline(self):
        for name in ['chart.png', 'scan.pdf', 'text.pdf']:
            value = files.view_file(self.wid, self.fid(name)); self.assertEqual(value['mimeType'], 'image/jpeg'); self.assertLessEqual(value['width'], 1568)
            self.assertGreater(len(base64.b64decode(value['data'])), 1000)
        outline = files.outline_file(self.wid, self.fid('report.docx'))
        self.assertEqual(outline['image_count'], 1)
        self.assertEqual(files.view_file(self.wid, self.fid('report.docx'), asset=1)['width'], 600)
    def test_08_batch_bounded(self):
        result = files.read_batch(self.wid, [{'file_id': self.fid('long.txt'), 'start': 2, 'max_chars': 1000}, {'file_id': self.fid('report.docx')}], max_chars=512)
        self.assertEqual(result['chars'], 512); self.assertTrue(result['results'][1]['skipped'])
        mixed=files.read_batch(self.wid,[{'file_id':self.fid('invalid.pdf')},{'file_id':self.fid('report.docx')}])
        self.assertIn('error',mixed['results'][0]);self.assertIn('Research heading',str(mixed['results'][1]))
    def test_09_size_limit(self):
        file = self.source / 'large.txt'
        with file.open('wb') as f: f.truncate(101 * 1024 * 1024)
        try:
            with self.assertRaises(ValueError): self.read('large.txt')
            (self.store / 'preferences.json').write_text('{"maxFileMB":102}')
            self.assertFalse(files.list_files(self.wid, self.sid + ':', limit=100)['entries'][[e['name'] for e in files.list_files(self.wid, self.sid + ':', limit=100)['entries']].index('large.txt')]['overLimit'])
        finally:
            file.unlink(); (self.store / 'preferences.json').unlink(missing_ok=True)
    def test_10_analysis_paper(self):
        source = self.source / 'revenue.xlsx'; before = hashlib.sha256(source.read_bytes()).hexdigest()
        result = analyze(self.wid, self.fid('revenue.xlsx'), calculations=[dict(name='Profit', expression='Revenue - Cost')], group_by=['Region'], pivot_column='Year', metrics=[dict(column='Profit', op='sum', name='Profit')])
        self.assertEqual(result['result_preview'], [['East', 40.0, 50.0], ['West', 30.0, 45.0]])
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)
        target, _ = files.resolve(self.wid, result['file_id']); self.assertTrue(target.exists())
        import openpyxl
        wb = openpyxl.load_workbook(target, data_only=True); self.assertEqual(wb['计算明细']['E3'].value, 50); self.assertEqual(wb['透视汇总']['C2'].value, 50); self.assertIn('说明与来源', wb.sheetnames); wb.close()
        wb = openpyxl.load_workbook(target, data_only=False); self.assertEqual(wb['计算明细']['E3'].value, '=(C3-D3)'); wb.close()
    def test_11_analysis_rejects_unsafe_or_stale_calcs(self):
        for expression in ['__import__("os").system("x")', 'Revenue ** 1000', 'Revenue / 0']:
            with self.assertRaises((ValueError, SyntaxError)): analyze(self.wid, self.fid('revenue.xlsx'), calculations=[dict(name='Bad', expression=expression)])
        result = analyze(self.wid, self.fid('revenue.xlsx'), sheet='Formulas')
        self.assertEqual(result['result_preview'], [[120], [220]])
        self.assertEqual(result['source']['recalculatedFormulas'], 1)
        import openpyxl
        target, _ = files.resolve(self.wid, result['file_id'])
        wb = openpyxl.load_workbook(target, data_only=False)
        self.assertEqual(wb['源公式复核']['B2'].value, '=SUM(A1:A2)')
        self.assertEqual(wb['源公式复核']['B2'].data_type, 's'); wb.close()
        with self.assertRaises(ValueError): analyze(self.wid, columns=['A'], rows=[[None]], metrics=[dict(column='A', op='sum')])
    def test_12_new_paper_and_injection(self):
        result = analyze(self.wid, columns=['Label', 'A'], rows=[['=HYPERLINK("https://invalid")', 12]], calculations=[dict(name='B', expression='round(A / 5, 2)')])
        target, _ = files.resolve(self.wid, result['file_id'])
        import openpyxl
        wb = openpyxl.load_workbook(target, data_only=False); self.assertEqual(wb['源数据']['A2'].data_type, 's'); wb.close()
        wb = openpyxl.load_workbook(target, data_only=True); self.assertEqual(wb['计算明细']['C2'].value, 2.4); wb.close()
        repeated=analyze(self.wid,result['file_id'],sheet='源数据')
        self.assertEqual(repeated['source']['recalculatedFormulas'],0)
        self.assertEqual(repeated['result_preview'][0][0],'=HYPERLINK("https://invalid")')
    def test_13_embedded_pdf_and_external_link(self):
        data = base64.b64encode((self.source / 'text.pdf').read_bytes()).decode()
        package = self.base / 'snapshot.xml'
        package.write_text(f'<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/embeddings/test.pdf"><pkg:binaryData>{data}</pkg:binaryData></pkg:part><pkg:part pkg:name="/_rels/.rels"><pkg:xmlData><Relationships><Relationship Type="test/oleObject" TargetMode="External" Target="https://invalid.local/private"/></Relationships></pkg:xmlData></pkg:part></pkg:package>', encoding='utf-8')
        listed = process(str(package)); self.assertEqual(listed['parts'][0]['kind'], 'embedded_document'); self.assertFalse(listed['links'][0]['followed'])
        self.assertIn('120', str(process(str(package), 'read')))
        self.assertEqual(process(str(package), 'view')['mimeType'], 'image/jpeg')
    def test_14_zip_bomb(self):
        target = self.source / 'bomb.docx'
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z: z.writestr('word/document.xml', b'x' * 4_000_000)
        try:
            with self.assertRaises(ValueError): self.read('bomb.docx')
        finally: target.unlink()

    def test_15_formula_subset(self):
        from spreadsheet_calc import FormulaGrid
        from decimal import Decimal
        grid = FormulaGrid([[10, 20, '=SUM(A1:B1)'], [2, 0, '=IF(B2=0,ROUND(C1/A2,2),1/0)']])
        self.assertEqual(grid.cell('C1'), 30); self.assertEqual(grid.cell('C2'), 15)
        self.assertEqual(grid.expression('SUM(B2:A1)'),32)
        for expression, result in [('SUM(A1:B2)',32),('COUNT(A1:B2)',4),('AVERAGE(A1:B2)',8),('MIN(A1:B2)',0),('MAX(A1:B2)',20),('ROUND(-2.345,2)',Decimal('-2.35')),('ABS(-3)+10%',Decimal('3.1')),('AND(A1>0,B2=0)',True)]:
            self.assertEqual(grid.expression(expression), result, expression)
        for expression in ['VLOOKUP(A1,A1:B2,2)', 'Other!A1', '[Book]A1', '__import__("os")', '1/0','-2^2', 'SUM("4")', 'SUM(TRUE)','XFE1']:
            with self.assertRaises(ValueError, msg=expression): grid.expression(expression)
        with self.assertRaises(ValueError): FormulaGrid([['=B1','=A1']]).cell('A1')

    def test_16_recursive_budgets_and_virtual_listing(self):
        sub = self.source / 'depth-a' / 'depth-b' / 'depth-c'
        sub.mkdir(parents=True, exist_ok=True)
        (sub / 'deep-marker.md').write_text('no automatic read')
        names, cursor, limited = [], '', False
        for i in range(200):
            result=files.find_files(self.wid, max_depth=1, limit=2, cursor=cursor)
            names.extend(e['name'] for e in result['entries']); limited |= result['depth_or_queue_limited']
            self.assertLessEqual(result['scanned'],2000)
            cursor=result['next_cursor']
            if not cursor: break
        else: self.fail('recursive pagination did not complete')
        self.assertTrue(limited); self.assertNotIn('deep-marker.md',names)
        found=files.find_files(self.wid, query='deep-marker', max_depth=3)
        self.assertEqual([e['name'] for e in found['entries']],['deep-marker.md'])
        self.assertNotIn('no automatic read',json.dumps(found))
        grant=json.loads(self.grant.read_text())
        virtual=uuid.uuid4().hex
        grant['collections']=[dict(id=virtual,name='资料',parent='',order=1)]
        grant['sources'][0]['parent']=virtual
        self.grant.write_text(json.dumps(grant))
        try:
            self.assertEqual(files.list_files(self.wid,'virtual:'+virtual)['entries'][0]['id'],self.sid+':')
            self.assertEqual(files.find_files(self.wid,query='deep-marker',max_depth=3)['entries'][0]['name'],'deep-marker.md')
        finally:
            grant['sources'][0].pop('parent');grant.pop('collections');self.grant.write_text(json.dumps(grant))

    def test_17_legacy_office_fixtures_when_available(self):
        for name, marker in [('legacy.doc','Office evidence'),('legacy.ppt','ORCHID-731')]:
            if (self.source/name).exists(): self.assertIn(marker,str(self.read(name,count=50)))


if __name__ == '__main__':
    unittest.main(verbosity=2)
