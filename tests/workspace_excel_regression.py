"""Real hidden Excel integration, defaulting to the Excel 2010 feature subset."""
import hashlib
import json
import os
from pathlib import Path
import sys
import unittest
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'bridge'))
import workspace_excel as engine
import workspace_files as files


class ExcelRegression(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = ROOT/'.temp'/('workspace-excel-'+uuid.uuid4().hex)
        cls.folder.mkdir(parents=True)
        cls.wid = uuid.uuid4().hex
        os.environ['GRIDPILOT_WORKSPACES'] = str(cls.folder)
        (cls.folder/(cls.wid+'.json')).write_text(json.dumps(dict(id=cls.wid,sources=[])),encoding='utf-8')
        cls.caps = engine.info(cls.wid)
        assert cls.caps['visible'] is False
        assert float(cls.caps['version']) >= 14
        print('EXCEL',json.dumps(cls.caps),flush=True)

    def test_01_native_workbook(self):
        actions = [
            dict(op='rename_sheet',sheet='Sheet1',name='Sales'),
            dict(op='write_range',sheet='Sales',values=[['Region','Year','Revenue','Cost','Profit'],['East',2025,100,60,None],['East',2026,120,70,None],['West',2025,80,50,None],['West',2026,105,60,None]]),
            dict(op='formula',sheet='Sales',range='E2:E5',formula='=C2-D2'),
            dict(op='table',sheet='Sales',range='A1:E5',name='SalesData'),
            dict(op='pivot',source_sheet='Sales',source_range='A1:E5',target_sheet='Pivot',name='ProfitPivot',rows=['Region'],columns=['Year'],values=[dict(field='Profit',aggregate='sum',name='Total Profit')]),
            dict(op='add_sheet',name='Notes'),
            dict(op='write_range',sheet='Notes',values=[['Text'],['=1+1'],['001']]),
            dict(op='copy_sheet',sheet='Notes',name='Notes copy'),
            dict(op='delete_sheet',sheet='Notes copy'),
            dict(op='describe',sheet='Sales',range='A1:E5',columns=['Profit'],target_sheet='Stats'),
            dict(op='correlation',sheet='Sales',range='A1:E5',columns=['Revenue','Cost'],target_sheet='Correlation'),
            dict(op='transform',sheet='Sales',range='A1:E5',target_sheet='Filtered',filters=[dict(column='Profit',op='ge',value=40)],sort_by=[dict(column='Profit',descending=True)],select=['Region','Year','Profit']),
            dict(op='chart',sheet='Sales',range='A1:E5',target_sheet='Charts',title='Profit'),
            dict(op='format',sheet='Sales',range='C2:E5',number_format='#,##0.00',autofit=True),
        ]
        result = engine.edit(self.wid,actions,output_name='NativePivot')
        type(self).created = result
        report = engine.read(self.wid,result['file_id'],'Sales','A1:E5')
        self.assertEqual([r[4] for r in report['values'][1:]],[40,50,30,45])
        self.assertEqual(report['formulas'][1]['formula'],'=C3-D3')
        pivot = engine.read(self.wid,result['file_id'],'Pivot','A3:D7')
        self.assertEqual(pivot['values'][1][:3],['Region',2025,2026])
        self.assertEqual(pivot['values'][2:4],[['East',40,50,90],['West',30,45,75]])
        self.assertEqual(pivot['values'][4][1:],[70,95,165])
        notes = engine.read(self.wid,result['file_id'],'Notes','A1:A3')
        self.assertEqual(notes['values'],[['Text'],['=1+1'],['001']])
        self.assertEqual(notes['formulas'],[])
        filtered = engine.read(self.wid,result['file_id'],'Filtered','A1:C4')['values']
        self.assertEqual(filtered,[['Region','Year','Profit'],['East',2026,50],['West',2026,45],['East',2025,40]])
        stats = engine.read(self.wid,result['file_id'],'Stats','A1:K2')['values'][1]
        self.assertEqual(stats[:8],['Profit',4,0,165,41.25,42.5,30,50])
        self.assertEqual([s['charts'] for s in result['sheets'] if s['name']=='Charts'],[1])
        path,_ = files.validated(self.wid,result['file_id'])
        with zipfile.ZipFile(path) as z:
            self.assertTrue(any(n.startswith('xl/pivotTables/pivotTable') for n in z.namelist()))
            self.assertTrue(any(n.startswith('xl/pivotCache/pivotCacheRecords') for n in z.namelist()))
            self.assertFalse(any(n.startswith('xl/externalLinks/') for n in z.namelist()))
        (self.folder/'created.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')

    def test_02_refresh_and_reopen(self):
        created = self.created
        path,stamp = files.validated(self.wid,created['file_id'])
        before = hashlib.sha256(path.read_bytes()).hexdigest()
        updated = engine.edit(self.wid,[dict(op='write_range',sheet='Sales',cell='C2',values=[[150]]),dict(op='refresh_pivot',sheet='Pivot',name='ProfitPivot')],file_id=created['file_id'],expected_revision=stamp)
        rows = engine.read(self.wid,updated['file_id'],'Pivot','A3:D7')['values']
        self.assertEqual(rows[2],['East',90,50,140])
        self.assertEqual(rows[-1][1:],[120,95,215])
        self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(),before)
        self.assertNotEqual(updated['file_id'],created['file_id'])
        self.assertEqual(len([s for s in updated['sheets'] if s['name'].startswith('GridPilot_Audit')]),2)

    def test_03_compatible_transforms(self):
        result = engine.edit(self.wid,[
            dict(op='write_range',sheet='Sheet1',values=[['ID','Region','Amount'],['A',' East ','10'],['A',' East ','10'],['B','West',None],['C','East','30']]),
            dict(op='add_sheet',name='Lookup'),dict(op='write_range',sheet='Lookup',values=[['Code','Owner'],['a','Alice'],['B','Bob']]),
            dict(op='transform',sheet='Sheet1',range='A1:C5',trim=['Region'],numeric=['Amount'],fill=[dict(column='Amount',value=0)],lookup=dict(sheet='Lookup',range='A1:B3',left_key='ID',right_key='Code',columns=['Owner'],if_not_found='Unknown'),unique_by=['ID'],sort_by=[dict(column='Amount',descending=True)],target_sheet='Joined'),
            dict(op='copy_sheet',sheet='Joined',name='NativeOps'),dict(op='sort',sheet='NativeOps',range='A1:D4',sort_by=[dict(column='Amount')]),
            dict(op='filter',sheet='NativeOps',range='A1:D4',filters=[dict(column='Amount',op='gt',value=0)]),
            dict(op='deduplicate',sheet='Sheet1',range='A1:C5',columns=['ID']),
        ])
        rows = engine.read(self.wid,result['file_id'],'Joined','A1:D4')['values']
        self.assertEqual(rows,[['ID','Region','Amount','Owner'],['C','East',30,'Unknown'],['A','East',10,'Alice'],['B','West',0,'Bob']])
        rows = engine.read(self.wid,result['file_id'],'NativeOps','A1:D4')['values']
        self.assertEqual([r[2] for r in rows[1:]],[0,10,30])

    def test_04_validation_and_atomic_failure(self):
        output = self.folder/'results'/self.wid
        before = set(output.iterdir())
        with self.assertRaisesRegex(ValueError,'本批次未保存'):
            engine.edit(self.wid,[dict(op='add_sheet',name='Unsaved'),dict(op='formula',sheet='Unsaved',range='A1',formula='=1/0')])
        self.assertEqual(set(output.iterdir()),before)
        with self.assertRaises(ValueError): engine.validate_formula('=XLOOKUP(A1,B1:B3,C1:C3)')
        for bad in ['=WEBSERVICE("https://example.com")','=CALL("x")',"='[outside.xlsx]Data'!A1"]:
            with self.assertRaises(ValueError): engine.validate_formula(bad,'installed',engine.MODERN_PROBES)
        for bad in ['A:A','A1:XFD1048576','B2:A1',"'Other'!A1"]:
            with self.assertRaises(ValueError): engine.bounds(bad)
        with self.assertRaises(ValueError): engine.sheet_name('../bad')
        with self.assertRaises(ValueError): engine.read(self.wid,self.created['file_id'],expected_revision='stale')
        with self.assertRaisesRegex(ValueError,'工作表不存在'):
            engine.edit(self.wid,[dict(op='write_range',sheet='Missing',values=[[1]])])

    def test_05_modern_formulas_only_when_supported(self):
        if 'SORT' not in self.caps['modern_functions']:
            self.skipTest('Installed Excel does not support SORT')
        result = engine.edit(self.wid,[dict(op='write_range',sheet='Sheet1',values=[[3],[1],[2]]),dict(op='formula',sheet='Sheet1',range='C1',formula='=SORT(A1:A3)',dynamic=True)],compatibility='installed')
        report = engine.read(self.wid,result['file_id'],'Sheet1','C1:C3')
        self.assertEqual(report['values'],[[1],[2],[3]])
        json.dumps(report)

    def test_06_existing_excel_is_untouched(self):
        import pythoncom
        import win32com.client
        import win32process
        pythoncom.CoInitialize()
        guard = win32com.client.DispatchEx('Excel.Application')
        guard.Visible = False
        book = guard.Workbooks.Add(-4167)
        book.Worksheets(1).Range('A1').Value2 = 'Existing workbook sentinel'
        try:
            baseline = set(win32process.EnumProcesses())
            engine.info(self.wid)
            self.assertEqual(book.Worksheets(1).Range('A1').Value2,'Existing workbook sentinel')
            self.assertIn(win32process.GetWindowThreadProcessId(guard.Hwnd)[1],baseline)
            self.assertFalse(guard.Visible)
        finally:
            book.Close(False)
            guard.Quit()
            pythoncom.CoUninitialize()

    def test_07_worker_termination_closes_only_owned_excel(self):
        import subprocess
        import win32api
        import win32event
        code = "import sys,time,win32process;sys.path.insert(0,sys.argv[1]);from workspace_excel import application\nwith application() as app:\n print(win32process.GetWindowThreadProcessId(app.Hwnd)[1],flush=True);time.sleep(30)"
        child = subprocess.Popen([sys.executable,'-c',code,str(ROOT/'bridge')],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,creationflags=subprocess.CREATE_NO_WINDOW)
        handle = None
        try:
            line = child.stdout.readline().strip()
            self.assertTrue(line.isdigit(),'Worker must publish its owned Excel PID')
            handle = win32api.OpenProcess(0x100000,False,int(line))
            child.kill()
            child.wait(timeout=10)
            self.assertEqual(win32event.WaitForSingleObject(handle,10000),win32event.WAIT_OBJECT_0,'Job ownership must terminate Excel when its worker is killed')
        finally:
            if child.poll() is None:
                child.kill();child.wait(timeout=10)
            if handle is not None: handle.Close()
            child.stdout.close();child.stderr.close()

    def test_08_external_parts_and_formula_references(self):
        source,_ = files.validated(self.wid,self.created['file_id'])
        unsafe = self.folder/'external-parts.xlsx'
        with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(unsafe,'w') as target_zip:
            for name in source_zip.namelist():
                target_zip.writestr(name,source_zip.read(name))
            target_zip.writestr('xl/worksheets/_rels/unsafe.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="image" Target="https://example.invalid/image.png" TargetMode="External"/></Relationships>')
        with self.assertRaisesRegex(ValueError,'外部链接'): engine.check_input(unsafe)
        with self.assertRaises(ValueError): engine.validate_formula("='C:/outside.xlsx'!A1")


if __name__=='__main__':
    unittest.main(verbosity=2)
