"""Real parser-family fixtures and protocol tests, including complete long cells."""
import json, os, sys, uuid, zipfile, unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'bridge'))
sys.path.insert(0,str(ROOT/'tests'))
import workspace_files as files
from workspace_client import WorkspaceClient,payload

class Formats(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder=ROOT/'.temp'/('formats-'+uuid.uuid4().hex);cls.folder.mkdir(parents=True)
        cls.store=cls.folder/'store';cls.store.mkdir();cls.wid=uuid.uuid4().hex;cls.sid=uuid.uuid4().hex
        (cls.store/(cls.wid+'.json')).write_text(json.dumps(dict(id=cls.wid,sources=[dict(id=cls.sid,path=str(cls.folder),kind='folder')])))
        os.environ['GRIDPILOT_WORKSPACES']=str(cls.store)
    def read_all(self,name):
        args=dict(workspace_id=self.wid,file_id=files.file_id(dict(id=self.sid),name),count=50,max_chars=1000)
        values=[]
        for _ in range(100):
            page=files.read_file(**args);values.extend(p['text'] for p in page['units'])
            if not page['next']:return ''.join(values)
            args.update(page['next']);args['expected_revision']=page['revision']
        self.fail('Unbounded continuation')
    def test_01_text_extensions(self):
        for ext in sorted(files.TEXT):
            name='text'+ext;(self.folder/name).write_text('中文 evidence 123\nsecond line',encoding='utf-8')
            self.assertIn('中文 evidence 123',self.read_all(name),ext)
    def test_02_long_excel_and_templates(self):
        from openpyxl import Workbook
        for ext in ['xlsx','xlsm','xltx','xltm']:
            book=Workbook();book.active['A1']='HEAD'+('长'*12000)+'TAIL';book.active['B1']='=1+2'
            hidden=book.create_sheet('Hidden');hidden.sheet_state='hidden';hidden['A1']='HIDDEN-EVIDENCE'
            name='long.'+ext;book.save(self.folder/name);book.close()
            text=self.read_all(name);self.assertIn('长'*12000+'TAIL',text);self.assertIn('=1+2',text);self.assertIn('HIDDEN-EVIDENCE',text)
    def test_03_odf_numbers_and_templates(self):
        content=b'<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:p>ODF-EVIDENCE</text:p><table:table-row><table:table-cell office:value="123.5" table:formula="of:=1+2"/></table:table-row></office:document>'
        for ext in ['odt','ods','odp','ott','ots','otp','fodt','fods','fodp']:
            name='open.'+ext
            if ext.startswith('f'): (self.folder/name).write_bytes(content)
            else:
                with zipfile.ZipFile(self.folder/name,'w') as z:z.writestr('content.xml',content)
            text=self.read_all(name);self.assertIn('ODF-EVIDENCE',text);self.assertIn('123.5',text);self.assertIn('of:=1+2',text)
    def test_04_mail_and_web_archive(self):
        from email.message import EmailMessage
        m=EmailMessage();m['Subject']='研究证据';m.set_content('MAIL-EVIDENCE');m.add_alternative('<p>HTML-EVIDENCE</p><script>DO-NOT-EXECUTE</script>',subtype='html');m.add_attachment(b'ATTACHMENT-BODY',maintype='application',subtype='octet-stream',filename='attachment.bin')
        for ext in ['eml','mht','mhtml']:
            name='mail.'+ext;(self.folder/name).write_bytes(m.as_bytes());text=self.read_all(name)
            self.assertIn('MAIL-EVIDENCE',text);self.assertIn('HTML-EVIDENCE',text);self.assertIn('attachment.bin',text);self.assertNotIn('ATTACHMENT-BODY',text);self.assertNotIn('DO-NOT-EXECUTE',text)
    def test_05_image_families(self):
        from PIL import Image
        for ext in sorted(files.IMAGES):
            name='image'+ext;im=Image.new('RGB',(32,32),'#35674a');im.save(self.folder/name)
            result=files.view_file(self.wid,files.file_id(dict(id=self.sid),name));self.assertEqual(result['width'],32,ext)
    def test_06_office_templates(self):
        from docx import Document
        from pptx import Presentation
        doc=Document();doc.add_paragraph('WORD-EVIDENCE')
        for ext in ['docx','docm','dotx','dotm']:
            name='word.'+ext;doc.save(self.folder/name);self.assertIn('WORD-EVIDENCE',self.read_all(name))
        deck=Presentation();slide=deck.slides.add_slide(deck.slide_layouts[1]);slide.shapes.title.text='SLIDE-EVIDENCE'
        for ext in ['pptx','pptm','potx','potm','ppsx','ppsm']:
            name='slides.'+ext;deck.save(self.folder/name)
            types={'pptm':'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml','potx':'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml','potm':'application/vnd.ms-powerpoint.template.macroEnabled.main+xml','ppsx':'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml','ppsm':'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml'}
            if ext in types:
                with zipfile.ZipFile(self.folder/name) as z:parts={n:z.read(n) for n in z.namelist()}
                parts['[Content_Types].xml']=parts['[Content_Types].xml'].replace(b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',types[ext].encode())
                with zipfile.ZipFile(self.folder/name,'w') as z:
                    for n,data in parts.items():z.writestr(n,data)
            self.assertIn('SLIDE-EVIDENCE',self.read_all(name))
    def test_07_stdio_and_resources(self):
        c=WorkspaceClient(ROOT/'dist',self.store)
        try:
            tools=c.request('tools/list',{})['tools'];self.assertEqual(len(tools),12)
            info=payload(c.call('workspace_info'));self.assertEqual(set(info['formats']),files.SUPPORTED)
            page=payload(c.call('workspace_list',workspace_id=self.wid));self.assertTrue(page['entries'])
            resources=c.request('resources/list',{})['resources'];self.assertIn('gridpilot://workspace-guide',[r['uri'] for r in resources])
            self.assertIn('formats',c.request('resources/read',dict(uri='gridpilot://workspace-guide'))['contents'][0]['text'])
        finally:c.close()
    def test_08_readonly_protocol(self):
        old=os.environ.get('GRIDPILOT_MCP_READ_ONLY');os.environ['GRIDPILOT_MCP_READ_ONLY']='1'
        c=WorkspaceClient(ROOT/'dist',self.store)
        try:
            names=[t['name'] for t in c.request('tools/list',{})['tools']]
            self.assertNotIn('workspace_analyze',names);self.assertNotIn('workspace_excel_edit',names)
            self.assertTrue(c.call('workspace_analyze',workspace_id=self.wid,columns=['A'],rows=[[1]]).get('isError'))
        finally:
            c.close()
            if old is None:os.environ.pop('GRIDPILOT_MCP_READ_ONLY',None)
            else:os.environ['GRIDPILOT_MCP_READ_ONLY']=old
    def test_09_native_binary_excel(self):
        import workspace_excel
        with workspace_excel.application() as excel:
            book=excel.Workbooks.Add(-4167)
            try:
                book.Worksheets(1).Range('A1').Value2='BINARY-EVIDENCE'
                book.Worksheets(1).Range('B1').Formula='=123+4'
                for ext,fmt in [('xlsb',50),('xls',56),('xlt',17),('xltx',54)]:
                    name='native.'+ext;book.SaveAs(str(self.folder/name),FileFormat=fmt)
                    text=self.read_all(name);self.assertIn('BINARY-EVIDENCE',text);self.assertIn('127',text)
            finally:book.Close(False)

if __name__=='__main__':unittest.main(verbosity=2)
