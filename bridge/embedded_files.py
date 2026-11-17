"""Extract reference bytes from Word Flat OPC snapshots, never activate OLE."""
import base64
import io
import json
from pathlib import Path
import re
import struct
import tempfile
import zipfile
import olefile
import workspace_files as files


def native_payload(data):
    pos = 6  # NativeDataSize + flags
    def string():
        nonlocal pos
        end = data.find(b'\0', pos, min(len(data), pos + 4096))
        if end < 0:
            raise ValueError('OLE 包文件名无效。')
        value = data[pos:end]
        pos = end + 1
        return value
    name = string().decode('utf-8', errors='replace')
    string()  # Source path is metadata, never opened.
    pos += 8
    string()  # Temporary path is metadata, never opened.
    size = struct.unpack_from('<I', data, pos)[0]
    pos += 4
    if size > len(data) - pos:
        raise ValueError('OLE 包文件内容已截断。')
    return name, data[pos:pos + size]


def unwrap(data, name='embedded', depth=0):
    if depth > 3:
        raise ValueError('OLE 嵌套过深，请另存附件后读取。')
    if data.startswith(b'%PDF-'):
        return Path(name).stem + '.pdf', data
    if data.startswith(b'PK\x03\x04'):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = set(z.namelist())
            for part, extension in [('word/document.xml', '.docx'), ('xl/workbook.xml', '.xlsx'), ('ppt/presentation.xml', '.pptx')]:
                if part in names:
                    return Path(name).stem + extension, data
        raise ValueError('嵌入 ZIP 不是支持的 Office 文档。')
    if olefile.isOleFile(io.BytesIO(data)):
        with olefile.OleFileIO(io.BytesIO(data)) as ole:
            for stream in ['\x01Ole10Native', 'Package', 'CONTENTS', 'Contents']:
                if ole.exists(stream):
                    raw = ole.openstream(stream).read()
                    if stream == '\x01Ole10Native':
                        label, raw = native_payload(raw)
                        return unwrap(raw, label, depth + 1)
                    return unwrap(raw, name, depth + 1)
            for stream, extension in [('WordDocument', '.doc'), ('Workbook', '.xls'), ('Book', '.xls'), ('PowerPoint Document', '.ppt')]:
                if ole.exists(stream):
                    return Path(name).stem + extension, data
        raise ValueError('该 OLE 对象没有可提取的文档流；仅有图标不代表读取了内容。')
    extension = Path(name).suffix.lower()
    if extension in files.TEXT | files.IMAGES | {'.html', '.htm', '.rtf'}:
        return Path(name).name, data
    raise ValueError('无法识别嵌入对象的实际文件格式。未启动 OLE 应用。')


def process(package_path, operation='list', part=1, start=1, count=3, max_chars=6000, query='', page=1, asset=0, offset=0):
    package = Path(package_path)
    if package.stat().st_size > 200 * 1024 * 1024:
        raise ValueError('对象快照过大。')
    from lxml import etree
    # Large base64 text nodes are normal for Office embeddings. The host and
    # worker bound the snapshot and decoded payload sizes independently.
    tree = etree.parse(str(package), parser=etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True))
    ns = {'pkg': 'http://schemas.microsoft.com/office/2006/xmlPackage'}
    binaries = []
    for element in tree.xpath('/pkg:package/pkg:part', namespaces=ns):
        name = element.get('{'+ns['pkg']+'}name', '')
        if '/embeddings/' in name or '/media/' in name:
            value = element.find('pkg:binaryData', ns)
            if value is not None:
                binaries.append((name, value.text or ''))
    links = [dict(kind='external_link', followed=False, note='外部链接不自动读取；请将目标文件显式添加到工作区。') for e in tree.xpath('//*[local-name()="Relationship"]') if e.get('TargetMode') == 'External' and (e.get('Type', '').endswith('/oleObject') or e.get('Type', '').endswith('/package') or e.get('Type', '').endswith('/image'))]
    # Document payloads precede preview/icon images. Image parts remain inspectable.
    binaries.sort(key=lambda pair: ('/embeddings/' not in pair[0], pair[0]))
    if operation == 'list':
        return dict(parts=[dict(part=i + 1, package_part=name, kind='embedded_document' if '/embeddings/' in name else 'image_or_preview', estimated_bytes=len(text) * 3 // 4) for i, (name, text) in enumerate(binaries)], links=links,
                    note='List only; no document contents read. Images may be OLE icons/previews, not the embedded document. Select an embedded_document part to read actual data.')
    if type(part) is not int or not 1 <= part <= len(binaries):
        raise ValueError('嵌入部分编号无效，请先 list。')
    name, text = binaries[part - 1]
    limit = files.preferences()['maxFileMB'] * 1024 * 1024
    if len(text) * 3 // 4 > limit + 4:
        raise ValueError('嵌入文件超过工作区文件大小上限。')
    raw = base64.b64decode(re.sub(r'\s+', '', text), validate=True)
    filename, data = unwrap(raw, Path(name).name)
    if len(data) > limit:
        raise ValueError('嵌入文件超过大小上限。')
    # Ephemeral grant only within a worker; removed even if parsing raises.
    with tempfile.TemporaryDirectory(prefix='embedded-', dir=files.home()) as temporary:
        target = Path(temporary) / ('reference' + Path(filename).suffix)
        target.write_bytes(data)
        work_id = __import__('uuid').uuid4().hex
        grant = files.home() / (work_id + '.json')
        source = dict(id=__import__('uuid').uuid4().hex, path=str(target), kind='file')
        grant.write_text(json.dumps(dict(id=work_id, sources=[source])), encoding='utf-8')
        try:
            common = dict(workspace_id=work_id, file_id=source['id'] + ':')
            if operation == 'read':
                result = files.read_file(**common, start=start, count=count, max_chars=max_chars, offset=offset)
            elif operation == 'outline':
                result = files.outline_file(**common, start=start, count=min(count, 100))
            elif operation == 'search':
                result = files.search_file(**common, query=query, start=start)
            elif operation == 'view':
                result = files.view_file(**common, page=page, asset=asset)
            else:
                raise ValueError('操作须为 list / outline / read / search / view。')
            result.pop('file_id', None)
            result['name'] = filename
            result['part'] = part
            result['package_part'] = name
            return result
        finally:
            grant.unlink(missing_ok=True)
