"""Read legacy Office binary structures without starting Office or executing OLE.

MS-DOC: FIB -> CLX -> current piece table (not a raw string scrape).
MS-PPT: CurrentUser -> edit chain -> current persist map -> slide list.
Ancillary layout, revision markup and legacy slide notes are not reconstructed.
"""
import re
import struct
import olefile


def u16(data, offset):
    return struct.unpack_from('<H', data, offset)[0]


def u32(data, offset):
    return struct.unpack_from('<I', data, offset)[0]


def doc_units(path):
    with olefile.OleFileIO(str(path)) as ole:
        word = ole.openstream('WordDocument').read()
        if u16(word, 0) != 0xA5EC:
            raise ValueError('仅支持 Word 97–2003 DOC，请另存为 DOCX。')
        flags = u16(word, 10)
        if flags & 0x8100:
            raise ValueError('DOC 已加密，请先解密。')
        table = ole.openstream('1Table' if flags & 0x200 else '0Table').read()
        offset = 34 + u16(word, 32) * 2
        longs = offset + 2
        ccp_text = u32(word, longs + 12)
        pairs = longs + u16(word, offset) * 4
        if u16(word, pairs) < 34:
            raise ValueError('DOC 版本过旧，请另存为 DOCX。')
        fc, size = struct.unpack_from('<II', word, pairs + 2 + 33 * 8)
        clx = table[fc:fc + size]
        pos = 0
        while pos < len(clx) and clx[pos] == 1:
            pos += 3 + u16(clx, pos + 1)
        if pos >= len(clx) or clx[pos] != 2:
            raise ValueError('DOC 正文索引不可用，请另存为 DOCX。')
        length = u32(clx, pos + 1)
        plc = clx[pos + 5:pos + 5 + length]
        if len(plc) != length or (length - 4) % 12:
            raise ValueError('DOC 正文索引损坏。')
        count = (length - 4) // 12
        chunks = []
        for i in range(count):
            start, end = u32(plc, 4 * i), u32(plc, 4 * (i + 1))
            if start >= ccp_text:
                break
            length = min(end, ccp_text) - start
            if length < 0:
                raise ValueError('DOC 正文索引顺序无效。')
            location = u32(plc, 4 * (count + 1) + 8 * i + 2)
            compressed = bool(location & 0x40000000)
            location &= 0x3FFFFFFF
            if compressed:
                location //= 2
            raw = word[location:location + length * (1 if compressed else 2)]
            if len(raw) != length * (1 if compressed else 2):
                raise ValueError('DOC 正文已截断。')
            chunks.append(raw.decode('cp1252' if compressed else 'utf-16le', errors='replace'))
        text = ''.join(chunks)
        # Field instructions are not body text; preserve cached visible results.
        text = re.sub('\x13[^\x13\x14\x15]*\x14', '', text)
        text = re.sub('\x13[^\x13\x15]*\x15', '', text).replace('\x15', '')
        text = text.replace('\x07', '\t').replace('\x0b', '\n')
        text = re.sub('[\x00-\x06\x08\x0e-\x1f]', '', text)
        for i, paragraph in enumerate(text.split('\r'), 1):
            if paragraph.strip():
                yield f'legacy DOC body paragraph {i}', paragraph


def records(data, start, end):
    while start + 8 <= end:
        flags, kind, size = struct.unpack_from('<HHI', data, start)
        stop = start + 8 + size
        if stop > end or stop <= start:
            raise ValueError('PPT 记录已损坏。')
        yield start, flags, kind, start + 8, stop
        start = stop


def ppt_units(path):
    with olefile.OleFileIO(str(path)) as ole:
        if ole.exists('EncryptedSummary'):
            raise ValueError('PPT 已加密，请先解密。')
        data = ole.openstream('PowerPoint Document').read()
        current = ole.openstream('Current User').read()
        edit = u32(current, 16)
        doc_id = u32(data, edit + 24)
        persist, seen = {}, set()
        while edit:
            if edit in seen or len(seen) > 10000:
                raise ValueError('PPT 编辑索引损坏。')
            seen.add(edit)
            if u16(data, edit + 2) != 4085:
                raise ValueError('PPT 编辑索引无效。')
            directory = u32(data, edit + 20)
            end, pos = directory + 8 + u32(data, directory + 4), directory + 8
            if end > len(data):
                raise ValueError('PPT 持久化索引损坏。')
            while pos < end:
                header = u32(data, pos)
                first, count = header & 0xFFFFF, header >> 20
                pos += 4
                if not count or pos + count * 4 > end:
                    raise ValueError('PPT 持久化记录损坏。')
                for i in range(count):
                    persist.setdefault(first + i, u32(data, pos + i * 4))
                pos += count * 4
            edit = u32(data, edit + 16)
        if doc_id not in persist:
            raise ValueError('无法定位 PPT 当前文档。')
        doc = persist[doc_id]
        slides = []
        for _, flags, kind, begin, end in records(data, doc + 8, doc + 8 + u32(data, doc + 4)):
            if kind != 4080 or flags >> 4 != 0:
                continue
            for _, _, subkind, subbegin, subend in records(data, begin, end):
                if subkind == 1011:
                    slides.append([u32(data, subbegin), []])
                elif slides and subkind in (4000, 4008):
                    slides[-1][1].append(data[subbegin:subend].decode('utf-16le' if subkind == 4000 else 'cp1252', errors='replace'))
        def text_records(start, end, depth=0):
            if depth > 32:
                raise ValueError('PPT 对象嵌套过深。')
            for _, flags, kind, begin, stop in records(data, start, end):
                if kind in (4000, 4008):
                    yield data[begin:stop].decode('utf-16le' if kind == 4000 else 'cp1252', errors='replace')
                elif flags & 15 == 15:
                    yield from text_records(begin, stop, depth + 1)
        for i, (ref, text) in enumerate(slides, 1):
            if ref not in persist:
                raise ValueError('PPT 幻灯片索引缺失。')
            position = persist[ref]
            text.extend(text_records(position + 8, position + 8 + u32(data, position + 4)))
            unique = list(dict.fromkeys(part for part in text if part.strip()))
            yield f'legacy PPT slide {i} (text only)', '\n'.join(unique) or '[No text. Export this slide to PDF for visual inspection.]'
