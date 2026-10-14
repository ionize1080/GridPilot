"""Build repeatable full/clean portable ZIPs from an already tested dist directory."""
from pathlib import Path
import gzip, hashlib, html.parser, json, shutil, zipfile

ROOT = Path(__file__).resolve().parent
VERSION = '1.1.0'
OUT = ROOT / ('release-v' + VERSION)
DOCS = ['README.txt', 'RELEASE-v1.1.0.md', 'COMPATIBILITY.txt', 'COPYING.txt', 'MASKING.txt', 'SEGMENTATION.txt', 'VALIDATION.txt', 'SETTINGS.txt', 'DATABASE.txt', 'Check-Environment.ps1', 'EXCEL-DNA-LICENSE.txt', 'JIEBA-NET-LICENSE.txt', 'JIEBA-LICENSE.txt', 'NEWTONSOFT-JSON-LICENSE.txt', 'THIRD-PARTY-NOTICES.txt']

class Text(html.parser.HTMLParser):
    def __init__(self):
        super().__init__(); self.parts = []; self.skip = 0
    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'): self.skip += 1
        if tag in ('h1', 'h2', 'h3', 'p', 'br', 'pre', 'li', 'div'): self.parts.append('\n')
    def handle_endtag(self, tag):
        if tag in ('script', 'style'): self.skip -= 1
        if tag in ('h1', 'h2', 'h3', 'p', 'pre', 'li', 'div'): self.parts.append('\n')
    def handle_data(self, data):
        if not self.skip: self.parts.append(data)

def digest(file):
    return hashlib.sha256(file.read_bytes()).hexdigest()

def package():
    OUT.mkdir(exist_ok=True)
    notices = []
    for file in [ROOT / p for p in DOCS if 'LICENSE' in p or 'NOTICES' in p] + [ROOT / 'dist/Browser/LICENSE', ROOT / 'dist/Browser/LICENSES.chromium.html']:
        content = file.read_text(encoding='utf-8-sig')
        if file.suffix == '.html':
            parser = Text(); parser.feed(content); content = ''.join(parser.parts)
        notices.append({'title': ('Electron / ' if file.parent.name == 'Browser' else '') + file.name, 'text': content})
    bundle = gzip.compress(json.dumps(notices, ensure_ascii=False).encode('utf-8'), mtime=0)
    (ROOT / 'dist/App/legal.bin').write_bytes(bundle)
    for clean in (False, True):
        suffix = 'clean' if clean else 'portable'
        base = OUT / ('GridPilot-' + VERSION + '-' + suffix)
        # Never silently package stale output from an earlier attempt.
        if base.exists():
            assert base.resolve().parent == OUT.resolve() and base.name.startswith('GridPilot-1.1.0-')
            shutil.rmtree(base)
        base.mkdir()
        for file in (ROOT / 'dist').iterdir():
            if file.is_file() and file.suffix in ('.dll', '.xll'): shutil.copy2(file, base / file.name)
        for folder in ('App', 'Browser', 'Resources'):
            shutil.copytree(ROOT / 'dist' / folder, base / folder)
        if clean:
            # License texts remain bundled and readable in Settings > About. Dictionary .txt files are runtime data.
            for name in ('LICENSE', 'LICENSES.chromium.html', 'version'):
                (base / 'Browser' / name).unlink(missing_ok=True)
        else:
            for name in DOCS: shutil.copy2(ROOT / name, base / name)
            (base / 'SHA256SUMS.txt').write_text(''.join(digest(p) + '  ' + p.relative_to(base).as_posix() + '\n' for p in sorted(base.rglob('*')) if p.is_file()), encoding='ascii')
        archive = OUT / ('GridPilot-' + VERSION + '-' + suffix + '-win-x64.zip')
        with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for file in sorted(base.rglob('*')):
                if file.is_file(): z.write(file, file.relative_to(OUT))
        with zipfile.ZipFile(archive) as z:
            assert z.testzip() is None
            names = z.namelist()
            assert any(n.endswith('/App/legal.bin') for n in names)
            assert any(n.endswith('/Browser/electron.exe') for n in names)
            assert len([n for n in names if n.endswith('.xll')]) == 2
            if clean: assert not any(n.lower().endswith('.md') or 'license' in Path(n).name.lower() or 'notice' in Path(n).name.lower() for n in names)
        print(archive.name, archive.stat().st_size, digest(archive), flush=True)
    (OUT / 'SHA256SUMS.txt').write_text(''.join(digest(p) + '  ' + p.name + '\n' for p in sorted(OUT.glob('*.zip'))), encoding='ascii')

if __name__ == '__main__': package()
