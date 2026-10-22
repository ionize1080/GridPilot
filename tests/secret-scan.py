"""Opt-in exact credential leak check; never prints credential bytes or file contents."""
import os, re, subprocess, zipfile
from pathlib import Path

root = Path(__file__).resolve().parent.parent
source = os.environ.get('GRIDPILOT_TEST_KEY_FILE')
if not source:
    raise SystemExit('Set the explicit credential file outside the repository.')
match = re.search(r'sk-[A-Za-z0-9_-]+', Path(source).read_text(encoding='utf-8-sig'))
if not match:
    raise SystemExit('Credential format not recognized; contents withheld.')
patterns = [match.group().encode(encoding) for encoding in ('utf-8', 'utf-16-le', 'utf-16-be')]
overlap = max(map(len, patterns)) - 1
count = 0

def inspect(stream, name):
    global count
    count += 1
    tail = b''
    while block := stream.read(4 * 1024 * 1024):
        block = tail + block
        if any(pattern in block for pattern in patterns):
            raise SystemExit('Credential detected; release blocked. File: ' + str(name))
        tail = block[-overlap:]

tracked = subprocess.check_output(['git', 'ls-files', '-z', '--cached', '--others', '--exclude-standard'], cwd=root).decode().split('\0')
for relative in filter(None, tracked):
    file = root / relative
    if file.is_file():
        with file.open('rb') as stream:
            inspect(stream, relative)
for file in (root / 'dist').rglob('*'):
    if file.is_file():
        with file.open('rb') as stream:
            inspect(stream, file.relative_to(root))
for file in (root / 'release-v2.1.0').glob('*.zip'):
    with zipfile.ZipFile(file) as archive:
        for item in archive.infolist():
            if not item.is_dir():
                with archive.open(item) as stream:
                    inspect(stream, file.name + ':' + item.filename)
print('PASS credential absent from', count, 'source/runtime/archive entries (UTF-8 and UTF-16); credential withheld')
