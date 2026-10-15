"""Reproduce the frozen v1 initials resource. Never update an existing mapping version in place."""
from pathlib import Path
import hashlib
import io
import unicodedata
import urllib.request
import zipfile

URL = 'https://www.unicode.org/Public/17.0.0/ucd/Unihan.zip'
SHA256 = 'f7a48b2b545acfaa77b2d607ae28747404ce02baefee16396c5d2d7a8ef34b5e'
ROOT = Path(__file__).resolve().parent.parent

def main():
    with urllib.request.urlopen(URL, timeout=60) as response:
        data = response.read()
    if hashlib.sha256(data).hexdigest() != SHA256:
        raise ValueError('Unihan archive checksum mismatch')
    rows = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for line in archive.read('Unihan_Readings.txt').decode('utf-8').splitlines():
            if '\tkMandarin\t' not in line:
                continue
            cp, _, reading = line.split('\t')
            initial = unicodedata.normalize('NFD', reading)[0].upper()
            if 'A' <= initial <= 'Z':
                rows.append(cp[2:] + '\t' + initial)
    if len(rows) != 44348:
        raise ValueError('Unexpected initials count')
    target = ROOT / 'src/SheetPilot.Core/Resources/pinyin-initials-v1.tsv'
    target.write_text('\n'.join(rows) + '\n', encoding='ascii')
    print('Generated', len(rows), 'frozen initials; see UNICODE-LICENSE.txt')

if __name__ == '__main__':
    main()
