"""Check the prepared portable package without reading any user profile."""
from pathlib import Path
import hashlib,json,re,sys,zipfile
ROOT=Path(__file__).resolve().parents[1]
base=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else ROOT/'dist'
assert base.is_dir()
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for folder in ['App','Browser','Chatbox','Resources','Python','bridge']:
    assert (base/folder).is_dir(),folder
for name in ['GridPilot-x64.xll','GridPilot-x86.xll','Python/python.exe','bridge/workspace_server.py','Chatbox/LICENSE-PAGEPILOT.txt']:
    assert (base/name).is_file(),name
for source in (ROOT/'bridge').glob('*.py'):
    assert sha(source)==sha(base/'bridge'/source.name),source.name
for name in ['bootstrap.cjs','tools.cjs','catalog.cjs','stdio.cjs','workspace-main.cjs']:
    assert sha(ROOT/'integrations/chatbox'/name)==sha(base/'Chatbox'/name),name
provenance=json.loads((base/'Chatbox/provenance.json').read_text(encoding='utf-8'))
last={p['file']:p['patchedSha256'] for p in provenance['patches']}
for name,digest in last.items():assert sha(base/'Chatbox/native'/name)==digest,name
for file in [base/'Chatbox/bootstrap.cjs',*(base/'bridge').glob('*.py'),*(base/'App').glob('*.js')]:
    text=file.read_text('utf-8')
    assert not any(marker in text for marker in ['synthetic-test-only','remote-debugging-port','GRIDPILOT_TEST_KEY_FILE']),file.name
    assert not re.search(r'gh[pousr]_[A-Za-z0-9]{30,}|sk-[A-Za-z0-9_-]{24,}',text),file.name
for relative in ['Workspaces','Recovery','Chatbox/config.json','Chatbox/Workspaces','Python/Workspaces','App/config.json']:
    assert not (base/relative).exists(),relative
for file in base.iterdir():
    assert not file.name.endswith(('.tools.json','.navigation.json','.settings.json')),file.name
manifest=base/'SHA256SUMS.txt'
if manifest.exists():
    checked=0
    for line in manifest.read_text(encoding='ascii').splitlines():
        digest,relative=line.split('  ',1)
        path=(base/relative).resolve()
        assert path.is_relative_to(base.resolve()),relative
        assert sha(path)==digest.lower(),relative
        checked+=1
    print('PASS package manifest:',checked,'entries')
print('PASS release source/runtime match, final Chatbox patch hashes, required components, no user settings or test instrumentation')
