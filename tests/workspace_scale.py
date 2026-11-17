"""Bounded recursive discovery on a large folder, without reading its content."""
import json, os, sys, time, uuid
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'bridge'))
import workspace_files as files
base=ROOT/'.temp'/'workspace-scale';base.mkdir(parents=True,exist_ok=True)
store=base/'store';store.mkdir(exist_ok=True);source=base/'sources';source.mkdir(exist_ok=True)
for i in range(2505):
    p=source/f'synthetic-{i:04}.md'
    if not p.exists():p.write_text('PRIVATE-BODY-NEVER-ENUMERATED')
wid,sid=uuid.uuid4().hex,uuid.uuid4().hex
os.environ['GRIDPILOT_WORKSPACES']=str(store)
(store/(wid+'.json')).write_text(json.dumps(dict(id=wid,sources=[dict(id=sid,path=str(source),kind='folder')])))
cursor='';batches=[];started=time.monotonic()
while True:
    r=files.find_files(wid,query='absent-marker',cursor=cursor)
    assert r['scanned']<=2000 and not r['entries'] and 'PRIVATE-BODY' not in json.dumps(r)
    batches.append(dict(scanned=r['scanned'],elapsed_ms=r['elapsed_ms']));cursor=r['next_cursor']
    if not cursor:break
    assert len(batches)<20
assert sum(b['scanned'] for b in batches)==2506
print(json.dumps(dict(passed=True,files=2505,batches=batches,total_ms=round((time.monotonic()-started)*1000))))
