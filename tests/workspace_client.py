"""Test client for the shipped stdio workspace server. No HTTP listener."""
import json
import os
from pathlib import Path
import queue
import subprocess
import threading


class WorkspaceClient:
    def __init__(self, root, store):
        root = Path(root)
        self.process = subprocess.Popen([str(root / 'Python/python.exe'), str(root / 'bridge/workspace_server.py')],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8',
            env=dict(os.environ, GRIDPILOT_WORKSPACES=str(store), PYTHONIOENCODING='utf-8'), creationflags=subprocess.CREATE_NO_WINDOW)
        self.output = queue.Queue(); self.errors = []; self.sequence = 0
        threading.Thread(target=lambda: [self.output.put(line) for line in self.process.stdout], daemon=True).start()
        threading.Thread(target=lambda: [self.errors.append(line) for line in self.process.stderr], daemon=True).start()
        try:
            self.request('initialize', dict(protocolVersion='2025-06-18', capabilities={}, clientInfo=dict(name='GridPilot Workspace QA', version='2.2.0')))
        except Exception as exc:
            self.close()
            raise RuntimeError('MCP initialization failed: '+''.join(self.errors[-10:])) from exc
        self.process.stdin.write('{"jsonrpc":"2.0","method":"notifications/initialized"}\n'); self.process.stdin.flush()
    def request(self, method, params):
        self.sequence += 1
        self.process.stdin.write(json.dumps(dict(jsonrpc='2.0', id=self.sequence, method=method, params=params)) + '\n'); self.process.stdin.flush()
        while True:
            response = json.loads(self.output.get(timeout=75))
            if response.get('id') != self.sequence: continue
            if 'error' in response: raise RuntimeError(str(response['error']))
            return response['result']
    def call(self, name, **arguments): return self.request('tools/call', dict(name=name, arguments=arguments))
    def close(self):
        self.process.stdin.close()
        try: self.process.wait(timeout=5)
        except subprocess.TimeoutExpired: self.process.kill(); self.process.wait()
        self.process.stdout.close(); self.process.stderr.close()


def payload(result):
    if result.get('isError'): raise AssertionError(result)
    if 'structuredContent' in result: return result['structuredContent']
    return json.loads(next(p['text'] for p in result['content'] if p['type'] == 'text'))
