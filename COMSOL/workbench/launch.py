"""Local-only Python FEM workbench. No cloud account or COMSOL license needed."""
from pathlib import Path
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
from concurrent.futures import ProcessPoolExecutor
from multiprocessing import get_context
from threading import Lock
from urllib.parse import urlsplit
import argparse,json,mimetypes,uuid,time,webbrowser,importlib.util
ROOT=Path(__file__).resolve().parent


def capabilities():
    extra={}
    for name in ['skfem','gmsh']:
        try:__import__(name);extra[name]=True
        except (ImportError,OSError) as e:extra[name]=str(e)
    return dict(ready=True,version='0.1.0',engine='Local Python FEM',models=dict(single=True,hcf=True,mcf=all(v is True for v in extra.values())),optional_dependencies=extra)


class WorkbenchServer(ThreadingHTTPServer):
    daemon_threads=True
    def __init__(self,address):
        self.jobs={};self.job_lock=Lock();self.pool=ProcessPoolExecutor(max_workers=1,mp_context=get_context('spawn'))
        self.capabilities=capabilities()
        super().__init__(address,Handler)
    def server_close(self):
        self.pool.shutdown(wait=True,cancel_futures=True);super().server_close()


class Handler(BaseHTTPRequestHandler):
    def log_message(self,format,*args):print(self.log_date_time_string(),format%args,flush=True)
    def allowed(self):
        port=self.server.server_address[1]
        if self.headers.get('Host') not in [f'127.0.0.1:{port}',f'localhost:{port}']:return False
        origin=self.headers.get('Origin')
        return not origin or origin in [f'http://127.0.0.1:{port}',f'http://localhost:{port}']
    def send_bytes(self,data,mime='application/json',status=200):
        self.send_response(status);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff');self.end_headers();self.wfile.write(data)
    def json(self,data,status=200):self.send_bytes(json.dumps(data,ensure_ascii=False,allow_nan=False).encode(),'application/json; charset=utf-8',status)
    def do_GET(self):
        if not self.allowed():return self.json({'error':'Localhost origin required'},403)
        path=urlsplit(self.path).path
        if path=='/api/status':return self.json(self.server.capabilities)
        if path.startswith('/api/jobs/'):
            key=path.rsplit('/',1)[-1]
            with self.server.job_lock:
                entry=self.server.jobs.get(key)
                if entry is None:return self.json({'error':'작업을 찾을 수 없습니다.'},404)
                future=entry['future']
                if not future.done():return self.json(dict(id=key,status='running' if future.running() else 'queued',started_at=entry['started_at']))
                try:return self.json(dict(id=key,status='completed',result=future.result()))
                except Exception as e:return self.json(dict(id=key,status='failed',error=str(e)))
        if path.startswith('/api/'):return self.json({'error':'Unknown API'},404)
        if path in ['/','/index.html']:path='/index.html'
        target=(ROOT/path.lstrip('/')).resolve()
        if not target.is_relative_to(ROOT) or not target.is_file() or target.suffix not in ['.html','.css','.js','.json','.png','.svg','.ico']:
            return self.json({'error':'Not found'},404)
        # Source and evidence folders are not part of the app's static surface.
        if target.parent!=ROOT:return self.json({'error':'Not found'},404)
        mime=mimetypes.guess_type(target.name)[0] or 'application/octet-stream'
        if target.suffix in ['.html','.css','.js','.json','.svg']:mime+='; charset=utf-8'
        self.send_bytes(target.read_bytes(),mime)
    def do_POST(self):
        if not self.allowed():return self.json({'error':'Localhost origin required'},403)
        if urlsplit(self.path).path!='/api/jobs':return self.json({'error':'Unknown API'},404)
        try:
            if not self.headers.get('Content-Type','').startswith('application/json'):raise ValueError('application/json 요청이 필요합니다.')
            size=int(self.headers.get('Content-Length','0'))
            if not 0<size<=65536:raise ValueError('잘못된 요청 크기입니다.')
            payload=json.loads(self.rfile.read(size));raw=payload['inputs'];action=payload.get('action','solve')
            from engine import normalize,run_job
            c=normalize(raw)
            if action not in ['solve','mesh','convergence','domain','sweep']:raise ValueError('알 수 없는 작업입니다.')
            if not self.server.capabilities['models'][c['model']]:raise ValueError('MCF용 scikit-fem/Gmsh 또는 OS 라이브러리가 설치되지 않았습니다. README를 확인하십시오.')
            with self.server.job_lock:
                if any(not x['future'].done() for x in self.server.jobs.values()):return self.json({'error':'현재 계산을 완료한 뒤 다음 작업을 실행하십시오.'},409)
                while len(self.server.jobs)>=12:self.server.jobs.pop(next(iter(self.server.jobs)))
                key=uuid.uuid4().hex;future=self.server.pool.submit(run_job,c,action)
                self.server.jobs[key]=dict(future=future,started_at=time.time())
            self.json(dict(id=key,status='queued'),202)
        except (ValueError,KeyError,TypeError) as e:self.json({'error':str(e)},400)


def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--port',type=int,default=8765);p.add_argument('--no-browser',action='store_true');args=p.parse_args()
    with WorkbenchServer(('127.0.0.1',args.port)) as server:
        url=f'http://127.0.0.1:{server.server_address[1]}/'
        print('LS Wave Optics Workbench v0.1.0 — '+url,flush=True)
        print('실제 계산은 이 PC에서 실행합니다. 종료: Ctrl+C',flush=True)
        if not args.no_browser:webbrowser.open(url)
        try:server.serve_forever()
        except KeyboardInterrupt:pass


if __name__=='__main__':main()
