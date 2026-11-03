"""Browser integration QA for normal CI hosts (not the managed preview host).
Starts the real local Python server, verifies recalculation, stale-result state,
exports, all three models, studies, static preview and mobile overflow.
"""
from pathlib import Path
import sys,subprocess,time,urllib.request,json,contextlib
from playwright.sync_api import sync_playwright,expect
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'evidence';OUT.mkdir(exist_ok=True)


def wait_http(url):
    for _ in range(150):
        try:
            with urllib.request.urlopen(url,timeout=1) as r:
                if r.status==200:return
        except OSError:pass
        time.sleep(.2)
    raise RuntimeError('HTTP server did not start: '+url)


def main():
    records=[];errors=[]
    with (OUT/'browser_server.log').open('w') as log:
        server=subprocess.Popen([sys.executable,str(ROOT/'launch.py'),'--no-browser','--port','8765'],stdout=log,stderr=subprocess.STDOUT)
        static=subprocess.Popen([sys.executable,'-m','http.server','8766','--bind','127.0.0.1','--directory',str(ROOT)],stdout=log,stderr=subprocess.STDOUT)
        try:
            wait_http('http://127.0.0.1:8765/api/status');wait_http('http://127.0.0.1:8766/')
            with sync_playwright() as p:
                browser=p.chromium.launch();ctx=browser.new_context(viewport=dict(width=1600,height=1000),accept_downloads=True)
                page=ctx.new_page();page.on('pageerror',lambda e:errors.append(str(e)))
                page.goto('http://127.0.0.1:8765/');expect(page.locator('#connection')).to_have_text('Python 엔진 연결됨',timeout=20000)
                def run(action):
                    page.locator(f'[data-action="{action}"]').first.click()
                    page.wait_for_function('!state.busy && document.querySelector("#log-state").textContent!=="준비"',timeout=180000)
                    assert page.locator('#log-state').inner_text()=='완료',page.locator('#log').inner_text()
                    page.wait_for_function('document.querySelector("#plot-image").complete && document.querySelector("#plot-image").naturalWidth>0')
                page.screenshot(path=str(OUT/'ui_geometry.png'));records.append('online engine and geometry')
                run('solve');assert page.locator('.metric-card').count()==5
                result=page.evaluate('state.result');assert result['metrics'][0]['error']<2e-7
                page.screenshot(path=str(OUT/'ui_single_result.png'));records.append('real single-core solve, oracle error and saved PNG')
                with page.expect_download() as dl:page.locator('#export-csv').click()
                dl.value.save_as(str(OUT/'downloaded_result.csv'));assert 'neff' in (OUT/'downloaded_result.csv').read_text(encoding='utf-8-sig')
                with page.expect_download() as dl:page.locator('#save-model').click()
                model_file=OUT/'downloaded_model.json';dl.value.save_as(str(model_file));assert json.loads(model_file.read_text())['inputs']['radius_um']==4.1
                page.locator('#radius_um').fill('3.8');expect(page.locator('#export-json')).to_be_disabled();expect(page.locator('#empty-title')).to_have_text('입력이 변경되었습니다');records.append('changed input hides old plot and disables exports')
                page.locator('#model-file').set_input_files(str(model_file));expect(page.locator('#radius_um')).to_have_value('4.1');records.append('model save and reopen')
                page.locator('#preset').select_option('mcf');run('mesh');assert page.evaluate('state.result.details.elements')>1000
                run('solve');m={x['key']:x for x in page.evaluate('state.result.metrics')};assert m['Lc']['error']<.05
                page.screenshot(path=str(OUT/'ui_mcf_result.png'));records.append('actual curved P2 mesh and two-core supermodes')
                page.locator('#preset').select_option('hcf');expect(page.locator('#domain-tool')).to_be_disabled();run('solve')
                first=page.evaluate('state.result.metrics.find(x=>x.key==="D").value');page.locator('#hcf_radius_um').fill('20');run('solve')
                second=page.evaluate('state.result.metrics.find(x=>x.key==="D").value');assert abs(first-second)>1;records.append('HCF radius change recomputes dispersion; no false domain study')
                run('convergence');rows=page.evaluate('state.result.rows');assert rows[-1]['neff_error']<rows[0]['neff_error'];records.append('numerical mesh convergence')
                page.locator('#group-study').evaluate('(e)=>e.open=true');page.locator('#sweep_points').fill('3');run('sweep');assert page.evaluate('state.result.rows.length')==3;records.append('wavelength sweep')
                offline=ctx.new_page();offline.on('pageerror',lambda e:errors.append(str(e)));offline.goto('http://127.0.0.1:8766/');expect(offline.locator('#connection')).to_have_text('웹 미리보기',timeout=20000);expect(offline.locator('#result-badge')).to_have_text('저장된 실제 실행 결과');expect(offline.locator('[data-action="solve"]').first).to_be_disabled()
                offline.screenshot(path=str(OUT/'ui_static_preview.png'));offline.locator('#wavelength_nm').fill('1570');expect(offline.locator('#export-json')).to_be_disabled();expect(offline.locator('#empty-title')).to_have_text('입력이 변경되었습니다');records.append('static preview labels saved evidence; never pretends to solve')
                offline.set_viewport_size(dict(width=390,height=844));assert offline.evaluate('document.documentElement.scrollWidth<=innerWidth+1');offline.screenshot(path=str(OUT/'ui_mobile.png'),full_page=True);records.append('mobile layout has no page overflow')
                assert not errors,errors
                browser.close()
            (OUT/'browser_qa.json').write_text(json.dumps(dict(passed=True,checks=records,js_errors=errors,environment='GitHub Actions Chromium; actual localhost Python backend'),indent=2,ensure_ascii=False)+'\n')
            print('Browser checks passed:',len(records))
        finally:
            for proc in (server,static):
                proc.terminate()
                try:proc.wait(timeout=15)
                except subprocess.TimeoutExpired:proc.kill();proc.wait()


if __name__=='__main__':main()
