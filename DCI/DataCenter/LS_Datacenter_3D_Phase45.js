/* Phase 4-5 · v4.7.0 measured FPS panel; device-specific evidence only. */
(function(){
'use strict';
const T=window.__LS3D_TEST__,C=window.LS3D_CONFIG,cache=window.LS3D_PHASE4_CACHE,view=document.getElementById('viewport');
if(!T||!C||!view)return;
const $=id=>document.getElementById(id);
const panel=document.createElement('aside');panel.id='phase5-panel';panel.className='phase58-panel';panel.hidden=true;
panel.innerHTML=[
'<header><b>PHASE 4–5 · 렌더링 / 기기 성능</b><button id="p5-close" type="button" aria-label="닫기">✕</button></header>',
'<p class="p58-note">현재 사용 중인 PC·휴대폰에서 직접 측정합니다. 장치 정보는 브라우저가 공개한 범위만 표시하며 서버에 전송하지 않습니다.</p>',
'<div class="p58-metrics"><div><small>평균 FPS</small><strong id="p5-fps">—</strong></div><div><small>p95 프레임</small><strong id="p5-p95">—</strong></div><div><small>정적 GPU 업로드</small><strong id="p5-static">—</strong></div><div><small>동적 정점</small><strong id="p5-verts">—</strong></div></div>',
'<label>렌더 품질<select id="p5-quality"><option value="low">낮음 · 저전력</option><option value="medium">보통</option><option value="high">높음</option></select></label>',
'<label><input type="checkbox" id="p5-cache" checked> 정적 Geometry GPU 캐시</label>',
'<div class="p58-actions"><button id="p5-run" type="button">10초 성능 측정</button><button id="p5-stop" type="button">취소</button><button id="p5-json" type="button">JSON 저장</button></div>',
'<p id="p5-status" aria-live="polite">실기기 측정 준비. 브라우저의 하드웨어 가속을 켜 주세요.</p>',
'<pre id="p5-device"></pre>',
'<small class="p58-foot">rAF 프레임 시간은 GPU 순수 계산 시간이 아닙니다. SwiftShader 헤드리스 결과는 실제 Windows·Android·iOS 기기의 성능을 대표하지 않습니다.</small>'
].join('');
document.body.appendChild(panel);
const btn=document.createElement('button');btn.id='phase5-open';btn.className='phase58-launch p5';btn.type='button';btn.textContent='▥ FPS / 기기 검증';btn.setAttribute('aria-controls','phase5-panel');view.appendChild(btn);
let opened=false,record=null,active=null,lastUpdate=0;
function device(){
 const cv=document.querySelector('#viewport canvas')||document.querySelector('canvas');
 let gpu='브라우저 비공개',gl=cv?.getContext('webgl')||cv?.getContext('experimental-webgl');
 try{const ext=gl?.getExtension('WEBGL_debug_renderer_info');if(ext)gpu=String(gl.getParameter(ext.UNMASKED_RENDERER_WEBGL)||gpu).slice(0,160)}catch(_){}
 return{userAgent:navigator.userAgent.slice(0,220),viewport:{width:innerWidth,height:innerHeight,dpr:devicePixelRatio||1},gpu,webgl:!!gl,cores:navigator.hardwareConcurrency||null,memoryGiB:navigator.deviceMemory||null};
}
function refresh(){
 $('p5-quality').value=C.quality||'medium';$('p5-cache').checked=!!cache?.enabled;
 $('p5-static').textContent=(cache?.gpu?.uploads??0)+'회';$('p5-verts').textContent=String(window.LS3D_RENDER_METRICS?.triVertices??'—');
 $('p5-device').textContent=JSON.stringify(device(),null,2);
}
function toggle(v){opened=v;panel.hidden=!v;btn.setAttribute('aria-expanded',String(v));if(v)refresh()}
btn.onclick=()=>toggle(!opened);$('p5-close').onclick=()=>toggle(false);
$('p5-quality').onchange=e=>{C.quality=e.target.value;refresh()};
$('p5-cache').onchange=e=>{if(cache){cache.enabled=e.target.checked;cache.clear?.();refresh()}};
function summarize(dts){if(dts.length<6)return null;const s=dts.slice().sort((a,b)=>a-b),at=p=>s[Math.min(s.length-1,Math.floor((s.length-1)*p))];return{frames:dts.length,fps:1000*dts.length/dts.reduce((a,b)=>a+b,0),medianMs:at(.5),p95Ms:at(.95),worstMs:s.at(-1)}}
function finish(cancel=false){
 if(!active)return;cancelAnimationFrame(active.raf);
 const result=summarize(active.deltas);
 record={schema:'lsdc/phase5-device-perf-v1',at:new Date().toISOString(),elapsedSec:(performance.now()-active.start)/1000,completed:!cancel&&!!result,
 device:device(),quality:C.quality,staticCacheEnabled:!!cache?.enabled,sample:result,renderMetrics:{...window.LS3D_RENDER_METRICS},
 staticCache:{hits:cache?.hits,misses:cache?.misses,uploads:cache?.gpu?.uploads,vertices:cache?.gpu?.counts?.reduce((a,b)=>a+b,0)}};
 active=null;
 $('p5-status').textContent=cancel?'측정을 취소했습니다.':'실기기 측정 완료. JSON 결과를 내려받을 수 있습니다.';
 if(result){$('p5-fps').textContent=result.fps.toFixed(1);$('p5-p95').textContent=result.p95Ms.toFixed(1)+' ms'}
 refresh();
}
function start(){
 if(active)finish(true);record=null;
 active={start:performance.now(),last:0,deltas:[],raf:0};$('p5-status').textContent='10초 측정 중 · 다른 탭으로 이동하지 마세요.';
 function frame(t){
  if(!active)return;if(document.hidden){finish(true);$('p5-status').textContent='탭이 비활성화되어 측정을 중지했습니다.';return}
  if(active.last){let ms=t-active.last;if(ms>0&&ms<500)active.deltas.push(ms)}
  active.last=t;if(t-active.start>=10000)finish(false);else active.raf=requestAnimationFrame(frame);
 }
 active.raf=requestAnimationFrame(frame);
}
$('p5-run').onclick=start;$('p5-stop').onclick=()=>finish(true);
$('p5-json').onclick=()=>{
 if(!record){$('p5-status').textContent='10초 측정 완료 후 저장해 주세요.';return}
 const blob=new Blob([JSON.stringify(record,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download='LS_Datacenter_device_FPS_'+new Date().toISOString().slice(0,10)+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
};
window.addEventListener('ls3d-tick',()=>{const t=performance.now();if(!opened||t-lastUpdate<600)return;lastUpdate=t;refresh()});
window.LS3D_PHASE5={open:()=>toggle(true),close:()=>toggle(false),measure:start,cancel:()=>finish(true),get device(){return device()},get result(){return record},get measuring(){return !!active},get cache(){return cache}};
})();