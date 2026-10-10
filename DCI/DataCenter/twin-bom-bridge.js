(function(){
'use strict';
const $=id=>document.getElementById(id);
const panel=document.createElement('section');panel.id='twinBomImportPanel';
panel.innerHTML='<div class="tb-head"><div><b>3D Digital Twin 구성 가져오기</b><small>교육·설계 검토용 가정값을 BOM 산출 입력 초안으로 매핑합니다.</small></div><button type="button" id="tbImportBtn">JSON 가져오기</button><input id="tbFile" type="file" accept=".json,application/json" hidden></div><div id="tbStatus" role="status">lsdc-twin-bom/1.0 JSON 파일 또는 3D 트윈의 공유 링크를 불러오세요.</div><div class="tb-actions"><button type="button" id="tbRun" disabled>BOM 입력값 적용 · 계산</button><a href="./LSDC-TWIN-BOM-SCHEMA.md" target="_blank" rel="noopener">공유 스키마 보기</a></div>';
const security=document.querySelector('.security');if(security)security.after(panel);
const style=document.createElement('style');style.textContent='#twinBomImportPanel{margin:10px auto 14px;max-width:1400px;padding:12px 16px;border:1px solid #35637b;border-radius:12px;background:linear-gradient(100deg,#10283d,#122237);color:#d9eff9;font:12px system-ui}.tb-head{display:flex;align-items:center;justify-content:space-between;gap:10px}.tb-head b{display:block;font-size:13px}.tb-head small{display:block;color:#9bb7c8;margin-top:4px}.tb-head button,.tb-actions button{background:#185267;color:#eaffff;border:1px solid #39889a;border-radius:8px;padding:8px 11px;font-weight:800;cursor:pointer}.tb-actions{display:flex;align-items:center;gap:12px;margin-top:9px}.tb-actions button:disabled{opacity:.45;cursor:not-allowed}.tb-actions a{color:#81d9e1}#tbStatus{margin-top:8px;color:#acc4d3;font-size:11px;line-height:1.5}.tb-error{color:#ffacb7!important}.tb-ok{color:#80e1c2!important}@media(max-width:520px){#twinBomImportPanel{margin:8px;padding:10px}.tb-head{align-items:flex-start;flex-direction:column}.tb-head button{width:100%}.tb-actions{flex-wrap:wrap}}';document.head.append(style);
let packet=null;
function setVal(id,val){const e=$(id);if(!e||val==null||Number.isNaN(val))return;if(e.type==='checkbox')e.checked=val===true||val==='on';else e.value=String(val);e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))}
function decodeHash(){const m=location.hash.match(/(?:^#|&)twin=([^&]+)/);if(!m)return null;const b=m[1].replace(/-/g,'+').replace(/_/g,'/'),p=b+'='.repeat((4-b.length%4)%4);return JSON.parse(decodeURIComponent(escape(atob(p))))}
function mapPacket(o){
 if(!o||o.schemaVersion!=='lsdc-twin-bom/1.0')throw Error('지원하지 않는 공유 JSON 스키마 버전입니다.');
 const f=o.facility||{},cool=f.cooling||{},top=o.optical?.topology||{},ln=Array.isArray(top.links)?top.links:[];
 const racks=Number(f.rackCount),kw=Number(f.rackPowerKw),ratio=Number(f.gpuServerRatio);
 if(!Number.isFinite(racks)||racks<1||racks>500||!Number.isFinite(kw)||kw<0||kw>200||!Number.isFinite(ratio)||ratio<0||ratio>1)throw Error('랙 수(1–500), 랙 전력(0–200 kW), GPU 비율(0–1)을 확인해 주세요.');
 const gpu=Math.max(1,Math.round(racks*ratio*8)),capacity=Number(cool.capacityKw??f.coolingKw)||kw*racks;
 const coolingMode=['air','rear','dlc','mixed'].includes(cool.mode)?cool.mode:'air';
 setVal('targetGPU',gpu);setVal('rackRU',48);setVal('rackPowerKw',kw);
 setVal('rackCoolingKw',capacity/racks);
 setVal('coolingMode',coolingMode==='rear'?'rear-door':coolingMode==='dlc'?'dlc':coolingMode==='air'?'air':'auto');
 setVal('facilityRedundancy',['N','N+1','2N'].includes(f.redundancy)?f.redundancy:'N+1');setVal('facilityEnabled','on');
 setVal('topology',top.type==='fat-tree'?'clos3':'rail');
 const coreMax=Math.max(12,...ln.map(x=>Number(x.cores)||0));
 const coreOptions=[...$('trunkFiberCount').options].map(x=>Number(x.value));setVal('trunkFiberCount',coreOptions.filter(x=>x<=coreMax).pop()||coreOptions[0]?.value||12);
 const server=ln.find(x=>/server-tor/.test(String(x.id))||/DAC|AOC/.test(String(x.type))),up=ln.filter(x=>/spine|leaf|agg/i.test(String(x.id)+' '+String(x.from)+' '+String(x.to)));
 if(server)setVal('serverDistanceM',Number(server.lengthM)||30);
 if(up.length)setVal('leafSpineDistanceM',Math.round(up.reduce((a,x)=>a+(Number(x.lengthM)||0),0)/up.length));
 packet=o;const status=$('tbStatus');status.className='tb-ok';
 const mode=o.optical?.cpoMode==='cpo'?'CPO':'플러거블';
 status.textContent='가져옴: '+racks+' 랙 · '+kw+' kW/랙 · GPU '+gpu+'개 초안(랙당 GPU 서버 1대 × 서버당 8 GPU 가정, BOM 입력 최소 1개) · 냉각 '+coolingMode+' / '+capacity+' kW · '+(f.redundancy||'N+1')+' · '+mode+' · 광 링크 '+ln.length+'개. 결과는 BOM 도구의 기존 계산 기준을 따르며 입력 검토 후 실행하세요.';
 $('tbRun').disabled=false;
}
$('tbImportBtn').onclick=()=>$('tbFile').click();
$('tbFile').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{mapPacket(JSON.parse(r.result))}catch(err){const s=$('tbStatus');s.className='tb-error';s.textContent='JSON을 가져오지 못했습니다: '+err.message}};r.readAsText(f);e.target.value=''};
$('tbRun').onclick=()=>{if(!packet)return;const run=$('run');if(run)run.click();$('tbStatus').textContent+=' · BOM 초안 계산을 실행했습니다.'};
try{const fromHash=decodeHash();if(fromHash)mapPacket(fromHash)}catch(err){const s=$('tbStatus');s.className='tb-error';s.textContent='공유 링크를 읽지 못했습니다: '+err.message}
})();