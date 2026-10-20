'use strict';
const $=id=>document.getElementById(id);
const state={online:false,busy:false,model:'single',catalog:null,result:null,dirty:false,view:'geometry',zoom:1,preset:'single'};
const names={single:'단일 코어 모델',mcf:'2코어 MCF 모델',hcf:'이상적 HCF 모세관'};
const descriptions={single:'스칼라 LP01 · 방사형 P1 FEM',mcf:'짝/홀 슈퍼모드 · 곡선 P2 FEM',hcf:'J0 / Dirichlet · 방사형 P1 FEM'};
const physics={single:'약도파 스칼라 LP01. 중심 정칙 조건, 외곽 Dirichlet입니다. 2D 화면은 실제 방사형 장을 회전 대칭으로 재구성합니다.',mcf:'동일 직선 코어 두 개의 스칼라 모드. 1/4 영역에서 y 대칭, x 짝/홀 조건, 외곽 Dirichlet입니다. 벡터 편광·랜덤 누화는 제외합니다.',hcf:'이상적 모세관의 J0/Dirichlet 모델. 중심 정칙, 실제 벽에서 ψ=0. 유리벽·반공진·PML·손실은 계산하지 않습니다.'};
const titles={geometry:'단면과 해석 영역',mesh:'실제 계산 메시',mode:'FEM 모드 분포',profile:'수치 장 / 전력 전달',study:'조건별 해석 결과'};
const actionNames={solve:'모드 해석',mesh:'메시 생성',convergence:'메시 수렴',domain:'영역 수렴',sweep:'파장 스윕'};
const num=(x,n=7)=>Number.isFinite(x)?(x!==0&&(Math.abs(x)<1e-4||Math.abs(x)>1e7)?x.toExponential(4):Number(x.toPrecision(n)).toLocaleString('en-US',{maximumFractionDigits:10})):'—';
function toast(message){$('toast').textContent=message;$('toast').hidden=false;setTimeout(()=>$('toast').hidden=true,6500);}
function log(message,bad=false){const p=document.createElement('p');p.textContent=new Date().toLocaleTimeString('ko-KR',{hour12:false})+'  '+message;if(bad)p.className='bad';$('log').append(p);$('log').scrollTop=$('log').scrollHeight;}
function inputs(){const c={model:state.model};for(const k of Object.keys(state.catalog.defaults))if(k!=='model')c[k]=Number($(k).value);return c;}
function setInputs(c){state.model=c.model;for(const k of Object.keys(state.catalog.defaults))if(k!=='model')$(k).value=c[k];updateModel();drawGeometry();}
function setBusy(b){state.busy=b;$('busy-indicator').hidden=!b;document.querySelectorAll('.backend').forEach(bu=>bu.disabled=b||!state.online||bu.dataset.action==='domain'&&state.model==='hcf');document.querySelectorAll('#parameters input,#preset,#new-model,#open-model').forEach(e=>e.disabled=b);if(!b)updateModel();updateExports();}
function updateExports(){document.querySelectorAll('.result-action').forEach(b=>b.disabled=state.busy||!state.result||state.dirty);}
function markDirty(){if(!state.result)return;state.dirty=true;$('result-badge').textContent='입력 변경 · 재계산 필요';$('tree-result').textContent='재계산 필요';$('metrics').classList.add('stale');updateExports();if(state.view!=='geometry')setView(state.view);}
function updateModel(){
 const model=state.model;$('tree-name').textContent=names[model];$('model-subtitle').textContent=descriptions[model];$('tree-mesh').textContent=model==='mcf'?'곡선 P2 · 1/4 영역':'방사형 P1';$('scope-short').textContent=descriptions[model];$('physics-note').lastElementChild.textContent=physics[model];
 $('outer-label').innerHTML=model==='mcf'?'해석영역 반폭 L <small>µm</small>':'해석영역 반경 R <small>µm</small>';
 document.querySelectorAll('[data-models]').forEach(el=>{el.hidden=!el.dataset.models.split(' ').includes(model);el.querySelectorAll('input').forEach(x=>x.disabled=state.busy||el.hidden);});
 $('domain-tool').disabled=state.busy||!state.online||model==='hcf';
}
function drawGeometry(){
 const c=inputs(),svg=$('geometry-svg');if(Object.values(c).some(v=>typeof v==='number'&&!Number.isFinite(v))){svg.innerHTML='<text x="360" y="220" text-anchor="middle" fill="#b35342">입력값을 확인하십시오.</text>';return;}
 const L=state.model==='hcf'?c.hcf_radius_um: c.outer_um;const scale=165/L,cx=360,cy=218,a=(state.model==='hcf'?c.hcf_radius_um:c.radius_um)*scale;
 let objects='';const outerStyle='fill="#e9f0f8" fill-opacity=".7" stroke="#8aa2be" stroke-width="1.2"';
 if(state.model==='mcf')objects+=`<rect x="${cx-L*scale}" y="${cy-L*scale}" width="${2*L*scale}" height="${2*L*scale}" rx="1" ${outerStyle}/>`;
 else objects+=`<circle cx="${cx}" cy="${cy}" r="${L*scale}" ${outerStyle}/>`;
 const centers=state.model==='mcf'?[-c.pitch_um/2,c.pitch_um/2]:[0];
 for(const x of centers)objects+=`<circle cx="${cx+x*scale}" cy="${cy}" r="${a}" fill="${state.model==='hcf'?'#fff':'#80acd6'}" fill-opacity=".65" stroke="#2b639b" stroke-width="1.6"/>`;
 const b=L*scale;
 svg.innerHTML=`<g><line x1="${cx-b-28}" x2="${cx+b+28}" y1="${cy}" y2="${cy}" stroke="#b4c6d8" stroke-dasharray="3 5"/><line x1="${cx}" x2="${cx}" y1="${cy-b-20}" y2="${cy+b+25}" stroke="#b4c6d8" stroke-dasharray="3 5"/>${objects}<text x="25" y="30" font-size="10" fill="#8a9eb3">GEOMETRY · µm</text><text x="25" y="49" font-size="10" fill="#6a829b">${state.model==='mcf'?'전체 단면 표시 / 실제 FEM은 1/4 영역':state.model==='hcf'?'중공 반경 = 물리적 외벽':'R은 계산 경계이며 실제 클래딩 반경과 다릅니다'}</text><text x="${cx+b+12}" y="${cy-7}" font-size="12" fill="#718aa3">x</text><text x="${cx+9}" y="${cy-b-9}" font-size="12" fill="#718aa3">y</text><line x1="${cx-b}" x2="${cx+b}" y1="408" y2="408" stroke="#8a9fb4"/><line x1="${cx-b}" x2="${cx-b}" y1="403" y2="413" stroke="#8a9fb4"/><line x1="${cx+b}" x2="${cx+b}" y1="403" y2="413" stroke="#8a9fb4"/><text x="${cx}" y="429" text-anchor="middle" font-size="11" fill="#728ba5">${num(2*L,4)} µm</text><text x="696" y="30" text-anchor="end" font-size="11" fill="#4b739b">λ = ${num(c.wavelength_nm,5)} nm</text><text x="696" y="49" text-anchor="end" font-size="10" fill="#7a8fa5">${state.model==='hcf'?'n = '+num(c.n_air,6):'a = '+num(c.radius_um,4)+' µm · Δn = '+num(c.delta_n,5)}</text></g>`;
}
function setView(view){
 state.view=view;state.zoom=1;document.documentElement.style.setProperty('--zoom','1');$('view-heading').textContent=titles[view];
 document.querySelectorAll('[data-view]').forEach(b=>{b.classList.toggle('active',b.dataset.view===view);b.setAttribute('aria-selected',String(b.dataset.view===view));});
 $('geometry-view').hidden=view!=='geometry';$('plot-view').hidden=true;$('empty-view').hidden=true;if(view==='geometry')return;
 let plot=state.result?.plots?.[view];if(view==='study')plot=state.result?.plots?.[state.result?.action];
 if(plot&&!state.dirty){$('plot-image').src=plot.image;$('plot-caption').textContent=plot.caption;$('plot-view').hidden=false;}
 else{$('empty-view').hidden=false;$('empty-title').textContent=state.dirty?'입력이 변경되었습니다':'먼저 해당 계산을 실행하세요';$('empty-message').textContent=state.dirty?'이전 결과를 현재 입력의 계산값으로 표시하지 않습니다. Python 엔진에서 다시 계산하세요.':state.online?'메시는 ‘메시 생성’, 모드는 ‘모드 해석’, 연구 그래프는 수렴 또는 파장 스윕으로 생성합니다.':'이 화면은 GitHub 미리보기입니다. 회사 PC에서 Python 앱을 실행하면 계산할 수 있습니다.';}
}
function table(headers,rows){const t=document.createElement('table'),thead=document.createElement('thead'),head=document.createElement('tr');for(const x of headers){const th=document.createElement('th');th.textContent=x;head.append(th);}thead.append(head);t.append(thead);const tb=document.createElement('tbody');for(const row of rows){const tr=document.createElement('tr');for(const x of row){const td=document.createElement('td');td.textContent=x;tr.append(td);}tb.append(tr);}t.append(tb);return t;}
function showResult(result,saved=false){
 state.result=result;
 state.dirty=Object.keys(result.inputs).some(k=>result.inputs[k]!==inputs()[k]);
 $('metrics').replaceChildren();$('metrics').classList.toggle('stale',state.dirty);$('result-title').textContent=result.action==='solve'?'해석 결과':result.action==='mesh'?'메시 정보':'연구 결과 · 마지막 계산 조건';
 for(const m of result.metrics??[]){const card=document.createElement('div');card.className='metric-card';const label=document.createElement('div');label.className='label';label.textContent=m.label;const value=document.createElement('strong');value.textContent=num(m.value,m.key.includes('neff')?10:7);const unit=document.createElement('small');unit.textContent=m.unit;value.append(unit);const err=document.createElement('div');err.className='error';err.textContent=m.error!==undefined?`${m.error_kind==='absolute'?'절대 오차':'상대 오차'} ${num(m.error,4)}${m.error_kind==='absolute'?'':'%'}`:'직접 계산값 · 독립 오차 미표시';card.append(label,value,err);$('metrics').append(card);}
 if(!result.metrics?.length){const info=document.createElement('div');info.className='metric-placeholder';info.textContent=`실제 메시 생성 완료 · ${num(result.details?.elements)} 요소`;$('metrics').append(info);}
 $('result-badge').textContent=state.dirty?'입력 변경 · 재계산 필요':saved?'저장된 실제 실행 결과':`${num(result.runtime_seconds,3)} s · Python FEM`;$('tree-result').textContent=state.dirty?'재계산 필요':saved?'저장 결과':'해석 완료';
 const body=$('numerical-body');body.replaceChildren();const scope=document.createElement('p');scope.textContent=result.scope+' / '+result.reference_scope;body.append(scope);
 if(result.action!=='solve'&&result.action!=='mesh'){const last=document.createElement('p');last.textContent='위 수치 카드는 마지막 연구 조건입니다. '+(result.last_condition?'λ='+num(result.last_condition.wavelength_nm)+' nm, R/L='+num(result.last_condition.outer_um)+' µm. ':'')+'전체 조건은 아래 표 및 JSON에 저장됩니다.';body.append(last);}
 const metricRows=(result.metrics??[]).map(m=>[m.label,m.reference===undefined?'—':num(m.reference),m.error===undefined?'미검증':num(m.error,5)+(m.error_kind==='absolute'?' (절대)':'%')]);if(metricRows.length)body.append(table(['항목','독립 기준','차이'],metricRows));
 if(result.rows?.length){const keys=Object.keys(result.rows[0]);body.append(table(keys,result.rows.map(r=>keys.map(k=>num(r[k])))));}
 const detail=document.createElement('p');detail.textContent=Object.entries(result.details??{}).map(([k,v])=>k+' = '+num(v,5)).join(' · ');body.append(detail);
 for(const w of result.warnings??[]){const p=document.createElement('p');p.textContent='• '+w;body.append(p);}
 const config=document.createElement('code');config.textContent=JSON.stringify(result.inputs,null,2);body.append(config);updateExports();
 if(result.action==='mesh')setView('mesh');else if(result.action==='solve')setView('mode');else setView('study');
}
async function run(action){
 if(!state.online||state.busy)return;if(!$('parameters').reportValidity())return;const c=inputs();setBusy(true);$('status-text').textContent=actionNames[action]+' 진행 중 — 동일 PC의 Python 계산';$('log-state').textContent='계산 중';log(actionNames[action]+' 시작 · '+names[state.model]);
 try{const res=await fetch('api/jobs',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inputs:c,action})});const task=await res.json();if(!res.ok)throw new Error(task.error||'작업 요청 실패');
 let completed=false;for(let i=0;i<1200;i++){await new Promise(r=>setTimeout(r,500));const response=await fetch('api/jobs/'+task.id);const job=await response.json();if(!response.ok)throw new Error(job.error||'작업 상태 오류');if(job.status==='failed')throw new Error(job.error);if(job.status==='completed'){showResult(job.result);completed=true;log(`${actionNames[action]} 완료 · ${num(job.result.runtime_seconds,3)} s · 수치 기준과 오차를 확인하십시오.`);$('status-text').textContent='계산 완료 · 결과와 입력 기록이 연결되었습니다.';$('log-state').textContent='완료';break;}}
 if(!completed)throw new Error('작업 확인 시간이 초과되었습니다. 계산이 서버에서 계속될 수 있으므로 앱을 확인하십시오.');
 }catch(e){toast(e.message);log(e.message,true);$('status-text').textContent='해석 실패 · 로그 및 입력을 확인하십시오.';$('log-state').textContent='오류';}finally{setBusy(false);}
}
function download(name,content,type){const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),5000);}
function choosePreset(id){const preset=state.catalog.presets.find(p=>p.id===id);if(!preset)return;state.preset=id;setInputs(preset.values);$('preset-note').textContent=preset.note||'수치 검증용 이상화 모델입니다. 실제 제품의 굴절률 단면을 의미하지 않습니다.';markDirty();setView('geometry');log('모델 선택 · '+preset.name);}
async function init(){
 try{state.catalog=await (await fetch('catalog.json')).json();setInputs(state.catalog.defaults);}catch(e){toast('모델 설정을 불러오지 못했습니다. 파일을 직접 열지 말고 Python 앱 또는 GitHub Pages를 사용하십시오.');return;}
 document.querySelectorAll('[data-action]').forEach(b=>b.addEventListener('click',()=>run(b.dataset.action)));
 document.querySelectorAll('[data-view]').forEach(b=>b.addEventListener('click',()=>setView(b.dataset.view)));
 document.querySelectorAll('[data-section]').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('.tree-item').forEach(x=>x.classList.toggle('active',x===b));const s=b.dataset.section;if(s==='results'){$('numerical-details').open=true;$('result-title').scrollIntoView({block:'nearest',behavior:'smooth'});}else if(s==='physics')$('physics-note').scrollIntoView({block:'nearest'});else if(s==='geometry'){setView('geometry');$('group-geometry').open=true;}else{const g=$('group-'+s);if(g){g.open=true;g.scrollIntoView({block:'nearest',behavior:'smooth'});g.querySelector('input')?.focus();}}}));
 $('parameters').addEventListener('submit',e=>e.preventDefault());$('parameters').addEventListener('input',()=>{drawGeometry();markDirty();});$('preset').addEventListener('change',()=>choosePreset($('preset').value));
 $('geometry-tool').onclick=()=>setView('geometry');$('new-model').onclick=()=>{$('preset').value='single';choosePreset('single');};
 $('zoom-in').onclick=()=>{state.zoom=Math.min(3,state.zoom+.25);document.documentElement.style.setProperty('--zoom',state.zoom);};$('zoom-out').onclick=()=>{state.zoom=Math.max(.5,state.zoom-.25);document.documentElement.style.setProperty('--zoom',state.zoom);};$('zoom-reset').onclick=()=>{state.zoom=1;document.documentElement.style.setProperty('--zoom','1');};
 $('save-model').onclick=()=>download('LS_WaveOptics_Model.json',JSON.stringify({format:'ls-wave-optics-model',version:1,inputs:inputs()},null,2),'application/json');
 $('open-model').onclick=()=>$('model-file').click();$('model-file').onchange=async e=>{try{const file=e.target.files[0];if(!file)return;if(file.size>65536)throw new Error('모델 파일이 너무 큽니다.');const obj=JSON.parse(await file.text());if(obj.format!=='ls-wave-optics-model'||obj.version!==1||!obj.inputs)throw new Error('지원하는 모델 JSON 형식이 아닙니다.');const keys=Object.keys(state.catalog.defaults);if(Object.keys(obj.inputs).length!==keys.length||keys.some(k=>!(k in obj.inputs)))throw new Error('모델 입력 항목이 일치하지 않습니다.');if(!['single','mcf','hcf'].includes(obj.inputs.model)||keys.filter(k=>k!=='model').some(k=>!Number.isFinite(obj.inputs[k])))throw new Error('모델 입력값이 잘못되었습니다.');setInputs(obj.inputs);$('preset-note').textContent='사용자가 불러온 모델입니다. 실행 전 범위와 경계조건을 확인하십시오.';markDirty();setView('geometry');log('모델 파일 불러오기 완료');}catch(err){toast(err.message);}e.target.value='';};
 $('export-json').onclick=()=>download('LS_WaveOptics_Result.json',JSON.stringify(state.result,null,2),'application/json');
 $('export-csv').onclick=()=>{const r=state.result;let rows;if(r.rows?.length){const keys=Object.keys(r.rows[0]);rows=[keys,...r.rows.map(x=>keys.map(k=>x[k]))];}else{rows=[['key','label','value','unit','reference','error','error_kind'],...(r.metrics??[]).map(m=>[m.key,m.label,m.value,m.unit,m.reference??'',m.error??'',m.error_kind??''])];}const csv=rows.map(row=>row.map(v=>'"'+String(v).replaceAll('"','""')+'"').join(',')).join('\r\n');download('LS_WaveOptics_Result.csv','\uFEFF'+csv,'text/csv;charset=utf-8');};
 $('export-png').onclick=()=>{const r=state.result,p=r?.plots?.[state.view]||r?.plots?.[r.action]||r?.plots?.mode||Object.values(r?.plots||{})[0];if(!p)return toast('저장할 그래프가 없습니다.');const a=document.createElement('a');a.href=p.image;a.download='LS_WaveOptics_Plot.png';a.click();};
 $('log-toggle').onclick=()=>{const hidden=!$('log').hidden;$('log').hidden=hidden;$('log-toggle').setAttribute('aria-expanded',String(!hidden));};
 try{const controller=new AbortController();const timeout=setTimeout(()=>controller.abort(),3000);let response;try{response=await fetch('api/status',{signal:controller.signal});}finally{clearTimeout(timeout);}if(!response.ok)throw new Error('No backend');const server=await response.json();if(!server.ready)throw new Error('Not ready');state.online=true;$('connection').textContent='Python 엔진 연결됨';$('connection').className='connection online';$('status-text').textContent='로컬 Python FEM 준비 완료';log('실제 Python 엔진 연결 · 계산은 이 PC에서 실행됩니다.');if(!server.models.mcf)log('MCF 추가 의존성이 준비되지 않았습니다. README의 설치 방법을 확인하십시오.',true);
 }catch(e){$('offline').hidden=false;$('connection').textContent='웹 미리보기';$('connection').className='connection offline';$('status-text').textContent='저장된 예시 탐색 · 새 FEM 계산은 로컬 Python 필요';log('Python 엔진 미연결 · 저장된 실제 실행 예시를 불러옵니다.');try{const demo=await (await fetch('sample_result.json')).json();showResult(demo,true);}catch(err){log('저장된 예시가 없습니다. 단면 미리보기는 사용할 수 있습니다.');}}
 setBusy(false);if(state.online)setView('geometry');
}
init();
