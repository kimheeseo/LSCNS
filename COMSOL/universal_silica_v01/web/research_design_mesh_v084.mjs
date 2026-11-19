/* STEP 1–2 v0.8.4. Draw only actual Python Gmsh output; no browser FEM. */
export function validateMeshPreview(data){
  if(data?.schema!=='LSCNS_mesh_preview_v084'||!data.input||!data.quality)throw Error('mesh_preview.json v0.8.4 형식이 필요합니다');
  if(!Array.isArray(data.triangles)||!data.triangles.length||data.triangles.length>6000)throw Error('메시 삼각형 표본 수 오류');
  const point=p=>Array.isArray(p)&&p.length===2&&p.every(x=>Number.isFinite(x)&&Math.abs(x)<=1e6);
  for(const t of data.triangles)if(!Array.isArray(t.vertices_um)||t.vertices_um.length!==3||!t.vertices_um.every(point))throw Error('삼각형 좌표 유한성 오류');
  if(!data.boundary_edges||typeof data.boundary_edges!=='object')throw Error('경계 라벨 누락');
  let count=0;
  for(const edges of Object.values(data.boundary_edges)){
    if(!Array.isArray(edges)||(count+=edges.length)>200000)throw Error('경계 크기 오류');
    for(const e of edges)if(!Array.isArray(e)||e.length!==2||!e.every(point))throw Error('경계 좌표 오류');
  }
  if(!Number.isInteger(data.quality.triangles)||data.quality.triangles<data.triangles.length)throw Error('전체 요소 수 오류');
  return data;
}

const canonical=x=>JSON.stringify(sort(x));
function sort(x){return Array.isArray(x)?x.map(sort):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,sort(x[k])])):x;}
export function inputsMatch(a,b){return canonical(a)===canonical(b);}

function draw(canvas,data){
  const ctx=canvas.getContext('2d'),w=canvas.width,h=canvas.height;
  ctx.clearRect(0,0,w,h);ctx.fillStyle='#081727';ctx.fillRect(0,0,w,h);
  const points=data.triangles.flatMap(t=>t.vertices_um);
  const edgePoints=Object.values(data.boundary_edges).flat(2);
  let radius=1;
  for(const p of [...points,...edgePoints])radius=Math.max(radius,Math.abs(p[0]),Math.abs(p[1]));
  const point=p=>[w/2+p[0]/radius*w*.46,h/2-p[1]/radius*h*.46];
  const colour=m=>m==='pml'?'#45556a':m==='air'?'#78cdea':m==='core'?'#f6ba5d':m==='stress_rod'?'#d592cf':'#5a9aa2';
  for(const t of data.triangles){
    ctx.beginPath();t.vertices_um.forEach((p,i)=>ctx[i?'lineTo':'moveTo'](...point(p)));ctx.closePath();
    ctx.fillStyle=colour(t.material);ctx.globalAlpha=.35;ctx.fill();ctx.globalAlpha=.65;ctx.strokeStyle='#80a6bd';ctx.lineWidth=.35;ctx.stroke();
  }
  ctx.globalAlpha=1;
  for(const [name,edges]of Object.entries(data.boundary_edges)){
    ctx.strokeStyle=name==='PML_INTERFACE'?'#ffa450':name.startsWith('OUTER_')?'#f5eeee':'#6ef1ba';ctx.lineWidth=1;
    ctx.beginPath();for(const [a,b]of edges){ctx.moveTo(...point(a));ctx.lineTo(...point(b));}ctx.stroke();
  }
}

export function mountDesignMesh(section){
  const panel=document.createElement('div');panel.id='rf-design-mesh-v084';
  panel.style.cssText='border:1px solid #41677e;border-radius:9px;padding:14px;margin:14px 0';
  panel.innerHTML='<h3>STEP 1·2 · 설계·메시 검사 v0.8.4</h3>'+
    '<p>25개 프리셋은 연구용 형상입니다. Kagome·conjoined·bow-tie 등 일부 명칭은 근사 단면이며, Python의 inspect 명령에서 상세 한계를 확인합니다.</p>'+
    '<p>메시 생성: Python cli.py mesh 설계.json --out out_mesh<br>전 도메인 파장 기준 적용: --resolution-policy wavelength<br>3단계 기하 메시 비교: Python cli.py mesh-study 설계.json --out out_mesh_study</p>'+
    '<label class="rf-file" for="rf-mesh-open">실제 mesh_preview.json 열기</label> <input type="file" accept=".json" id="rf-mesh-open" hidden>'+
    '<p id="rf-mesh-status" role="status">Python 메시 결과를 선택하세요. 웹은 메시·벡터 FEM을 실행하지 않습니다.</p>'+
    '<canvas id="rf-mesh-canvas" width="600" height="600" style="width:100%;max-width:600px;height:auto;display:none" aria-label="실제 Gmsh 메시와 물리 경계"></canvas>'+
    '<pre id="rf-mesh-quality" style="white-space:pre-wrap;word-break:break-word"></pre>';
  section.querySelector('#rf-json').insertAdjacentElement('afterend',panel);
  const input=panel.querySelector('input'),status=panel.querySelector('#rf-mesh-status'),canvas=panel.querySelector('canvas'),quality=panel.querySelector('pre');
  let loaded=false;
  section.querySelector('#rf-json').addEventListener('input',()=>{
    if(!loaded)return;loaded=false;
    canvas.style.display='none';quality.textContent='';status.textContent='설계 입력 변경: 이전 메시 결과가 무효화됐습니다. Python에서 다시 생성한 파일을 불러오세요.';
  });
  input.addEventListener('change',async()=>{
    try{
      const f=input.files[0];if(!f)return;if(f.size>20000000)throw Error('20 MB 제한');
      const data=validateMeshPreview(JSON.parse(await f.text()));
      const same=inputsMatch(data.input,JSON.parse(section.querySelector('#rf-json').value));
      draw(canvas,data);canvas.style.display='block';loaded=true;
      const q=data.quality,fmt=x=>typeof x==='number'&&Number.isFinite(x)?x.toPrecision(5):'—';
      const counts=data.triangles.length+'/'+q.triangles;
      status.textContent=(same?'현재 입력과 일치. ':'현재 입력과 다름: 가져온 파일의 설계 메시를 표시합니다. ')+
        '삼각형 '+counts+'개 표시. 경계는 전체 표시. 모드·손실 수렴 인증이 아닙니다.';
      quality.textContent=['메시 정책: '+q.resolution_policy,'최소/최대 변 길이 [µm]: '+fmt(q.edge_min_um)+' / '+fmt(q.edge_max_um),
        '최대 종횡비: '+fmt(q.aspect_ratio_max),'파장 기준: '+q.wavelength_resolution?.status,
        '원형 벽 법선 표본 검사: '+q.wall_resolution?.status,'경계: '+Object.keys(data.boundary_edges).join(', '),
        ...(q.design_diagnostics?.approximations||[]),...(q.notes||[])].join('\n');
    }catch(e){loaded=false;canvas.style.display='none';quality.textContent='';status.textContent='메시 읽기 실패: '+e.message;}
    finally{input.value='';}
  });
}
