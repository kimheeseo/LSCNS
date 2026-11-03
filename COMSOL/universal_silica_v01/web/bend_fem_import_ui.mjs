/**
 * Import independently computed 2D scalar FEM/PML results from pinned Colab.
 * Browser never runs the Python sparse eigensolver; it reads its JSON output.
 * Numerical 2% checks are not product certification.
 */
import {bendLossFromComplexNeff} from './bend_physics_bridge.mjs?v=20261103-velamuri-v1';
const $=id=>document.getElementById(id);
let last=null;
function finite(v){return typeof v==='number'&&Number.isFinite(v)}
function status(msg,ok=false){$('bend2d-status').textContent=msg;$('bend2d-status').className='eng-status '+(ok?'success':'error');}
function reset(){
 for(const id of ['bend2d-neff','bend2d-imag','bend2d-loss','bend2d-mesh','bend2d-pml','bend2d-conclusion'])$(id).textContent='—';
 $('bend2d-summary').textContent='계산된 Python JSON 파일을 선택하세요.';
 const c=$('bend2d-sweep'),ctx=c.getContext('2d');
 ctx.fillStyle='#101b2d';ctx.fillRect(0,0,c.width,c.height);
 ctx.fillStyle='#a6bdd3';ctx.font='18px system-ui';ctx.textAlign='center';
 ctx.fillText('독립 FEM/PML 스윕 결과 대기',c.width/2,c.height/2);
}
function curve(chart,rows,accessor,title,offset,width){
 if(!rows.length)return;
 const ctx=chart.getContext('2d'),W=width,H=chart.height;
 const L=offset+65,R=offset+W-22,T=36,B=H-57;
 const ys=rows.map(r=>r.bend_loss_dB_per_turn);
 if(ys.some(y=>!finite(y)||y<0))return;
 const bottom=Math.min(...ys)*.88,top=Math.max(...ys)*1.12+1e-8;
 const x=i=>L+(R-L)*(rows.length===1?.5:i/(rows.length-1));
 const y=v=>B-(v-bottom)/(top-bottom)*(B-T);
 ctx.strokeStyle='#30465e';ctx.lineWidth=1;ctx.strokeRect(L,T,R-L,B-T);
 ctx.textAlign='center';ctx.fillStyle='#d7e9ff';ctx.font='15px system-ui';
 ctx.fillText(title,offset+W/2,20);
 ctx.font='12px system-ui';ctx.fillStyle='#9bbdd5';
 for(let i=0;i<rows.length;i++)ctx.fillText(accessor(rows[i]),x(i),B+22);
 ctx.textAlign='right';
 for(let j=0;j<3;j++)ctx.fillText((bottom+(top-bottom)*(2-j)/2).toFixed(2),L-7,T+(B-T)*j/2+4);
 ctx.strokeStyle='#53d1ce';ctx.lineWidth=2.5;ctx.beginPath();
 rows.forEach((r,i)=>i?ctx.lineTo(x(i),y(r.bend_loss_dB_per_turn)):ctx.moveTo(x(i),y(r.bend_loss_dB_per_turn)));
 ctx.stroke();
 rows.forEach((r,i)=>{ctx.beginPath();ctx.fillStyle='#ffc083';ctx.arc(x(i),y(r.bend_loss_dB_per_turn),5,0,2*Math.PI);ctx.fill();});
}
function load(data,label){
 if(data?.file_kind!=='LSCNS_BEND_2D_SCALAR_FEM_PML_V1')throw Error('지원하지 않는 JSON 유형입니다. Colab의 FEM 결과 JSON을 사용하세요.');
 const base=data.base,im=base?.neff_imag,reference=base?.input;
 if(!reference||!finite(im)||!finite(base.neff_real)||!finite(reference.wavelength_nm)||
   !finite(reference.bend_diameter_mm)||!Array.isArray(data.mesh)||!Array.isArray(data.pml))
   throw Error('FEM 결과 JSON의 필수 데이터가 누락되었습니다.');
 const ownLoss=base.bend_loss_dB_per_turn;
 if(im>=0&&finite(ownLoss)){
  const computed=bendLossFromComplexNeff({
   wavelengthNm:reference.wavelength_nm,bendDiameterMm:reference.bend_diameter_mm,
   neffImagMagnitude:im});
  if(Math.abs(computed.lossDbPerTurn-ownLoss)>1e-6*Math.max(1,Math.abs(ownLoss)))
   throw Error('수입된 손실과 복소 굴절률 환산식이 맞지 않습니다.');
 }
 if(data.mesh.length<2||data.pml.length<2)throw Error('수렴 비교에는 메시 및 PML 결과가 각각 2개 이상 필요합니다.');
 const meshChange=data.mesh_last_two_change_pct,pmlChange=data.pml_last_two_change_pct;
 if(!finite(meshChange)||!finite(pmlChange))throw Error('수렴 변화율이 유효하지 않습니다.');
 const accepted=(im>=0&&finite(ownLoss)&&finite(base.core_fraction)&&data.last_two_sweeps_passed===true&&finite(data.tolerance_percent)&&meshChange<=data.tolerance_percent&&pmlChange<=data.tolerance_percent);
 last=data;
 $('bend2d-neff').textContent=base.neff_real.toFixed(9);
 $('bend2d-imag').textContent=im.toExponential(6);
 $('bend2d-loss').textContent=ownLoss===null||!finite(ownLoss)?'산출 불가':ownLoss.toFixed(5)+' dB/turn';
 $('bend2d-mesh').textContent=meshChange.toFixed(3)+'%';
 $('bend2d-pml').textContent=pmlChange.toFixed(3)+'%';
 $('bend2d-conclusion').textContent=accepted?'2개 수치 기준 통과 (제품 검증 아님)':'수렴 미완료 — 수치 결과를 제품 성능으로 사용 금지';
 $('bend2d-summary').textContent=
  label+' · '+base.solver_level+' · 굽힘 직경 '+reference.bend_diameter_mm+' mm · 파장 '+
  reference.wavelength_nm+' nm · 코어 에너지 분율 '+(100*base.core_fraction).toFixed(1)+'%. '+
  'Full-vector FEM/GEBT·실제 RIP·물리 클래딩/코팅·도메인 검증은 별도로 필요합니다.';
 const c=$('bend2d-sweep'),ctx=c.getContext('2d');
 ctx.fillStyle='#101b2d';ctx.fillRect(0,0,c.width,c.height);
 curve(c,data.mesh,r=>(r.grid_h_requested_um??r.input?.grid_h_um??0).toString()+'µm',
   'Mesh refinement (h)',0,c.width/2);
 curve(c,data.pml,r=>(r.pml_thickness_requested_um??r.input?.pml_thickness_um??0).toFixed(1)+'µm',
   'PML thickness',c.width/2,c.width/2);
 status(accepted?'수치 두 조건 만족 · 물리적 검증 필요':'수렴 미완료: 메시/PML 기준 미충족 또는 음의 Im(n_eff)',accepted);
}
$('bend2d-file').addEventListener('change',async e=>{
 const file=e.target.files?.[0];if(!file)return;
 try{
  if(file.size>8_000_000)throw Error('파일이 8MB보다 큽니다.');
  load(JSON.parse(await file.text()),file.name);
 }catch(err){last=null;reset();status('불러오기 실패: '+err.message);}
 e.target.value='';
});
$('bend2d-example').addEventListener('click',async()=>{
 try{
  status('GitHub 시연용 JSON 로딩 중');
  const response=await fetch('../bending_fem/example_validation.json?v=20261103');
  if(!response.ok)throw Error('예제 파일 로드 실패 HTTP '+response.status);
  load(await response.json(),'실제 시험 실행 사례 · 제조사 모델 아님');
 }catch(e){last=null;reset();status('예제 읽기 실패: '+e.message);}
});
reset();
