(function browserComputeSilicon(){
'use strict';
const GPU={
 nvidia:{vendor:'NVIDIA',family:'Hopper / Blackwell',note:'Current DGX/HGX profiles in this tool are NVIDIA-based.'},
 amd:{vendor:'AMD',family:'Instinct MI350X / MI355X (MI350 Series)',note:'Physical AI accelerator procurement reference. Use an OEM/server platform profile before applying rack power, cooling, and fabric quantities.'},
 intel:{vendor:'Intel',family:'Data Center GPU Max Series',note:'Supplier reference; lifecycle/availability must be checked for a real procurement.'},
 biren:{vendor:'Biren Technology',family:'BR100 family',note:'Supplier reference; regional availability/compliance must be checked.'},
 google:{vendor:'Google Cloud',family:'TPU7x (Ironwood) / TPU v6e (Trillium)',note:'Cloud accelerator reference. TPU is consumed as Google Cloud capacity/reservation, not as a standard field-purchased GPU BOM line item.'}
};
const CPU={
 intel:{vendor:'Intel',family:'Xeon 6',note:'Server CPU reference for AI/HPC/data-center host systems.'},
 amd:{vendor:'AMD',family:'EPYC 9005 Series',note:'Server CPU reference.'},
 nvidia:{vendor:'NVIDIA',family:'Grace CPU',note:'Arm-based data-center CPU reference.'},
 ampere:{vendor:'Ampere Computing',family:'AmpereOne',note:'Arm server CPU reference.'}
};
function q(id){return document.getElementById(id)}
function txt(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function mount(){
 if(q('v61-compute-silicon'))return q('v61-compute-silicon');
 var anchor=q('v60-system-engineering')||q('v48-power-model')||q('v48-role-model');if(!anchor)return null;
 var p=document.createElement('section');p.id='v61-compute-silicon';
 p.innerHTML='<h3>Compute Silicon Suppliers · GPU / AI Accelerator / Server CPU</h3>'+
 '<div class="v61-sub">GPU 공급업체 생태계와 서버 CPU를 BOM 설계 화면에 분리해 표시합니다. 아래 Supplier Reference는 조달/아키텍처 비교용이며, 검증된 vendor-specific system profile이 없는 경우 기존 랙·전력·네트워크 계산값을 임의로 바꾸지 않습니다.</div>'+
 '<div class="v61-grid">'+
  '<div class="v61-card"><h4>GPU / AI Accelerator Suppliers</h4>'+
   '<div class="v61-row"><label>Reference vendor</label><select id="v61-gpu-vendor"><option value="nvidia">NVIDIA</option><option value="amd">AMD Instinct</option><option value="google">Google TPU</option><option value="intel">Intel</option><option value="biren">Biren Technology</option></select></div>'+
   '<div class="v61-row"><label>Product family</label><div id="v61-gpu-family"></div></div>'+
   '<div class="v61-chips"><span class="v61-chip primary">NVIDIA</span><span class="v61-chip">AMD Instinct</span><span class="v61-chip">Google TPU</span><span class="v61-chip">Intel</span><span class="v61-chip">Biren Technology</span></div><div class="v61-note" id="v61-gpu-note"></div>'+
  '</div>'+
  '<div class="v61-card"><h4>Server CPU Suppliers</h4>'+
   '<div class="v61-row"><label>Reference vendor</label><select id="v61-cpu-vendor"><option value="intel">Intel</option><option value="amd">AMD</option><option value="nvidia">NVIDIA</option><option value="ampere">Ampere Computing</option></select></div>'+
   '<div class="v61-row"><label>Product family</label><div id="v61-cpu-family"></div></div>'+
   '<div class="v61-chips"><span class="v61-chip primary">Intel · Xeon 6</span><span class="v61-chip">AMD · EPYC 9005</span><span class="v61-chip">NVIDIA · Grace</span><span class="v61-chip">Ampere · AmpereOne</span></div><div class="v61-note" id="v61-cpu-note"></div>'+
  '</div>'+
 '</div>';
 anchor.insertAdjacentElement('afterend',p);
 q('v61-gpu-vendor').addEventListener('change',render);
 q('v61-cpu-vendor').addEventListener('change',render);
 return p;
}
function upsert(cat,item,qty,key){
 var body=q('bom');if(!body)return;
 var row=body.querySelector('tr[data-v61-row="'+key+'"]');
 if(!row){row=document.createElement('tr');row.setAttribute('data-v61-row',key);row.innerHTML='<td></td><td></td><td></td>';body.appendChild(row)}
 var td=row.querySelectorAll('td');if(td.length>=3){td[0].textContent=cat;td[1].textContent=item;td[2].textContent=qty}
}
function render(){
 var p=mount();if(!p)return;
 var g=GPU[q('v61-gpu-vendor').value]||GPU.nvidia,c=CPU[q('v61-cpu-vendor').value]||CPU.intel;
 q('v61-gpu-family').textContent=g.vendor+' · '+g.family;
 q('v61-gpu-note').textContent=g.note;
 q('v61-cpu-family').textContent=c.vendor+' · '+c.family;
 q('v61-cpu-note').textContent=c.note;
 upsert('GPU / AI Accelerator','GPU supplier reference · '+g.vendor+' · '+g.family,'Reference','gpu-supplier');
 upsert('Server CPU','Server CPU supplier reference · '+c.vendor+' · '+c.family,'Reference','cpu-supplier');
 try{
  window.currentDesignSnapshot=window.currentDesignSnapshot||{};
  window.currentDesignSnapshot.results=window.currentDesignSnapshot.results||{};
  window.currentDesignSnapshot.results.computeSuppliers={gpu:{vendor:g.vendor,family:g.family},cpu:{vendor:c.vendor,family:c.family},referenceOnly:true};
 }catch(e){}
}
function start(){if(!mount()){setTimeout(start,400);return}render();document.addEventListener('click',function(e){if(e.target&&e.target.closest&&e.target.closest('button'))setTimeout(render,100)},true)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();