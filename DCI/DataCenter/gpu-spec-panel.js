(() => {
'use strict';
const VERSION='7.3.2';
const PROFILES={
  h200:{
    model:'NVIDIA DGX H200',gpu:'H200 SXM',arch:'Hopper',badge:'H200',
    gpus:'8 / system',memory:'141 GB HBM3e / GPU · 1,128 GB total',
    bandwidth:'4.8 TB/s / GPU · 38.4 TB/s aggregate',
    compute:'FP8 Tensor: 3.958 PFLOPS / GPU (sparsity)',
    interconnect:'NVLink 900 GB/s / GPU',
    network:'Compute fabric: 8 × 400G logical links presented through 4 physical OSFP twin-port cages.',
    ru:'8 RU',power:'10.2 kW max / system',
    source:'https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html',
    note:'H200 GPU memory/bandwidth follows NVIDIA H200 SXM specs; DGX H200 rack/power/network context follows the DGX H100/H200 user guide.'
  },
  b200:{
    model:'NVIDIA DGX B200',gpu:'Blackwell GPU',arch:'Blackwell',badge:'B200',
    gpus:'8 / system',memory:'180 GB HBM3e / GPU · 1,440 GB total',
    bandwidth:'8 TB/s / GPU · 64 TB/s aggregate',
    compute:'FP8: 72 PFLOPS / system · FP4: 144 PFLOPS / system',
    interconnect:'NVLink aggregate: 14.4 TB/s',
    network:'Compute fabric: 8 × 400G ConnectX-7 logical links through 4 physical OSFP ports (twin-port presentation).',
    ru:'10 RU',power:'14.3 kW max / system',
    source:'https://www.nvidia.com/ko-kr/data-center/dgx-b200/',
    note:'This card intentionally separates logical 400G links from physical OSFP cages so optics/cage counts are not conflated.'
  },
  b300:{
    model:'NVIDIA DGX B300',gpu:'B300 Blackwell Ultra',arch:'Blackwell Ultra',badge:'B300',
    gpus:'8 / system',memory:'288 GB HBM3e / GPU · 2.3 TB total',
    bandwidth:'Up to 8 TB/s / GPU · up to 64 TB/s aggregate',
    compute:'FP8 training: 72 PFLOPS / system · FP4 inference: 144 PFLOPS / system',
    interconnect:'2 × 5th-gen NVSwitch / NVLink',
    network:'Compute: 8 × 800G ConnectX-8 via 8 OSFP ports · Storage/management: 2 × 400G through BlueField-3.',
    ru:'10 RU',power:'14.5 kW NVIDIA physical-spec consumption',
    source:'https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html',
    note:'The 14.5 kW value is the published DGX B300 physical specification. Facility peak-provisioning should remain a separate design envelope.'
  }
};
const I18N={
 ko:{kicker:'선택 GPU / 시스템 사양',selected:'선택 장비',gpus:'GPU 수',memory:'GPU 메모리',bw:'메모리 대역폭',compute:'AI 연산',interconnect:'GPU 인터커넥트',ru:'Rack 크기',power:'시스템 전력',network:'Network role / physical interface',source:'NVIDIA 공식 사양',unsupported:'현재 상세 사양 패널은 H200 / B200 / B300 verified profile에 우선 적용됩니다.',note:'설계 계산과 별도로 선택 즉시 갱신됩니다.'},
 en:{kicker:'Selected GPU / System Specs',selected:'Selected equipment',gpus:'GPU count',memory:'GPU memory',bw:'Memory bandwidth',compute:'AI compute',interconnect:'GPU interconnect',ru:'Rack size',power:'System power',network:'Network role / physical interface',source:'NVIDIA official specs',unsupported:'Detailed verified cards are currently prioritized for H200 / B200 / B300.',note:'Updates immediately when the selected system changes.'},
 ja:{kicker:'選択GPU / システム仕様',selected:'選択装置',gpus:'GPU数',memory:'GPUメモリ',bw:'メモリ帯域幅',compute:'AI演算',interconnect:'GPUインターコネクト',ru:'ラックサイズ',power:'システム電力',network:'Network role / physical interface',source:'NVIDIA公式仕様',unsupported:'詳細な検証済みカードは現在H200 / B200 / B300を優先しています。',note:'システム選択時に即時更新されます。'},
 zh:{kicker:'所选 GPU / 系统规格',selected:'所选设备',gpus:'GPU 数量',memory:'GPU 内存',bw:'内存带宽',compute:'AI 算力',interconnect:'GPU 互连',ru:'机架尺寸',power:'系统功耗',network:'Network role / physical interface',source:'NVIDIA 官方规格',unsupported:'当前详细验证卡优先支持 H200 / B200 / B300。',note:'选择系统后立即更新。'},
 de:{kicker:'Ausgewählte GPU / Systemspezifikation',selected:'Ausgewähltes System',gpus:'GPU-Anzahl',memory:'GPU-Speicher',bw:'Speicherbandbreite',compute:'AI-Rechenleistung',interconnect:'GPU-Interconnect',ru:'Rack-Größe',power:'Systemleistung',network:'Network role / physical interface',source:'Offizielle NVIDIA-Spezifikation',unsupported:'Detaillierte verifizierte Karten sind derzeit für H200 / B200 / B300 priorisiert.',note:'Wird bei Änderung des Systems sofort aktualisiert.'}
};
function textOf(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function lang(){const x=window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko';const s=String(x).toLowerCase();if(s.startsWith('en'))return'en';if(s.startsWith('ja')||s.startsWith('jp'))return'ja';if(s.startsWith('zh')||s.startsWith('cn'))return'zh';if(s.startsWith('de'))return'de';return'ko'}
function selected(){const e=document.getElementById('systemId');if(!e)return{id:'',label:'-'};const o=e.options&&e.options[e.selectedIndex];return{id:String(e.value||'').toLowerCase(),label:o?textOf(o):String(e.value||'-')}}
function mount(){
  if(document.getElementById('gpu-spec-panel'))return true;
  const metric=document.getElementById('mGpu');if(!metric)return false;
  const card=metric.closest('.card');if(!card)return false;
  const metrics=card.querySelector('.metrics'),arch=card.querySelector('.archStrip'),warn=card.querySelector('#resultWarning');
  if(!metrics)return false;
  const layout=document.createElement('div');layout.className='gpu-result-layout';
  const left=document.createElement('div');left.className='gpu-result-left';
  card.insertBefore(layout,metrics);layout.appendChild(left);
  left.appendChild(metrics);if(arch)left.appendChild(arch);if(warn)left.appendChild(warn);
  const panel=document.createElement('aside');panel.id='gpu-spec-panel';layout.appendChild(panel);
  return true;
}
function cell(k,v){return'<div class="gsp-cell"><div class="gsp-k">'+k+'</div><div class="gsp-v">'+v+'</div></div>'}
function render(){
  if(!mount())return;
  const p=document.getElementById('gpu-spec-panel'),sel=selected(),d=PROFILES[sel.id],t=I18N[lang()]||I18N.ko;
  if(!d){
    p.innerHTML='<div class="gsp-head"><div><div class="gsp-kicker">'+t.kicker+'</div><h3>'+t.selected+': '+sel.label+'</h3></div><span class="gsp-badge">v'+VERSION+'</span></div><div class="gsp-gpu"><div class="gsp-chip"></div><div><div class="gsp-model">'+sel.label+'</div><div class="gsp-arch">'+t.unsupported+'</div></div></div><div class="gsp-note">'+t.note+'</div>';
    return;
  }
  p.innerHTML='<div class="gsp-head"><div><div class="gsp-kicker">'+t.kicker+'</div><h3>'+d.model+'</h3></div><span class="gsp-badge">'+d.badge+'</span></div>'+
  '<div class="gsp-gpu"><div class="gsp-chip"></div><div><div class="gsp-model">'+d.gpu+'</div><div class="gsp-arch">'+d.arch+' · '+d.model+'</div></div></div>'+
  '<div class="gsp-grid">'+cell(t.gpus,d.gpus)+cell(t.memory,d.memory)+cell(t.bw,d.bandwidth)+cell(t.compute,d.compute)+cell(t.interconnect,d.interconnect)+cell(t.ru,d.ru)+cell(t.power,d.power)+'</div>'+
  '<div class="gsp-network"><b>'+t.network+'</b><br>'+d.network+'</div>'+
  '<div class="gsp-note">'+d.note+' '+t.note+'</div>'+
  '<a class="gsp-source" href="'+d.source+'" target="_blank" rel="noopener">'+t.source+' ↗</a>';
}
function start(){let last='';const update=()=>{const s=selected().id+'|'+lang();if(s!==last){last=s;render()}};render();const e=document.getElementById('systemId');if(e)e.addEventListener('change',()=>{last='';update()});document.addEventListener('click',()=>setTimeout(update,60),true);document.addEventListener('change',()=>setTimeout(update,30),true);setInterval(update,1000)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();