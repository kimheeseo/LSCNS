(function browserPatch(){
'use strict';

let coolingOpen=false;

const I18N={
  ko:{button:'스케줄',title:'데이터센터 주요 일정',domestic:'국내',overseas:'국외',empty:'향후 21일 내 확인된 주요 일정이 없습니다.'},
  en:{button:'Schedule',title:'Data Center Schedule',domestic:'Korea',overseas:'Global',empty:'No verified major events in the next 21 days.'},
  ja:{button:'スケジュール',title:'データセンター主要日程',domestic:'韓国',overseas:'海外',empty:'今後21日以内に確認された主要日程はありません。'},
  zh:{button:'日程',title:'数据中心主要日程',domestic:'韩国',overseas:'海外',empty:'未来21天内没有已确认的主要活动。'},
  de:{button:'Termine',title:'Data-Center-Termine',domestic:'Korea',overseas:'International',empty:'Keine bestätigten wichtigen Termine in den nächsten 21 Tagen.'}
};

/* Verified official-source schedule seed, filtered at runtime to today + 21 days. */
const EVENTS=[
  {start:'2026-09-30',end:'2026-09-30',region:'overseas',title:'OCP Educational Webinar: Energy Storage Across the Modern Data Center',url:'https://www.opencompute.org/events/upcoming-events'},
  {start:'2026-10-08',end:'2026-10-08',region:'overseas',title:'AMD Embedded Computing Summit – Paris',url:'https://www.amd.com/en/corporate/events.html'},
  {start:'2026-10-12',end:'2026-10-15',region:'overseas',title:'2026 OCP Global Summit',url:'https://www.opencompute.org/summit/global-summit'},
  {start:'2026-10-12',end:'2026-10-15',region:'overseas',title:'NVIDIA at OCP Global Summit 2026',url:'https://www.nvidia.com/en-us/events/ocp-summit/'},
  {start:'2026-10-14',end:'2026-10-14',region:'overseas',title:'OCP Future Technologies Symposium 2026',url:'https://www.opencompute.org/summit/global-summit/call-for-content'},
  {start:'2026-10-20',end:'2026-10-22',region:'domestic',title:'Nokia SReXperts APAC 2026',url:'https://www.nokia.com/events/srexperts/apac/'},
  {start:'2026-10-20',end:'2026-10-22',region:'overseas',title:'NVIDIA GTC Berlin 2026',url:'https://www.nvidia.com/en-eu/gtc/'}
];

function q(id){return document.getElementById(id)}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}
function uiLang(){
  const raw=String(window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko').toLowerCase();
  if(raw.includes('en'))return'en';
  if(raw.includes('ja')||raw.includes('jp'))return'ja';
  if(raw.includes('zh')||raw.includes('cn'))return'zh';
  if(raw.includes('de'))return'de';
  return'ko';
}
function t(){return I18N[uiLang()]||I18N.ko}

function normalizeHeader(){
  document.querySelectorAll('h1').forEach(h=>{
    const txt=(h.textContent||'').replace(/\s+/g,' ').trim();
    if(/^DC BOM Designer(?:\s+v\d+(?:\.\d+){1,3})?$/i.test(txt) && txt!=='DC BOM Designer'){
      h.textContent='DC BOM Designer';
    }
  });
  const cleanTitle=(document.title||'').replace(/DC BOM Designer\s+v\d+(?:\.\d+){1,3}/ig,'DC BOM Designer');
  if(document.title!==cleanTitle)document.title=cleanTitle;

  const exact='공개 UI + 비공개 계산 엔진 + Rack/Cooling/Optical 시각화';
  document.querySelectorAll('p,span,div').forEach(el=>{
    if(el.children&&el.children.length)return;
    const txt=(el.textContent||'').replace(/\s+/g,' ').trim();
    if(txt===exact)el.textContent='Rack/Cooling/Optical 시각화';
  });
}

function removeLegacyCooling(){
  const old=q('v51-cooling-architecture');
  if(old)old.remove();
}

function hideInlineMultifiberCandidates(){
  const rx=/^multi[- ]?fiber\s+cable\s+candidates$/i;
  const nodes=[...document.querySelectorAll('h1,h2,h3,h4,h5,h6,th,caption,div,span')];
  nodes.forEach(node=>{
    if(node.closest&&node.closest('#multifiber-modal,#multifiber-inline-host'))return;
    const txt=(node.textContent||'').replace(/\s+/g,' ').trim();
    if(!rx.test(txt))return;

    let p=node;
    for(let i=0;i<7&&p;i++,p=p.parentElement){
      if(p.id==='multifiber-modal'||p.id==='multifiber-inline-host')break;
      if(p.querySelector&&p.querySelector('table')){
        p.setAttribute('data-v73-hidden-inline-multifiber','1');
        p.style.setProperty('display','none','important');
        break;
      }
    }
  });
}

function applyCooling(scroll){
  if(document.body)document.body.setAttribute('data-v73-cooling-state',coolingOpen?'open':'closed');
  const panel=q('v53-cooling-architecture');
  if(panel){
    panel.classList.remove('v70-open','v71-cooling-open');
    panel.classList.toggle('v73-cooling-open',coolingOpen);
    panel.hidden=!coolingOpen;
    panel.style.setProperty('display',coolingOpen?'block':'none','important');
  }
  const btn=q('v68-cooling-btn');
  if(btn)btn.setAttribute('aria-expanded',coolingOpen?'true':'false');
  if(scroll&&coolingOpen&&panel)setTimeout(()=>panel.scrollIntoView({behavior:'smooth',block:'start'}),20);
}
function hideCooling(){coolingOpen=false;applyCooling(false)}
function toggleCooling(){coolingOpen=!coolingOpen;applyCooling(coolingOpen)}

function localStartOfDay(){const d=new Date();d.setHours(0,0,0,0);return d}
function parseDate(s){const p=s.split('-').map(Number);return new Date(p[0],p[1]-1,p[2])}
function visibleEvents(){
  const from=localStartOfDay();
  const until=new Date(from);until.setDate(until.getDate()+21);until.setHours(23,59,59,999);
  return EVENTS.filter(e=>parseDate(e.end)>=from&&parseDate(e.start)<=until)
    .sort((a,b)=>a.start.localeCompare(b.start)||a.title.localeCompare(b.title))
    .slice(0,10);
}
function groupHtml(items,region,label){
  const rows=items.filter(e=>e.region===region);
  if(!rows.length)return'';
  return '<section class="v73-group"><h3>'+esc(label)+'</h3><div class="v73-list">'+
    rows.map(e=>'<a class="v73-item" href="'+esc(e.url)+'" target="_blank" rel="noopener noreferrer">'+esc(e.title)+'</a>').join('')+
    '</div></section>';
}
function renderSchedule(){
  const L=t(),m=q('v73-schedule-modal'),btn=q('v73-schedule-btn');
  if(btn){const x=btn.querySelector('.v73-label');if(x)x.textContent=L.button}
  if(!m)return;
  const h=m.querySelector('h2'),body=m.querySelector('.v73-body');
  if(h)h.textContent=L.title;
  if(!body)return;
  const items=visibleEvents();
  body.innerHTML=items.length
    ?groupHtml(items,'domestic',L.domestic)+groupHtml(items,'overseas',L.overseas)
    :'<div class="v73-empty">'+esc(L.empty)+'</div>';
}
function closeSchedule(){const m=q('v73-schedule-modal');if(m)m.hidden=true}
function openSchedule(){renderSchedule();const m=q('v73-schedule-modal');if(m)m.hidden=false}

function mountSchedule(){
  const bar=q('v63-top-links');
  if(!bar)return false;

  let btn=q('v73-schedule-btn');
  if(!btn){
    btn=document.createElement('button');
    btn.type='button';
    btn.id='v73-schedule-btn';
    btn.innerHTML='<span class="v73-dot" aria-hidden="true"></span><span class="v73-label"></span>';
    bar.insertBefore(btn,bar.firstChild);
    btn.addEventListener('click',openSchedule);
  }

  let m=q('v73-schedule-modal');
  if(!m){
    m=document.createElement('div');
    m.id='v73-schedule-modal';
    m.hidden=true;
    m.innerHTML='<div class="v73-card" role="dialog" aria-modal="true" aria-labelledby="v73-title"><div class="v73-head"><h2 id="v73-title"></h2><button type="button" class="v73-close" aria-label="Close">×</button></div><div class="v73-body"></div></div>';
    document.body.appendChild(m);
    m.querySelector('.v73-close').addEventListener('click',closeSchedule);
    m.addEventListener('click',e=>{if(e.target===m)closeSchedule()});
  }
  renderSchedule();
  return true;
}

function housekeeping(){
  normalizeHeader();
  removeLegacyCooling();
  hideInlineMultifiberCandidates();
  applyCooling(false);
  mountSchedule();
}

function start(){
  if(document.body)document.body.setAttribute('data-v73-cooling-state','closed');
  normalizeHeader();
  removeLegacyCooling();
  hideInlineMultifiberCandidates();
  hideCooling();

  document.addEventListener('click',e=>{
    const target=e.target&&e.target.closest?e.target.closest('#v68-cooling-btn,#v68-logic-btn,#v68-structured-btn,[data-lang],[data-language],.lang,.language,.lang-btn'):null;
    if(!target)return;
    if(target.id==='v68-cooling-btn'){
      e.preventDefault();
      e.stopImmediatePropagation();
      toggleCooling();
      return;
    }
    if(target.id==='v68-logic-btn'||target.id==='v68-structured-btn')hideCooling();
    if(target.matches&&target.matches('[data-lang],[data-language],.lang,.language,.lang-btn'))setTimeout(renderSchedule,60);
  },true);

  document.addEventListener('keydown',e=>{if(e.key==='Escape')closeSchedule()});

  let tries=0;
  const mountTimer=setInterval(()=>{
    tries++;
    normalizeHeader();
    removeLegacyCooling();
    hideInlineMultifiberCandidates();
    applyCooling(false);
    if(mountSchedule()||tries>40)clearInterval(mountTimer);
  },250);

  const mo=new MutationObserver(()=>{
    normalizeHeader();
    removeLegacyCooling();
    hideInlineMultifiberCandidates();
    applyCooling(false);
  });
  mo.observe(document.documentElement,{childList:true,subtree:true,characterData:true});

  setInterval(housekeeping,500);
}

if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();