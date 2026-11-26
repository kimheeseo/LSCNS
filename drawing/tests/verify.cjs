/* 실제 Three.js r128 수학·지오메트리와 최소 DOM을 이용한 회귀 검사.
   WebGL 화면, 브라우저 레이아웃, 실제 터치 제스처 검사를 대체하지 않습니다. */
'use strict';
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
if(!process.argv[2])throw Error('사용법: node drawing/tests/verify.cjs /absolute/path/three.min.js');
const THREE=require(path.resolve(process.argv[2]));assert.equal(THREE.REVISION,'128');
const html=fs.readFileSync(path.join(__dirname,'../index.html'),'utf8'),script=html.match(/<script>([\s\S]*?)<\/script>/)[1];
const ctx=new Proxy({}, {get:(t,k)=>t[k]??(()=>{}),set:(t,k,v)=>(t[k]=v,true)});
class Element {
 constructor(id=''){this.id=id;this.children=[];this.hidden=false;this.value='';this.textContent='';this.style={};this.dataset={};this.attrs={};this.clientWidth=id==='viewport'?816:260;this.clientHeight=id==='viewport'?854:160;this.offsetWidth=110;this.listeners={};this.classes=new Set();this.classList={add:x=>this.classes.add(x),remove:x=>this.classes.delete(x),contains:x=>this.classes.has(x),toggle:(x,v)=>{const b=v===undefined?!this.classes.has(x):v;b?this.classes.add(x):this.classes.delete(x);return b;}};}
 set className(v){this.classes=new Set(v.split(' '));}get className(){return [...this.classes].join(' ');}
 setAttribute(k,v){this.attrs[k]=v;}append(v){this.children.push(v);}addEventListener(k,v){this.listeners[k]=v;}getContext(){return ctx;}scrollIntoView(){}setPointerCapture(){}getBoundingClientRect(){return {left:0,top:0,width:this.clientWidth,height:this.clientHeight};}
}
const elements=new Map(),el=id=>{if(!elements.has(id))elements.set(id,new Element(id));return elements.get(id);};
for(const id of html.matchAll(/id="([^"]+)"/g))el(id[1]);
const presets=['all','furnace','coating','capstan','spool'].map(v=>{const e=new Element();e.dataset.view=v;return e;});presets[0].classList.add('active');
const panels=['controlsPanel','processPanel'].map(v=>{const e=new Element();e.dataset.panel=v;return e;});
const stepCard=new Element(),document={hidden:false,documentElement:{dataset:{}},getElementById:el,createElement:()=>new Element(),querySelectorAll:s=>s==='.stage'?el('stageList').children:s==='[data-view]'?presets:s==='[data-panel]'?panels:[],querySelector:s=>s==='.step-card'?stepCard:presets[0],addEventListener(){}};
THREE.WebGLRenderer=class {constructor(){this.info={memory:{geometries:0},render:{calls:0}};}setPixelRatio(){}setSize(){}render(scene){const ids=new Set();scene.traverse(o=>{if(o.geometry)ids.add(o.geometry.uuid);});this.info.memory.geometries=ids.size;}};
let frame,now=0;const window={THREE},sandbox={THREE,window,document,devicePixelRatio:1,matchMedia:()=>({matches:false}),ResizeObserver:class{observe(){}},performance:{now:()=>now},requestAnimationFrame:f=>{frame=f;},console};
vm.runInNewContext(script,sandbox,{timeout:10000});const api=window.DrawingTower;assert.ok(api);let checks=0;
function check(name,f){f();checks++;console.log('PASS '+name);}
function input(id,value){el(id).value=String(value);el(id).listeners.input({target:el(id)});}
function click(id){el(id).onclick();}
function tick(ms){now+=ms;frame(now);}
check('r128 / 공정 10단계',()=>{assert.equal(api.stages.length,10);assert.equal(el('stageList').children.length,10);assert.equal(el('crossStage').children.length,10);});
check('기본 125 / 190 / 245 µm',()=>{const p=api.snapshot();assert.equal(p.df,125);assert.equal(p.primary,190);assert.equal(p.coated,245);assert.equal(p.tension,.2);});
check('질량보존과 단위 변환 500개 조건',()=>{for(let i=0;i<500;i++){const vf=.5+(i%196)/10,dp=25+(i%76),vp=.02+(i%300)/10,p=api.calculate(vf,vp,dp);const left=vp/60000*(dp/1000)**2,right=vf*(p.df/1e6)**2;assert.ok(Math.abs(left-right)<1e-18);}});
check('인출 속도 두 배와 자동 공급 맞춤',()=>{input('vf',20);assert.ok(Math.abs(api.snapshot().df-88.388347648)<1e-8);assert.ok(el('diameterStatus').classList.contains('bad'));click('match');assert.equal(api.snapshot().vp,7.5);assert.equal(api.snapshot().df,125);});
check('허용 범위 124 / 126 경계와 이탈',()=>{input('vf',10);for(const d of [123.99,124,125,126,126.01]){const vp=10*60000*(d/50000)**2;input('vp',vp);assert.equal(el('diameterStatus').classList.contains('bad'),d<124||d>126);}});
check('25–100 mm / 0.5–20 m/s 공급 맞춤 범위',()=>{for(const dp of [25,50,100])for(const vf of [.5,10,20]){const p=api.calculate(vf,3.75,dp);assert.ok(p.balancedVp>=.02&&p.balancedVp<=30+1e-9);assert.ok(Math.abs(api.calculate(vf,p.balancedVp,dp).df-125)<1e-9);}});
click('reset');tick(1000);
check('생산 길이와 공급 소모 적분',()=>{assert.equal(api.snapshot().length,10);assert.equal(api.snapshot().consumed,.0625);});
check('일시정지에서 길이·휠 상태 정지',()=>{click('play');const p=api.snapshot();tick(2000);assert.equal(api.snapshot().length,p.length);assert.equal(api.snapshot().phase,p.phase);});
check('5배 가상 시간 적분',()=>{input('timeScale',5);click('play');const length=api.snapshot().length;tick(1000);assert.equal(api.snapshot().length-length,50);});
check('10단계 이전·다음 경계 및 코팅 단면',()=>{click('reset');assert.equal(el('prev').disabled,true);for(let i=0;i<9;i++)click('next');assert.equal(api.snapshot().selected,9);assert.equal(el('next').disabled,true);assert.match(el('crossText').textContent,/245.0/);for(let i=0;i<9;i++)click('prev');assert.equal(el('prev').disabled,true);});
check('보기 토글 / 초기화',()=>{for(const id of ['transparent','crossToggle','scaleToggle','stepMode'])click(id);assert.equal(api.snapshot().transparent,false);assert.equal(api.snapshot().exaggerated,false);assert.equal(el('crossPanel').hidden,true);click('reset');assert.equal(api.snapshot().transparent,true);assert.equal(api.snapshot().stepMode,false);assert.equal(el('crossPanel').hidden,false);});
check('반복 입력·댄서 경로·직경 변경 시 지오메트리 수 고정',()=>{tick(16);const before=api.snapshot().geometries;for(let i=0;i<50;i++){input('vf',.5+i*.39);input('vp',.02+i*.599);input('dp',25+i%76);input('furnaceTemp',1800+i%41*10);tick(16);}assert.equal(api.snapshot().geometries,before);assert.ok(Number.isFinite(api.snapshot().df));});
check('일시정지 중에도 단계 카메라 애니메이션 완료',()=>{click('reset');click('play');click('next');for(let i=0;i<10;i++)tick(100);assert.equal(api.snapshot().length,0);assert.equal(api.snapshot().selected,1);});
console.log(JSON.stringify({passed:checks,revision:THREE.REVISION,geometries:api.snapshot().geometries,note:'최소 DOM, 실제 r128 지오메트리, 렌더러 대역 사용. GPU/실기기 성능은 미검증.'},null,2));
