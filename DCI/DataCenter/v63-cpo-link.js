(function browserPatch(){
'use strict';
const URL='https://cpo-supply-chain.onrender.com/';
function mount(){
 const existing=document.getElementById('v63-cpo-link');
 if(existing){existing.href=URL;existing.title='Open CPO (Co-Packaged Optics) Supply Chain';return true;}
 const existingBar=document.getElementById('v63-top-links');
 if(existingBar){
  const a=document.createElement('a');
  a.id='v63-cpo-link';a.href=URL;a.target='_blank';a.rel='noopener noreferrer';a.title='Open CPO (Co-Packaged Optics) Supply Chain';
  a.innerHTML='<span class="v63-cpo-dot" aria-hidden="true"></span><span>CPO (Co-Packaged Optics)</span><span class="v63-cpo-arrow" aria-hidden="true">↗</span>';
  existingBar.appendChild(a);return true;
 }
 if(!document.body)return false;
 const bar=document.createElement('nav');
 bar.id='v63-top-links';
 bar.setAttribute('aria-label','Related tools');
 const a=document.createElement('a');
 a.id='v63-cpo-link';
 a.href=URL;
 a.target='_blank';
 a.rel='noopener noreferrer';
 a.title='Open CPO (Co-Packaged Optics) Supply Chain';
 a.innerHTML='<span class="v63-cpo-dot" aria-hidden="true"></span><span>CPO (Co-Packaged Optics)</span><span class="v63-cpo-arrow" aria-hidden="true">↗</span>';
 bar.appendChild(a);
 document.body.insertBefore(bar,document.body.firstChild);
 return true;
}
function start(){
 if(!mount()){
  const timer=setInterval(()=>{if(mount())clearInterval(timer)},250);
  setTimeout(()=>clearInterval(timer),10000);
 }
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();