(()=>{
  'use strict';
  let deferredPrompt=null;
  const $=id=>document.getElementById(id);
  const standalone=()=>window.matchMedia('(display-mode: standalone)').matches||window.navigator.standalone===true;
  function hint(message){
    let el=$('dcInstallHint');
    if(!el){el=document.createElement('div');el.id='dcInstallHint';el.className='dc-install-hint';document.body.appendChild(el);}
    el.textContent=message;el.hidden=false;
    clearTimeout(hint.timer);hint.timer=setTimeout(()=>{el.hidden=true},5000);
  }
  function button(){
    const b=$('dcInstallBtn');
    if(!b)return null;
    b.hidden=false;
    if(standalone()){b.textContent='앱 실행 중';b.disabled=true;b.setAttribute('aria-disabled','true');}
    else {b.textContent='앱 설치';b.disabled=false;b.removeAttribute('aria-disabled');}
    return b;
  }
  async function install(){
    const b=button();if(!b||standalone())return;
    if(deferredPrompt){
      deferredPrompt.prompt();
      const choice=await deferredPrompt.userChoice.catch(()=>null);
      deferredPrompt=null;
      if(choice&&choice.outcome==='accepted') hint('앱 설치를 시작했습니다.');
      else hint('설치를 취소했습니다. 필요할 때 다시 누르세요.');
      return;
    }
    const ua=navigator.userAgent||'';
    if(/iphone|ipad|ipod/i.test(ua)) hint('Safari의 공유 버튼 → “홈 화면에 추가”를 선택하세요.');
    else hint('브라우저 메뉴의 “앱 설치” 또는 “홈 화면에 추가”를 선택하세요.');
  }
  window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferredPrompt=e;button();});
  window.addEventListener('appinstalled',()=>{deferredPrompt=null;button();hint('DataCenter BOM 앱 설치가 완료되었습니다.');});
  if('serviceWorker' in navigator){
    window.addEventListener('load',()=>navigator.serviceWorker.register('./sw.js',{scope:'./'}).catch(()=>{}),{once:true});
  }
  function start(){const b=button();if(b)b.addEventListener('click',install);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
