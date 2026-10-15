// LS Datacenter 3D — responsive rack controls disclosure.
// The panel, body and button are provided by scene-adapter; this module does
// not modify any rack/server state or spawn a second rendering loop.
export function createRackPanelDisclosure({panel,body,button,title,collapsed=false}){
  if(!panel||!body||!button)throw new Error('Rack disclosure requires panel, body and button');
  let isCollapsed=!!collapsed;
  function paint(){
    body.hidden=isCollapsed;
    panel.dataset.rackCollapsed=String(isCollapsed);
    panel.style.width=isCollapsed?'min(224px,calc(100% - 16px))':'min(302px,calc(100% - 20px))';
    panel.style.padding=isCollapsed?'7px 9px':'10px';
    panel.style.maxHeight=isCollapsed?'62px':'min(76vh,550px)';
    panel.style.overflowY=isCollapsed?'hidden':'auto';
    button.textContent=isCollapsed?'펼치기 ▾':'접기 ▴';
    button.setAttribute('aria-expanded',String(!isCollapsed));
    button.setAttribute('aria-controls',body.id);
    button.setAttribute('aria-label',isCollapsed?'랙 조작 패널 펼치기':'랙 조작 패널 접기');
    if(title)title.title=isCollapsed?'눌러서 랙 조작 패널을 펼칠 수 있습니다':'';
  }
  function setCollapsed(value){
    isCollapsed=!!value;paint();return isCollapsed;
  }
  function toggle(){return setCollapsed(!isCollapsed)}
  function summary(){return {collapsed:isCollapsed,expanded:!isCollapsed}}
  paint();
  return {toggle,setCollapsed,summary,get collapsed(){return isCollapsed}};
}
