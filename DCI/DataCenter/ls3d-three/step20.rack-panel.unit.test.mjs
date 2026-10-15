// LS Datacenter 3D: rack control collapse/expand regression.
// node --test DCI/DataCenter/ls3d-three/step20.rack-panel.unit.test.mjs
// DOM mock only: browser/WebGL screenshots are a separate acceptance gate.
import test from 'node:test';
import assert from 'node:assert/strict';
import {createRackPanelDisclosure} from './rack-panel-disclosure.mjs';
import {layoutSceneOverlays} from './scene-viewport-policy.mjs';

function element(){
  const attrs={};
  return {style:{},dataset:{},hidden:false,id:'',
    textContent:'',title:'',
    setAttribute(k,v){attrs[k]=v;},
    getAttribute(k){return attrs[k];},
    querySelector(){return null;}
  };
}
function fixture(collapsed){
  const panel=element(),body=element(),button=element(),title=element();
  body.id='ls3dRackMotionBody';
  const disclosure=createRackPanelDisclosure({panel,body,button,title,collapsed});
  return {panel,body,button,title,disclosure};
}

test('mobile starts collapsed, uses aria-controls, retains operable expand button',()=>{
  const {panel,body,button,disclosure}=fixture(true);
  assert.equal(disclosure.collapsed,true);
  assert.equal(body.hidden,true);
  assert.equal(panel.dataset.rackCollapsed,'true');
  assert.equal(panel.style.width,'min(224px,calc(100% - 16px))');
  assert.match(button.textContent,/펼치기/);
  assert.equal(button.getAttribute('aria-expanded'),'false');
  assert.equal(button.getAttribute('aria-controls'),'ls3dRackMotionBody');
});

test('one button toggles expanded rack controls, no change to server or rack state',()=>{
  const {panel,body,button,disclosure}=fixture(true);
  assert.equal(disclosure.toggle(),false);
  assert.equal(body.hidden,false);
  assert.equal(panel.dataset.rackCollapsed,'false');
  assert.match(button.textContent,/접기/);
  assert.equal(button.getAttribute('aria-expanded'),'true');
  assert.equal(panel.style.width,'min(302px,calc(100% - 20px))');
  assert.equal(disclosure.toggle(),true);
  assert.equal(body.hidden,true);
  assert.equal(disclosure.summary().expanded,false);
});

test('desktop starts expanded, and user collapse choice persists across rerenders',()=>{
  const {body,button,disclosure}=fixture(false);
  assert.equal(body.hidden,false);
  assert.equal(disclosure.summary().expanded,true);
  disclosure.setCollapsed(true);
  for(let i=0;i<50;i++)disclosure.summary();
  assert.equal(disclosure.collapsed,true);
  assert.equal(button.getAttribute('aria-label'),'랙 조작 패널 펼치기');
});

test('responsive layout reserves only compact title bar above selected rack',()=>{
  const {panel,disclosure}=fixture(true);
  const colo=element(),mep=element();
  const md=element();md.open=false;
  mep.querySelector=selector=>selector==='details'?md:null;
  const result=layoutSceneOverlays({width:390,height:844,colo,mep,rack:panel});
  assert.equal(result.rackCollapsed,true);
  assert.equal(panel.style.maxHeight,'62px');
  assert.equal(panel.style.overflowY,'hidden');
  assert.equal(mep.style.bottom,'80px');
  disclosure.toggle();
  const expanded=layoutSceneOverlays({width:390,height:844,colo,mep,rack:panel});
  assert.equal(expanded.rackCollapsed,false);
  assert.equal(mep.style.bottom,'calc(32% + 16px)');
  assert.equal(panel.style.overflowY,'auto');
});
