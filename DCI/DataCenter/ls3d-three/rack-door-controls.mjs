// STEP 11 — independent rear-door hinge motion; front motion retains the original host contract.
// Standalone door controller does not create a new animation loop.
export function createRackDoorControls() {
  const registry=new Map(),stats={frames:0,toggled:0,rearOpen:0};
  const clamp=v=>Math.max(0,Math.min(1,Number(v)||0));
  function ensure(id){if(!registry.has(id))registry.set(id,{targetRear:0,rear:0});return registry.get(id);}
  function setRear(id,open){const s=ensure(id);s.targetRear=open?1:0;stats.toggled++;return s.targetRear;}
  function toggleRear(id){const s=ensure(id);return setRear(id,!s.targetRear);}
  function tick(id,moving,legacy,dt) {
    const state=ensure(id);
    const a=1-Math.exp(-Math.min(.12,Math.max(0,dt||0))*8);
    state.rear+=(state.targetRear-state.rear)*a;
    if(moving.rearDoor)moving.rearDoor.rotation.y=state.rear*Math.PI*.57;
    if(moving.hinge)moving.hinge.rotation.y=-clamp(legacy?.door)*Math.PI*.53;
    stats.frames++;
    stats.rearOpen=[...registry.values()].filter(s=>s.rear>.05).length;
    return {front:clamp(legacy?.door),rear:state.rear,targetRear:state.targetRear};
  }
  function state(id){const s=ensure(id);return {...s};}
  function drop(id){registry.delete(id);}
  function reset(){registry.clear();}
  return {tick,toggleRear,setRear,state,drop,reset,stats};
}
