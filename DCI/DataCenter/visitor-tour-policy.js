// Shared deterministic visitor camera / gate demonstration policy (seconds).
(function(root){
  const duration=25;
  function gateMotion(elapsed){
    const t=Math.max(0,Number(elapsed)||0),finished=t>=duration;
    const phase=t%5;
    const openness=finished?0:phase<.5?0:phase<1.2?(phase-.5)/.7:phase<2.8?1:phase<3.5?1-(phase-2.8)/.7:0;
    return {openness,completed:Math.min(5,Math.floor(t/5)),finished,duration};
  }
  function gatePosition(hall){return {x:hall.pos[0],z:hall.pos[1]+hall.size[1]/2+2.3};}
  function loadingShot(a){return {eye:[a.x-a.w*.20,1.55,a.z+a.d*.359+5.4],target:[a.x-a.w*.20,.95,a.z+a.d*.20]};}
  root.LS3D_VISITOR_POLICY={gateMotion,gatePosition,loadingShot,gateDuration:duration};
})(globalThis);
