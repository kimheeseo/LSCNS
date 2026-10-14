// STEP 08 — deterministic, per-worker equipment inspection state machine.
// Does not mutate original routes, equipment status, metrics or incident rules.
// Real sequence at each existing worker pause: approach -> align -> inspect -> tablet -> depart.
export function createInspectionFSM(T){
 const states=new Map(),history=[],limits={align:.50,inspect:.76,tablet:.64};
 const metrics={transitions:0,completed:0,restarts:0,alerts:0};
 const zonePrefs={
  'eng-hall':/rack|tor|odf|crah|storage|ups|pdu/i,
  'eng-utility':/xfmr|switchgear|generator|ats|battery/i,
  'eng-ops':/noc|core|spine|odf|security|mmr/i,
  'eng-campus':/gate|entry|xray|metal|sign|desk/i
 };
 const incidentPriority={
  power:/switchgear|generator|ats|ups|xfmr/i,
  cooling:/tower|chiller|cdu|crah|crac|pump/i,
  network:/core|spine|leaf|odf|carrier/i,
  fire:/fire|security|gate/i,
  'dlc-leak':/cdu|coolant|pump/i,
  'rack-hotspot':/rack|gpu|tor/i
 };
 function pickTarget(n,pose,scenario='normal'){
  const all=(T.visibleAssets?.()||[]).filter(a=>a.zone===n.zone&&!a.hidden);
  const prefer=incidentPriority[scenario]||zonePrefs[n.id]||/rack|ups|cool/i;
  const matches=all.filter(a=>prefer.test(a.id+' '+a.type+' '+a.name));
  const list=matches.length?matches:all;
  return list.slice().sort((a,b)=>
   Math.hypot(a.x-pose.x,a.z-pose.z)-Math.hypot(b.x-pose.x,b.z-pose.z))[0]||null;
 }
 function log(n,state,target,at){
  history.push({id:n.id,state,targetId:target?.id||null,time:Math.round(at*100)/100});
  if(history.length>120)history.shift();metrics.transitions++;
 }
 function update(n,pose,clock,scenario='normal'){
  if(!n||n.point||!pose||!Number.isFinite(pose.x)||!Number.isFinite(pose.z))return pose;
  const time=Number.isFinite(clock)?clock:0,mode=T.facilityMode||'aidc';
  let s=states.get(n.id);
  if(s&&(time<s.lastTime-.1||s.mode!==mode)){
    states.delete(n.id);s=null;metrics.restarts++;
  }
  const moving=pose.action==='walk';
  if(!s){
    s={state:'move',since:time,lastTime:time,mode,target:null,completed:0,
       scenario,origin:{x:pose.x,z:pose.z}};
    states.set(n.id,s);log(n,'move',null,time);
  }
  s.lastTime=time;
  if(moving){
    if(s.state!=='move'){
      // Completion is recorded at tablet -> depart, never twice on resumed movement.
      s.state='depart';s.since=time;log(n,'depart',s.target,time);
    }
    if(s.state==='depart'&&time-s.since>.25){
      s.state='move';s.target=null;log(n,'move',null,time);
    }
    return {...pose,inspectionState:s.state,targetId:s.target?.id||null};
  }
  if(s.state==='depart'){
    // Remain in departing/rest state until the legacy navigation actually resumes.
    // Otherwise a long stationary pose would restart the entire sequence every frame.
    return {...pose,action:'pause',inspectionState:'depart',targetId:s.target?.id||null,
      inspectionAlert:scenario!=='normal'};
  }
  if(s.state==='move'){
    s.target=pickTarget(n,pose,scenario);
    s.state='align';s.since=time;s.scenario=scenario;
    s.origin={x:pose.x,z:pose.z};
    if(s.scenario!=='normal')metrics.alerts++;
    log(n,'align',s.target,time);
  }else if(s.scenario!==scenario && scenario!=='normal'){
    s.target=pickTarget(n,pose,scenario);
    s.scenario=scenario;s.state='align';s.since=time;metrics.alerts++;
    log(n,'align',s.target,time);
  }
  const age=Math.max(0,time-s.since);
  if(s.state==='align'&&age>=limits.align){
    s.state='inspect';s.since=time;log(n,'inspect',s.target,time);
  }else if(s.state==='inspect'&&age>=limits.inspect){
    s.state='tablet';s.since=time;log(n,'tablet',s.target,time);
  }else if(s.state==='tablet'&&age>=limits.tablet){
    s.state='depart';s.since=time;metrics.completed++;log(n,'depart',s.target,time);
  }
  const target=s.target;
  // Rotate in place toward the selected nearby cabinet, without teleporting
  // or changing A* collision-checked x/z coordinates.
  const angle=target?Math.atan2(target.x-pose.x,target.z-pose.z):pose.heading;
  const facing=Number.isFinite(angle)?angle:pose.heading;
  const action=s.state==='inspect'||s.state==='tablet'?'inspect':
    s.state==='align'?'pause':'pause';
  return {...pose,heading:facing,action,inspectionState:s.state,
    targetId:target?.id||null,showTablet:s.state==='tablet',
    inspectionAlert:scenario!=='normal'};
 }
 function diagnostic(){
  return {metrics:{...metrics},workers:[...states.entries()].map(([id,s])=>({
    id,state:s.state,targetId:s.target?.id||null,scenario:s.scenario,
    since:+s.since.toFixed(2)})),recent:history.slice(-15)};
 }
 function reset(){states.clear();history.length=0;}
 return {update,diagnostic,reset};
}
