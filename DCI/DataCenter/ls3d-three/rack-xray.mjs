// Visual state only: never mutates hinge angles, visibility, or mechanical state.
export function createRackXray() {
  let active=null;const originals=new Map();
  function clear(){for(const [mesh,{original,clones}] of originals){mesh.material=original;for(const material of clones)material.dispose();}originals.clear();active=null;}
  function set(id,doors=[]){
    clear();if(!id)return false;active=id;
    for(const door of doors)door?.traverse(mesh=>{
      if(!mesh.isMesh||!mesh.material)return;
      const original=mesh.material,clones=(Array.isArray(original)?original:[original]).map(material=>{
        const clone=material.clone();clone.transparent=true;clone.opacity=.12;clone.depthWrite=false;return clone;
      });
      originals.set(mesh,{original,clones});mesh.material=Array.isArray(original)?clones:clones[0];
    });return true;
  }
  return {set,clear,get active(){return active;}};
}
