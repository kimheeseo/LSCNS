(() => {
 'use strict';
 function render(){
  let b=document.getElementById('bom-engineering-extension');
  if(!b){b=document.createElement('section');b.id='bom-engineering-extension';document.getElementById('engineering-audit')?.insertAdjacentElement('afterend',b);}
  const r=window.DCDesign;if(!r?.usable||r.input.scenario==='colo'){b.innerHTML='';return;}
  b.innerHTML='<h3>Optical quantities — actual design snapshot</h3><table><thead><tr><th>Segment</th><th>Logical links</th><th>Physical modules (installed / purchase)</th><th>Active / occupied fibers</th><th>Trunks (installed / spare)</th></tr></thead><tbody>'+Object.entries(r.optical).map(([key,x])=>`<tr><td>${key}</td><td>${x.links}</td><td>${x.installedModules} / ${x.purchasedModules}</td><td>${x.trunk.requiredFibers} / ${x.trunk.occupiedFibers}</td><td>${x.trunk.installedCableCount} / ${x.trunk.spareCableCount}</td></tr>`).join('')+'</tbody></table><p>Active fibers and connector footprint differ. Trunks are packed within one endpoint pair; spare stock is excluded from installed loss and power. Exact module form factor, breakout, polarity and pinning remain RFQ.</p>';
 }
 document.addEventListener('dc:design',render);document.addEventListener('DOMContentLoaded',render);render();
})();
