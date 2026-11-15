// Reference-inspired finish palette. Functional cable/LED colors stay semantic.
export function industrialColor(name,color,context='equipment'){
 const n=String(name).toLowerCase();
 if(context==='architecture'){
  if(/asphalt|parking|grass|garden|road|marking|line|car-/.test(n))return color;
  if(/floor|foundation|slab|tile/.test(n))return '#a8adb0';
  if(/wall|facade|ceiling|roof/.test(n))return '#d9dddf';
  if(/beam|column|gantry/.test(n))return '#656b70';
  return color;
 }
 if(/led|indicator|signal|label|badge|light|cord|coolant|outlet|fiber|optic|lc-|mtp|pigtail/.test(n))return color;
 if(/board|pcb|memory/.test(n))return '#287858';
 if(/fan|grille|louver|vent|bezel|port-face|socket|intake|blanking-plate/.test(n))return '#30363a';
 if(/handle|rail|fin|coldplate|heatsink|mount|fastener|cage/.test(n))return '#b9c0c4';
 if(/panel|door|side|plate|shell|cabinet|chassis|housing|case|enclosure|body|stile|header|kickplate/.test(n))return '#dce0e1';
 return color;
}
