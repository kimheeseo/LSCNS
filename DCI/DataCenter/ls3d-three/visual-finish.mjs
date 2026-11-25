/* Surface/lighting only; no meshes, model colors, cameras or simulation state. */
export function createVisualFinish(THREE,renderer,scene){
  const pixels=new Uint8Array(64*64*4);
  for(let y=0;y<64;y++)for(let x=0;x<64;x++){
    const i=(y*64+x)*4,v=124+((x*17+y*31+x*y)%9);
    pixels[i]=pixels[i+1]=pixels[i+2]=v;pixels[i+3]=255;
  }
  const grain=new THREE.DataTexture(pixels,64,64,THREE.RGBAFormat);
  grain.wrapS=grain.wrapT=THREE.RepeatWrapping;grain.repeat.set(4,4);
  grain.magFilter=THREE.LinearFilter;grain.minFilter=THREE.LinearMipmapLinearFilter;
  grain.generateMipmaps=true;grain.needsUpdate=true;
  // A small procedural reflection source supplies neutral service-light bands.
  // It is an environment map, never an image plane or an equipment texture.
  const w=256,h=128,data=new Uint8Array(w*h*4);
  for(let y=0;y<h;y++)for(let x=0;x<w;x++){
    const t=y/(h-1),band=y>28&&y<43&&((x>25&&x<70)||(x>150&&x<200));
    const level=band?200:Math.round(18+36*(1-t));const i=(y*w+x)*4;
    data[i]=level;data[i+1]=Math.min(255,level+3);data[i+2]=Math.min(255,level+7);data[i+3]=255;
  }
  const source=new THREE.DataTexture(data,w,h,THREE.RGBAFormat);
  source.mapping=THREE.EquirectangularReflectionMapping;source.colorSpace=THREE.SRGBColorSpace;source.needsUpdate=true;
  let target=null;
  try{
    const pmrem=new THREE.PMREMGenerator(renderer);
    try{target=pmrem.fromEquirectangular(source);scene.environment=target.texture;}
    finally{pmrem.dispose();}
  }catch(_){/* Directional lighting remains sufficient on limited drivers. */}
  source.dispose();
  return {
    materialOptions(metal,rough,transparent){
      if(transparent)return {metalness:metal,roughness:rough};
      return {metalness:metal,roughness:Math.min(.96,Math.max(.28,rough+.035)),
        bumpMap:grain,bumpScale:.0025,envMapIntensity:metal>=.5?.55:.3};
    },
    dispose(){if(scene.environment===target?.texture)scene.environment=null;target?.dispose();grain.dispose();}
  };
}
